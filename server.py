#!/usr/bin/env python3
import warnings

warnings.simplefilter("ignore", DeprecationWarning)

import cgi
import json
import os
import re
import shutil
import uuid
from io import BytesIO
from datetime import datetime, timezone
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from openpyxl import load_workbook


BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
DEFAULT_RUNTIME_DIR = Path("/tmp/partner-lead-portal") if os.environ.get("VERCEL") else BASE_DIR
DATA_DIR = Path(os.environ.get("LEAD_DATA_DIR", DEFAULT_RUNTIME_DIR / "data"))
EXPORT_DIR = Path(os.environ.get("LEAD_EXPORT_DIR", DEFAULT_RUNTIME_DIR / "exports"))
DATA_FILE = DATA_DIR / "leads.json"
TEMPLATE_PATH = Path(
    os.environ.get(
        "LEAD_TEMPLATE_PATH",
        BASE_DIR / "templates" / "oversea_lead_import_template.xlsx",
    )
)
ADMIN_KEY = os.environ.get("ADMIN_KEY", "lark-admin")
PUBLIC_ONLY = os.environ.get("PUBLIC_ONLY") == "1"

FIELD_TO_HEADER = {
    "firstName": "First Name",
    "lastName": "Last Name",
    "country": "Country",
    "companyName": "Company Name",
    "mobileNumber": "Mobile Number",
    "workEmail": "Work Email",
    "jobTitle": "Job Title",
    "companySize": "Company Size",
    "industry": "Industry",
    "subIndustry": "Sub Industry",
    "trackingCode": "Tracking Code",
}

REQUIRED_FIELDS = [
    "firstName",
    "lastName",
    "country",
    "companyName",
    "workEmail",
    "jobTitle",
    "companySize",
    "industry",
    "subIndustry",
]

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
MAX_UPLOAD_BYTES = 8 * 1024 * 1024


def utc_now():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def normalize_company(value):
    value = (value or "").lower()
    value = re.sub(r"[^a-z0-9]+", " ", value)
    suffixes = {
        "inc",
        "ltd",
        "limited",
        "co",
        "company",
        "corp",
        "corporation",
        "pte",
        "plc",
        "llc",
        "gmbh",
        "sdn",
        "bhd",
    }
    return " ".join(part for part in value.split() if part not in suffixes)


def read_column_values(ws, column, start_row=2):
    values = []
    for row in range(start_row, ws.max_row + 1):
        value = ws[f"{column}{row}"].value
        if value not in (None, ""):
            values.append(str(value).strip())
    return values


def load_schema():
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"Template not found: {TEMPLATE_PATH}")

    # This template stores large enum ranges in a way that openpyxl can
    # under-report in read-only mode, so load normally and keep the schema exact.
    wb = load_workbook(TEMPLATE_PATH, read_only=False, data_only=True)
    country_ws = wb["Country Enums"]
    job_ws = wb["Job Title Enums"]
    size_ws = wb["Company Size Enums"]
    industry_ws = wb["Industry & Sub Industry Enums"]

    countries = [
        {
            "value": str(country_ws[f"A{row}"].value).strip(),
            "label": str(country_ws[f"C{row}"].value or country_ws[f"B{row}"].value or "").strip(),
        }
        for row in range(2, country_ws.max_row + 1)
        if country_ws[f"A{row}"].value
    ]

    job_titles = [
        {
            "value": str(job_ws[f"A{row}"].value).strip(),
            "label": str(job_ws[f"B{row}"].value or job_ws[f"A{row}"].value).strip(),
        }
        for row in range(2, job_ws.max_row + 1)
        if job_ws[f"A{row}"].value
    ]

    company_sizes = read_column_values(size_ws, "A")

    industries = {}
    for row in range(2, industry_ws.max_row + 1):
        industry = industry_ws[f"A{row}"].value
        sub_industry = industry_ws[f"B{row}"].value
        industry_label = industry_ws[f"C{row}"].value
        sub_label = industry_ws[f"D{row}"].value
        if not industry or not sub_industry:
            continue
        industry = str(industry).strip()
        if industry not in industries:
            industries[industry] = {
                "value": industry,
                "label": str(industry_label or industry).strip(),
                "subIndustries": [],
            }
        industries[industry]["subIndustries"].append(
            {
                "value": str(sub_industry).strip(),
                "label": str(sub_label or sub_industry).strip(),
            }
        )

    return {
        "templatePath": str(TEMPLATE_PATH),
        "fields": FIELD_TO_HEADER,
        "requiredFields": REQUIRED_FIELDS,
        "countries": countries,
        "jobTitles": job_titles,
        "companySizes": company_sizes,
        "industries": list(industries.values()),
    }


SCHEMA = load_schema()


def load_leads():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    if not DATA_FILE.exists():
        return []
    with DATA_FILE.open("r", encoding="utf-8") as f:
        return json.load(f)


def save_leads(leads):
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    temp = DATA_FILE.with_suffix(".tmp")
    with temp.open("w", encoding="utf-8") as f:
        json.dump(leads, f, indent=2, ensure_ascii=False)
    temp.replace(DATA_FILE)


def schema_sets():
    countries = {item["value"] for item in SCHEMA["countries"]}
    job_titles = {item["value"] for item in SCHEMA["jobTitles"]}
    company_sizes = set(SCHEMA["companySizes"])
    industries = {item["value"]: {sub["value"] for sub in item["subIndustries"]} for item in SCHEMA["industries"]}
    return countries, job_titles, company_sizes, industries


def duplicate_report(fields, exclude_id=None):
    email = (fields.get("workEmail") or "").strip().lower()
    company_key = normalize_company(fields.get("companyName"))
    country = (fields.get("country") or "").strip()
    leads = load_leads()
    email_matches = []
    company_matches = []

    for lead in leads:
        if exclude_id and lead["id"] == exclude_id:
            continue
        existing = lead.get("fields", {})
        if email and email == (existing.get("workEmail") or "").strip().lower():
            email_matches.append(summarize_lead(lead))
        if company_key and country and country == existing.get("country") and company_key == normalize_company(existing.get("companyName")):
            company_matches.append(summarize_lead(lead))

    return {"email": email_matches, "company": company_matches}


def public_duplicate_report(duplicates):
    return {
        "email": bool(duplicates.get("email")),
        "company": bool(duplicates.get("company")),
    }


def summarize_lead(lead):
    fields = lead.get("fields", {})
    return {
        "id": lead.get("id"),
        "partner": lead.get("partner"),
        "status": lead.get("status"),
        "workEmail": fields.get("workEmail"),
        "companyName": fields.get("companyName"),
        "createdAt": lead.get("createdAt"),
    }


def validate_lead(fields, block_duplicate_email=True):
    errors = {}
    warnings = {}
    countries, job_titles, company_sizes, industries = schema_sets()

    cleaned = {key: str(fields.get(key, "")).strip() for key in FIELD_TO_HEADER}

    for field in REQUIRED_FIELDS:
        if not cleaned[field]:
            errors[field] = "Required"

    if cleaned["workEmail"] and not EMAIL_RE.match(cleaned["workEmail"]):
        errors["workEmail"] = "Invalid email format"

    if cleaned["country"] and cleaned["country"] not in countries:
        errors["country"] = "Country must match template enum"
    if cleaned["jobTitle"] and cleaned["jobTitle"] not in job_titles:
        errors["jobTitle"] = "Job title must match template enum"
    if cleaned["companySize"] and cleaned["companySize"] not in company_sizes:
        errors["companySize"] = "Company size must match template enum"
    if cleaned["industry"] and cleaned["industry"] not in industries:
        errors["industry"] = "Industry must match template enum"
    if cleaned["industry"] in industries and cleaned["subIndustry"] and cleaned["subIndustry"] not in industries[cleaned["industry"]]:
        errors["subIndustry"] = "Sub industry must belong to selected industry"

    duplicates = duplicate_report(cleaned)
    if duplicates["email"]:
        message = "This work email already exists"
        if block_duplicate_email:
            errors["workEmail"] = message
        else:
            warnings["workEmail"] = message
    if duplicates["company"]:
        warnings["companyName"] = "This company and country already exist"

    return cleaned, errors, warnings, duplicates


def make_export(status):
    EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    leads = load_leads()
    selected = [lead for lead in leads if status == "all" or lead.get("status") == status]

    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    output_path = EXPORT_DIR / f"lead-export-{status}-{timestamp}.xlsx"
    shutil.copyfile(TEMPLATE_PATH, output_path)

    wb = load_workbook(output_path)
    ws = wb["Leads"]

    if ws.max_row > 2:
        ws.delete_rows(3, ws.max_row - 2)

    for index, lead in enumerate(selected, start=3):
        fields = lead.get("fields", {})
        for col_index, key in enumerate(FIELD_TO_HEADER, start=1):
            ws.cell(row=index, column=col_index, value=fields.get(key, ""))

    wb.save(output_path)
    return output_path, len(selected)


def parse_uploaded_workbook(file_bytes):
    workbook = load_workbook(BytesIO(file_bytes), read_only=True, data_only=True)
    ws = workbook["Leads"] if "Leads" in workbook.sheetnames else workbook.active
    header_cells = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), [])
    headers = {str(value).strip(): index for index, value in enumerate(header_cells) if value not in (None, "")}
    missing = [header for header in FIELD_TO_HEADER.values() if header not in headers]
    if missing:
        raise ValueError(f"Missing required template columns: {', '.join(missing)}")

    row2 = [value for value in next(ws.iter_rows(min_row=2, max_row=2, values_only=True), [])]
    start_row = 3 if any("Notice" in str(value) for value in row2 if value is not None) else 2

    parsed = []
    for row_number, row in enumerate(ws.iter_rows(min_row=start_row, values_only=True), start=start_row):
        fields = {}
        has_value = False
        for field, header in FIELD_TO_HEADER.items():
            value = row[headers[header]] if headers[header] < len(row) else None
            if value not in (None, ""):
                has_value = True
            fields[field] = "" if value is None else str(value).strip()
        if has_value:
            parsed.append({"row": row_number, "fields": fields})
    return parsed


def import_uploaded_leads(file_bytes, partner):
    parsed_rows = parse_uploaded_workbook(file_bytes)
    leads = load_leads()
    imported = []
    failed = []
    seen_emails = set()

    for item in parsed_rows:
        cleaned, errors, warnings, duplicates = validate_lead(item["fields"])
        email_key = cleaned.get("workEmail", "").lower()
        if email_key and email_key in seen_emails:
            errors["workEmail"] = "Duplicate email inside uploaded file"
        if errors:
            failed.append(
                {
                    "row": item["row"],
                    "email": cleaned.get("workEmail", ""),
                    "companyName": cleaned.get("companyName", ""),
                    "errors": errors,
                }
            )
            continue

        seen_emails.add(email_key)
        lead = {
            "id": uuid.uuid4().hex,
            "partner": partner,
            "status": "pending",
            "createdAt": utc_now(),
            "updatedAt": utc_now(),
            "reviewNote": "",
            "warnings": warnings,
            "fields": cleaned,
        }
        imported.append(lead)

    if imported:
        save_leads(imported + leads)

    return {
        "parsed": len(parsed_rows),
        "imported": len(imported),
        "failed": len(failed),
        "errors": failed[:50],
    }


class LeadPortalHandler(BaseHTTPRequestHandler):
    server_version = "PartnerLeadPortal/0.1"

    def log_message(self, fmt, *args):
        print("[%s] %s" % (self.log_date_time_string(), fmt % args))

    def send_json(self, payload, status=HTTPStatus.OK):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def read_json(self):
        length = int(self.headers.get("Content-Length", "0"))
        if length == 0:
            return {}
        raw = self.rfile.read(length)
        return json.loads(raw.decode("utf-8"))

    def read_multipart_upload(self):
        length = int(self.headers.get("Content-Length", "0"))
        if length > MAX_UPLOAD_BYTES:
            raise ValueError("File is too large. Maximum upload size is 8 MB.")
        environ = {
            "REQUEST_METHOD": "POST",
            "CONTENT_TYPE": self.headers.get("Content-Type", ""),
            "CONTENT_LENGTH": str(length),
        }
        form = cgi.FieldStorage(fp=self.rfile, headers=self.headers, environ=environ)
        upload = form["file"] if "file" in form else None
        if upload is None or not getattr(upload, "filename", ""):
            raise ValueError("No Excel file uploaded.")
        if not upload.filename.lower().endswith(".xlsx"):
            raise ValueError("Only .xlsx files are supported.")
        partner = form.getfirst("partner", "Unknown Partner").strip() or "Unknown Partner"
        return upload.file.read(), partner

    def is_admin_request(self):
        parsed = urlparse(self.path)
        query = parse_qs(parsed.query)
        provided = self.headers.get("X-Admin-Key") or query.get("admin_key", [""])[0]
        return provided == ADMIN_KEY

    def require_admin(self):
        if self.is_admin_request():
            return True
        self.send_json({"error": "Admin access required"}, HTTPStatus.UNAUTHORIZED)
        return False

    def serve_static(self, parsed):
        path = parsed.path
        if PUBLIC_ONLY and path == "/admin":
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        if path in {"/", "/partner", "/admin"}:
            path = "/index.html"
        file_path = (STATIC_DIR / path.lstrip("/")).resolve()
        if not str(file_path).startswith(str(STATIC_DIR.resolve())) or not file_path.exists():
            self.send_error(HTTPStatus.NOT_FOUND)
            return

        content_type = "text/plain; charset=utf-8"
        if file_path.suffix == ".html":
            content_type = "text/html; charset=utf-8"
        elif file_path.suffix == ".css":
            content_type = "text/css; charset=utf-8"
        elif file_path.suffix == ".js":
            content_type = "application/javascript; charset=utf-8"

        body = file_path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        parsed = urlparse(self.path)
        query = parse_qs(parsed.query)

        if parsed.path == "/api/schema":
            self.send_json(SCHEMA)
            return

        if parsed.path == "/api/leads":
            if PUBLIC_ONLY:
                self.send_error(HTTPStatus.NOT_FOUND)
                return
            if not self.require_admin():
                return
            leads = load_leads()
            self.send_json({"leads": leads})
            return

        if parsed.path == "/api/check":
            fields = {
                "workEmail": query.get("workEmail", [""])[0],
                "companyName": query.get("companyName", [""])[0],
                "country": query.get("country", [""])[0],
            }
            self.send_json({"duplicates": public_duplicate_report(duplicate_report(fields))})
            return

        if parsed.path == "/api/export":
            if PUBLIC_ONLY:
                self.send_error(HTTPStatus.NOT_FOUND)
                return
            if not self.require_admin():
                return
            status = query.get("status", ["approved"])[0]
            if status not in {"approved", "pending", "rejected", "all"}:
                self.send_json({"error": "Invalid export status"}, HTTPStatus.BAD_REQUEST)
                return
            output_path, count = make_export(status)
            self.send_json({"path": str(output_path), "count": count})
            return

        if parsed.path.startswith("/exports/"):
            if PUBLIC_ONLY:
                self.send_error(HTTPStatus.NOT_FOUND)
                return
            if not self.require_admin():
                return
            file_path = (EXPORT_DIR / parsed.path.removeprefix("/exports/")).resolve()
            if not str(file_path).startswith(str(EXPORT_DIR.resolve())) or not file_path.exists():
                self.send_error(HTTPStatus.NOT_FOUND)
                return
            body = file_path.read_bytes()
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", f'attachment; filename="{file_path.name}"')
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return

        self.serve_static(parsed)

    def do_POST(self):
        parsed = urlparse(self.path)

        if parsed.path == "/api/leads":
            payload = self.read_json()
            partner = str(payload.get("partner", "")).strip() or "Unknown Partner"
            cleaned, errors, warnings, duplicates = validate_lead(payload.get("fields", {}))
            if errors:
                self.send_json(
                    {
                        "errors": errors,
                        "warnings": warnings,
                        "duplicates": public_duplicate_report(duplicates),
                    },
                    HTTPStatus.BAD_REQUEST,
                )
                return

            leads = load_leads()
            lead = {
                "id": uuid.uuid4().hex,
                "partner": partner,
                "status": "pending",
                "createdAt": utc_now(),
                "updatedAt": utc_now(),
                "reviewNote": "",
                "warnings": warnings,
                "fields": cleaned,
            }
            leads.insert(0, lead)
            save_leads(leads)
            self.send_json({"lead": lead, "warnings": warnings}, HTTPStatus.CREATED)
            return

        if parsed.path == "/api/upload":
            try:
                file_bytes, partner = self.read_multipart_upload()
                result = import_uploaded_leads(file_bytes, partner)
            except Exception as exc:
                self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)
                return
            self.send_json(result, HTTPStatus.CREATED)
            return

        self.send_error(HTTPStatus.NOT_FOUND)

    def do_PATCH(self):
        parsed = urlparse(self.path)
        if not parsed.path.startswith("/api/leads/"):
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        if PUBLIC_ONLY:
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        if not self.require_admin():
            return

        lead_id = parsed.path.split("/")[-1]
        payload = self.read_json()
        leads = load_leads()
        for lead in leads:
            if lead["id"] == lead_id:
                status = payload.get("status")
                if status:
                    if status not in {"pending", "approved", "rejected"}:
                        self.send_json({"error": "Invalid status"}, HTTPStatus.BAD_REQUEST)
                        return
                    lead["status"] = status
                if "reviewNote" in payload:
                    lead["reviewNote"] = str(payload.get("reviewNote") or "").strip()
                lead["updatedAt"] = utc_now()
                save_leads(leads)
                self.send_json({"lead": lead})
                return

        self.send_json({"error": "Lead not found"}, HTTPStatus.NOT_FOUND)


def main():
    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", "8787"))
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    server = ThreadingHTTPServer((host, port), LeadPortalHandler)
    print(f"Partner Lead Portal running at http://{host}:{port}")
    print(f"Template: {TEMPLATE_PATH}")
    if PUBLIC_ONLY:
        print("Public-only mode: admin routes and export APIs are disabled")
    server.serve_forever()


if __name__ == "__main__":
    main()
