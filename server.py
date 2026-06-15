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
CORS_ALLOW_ORIGIN = os.environ.get("CORS_ALLOW_ORIGIN", "")

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

FIELD_ALIASES = {
    "firstName": ["first", "given name", "given", "first_name", "firstname", "名"],
    "lastName": ["last", "family name", "surname", "last_name", "lastname", "姓"],
    "country": ["country code", "country/region", "region", "market", "国家", "国家/地区"],
    "companyName": ["company", "account", "organization", "organisation", "company_name", "公司", "公司名称"],
    "mobileNumber": ["mobile", "phone", "phone number", "mobile phone", "tel", "telephone", "手机号", "电话"],
    "workEmail": ["email", "work mail", "business email", "work_email", "mail", "邮箱", "工作邮箱"],
    "jobTitle": ["title", "role", "position", "job", "job_title", "职位", "职务"],
    "companySize": ["size", "employee size", "employees", "headcount", "company_size", "公司规模"],
    "industry": ["sector", "vertical", "行业"],
    "subIndustry": ["subindustry", "sub industry", "sub-sector", "sub sector", "子行业"],
    "trackingCode": ["tracking", "campaign", "campaign code", "tracking_code", "活动代码"],
}

JOB_TITLE_ALIASES = {
    "chief executive officer": "CEO",
    "founder": "CEO",
    "owner": "CEO",
    "chief technology officer": "CTO",
    "chief information officer": "CIO",
    "chief financial officer": "CFO",
    "chief marketing officer": "CMO",
    "chief operating officer": "COO",
    "manager": "Manager",
    "director": "Director",
    "engineer": "Engineering",
    "employee": "Employee",
    "student": "Student",
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


def normalize_lookup_key(value):
    return re.sub(r"[\W_]+", "", str(value or "").strip().lower(), flags=re.UNICODE)


def normalize_header(value):
    return normalize_lookup_key(value)


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


def enum_lookup_maps():
    country_map = {}
    for item in SCHEMA["countries"]:
        country_map[normalize_lookup_key(item["value"])] = item["value"]
        country_map[normalize_lookup_key(item["label"])] = item["value"]

    job_map = {}
    for item in SCHEMA["jobTitles"]:
        job_map[normalize_lookup_key(item["value"])] = item["value"]
        job_map[normalize_lookup_key(item["label"])] = item["value"]
    for alias, value in JOB_TITLE_ALIASES.items():
        job_map[normalize_lookup_key(alias)] = value

    size_map = {normalize_lookup_key(size): size for size in SCHEMA["companySizes"]}
    size_map.update(
        {
            normalize_lookup_key("1 to 10"): "1-10",
            normalize_lookup_key("11 to 20"): "11-20",
            normalize_lookup_key("21 to 50"): "21-50",
            normalize_lookup_key("51 to 100"): "51-100",
            normalize_lookup_key("101 to 250"): "101-250",
            normalize_lookup_key("251 to 500"): "251-500",
            normalize_lookup_key("501 to 999"): "501-999",
            normalize_lookup_key("1000 to 4999"): "1000-4999",
            normalize_lookup_key("5000 to 9999"): "5000-9999",
            normalize_lookup_key("10000+"): ">10000",
            normalize_lookup_key("> 10000"): ">10000",
            normalize_lookup_key("more than 10000"): ">10000",
        }
    )

    industry_map = {}
    subindustry_map = {}
    subindustry_by_industry = {}
    for industry in SCHEMA["industries"]:
        industry_map[normalize_lookup_key(industry["value"])] = industry["value"]
        industry_map[normalize_lookup_key(industry["label"])] = industry["value"]
        subindustry_by_industry[industry["value"]] = {}
        for sub in industry["subIndustries"]:
            subindustry_map[normalize_lookup_key(sub["value"])] = sub["value"]
            subindustry_map[normalize_lookup_key(sub["label"])] = sub["value"]
            subindustry_by_industry[industry["value"]][normalize_lookup_key(sub["value"])] = sub["value"]
            subindustry_by_industry[industry["value"]][normalize_lookup_key(sub["label"])] = sub["value"]

    return {
        "country": country_map,
        "jobTitle": job_map,
        "companySize": size_map,
        "industry": industry_map,
        "subIndustry": subindustry_map,
        "subIndustryByIndustry": subindustry_by_industry,
    }


def map_enum_value(field, value, maps, warnings):
    original = str(value or "").strip()
    if not original:
        return ""
    mapped = maps[field].get(normalize_lookup_key(original))
    if mapped:
        if mapped != original:
            warnings[field] = f"Cleaned '{original}' to '{mapped}'"
        return mapped
    return original


def clean_upload_fields(fields, missing_columns):
    maps = enum_lookup_maps()
    cleaned = {key: str(fields.get(key, "")).strip() for key in FIELD_TO_HEADER}
    warnings = {}

    if cleaned["workEmail"]:
        lowered = cleaned["workEmail"].replace("mailto:", "").strip().lower()
        if lowered != cleaned["workEmail"]:
            warnings["workEmail"] = f"Cleaned email '{cleaned['workEmail']}' to '{lowered}'"
        cleaned["workEmail"] = lowered

    for field in ["country", "jobTitle", "companySize", "industry"]:
        cleaned[field] = map_enum_value(field, cleaned[field], maps, warnings)

    if cleaned["subIndustry"]:
        by_industry = maps["subIndustryByIndustry"].get(cleaned["industry"], {})
        original = cleaned["subIndustry"]
        mapped = by_industry.get(normalize_lookup_key(original)) or maps["subIndustry"].get(normalize_lookup_key(original))
        if mapped:
            if mapped != original:
                warnings["subIndustry"] = f"Cleaned '{original}' to '{mapped}'"
            cleaned["subIndustry"] = mapped

    for header in missing_columns:
        field = next((key for key, value in FIELD_TO_HEADER.items() if value == header), None)
        if field:
            warnings[field] = f"Missing column '{header}'; treated as blank"

    return cleaned, warnings


def warnings_from_errors(errors):
    return {field: message for field, message in errors.items()}


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


def upload_header_candidates(field):
    return [FIELD_TO_HEADER[field], *FIELD_ALIASES.get(field, [])]


def build_upload_header_index(header_cells):
    normalized_headers = {
        normalize_header(value): index
        for index, value in enumerate(header_cells)
        if value not in (None, "")
    }
    header_index = {}
    for field in FIELD_TO_HEADER:
        for candidate in upload_header_candidates(field):
            key = normalize_header(candidate)
            if key in normalized_headers:
                header_index[field] = normalized_headers[key]
                break
    return header_index


def looks_like_instruction_row(row):
    values = [str(value or "").strip().lower() for value in row if value not in (None, "")]
    if not values:
        return False
    joined = " ".join(values)
    has_email = any(EMAIL_RE.match(value) for value in values)
    instruction_tokens = ["notice", "required", "sample", "example", "do not delete", "note"]
    repeated_placeholder = len(set(values)) == 1 and values[0] in {"note", "sample", "example", "required"}
    return not has_email and (repeated_placeholder or any(token in joined for token in instruction_tokens))


def parse_uploaded_workbook(file_bytes):
    workbook = load_workbook(BytesIO(file_bytes), read_only=True, data_only=True)
    ws = workbook["Leads"] if "Leads" in workbook.sheetnames else workbook.active
    header_cells = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), [])
    headers = build_upload_header_index(header_cells)
    recognized = [FIELD_TO_HEADER[field] for field in FIELD_TO_HEADER if field in headers]
    missing = [FIELD_TO_HEADER[field] for field in FIELD_TO_HEADER if field not in headers]
    if not recognized:
        raise ValueError("No matching lead columns found. Keep at least one template header in row 1.")

    row2 = [value for value in next(ws.iter_rows(min_row=2, max_row=2, values_only=True), [])]
    start_row = 3 if looks_like_instruction_row(row2) else 2

    parsed = []
    for row_number, row in enumerate(ws.iter_rows(min_row=start_row, values_only=True), start=start_row):
        fields = {}
        has_value = False
        for field, header in FIELD_TO_HEADER.items():
            if field not in headers:
                fields[field] = ""
                continue
            value = row[headers[field]] if headers[field] < len(row) else None
            if value not in (None, ""):
                has_value = True
            fields[field] = "" if value is None else str(value).strip()
        if has_value:
            parsed.append({"row": row_number, "fields": fields})
    return parsed, missing


def import_uploaded_leads(file_bytes, partner):
    parsed_rows, missing_columns = parse_uploaded_workbook(file_bytes)
    leads = load_leads()
    imported = []
    failed = []
    row_warnings = []
    seen_emails = set()

    for item in parsed_rows:
        cleaned, cleanup_warnings = clean_upload_fields(item["fields"], missing_columns)
        cleaned, errors, validation_warnings, duplicates = validate_lead(cleaned, block_duplicate_email=False)
        warnings = {**cleanup_warnings, **validation_warnings, **warnings_from_errors(errors)}
        email_key = cleaned.get("workEmail", "").lower()
        if email_key and email_key in seen_emails:
            warnings["workEmail"] = "Duplicate email inside uploaded file"

        if email_key:
            seen_emails.add(email_key)
        if warnings:
            row_warnings.append(
                {
                    "row": item["row"],
                    "email": cleaned.get("workEmail", ""),
                    "companyName": cleaned.get("companyName", ""),
                    "warnings": warnings,
                }
            )
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
        "missingColumns": missing_columns,
        "warnings": row_warnings[:50],
        "errors": failed[:50],
    }


class LeadPortalHandler(BaseHTTPRequestHandler):
    server_version = "PartnerLeadPortal/0.1"

    def log_message(self, fmt, *args):
        print("[%s] %s" % (self.log_date_time_string(), fmt % args))

    def send_json(self, payload, status=HTTPStatus.OK):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_cors_headers()
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def send_cors_headers(self):
        if not CORS_ALLOW_ORIGIN:
            return
        self.send_header("Access-Control-Allow-Origin", CORS_ALLOW_ORIGIN)
        self.send_header("Access-Control-Allow-Methods", "GET, POST, PATCH, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, X-Admin-Key")
        self.send_header("Access-Control-Max-Age", "86400")

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
        self.send_cors_headers()
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(HTTPStatus.NO_CONTENT)
        self.send_cors_headers()
        self.end_headers()

    def do_GET(self):
        parsed = urlparse(self.path)
        query = parse_qs(parsed.query)

        if parsed.path == "/api/schema":
            self.send_json(SCHEMA)
            return

        if parsed.path == "/healthz":
            self.send_json({"status": "ok"})
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
            self.send_cors_headers()
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
