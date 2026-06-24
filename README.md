# Partner Lead Portal

Fastest-usable MVP for partner lead intake:

- partner web form with required fields and template enum dropdowns
- partner Excel upload that parses the template and submits valid rows
- real-time duplicate checks for work email and company plus country
- local admin review table with pending, approved, rejected statuses
- one-click export to the existing `oversea_lead_import_template.xlsx` structure

## Run

```bash
cd /Users/bytedance/Documents/Bytedance/partner-lead-portal
/Users/bytedance/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 server.py
```

Open:

```text
http://127.0.0.1:8787/partner
```

Partner-specific link:

```text
http://127.0.0.1:8787/partner?partner=Partner%20Name
```

Admin review:

```text
http://127.0.0.1:8787/admin
```

Current default behavior:

```text
ADMIN_AUTH_DISABLED=1
```

This means admin API checks are disabled by default. Anyone who can open `/admin` can view leads, update status, and export approved leads. Set `ADMIN_AUTH_DISABLED=0` and configure `ADMIN_KEY` if you want to re-enable admin protection.

To restore admin protection:

```bash
ADMIN_AUTH_DISABLED=0 ADMIN_KEY='your-strong-admin-key' /Users/bytedance/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 server.py
```

## Data

- Lead database: `data/leads.json`
- Excel exports: `exports/`
- Source template: `/Users/bytedance/Downloads/oversea_lead_import_template.xlsx`

## Partner Excel Upload

Partners can upload an `.xlsx` file from `/partner`.

Rules:

- The left panel collects both `Partner Name` and `Activity Name`; review pages group leads by Activity Name and display Partner Name as the source.
- Use the same import template columns.
- Keep row 1 as headers.
- Keep row 2 as the template note/sample row.
- Add actual lead data from row 3.
- Download the standard template from `/template`.
- Every field except `Tracking Code` is required.
- For Excel uploads, valid rows are submitted and rows with errors are skipped with row-level error messages.
- Duplicate leads are blocked when first name + last name, work email, or mobile number already exists.
- Missing columns, blank required values, enum mismatches, duplicate email/company, and cleaned values are returned as row-level warnings for admin review.
- Job titles, company sizes, industries, and sub industries are cleaned against the template enum with exact aliases, keyword rules, numeric/range parsing, and conservative fuzzy matching before review/export.
- Each Excel upload creates one upload batch. Leads from the same file share the same `uploadBatchId`.

Set another template path with:

```bash
LEAD_TEMPLATE_PATH=/path/to/template.xlsx python3 server.py
```

## Current Security Boundary

This version separates the public partner form from admin review:

- `/partner` is public lead submission.
- `/admin` is open by default for review, status changes, and Excel export.

This is enough for MVP testing, but it is not Lark identity security. Anyone with the review URL can view, update, and export leads. To make admin review visible only to your Lark team in production, set `ADMIN_AUTH_DISABLED=0`, configure `ADMIN_KEY`, and put this app behind one of:

- Lark SSO / Lark app login
- company VPN or internal network
- Cloudflare Access / Google Workspace / Okta style access control

## Sharing

`127.0.0.1` only works on your own computer. Partners cannot open that URL.

For quick internal testing on the same network, run with:

```bash
HOST=0.0.0.0 ADMIN_AUTH_DISABLED=1 /Users/bytedance/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 server.py
```

Then share your machine's LAN IP:

```text
http://YOUR-LAN-IP:8787/partner?partner=Partner%20Name
```

For external partners, deploy it to a real HTTPS host. Share only the partner URL externally:

```text
https://your-domain.example/partner?partner=Partner%20Name
```

Keep admin URL inside your Lark group:

```text
https://your-domain.example/admin
```

## Vercel Deployment

This repo includes a Vercel adapter:

- `api/index.py`
- `vercel.json`

After pushing to GitHub, import the repository in Vercel.

Recommended Vercel settings:

```text
Framework Preset: Other
Build Command: leave empty
Output Directory: leave empty
Install Command: pip install -r requirements.txt
```

Environment variable:

```text
ADMIN_AUTH_DISABLED=1
```

After deployment, open:

```text
https://<your-vercel-domain>/partner
```

Important: Vercel serverless storage is not persistent. Submissions may not be reliably retained across cold starts or deployments unless this app is connected to Feishu/Lark Base, a database, or Vercel storage. Use Vercel for the permanent public page; use an external data store for production lead storage.

## Batch Export

Export approved leads from one upload batch:

```text
/api/export?status=approved&batchId=<uploadBatchId>
```
