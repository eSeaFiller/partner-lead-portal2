# Permanent Deployment

Use Render for the fastest permanent HTTPS URL.

## What This Deploys

- Public partner form: `/partner`
- Excel upload parsing: `/api/upload`
- Internal review page: `/admin`
- Admin and export APIs open by default
- Persistent data and exports stored on a Render disk at `/data`

## Required Render Settings

Create a Render Web Service from this project root.

```text
Language: Python 3
Build Command: pip install -r requirements.txt
Start Command: python server.py
Region: Singapore
```

Environment variables:

```text
HOST=0.0.0.0
ADMIN_AUTH_DISABLED=1
LEAD_DATA_DIR=/data/leads
LEAD_EXPORT_DIR=/data/exports
```

Persistent disk:

```text
Name: lead-data
Mount Path: /data
Size: 1 GB
```

Python version:

```text
3.12.13
```

This is set in `.python-version`.

## URLs After Deployment

Render gives a permanent HTTPS domain:

```text
https://<service-name>.onrender.com
```

Share this with partners:

```text
https://<service-name>.onrender.com/partner?partner=Partner%20Name
```

Keep this internal:

```text
https://<service-name>.onrender.com/admin
```

## Important Boundary

This is permanent hosting, not full Lark SSO. With `ADMIN_AUTH_DISABLED=1`, anyone who can open `/admin` or call the admin APIs can view, update, and export leads. To restore backend protection, set `ADMIN_AUTH_DISABLED=0`, configure `ADMIN_KEY`, and put the review page behind Lark SSO, Cloudflare Access, company VPN, or Feishu/Lark Base.
