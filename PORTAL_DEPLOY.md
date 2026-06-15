# Own Portal Deployment

Use this when you want the full portal to run on your own server or company-managed hosting.

This mode supports the complete feature set tested locally:

- `/partner` public lead form
- single-lead submission
- Excel upload and row-level parsing
- enum and required-field validation
- duplicate email/company checks
- `/admin` review page
- approve, reject, pending status changes
- approved-lead Excel export using the import template
- persistent local data and export files

By default, backend admin checks are disabled. Anyone who can open `/admin` can view, update, and export leads.

## Docker Run

Build:

```bash
docker build -t partner-lead-portal .
```

Run:

```bash
docker run -d \
  --name partner-lead-portal \
  -p 8787:8787 \
  -e ADMIN_AUTH_DISABLED=1 \
  -v partner_lead_data:/data \
  partner-lead-portal
```

Open:

```text
http://YOUR-SERVER:8787/partner
http://YOUR-SERVER:8787/admin
```

## Docker Compose

Start:

```bash
docker compose up -d --build
```

## Health Check

```text
http://YOUR-SERVER:8787/healthz
```

Expected response:

```json
{"status": "ok"}
```

## Production Notes

Put the service behind HTTPS before sharing externally. Recommended options:

- company reverse proxy
- Nginx + TLS certificate
- Cloudflare Tunnel with a named tunnel and fixed domain
- Kubernetes or internal app platform ingress

For admin access, the current default is intentionally open. If this becomes production-critical, set `ADMIN_AUTH_DISABLED=0`, configure `ADMIN_KEY`, and put `/admin` behind Lark SSO, VPN, Cloudflare Access, or your company identity layer.
