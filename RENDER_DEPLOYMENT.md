# Render Deployment Notes

ETQAN now uses PostgreSQL only for permanent data. The included `render.yaml` deploys the app as a Docker web service with a PostgreSQL database and a persistent disk.

Required Render environment variables:

```text
DATABASE_URL=<Render PostgreSQL internal connection string>
SECRET_KEY=<long random secret>
ADMIN_EMAILS=admin@example.com,another-admin@example.com
UPLOAD_FOLDER=/var/data/uploads
ORG_LOGO_FOLDER=/var/data/uploads/organization_logos
RESEND_API_KEY=<Resend API key>
RESEND_FROM_EMAIL=onboarding@resend.dev
CONTACT_TO_EMAIL=ghad.ai.research@gmail.com
APP_PUBLIC_URL=https://your-etqan-service.onrender.com
```

For production email, verify your own domain in Resend and replace `RESEND_FROM_EMAIL` with an address on that domain. `onboarding@resend.dev` is useful for initial testing.

Recommended Render setup:

1. Connect the GitHub repository to Render.
2. Use the Blueprint option so Render reads `render.yaml`.
3. Confirm the web service, PostgreSQL database, and `/var/data` persistent disk.
4. Set `ADMIN_EMAILS` to the email address that should access `/admin`.
5. Deploy.

The Docker image installs:

- Python 3.12
- Tesseract OCR with Arabic and English language data
- Poppler utilities for PDF image conversion
- Python packages from `requirements.txt`

If the Render UI does not show Docker when creating a service manually, use **New Blueprint** instead of **New Web Service**. The repository includes `render.yaml`, and the blueprint tells Render to use the Dockerfile automatically.

Native Python fallback:

- Build command: `pip install -r requirements.txt`
- Start command: `gunicorn course_report:app --bind 0.0.0.0:$PORT --workers 2 --threads 4 --timeout 180`
- This can run the app, but scanned Arabic PDF OCR depends on system packages. `Aptfile` lists the needed packages for environments that support apt buildpacks, but Docker/Blueprint is the reliable option.

Permanent data stored in PostgreSQL:

- User accounts and subscription status
- Organization identity colors and logo file names
- Saved courses
- Saved reports
- Contact requests and attachment metadata

Files stored on the Persistent Disk:

- Uploaded organization logos
- Contact request attachments
- Temporary uploaded assessment files while a session is active

Admin dashboard:

- URL: `/admin`
- Access: only emails listed in `ADMIN_EMAILS`
- Shows user count, saved courses, saved reports, and contact requests
