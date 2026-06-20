# Render Deployment Notes

ETQAN now uses PostgreSQL only for permanent data. The included `render.yaml` deploys the app as a Docker web service with a PostgreSQL database and a persistent disk.

Required Render environment variables:

```text
DATABASE_URL=<Render PostgreSQL internal connection string>
SECRET_KEY=<long random secret>
ADMIN_EMAILS=admin@example.com,another-admin@example.com
UPLOAD_FOLDER=/var/data/uploads
ORG_LOGO_FOLDER=/var/data/uploads/organization_logos
```

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
