# ETQAN

Educational Transformation & Quality ANalytics

A Flask app for uploading assessment reports, mapping questions to CLOs, calculating CLO attainment, and exporting formal CSV/PDF reports.

## Local Run

```powershell
pip install -r requirements.txt
python app.py
```

The local app starts at `http://127.0.0.1:8093`.

## Render

Use this start command:

```bash
gunicorn course_report:app
```
