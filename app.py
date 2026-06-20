import os

from course_report import app


if __name__ == "__main__":
    port = int(os.environ.get("PORT") or 8093)
    print(f"Starting ETQAN v2 on http://127.0.0.1:{port}")
    app.run(port=port, debug=True)
