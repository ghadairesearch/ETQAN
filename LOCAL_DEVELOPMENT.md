# Local Development

ETQAN uses PostgreSQL only. SQLite is no longer used.

## Option 1: Run everything with Docker Compose

```powershell
docker compose up --build
```

Then open:

```text
http://127.0.0.1:8093
```

The local admin email in `docker-compose.yml` is:

```text
admin@example.com
```

Create an account using that email if you want to access:

```text
http://127.0.0.1:8093/admin
```

## Option 2: Run Flask locally with your own PostgreSQL

1. Create a PostgreSQL database named `etqan`.
2. Copy `.env.local.example` to `.env`.
3. Edit `DATABASE_URL` if your username, password, or database name is different.
4. Run:

```powershell
python app.py
```

Then open:

```text
http://127.0.0.1:8093
```

If `DATABASE_URL` is missing, the app will not start because PostgreSQL is required.
