# Railway PostgreSQL Setup

1. Add PostgreSQL service in Railway.
2. In web service Variables, add:

```env
DATABASE_URL=${{Postgres.DATABASE_URL}}
SECRET_KEY=change-this-secret
```

3. Start Command:

```bash
gunicorn app:app --bind 0.0.0.0:$PORT
```

4. Deploy. The app will create tables automatically at startup.
5. Backup route if needed: `/init-db`

Default users:
- admin / admin123
- user / user123

Do not commit `.venv`, `*.db`, or `.env`.
