FROM python:3.13-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1

WORKDIR /app

COPY requirements.txt .
RUN pip install -r requirements.txt

COPY app.py build_event_db.py load_to_sqlite.py make_password.py ./
COPY templates ./templates

# data dir for the sqlite db — bind-mount this in production
RUN mkdir -p /data
ENV NASH_DB_PATH=/data/events.db

EXPOSE 5050

# non-root for hygiene; HOME=/tmp so gunicorn's control socket has a writable dir
RUN useradd -r -u 10001 nash && chown -R nash:nash /app /data
USER nash
ENV HOME=/tmp

CMD ["gunicorn", "--bind", "0.0.0.0:5050", "--workers", "2", "--worker-tmp-dir", "/tmp", "--access-logfile", "-", "app:create_app()"]
