#!/bin/bash
set -e

HOST="${POSTGRES_HOST:-db}"
PORT="${POSTGRES_PORT:-5432}"

echo "Waiting for PostgreSQL at $HOST:$PORT..."
while ! python -c "import socket; s = socket.socket(); s.settimeout(2); s.connect(('$HOST', $PORT)); s.close()" 2>/dev/null; do
    sleep 1
done
echo "PostgreSQL is ready."

echo "Applying migrations..."
python manage.py migrate --noinput

echo "Collecting static files..."
python manage.py collectstatic --noinput

echo "Starting Gunicorn with ${GUNICORN_WORKERS:-2} workers..."
exec gunicorn remote_mining_management_platform.wsgi:application \
    --bind 0.0.0.0:8000 \
    --workers "${GUNICORN_WORKERS:-2}" \
    --access-logfile - \
    --error-logfile -
