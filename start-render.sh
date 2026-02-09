#!/usr/bin/env bash
# Render start script — writes Google credentials from env var, then launches gunicorn
set -o errexit

# If GOOGLE_CREDENTIALS_JSON is set, write it to a file so the app can find it
if [ -n "$GOOGLE_CREDENTIALS_JSON" ]; then
    echo "$GOOGLE_CREDENTIALS_JSON" > Credentials/service-account.json
    export SHM_GOOGLE_CREDENTIALS="Credentials/service-account.json"
    echo "Google credentials written from environment variable"
fi

# Render sets PORT; default to 5000 for local testing
PORT="${PORT:-5000}"

exec gunicorn api:app \
    --bind "0.0.0.0:${PORT}" \
    --workers 2 \
    --timeout 120 \
    --access-logfile - \
    --error-logfile -
