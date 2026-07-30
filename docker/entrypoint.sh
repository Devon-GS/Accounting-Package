#!/bin/sh
set -e

python -c "from app import bootstrap_storage; bootstrap_storage()"

exec gunicorn --workers 4 --bind 0.0.0.0:5000 app:app
