#!/bin/sh
set -e

until pg_isready -h db-primary -p 5432; do
	sleep 2
done

find "$PGDATA" -mindepth 1 -maxdepth 1 -exec rm -rf {} +
PGPASSWORD="$POSTGRES_REPLICATION_PASSWORD" pg_basebackup \
	-h db-primary \
	-D "$PGDATA" \
	-U replicator \
	-Fp -Xs -P -R
