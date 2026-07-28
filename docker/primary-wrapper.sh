#!/bin/sh
set -e

if [ -f "$PGDATA/pg_hba.conf" ] && ! grep -q '^host replication replicator all scram-sha-256$' "$PGDATA/pg_hba.conf"; then
	echo "host replication replicator all scram-sha-256" >> "$PGDATA/pg_hba.conf"
fi

/usr/local/bin/docker-entrypoint.sh postgres -c wal_level=replica -c max_wal_senders=5 -c max_replication_slots=5 &
postgres_pid=$!

until pg_isready -h 127.0.0.1 -U "$POSTGRES_USER" -d "$POSTGRES_DB" >/dev/null 2>&1; do
	sleep 1
done

if [ "$(PGPASSWORD="$POSTGRES_PASSWORD" psql -h 127.0.0.1 -tAc "SELECT 1 FROM pg_roles WHERE rolname = 'replicator'" --username "$POSTGRES_USER" --dbname "$POSTGRES_DB")" = "1" ]; then
	replication_password_sql=$(printf "%s" "$POSTGRES_REPLICATION_PASSWORD" | sed "s/'/''/g")
	PGPASSWORD="$POSTGRES_PASSWORD" psql -h 127.0.0.1 -v ON_ERROR_STOP=1 --username "$POSTGRES_USER" --dbname "$POSTGRES_DB" -c "ALTER ROLE replicator WITH REPLICATION LOGIN PASSWORD '$replication_password_sql';"
else
	replication_password_sql=$(printf "%s" "$POSTGRES_REPLICATION_PASSWORD" | sed "s/'/''/g")
	PGPASSWORD="$POSTGRES_PASSWORD" psql -h 127.0.0.1 -v ON_ERROR_STOP=1 --username "$POSTGRES_USER" --dbname "$POSTGRES_DB" -c "CREATE ROLE replicator WITH REPLICATION LOGIN PASSWORD '$replication_password_sql';"
fi

PGPASSWORD="$POSTGRES_PASSWORD" psql -h 127.0.0.1 -v ON_ERROR_STOP=1 --username "$POSTGRES_USER" --dbname "$POSTGRES_DB" -c "SELECT pg_reload_conf();"

touch /tmp/replication-ready
wait "$postgres_pid"
