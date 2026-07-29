#!/bin/sh
set -e

# Only run pg_basebackup if the data directory is completely empty
if [ -z "$(ls -A "$PGDATA" 2>/dev/null)" ]; then
	echo "Data directory empty. Bootstrapping replica from primary..."
	
	until pg_isready -h db-primary -p 5432 -U accounting >/dev/null 2>&1; do
		sleep 2
	done

	PGPASSWORD="$POSTGRES_REPLICATION_PASSWORD" pg_basebackup \
		-h db-primary \
		-D "$PGDATA" \
		-U replicator \
		-Fp -Xs -P -R
fi

# Hand over execution to the standard postgres entrypoint
exec /usr/local/bin/docker-entrypoint.sh postgres