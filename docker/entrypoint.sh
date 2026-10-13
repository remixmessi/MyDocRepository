#!/bin/sh
set -eu

# 1Panel may create bind-mounted directories as root on first deployment.
if [ "$(id -u)" = "0" ]; then
    mkdir -p /app/data
    chown -R node:node /app/data
    exec gosu node "$@"
fi

exec "$@"
