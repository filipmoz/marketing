#!/bin/bash
set -e

# Create data directory if it doesn't exist and fix permissions
# This runs as root (before USER switch)
mkdir -p /app/data
chown -R appuser:appuser /app/data || true
chmod -R 755 /app/data || true

# Switch to appuser and execute the main command
exec gosu appuser "$@"

