#!/bin/bash

# Create necessary directories
mkdir -p uploads
mkdir -p output

# Set permissions
chmod -R 755 uploads
chmod -R 755 output

# Start the application
gunicorn app:app \
    --workers 4 \
    --worker-class uvicorn.workers.UvicornWorker \
    --bind 0.0.0.0:$PORT \
    --timeout 120 \
    --log-level info \
    --access-logfile - \
    --error-logfile - 