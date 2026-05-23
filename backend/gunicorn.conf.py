# gunicorn.conf.py — Render-compatible config
import os

bind = f"0.0.0.0:{os.environ.get('PORT', '5000')}"
workers = 1          # Free tier has 512MB RAM — 1 worker is safer
worker_class = "sync"
timeout = 120        # Vision API calls can be slow
keepalive = 5
loglevel = "info"
accesslog = "-"      # Log to stdout so Render captures it
errorlog = "-"
