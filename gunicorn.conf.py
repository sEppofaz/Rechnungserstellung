# Wird automatisch geladen (WorkingDirectory /opt/kargl-invoice/src).
# OCR mit Opus 5.5 + Denken kann >30 s (gunicorn-Standard) dauern; nginx /kargl/ wartet 120 s.
timeout = 120
