# -*- coding: utf-8 -*-
"""Application config — values loaded from .env."""

from __future__ import annotations

import os
import tempfile
from pathlib import Path

from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

SECRET_KEY = os.getenv("SECRET_KEY") or "dev-insecure-change-me"
SUPABASE_URL = (os.getenv("SUPABASE_URL") or "").strip()
SUPABASE_SERVICE_ROLE_KEY = (os.getenv("SUPABASE_SERVICE_ROLE_KEY") or "").strip()

# Lemon Squeezy
LEMONSQUEEZY_CHECKOUT_URL = (os.getenv("LEMONSQUEEZY_CHECKOUT_URL") or "").strip()
LEMONSQUEEZY_WEBHOOK_SECRET = (os.getenv("LEMONSQUEEZY_WEBHOOK_SECRET") or "").strip()

# User uploads must live under a writable temp dir (Vercel project FS is read-only).
# Default: <tempdir>/Report2Slide/Uploads  →  /tmp/Report2Slide/Uploads on Linux/Vercel
_UPLOAD_ENV = (os.getenv("UPLOAD_ROOT") or "").strip()
if _UPLOAD_ENV:
    UPLOAD_ROOT = Path(_UPLOAD_ENV)
else:
    UPLOAD_ROOT = Path(tempfile.gettempdir()) / "Report2Slide" / "Uploads"

try:
    UPLOAD_ROOT.mkdir(parents=True, exist_ok=True)
except OSError as e:
    # Surface early in logs; per-request mkdir in user_upload_dir is the fallback.
    print(f"[config] Could not create UPLOAD_ROOT={UPLOAD_ROOT}: {e}")

ALLOWED_EXTENSIONS = {"xlsx", "xls"}
