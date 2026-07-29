# -*- coding: utf-8 -*-
"""Uygulama yapılandırması — değerler .env dosyasından okunur."""

from __future__ import annotations

import os
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

# Kullanıcı dosyaları: Uploads/<user_id>/
UPLOAD_ROOT = BASE_DIR / "Uploads"

ALLOWED_EXTENSIONS = {"xlsx", "xls"}
