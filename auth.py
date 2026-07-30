# -*- coding: utf-8 -*-
"""Supabase Auth yardımcıları ve login_required decorator."""

from __future__ import annotations

import re
from functools import wraps
from typing import Any

from flask import jsonify, redirect, request, session, url_for
from supabase import Client, create_client

import config

_supabase: Client | None = None

_UUID_SAFE = re.compile(r"^[0-9a-fA-F-]{8,64}$")


def get_supabase() -> Client:
    """Tekil Supabase istemcisi (service role)."""
    global _supabase
    if _supabase is None:
        if not config.SUPABASE_URL or not config.SUPABASE_SERVICE_ROLE_KEY:
            raise RuntimeError(
                "SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY must be set in the .env file."
            )
        _supabase = create_client(config.SUPABASE_URL, config.SUPABASE_SERVICE_ROLE_KEY)
    return _supabase


def is_authenticated() -> bool:
    return bool(session.get("user_id"))


def current_user_id() -> str | None:
    uid = session.get("user_id")
    return str(uid) if uid else None


def current_user_email() -> str | None:
    return session.get("user_email")


def set_user_session(user: Any, access_token: str | None = None, refresh_token: str | None = None) -> None:
    """Supabase user nesnesinden Flask session yazar."""
    session.clear()
    session["user_id"] = str(user.id)
    session["user_email"] = getattr(user, "email", None) or ""
    if access_token:
        session["access_token"] = access_token
    if refresh_token:
        session["refresh_token"] = refresh_token
    session.permanent = True


def clear_user_session() -> None:
    session.clear()


def sanitize_user_id(user_id: str) -> str:
    """Klasör adı için user_id doğrula (path traversal engeli)."""
    uid = str(user_id).strip()
    if not _UUID_SAFE.match(uid):
        # Supabase UUID dışı id gelirse alfanumerik + tire ile sınırla
        cleaned = "".join(c for c in uid if c.isalnum() or c in "-_")
        if not cleaned or ".." in cleaned:
            raise ValueError("Invalid user ID.")
        return cleaned[:128]
    return uid


def user_upload_dir(user_id: str | None = None) -> str:
    """
    Kullanıcıya özel upload kökü: Uploads/<user_id>/
    Klasör yoksa oluşturur.
    """
    uid = user_id or current_user_id()
    if not uid:
        raise RuntimeError("Authentication required.")
    safe = sanitize_user_id(uid)
    path = config.UPLOAD_ROOT / safe
    path.mkdir(parents=True, exist_ok=True)
    return str(path)


def _wants_json_error() -> bool:
    """API / XHR isteklerinde JSON 401 dön."""
    path = request.path or ""
    if path.startswith("/api/") or path in ("/upload", "/create_presentation"):
        return True
    if request.headers.get("X-Requested-With") == "XMLHttpRequest":
        return True
    accept = request.headers.get("Accept", "")
    if "application/json" in accept and "text/html" not in accept:
        return True
    return False


def login_required(view):
    """Giriş zorunlu; aksi halde login sayfasına veya 401 JSON."""

    @wraps(view)
    def wrapped(*args, **kwargs):
        if not is_authenticated():
            if _wants_json_error():
                return jsonify({"error": "Authentication required. Please log in."}), 401
            return redirect(url_for("login", next=request.path))
        return view(*args, **kwargs)

    return wrapped


def auth_error_message(exc: Exception) -> str:
    """Supabase / Auth hatalarını kullanıcıya okunur metne çevir."""
    msg = str(exc).strip() or "Authentication error."
    lower = msg.lower()
    if "invalid login" in lower or "invalid_credentials" in lower:
        return "Invalid email or password."
    if "user already registered" in lower or "already been registered" in lower:
        return "This email is already registered."
    if "password" in lower and ("least" in lower or "weak" in lower or "short" in lower):
        return "Password is too short or does not meet security requirements (min. 6 characters)."
    if "email" in lower and "invalid" in lower:
        return "Invalid email address."
    if "supabase_url" in lower or "service_role" in lower or ".env" in lower:
        return "Server configuration is incomplete. Please ask an admin to check the .env file."
    return msg
