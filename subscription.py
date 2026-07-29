# -*- coding: utf-8 -*-
"""
Lemon Squeezy abonelik + Supabase profiles plan yönetimi.

Supabase'de `supabase_schema.sql` dosyasını çalıştırın.
"""

from __future__ import annotations

import hashlib
import hmac
import logging
from datetime import datetime, timezone
from functools import wraps
from typing import Any

from flask import jsonify, redirect, request, url_for

import config
from auth import (
    _wants_json_error,
    current_user_email,
    current_user_id,
    get_supabase,
    is_authenticated,
    login_required,
)

logger = logging.getLogger(__name__)

PRO_STATUSES = frozenset({"active", "on_trial", "paused"})
# cancelled: dönem bitene kadar erişim (ends_at gelecekteyse)
KEEP_ACCESS_STATUSES = frozenset({"active", "on_trial", "paused", "cancelled"})
HANDLED_EVENTS = frozenset({
    "subscription_created",
    "subscription_updated",
    "subscription_cancelled",
})


def verify_lemonsqueezy_signature(raw_body: bytes, signature_header: str | None) -> bool:
    """X-Signature = HMAC-SHA256(raw_body, webhook_secret) hex."""
    secret = config.LEMONSQUEEZY_WEBHOOK_SECRET
    if not secret:
        logger.error("LEMONSQUEEZY_WEBHOOK_SECRET tanımlı değil.")
        return False
    if not signature_header:
        return False
    digest = hmac.new(
        secret.encode("utf-8"),
        raw_body,
        hashlib.sha256,
    ).hexdigest()
    return hmac.compare_digest(digest, signature_header.strip())


def _parse_ends_at(ends_at: Any) -> datetime | None:
    if not ends_at:
        return None
    try:
        s = str(ends_at).replace("Z", "+00:00")
        dt = datetime.fromisoformat(s)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt
    except Exception:
        return None


def plan_from_subscription_status(status: str | None, ends_at: Any = None) -> str:
    """Lemon Squeezy subscription status → free | pro."""
    st = (status or "").strip().lower()
    if st in PRO_STATUSES:
        return "pro"
    if st == "cancelled":
        end = _parse_ends_at(ends_at)
        if end and end > datetime.now(timezone.utc):
            return "pro"
        # ends_at yoksa iptal edilmiş ama henüz expired event gelmemiş olabilir
        return "pro" if end is None else "free"
    return "free"


def ensure_profile(user_id: str, email: str | None = None) -> dict[str, Any]:
    """profiles satırı yoksa free plan ile oluşturur."""
    sb = get_supabase()
    res = sb.table("profiles").select("*").eq("id", user_id).limit(1).execute()
    rows = res.data or []
    if rows:
        row = rows[0]
        if email and row.get("email") != email:
            sb.table("profiles").update({"email": email}).eq("id", user_id).execute()
            row["email"] = email
        return row
    payload = {
        "id": user_id,
        "email": (email or "").strip().lower() or None,
        "plan": "free",
        "subscription_status": None,
    }
    ins = sb.table("profiles").insert(payload).execute()
    if ins.data:
        return ins.data[0]
    return payload


def get_profile(user_id: str | None = None) -> dict[str, Any] | None:
    uid = user_id or current_user_id()
    if not uid:
        return None
    try:
        sb = get_supabase()
        res = sb.table("profiles").select("*").eq("id", uid).limit(1).execute()
        rows = res.data or []
        if rows:
            return rows[0]
        email = current_user_email() if uid == current_user_id() else None
        return ensure_profile(uid, email)
    except Exception as e:
        logger.warning("get_profile failed: %s", e)
        return None


def user_has_pro_access(user_id: str | None = None) -> bool:
    profile = get_profile(user_id)
    if not profile:
        return False
    plan = (profile.get("plan") or "free").lower()
    status = (profile.get("subscription_status") or "").lower()
    if plan != "pro":
        return False
    if status and status not in KEEP_ACCESS_STATUSES and status not in ("",):
        # expired / unpaid vb.
        if status in ("expired", "unpaid"):
            return False
    return True


def checkout_url_for_current_user() -> str:
    """Checkout URL + e-posta + custom user_id (webhook eşleştirmesi için)."""
    base = config.LEMONSQUEEZY_CHECKOUT_URL
    if not base:
        return ""
    from urllib.parse import parse_qsl, urlencode, urlparse, urlunparse

    parsed = urlparse(base)
    pairs = list(parse_qsl(parsed.query, keep_blank_values=True))
    # Aynı anahtarları yeniden yazmak için önce temizle
    keys_to_set = {"checkout[email]", "checkout[custom][user_id]"}
    pairs = [(k, v) for k, v in pairs if k not in keys_to_set]
    email = current_user_email()
    uid = current_user_id()
    if email:
        pairs.append(("checkout[email]", email))
    if uid:
        pairs.append(("checkout[custom][user_id]", uid))
    return urlunparse(parsed._replace(query=urlencode(pairs)))


def pro_required(view):
    """Giriş + Pro abonelik zorunlu. Değilse /pricing veya 403 JSON."""

    @login_required
    @wraps(view)
    def wrapped(*args, **kwargs):
        uid = current_user_id()
        try:
            if uid:
                ensure_profile(uid, current_user_email())
        except Exception as e:
            logger.warning("ensure_profile in pro_required: %s", e)

        if user_has_pro_access(uid):
            return view(*args, **kwargs)

        if _wants_json_error():
            return jsonify({
                "error": "Bu özellik Pro abonelik gerektirir.",
                "code": "pro_required",
                "pricing_url": url_for("pricing"),
            }), 403
        return redirect(url_for("pricing", reason="pro_required", next=request.path))

    return wrapped


def _find_profile_for_webhook(
    *,
    user_id: str | None,
    email: str | None,
    customer_id: str | None,
) -> dict[str, Any] | None:
    sb = get_supabase()
    if user_id:
        res = sb.table("profiles").select("*").eq("id", user_id).limit(1).execute()
        if res.data:
            return res.data[0]
    if email:
        res = (
            sb.table("profiles")
            .select("*")
            .eq("email", email.strip().lower())
            .limit(1)
            .execute()
        )
        if res.data:
            return res.data[0]
    if customer_id:
        res = (
            sb.table("profiles")
            .select("*")
            .eq("lemon_customer_id", str(customer_id))
            .limit(1)
            .execute()
        )
        if res.data:
            return res.data[0]
    return None


def apply_subscription_webhook(payload: dict[str, Any]) -> dict[str, Any]:
    """
    Lemon Squeezy webhook gövdesini işler; profiles.plan günceller.
    Dönen dict: {ok, event, profile_id, plan, ...}
    """
    meta = payload.get("meta") or {}
    event_name = (meta.get("event_name") or "").strip()
    if event_name not in HANDLED_EVENTS:
        return {"ok": True, "skipped": True, "event": event_name, "reason": "unhandled_event"}

    data = payload.get("data") or {}
    attrs = data.get("attributes") or {}
    custom = meta.get("custom_data") or {}
    if not isinstance(custom, dict):
        custom = {}

    user_id = custom.get("user_id") or custom.get("userId")
    if user_id is not None:
        user_id = str(user_id).strip() or None

    email = (attrs.get("user_email") or attrs.get("customer_email") or "").strip().lower() or None
    customer_id = attrs.get("customer_id")
    if customer_id is not None:
        customer_id = str(customer_id)

    subscription_id = str(data.get("id") or attrs.get("first_subscription_item_id") or "") or None
    status = (attrs.get("status") or "").strip().lower()
    ends_at = attrs.get("ends_at") or attrs.get("renews_at")

    if event_name == "subscription_cancelled" and not status:
        status = "cancelled"

    plan = plan_from_subscription_status(status, ends_at)
    if event_name == "subscription_cancelled":
        # İptal event'inde dönem bitmediyse pro kalabilir
        plan = plan_from_subscription_status(status or "cancelled", ends_at)

    profile = _find_profile_for_webhook(
        user_id=user_id,
        email=email,
        customer_id=customer_id,
    )
    if not profile:
        logger.warning(
            "Webhook: profil bulunamadı event=%s email=%s user_id=%s customer_id=%s",
            event_name,
            email,
            user_id,
            customer_id,
        )
        return {
            "ok": False,
            "error": "profile_not_found",
            "event": event_name,
            "email": email,
            "user_id": user_id,
        }

    update = {
        "plan": plan,
        "subscription_status": status or None,
        "lemon_customer_id": customer_id,
        "lemon_subscription_id": subscription_id,
        "updated_at": datetime.now(timezone.utc).isoformat(),
    }
    if email:
        update["email"] = email

    sb = get_supabase()
    sb.table("profiles").update(update).eq("id", profile["id"]).execute()
    logger.info(
        "Webhook OK event=%s profile=%s plan=%s status=%s",
        event_name,
        profile["id"],
        plan,
        status,
    )
    return {
        "ok": True,
        "event": event_name,
        "profile_id": profile["id"],
        "plan": plan,
        "subscription_status": status,
    }
