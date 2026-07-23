"""Telegram usage notifications for the app owner.

Reads bot credentials from Streamlit secrets:

    [telegram]
    bot_token = "123456:ABC..."
    chat_id = "123456789"

Every call fails silently — a broken/missing notification must never
break or slow down file processing for the user.
"""
import requests
import streamlit as st

_TIMEOUT_S = 5


def _config():
    try:
        cfg = st.secrets["telegram"]
        return cfg["bot_token"], cfg["chat_id"]
    except Exception:
        return None, None  # secrets not configured (e.g. local dev)


def notify(text: str) -> None:
    """Send `text` to the owner's Telegram chat. Silent no-op when
    secrets are missing; swallows network errors."""
    token, chat_id = _config()
    if not token or not chat_id:
        return
    try:
        requests.post(
            f"https://api.telegram.org/bot{token}/sendMessage",
            json={"chat_id": chat_id, "text": text},
            timeout=_TIMEOUT_S,
        )
    except Exception:
        pass  # notifications are best-effort only
