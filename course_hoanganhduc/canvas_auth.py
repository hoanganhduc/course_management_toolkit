# -*- coding: utf-8 -*-
# Canvas API: https://canvas.instructure.com/doc/api/

"""Shared Canvas client helpers."""

import re

from canvasapi import Canvas


def validate_canvas_url(api_url):
    if not api_url or not isinstance(api_url, str):
        raise ValueError("Canvas API URL is required.")
    if not re.match(r"^https?://", api_url.strip(), re.IGNORECASE):
        raise ValueError("Canvas API URL must start with http:// or https://")
    return api_url.rstrip("/")


def require_canvas_config(api_url, api_key, course_id):
    api_url = validate_canvas_url(api_url)
    if not api_key:
        raise ValueError("Canvas API key is required.")
    if not course_id:
        raise ValueError("Canvas course ID is required.")
    return api_url, api_key, course_id


def get_canvas_client(api_url, api_key, verbose=False):
    api_url = validate_canvas_url(api_url)
    if not api_key:
        raise ValueError("Canvas API key is required.")
    try:
        return Canvas(api_url, api_key)
    except Exception as exc:
        if verbose:
            print(f"[CanvasAuth] Failed to initialize Canvas client: {exc}")
        raise
