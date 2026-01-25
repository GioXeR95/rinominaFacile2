"""Minimal Gemini client using the public REST API.

This module avoids external dependencies (uses urllib) to POST a generateContent
request to Gemini 1.5 Flash. It sends the provided file as inline binary data
alongside a text prompt.
"""

from __future__ import annotations

import base64
import json
import mimetypes
import os
from urllib import request, error


GEMINI_MODEL = "gemini-1.5-flash"
GEMINI_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta"


def _guess_mime_type(file_path: str) -> str:
	mime, _ = mimetypes.guess_type(file_path)
	return mime or "application/octet-stream"


def _build_payload(prompt: str, file_path: str) -> dict:
	mime_type = _guess_mime_type(file_path)
	with open(file_path, "rb") as f:
		data_b64 = base64.b64encode(f.read()).decode("ascii")

	return {
		"contents": [
			{
				"parts": [
					{"text": prompt},
					{
						"inline_data": {
							"mime_type": mime_type,
							"data": data_b64,
						}
					},
				]
			}
		]
	}


def _post_json(url: str, payload: dict) -> dict:
	body = json.dumps(payload).encode("utf-8")
	req = request.Request(url, data=body, headers={"Content-Type": "application/json"})
	with request.urlopen(req, timeout=60) as resp:
		return json.loads(resp.read().decode("utf-8"))


def _extract_text(response_json: dict) -> str:
	# Gemini REST responses structure: candidates -> content -> parts -> text
	candidates = response_json.get("candidates") or []
	if not candidates:
		return ""
	content = candidates[0].get("content") or {}
	parts = content.get("parts") or []
	texts = [p.get("text", "") for p in parts if isinstance(p, dict)]
	return "".join(texts).strip()


def analyze_file_with_gemini(api_key: str, file_path: str, prompt: str) -> tuple[bool, str]:
	"""Send the file and prompt to Gemini and return (success, text_or_error)."""
	if not api_key:
		return False, "Missing Gemini API key"
	if not os.path.exists(file_path):
		return False, "File not found"

	url = f"{GEMINI_ENDPOINT}/models/{GEMINI_MODEL}:generateContent?key={api_key}"

	try:
		payload = _build_payload(prompt, file_path)
		response_json = _post_json(url, payload)
		text = _extract_text(response_json)
		if not text:
			return False, "Empty response from Gemini"
		return True, text
	except error.HTTPError as e:
		try:
			detail = e.read().decode("utf-8")
		except Exception:
			detail = str(e)
		return False, f"Gemini HTTP error: {e.code} - {detail}"
	except Exception as exc:
		return False, f"Gemini request failed: {exc}"


__all__ = ["analyze_file_with_gemini"]
