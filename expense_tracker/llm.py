"""Thin Ollama client. Talks only to the local Ollama HTTP API."""

from __future__ import annotations

import json
from typing import Any

import httpx


class OllamaError(RuntimeError):
    """Raised for any failure talking to Ollama, with a human-readable message."""


class OllamaClient:
    def __init__(self, base_url: str, model: str, timeout_s: float = 120.0):
        self.base_url = base_url.rstrip("/")
        self.model = model
        self.timeout_s = timeout_s

    def health(self) -> list[str]:
        """Return installed model names. Raises OllamaError if the server is unreachable."""
        try:
            r = httpx.get(f"{self.base_url}/api/tags", timeout=5.0)
            r.raise_for_status()
        except httpx.HTTPError as e:
            raise OllamaError(
                f"Cannot reach Ollama at {self.base_url}. Start it with `ollama serve`. ({e})"
            ) from e
        return [m.get("name", "") for m in r.json().get("models", [])]

    def chat_json(self, system: str, user: str, schema: dict[str, Any]) -> dict[str, Any]:
        """Call /api/chat with a JSON schema in `format` and return the parsed JSON object."""
        payload = {
            "model": self.model,
            "stream": False,
            "format": schema,
            "options": {"temperature": 0},
            "messages": [
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
        }
        try:
            r = httpx.post(f"{self.base_url}/api/chat", json=payload, timeout=self.timeout_s)
        except httpx.ConnectError as e:
            raise OllamaError(
                f"Cannot reach Ollama at {self.base_url}. Start it with `ollama serve`."
            ) from e
        except httpx.TimeoutException as e:
            raise OllamaError(
                f"Ollama timed out after {self.timeout_s:.0f}s. The model may still be loading."
            ) from e

        if r.status_code == 404:
            raise OllamaError(f"Model '{self.model}' not found. Run `ollama pull {self.model}`.")
        if r.status_code >= 400:
            raise OllamaError(f"Ollama returned HTTP {r.status_code}: {r.text[:200]}")

        try:
            content = r.json()["message"]["content"]
            return json.loads(content)
        except (KeyError, ValueError, TypeError) as e:
            raise OllamaError(f"Ollama returned unparseable output: {e}") from e
