"""
ai_engine.py
------------
All Ollama interaction lives here.
The rest of the codebase does not import requests directly.

Edge cases handled:
  - Ollama server unreachable
  - Ollama returns non-JSON when JSON is expected
  - Model not found (404)
  - Request timeout
  - Malformed response body
"""

import json

import requests

from core import SkuGroup, COL_STOCK, COL_LOCATION


# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

AI_HOST    = "http://10.10.0.130:11434"
DEFAULT_MODEL  = "llama3:latest"
KNOWN_MODELS   = ["llama3:latest", "llama4:latest"]
REQUEST_TIMEOUT = 120  # seconds

SYSTEM_PROMPT = (
    "You are a warehouse inventory analyst. "
    "Provide concise, factual analysis of stock data. "
    "Do not use markdown, bullet points, or headers. "
    "Write in plain paragraphs only."
)


# ---------------------------------------------------------------------------
# Connection check
# ---------------------------------------------------------------------------

class OllamaError(Exception):
    pass


def check_connection() -> tuple[bool, list[str], str]:
    """
    Probe Ollama server.

    Returns:
        (reachable: bool, model_names: list[str], error_message: str)
    """
    try:
        resp = requests.get(f"{AI_HOST}/api/tags", timeout=5)
        resp.raise_for_status()
        data = resp.json()
        models = [m["name"] for m in data.get("models", [])]
        return True, models, ""
    except requests.exceptions.ConnectionError:
        return False, [], f"Cannot connect to Ollama at {AI_HOST}."
    except requests.exceptions.Timeout:
        return False, [], f"Connection to {AI_HOST} timed out."
    except requests.exceptions.HTTPError as e:
        return False, [], f"Ollama server returned error: {e}"
    except Exception as e:
        return False, [], f"Unexpected error checking Ollama: {e}"


# ---------------------------------------------------------------------------
# Low-level chat
# ---------------------------------------------------------------------------

def _chat(model: str, user_prompt: str) -> str:
    """
    Send one user message to Ollama and return the text reply.
    Raises OllamaError on any failure.
    """
    payload = {
        "model": model,
        "stream": False,
        "system": SYSTEM_PROMPT,
        "messages": [{"role": "user", "content": user_prompt}],
    }

    try:
        resp = requests.post(
            f"{AI_HOST}/api/chat",
            json=payload,
            timeout=REQUEST_TIMEOUT,
        )
    except requests.exceptions.ConnectionError:
        raise OllamaError(f"Cannot reach Ollama at {AI_HOST}.")
    except requests.exceptions.Timeout:
        raise OllamaError(f"Ollama request timed out after {REQUEST_TIMEOUT}s.")
    except Exception as e:
        raise OllamaError(f"Request failed: {e}")

    if resp.status_code == 404:
        raise OllamaError(
            f"Model '{model}' not found on server. "
            f"Run: ollama pull {model}"
        )
    if not resp.ok:
        raise OllamaError(
            f"Ollama returned HTTP {resp.status_code}: {resp.text[:200]}"
        )

    try:
        body = resp.json()
    except Exception:
        raise OllamaError("Ollama returned a non-JSON response body.")

    try:
        return body["message"]["content"].strip()
    except (KeyError, TypeError):
        raise OllamaError(f"Unexpected response structure: {str(body)[:200]}")


# ---------------------------------------------------------------------------
# High-level AI functions
# ---------------------------------------------------------------------------

def ai_analyze(model: str, sku_groups: list[SkuGroup]) -> str:
    """
    Ask the model to describe why each SKU needs rebalancing.
    Falls back to an error string on failure — never raises.
    """
    if not sku_groups:
        return "No SKUs required rebalancing."

    data_lines = []
    for g in sku_groups:
        data_lines.append(
            f"SKU '{g.sku_name}': {len(g.locations)} locations, "
            f"stocks {g.old_stocks}, total {g.total}."
        )

    prompt = (
        "Analyze the following SKU stock data from a warehouse inventory system. "
        "For each SKU, explain why the stock distribution is uneven and why "
        "rebalancing across locations is operationally important. "
        "Write 2 to 3 plain sentences per SKU. Do not use bullet points or lists.\n\n"
        + "\n".join(data_lines)
    )

    try:
        return _chat(model, prompt)
    except OllamaError as e:
        return f"AI analysis unavailable: {e}"


def ai_summary(model: str, sku_groups: list[SkuGroup], dry_run: bool) -> str:
    """
    Ask the model to produce a short manager-level summary of the operation.
    Falls back to an error string on failure — never raises.
    """
    total_rows  = sum(len(g.locations) for g in sku_groups)
    total_units = sum(g.total for g in sku_groups)
    mode_str    = "preview (dry run, no file written)" if dry_run else "applied and saved"

    details = []
    for g in sku_groups:
        details.append(
            f"SKU '{g.sku_name}': {g.total} total units across "
            f"{len(g.locations)} locations, rebalanced to {g.new_stocks}."
        )

    prompt = (
        f"Write a 2-sentence operational summary for a warehouse manager. "
        f"Mode: {mode_str}. "
        f"Total SKUs rebalanced: {len(sku_groups)}. "
        f"Total location rows updated: {total_rows}. "
        f"Total stock units redistributed: {total_units:,}. "
        f"Details: {' '.join(details)} "
        f"Be factual, professional, and concise. No bullet points."
    )

    try:
        return _chat(model, prompt)
    except OllamaError as e:
        return f"AI summary unavailable: {e}"