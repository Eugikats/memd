"""Azure OpenAI calls for the Excel engine (fill plan + job summary). Same auth as the chat agent."""
from __future__ import annotations

import json
import os

_client = None


def available() -> bool:
    return bool(os.environ.get("AZURE_OPENAI_ENDPOINT"))


def _azure():
    global _client
    if _client is None:
        from openai import AzureOpenAI
        kw = {"azure_endpoint": os.environ["AZURE_OPENAI_ENDPOINT"],
              "api_version": os.environ.get("AZURE_OPENAI_API_VERSION", "2025-04-01-preview"),
              "timeout": 240.0, "max_retries": 1}
        if os.environ.get("AZURE_OPENAI_API_KEY"):
            kw["api_key"] = os.environ["AZURE_OPENAI_API_KEY"]
        else:
            from azure.identity import DefaultAzureCredential, get_bearer_token_provider
            kw["azure_ad_token_provider"] = get_bearer_token_provider(
                DefaultAzureCredential(), "https://cognitiveservices.azure.com/.default")
        _client = AzureOpenAI(**kw)
    return _client


def _deployment():
    return os.environ.get("AZURE_OPENAI_DEPLOYMENT", "gpt-5.4")


def chat_json(system: str, user: str, schema: dict, name: str = "plan", max_tokens: int = 16000) -> dict:
    r = _azure().chat.completions.create(
        model=_deployment(),
        messages=[{"role": "system", "content": system}, {"role": "user", "content": user}],
        response_format={"type": "json_schema", "json_schema": {"name": name, "schema": schema, "strict": True}},
        max_completion_tokens=max_tokens)
    return json.loads(r.choices[0].message.content or "{}")


def chat_text(system: str, user: str, max_tokens: int = 3000) -> str:
    r = _azure().chat.completions.create(
        model=_deployment(),
        messages=[{"role": "system", "content": system}, {"role": "user", "content": user}],
        max_completion_tokens=max_tokens)
    return (r.choices[0].message.content or "").strip()
