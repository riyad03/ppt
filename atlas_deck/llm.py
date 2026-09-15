"""Factory for the Qwen model used by the graph.

A single switch point (`LLM_PROVIDER`) between the dev machine (Ollama, local
and free) and a future deployment (OpenAI-compatible endpoint, e.g. vLLM on
Runpod, or OpenRouter's free tier). The rest of the pipeline (`graph.py`)
only knows the standard LangChain interface (`with_structured_output`) and
has nothing to change depending on the provider.
"""

from __future__ import annotations

import os
from typing import Any

from dotenv import load_dotenv

load_dotenv()


def get_llm() -> Any:
    provider = os.getenv("LLM_PROVIDER", "ollama").lower()
    model = os.getenv("QWEN_MODEL", "qwen2.5:7b-instruct")

    # temperature=0 on a local model would make the `plan` node's retry
    # logic pointless: a validation failure would reproduce identically on
    # every retry, with no rewording able to change the outcome.
    temperature = float(os.getenv("QWEN_TEMPERATURE", "0.3"))

    if provider == "ollama":
        from langchain_ollama import ChatOllama

        return ChatOllama(
            model=model,
            base_url=os.getenv("OLLAMA_BASE_URL", "http://localhost:11434"),
            temperature=temperature,
        )

    if provider == "openai_compatible":
        from langchain_openai import ChatOpenAI

        _check_cost_guardrail(model, os.getenv("OPENAI_BASE_URL", ""))

        return ChatOpenAI(
            model=model,
            base_url=os.getenv("OPENAI_BASE_URL"),
            api_key=os.getenv("OPENAI_API_KEY", "not-needed"),
            temperature=temperature,
        )

    raise ValueError(
        f"Unknown LLM_PROVIDER: {provider!r} (expected 'ollama' or 'openai_compatible')"
    )


def _check_cost_guardrail(model: str, base_url: str) -> None:
    """Loss-aversion rule: never a billed call without explicit confirmation.

    A self-hosted Runpod/vLLM endpoint isn't affected (no per-call billing).
    On OpenRouter, a model without the `:free` suffix is paid — refused by
    default, including on a config mistake (e.g. switching provider without
    updating `QWEN_MODEL`, which would otherwise send an Ollama model id
    straight to OpenRouter as-is).
    """
    if "openrouter.ai" not in base_url:
        return
    if model.endswith(":free"):
        return
    if os.getenv("ALLOW_PAID_MODEL", "false").lower() == "true":
        return
    raise RuntimeError(
        f"Cost guardrail: QWEN_MODEL={model!r} on OpenRouter is not a free "
        "model (missing ':free' suffix) and no paid call was confirmed. Use "
        "a ':free' model, or set ALLOW_PAID_MODEL=true in .env if the spend "
        "is accepted."
    )
