from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Iterable

from .models import APICallUsage


PRICING_SOURCE = "https://developers.openai.com/api/docs/pricing"
PRICING_CHECKED_ON = "2026-09-28"


@dataclass(frozen=True, slots=True)
class TokenPricing:
    input: float
    cached_input: float
    cache_write_input: float
    output: float


# Standard API prices in USD per one million tokens.
MODEL_PRICING = {
    "gpt-6-sol": TokenPricing(2.0, 0.2, 2.5, 10.0),
    "gpt-6-luna": TokenPricing(0.1, 0.01, 0.125, 0.5),
}


def call_usage_from_response(
    response: Any,
    *,
    page_number: int,
    stage: str,
    fallback_model: str,
) -> APICallUsage | None:
    usage = getattr(response, "usage", None)
    if usage is None:
        return None
    input_details = getattr(usage, "input_tokens_details", None)
    output_details = getattr(usage, "output_tokens_details", None)
    return APICallUsage(
        page_number=page_number,
        stage=stage,
        model=str(getattr(response, "model", None) or fallback_model),
        input_tokens=_nonnegative(getattr(usage, "input_tokens", 0)),
        cached_input_tokens=_nonnegative(getattr(input_details, "cached_tokens", 0)),
        cache_write_input_tokens=_nonnegative(
            getattr(input_details, "cache_write_tokens", 0)
        ),
        output_tokens=_nonnegative(getattr(usage, "output_tokens", 0)),
        reasoning_tokens=_nonnegative(getattr(output_details, "reasoning_tokens", 0)),
        total_tokens=_nonnegative(getattr(usage, "total_tokens", 0)),
    )


def estimated_call_cost(call: APICallUsage) -> float | None:
    pricing = _pricing_for_model(call.model)
    if pricing is None:
        return None
    uncached, cached, cache_write = _input_partition(call)
    return (
        uncached * pricing.input
        + cached * pricing.cached_input
        + cache_write * pricing.cache_write_input
        + call.output_tokens * pricing.output
    ) / 1_000_000


def usage_report(
    calls: Iterable[APICallUsage],
    *,
    request_attempts: int | None = None,
) -> dict[str, Any]:
    ordered = sorted(calls, key=lambda call: (call.page_number, _stage_order(call.stage)))
    attempts = max(len(ordered), request_attempts if request_attempts is not None else len(ordered))
    unmetered_attempts = attempts - len(ordered)
    known_costs = [estimated_call_cost(call) for call in ordered]
    known_minimum_cost = sum(cost or 0.0 for cost in known_costs)
    has_unknown_cost = unmetered_attempts > 0 or any(
        cost is None for cost in known_costs
    )
    summary = {
        "api_calls": attempts,
        "metered_responses": len(ordered),
        "unmetered_attempts": unmetered_attempts,
        "input_tokens": sum(call.input_tokens for call in ordered),
        "cached_input_tokens": sum(call.cached_input_tokens for call in ordered),
        "cache_write_input_tokens": sum(
            call.cache_write_input_tokens for call in ordered
        ),
        "output_tokens": sum(call.output_tokens for call in ordered),
        "reasoning_tokens": sum(call.reasoning_tokens for call in ordered),
        "total_tokens": sum(call.total_tokens for call in ordered),
        "estimated_cost_usd": (
            None
            if has_unknown_cost
            else round(known_minimum_cost, 6)
        ),
        "known_minimum_cost_usd": round(known_minimum_cost, 6),
    }
    by_model = []
    for model in sorted({call.model for call in ordered}):
        model_calls = [call for call in ordered if call.model == model]
        model_costs = [estimated_call_cost(call) for call in model_calls]
        by_model.append(
            {
                "model": model,
                "api_calls": len(model_calls),
                "input_tokens": sum(call.input_tokens for call in model_calls),
                "output_tokens": sum(call.output_tokens for call in model_calls),
                "reasoning_tokens": sum(call.reasoning_tokens for call in model_calls),
                "estimated_cost_usd": (
                    None
                    if any(cost is None for cost in model_costs)
                    else round(sum(cost or 0.0 for cost in model_costs), 6)
                ),
            }
        )
    call_rows = []
    for call, cost in zip(ordered, known_costs):
        row = asdict(call)
        uncached, cached, cache_write = _input_partition(call)
        row["uncached_input_tokens"] = uncached
        row["cached_input_tokens"] = cached
        row["cache_write_input_tokens"] = cache_write
        row["estimated_cost_usd"] = None if cost is None else round(cost, 8)
        call_rows.append(row)
    return {
        "pricing": {
            "currency": "USD",
            "basis": "Default service tier, standard API rates per 1M tokens",
            "checked_on": PRICING_CHECKED_ON,
            "source": PRICING_SOURCE,
            "note": (
                "Estimated cost may differ from the billing dashboard because of "
                "regional processing or service tiers. If an attempted request did "
                "not return token usage, estimated_cost_usd is null and "
                "known_minimum_cost_usd covers only metered responses. Reasoning "
                "tokens are already included "
                "in output_tokens and are not charged twice here."
            ),
        },
        "summary": summary,
        "by_model": by_model,
        "calls": call_rows,
    }


def write_usage_report(
    path: Path,
    calls: Iterable[APICallUsage],
    *,
    request_attempts: int | None = None,
) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(
            usage_report(calls, request_attempts=request_attempts),
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    return path


def _pricing_for_model(model: str) -> TokenPricing | None:
    normalized = model.strip().lower()
    for name, pricing in MODEL_PRICING.items():
        if normalized == name or normalized.startswith(f"{name}-"):
            return pricing
    return None


def _stage_order(stage: str) -> int:
    return {"extraction": 0, "verification": 1}.get(stage, 2)


def _input_partition(call: APICallUsage) -> tuple[int, int, int]:
    cached = min(call.input_tokens, call.cached_input_tokens)
    cache_write = min(
        max(0, call.input_tokens - cached),
        call.cache_write_input_tokens,
    )
    uncached = max(0, call.input_tokens - cached - cache_write)
    return uncached, cached, cache_write


def _nonnegative(value: Any) -> int:
    try:
        return max(0, int(value or 0))
    except (TypeError, ValueError):
        return 0
