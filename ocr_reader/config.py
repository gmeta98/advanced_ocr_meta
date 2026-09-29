from __future__ import annotations

import os
from dataclasses import dataclass, replace
from pathlib import Path

from dotenv import load_dotenv


PROJECT_ROOT = Path(__file__).resolve().parents[1]


@dataclass(frozen=True, slots=True)
class ModelPreset:
    model: str
    reasoning_effort: str


MODEL_PRESETS = {
    "GPT-6 Sol": ModelPreset("gpt-6-sol", "medium"),
    "GPT-6 Luna": ModelPreset("gpt-6-luna", "low"),
}
DEFAULT_MODEL_PRESET = "GPT-6 Sol"


def _resolve_env_file() -> Path | None:
    explicit = _get_setting("OCR_ENV_FILE")
    if explicit:
        candidate = Path(explicit).expanduser()
        if candidate.is_file():
            return candidate

    local = PROJECT_ROOT / ".env"
    if local.is_file():
        return local

    pointer = PROJECT_ROOT / ".env.source"
    if pointer.is_file():
        raw = pointer.read_text(encoding="utf-8").strip()
        if raw:
            candidate = Path(raw).expanduser()
            if candidate.is_file():
                return candidate
    return None


def load_environment() -> Path | None:
    env_file = _resolve_env_file()
    if env_file:
        load_dotenv(env_file, override=False)
    return env_file


@dataclass(frozen=True, slots=True)
class Settings:
    api_key: str
    model: str = "gpt-6-sol"
    reasoning_effort: str = "medium"
    max_workers: int = 2
    render_dpi: int = 260
    verify_ocr: bool = False
    enhance_scans: bool = True
    include_native_text: bool = True

    @classmethod
    def from_environment(cls, *, require_api_key: bool = True) -> "Settings":
        load_environment()
        key = (_get_setting("OPENAI_API_KEY", "") or "").strip()
        if require_api_key and not key:
            raise RuntimeError(
                "OPENAI_API_KEY is missing. Add it to Streamlit app secrets, "
                "a local .env file, or an environment file referenced by .env.source."
            )
        workers = _bounded_int(
            _get_setting("OCR_MAX_WORKERS"), default=2, low=1, high=6
        )
        dpi = _bounded_int(
            _get_setting("OCR_RENDER_DPI"), default=260, low=120, high=450
        )
        effort = (_get_setting("OCR_REASONING_EFFORT", "medium") or "medium").strip().lower()
        if effort not in {"none", "low", "medium", "high", "xhigh", "max"}:
            effort = "medium"
        return cls(
            api_key=key,
            model=(_get_setting("OPENAI_MODEL", "gpt-6-sol") or "gpt-6-sol").strip(),
            reasoning_effort=effort,
            max_workers=workers,
            render_dpi=dpi,
            verify_ocr=_boolean(_get_setting("OCR_VERIFY_PASS"), default=False),
            enhance_scans=_boolean(_get_setting("OCR_ENHANCE_SCANS"), default=True),
            include_native_text=_boolean(
                _get_setting("OCR_INCLUDE_NATIVE_TEXT"), default=True
            ),
        )


def apply_quality_preset(settings: Settings, quality: str) -> Settings:
    normalized = quality.strip().lower()
    if normalized == "balanced":
        return replace(
            settings,
            render_dpi=260,
            verify_ocr=False,
            enhance_scans=True,
        )
    if normalized == "maximum":
        return replace(
            settings,
            render_dpi=360,
            max_workers=1,
            verify_ocr=True,
            enhance_scans=True,
        )
    raise ValueError(f"Unknown OCR quality preset: {quality}")


def _bounded_int(value: str | None, *, default: int, low: int, high: int) -> int:
    try:
        parsed = int(value) if value is not None else default
    except ValueError:
        parsed = default
    return max(low, min(high, parsed))


def _boolean(value: str | None, *, default: bool) -> bool:
    if value is None:
        return default
    normalized = value.strip().lower()
    if normalized in {"1", "true", "yes", "on"}:
        return True
    if normalized in {"0", "false", "no", "off"}:
        return False
    return default


def _get_setting(name: str, default: str | None = None) -> str | None:
    """Read local environment variables or root-level Streamlit Cloud secrets."""
    environment_value = os.getenv(name)
    if environment_value is not None:
        return environment_value

    try:
        import streamlit as st

        if name in st.secrets:
            secret_value = st.secrets[name]
            return default if secret_value is None else str(secret_value)
    except Exception:
        # Streamlit secrets are unavailable during CLI use and local unit tests.
        pass
    return default
