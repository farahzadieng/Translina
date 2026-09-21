#!/usr/bin/env python3
"""Local hybrid document translator.

Run:
    python app.py --config translator.config.json

Required packages:
    flask ctranslate2 transformers sentencepiece llama-cpp-python lxml python-docx pymupdf

The web application is intentionally self-contained: configuration, model
orchestration, document adapters, job persistence, SSE progress, and HTML live
in this file so it can be moved together with ``translator.config.json``.
"""

from __future__ import annotations

import argparse
import copy
import gc
import hashlib
import html
import json
import os
import platform
import queue
import re
import secrets
import shutil
import sqlite3
import subprocess
import sys
import threading
import time
import uuid
import zipfile
from collections.abc import Iterable, Sequence
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

LANGUAGES = {"fa": "فارسی", "en": "English", "ar": "العربية"}
RTL_LANGUAGES = {"fa", "ar"}
APP_VERSION = "3.1.0-direct-gguf-metal"


class TranslatorError(RuntimeError):
    """A user-facing translation failure."""


class JobCancelled(TranslatorError):
    """Raised at a safe checkpoint when a user cancels a job."""


class JobPaused(TranslatorError):
    """Raised at a safe checkpoint after the user requests a pause."""


class ModelCapacityError(TranslatorError):
    """The rendered request cannot fit the model or available memory."""


class ModelLoadError(TranslatorError):
    """The configured model or llama.cpp backend could not be loaded."""


class SegmentIntegrityError(TranslatorError):
    """A protected DOCX formatting boundary was not preserved by the model."""


def _deep_merge(base: dict[str, Any], overlay: dict[str, Any]) -> dict[str, Any]:
    result = copy.deepcopy(base)
    for key, value in overlay.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = _deep_merge(result[key], value)
        else:
            result[key] = copy.deepcopy(value)
    return result


def _expand_value(value: Any, variables: dict[str, str], base_dir: Path) -> Any:
    if isinstance(value, dict):
        return {
            key: _expand_value(item, variables, base_dir) for key, item in value.items()
        }
    if isinstance(value, list):
        return [_expand_value(item, variables, base_dir) for item in value]
    if not isinstance(value, str):
        return value
    expanded = value
    for key, replacement in variables.items():
        expanded = expanded.replace("${" + key + "}", replacement)
    expanded = os.path.expandvars(os.path.expanduser(expanded))
    looks_like_path = (
        expanded.startswith((".", "..", "/", "~"))
        # or "/" in expanded
        or "\\" in expanded
        or expanded.lower().endswith((".gguf", ".db"))
    )
    if looks_like_path and not Path(expanded).is_absolute():
        expanded = str((base_dir / expanded).resolve())
    return expanded


def load_config(path: str | Path) -> dict[str, Any]:
    config_path = Path(path).expanduser().resolve()
    if not config_path.is_file():
        raise TranslatorError(f"Config file not found: {config_path}")
    try:
        raw = json.loads(config_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise TranslatorError(f"Cannot read config file: {exc}") from exc
    if not isinstance(raw, dict):
        raise TranslatorError("The root of the config file must be a JSON object.")

    base_dir = config_path.parent
    configured_root = raw.get("paths", {}).get("model_root", "./models")
    model_root = Path(os.path.expandvars(os.path.expanduser(str(configured_root))))
    if not model_root.is_absolute():
        model_root = (base_dir / model_root).resolve()
    variables = {"MODEL_ROOT": str(model_root), "APP_DIR": str(base_dir)}
    expanded = _expand_value(raw, variables, base_dir)
    expanded.setdefault("paths", {})["model_root"] = str(model_root)
    expanded["_config_path"] = str(config_path)
    expanded["_app_dir"] = str(base_dir)
    return expanded


def total_memory_bytes() -> int:
    """Return physical memory without requiring psutil."""
    try:
        if sys.platform == "darwin":
            return int(
                subprocess.check_output(
                    ["sysctl", "-n", "hw.memsize"], text=True
                ).strip()
            )
        if os.name == "nt":
            import ctypes

            class MemoryStatus(ctypes.Structure):
                _fields_ = [
                    ("length", ctypes.c_ulong),
                    ("memory_load", ctypes.c_ulong),
                    ("total_physical", ctypes.c_ulonglong),
                    ("available_physical", ctypes.c_ulonglong),
                    ("total_page_file", ctypes.c_ulonglong),
                    ("available_page_file", ctypes.c_ulonglong),
                    ("total_virtual", ctypes.c_ulonglong),
                    ("available_virtual", ctypes.c_ulonglong),
                    ("available_extended_virtual", ctypes.c_ulonglong),
                ]

            status = MemoryStatus()
            status.length = ctypes.sizeof(status)
            ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(status))
            return int(status.total_physical)
        pages = os.sysconf("SC_PHYS_PAGES")
        page_size = os.sysconf("SC_PAGE_SIZE")
        return int(pages * page_size)
    except (OSError, ValueError, AttributeError, subprocess.SubprocessError):
        return 8 * 1024**3



def _model_exists(config: dict[str, Any], model_key: str | None) -> bool:
    """Return whether a configured local model exists without loading it."""
    if not model_key:
        return False
    model = config.get("models", {}).get(model_key, {})
    if not isinstance(model, dict):
        return False
    backend = str(model.get("backend", "")).lower()
    if backend == "ctranslate2":
        model_path = Path(str(model.get("path", "")))
        return model_path.is_dir() and (model_path / "model.bin").is_file()
    if backend == "llama_cpp":
        model_path = Path(str(model.get("path", "")))
        return model_path.is_file() and model_path.suffix.lower() == ".gguf"
    return False


def choose_profile(
    config: dict[str, Any], total_memory_bytes: int | None = None
) -> tuple[str, dict[str, Any]]:
    memory = (
        total_memory_bytes
        if total_memory_bytes is not None
        else globals()["total_memory_bytes"]()
    )
    threshold = (
        float(config.get("runtime", {}).get("low_memory_threshold_gb", 12)) * 1024**3
    )
    requested = str(config.get("runtime", {}).get("memory_profile", "auto")).lower()
    if requested in {"low", "high"}:
        desired = requested
    elif sys.platform == "darwin":
        # Unified memory must hold the OS, GGUF weights and KV cache. Be conservative
        # on 16/24 GB Macs; Metal still accelerates the GGUF models in the low profile.
        desired = "low" if memory <= 24 * 1024**3 else "high"
    else:
        desired = "low" if memory <= threshold else "high"

    profiles = config.get("profiles", {})
    profile = copy.deepcopy(profiles.get(desired, {}))
    if not profile:
        raise TranslatorError(f"Runtime profile '{desired}' is not configured.")

    required_roles = ["translator", "analyst", "reviewer"]
    missing = [
        role
        for role in required_roles
        if not _model_exists(config, profile.get(role))
    ]
    if missing and desired == "high":
        desired = "low"
        profile = copy.deepcopy(profiles.get("low", {}))
        missing = [
            role
            for role in required_roles
            if not _model_exists(config, profile.get(role))
        ]
    if missing:
        raise TranslatorError(
            "Selected profile is missing usable models for: " + ", ".join(missing)
        )
    polisher = profile.get("polisher")
    if polisher and not _model_exists(config, polisher):
        profile["polisher"] = None
    return desired, profile


def direction_for(language: str) -> str:
    if language not in LANGUAGES:
        raise TranslatorError(f"Unsupported language: {language}")
    return "rtl" if language in RTL_LANGUAGES else "ltr"


def is_loopback_host(host: str) -> bool:
    normalized = host.strip().lower().strip("[]")
    return normalized in {"localhost", "127.0.0.1", "::1"}



def detect_llama_acceleration() -> str:
    """Return the llama.cpp acceleration mode exposed by the installed build."""
    try:
        from llama_cpp import llama_cpp as llama_backend
    except (ImportError, OSError):
        return "unavailable"
    try:
        gpu = bool(llama_backend.llama_supports_gpu_offload())
    except (AttributeError, OSError):
        gpu = False
    if not gpu:
        return "cpu"
    return "metal" if sys.platform == "darwin" else "gpu-offload"


def detect_runtime_backend() -> str:
    """Report the local inference stack available in this Python environment."""
    try:
        import ctranslate2  # noqa: F401
        import transformers  # noqa: F401
        import sentencepiece  # noqa: F401
        import llama_cpp  # noqa: F401
    except (ImportError, OSError):
        return "incomplete"
    return f"ctranslate2+llama.cpp({detect_llama_acceleration()})"


def output_name(
    filename: str, target_language: str, default_suffix: str | None = None
) -> str:
    safe = Path(filename).name
    path = Path(safe)
    suffix = path.suffix or (default_suffix or "")
    stem = path.name[: -len(path.suffix)] if path.suffix else path.name
    stem = (
        re.sub(r"[^\w. -]+", "_", stem, flags=re.UNICODE).strip(" .") or "translation"
    )
    return f"{stem}.{target_language}{suffix}"


def semantic_split(text: str, max_chars: int) -> list[str]:
    """Split on natural boundaries while preserving the input byte-for-byte."""
    if max_chars < 1:
        raise ValueError("max_chars must be positive")
    if len(text) <= max_chars:
        return [text] if text else []
    boundaries = list(re.finditer(r"(?:\r?\n\s*\r?\n|(?<=[.!?؟؛。！？])\s+|\s+)", text))
    pieces: list[str] = []
    start = 0
    while start < len(text):
        hard_end = min(start + max_chars, len(text))
        if hard_end == len(text):
            pieces.append(text[start:])
            break
        candidates = [
            match.end() for match in boundaries if start < match.end() <= hard_end
        ]
        end = candidates[-1] if candidates else hard_end
        pieces.append(text[start:end])
        start = end
    return pieces


def extract_json_object(raw: str) -> dict[str, Any]:
    if not isinstance(raw, str):
        raise TranslatorError("The model returned a non-text response.")
    decoder = json.JSONDecoder()
    for index, char in enumerate(raw):
        if char != "{":
            continue
        try:
            value, _ = decoder.raw_decode(raw[index:])
        except json.JSONDecodeError:
            continue
        if isinstance(value, dict):
            return value
    raise TranslatorError("The model did not return a valid JSON object.")


def distribute_text(text: str, source_segments: Sequence[str]) -> list[str]:
    """Distribute text across formatting segments without losing characters."""
    count = len(source_segments)
    if count == 0:
        return []
    if count == 1:
        return [text]
    weights = [
        max(1, len(segment.strip()) or len(segment)) for segment in source_segments
    ]
    total = sum(weights)
    cuts: list[int] = []
    previous = 0
    for partial in range(1, count):
        target = round(len(text) * sum(weights[:partial]) / total)
        target = max(previous, min(target, len(text)))
        if 0 < target < len(text):
            right = text.find(" ", target)
            left = text.rfind(" ", previous, target + 1)
            choices = [
                position
                for position in (left, right)
                if previous < position < len(text)
            ]
            if choices:
                target = min(choices, key=lambda position: abs(position - target))
        cuts.append(target)
        previous = target
    result: list[str] = []
    start = 0
    for end in cuts + [len(text)]:
        result.append(text[start:end])
        start = end
    return result


@dataclass
class TranslationUnit:
    unit_id: str
    segments: list[str]
    metadata: dict[str, Any] = field(default_factory=dict)

    @property
    def text(self) -> str:
        return "".join(self.segments)



def _model_manifest_entry(config: dict[str, Any], model_key: str | None) -> dict[str, Any]:
    if not model_key:
        return {"key": None}
    model = copy.deepcopy(config.get("models", {}).get(model_key, {}))
    backend = str(model.get("backend", ""))
    entry: dict[str, Any] = {"key": model_key, "backend": backend}
    if backend == "ctranslate2":
        path = Path(str(model.get("path", "")))
        entry["path"] = str(path)
        stat_path = path / "model.bin"
    elif backend == "llama_cpp":
        path = Path(str(model.get("path", "")))
        entry["path"] = str(path)
        stat_path = path
    else:
        return entry
    try:
        stat = stat_path.stat()
        entry.update({"size": stat.st_size, "mtime_ns": stat.st_mtime_ns})
    except OSError:
        entry.update({"size": -1, "mtime_ns": -1})
    return entry


def build_checkpoint_manifest(
    config: dict[str, Any],
    profile: dict[str, Any],
    units: Sequence[TranslationUnit],
    input_budget: int,
) -> dict[str, Any]:
    unit_hashes = {
        unit.unit_id: hashlib.sha256(
            json.dumps(unit.segments, ensure_ascii=False, separators=(",", ":")).encode(
                "utf-8"
            )
        ).hexdigest()
        for unit in units
    }
    return {
        "version": 3,
        "route": {
            "translator": _model_manifest_entry(config, profile.get("translator")),
            "analyst": _model_manifest_entry(config, profile.get("analyst")),
            "reviewer": _model_manifest_entry(config, profile.get("reviewer")),
            "polisher": _model_manifest_entry(config, profile.get("polisher")),
            "input_budget": int(input_budget),
        },
        "units": unit_hashes,
    }


def reusable_checkpoint_translations(
    state: dict[str, Any], manifest: dict[str, Any]
) -> tuple[
    dict[str, list[str]],
    dict[str, list[str]],
    dict[str, list[str]],
    dict[str, str],
]:
    previous = state.get("manifest")
    if not isinstance(previous, dict) or previous.get("route") != manifest.get("route"):
        return {}, {}, {}, {}
    previous_units = previous.get("units", {})
    current_units = manifest.get("units", {})
    if not isinstance(previous_units, dict) or not isinstance(current_units, dict):
        return {}, {}, {}, {}
    valid_ids = {
        unit_id
        for unit_id, digest in current_units.items()
        if previous_units.get(unit_id) == digest
    }

    def filtered(key: str) -> dict[str, list[str]]:
        values = state.get(key, {})
        if not isinstance(values, dict):
            return {}
        return {
            unit_id: [str(item) for item in value]
            for unit_id, value in values.items()
            if unit_id in valid_ids and isinstance(value, list)
        }

    cache = state.get("mt_cache", {})
    mt_cache = (
        {str(k): str(v) for k, v in cache.items() if isinstance(v, str)}
        if isinstance(cache, dict)
        else {}
    )
    return filtered("drafts"), filtered("finals"), filtered("polished"), mt_cache


def adaptive_batches(
    units: Sequence[TranslationUnit],
    token_counter: Any,
    token_budget: int,
    max_units: int,
) -> list[list[TranslationUnit]]:
    if token_budget < 1 or max_units < 1:
        raise ValueError("token_budget and max_units must be positive")
    batches: list[list[TranslationUnit]] = []
    current: list[TranslationUnit] = []
    current_tokens = 0
    for unit in units:
        unit_tokens = max(1, int(token_counter(unit.text)))
        if current and (
            current_tokens + unit_tokens > token_budget or len(current) >= max_units
        ):
            batches.append(current)
            current = []
            current_tokens = 0
        current.append(unit)
        current_tokens += unit_tokens
    if current:
        batches.append(current)
    return batches


def context_input_budget(
    *,
    context_size: int,
    fixed_tokens: int,
    reserved_output_tokens: int,
    safety_tokens: int,
) -> int:
    available = context_size - fixed_tokens - reserved_output_tokens - safety_tokens
    output_balanced = int(reserved_output_tokens * 0.8)
    budget = min(available, output_balanced)
    if budget < 64:
        raise ModelCapacityError(
            "The configured context window has no safe room for document text."
        )
    return budget


def truncate_to_token_budget(text: str, token_counter: Any, budget: int) -> str:
    if budget < 1:
        return ""
    if int(token_counter(text)) <= budget:
        return text
    low, high = 0, len(text)
    while low < high:
        midpoint = (low + high + 1) // 2
        if int(token_counter(text[:midpoint])) <= budget:
            low = midpoint
        else:
            high = midpoint - 1
    return text[:low]


def fit_previous_context(
    previous: Sequence[dict[str, str]], token_counter: Any, budget: int
) -> list[dict[str, str]]:
    selected: list[dict[str, str]] = []
    for item in reversed(previous):
        candidate = [
            {
                "source": str(item.get("source", "")),
                "target": str(item.get("target", "")),
            }
        ] + selected
        encoded = json.dumps(candidate, ensure_ascii=False)
        if int(token_counter(encoded)) <= budget:
            selected = candidate
            continue
        if selected:
            break
        overhead = int(token_counter(json.dumps([{"source": "", "target": ""}])))
        content_budget = max(0, budget - overhead)
        source = truncate_to_token_budget(
            str(item.get("source", "")), token_counter, content_budget // 2
        )
        target = truncate_to_token_budget(
            str(item.get("target", "")),
            token_counter,
            content_budget - content_budget // 2,
        )
        clipped = [{"source": source, "target": target}]
        while (source or target) and int(
            token_counter(json.dumps(clipped, ensure_ascii=False))
        ) > budget:
            if len(source) >= len(target) and source:
                source = source[:-1]
            elif target:
                target = target[:-1]
            clipped = [{"source": source, "target": target}]
        selected = (
            clipped
            if int(token_counter(json.dumps(clipped, ensure_ascii=False))) <= budget
            else []
        )
        break
    return selected


def _pieces_under_budget(text: str, token_counter: Any, token_budget: int) -> list[str]:
    if not text:
        return [""]
    pending = [text]
    result: list[str] = []
    while pending:
        piece = pending.pop(0)
        tokens = max(1, int(token_counter(piece)))
        if tokens <= token_budget or len(piece) <= 1:
            result.append(piece)
            continue
        max_chars = max(
            1, min(len(piece) - 1, int(len(piece) * token_budget / tokens * 0.88))
        )
        split = semantic_split(piece, max_chars)
        if len(split) == 1 and split[0] == piece:
            midpoint = max(1, len(piece) // 2)
            split = [piece[:midpoint], piece[midpoint:]]
        pending = split + pending
    return result


def prepare_units_for_budget(
    units: Sequence[TranslationUnit], token_counter: Any, token_budget: int
) -> tuple[list[TranslationUnit], dict[str, dict[str, Any]]]:
    """Split an oversized unit without losing its formatting-segment map."""
    prepared: list[TranslationUnit] = []
    mapping: dict[str, dict[str, Any]] = {}
    for unit in units:
        if max(1, int(token_counter(unit.text))) <= token_budget:
            prepared.append(unit)
            mapping[unit.unit_id] = {
                "original_id": unit.unit_id,
                "segment_indices": list(range(len(unit.segments))),
            }
            continue
        fragments: list[tuple[int, str]] = []
        for segment_index, segment in enumerate(unit.segments):
            for piece in _pieces_under_budget(segment, token_counter, token_budget):
                fragments.append((segment_index, piece))
        groups: list[list[tuple[int, str]]] = []
        current: list[tuple[int, str]] = []
        for fragment in fragments:
            proposed = current + [fragment]
            proposed_text = "".join(value for _, value in proposed)
            if current and int(token_counter(proposed_text)) > token_budget:
                groups.append(current)
                current = []
            current.append(fragment)
        if current:
            groups.append(current)
        original_protected = set(unit.metadata.get("protected_segment_indices", []))
        for part_index, group in enumerate(groups):
            virtual_id = f"{unit.unit_id}::part:{part_index}"
            protected_indices = [
                virtual_index
                for virtual_index, (original_index, _) in enumerate(group)
                if original_index in original_protected
            ]
            virtual = TranslationUnit(
                virtual_id,
                [value for _, value in group],
                {
                    **unit.metadata,
                    "original_unit_id": unit.unit_id,
                    "part_index": part_index,
                    "protected_segment_indices": protected_indices,
                },
            )
            prepared.append(virtual)
            mapping[virtual_id] = {
                "original_id": unit.unit_id,
                "segment_indices": [index for index, _ in group],
            }
    return prepared, mapping


def reassemble_split_units(
    originals: Sequence[TranslationUnit],
    prepared: Sequence[TranslationUnit],
    mapping: dict[str, dict[str, Any]],
    translations: dict[str, list[str]],
) -> dict[str, list[str]]:
    result = {unit.unit_id: ["" for _ in unit.segments] for unit in originals}
    for virtual in prepared:
        details = mapping[virtual.unit_id]
        original_id = str(details["original_id"])
        indices = list(details["segment_indices"])
        values = normalize_translation_segments(
            virtual, translations.get(virtual.unit_id, virtual.segments)
        )
        for original_index, value in zip(indices, values):
            result[original_id][int(original_index)] += value
    return result


def normalize_translation_segments(unit: TranslationUnit, value: Any) -> list[str]:
    if isinstance(value, str):
        candidate = [value]
    elif isinstance(value, list):
        candidate = [str(item) for item in value]
    else:
        candidate = []
    if len(candidate) == len(unit.segments):
        return candidate
    if unit.metadata.get("protected_segment_indices"):
        raise SegmentIntegrityError(
            f"Protected formatting segments changed for unit {unit.unit_id}."
        )
    joined = "".join(candidate).strip()
    if not joined:
        raise TranslatorError(f"No translation was returned for unit {unit.unit_id}.")
    return distribute_text(joined, unit.segments)


_NUMBER_TRANSLATION = str.maketrans(
    "٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹−٬٫⁄：",
    "01234567890123456789-,./:",
)
_NUMBER_PATTERN = re.compile(
    r"(?<!\w)[+\-−]?(?:[0-9٠-٩۰-۹][0-9٠-٩۰-۹,٬.٫/⁄:：\-]*[0-9٠-٩۰-۹]|[0-9٠-٩۰-۹])(?!\w)"
)


def canonical_numbers(text: str) -> list[str]:
    """Return comparable numbers while accepting localized Persian/Arabic digits."""
    return [match.group(0).translate(_NUMBER_TRANSLATION) for match in _NUMBER_PATTERN.finditer(text)]



def quality_flags(
    source: str, target: str, target_language: str | None = None
) -> list[str]:
    flags: list[str] = []
    if not target.strip():
        return ["empty"]
    source_numbers = canonical_numbers(source)
    target_numbers = canonical_numbers(target)
    if sorted(source_numbers) != sorted(target_numbers):
        flags.append("numbers")
    source_urls = re.findall(r"https?://[^\s<>()]+", source)
    target_urls = re.findall(r"https?://[^\s<>()]+", target)
    if source_urls != target_urls:
        flags.append("urls")
    source_emails = re.findall(r"[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}", source)
    target_emails = re.findall(r"[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}", target)
    if source_emails != target_emails:
        flags.append("emails")
    if re.search(r"ZXQPH[A-Z]+QXZ", target):
        flags.append("placeholder")
    ratio = len(target.strip()) / max(1, len(source.strip()))
    if len(source.strip()) > 80 and (ratio < 0.28 or ratio > 3.8):
        flags.append("length")
    if target_language in {"fa", "ar"} and len(source) > 100:
        latin_words = re.findall(r"\b[A-Za-z]{4,}\b", target)
        source_latin = re.findall(r"\b[A-Za-z]{4,}\b", source)
        if len(latin_words) >= max(8, int(len(source_latin) * 0.65)):
            flags.append("source_leakage")
    return flags


def glossary_flags(source: str, target: str, analysis: dict[str, Any]) -> list[str]:
    flags: list[str] = []
    source_folded = source.casefold()
    target_folded = target.casefold()
    terms = analysis.get("terms", []) if isinstance(analysis, dict) else []
    if not isinstance(terms, list):
        return flags
    for term in terms:
        if not isinstance(term, dict):
            continue
        # Terms inferred by a small model are useful guidance, not ground truth.
        # Only an explicitly user-locked glossary entry may block delivery.
        if term.get("locked") is not True:
            continue
        source_term = str(term.get("source", "")).strip()
        target_term = str(term.get("target", "")).strip()
        if (
            len(source_term) >= 2
            and target_term
            and source_term.casefold() in source_folded
            and target_term.casefold() not in target_folded
        ):
            flags.append(f"terminology:{source_term}→{target_term}")
    return flags



class ModelManager:
    """Local runtime: CTranslate2 for MADLAD and llama.cpp for GGUF LLMs.

    Exactly one GGUF model is resident at a time. This is important on Apple
    Silicon unified memory and on smaller NVIDIA cards.
    """

    def __init__(
        self,
        config: dict[str, Any],
        profile: dict[str, Any],
        status_callback: Any | None = None,
    ):
        self.config = config
        self.profile = profile
        self.status_callback = status_callback or (lambda *args: None)
        self._lock = threading.RLock()
        self._mt_models: dict[str, Any] = {}
        self._tokenizers: dict[str, Any] = {}
        self._llm: Any = None
        self._llm_key: str | None = None

    def _notify(self, state: str, model_key: str) -> None:
        self.status_callback(state, model_key)

    def _model_config(self, model_key: str) -> dict[str, Any]:
        model = self.config.get("models", {}).get(model_key)
        if not isinstance(model, dict):
            raise ModelLoadError(f"Model '{model_key}' is not configured.")
        return model

    def _load_tokenizer(self, model_key: str) -> Any:
        with self._lock:
            if model_key in self._tokenizers:
                return self._tokenizers[model_key]
            model = self._model_config(model_key)
            if str(model.get("backend", "")).lower() != "ctranslate2":
                return None
            path = Path(str(model.get("tokenizer_path") or model.get("path") or ""))
            if not path.is_dir():
                raise ModelLoadError(f"Tokenizer directory not found: {path}")
            try:
                from transformers import AutoTokenizer
            except ImportError as exc:
                raise ModelLoadError(
                    "transformers is not installed. Install transformers and sentencepiece."
                ) from exc
            try:
                tokenizer = AutoTokenizer.from_pretrained(
                    str(path), local_files_only=True, use_fast=True
                )
            except Exception:
                try:
                    tokenizer = AutoTokenizer.from_pretrained(
                        str(path), local_files_only=True, use_fast=False
                    )
                except Exception as exc:
                    raise ModelLoadError(
                        f"Could not load tokenizer for '{model_key}': {exc}"
                    ) from exc
            self._tokenizers[model_key] = tokenizer
            return tokenizer

    @staticmethod
    def _resolve_ct2_runtime(model: dict[str, Any]) -> tuple[str, str]:
        requested = str(model.get("device", "auto")).lower()
        compute = str(model.get("compute_type", "auto"))
        if requested != "auto":
            return requested, compute
        # CTranslate2 prebuilt GPU support is CUDA/NVIDIA. Apple Silicon uses ARM64 CPU.
        if sys.platform == "darwin":
            return "cpu", str(model.get("cpu_compute_type", "int8"))
        try:
            import ctranslate2
            cuda_types = ctranslate2.get_supported_compute_types("cuda")
            if cuda_types:
                preferred = str(model.get("cuda_compute_type", "int8_float16"))
                if preferred in cuda_types:
                    return "cuda", preferred
                if "float16" in cuda_types:
                    return "cuda", "float16"
                return "cuda", sorted(cuda_types)[0]
        except Exception:
            pass
        return "cpu", str(model.get("cpu_compute_type", "int8"))

    def _load_mt(self, model_key: str) -> Any:
        with self._lock:
            if model_key in self._mt_models:
                return self._mt_models[model_key]
            model = self._model_config(model_key)
            if str(model.get("backend", "")).lower() != "ctranslate2":
                raise ModelLoadError(f"Model '{model_key}' is not a CTranslate2 model.")
            path = Path(str(model.get("path", "")))
            if not path.is_dir() or not (path / "model.bin").is_file():
                raise ModelLoadError(f"CTranslate2 model directory not found: {path}")
            self._notify("loading", model_key)
            try:
                import ctranslate2
            except ImportError as exc:
                self._notify("error", model_key)
                raise ModelLoadError("ctranslate2 is not installed.") from exc
            self._load_tokenizer(model_key)
            device, compute_type = self._resolve_ct2_runtime(model)
            kwargs = {
                "device": device,
                "compute_type": compute_type,
                "inter_threads": int(model.get("inter_threads", 1)),
                "intra_threads": int(
                    model.get("intra_threads", max(1, (os.cpu_count() or 4) - 1))
                ),
            }
            try:
                translator = ctranslate2.Translator(str(path), **kwargs)
            except Exception as first_exc:
                if device == "cuda" and bool(model.get("allow_cpu_fallback", True)):
                    try:
                        translator = ctranslate2.Translator(
                            str(path),
                            device="cpu",
                            compute_type=str(model.get("cpu_compute_type", "int8")),
                            inter_threads=1,
                            intra_threads=max(1, (os.cpu_count() or 4) - 1),
                        )
                    except Exception as second_exc:
                        self._notify("error", model_key)
                        raise ModelLoadError(
                            f"Could not load MT model on GPU ({first_exc}) or CPU ({second_exc})."
                        ) from second_exc
                else:
                    self._notify("error", model_key)
                    raise ModelLoadError(
                        f"Could not load MT model '{model_key}': {first_exc}"
                    ) from first_exc
            self._mt_models[model_key] = translator
            self._notify("loaded", model_key)
            return translator

    @staticmethod
    def _gpu_layers_for(model: dict[str, Any]) -> int:
        configured = model.get("gpu_layers", "auto")
        if configured != "auto":
            return int(configured)
        return -1 if detect_llama_acceleration() in {"metal", "gpu-offload"} else 0

    def _load_llm(self, model_key: str) -> Any:
        with self._lock:
            if self._llm is not None and self._llm_key == model_key:
                return self._llm
            self._unload_llm_locked()
            model = self._model_config(model_key)
            if str(model.get("backend", "")).lower() != "llama_cpp":
                raise ModelLoadError(f"Model '{model_key}' is not configured for llama.cpp.")
            path = Path(str(model.get("path", "")))
            if not path.is_file():
                raise ModelLoadError(f"GGUF model file not found: {path}")
            self._notify("loading", model_key)
            try:
                from llama_cpp import Llama
            except (ImportError, OSError) as exc:
                self._notify("error", model_key)
                raise ModelLoadError(
                    "llama-cpp-python is not installed or could not be loaded."
                ) from exc

            threads = model.get("threads", "auto")
            n_threads = (
                max(1, (os.cpu_count() or 4) - 2)
                if threads == "auto"
                else max(1, int(threads))
            )
            kwargs: dict[str, Any] = {
                "model_path": str(path),
                "n_ctx": int(model.get("context_size", 8192)),
                "n_batch": int(model.get("batch_size", 512)),
                "n_threads": n_threads,
                "n_gpu_layers": self._gpu_layers_for(model),
                "use_mmap": bool(model.get("use_mmap", True)),
                "use_mlock": bool(model.get("use_mlock", False)),
                "verbose": bool(model.get("verbose", False)),
            }
            if model.get("chat_format"):
                kwargs["chat_format"] = str(model["chat_format"])
            try:
                self._llm = Llama(**kwargs)
            except Exception as exc:
                self._llm = None
                self._llm_key = None
                self._notify("error", model_key)
                hint = ""
                if sys.platform == "darwin":
                    hint = " Verify that llama-cpp-python was built with Metal support."
                raise ModelLoadError(
                    f"Could not load GGUF model '{model_key}': {exc}.{hint}"
                ) from exc
            self._llm_key = model_key
            self._notify("loaded", model_key)
            return self._llm

    def count_tokens(self, model_key: str, text: str) -> int:
        model = self._model_config(model_key)
        backend = str(model.get("backend", "")).lower()
        if backend == "ctranslate2":
            tokenizer = self._load_tokenizer(model_key)
            return len(tokenizer.encode(text, add_special_tokens=True))
        if backend == "llama_cpp":
            llm = self._load_llm(model_key)
            try:
                return len(llm.tokenize(text.encode("utf-8"), add_bos=False))
            except TypeError:
                return len(llm.tokenize(text.encode("utf-8")))
        return max(1, int(len(text) / 2.6) + 8)

    def source_token_limit(self, model_key: str) -> int:
        model = self._model_config(model_key)
        return max(64, int(model.get("max_source_tokens", 420)))

    def translate_texts(
        self,
        model_key: str,
        texts: Sequence[str],
        target_language: str,
    ) -> list[str]:
        if not texts:
            return []
        model = self._model_config(model_key)
        if str(model.get("backend", "")).lower() != "ctranslate2":
            raise TranslatorError("The translation role must use a CTranslate2 model.")
        translator = self._load_mt(model_key)
        tokenizer = self._load_tokenizer(model_key)
        target_tags = model.get("target_tags", {})
        target_code = str(target_tags.get(target_language, target_language))
        token_batches: list[list[str]] = []
        limit = self.source_token_limit(model_key)
        for text in texts:
            prompt = f"<2{target_code}> {text}"
            ids = tokenizer.encode(prompt, add_special_tokens=True)
            if len(ids) > limit:
                raise ModelCapacityError(
                    f"MT input is {len(ids)} tokens, above configured limit {limit}."
                )
            token_batches.append(tokenizer.convert_ids_to_tokens(ids))
        try:
            results = translator.translate_batch(
                token_batches,
                beam_size=max(1, int(model.get("beam_size", 2))),
                max_decoding_length=max(64, int(model.get("max_decoding_length", 640))),
                repetition_penalty=float(model.get("repetition_penalty", 1.05)),
                no_repeat_ngram_size=max(0, int(model.get("no_repeat_ngram_size", 0))),
                batch_type="tokens",
                max_batch_size=max(1, int(model.get("max_batch_size", 2048))),
            )
        except Exception as exc:
            message = str(exc).lower()
            error_type = (
                ModelCapacityError
                if any(x in message for x in ("memory", "allocate", "cuda", "length"))
                else TranslatorError
            )
            raise error_type(f"CTranslate2 translation failed: {exc}") from exc
        outputs: list[str] = []
        for result in results:
            tokens = result.hypotheses[0]
            ids = tokenizer.convert_tokens_to_ids(tokens)
            value = tokenizer.decode(ids, skip_special_tokens=True).strip()
            outputs.append(value)
        return outputs

    def generate_json(
        self,
        model_key: str,
        messages: list[dict[str, str]],
        max_tokens: int,
        temperature: float,
    ) -> dict[str, Any]:
        model = self._model_config(model_key)
        if str(model.get("backend", "")).lower() != "llama_cpp":
            raise TranslatorError("LLM analysis/review roles must use a GGUF llama.cpp model.")
        llm = self._load_llm(model_key)
        kwargs: dict[str, Any] = {
            "messages": messages,
            "max_tokens": int(max_tokens),
            "temperature": float(temperature),
            "top_p": float(model.get("top_p", 0.9)),
            "repeat_penalty": float(model.get("repeat_penalty", 1.03)),
        }
        try:
            try:
                response = llm.create_chat_completion(
                    **kwargs, response_format={"type": "json_object"}
                )
            except (TypeError, ValueError):
                response = llm.create_chat_completion(**kwargs)
            choice = response["choices"][0]
            if str(choice.get("finish_reason", "")).lower() == "length":
                raise ModelCapacityError(
                    f"GGUF model '{model_key}' reached the output token limit."
                )
            content = str(choice.get("message", {}).get("content") or "")
            if not content.strip():
                raise TranslatorError(f"GGUF model '{model_key}' returned empty output.")
            return extract_json_object(content)
        except (TranslatorError, ModelCapacityError):
            raise
        except Exception as exc:
            message = str(exc).lower()
            error_type = (
                ModelCapacityError
                if any(x in message for x in ("context", "kv cache", "memory", "allocate", "oom"))
                else TranslatorError
            )
            raise error_type(f"GGUF generation failed ({model_key}): {exc}") from exc

    def _unload_llm_locked(self) -> None:
        model = self._llm
        key = self._llm_key
        self._llm = None
        self._llm_key = None
        if model is not None:
            try:
                close = getattr(model, "close", None)
                if callable(close):
                    close()
            finally:
                del model
                if key:
                    self._notify("unloaded", key)
            gc.collect()

    def unload(self, model_key: str | None = None) -> None:
        with self._lock:
            mt_keys = list(self._mt_models) if model_key is None else [model_key]
            for key in mt_keys:
                translator = self._mt_models.pop(key, None)
                if translator is not None:
                    del translator
                    self._notify("unloaded", key)
            if model_key is None or self._llm_key == model_key:
                self._unload_llm_locked()
            gc.collect()


TONE_GUIDANCE = {
    "auto": "Infer and preserve the document's appropriate register.",
    "formal": "Use polished, formal, publication-ready prose.",
    "legal": "Use precise legal language; preserve ambiguity only when it exists in the source.",
    "business": "Use clear, concise, professional business language.",
    "academic": "Use disciplined academic language and stable technical terminology.",
    "literary": "Preserve imagery, rhythm, voice, and emotional effect without adding ideas.",
    "conversational": "Use natural contemporary speech appropriate to the target language.",
    "screenplay": "Write speakable dialogue; preserve character voice, rhythm, subtext, and scene directions.",
}



class TranslationPipeline:
    """Specialized MT -> semantic post-edit -> deterministic QA -> Persian polish."""

    IMMUTABLE_PATTERN = re.compile(
        r"https?://[^\s<>()]+|[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}|"
        r"(?<!\w)[+\-−]?(?:[0-9٠-٩۰-۹][0-9٠-٩۰-۹,٬.٫/⁄:：\-]*[0-9٠-٩۰-۹]|[0-9٠-٩۰-۹])(?!\w)"
    )

    def __init__(
        self,
        config: dict[str, Any],
        profile: dict[str, Any],
        models: ModelManager,
        progress: Any | None = None,
        cancelled: Any | None = None,
        paused: Any | None = None,
        checkpoint: Any | None = None,
    ):
        self.config = config
        self.profile = profile
        self.models = models
        self.progress = progress or (lambda *args, **kwargs: None)
        self.cancelled = cancelled or (lambda: False)
        self.paused = paused or (lambda: False)
        self.checkpoint = checkpoint or (lambda *args, **kwargs: None)
        self.options = config.get("translation", {})
        self.quality_warnings: dict[str, list[str]] = {}
        self._polish_hints: set[str] = set()

    def _check_cancelled(self) -> None:
        if self.cancelled():
            raise JobCancelled("Translation cancelled.")
        if self.paused():
            raise JobPaused("Translation paused at a safe checkpoint.")

    def _llm_json_call(
        self,
        model_key: str,
        messages: list[dict[str, str]],
        max_tokens: int,
        temperature: float,
    ) -> dict[str, Any]:
        retries = max(0, int(self.options.get("max_retries", 2)))
        last_error: Exception | None = None
        for attempt in range(retries + 1):
            self._check_cancelled()
            try:
                return self.models.generate_json(
                    model_key,
                    messages,
                    max_tokens=max_tokens,
                    temperature=temperature,
                )
            except (TranslatorError, ModelLoadError) as exc:
                last_error = exc
                if attempt >= retries:
                    break
                messages = list(messages) + [
                    {
                        "role": "user",
                        "content": "Return exactly one valid JSON object and no markdown.",
                    }
                ]
        raise TranslatorError(f"Structured model call failed: {last_error}")

    def _representative_sample(self, units: Sequence[TranslationUnit]) -> str:
        nonempty = [unit.text for unit in units if unit.text.strip()]
        if not nonempty:
            return ""
        picks: list[str] = []
        indices = list(range(min(4, len(nonempty))))
        if len(nonempty) > 8:
            middle = len(nonempty) // 2
            indices.extend([max(0, middle - 1), middle])
        indices.extend(range(max(0, len(nonempty) - 4), len(nonempty)))
        seen: set[int] = set()
        for index in indices:
            if index not in seen:
                picks.append(nonempty[index])
                seen.add(index)
        sample = "\n\n---\n\n".join(picks)
        return sample[: max(2000, int(self.options.get("analysis_max_chars", 14000)))]

    def _analyze(
        self,
        units: Sequence[TranslationUnit],
        source_language: str,
        target_language: str,
        tone: str,
    ) -> dict[str, Any]:
        model_key = str(self.profile.get("analyst") or self.profile.get("reviewer"))
        sample = self._representative_sample(units)
        payload = {
            "source_language": LANGUAGES[source_language],
            "target_language": LANGUAGES[target_language],
            "requested_tone": tone,
            "document_sample": sample,
        }
        system = (
            "You are a senior translation analyst. Do not translate the document. "
            "Return JSON with keys domain, audience, summary, style, entities, and terms. "
            "terms must be an array of {source,target,locked:false}. Include only recurring "
            "multiword concepts whose translation is reasonably stable; prefer omission to guessing. "
            "For Persian target text, choose natural Persian terminology, not word-for-word calques."
        )
        data = self._llm_json_call(
            model_key,
            [
                {"role": "system", "content": system},
                {"role": "user", "content": json.dumps(payload, ensure_ascii=False)},
            ],
            max_tokens=min(1200, int(self.options.get("reserved_output_tokens", 1800))),
            temperature=0.05,
        )
        if not isinstance(data.get("terms"), list):
            data["terms"] = []
        return data

    def _locked_terms(self, analysis: dict[str, Any]) -> list[dict[str, str]]:
        terms: list[dict[str, str]] = []
        configured = self.options.get("locked_glossary", [])
        for raw in [
            *(configured if isinstance(configured, list) else []),
            *(analysis.get("terms", []) if isinstance(analysis.get("terms"), list) else []),
        ]:
            if not isinstance(raw, dict) or raw.get("locked") is not True:
                continue
            source = str(raw.get("source", "")).strip()
            target = str(raw.get("target", "")).strip()
            if source and target:
                terms.append({"source": source, "target": target})
        unique: dict[str, dict[str, str]] = {}
        for term in terms:
            unique[term["source"].casefold()] = term
        return list(unique.values())

    @staticmethod
    def _alpha_id(index: int) -> str:
        letters = []
        value = index
        while True:
            letters.append(chr(ord("A") + value % 26))
            value = value // 26 - 1
            if value < 0:
                return "".join(reversed(letters))

    def _protect_text(
        self, text: str, locked_terms: Sequence[dict[str, str]]
    ) -> tuple[str, dict[str, str]]:
        mapping: dict[str, str] = {}
        counter = [0]

        def placeholder(value: str) -> str:
            token = f"ZXQPH{self._alpha_id(counter[0])}QXZ"
            counter[0] += 1
            mapping[token] = value
            return token

        protected = self.IMMUTABLE_PATTERN.sub(lambda m: placeholder(m.group(0)), text)
        for term in sorted(locked_terms, key=lambda item: len(item["source"]), reverse=True):
            pattern = re.compile(re.escape(term["source"]), flags=re.IGNORECASE)
            protected = pattern.sub(lambda _m, target=term["target"]: placeholder(target), protected)
        return protected, mapping

    @staticmethod
    def _restore_text(text: str, mapping: dict[str, str]) -> str:
        restored = text
        for token, value in mapping.items():
            restored = restored.replace(token, value)
            restored = restored.replace(token.lower(), value)
        return restored

    def _mt_cache_key(
        self, text: str, target_language: str, locked_terms: Sequence[dict[str, str]]
    ) -> str:
        payload = {
            "text": text,
            "target": target_language,
            "terms": list(locked_terms),
        }
        return hashlib.sha256(
            json.dumps(payload, ensure_ascii=False, sort_keys=True).encode("utf-8")
        ).hexdigest()

    def _translate_long_text(
        self,
        model_key: str,
        protected_text: str,
        target_language: str,
    ) -> str:
        limit = self.models.source_token_limit(model_key)
        counter = lambda value: self.models.count_tokens(
            model_key, f"<2{target_language}> {value}"
        )
        pieces = _pieces_under_budget(protected_text, counter, max(64, limit - 12))
        outputs: list[str] = []
        batch_size = max(1, int(self.options.get("mt_long_piece_batch", 8)))
        for start in range(0, len(pieces), batch_size):
            self._check_cancelled()
            outputs.extend(
                self.models.translate_texts(
                    model_key, pieces[start : start + batch_size], target_language
                )
            )
        return " ".join(value.strip() for value in outputs if value.strip())

    def _translate_work_items(
        self,
        model_key: str,
        items: Sequence[tuple[str, str, int | None]],
        target_language: str,
        locked_terms: Sequence[dict[str, str]],
        mt_cache: dict[str, str],
    ) -> dict[tuple[str, int | None], str]:
        results: dict[tuple[str, int | None], str] = {}
        pending: list[tuple[str, int | None, str, dict[str, str], str]] = []
        limit = self.models.source_token_limit(model_key)
        for unit_id, text, segment_index in items:
            if not text.strip():
                results[(unit_id, segment_index)] = text
                continue
            cache_key = self._mt_cache_key(text, target_language, locked_terms)
            if cache_key in mt_cache:
                results[(unit_id, segment_index)] = mt_cache[cache_key]
                continue
            protected, mapping = self._protect_text(text, locked_terms)
            count = self.models.count_tokens(model_key, f"<2{target_language}> {protected}")
            if count > limit:
                translated = self._translate_long_text(
                    model_key, protected, target_language
                )
                translated = self._restore_text(translated, mapping)
                mt_cache[cache_key] = translated
                results[(unit_id, segment_index)] = translated
            else:
                pending.append((unit_id, segment_index, protected, mapping, cache_key))

        batch_size = max(1, int(self.options.get("mt_batch_items", 12)))
        for start in range(0, len(pending), batch_size):
            self._check_cancelled()
            block = pending[start : start + batch_size]
            translated = self.models.translate_texts(
                model_key, [item[2] for item in block], target_language
            )
            for item, value in zip(block, translated):
                unit_id, segment_index, _protected, mapping, cache_key = item
                restored = self._restore_text(value, mapping)
                mt_cache[cache_key] = restored
                results[(unit_id, segment_index)] = restored
        return results

    def _translate_batch_mt(
        self,
        model_key: str,
        batch: Sequence[TranslationUnit],
        target_language: str,
        locked_terms: Sequence[dict[str, str]],
        mt_cache: dict[str, str],
    ) -> dict[str, list[str]]:
        work: list[tuple[str, str, int | None]] = []
        protected_units: set[str] = set()
        for unit in batch:
            protected = bool(unit.metadata.get("protected_segment_indices"))
            if protected and len(unit.segments) > 1:
                protected_units.add(unit.unit_id)
                for index, segment in enumerate(unit.segments):
                    work.append((unit.unit_id, segment, index))
            else:
                work.append((unit.unit_id, unit.text, None))
        values = self._translate_work_items(
            model_key, work, target_language, locked_terms, mt_cache
        )
        result: dict[str, list[str]] = {}
        for unit in batch:
            if unit.unit_id in protected_units:
                result[unit.unit_id] = [
                    values.get((unit.unit_id, index), segment)
                    for index, segment in enumerate(unit.segments)
                ]
            else:
                value = values.get((unit.unit_id, None), unit.text)
                result[unit.unit_id] = (
                    [value]
                    if len(unit.segments) == 1
                    else distribute_text(value, unit.segments)
                )
        return result

    def _review_batch(
        self,
        model_key: str,
        batch: Sequence[TranslationUnit],
        drafts: dict[str, list[str]],
        source_language: str,
        target_language: str,
        tone: str,
        analysis: dict[str, Any],
        pre_flags: dict[str, list[str]],
        custom_instruction: str,
    ) -> tuple[dict[str, list[str]], set[str]]:
        locked = self._locked_terms(analysis)
        payload = {
            "source_language": LANGUAGES[source_language],
            "target_language": LANGUAGES[target_language],
            "tone": TONE_GUIDANCE.get(tone, tone),
            "document_profile": {
                "domain": analysis.get("domain", ""),
                "audience": analysis.get("audience", ""),
                "style": analysis.get("style", ""),
                "summary": analysis.get("summary", ""),
            },
            "locked_glossary": locked,
            "user_instruction": custom_instruction.strip(),
            "units": [
                {
                    "id": unit.unit_id,
                    "source": unit.segments,
                    "draft": drafts[unit.unit_id],
                    "qa_reasons": pre_flags.get(unit.unit_id, []),
                    "protected_segment_indices": unit.metadata.get(
                        "protected_segment_indices", []
                    ),
                }
                for unit in batch
            ],
        }
        system = (
            "You are the semantic post-editor of a professional translation system. SOURCE is the "
            "absolute authority; the machine-translation draft is only a proposal. Compare source and "
            "draft carefully. If the draft is already correct and natural, keep it with minimal changes. "
            "Fix mistranslation, omission, addition, negation, modality, terminology, idioms, register, "
            "and unnatural Persian. Never invent information. Preserve every number, URL, email, proper "
            "name, factual relation, and formatting segment count. Enforce locked_glossary exactly. "
            "Return JSON: {\"translations\": {id: [same number of segments]}, \"polish\": [ids]}. "
            "Put an id in polish only when the meaning is correct but Persian wording would benefit from "
            "a Persian-specialist stylistic pass. No explanations or markdown."
        )
        data = self._llm_json_call(
            model_key,
            [
                {"role": "system", "content": system},
                {"role": "user", "content": json.dumps(payload, ensure_ascii=False)},
            ],
            max_tokens=int(self.options.get("reserved_output_tokens", 1800)),
            temperature=float(self.options.get("editor_temperature", 0.05)),
        )
        translations = data.get("translations", {})
        if not isinstance(translations, dict):
            raise TranslatorError("Semantic reviewer returned no translations object.")
        result: dict[str, list[str]] = {}
        for unit in batch:
            candidate = translations.get(unit.unit_id, drafts[unit.unit_id])
            try:
                result[unit.unit_id] = normalize_translation_segments(unit, candidate)
            except SegmentIntegrityError:
                result[unit.unit_id] = drafts[unit.unit_id]
        polish_raw = data.get("polish", [])
        polish = {
            str(item)
            for item in polish_raw
            if isinstance(polish_raw, list) and str(item) in {u.unit_id for u in batch}
        }
        return result, polish

    def _polish_batch(
        self,
        model_key: str,
        batch: Sequence[TranslationUnit],
        candidates: dict[str, list[str]],
        source_language: str,
        target_language: str,
        tone: str,
        analysis: dict[str, Any],
        reasons: dict[str, list[str]],
    ) -> dict[str, list[str]]:
        payload = {
            "source_language": LANGUAGES[source_language],
            "target_language": LANGUAGES[target_language],
            "tone": TONE_GUIDANCE.get(tone, tone),
            "locked_glossary": self._locked_terms(analysis),
            "user_instruction": str(analysis.get("user_instruction", "")),
            "units": [
                {
                    "id": unit.unit_id,
                    "source": unit.segments,
                    "translation": candidates[unit.unit_id],
                    "reason": reasons.get(unit.unit_id, []),
                    "protected_segment_indices": unit.metadata.get(
                        "protected_segment_indices", []
                    ),
                }
                for unit in batch
            ],
        }
        system = (
            "You are a Persian language specialist performing the final linguistic polish. Do not "
            "re-translate from scratch. Preserve the source meaning and every fact exactly, but make the "
            "Persian read as fluent, contemporary, publication-ready native prose appropriate to the "
            "requested tone. Fix calques, awkward word order, collocations, punctuation and register. "
            "Never change numbers, names, URLs, emails, terminology or segment counts. Return only JSON "
            "{\"translations\": {id: [segments]}}."
        )
        data = self._llm_json_call(
            model_key,
            [
                {"role": "system", "content": system},
                {"role": "user", "content": json.dumps(payload, ensure_ascii=False)},
            ],
            max_tokens=int(self.options.get("reserved_output_tokens", 1800)),
            temperature=float(self.options.get("polisher_temperature", 0.08)),
        )
        translations = data.get("translations", {})
        if not isinstance(translations, dict):
            return {unit.unit_id: candidates[unit.unit_id] for unit in batch}
        result: dict[str, list[str]] = {}
        for unit in batch:
            value = translations.get(unit.unit_id, candidates[unit.unit_id])
            try:
                result[unit.unit_id] = normalize_translation_segments(unit, value)
            except SegmentIntegrityError:
                result[unit.unit_id] = candidates[unit.unit_id]
        return result

    def _collect_issues(
        self,
        units: Sequence[TranslationUnit],
        translations: dict[str, list[str]],
        analysis: dict[str, Any],
        target_language: str,
    ) -> dict[str, list[str]]:
        issues: dict[str, list[str]] = {}
        qa_analysis = dict(analysis)
        analysis_terms = analysis.get("terms", []) if isinstance(analysis.get("terms"), list) else []
        configured_terms = self.options.get("locked_glossary", [])
        if not isinstance(configured_terms, list):
            configured_terms = []
        qa_analysis["terms"] = [*analysis_terms, *configured_terms]
        for unit in units:
            target = "".join(translations.get(unit.unit_id, []))
            flags = [
                *quality_flags(unit.text, target, target_language),
                *glossary_flags(unit.text, target, qa_analysis),
            ]
            if flags:
                issues[unit.unit_id] = list(dict.fromkeys(flags))
        return issues

    def _repair_batch(
        self,
        model_key: str,
        batch: Sequence[TranslationUnit],
        candidates: dict[str, list[str]],
        source_language: str,
        target_language: str,
        analysis: dict[str, Any],
        flags: dict[str, list[str]],
    ) -> dict[str, list[str]]:
        payload = {
            "source_language": LANGUAGES[source_language],
            "target_language": LANGUAGES[target_language],
            "locked_glossary": self._locked_terms(analysis),
            "units": [
                {
                    "id": unit.unit_id,
                    "source": unit.segments,
                    "candidate": candidates[unit.unit_id],
                    "failed_checks": flags.get(unit.unit_id, []),
                }
                for unit in batch
            ],
        }
        system = (
            "Repair only the listed deterministic failures. SOURCE is authoritative. Preserve natural "
            "target-language prose, but restore missing or changed numbers, URLs, emails, locked terms "
            "or omitted content. Do not make unrelated stylistic rewrites. Return JSON "
            "{\"translations\": {id: [same number of segments]}}."
        )
        data = self._llm_json_call(
            model_key,
            [
                {"role": "system", "content": system},
                {"role": "user", "content": json.dumps(payload, ensure_ascii=False)},
            ],
            max_tokens=int(self.options.get("reserved_output_tokens", 1800)),
            temperature=0.0,
        )
        translations = data.get("translations", {})
        result: dict[str, list[str]] = {}
        for unit in batch:
            value = translations.get(unit.unit_id, candidates[unit.unit_id]) if isinstance(translations, dict) else candidates[unit.unit_id]
            try:
                result[unit.unit_id] = normalize_translation_segments(unit, value)
            except SegmentIntegrityError:
                result[unit.unit_id] = candidates[unit.unit_id]
        return result

    def run(
        self,
        units: Sequence[TranslationUnit],
        source_language: str,
        target_language: str,
        tone: str,
        custom_instruction: str = "",
        resume: dict[str, Any] | None = None,
    ) -> dict[str, list[str]]:
        if source_language == target_language:
            raise TranslatorError("Source and target languages must be different.")
        if source_language not in LANGUAGES or target_language not in LANGUAGES:
            raise TranslatorError("Unsupported source or target language.")
        nonempty = [unit for unit in units if unit.text.strip()]
        if not nonempty:
            return {unit.unit_id: list(unit.segments) for unit in units}

        state = resume if isinstance(resume, dict) else {}
        translator_key = str(self.profile["translator"])
        reviewer_key = str(self.profile["reviewer"])
        analyst_key = str(self.profile.get("analyst") or reviewer_key)
        polisher_key = self.profile.get("polisher")
        quality_mode = str(self.options.get("quality_mode", "balanced")).lower()
        if quality_mode not in {"fast", "balanced", "strict"}:
            quality_mode = "balanced"

        self.progress("analysis", 5, "تحلیل حوزه، لحن و اصطلاحات سند با مدل بازبین")
        analysis = state.get("analysis") if isinstance(state.get("analysis"), dict) else None
        if not analysis:
            analysis = self._analyze(nonempty, source_language, target_language, tone)
        if custom_instruction.strip():
            analysis["user_instruction"] = custom_instruction.strip()
        self.checkpoint("analysis", {"analysis": analysis})
        # Do not keep Gemma resident while CTranslate2 loads MADLAD.
        self.models.unload(analyst_key)

        # The MT model itself has a short encoder context (MADLAD is normally 512 tokens),
        # so batching is by units while oversized paragraphs are split internally.
        budget = self.models.source_token_limit(translator_key)
        working_units = nonempty
        manifest = build_checkpoint_manifest(self.config, self.profile, working_units, budget)
        drafts, finals, polished, mt_cache = reusable_checkpoint_translations(state, manifest)
        self.checkpoint(
            "manifest",
            {
                "analysis": analysis,
                "manifest": manifest,
                "drafts": drafts,
                "finals": finals,
                "polished": polished,
                "mt_cache": mt_cache,
            },
        )

        locked_terms = self._locked_terms(analysis)
        mt_unit_batch = max(1, int(self.options.get("mt_units_per_batch", 10)))
        mt_batches = [
            working_units[i : i + mt_unit_batch]
            for i in range(0, len(working_units), mt_unit_batch)
        ]
        for index, batch in enumerate(mt_batches):
            self._check_cancelled()
            pending = [unit for unit in batch if unit.unit_id not in drafts]
            if pending:
                drafts.update(
                    self._translate_batch_mt(
                        translator_key,
                        pending,
                        target_language,
                        locked_terms,
                        mt_cache,
                    )
                )
                self.checkpoint(
                    "draft",
                    {
                        "analysis": analysis,
                        "manifest": manifest,
                        "drafts": drafts,
                        "mt_cache": mt_cache,
                    },
                )
            self.progress(
                "translation",
                12 + int(38 * (index + 1) / max(1, len(mt_batches))),
                f"ترجمه تخصصی MADLAD: بخش {index + 1} از {len(mt_batches)}",
            )

        # Free the large translation model before loading Gemma through llama.cpp.
        self.models.unload(translator_key)
        pre_flags = self._collect_issues(working_units, drafts, analysis, target_language)
        self.progress("review", 52, "بازبینی معنایی ترجمه با Gemma 3")

        reviewer_context = int(
            self.config.get("models", {}).get(reviewer_key, {}).get("context_size", 32768)
        )
        review_budget = max(800, int(reviewer_context * 0.34))
        review_batches = adaptive_batches(
            working_units,
            lambda text: self.models.count_tokens(reviewer_key, text) * 2,
            review_budget,
            max(1, int(self.profile.get("max_units_per_request", 6))),
        )
        for index, batch in enumerate(review_batches):
            self._check_cancelled()
            if quality_mode == "fast":
                pending = [
                    unit
                    for unit in batch
                    if unit.unit_id not in finals and unit.unit_id in pre_flags
                ]
                untouched = [
                    unit
                    for unit in batch
                    if unit.unit_id not in finals and unit.unit_id not in pre_flags
                ]
                for unit in untouched:
                    finals[unit.unit_id] = drafts[unit.unit_id]
            else:
                pending = [unit for unit in batch if unit.unit_id not in finals]
            if pending:
                edited, polish_ids = self._review_batch(
                    reviewer_key,
                    pending,
                    drafts,
                    source_language,
                    target_language,
                    tone,
                    analysis,
                    pre_flags,
                    custom_instruction,
                )
                finals.update(edited)
                self._polish_hints.update(polish_ids)
                self.checkpoint(
                    "review",
                    {
                        "analysis": analysis,
                        "manifest": manifest,
                        "drafts": drafts,
                        "finals": finals,
                        "polish_hints": sorted(self._polish_hints),
                        "mt_cache": mt_cache,
                    },
                )
            self.progress(
                "review",
                52 + int(23 * (index + 1) / max(1, len(review_batches))),
                f"بازبینی معنایی: بخش {index + 1} از {len(review_batches)}",
            )

        # Resume can restore polish hints generated in a previous attempt.
        previous_hints = state.get("polish_hints", [])
        if isinstance(previous_hints, list):
            self._polish_hints.update(str(x) for x in previous_hints)

        after_review_flags = self._collect_issues(
            working_units, finals, analysis, target_language
        )
        selected_for_polish: set[str] = set(after_review_flags)
        selected_for_polish.update(self._polish_hints)
        if quality_mode == "strict" and target_language == "fa":
            # In strict mode Dorna sees every reviewed unit, but it still receives the
            # reviewed translation rather than translating from scratch.
            selected_for_polish.update(unit.unit_id for unit in working_units)

        if target_language == "fa" and polisher_key and selected_for_polish:
            self.models.unload(reviewer_key)
            polish_units = [
                unit for unit in working_units if unit.unit_id in selected_for_polish
            ]
            polish_batches = adaptive_batches(
                polish_units,
                lambda text: self.models.count_tokens(str(polisher_key), text) * 2,
                max(700, int(self.config.get("models", {}).get(str(polisher_key), {}).get("context_size", 8192) * 0.30)),
                max(1, int(self.profile.get("polisher_max_units", 4))),
            )
            for index, batch in enumerate(polish_batches):
                self._check_cancelled()
                pending = [unit for unit in batch if unit.unit_id not in polished]
                if pending:
                    polished.update(
                        self._polish_batch(
                            str(polisher_key),
                            pending,
                            finals,
                            source_language,
                            target_language,
                            tone,
                            analysis,
                            after_review_flags,
                        )
                    )
                    self.checkpoint(
                        "polish",
                        {
                            "analysis": analysis,
                            "manifest": manifest,
                            "drafts": drafts,
                            "finals": finals,
                            "polished": polished,
                            "polish_hints": sorted(self._polish_hints),
                            "mt_cache": mt_cache,
                        },
                    )
                self.progress(
                    "polish",
                    76 + int(12 * (index + 1) / max(1, len(polish_batches))),
                    f"روان‌سازی فارسی با Dorna: بخش {index + 1} از {len(polish_batches)}",
                )
            self.models.unload(str(polisher_key))

        combined = dict(finals)
        combined.update(polished)
        unresolved = self._collect_issues(
            working_units, combined, analysis, target_language
        )

        # Only deterministic integrity problems are auto-repaired. Stylistic flags
        # remain warnings and never prevent the user from receiving a file.
        blocking_codes = {"empty", "numbers", "urls", "emails", "placeholder"}
        repair_flags = {
            unit_id: [flag for flag in flags if flag.split(":", 1)[0] in blocking_codes]
            for unit_id, flags in unresolved.items()
        }
        repair_flags = {k: v for k, v in repair_flags.items() if v}
        repair_attempts = max(0, int(self.options.get("quality_repair_attempts", 1)))
        for attempt in range(repair_attempts):
            if not repair_flags:
                break
            self.models.unload(str(polisher_key) if polisher_key else None)
            repair_units = [unit for unit in working_units if unit.unit_id in repair_flags]
            repair_batches = adaptive_batches(
                repair_units,
                lambda text: self.models.count_tokens(reviewer_key, text) * 2,
                review_budget,
                max(1, int(self.profile.get("max_units_per_request", 6))),
            )
            for batch in repair_batches:
                repaired = self._repair_batch(
                    reviewer_key,
                    batch,
                    combined,
                    source_language,
                    target_language,
                    analysis,
                    repair_flags,
                )
                combined.update(repaired)
            unresolved = self._collect_issues(
                working_units, combined, analysis, target_language
            )
            repair_flags = {
                unit_id: [
                    flag
                    for flag in flags
                    if flag.split(":", 1)[0] in blocking_codes
                ]
                for unit_id, flags in unresolved.items()
            }
            repair_flags = {k: v for k, v in repair_flags.items() if v}
            self.progress(
                "quality",
                89 + min(4, attempt + 1),
                f"کنترل قطعی و ترمیم داده‌ها: تلاش {attempt + 1}",
            )

        self.quality_warnings = unresolved
        self.progress(
            "quality",
            94,
            (
                f"کنترل نهایی تمام شد؛ {sum(len(v) for v in unresolved.values())} هشدار باقی مانده"
                if unresolved
                else "کنترل نهایی بدون هشدار قطعی پایان یافت"
            ),
        )

        result = {unit.unit_id: list(unit.segments) for unit in units}
        result.update(combined)
        return result


class TextArtifact:
    kind = "text"

    def __init__(self, text: str, original_name: str = "translation.txt"):
        self.original_name = original_name
        self._parts: list[tuple[str, str]] = []
        self.units: list[TranslationUnit] = []
        for index, part in enumerate(re.split(r"(\r?\n(?:[ \t]*\r?\n)+)", text)):
            if not part:
                continue
            if re.fullmatch(r"\r?\n(?:[ \t]*\r?\n)+", part):
                self._parts.append(("literal", part))
            else:
                unit_id = f"text-{index}"
                self.units.append(TranslationUnit(unit_id, [part], {"kind": "text"}))
                self._parts.append(("unit", unit_id))

    def build(
        self,
        translations: dict[str, list[str]],
        output_path: str | Path,
        target_language: str,
        config: dict[str, Any],
    ) -> Path:
        del target_language, config
        chunks = []
        for kind, value in self._parts:
            chunks.append(
                value if kind == "literal" else "".join(translations.get(value, []))
            )
        output = Path(output_path)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text("".join(chunks), encoding="utf-8")
        return output


@dataclass
class _DocxParagraphRef:
    entry_name: str
    paragraph: Any
    text_nodes: list[Any]


class DocxArtifact:
    kind = "docx"
    W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    XML_NS = "http://www.w3.org/XML/1998/namespace"
    MAX_ZIP_ENTRIES = 10_000
    MAX_UNCOMPRESSED_BYTES = 512 * 1024 * 1024
    MAX_XML_BYTES = 64 * 1024 * 1024
    MAX_COMPRESSION_RATIO = 200

    def __init__(self, source_path: str | Path):
        try:
            from lxml import etree
        except ImportError as exc:
            raise TranslatorError("lxml is required for DOCX processing.") from exc
        self.etree = etree
        self.source_path = Path(source_path)
        if not self.source_path.is_file():
            raise TranslatorError(f"DOCX file not found: {self.source_path}")
        try:
            with zipfile.ZipFile(self.source_path) as package:
                self._infos = package.infolist()
                names = {info.filename for info in self._infos}
                required = {"[Content_Types].xml", "_rels/.rels", "word/document.xml"}
                if not required.issubset(names):
                    raise TranslatorError(
                        "The DOCX package is missing required OOXML parts."
                    )
                if len(self._infos) > self.MAX_ZIP_ENTRIES:
                    raise TranslatorError(
                        "The DOCX package contains too many ZIP entries."
                    )
                total_uncompressed = sum(info.file_size for info in self._infos)
                if total_uncompressed > self.MAX_UNCOMPRESSED_BYTES:
                    raise TranslatorError("The uncompressed DOCX package is too large.")
                for info in self._infos:
                    if (
                        info.filename.lower().endswith(".xml")
                        and info.file_size > self.MAX_XML_BYTES
                    ):
                        raise TranslatorError(
                            f"DOCX XML part is too large: {info.filename}"
                        )
                    ratio = info.file_size / max(1, info.compress_size)
                    if (
                        info.file_size > 1024 * 1024
                        and ratio > self.MAX_COMPRESSION_RATIO
                    ):
                        raise TranslatorError(
                            f"Suspicious DOCX compression ratio in {info.filename}."
                        )
                self._entries = {
                    info.filename: package.read(info.filename) for info in self._infos
                }
        except (OSError, zipfile.BadZipFile) as exc:
            raise TranslatorError(f"Invalid DOCX package: {exc}") from exc
        if "word/document.xml" not in self._entries:
            raise TranslatorError("The file is not a valid DOCX document.")
        self._roots: dict[str, Any] = {}
        self._refs: dict[str, _DocxParagraphRef] = {}
        self.units: list[TranslationUnit] = []
        candidates = [
            name
            for name in self._entries
            if re.fullmatch(
                r"word/(?:document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml",
                name,
            )
        ]
        paragraph_tag = self._qn("p")
        text_tag = self._qn("t")
        for entry_name in candidates:
            try:
                parser = etree.XMLParser(
                    resolve_entities=False,
                    no_network=True,
                    huge_tree=False,
                    recover=False,
                )
                root = etree.fromstring(self._entries[entry_name], parser=parser)
            except etree.XMLSyntaxError as exc:
                raise TranslatorError(f"Cannot parse {entry_name}: {exc}") from exc
            self._roots[entry_name] = root
            for paragraph_index, paragraph in enumerate(root.iter(paragraph_tag)):
                text_nodes = []
                for node in paragraph.iter(text_tag):
                    nearest = next(
                        (ancestor for ancestor in node.iterancestors(paragraph_tag)),
                        None,
                    )
                    if nearest is paragraph:
                        text_nodes.append(node)
                segments = [node.text or "" for node in text_nodes]
                if not text_nodes or not "".join(segments).strip():
                    continue
                protected_tags = {
                    self._qn("b"),
                    self._qn("i"),
                    self._qn("u"),
                    self._qn("color"),
                    self._qn("highlight"),
                    self._qn("vertAlign"),
                    self._qn("strike"),
                    self._qn("dstrike"),
                    self._qn("caps"),
                    self._qn("smallCaps"),
                }
                protected_indices: list[int] = []
                for segment_index, node in enumerate(text_nodes):
                    run = next(
                        (ancestor for ancestor in node.iterancestors(self._qn("r"))),
                        None,
                    )
                    hyperlink = next(
                        (
                            ancestor
                            for ancestor in node.iterancestors(self._qn("hyperlink"))
                        ),
                        None,
                    )
                    run_properties = (
                        run.find(self._qn("rPr")) if run is not None else None
                    )
                    styled = bool(
                        run_properties is not None
                        and any(child.tag in protected_tags for child in run_properties)
                    )
                    if hyperlink is not None or styled:
                        protected_indices.append(segment_index)
                unit_id = f"docx-{len(self.units)}"
                metadata = {
                    "entry": entry_name,
                    "paragraph_index": paragraph_index,
                    "protected_segment_indices": protected_indices,
                }
                self.units.append(TranslationUnit(unit_id, segments, metadata))
                self._refs[unit_id] = _DocxParagraphRef(
                    entry_name, paragraph, text_nodes
                )

    def _qn(self, local: str) -> str:
        return f"{{{self.W_NS}}}{local}"

    def _ensure_first(self, parent: Any, tag: str) -> Any:
        child = parent.find(tag)
        if child is None:
            child = self.etree.Element(tag)
            parent.insert(0, child)
        return child

    def _set_direction(
        self, paragraph: Any, target_language: str, fallback_font: str
    ) -> None:
        rtl = target_language in RTL_LANGUAGES
        ppr = self._ensure_first(paragraph, self._qn("pPr"))
        bidi = ppr.find(self._qn("bidi"))
        if rtl and bidi is None:
            bidi = self.etree.SubElement(ppr, self._qn("bidi"))
        if bidi is not None:
            bidi.set(self._qn("val"), "1" if rtl else "0")
        alignment = ppr.find(self._qn("jc"))
        if alignment is not None:
            value = alignment.get(self._qn("val"), "")
            if rtl and value == "left":
                alignment.set(self._qn("val"), "right")
            elif not rtl and value == "right":
                alignment.set(self._qn("val"), "left")
        table = next(
            (ancestor for ancestor in paragraph.iterancestors(self._qn("tbl"))), None
        )
        if table is not None:
            table_properties = self._ensure_first(table, self._qn("tblPr"))
            table_bidi = table_properties.find(self._qn("bidiVisual"))
            if table_bidi is None:
                table_bidi = self.etree.SubElement(
                    table_properties, self._qn("bidiVisual")
                )
            table_bidi.set(self._qn("val"), "1" if rtl else "0")
        for run in paragraph.iter(self._qn("r")):
            nearest = next(
                (ancestor for ancestor in run.iterancestors(self._qn("p"))), None
            )
            if nearest is not paragraph:
                continue
            rpr = self._ensure_first(run, self._qn("rPr"))
            rtl_node = rpr.find(self._qn("rtl"))
            if rtl and rtl_node is None:
                rtl_node = self.etree.SubElement(rpr, self._qn("rtl"))
            if rtl_node is not None:
                rtl_node.set(self._qn("val"), "1" if rtl else "0")
            lang = rpr.find(self._qn("lang"))
            if lang is None:
                lang = self.etree.SubElement(rpr, self._qn("lang"))
            lang.set(
                self._qn("val"),
                {"fa": "fa-IR", "ar": "ar-SA", "en": "en-US"}[target_language],
            )
            if fallback_font:
                fonts = rpr.find(self._qn("rFonts"))
                if fonts is None:
                    fonts = self.etree.Element(self._qn("rFonts"))
                    rpr.insert(0, fonts)
                if rtl:
                    fonts.set(self._qn("cs"), fallback_font)
                    fonts.set(self._qn("eastAsia"), fallback_font)
                elif not fonts.get(self._qn("ascii")):
                    fonts.set(self._qn("ascii"), fallback_font)
                    fonts.set(self._qn("hAnsi"), fallback_font)

    def build(
        self,
        translations: dict[str, list[str]],
        output_path: str | Path,
        target_language: str,
        config: dict[str, Any],
    ) -> Path:
        fallback = str(
            config.get("documents", {})
            .get("fallback_fonts", {})
            .get(target_language, "")
        )
        for unit in self.units:
            ref = self._refs[unit.unit_id]
            segments = normalize_translation_segments(
                unit, translations.get(unit.unit_id, unit.segments)
            )
            for node, text in zip(ref.text_nodes, segments):
                node.text = text
                if text.startswith(" ") or text.endswith(" "):
                    node.set(f"{{{self.XML_NS}}}space", "preserve")
                elif f"{{{self.XML_NS}}}space" in node.attrib:
                    del node.attrib[f"{{{self.XML_NS}}}space"]
            self._set_direction(ref.paragraph, target_language, fallback)
        entries = dict(self._entries)
        for entry_name, root in self._roots.items():
            entries[entry_name] = self.etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True
            )
        output = Path(output_path)
        output.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output, "w") as package:
            for info in self._infos:
                package.writestr(info, entries[info.filename])
        try:
            with zipfile.ZipFile(output) as check:
                if check.testzip() is not None:
                    raise TranslatorError("The generated DOCX package is corrupt.")
        except zipfile.BadZipFile as exc:
            raise TranslatorError(f"The generated DOCX is invalid: {exc}") from exc
        return output


class PdfArtifact:
    kind = "pdf"

    def __init__(self, source_path: str | Path):
        try:
            import pymupdf as fitz
        except ImportError as exc:
            raise TranslatorError("PyMuPDF is required for PDF processing.") from exc
        self.fitz = fitz
        self.source_path = Path(source_path)
        if not self.source_path.is_file():
            raise TranslatorError(f"PDF file not found: {self.source_path}")
        try:
            document = fitz.open(self.source_path)
        except Exception as exc:
            raise TranslatorError(f"Cannot open PDF: {exc}") from exc
        self.units: list[TranslationUnit] = []
        try:
            if document.needs_pass:
                raise TranslatorError(
                    "Password-protected PDFs are not supported in this version."
                )
            for page_index, page in enumerate(document):
                content = page.get_text("dict", flags=fitz.TEXTFLAGS_TEXT)
                for block_index, block in enumerate(content.get("blocks", [])):
                    if block.get("type") != 0:
                        continue
                    lines: list[str] = []
                    representative: dict[str, Any] = {}
                    for line in block.get("lines", []):
                        line_text = "".join(
                            str(span.get("text", "")) for span in line.get("spans", [])
                        )
                        if line_text:
                            lines.append(line_text)
                        if not representative and line.get("spans"):
                            representative = line["spans"][0]
                    text = "\n".join(lines)
                    if not text.strip():
                        continue
                    metadata = {
                        "page": page_index,
                        "block": block_index,
                        "bbox": list(
                            block.get(
                                "bbox",
                                (36, 36, page.rect.width - 36, page.rect.height - 36),
                            )
                        ),
                        "font": representative.get("font", "sans-serif"),
                        "size": float(representative.get("size", 11)),
                        "color": int(representative.get("color", 0)),
                    }
                    unit_id = f"pdf-{page_index}-{block_index}"
                    self.units.append(TranslationUnit(unit_id, [text], metadata))
        finally:
            document.close()
        if not self.units:
            raise TranslatorError(
                "This PDF has no selectable text. Scanned PDFs are not supported."
            )

    @staticmethod
    def _normalized_text(value: str) -> str:
        return "".join(
            character.casefold() for character in value if character.isalnum()
        )

    def _entry_html(
        self,
        unit: TranslationUnit,
        text: str,
        target_language: str,
        fallback: str,
        include_block_spacing: bool = True,
    ) -> str:
        size = max(7.0, float(unit.metadata.get("size", 11)))
        color = f"#{int(unit.metadata.get('color', 0)) & 0xFFFFFF:06x}"
        direction = direction_for(target_language)
        align = "right" if direction == "rtl" else "left"
        source_font = re.sub(r"^[A-Z]{6}\+", "", str(unit.metadata.get("font", "")))
        family = (
            fallback
            if direction == "rtl" or fallback == "LocalTargetFont"
            else (source_font or fallback)
        )
        family = family.replace("'", "")
        margin = "0 0 .7em 0" if include_block_spacing else "0"
        escaped = html.escape(text).replace("\n", "<br>")
        return (
            f"<div style=\"font-family:'{family}',sans-serif;font-size:{size}pt;color:{color};"
            f'direction:{direction};text-align:{align};line-height:1.3;margin:{margin}">'
            f"{escaped}</div>"
        )

    @staticmethod
    def _html_fits(
        page: Any,
        rect: Any,
        markup: str,
        min_scale: float,
        css: str = "",
        archive: Any | None = None,
    ) -> bool:
        try:
            kwargs: dict[str, Any] = {"scale_low": min_scale}
            if css:
                kwargs.update({"css": css, "archive": archive})
            result = page.insert_htmlbox(rect, markup, **kwargs)
        except (RuntimeError, ValueError):
            return False
        return bool(result and result[0] >= 0)

    def _trial_group_fits(
        self,
        width: float,
        height: float,
        markup: str,
        min_scale: float,
        css: str = "",
        archive: Any | None = None,
    ) -> bool:
        trial = self.fitz.open()
        try:
            page = trial.new_page(width=width, height=height)
            rect = self.fitz.Rect(36, 36, width - 36, height - 36)
            return self._html_fits(page, rect, markup, min_scale, css, archive)
        finally:
            trial.close()

    def _reflow_entries(
        self,
        output_doc: Any,
        source_page: Any,
        first_page: Any,
        entries: list[tuple[TranslationUnit, str]],
        target_language: str,
        fallback: str,
        min_scale: float,
        css: str = "",
        archive: Any | None = None,
    ) -> None:
        width, height = source_page.rect.width, source_page.rect.height
        readable_scale = max(0.85, min_scale)
        fragments: list[tuple[TranslationUnit, str]] = []
        pending = list(entries)
        while pending:
            unit, text = pending.pop(0)
            markup = self._entry_html(unit, text, target_language, fallback)
            if self._trial_group_fits(
                width, height, markup, readable_scale, css, archive
            ):
                fragments.append((unit, text))
                continue
            if len(text) <= 1:
                raise TranslatorError(
                    f"PDF reflow could not place translation unit {unit.unit_id}."
                )
            pieces = semantic_split(text, max(1, len(text) // 2))
            if len(pieces) == 1:
                midpoint = max(1, len(text) // 2)
                pieces = [text[:midpoint], text[midpoint:]]
            pending = [(unit, piece) for piece in pieces if piece] + pending

        groups: list[list[tuple[TranslationUnit, str]]] = []
        current: list[tuple[TranslationUnit, str]] = []
        for entry in fragments:
            proposed = current + [entry]
            markup = "".join(
                self._entry_html(unit, text, target_language, fallback)
                for unit, text in proposed
            )
            if current and not self._trial_group_fits(
                width, height, markup, readable_scale, css, archive
            ):
                groups.append(current)
                current = []
            current.append(entry)
        if current:
            groups.append(current)

        for group_index, group in enumerate(groups):
            page = (
                first_page
                if group_index == 0
                else output_doc.new_page(width=width, height=height)
            )
            rect = self.fitz.Rect(36, 36, width - 36, height - 36)
            markup = "".join(
                self._entry_html(unit, text, target_language, fallback)
                for unit, text in group
            )
            if not self._html_fits(
                page, rect, markup, readable_scale, css, archive
            ):
                raise TranslatorError(
                    "PDF reflow validation failed while placing translated text."
                )

    def build(
        self,
        translations: dict[str, list[str]],
        output_path: str | Path,
        target_language: str,
        config: dict[str, Any],
    ) -> Path:
        fitz = self.fitz
        source = fitz.open(self.source_path)
        output_doc = fitz.open()
        fallback = str(
            config.get("documents", {})
            .get("fallback_fonts", {})
            .get(target_language, "sans-serif")
        )
        font_css = ""
        font_archive: Any | None = None
        font_value = (
            config.get("documents", {}).get("font_files", {}).get(target_language)
        )
        if font_value:
            font_path = Path(str(font_value))
            if not font_path.is_file():
                raise TranslatorError(f"Configured font file not found: {font_path}")
            try:
                font_archive = fitz.Archive(str(font_path.parent))
            except Exception as exc:
                raise TranslatorError(f"Cannot open configured font: {exc}") from exc
            font_filename = font_path.name.replace("'", "")
            font_css = (
                "@font-face{font-family:'LocalTargetFont';"
                f"src:url('{font_filename}');}}"
                "*{font-family:'LocalTargetFont'!important;}"
            )
        min_scale = float(config.get("documents", {}).get("pdf_min_scale", 0.62))
        by_page: dict[int, list[TranslationUnit]] = {}
        for unit in self.units:
            by_page.setdefault(int(unit.metadata["page"]), []).append(unit)

        for page_index, source_page in enumerate(source):
            output_doc.insert_pdf(source, from_page=page_index, to_page=page_index)
            page = output_doc[-1]
            units = by_page.get(page_index, [])
            entries = [
                (unit, "".join(translations.get(unit.unit_id, unit.segments)))
                for unit in units
            ]

            trial = fitz.open()
            trial_page = trial.new_page(
                width=source_page.rect.width, height=source_page.rect.height
            )
            fixed_layout_fits = True
            for unit, text in entries:
                rect = fitz.Rect(unit.metadata["bbox"])
                size = max(7.0, float(unit.metadata.get("size", 11)))
                rect.y1 = min(
                    trial_page.rect.y1 - 4, max(rect.y1, rect.y0 + size * 1.8)
                )
                markup = self._entry_html(
                    unit,
                    text,
                    target_language,
                    fallback,
                    include_block_spacing=False,
                )
                if not self._html_fits(
                    trial_page,
                    rect,
                    markup,
                    min_scale,
                    font_css,
                    font_archive,
                ):
                    fixed_layout_fits = False
                    break
            trial.close()

            for unit in units:
                rect = fitz.Rect(unit.metadata["bbox"])
                page.add_redact_annot(rect, fill=(1, 1, 1))
            try:
                page.apply_redactions(images=0, graphics=0)
            except TypeError:
                page.apply_redactions(images=0)

            if fixed_layout_fits:
                for unit, text in entries:
                    rect = fitz.Rect(unit.metadata["bbox"])
                    size = max(7.0, float(unit.metadata.get("size", 11)))
                    rect.y1 = min(page.rect.y1 - 4, max(rect.y1, rect.y0 + size * 1.8))
                    markup = self._entry_html(
                        unit,
                        text,
                        target_language,
                        fallback,
                        include_block_spacing=False,
                    )
                    if not self._html_fits(
                        page,
                        rect,
                        markup,
                        min_scale,
                        font_css,
                        font_archive,
                    ):
                        raise TranslatorError(
                            f"PDF placement changed after preflight for unit {unit.unit_id}."
                        )
            elif entries:
                self._reflow_entries(
                    output_doc,
                    source_page,
                    page,
                    entries,
                    target_language,
                    fallback,
                    min_scale,
                    font_css,
                    font_archive,
                )

        metadata = source.metadata or {}
        if metadata:
            output_doc.set_metadata(metadata)
        output = Path(output_path)
        output.parent.mkdir(parents=True, exist_ok=True)
        try:
            output_doc.save(output, garbage=4, deflate=True)
        finally:
            output_doc.close()
            source.close()
        check = fitz.open(output)
        try:
            if check.page_count < 1:
                raise TranslatorError("The generated PDF has no pages.")
            extracted = self._normalized_text(
                "\n".join(page.get_text() for page in check)
            )
            expected_values = [
                self._normalized_text(
                    "".join(translations.get(unit.unit_id, unit.segments))
                )
                for unit in self.units
            ]
            if target_language in RTL_LANGUAGES:
                expected_length = sum(len(value) for value in expected_values)
                if expected_length >= 4 and len(extracted) < expected_length * 0.55:
                    raise TranslatorError("PDF RTL text coverage validation failed.")
                return output
            for unit, expected in zip(self.units, expected_values):
                if len(expected) < 4:
                    continue
                sample_size = min(32, len(expected))
                if (
                    expected[:sample_size] not in extracted
                    or expected[-sample_size:] not in extracted
                ):
                    raise TranslatorError(
                        f"PDF text coverage validation failed for unit {unit.unit_id}."
                    )
        finally:
            check.close()
        return output


def artifact_from_input(
    input_type: str, source: str | Path, original_name: str = "translation.txt"
) -> Any:
    if input_type == "text":
        return TextArtifact(str(source), original_name)
    if input_type == "docx":
        return DocxArtifact(source)
    if input_type == "pdf":
        return PdfArtifact(source)
    raise TranslatorError(f"Unsupported input type: {input_type}")


TERMINAL_JOB_STATES = {"completed", "failed", "cancelled"}
STREAM_END_STATES = TERMINAL_JOB_STATES | {"paused"}
JOB_PUBLIC_FIELDS = {
    "id",
    "status",
    "stage",
    "progress",
    "message",
    "error",
    "input_type",
    "original_name",
    "source_language",
    "target_language",
    "tone",
    "output_name",
    "created_at",
    "updated_at",
    "heartbeat_at",
    "paused_from_stage",
    "profile_name",
    "model_key",
    "model_state",
    "backend",
}


class DataDirectoryLock:
    def __init__(self, path: str | Path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.handle = self.path.open("a+b")
        try:
            if os.name == "nt":
                import msvcrt

                if self.path.stat().st_size == 0:
                    self.handle.write(b"0")
                    self.handle.flush()
                self.handle.seek(0)
                msvcrt.locking(self.handle.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl

                fcntl.flock(self.handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except (OSError, BlockingIOError) as exc:
            self.handle.close()
            raise TranslatorError(
                "Another translator process is already using this data directory."
            ) from exc

    def release(self) -> None:
        if self.handle.closed:
            return
        try:
            if os.name == "nt":
                import msvcrt

                self.handle.seek(0)
                msvcrt.locking(self.handle.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                import fcntl

                fcntl.flock(self.handle.fileno(), fcntl.LOCK_UN)
        finally:
            self.handle.close()

    def __del__(self) -> None:
        try:
            self.release()
        except OSError:
            pass


class JobStore:
    def __init__(self, data_dir: str | Path):
        self.data_dir = Path(data_dir)
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.jobs_dir = self.data_dir / "jobs"
        self.jobs_dir.mkdir(parents=True, exist_ok=True)
        self.db_path = self.data_dir / "jobs.sqlite3"
        self._initialize()

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.db_path, timeout=30)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA journal_mode=WAL")
        connection.execute("PRAGMA foreign_keys=ON")
        return connection

    def _initialize(self) -> None:
        with self._connect() as connection:
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS jobs (
                    id TEXT PRIMARY KEY,
                    status TEXT NOT NULL,
                    stage TEXT NOT NULL,
                    progress INTEGER NOT NULL,
                    message TEXT NOT NULL,
                    error TEXT NOT NULL,
                    input_type TEXT NOT NULL,
                    original_name TEXT NOT NULL,
                    source_path TEXT NOT NULL,
                    source_text TEXT NOT NULL,
                    source_language TEXT NOT NULL,
                    target_language TEXT NOT NULL,
                    tone TEXT NOT NULL,
                    custom_instruction TEXT NOT NULL,
                    output_path TEXT NOT NULL,
                    output_name TEXT NOT NULL,
                    cancel_requested INTEGER NOT NULL DEFAULT 0,
                    pause_requested INTEGER NOT NULL DEFAULT 0,
                    paused_from_stage TEXT NOT NULL DEFAULT '',
                    heartbeat_at REAL NOT NULL DEFAULT 0,
                    profile_name TEXT NOT NULL DEFAULT '',
                    model_key TEXT NOT NULL DEFAULT '',
                    model_state TEXT NOT NULL DEFAULT 'not_loaded',
                    backend TEXT NOT NULL DEFAULT '',
                    created_at REAL NOT NULL,
                    updated_at REAL NOT NULL
                )
                """
            )
            existing = {
                str(row[1])
                for row in connection.execute("PRAGMA table_info(jobs)").fetchall()
            }
            migrations = {
                "pause_requested": "INTEGER NOT NULL DEFAULT 0",
                "paused_from_stage": "TEXT NOT NULL DEFAULT ''",
                "heartbeat_at": "REAL NOT NULL DEFAULT 0",
                "profile_name": "TEXT NOT NULL DEFAULT ''",
                "model_key": "TEXT NOT NULL DEFAULT ''",
                "model_state": "TEXT NOT NULL DEFAULT 'not_loaded'",
                "backend": "TEXT NOT NULL DEFAULT ''",
            }
            for column, definition in migrations.items():
                if column not in existing:
                    connection.execute(
                        f"ALTER TABLE jobs ADD COLUMN {column} {definition}"
                    )
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS job_events (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    job_id TEXT NOT NULL,
                    status TEXT NOT NULL,
                    stage TEXT NOT NULL,
                    progress INTEGER NOT NULL,
                    message TEXT NOT NULL,
                    created_at REAL NOT NULL,
                    FOREIGN KEY(job_id) REFERENCES jobs(id) ON DELETE CASCADE
                )
                """
            )

    def create_job(
        self,
        *,
        input_type: str,
        original_name: str,
        source_path: str,
        source_text: str,
        source_language: str,
        target_language: str,
        tone: str,
        custom_instruction: str,
    ) -> str:
        job_id = uuid.uuid4().hex
        now = time.time()
        with self._connect() as connection:
            connection.execute(
                """
                INSERT INTO jobs (
                    id, status, stage, progress, message, error, input_type, original_name,
                    source_path, source_text, source_language, target_language, tone,
                    custom_instruction, output_path, output_name, heartbeat_at,
                    created_at, updated_at
                ) VALUES (?, 'queued', 'queued', 0, ?, '', ?, ?, ?, ?, ?, ?, ?, ?, '', '', ?, ?, ?)
                """,
                (
                    job_id,
                    "در صف پردازش",
                    input_type,
                    original_name,
                    source_path,
                    source_text,
                    source_language,
                    target_language,
                    tone,
                    custom_instruction,
                    now,
                    now,
                    now,
                ),
            )
            connection.execute(
                "INSERT INTO job_events (job_id,status,stage,progress,message,created_at) "
                "VALUES (?, 'queued', 'queued', 0, ?, ?)",
                (job_id, "در صف پردازش", now),
            )
        return job_id

    def get(self, job_id: str) -> dict[str, Any] | None:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT * FROM jobs WHERE id = ?", (job_id,)
            ).fetchone()
        return dict(row) if row else None

    def public(self, job: dict[str, Any]) -> dict[str, Any]:
        value = {key: job.get(key) for key in JOB_PUBLIC_FIELDS}
        heartbeat = float(job.get("heartbeat_at") or job.get("updated_at") or 0)
        heartbeat_age = max(0.0, time.time() - heartbeat)
        progress_age = max(0.0, time.time() - float(job.get("updated_at") or 0))
        value["heartbeat_age_seconds"] = round(heartbeat_age, 1)
        value["progress_age_seconds"] = round(progress_age, 1)
        status = str(job.get("status", ""))
        if status in {"running", "pausing"} and heartbeat_age > 10:
            health_state = "worker_unresponsive"
        elif status in {"running", "pausing"} and progress_age > 120:
            health_state = "long_running"
        elif status in {"running", "pausing"}:
            health_state = "working"
        else:
            health_state = status or "unknown"
        value["health_state"] = health_state
        value["is_stalled"] = health_state in {
            "worker_unresponsive",
            "long_running",
        }
        value["history"] = self.history(str(job["id"]))
        return value

    def list_jobs(self, limit: int = 30) -> list[dict[str, Any]]:
        with self._connect() as connection:
            rows = connection.execute(
                "SELECT * FROM jobs ORDER BY created_at DESC LIMIT ?",
                (max(1, min(100, int(limit))),),
            ).fetchall()
        return [self.public(dict(row)) for row in rows]

    def history(self, job_id: str, limit: int = 30) -> list[dict[str, Any]]:
        with self._connect() as connection:
            rows = connection.execute(
                "SELECT status,stage,progress,message,created_at FROM job_events "
                "WHERE job_id=? ORDER BY id DESC LIMIT ?",
                (job_id, max(1, min(100, int(limit)))),
            ).fetchall()
        return [dict(row) for row in reversed(rows)]

    def update(self, job_id: str, **changes: Any) -> None:
        allowed = {
            "status",
            "stage",
            "progress",
            "message",
            "error",
            "output_path",
            "output_name",
            "cancel_requested",
            "pause_requested",
            "paused_from_stage",
            "heartbeat_at",
            "profile_name",
            "model_key",
            "model_state",
            "backend",
        }
        values = {key: value for key, value in changes.items() if key in allowed}
        if not values:
            return
        if "progress" in values:
            values["progress"] = max(0, min(100, int(values["progress"])))
        values["updated_at"] = time.time()
        assignments = ", ".join(f"{key} = ?" for key in values)
        with self._connect() as connection:
            connection.execute(
                f"UPDATE jobs SET {assignments} WHERE id = ?",
                [*values.values(), job_id],
            )
            if any(
                key in values for key in ("status", "stage", "progress", "message")
            ):
                row = connection.execute(
                    "SELECT status,stage,progress,message FROM jobs WHERE id=?",
                    (job_id,),
                ).fetchone()
                if row:
                    previous = connection.execute(
                        "SELECT status,stage,progress,message FROM job_events "
                        "WHERE job_id=? ORDER BY id DESC LIMIT 1",
                        (job_id,),
                    ).fetchone()
                    snapshot = tuple(row)
                    if previous is None or tuple(previous) != snapshot:
                        connection.execute(
                            "INSERT INTO job_events "
                            "(job_id,status,stage,progress,message,created_at) "
                            "VALUES (?,?,?,?,?,?)",
                            (job_id, *snapshot, time.time()),
                        )

    def touch_heartbeat(self, job_id: str) -> None:
        now = time.time()
        with self._connect() as connection:
            connection.execute(
                "UPDATE jobs SET heartbeat_at=? WHERE id=?", (now, job_id)
            )

    def request_pause(self, job_id: str) -> bool:
        job = self.get(job_id)
        if not job or job["status"] not in {"queued", "running", "pausing"}:
            return False
        if job["status"] == "queued":
            self.update(
                job_id,
                status="paused",
                pause_requested=0,
                paused_from_stage=job["stage"],
                message="تسک پیش از شروع متوقف شد",
            )
        else:
            self.update(
                job_id,
                status="pausing",
                pause_requested=1,
                message="درخواست توقف ثبت شد؛ منتظر رسیدن به مرز امن",
            )
        return True

    def is_pause_requested(self, job_id: str) -> bool:
        job = self.get(job_id)
        return bool(job and job.get("pause_requested"))

    def mark_paused(self, job_id: str) -> None:
        job = self.get(job_id)
        if not job:
            return
        self.update(
            job_id,
            status="paused",
            pause_requested=0,
            paused_from_stage=job["stage"],
            message=f"در مرحلهٔ {job['stage']} متوقف شد؛ آمادهٔ ادامه",
        )

    def resume(self, job_id: str) -> bool:
        job = self.get(job_id)
        if not job or job["status"] not in {"paused", "failed"}:
            return False
        failed_retry = job["status"] == "failed"
        stage = (
            "checkpoint"
            if failed_retry
            else str(job.get("paused_from_stage") or job.get("stage") or "queued")
        )
        self.update(
            job_id,
            status="queued",
            stage=stage,
            pause_requested=0,
            cancel_requested=0,
            error="",
            model_state="unloaded",
            message=(
                "در صف تلاش دوباره از آخرین checkpoint"
                if failed_retry
                else f"در صف ادامه از مرحلهٔ {stage}"
            ),
        )
        return True

    def request_cancel(self, job_id: str) -> bool:
        job = self.get(job_id)
        if not job or job["status"] in TERMINAL_JOB_STATES:
            return False
        if job["status"] in {"queued", "paused"}:
            self.update(
                job_id,
                status="cancelled",
                stage="cancelled",
                cancel_requested=1,
                message="ترجمه لغو شد",
            )
            return True
        self.update(job_id, cancel_requested=1, message="درخواست لغو ثبت شد")
        return True

    def is_cancelled(self, job_id: str) -> bool:
        job = self.get(job_id)
        return bool(job and job.get("cancel_requested"))

    def recoverable(self) -> list[str]:
        with self._connect() as connection:
            rows = connection.execute(
                "SELECT id,stage FROM jobs "
                "WHERE status IN ('queued','running','pausing') ORDER BY created_at"
            ).fetchall()
        for row in rows:
            self.update(
                str(row["id"]),
                status="paused",
                pause_requested=0,
                cancel_requested=0,
                paused_from_stage=str(row["stage"]),
                model_state="unloaded",
                message="پس از راه‌اندازی مجدد متوقف ماند؛ برای ادامه Resume را بزنید",
            )
        return []

    def active_count(self) -> int:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT COUNT(*) AS count FROM jobs "
                "WHERE status IN ('queued','running','pausing')"
            ).fetchone()
        return int(row["count"] if row else 0)

    def delete_job(self, job_id: str) -> bool:
        job = self.get(job_id)
        if not job or job["status"] not in TERMINAL_JOB_STATES | {"paused"}:
            return False
        with self._connect() as connection:
            connection.execute("DELETE FROM jobs WHERE id=?", (job_id,))
        job_dir = self.jobs_dir / job_id
        if job_dir.is_dir():
            shutil.rmtree(job_dir, ignore_errors=True)
        source_value = str(job.get("source_path") or "")
        if source_value:
            source = Path(source_value)
            uploads_root = (self.data_dir / "uploads").resolve()
            try:
                resolved = source.resolve()
                if resolved.is_relative_to(uploads_root) and resolved.is_file():
                    resolved.unlink()
            except OSError:
                pass
        return True

    def clear_deletable(self) -> int:
        with self._connect() as connection:
            rows = connection.execute(
                "SELECT id FROM jobs WHERE status IN "
                "('paused','completed','failed','cancelled')"
            ).fetchall()
        deleted = 0
        for row in rows:
            deleted += int(self.delete_job(str(row["id"])))
        return deleted

    def cleanup(self, retention_days: int) -> None:
        cutoff = time.time() - max(1, retention_days) * 86400
        with self._connect() as connection:
            rows = connection.execute(
                "SELECT id, source_path FROM jobs WHERE updated_at < ? "
                "AND status IN ('completed','failed','cancelled')",
                (cutoff,),
            ).fetchall()
            connection.executemany(
                "DELETE FROM jobs WHERE id = ?", [(row["id"],) for row in rows]
            )
        for row in rows:
            path = self.jobs_dir / str(row["id"])
            if path.is_dir():
                shutil.rmtree(path, ignore_errors=True)
            source_value = str(row["source_path"] or "")
            if source_value:
                source = Path(source_value)
                uploads_root = (self.data_dir / "uploads").resolve()
                try:
                    resolved_source = source.resolve()
                    if (
                        resolved_source.is_relative_to(uploads_root)
                        and resolved_source.is_file()
                    ):
                        resolved_source.unlink()
                except OSError:
                    pass


class JobRunner:
    def __init__(self, config: dict[str, Any], store: JobStore):
        self.config = config
        self.store = store
        self.queue: queue.Queue[str] = queue.Queue()
        self.worker = threading.Thread(
            target=self._loop, name="translator-worker", daemon=True
        )
        self.worker.start()
        for job_id in store.recoverable():
            self.queue.put(job_id)

    def enqueue(self, job_id: str) -> None:
        self.queue.put(job_id)

    def _loop(self) -> None:
        while True:
            job_id = self.queue.get()
            try:
                self._process(job_id)
            finally:
                self.queue.task_done()

    def _checkpoint_path(self, job_id: str) -> Path:
        return self.store.jobs_dir / job_id / "checkpoint.json"

    def _load_checkpoint(self, job_id: str) -> dict[str, Any]:
        path = self._checkpoint_path(job_id)
        if not path.is_file():
            return {}
        try:
            value = json.loads(path.read_text(encoding="utf-8"))
            return value if isinstance(value, dict) else {}
        except (OSError, json.JSONDecodeError):
            return {}

    def _save_checkpoint(self, job_id: str, stage: str, values: dict[str, Any]) -> None:
        job_dir = self.store.jobs_dir / job_id
        job_dir.mkdir(parents=True, exist_ok=True)
        current = self._load_checkpoint(job_id)
        current.update(values)
        current["stage"] = stage
        temporary = job_dir / "checkpoint.tmp"
        temporary.write_text(json.dumps(current, ensure_ascii=False), encoding="utf-8")
        temporary.replace(self._checkpoint_path(job_id))

    def _process(self, job_id: str) -> None:
        job = self.store.get(job_id)
        if not job or job["status"] in TERMINAL_JOB_STATES | {"paused"}:
            return
        if self.store.is_cancelled(job_id):
            self.store.update(
                job_id, status="cancelled", stage="cancelled", message="ترجمه لغو شد"
            )
            return
        self.store.update(
            job_id,
            status="running",
            stage="preparing",
            progress=1,
            message="آماده‌سازی ورودی",
        )
        profile_name = ""
        model_manager: ModelManager | None = None
        heartbeat_stop = threading.Event()

        def heartbeat() -> None:
            while not heartbeat_stop.wait(2):
                current = self.store.get(job_id)
                if not current or current["status"] not in {"running", "pausing"}:
                    return
                self.store.touch_heartbeat(job_id)

        heartbeat_thread = threading.Thread(
            target=heartbeat,
            name=f"translator-heartbeat-{job_id[:8]}",
            daemon=True,
        )
        self.store.touch_heartbeat(job_id)
        heartbeat_thread.start()
        try:
            profile_name, profile = choose_profile(self.config)
            backend = detect_runtime_backend()
            self.store.update(
                job_id,
                profile_name=profile_name,
                model_key=str(profile.get("translator", "")),
                model_state="not_loaded",
                backend=backend,
            )
            input_source: str | Path = (
                job["source_text"]
                if job["input_type"] == "text"
                else job["source_path"]
            )
            artifact = artifact_from_input(
                job["input_type"], input_source, job["original_name"]
            )
            self.store.update(
                job_id,
                stage="structure",
                progress=5,
                message=f"{len(artifact.units)} بخش برای ترجمه استخراج شد — پروفایل {profile_name}",
            )

            def report(stage: str, progress: int, message: str) -> None:
                self.store.update(
                    job_id, stage=stage, progress=progress, message=message
                )

            def save(stage: str, values: dict[str, Any]) -> None:
                self._save_checkpoint(job_id, stage, values)

            previous_model_stage = ["analysis"]

            def model_status(state: str, model_key: str) -> None:
                current = self.store.get(job_id) or {}
                progress = int(current.get("progress") or 5)
                if state == "loading":
                    stage = str(current.get("stage") or "analysis")
                    if stage != "loading_model":
                        previous_model_stage[0] = stage
                    self.store.update(
                        job_id,
                        stage="loading_model",
                        progress=max(5, progress),
                        message=f"در حال بارگذاری مدل {model_key} روی {backend}",
                        model_key=model_key,
                        model_state="loading",
                        backend=backend,
                    )
                elif state == "loaded":
                    self.store.update(
                        job_id,
                        stage=previous_model_stage[0],
                        message=f"مدل {model_key} با موفقیت بارگذاری شد؛ پردازش ادامه دارد",
                        model_key=model_key,
                        model_state="loaded",
                        backend=backend,
                    )
                elif state == "error":
                    self.store.update(
                        job_id,
                        stage="model_error",
                        message=f"بارگذاری مدل {model_key} ناموفق بود",
                        model_key=model_key,
                        model_state="error",
                        backend=backend,
                    )
                else:
                    self.store.update(job_id, model_state="unloaded")

            model_manager = ModelManager(
                self.config, profile, status_callback=model_status
            )
            pipeline = TranslationPipeline(
                self.config,
                profile,
                model_manager,
                progress=report,
                cancelled=lambda: self.store.is_cancelled(job_id),
                paused=lambda: self.store.is_pause_requested(job_id),
                checkpoint=save,
            )
            checkpoint = self._load_checkpoint(job_id)
            try:
                translations = pipeline.run(
                    artifact.units,
                    job["source_language"],
                    job["target_language"],
                    job["tone"],
                    job["custom_instruction"],
                    resume=checkpoint,
                )
            except (JobCancelled, JobPaused):
                raise
            except ModelLoadError:
                if profile_name != "high":
                    raise
                low = copy.deepcopy(self.config.get("profiles", {}).get("low", {}))
                if not low or not _model_exists(self.config, low.get("translator")):
                    raise
                model_manager.unload()
                self.store.update(
                    job_id,
                    stage="fallback",
                    message="مدل پرحافظه اجرا نشد؛ ادامه با پروفایل کم‌حافظه",
                )
                model_manager = ModelManager(
                    self.config, low, status_callback=model_status
                )
                pipeline = TranslationPipeline(
                    self.config,
                    low,
                    model_manager,
                    progress=report,
                    cancelled=lambda: self.store.is_cancelled(job_id),
                    paused=lambda: self.store.is_pause_requested(job_id),
                    checkpoint=save,
                )
                translations = pipeline.run(
                    artifact.units,
                    job["source_language"],
                    job["target_language"],
                    job["tone"],
                    job["custom_instruction"],
                    resume=checkpoint,
                )
            self.store.update(
                job_id, stage="rebuilding", progress=95, message="بازسازی فایل خروجی"
            )
            job_dir = self.store.jobs_dir / job_id
            job_dir.mkdir(parents=True, exist_ok=True)
            suffix = {"text": ".txt", "docx": ".docx", "pdf": ".pdf"}[job["input_type"]]
            filename = output_name(job["original_name"], job["target_language"], suffix)
            output_path = job_dir / filename
            artifact.build(
                translations, output_path, job["target_language"], self.config
            )
            if not output_path.is_file() or output_path.stat().st_size == 0:
                raise TranslatorError(
                    "Output validation failed: the generated file is empty."
                )
            warning_count = sum(
                len(items)
                for items in getattr(pipeline, "quality_warnings", {}).values()
            )
            completion_message = (
                f"ترجمه آماده دانلود است؛ {warning_count} هشدار کنترل کیفیت ثبت شد"
                if warning_count
                else "ترجمه آماده دانلود است"
            )
            self.store.update(
                job_id,
                status="completed",
                stage="completed",
                progress=100,
                message=completion_message,
                output_path=str(output_path),
                output_name=filename,
                error="",
            )
        except JobCancelled:
            self.store.update(
                job_id, status="cancelled", stage="cancelled", message="ترجمه لغو شد"
            )
        except JobPaused:
            self.store.mark_paused(job_id)
        except Exception as exc:  # noqa: BLE001 - worker boundary must persist every failure
            self.store.update(
                job_id,
                status="failed",
                stage="failed",
                message="پردازش متوقف شد",
                error=str(exc),
            )
        finally:
            heartbeat_stop.set()
            if model_manager is not None:
                model_manager.unload()
            try:
                self.store.cleanup(
                    int(self.config.get("storage", {}).get("retention_days", 7))
                )
            except OSError:
                pass


def _safe_upload_name(filename: str, expected_suffix: str) -> str:
    name = Path(filename or "upload").name.replace("\x00", "")
    stem = Path(name).stem
    stem = re.sub(r"[^\w. -]+", "_", stem, flags=re.UNICODE).strip(" .") or "upload"
    return stem + expected_suffix



def _dependency_status() -> dict[str, bool]:
    import importlib.util

    return {
        "flask": importlib.util.find_spec("flask") is not None,
        "ctranslate2": importlib.util.find_spec("ctranslate2") is not None,
        "transformers": importlib.util.find_spec("transformers") is not None,
        "sentencepiece": importlib.util.find_spec("sentencepiece") is not None,
        "llama_cpp": importlib.util.find_spec("llama_cpp") is not None,
        "lxml": importlib.util.find_spec("lxml") is not None,
        "python_docx": importlib.util.find_spec("docx") is not None,
        "pymupdf": importlib.util.find_spec("pymupdf") is not None,
    }


def build_doctor_report(
    config: dict[str, Any], probe_model: bool = False
) -> dict[str, Any]:
    dependencies = _dependency_status()
    backend = detect_runtime_backend()
    llm_acceleration = detect_llama_acceleration() if dependencies.get("llama_cpp") else "unavailable"
    models: dict[str, Any] = {}
    for key, value in config.get("models", {}).items():
        if not isinstance(value, dict):
            continue
        model_backend = str(value.get("backend", ""))
        item: dict[str, Any] = {"backend": model_backend}
        if model_backend == "ctranslate2":
            path = Path(str(value.get("path", "")))
            item.update({"path": str(path), "exists": (path / "model.bin").is_file()})
            try:
                device, compute = ModelManager._resolve_ct2_runtime(value)
                item.update({"runtime_device": device, "compute_type": compute})
            except Exception as exc:
                item["runtime_error"] = str(exc)
        elif model_backend == "llama_cpp":
            path = Path(str(value.get("path", "")))
            item.update(
                {
                    "path": str(path),
                    "exists": path.is_file(),
                    "gpu_layers": ModelManager._gpu_layers_for(value),
                    "acceleration": llm_acceleration,
                }
            )
        else:
            item["exists"] = False
        models[key] = item

    errors: list[str] = []
    required = (
        "flask",
        "ctranslate2",
        "transformers",
        "sentencepiece",
        "llama_cpp",
        "lxml",
        "python_docx",
        "pymupdf",
    )
    missing_deps = [name for name in required if not dependencies.get(name)]
    if missing_deps:
        errors.append("Missing Python packages: " + ", ".join(missing_deps))

    try:
        profile_name, profile = choose_profile(config)
    except TranslatorError as exc:
        profile_name, profile = "unavailable", {}
        errors.append(str(exc))

    if sys.platform == "darwin" and dependencies.get("llama_cpp") and llm_acceleration != "metal":
        errors.append(
            "llama-cpp-python is installed without active Metal GPU offload. "
            "Reinstall a Metal-enabled arm64 build for GPU acceleration."
        )

    probe: dict[str, Any] = {"requested": bool(probe_model), "ok": None, "steps": {}}
    if probe_model and profile_name != "unavailable" and not missing_deps:
        manager = ModelManager(config, profile)
        try:
            translator_key = str(profile.get("translator", ""))
            reviewer_key = str(profile.get("reviewer", ""))
            polisher_key = str(profile.get("polisher") or "")

            mt = manager.translate_texts(translator_key, ["Hello world."], "fa")[0]
            probe["steps"]["madlad"] = {"ok": bool(mt.strip()), "sample": mt[:120]}
            manager.unload(translator_key)

            llm = manager.generate_json(
                reviewer_key,
                [{"role": "user", "content": 'Return only JSON: {"ok": true}'}],
                max_tokens=32,
                temperature=0.0,
            )
            probe["steps"]["reviewer"] = {"ok": llm.get("ok") is True}
            manager.unload(reviewer_key)

            if polisher_key:
                polish = manager.generate_json(
                    polisher_key,
                    [{"role": "user", "content": 'فقط این JSON را برگردان: {"ok": true}'}],
                    max_tokens=32,
                    temperature=0.0,
                )
                probe["steps"]["polisher"] = {"ok": polish.get("ok") is True}
                manager.unload(polisher_key)

            if not mt.strip() or llm.get("ok") is not True:
                raise TranslatorError("Direct-GGUF probe did not return expected outputs.")
            if polisher_key and probe["steps"].get("polisher", {}).get("ok") is not True:
                raise TranslatorError("Dorna GGUF probe did not return expected output.")
            probe["ok"] = True
        except Exception as exc:
            probe.update({"ok": False, "error": str(exc)})
            errors.append(f"Model probe failed: {exc}")
        finally:
            manager.unload()

    healthy = not errors and backend != "incomplete"
    return {
        "app_version": APP_VERSION,
        "status": "ok" if healthy else "incomplete",
        "platform": platform.platform(),
        "python_arch": platform.machine(),
        "memory_gb": round(total_memory_bytes() / 1024**3, 1),
        "backend": backend,
        "llm_acceleration": llm_acceleration,
        "profile": profile_name,
        "profile_config": profile,
        "dependencies": dependencies,
        "models": models,
        "probe": probe,
        "errors": errors,
    }


INDEX_HTML = r"""<!doctype html>
<html lang="fa" dir="rtl">
<head>
  <meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
  <meta name="csrf-token" content="__CSRF_TOKEN__">
  <title>مترجم محلی اسناد</title>
  <style>
    @font-face{font-family:Vazirmatn;src:url('/assets/app-font') format('truetype');font-display:swap}
    :root{--ink:#192522;--muted:#63716d;--paper:#f4f1e8;--card:#fffdf7;--line:#d9d4c6;--accent:#0f766e;--accent2:#115e59;--danger:#b42318}
    *{box-sizing:border-box}body{margin:0;background:radial-gradient(circle at 15% 0,#d7ebe2 0,transparent 35%),var(--paper);color:var(--ink);font-family:Vazirmatn,Tahoma,"Noto Sans Arabic",sans-serif;min-height:100vh}
    main{width:min(920px,calc(100% - 28px));margin:42px auto}.eyebrow{color:var(--accent);font-weight:700;font-size:.82rem;letter-spacing:.08em}h1{font-size:clamp(2rem,5vw,3.5rem);margin:.3rem 0 .7rem;line-height:1.15}header p{color:var(--muted);max-width:680px;line-height:1.9}
    .card{background:rgba(255,253,247,.94);border:1px solid var(--line);border-radius:22px;padding:24px;box-shadow:0 18px 55px rgba(25,37,34,.09);margin-top:25px}.grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:16px}.full{grid-column:1/-1}
    label{display:block;font-weight:700;font-size:.9rem;margin-bottom:8px}select,textarea,input[type=file],input[type=text]{width:100%;border:1px solid var(--line);border-radius:12px;background:#fff;padding:12px;color:var(--ink);font:inherit}textarea{min-height:160px;resize:vertical;line-height:1.8}select:focus,textarea:focus,input:focus{outline:3px solid rgba(15,118,110,.14);border-color:var(--accent)}
    .switch{display:flex;gap:8px;margin-bottom:12px}.switch button{background:#ece8dc;color:var(--muted);border:0;padding:8px 14px;border-radius:999px;cursor:pointer}.switch button.active{background:var(--ink);color:white}.actions{display:flex;gap:10px;align-items:center;margin-top:18px;flex-wrap:wrap}button.primary,a.download{background:var(--accent);color:#fff;border:0;border-radius:12px;padding:12px 20px;font:inherit;font-weight:700;cursor:pointer;text-decoration:none}button.primary:hover,a.download:hover{background:var(--accent2)}button.secondary{background:transparent;border:1px solid var(--line);border-radius:12px;padding:11px 18px;color:var(--danger);cursor:pointer;font:inherit}button.neutral{color:var(--ink)}.hidden{display:none!important}
    #statusCard{overflow:hidden}.statusline{display:flex;justify-content:space-between;gap:16px;align-items:center}.percent{font-size:2rem;font-weight:800;color:var(--accent)}.track{height:12px;background:#e5e1d6;border-radius:999px;margin:18px 0;overflow:hidden}.bar{height:100%;width:0;background:linear-gradient(90deg,var(--accent),#2dd4bf);transition:width .35s ease}.muted{color:var(--muted);font-size:.9rem}.error{color:var(--danger);white-space:pre-wrap;line-height:1.7}.details{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:8px 18px;background:#f3f0e7;border-radius:14px;padding:14px;margin-top:14px}.details div{font-size:.88rem}.health{margin-top:12px;padding:10px 12px;border-radius:10px;background:#e7f5ef;color:#0b5f55}.health.warn{background:#fff0d8;color:#8a4b08}.health.bad{background:#fee9e7;color:var(--danger)}.process{margin-top:14px;line-height:1.9}.history{margin:12px 0 0;padding:0 18px 0 0;max-height:210px;overflow:auto}.history li{margin:6px 0;color:var(--muted);font-size:.84rem}.jobs{display:grid;gap:8px;margin-top:12px}.jobrow{display:flex;justify-content:space-between;gap:12px;align-items:center;width:100%;text-align:right;border:1px solid var(--line);background:#fff;border-radius:12px;padding:10px 12px;font:inherit;cursor:pointer}.jobrow:hover{border-color:var(--accent)}.footer{margin:22px 4px;color:var(--muted);font-size:.8rem}.ltr{direction:ltr;text-align:left}
    button:focus-visible,a:focus-visible{outline:3px solid rgba(15,118,110,.28);outline-offset:2px}
    @media(max-width:700px){main{margin:24px auto}.card{padding:18px}.grid,.details{grid-template-columns:1fr}.full{grid-column:auto}.statusline{align-items:flex-start}.actions{flex-wrap:wrap}}
    @media(prefers-reduced-motion:reduce){*{scroll-behavior:auto!important;transition:none!important}}
  </style>
</head>
<body><main>
  <header><div class="eyebrow">LOCAL · PRIVATE · MADLAD + DIRECT GGUF · نسخه __APP_VERSION__</div><h1>مترجم محلی اسناد</h1><p>خط لولهٔ محلی MADLAD + Gemma 3 GGUF + Dorna GGUF؛ اجرای مستقیم llama.cpp با Metal روی مک و CUDA/GPU offload در سیستم‌های سازگار.</p><div id="runtimeHealth" class="muted">در حال بررسی آمادگی سیستم…</div></header>
  <section class="card" id="formCard"><form id="jobForm"><input type="hidden" name="csrf_token" value="__CSRF_TOKEN__">
    <div class="switch"><button type="button" class="active" data-type="text">متن</button><button type="button" data-type="docx">Word</button><button type="button" data-type="pdf">PDF</button></div>
    <input type="hidden" name="input_type" id="inputType" value="text">
    <div id="textInput"><label for="text">متن ورودی</label><textarea id="text" name="text" placeholder="متن را اینجا وارد کنید…"></textarea></div>
    <div id="fileInput" class="hidden"><label for="file">فایل ورودی</label><input id="file" name="file" type="file"></div>
    <div class="grid" style="margin-top:18px">
      <div><label for="source">زبان مبدأ</label><select id="source" name="source_language"><option value="en">English</option><option value="fa">فارسی</option><option value="ar">العربية</option></select></div>
      <div><label for="target">زبان مقصد</label><select id="target" name="target_language"><option value="fa">فارسی</option><option value="en">English</option><option value="ar">العربية</option></select></div>
      <div><label for="tone">لحن</label><select id="tone" name="tone"><option value="auto">تشخیص خودکار</option><option value="formal">رسمی</option><option value="legal">حقوقی</option><option value="business">تجاری</option><option value="academic">دانشگاهی</option><option value="literary">ادبی</option><option value="conversational">محاوره‌ای</option><option value="screenplay">فیلم‌نامه</option></select></div>
      <div><label for="instruction">توضیح اختیاری</label><input id="instruction" name="custom_instruction" type="text" placeholder="مثلاً برای مخاطب عمومی و روان"></div>
    </div>
    <div class="actions"><button class="primary" type="submit" id="start">شروع ترجمه</button><span class="muted" id="formMessage"></span></div>
  </form></section>
  <section class="card hidden" id="statusCard">
    <div class="statusline"><div><strong id="stage">آماده‌سازی</strong><div class="muted" id="message"></div></div><div class="percent"><span id="percent">0</span>٪</div></div>
    <div class="track"><div class="bar" id="bar"></div></div>
    <div id="health" class="health">سلامت اجرا: در حال بررسی</div>
    <div class="details"><div>وضعیت: <strong id="jobStatus">—</strong></div><div>آخرین مرحله: <strong id="pausedStage">—</strong></div><div>مدل: <strong id="model">—</strong></div><div>وضعیت مدل: <strong id="modelState">—</strong></div><div>Backend: <strong id="backend">—</strong></div><div>آخرین پیشرفت: <strong id="lastProgress">—</strong></div></div>
    <div class="process"><strong>فرایند:</strong> آماده‌سازی ← بارگذاری مدل ← تحلیل ← ترجمه ← ویرایش ← کنترل کیفیت ← ساخت خروجی</div>
    <div id="error" class="error"></div>
    <div class="actions"><button id="pause" class="secondary neutral" type="button">توقف امن</button><button id="resume" class="primary hidden" type="button">ادامه</button><button id="cancel" class="secondary" type="button">لغو کامل</button><button id="delete" class="secondary hidden" type="button">حذف تسک</button><button id="newTask" class="secondary neutral" type="button">تسک جدید</button><a id="download" class="download hidden">دانلود فایل ترجمه‌شده</a></div>
    <details><summary>تاریخچهٔ مراحل</summary><ol id="history" class="history"></ol></details>
  </section>
  <section class="card"><div class="statusline"><strong>تسک‌های اخیر</strong><button id="clearJobs" class="secondary" type="button">پاک‌کردن تسک‌های قابل حذف</button></div><div id="jobs" class="jobs"><span class="muted">تسکی ثبت نشده است.</span></div></section>
  <div class="footer">فایل‌ها و مدل‌ها روی همین دستگاه پردازش می‌شوند.</div>
</main><script>
const form=document.getElementById('jobForm'),inputType=document.getElementById('inputType'),textInput=document.getElementById('textInput'),fileInput=document.getElementById('fileInput'),csrfToken=document.querySelector('meta[name="csrf-token"]').content;
let currentJob=null,events=null;
const terminal=['completed','failed','cancelled'];
const stageNames={queued:'در صف',checkpoint:'ادامه از checkpoint',preparing:'آماده‌سازی',structure:'استخراج ساختار',loading_model:'بارگذاری مدل',model_error:'خطای مدل',analysis:'تحلیل متن',manifest:'ثبت checkpoint',translation:'ترجمهٔ اولیه',draft:'ترجمهٔ اولیه',context:'به‌روزرسانی حافظه',editing:'ویرایش',quality:'کنترل کیفیت',quality_repair:'اصلاح کیفیت',fallback:'مدل جایگزین',rebuilding:'ساخت خروجی',completed:'تکمیل‌شده',paused:'متوقف',failed:'خطا',cancelled:'لغوشده'};
const statusNames={queued:'در صف',running:'در حال اجرا',pausing:'در انتظار توقف امن',paused:'متوقف',completed:'تکمیل‌شده',failed:'ناموفق',cancelled:'لغوشده'};
const modelNames={not_loaded:'هنوز بارگذاری نشده',loading:'در حال بارگذاری',loaded:'بارگذاری شده',unloaded:'از حافظه خارج شده',error:'خطای بارگذاری'};
document.querySelectorAll('.switch button').forEach(btn=>btn.onclick=()=>{document.querySelectorAll('.switch button').forEach(x=>x.classList.remove('active'));btn.classList.add('active');inputType.value=btn.dataset.type;const isText=btn.dataset.type==='text';textInput.classList.toggle('hidden',!isText);fileInput.classList.toggle('hidden',isText);document.getElementById('file').accept=btn.dataset.type==='docx'?'.docx':btn.dataset.type==='pdf'?'.pdf':'';});
function setText(id,value){document.getElementById(id).textContent=(value===null||value===undefined)?'—':value}
function showState(job){
  currentJob=job.id;localStorage.setItem('translatorJobId',job.id);document.getElementById('statusCard').classList.remove('hidden');
  setText('stage',stageNames[job.stage]||job.stage||job.status);setText('message',job.message||'');setText('jobStatus',statusNames[job.status]||job.status);setText('pausedStage',stageNames[job.paused_from_stage]||job.paused_from_stage);setText('model',job.model_key);setText('modelState',modelNames[job.model_state]||job.model_state);setText('backend',job.backend);setText('lastProgress',Math.round(job.progress_age_seconds||0)+' ثانیه پیش');
  document.getElementById('percent').textContent=job.progress||0;document.getElementById('bar').style.width=(job.progress||0)+'%';document.getElementById('error').textContent=job.error||'';
  const health=document.getElementById('health');health.className='health';
  if(job.health_state==='worker_unresponsive'){health.textContent='سلامت اجرا: heartbeat پردازش قطع شده؛ احتمال توقف Worker وجود دارد.';health.classList.add('bad')}
  else if(job.health_state==='long_running'){health.textContent='سلامت اجرا: برنامه زنده است، اما بیش از دو دقیقه پیشرفت مرحله‌ای ثبت نشده؛ inference ممکن است طولانی یا متوقف شده باشد.';health.classList.add('warn')}
  else if(job.health_state==='working'){health.textContent='سلامت اجرا: Worker پاسخگو است و پردازش ادامه دارد.'}
  else if(job.status==='paused'){health.textContent='سلامت اجرا: تسک عمداً متوقف است و مدل از حافظه خارج می‌شود.'}
  else if(job.status==='queued'){health.textContent='سلامت اجرا: تسک در صف است و هنوز مدل را اشغال نکرده.'}
  else if(job.status==='failed'){health.textContent='سلامت اجرا: این تسک متوقف شده و اکنون هیچ پردازشی برای آن در حال اجرا نیست.';health.classList.add('bad')}
  else{health.textContent='سلامت اجرا: '+(statusNames[job.status]||job.status)}
  const done=job.status==='completed',paused=job.status==='paused',retryable=paused||job.status==='failed',active=['queued','running','pausing'].includes(job.status),deletable=paused||terminal.includes(job.status),resume=document.getElementById('resume');
  document.getElementById('download').classList.toggle('hidden',!done);document.getElementById('pause').classList.toggle('hidden',!['queued','running'].includes(job.status));resume.classList.toggle('hidden',!retryable);resume.textContent=job.status==='failed'?'تلاش دوباره از checkpoint':'ادامه';document.getElementById('cancel').classList.toggle('hidden',!active&&!paused);document.getElementById('delete').classList.toggle('hidden',!deletable);if(done)document.getElementById('download').href='/api/jobs/'+job.id+'/download';
  const history=document.getElementById('history');history.replaceChildren();(job.history||[]).slice(-12).reverse().forEach(item=>{const li=document.createElement('li');li.textContent=(stageNames[item.stage]||item.stage)+' — '+item.progress+'٪ — '+item.message;history.appendChild(li)});
}
async function loadJobs(){const res=await fetch('/api/jobs');if(!res.ok)return;const jobs=(await res.json()).jobs||[],box=document.getElementById('jobs');box.replaceChildren();if(!jobs.length){const empty=document.createElement('span');empty.className='muted';empty.textContent='تسکی ثبت نشده است.';box.appendChild(empty);return}jobs.forEach(job=>{const button=document.createElement('button');button.type='button';button.className='jobrow';const name=document.createElement('span');name.textContent=job.original_name+' · '+(stageNames[job.stage]||job.stage)+' · '+job.progress+'٪';const state=document.createElement('strong');state.textContent=statusNames[job.status]||job.status;button.append(name,state);button.onclick=()=>selectJob(job.id);box.appendChild(button)})}
function connectJob(jobId){currentJob=jobId;localStorage.setItem('translatorJobId',jobId);if(events)events.close();events=new EventSource('/api/jobs/'+jobId+'/events');events.addEventListener('progress',e=>{const job=JSON.parse(e.data);showState(job);if(terminal.includes(job.status)||job.status==='paused'){events.close();loadJobs()}});events.onerror=()=>{if(events)document.getElementById('message').textContent='ارتباط وضعیت موقتاً قطع شد؛ اتصال مجدد خودکار انجام می‌شود.'}}
async function selectJob(jobId){const res=await fetch('/api/jobs/'+jobId);if(!res.ok){await loadJobs();return}const job=await res.json();showState(job);if(['queued','running','pausing'].includes(job.status))connectJob(jobId);else if(events){events.close();events=null}}
async function jobAction(action,method='POST'){if(!currentJob)return;const res=await fetch('/api/jobs/'+currentJob+(action?'/'+action:''),{method,headers:{'X-CSRF-Token':csrfToken}});const data=await res.json();if(!res.ok){alert(data.error||'عملیات انجام نشد');return false}await selectJob(currentJob);await loadJobs();return true}
form.onsubmit=async e=>{e.preventDefault();setText('formMessage','');document.getElementById('error').textContent='';const res=await fetch('/api/jobs',{method:'POST',body:new FormData(form)}),data=await res.json();if(!res.ok){setText('formMessage',data.error||'خطا در ثبت کار');return}await loadJobs();connectJob(data.job_id)};
document.getElementById('pause').onclick=()=>jobAction('pause');document.getElementById('resume').onclick=async()=>{if(await jobAction('resume'))connectJob(currentJob)};document.getElementById('cancel').onclick=()=>jobAction('cancel');
document.getElementById('delete').onclick=async()=>{if(confirm('این تسک، checkpoint و فایل‌هایش حذف شود؟')&&await jobAction('', 'DELETE')){if(events)events.close();localStorage.removeItem('translatorJobId');currentJob=null;document.getElementById('statusCard').classList.add('hidden');await loadJobs()}};
document.getElementById('newTask').onclick=()=>{form.reset();document.querySelector('.switch button[data-type="text"]').click();document.getElementById('formCard').scrollIntoView({behavior:'smooth'});document.getElementById('text').focus()};
document.getElementById('clearJobs').onclick=async()=>{if(!confirm('همهٔ تسک‌های متوقف، تمام‌شده، ناموفق و لغوشده حذف شوند؟'))return;const res=await fetch('/api/jobs/clear',{method:'POST',headers:{'X-CSRF-Token':csrfToken}});if(res.ok){const data=await res.json();setText('formMessage',data.deleted+' تسک حذف شد');await loadJobs()}};
async function loadDoctor(){const res=await fetch('/api/doctor'),report=await res.json(),el=document.getElementById('runtimeHealth');const existing=Object.entries(report.models||{}).filter(([,value])=>value.exists).map(([key])=>key);el.textContent='آمادگی سیستم: '+(report.status==='ok'?'آماده':'ناقص')+' · Backend: '+report.backend+' · شتاب LLM: '+(report.llm_acceleration||'—')+' · پروفایل: '+report.profile+' · مدل‌ها: '+(existing.join('، ')||'هیچ‌کدام');if(report.errors&&report.errors.length)el.title=report.errors.join('\n')}
loadDoctor();loadJobs();const savedJob=localStorage.getItem('translatorJobId');if(savedJob)selectJob(savedJob);
</script></body></html>"""


def create_app(config: dict[str, Any], start_worker: bool = True) -> Any:
    try:
        from flask import (
            Flask,
            Response,
            jsonify,
            request,
            send_file,
            session,
            stream_with_context,
        )
    except ImportError as exc:
        raise TranslatorError(
            "Flask is not installed. Run 'pip install Flask'."
        ) from exc

    application = Flask(__name__)
    application.secret_key = secrets.token_bytes(32)
    application.config["TRUSTED_HOSTS"] = ["localhost", "127.0.0.1", "[::1]"]
    max_upload = int(config.get("runtime", {}).get("max_upload_mb", 100)) * 1024 * 1024
    application.config["MAX_CONTENT_LENGTH"] = max_upload
    data_dir = Path(config.get("paths", {}).get("data_dir", "./translator_data"))
    process_lock = (
        DataDirectoryLock(data_dir / "translator.lock") if start_worker else None
    )
    store = JobStore(data_dir)
    store.cleanup(int(config.get("storage", {}).get("retention_days", 7)))
    runner = JobRunner(config, store) if start_worker else None
    application.extensions["translator_job_store"] = store
    application.extensions["translator_job_runner"] = runner
    application.extensions["translator_process_lock"] = process_lock

    def csrf_is_valid() -> bool:
        expected = str(session.get("csrf_token", ""))
        supplied = str(
            request.form.get("csrf_token", "")
            or request.headers.get("X-CSRF-Token", "")
        )
        if (
            not expected
            or not supplied
            or not secrets.compare_digest(expected, supplied)
        ):
            return False
        origin = request.headers.get("Origin")
        if origin:
            hostname = (urlparse(origin).hostname or "").lower()
            if hostname not in {"localhost", "127.0.0.1", "::1"}:
                return False
        return True

    @application.get("/")
    def index() -> Response:
        token = str(session.get("csrf_token") or secrets.token_urlsafe(32))
        session["csrf_token"] = token
        response = Response(
            INDEX_HTML.replace(
                "__CSRF_TOKEN__", html.escape(token, quote=True)
            ).replace("__APP_VERSION__", APP_VERSION),
            mimetype="text/html",
        )
        response.headers["Cache-Control"] = "no-store, max-age=0"
        return response

    @application.get("/assets/app-font")
    def app_font() -> Any:
        font_value = config.get("documents", {}).get("font_files", {}).get("fa")
        font_path = Path(str(font_value or ""))
        if not font_value or not font_path.is_file():
            return Response(status=404)
        return send_file(font_path, mimetype="font/ttf", max_age=3600)

    @application.get("/api/doctor")
    def doctor() -> Any:
        report = build_doctor_report(config)
        return jsonify(report), (200 if report["status"] == "ok" else 503)

    @application.post("/api/jobs")
    def submit_job() -> Any:
        if not csrf_is_valid():
            return jsonify(
                {"error": "درخواست امنیتی معتبر نیست؛ صفحه را تازه‌سازی کنید."}
            ), 403
        input_type = str(request.form.get("input_type", "text")).lower()
        source_language = str(request.form.get("source_language", ""))
        target_language = str(request.form.get("target_language", ""))
        tone = str(request.form.get("tone", "auto"))
        custom_instruction = str(request.form.get("custom_instruction", ""))[:1000]
        if input_type not in {"text", "docx", "pdf"}:
            return jsonify({"error": "نوع ورودی پشتیبانی نمی‌شود."}), 400
        if source_language not in LANGUAGES or target_language not in LANGUAGES:
            return jsonify({"error": "زبان مبدأ یا مقصد معتبر نیست."}), 400
        if source_language == target_language:
            return jsonify({"error": "زبان مبدأ و مقصد باید متفاوت باشند."}), 400
        if tone not in TONE_GUIDANCE:
            return jsonify({"error": "لحن انتخاب‌شده معتبر نیست."}), 400
        queue_limit = max(1, int(config.get("runtime", {}).get("max_queued_jobs", 3)))
        if store.active_count() >= queue_limit:
            return jsonify(
                {"error": "صف ترجمه پر است؛ پس از پایان کار جاری دوباره تلاش کنید."}
            ), 429
        source_text = ""
        source_path = ""
        original_name = "translation.txt"
        if input_type == "text":
            source_text = str(request.form.get("text", ""))
            if not source_text.strip():
                return jsonify({"error": "متن ورودی خالی است."}), 400
            if len(source_text.encode("utf-8")) > max_upload:
                return jsonify({"error": "حجم متن بیشتر از حد مجاز است."}), 413
        else:
            uploaded = request.files.get("file")
            suffix = ".docx" if input_type == "docx" else ".pdf"
            if uploaded is None or not uploaded.filename:
                return jsonify({"error": "فایل ورودی انتخاب نشده است."}), 400
            if Path(uploaded.filename).suffix.lower() != suffix:
                return jsonify(
                    {"error": f"برای این ورودی فقط فایل {suffix} مجاز است."}
                ), 400
            original_name = _safe_upload_name(uploaded.filename, suffix)
            upload_dir = store.data_dir / "uploads"
            upload_dir.mkdir(parents=True, exist_ok=True)
            source_path = str(upload_dir / f"{uuid.uuid4().hex}{suffix}")
            uploaded.save(source_path)
            if Path(source_path).stat().st_size == 0:
                Path(source_path).unlink(missing_ok=True)
                return jsonify({"error": "فایل بارگذاری‌شده خالی است."}), 400
        job_id = store.create_job(
            input_type=input_type,
            original_name=original_name,
            source_path=source_path,
            source_text=source_text,
            source_language=source_language,
            target_language=target_language,
            tone=tone,
            custom_instruction=custom_instruction,
        )
        if runner is not None:
            runner.enqueue(job_id)
        return jsonify({"job_id": job_id}), 201

    @application.get("/api/jobs")
    def list_jobs() -> Any:
        return jsonify({"jobs": store.list_jobs()})

    @application.post("/api/jobs/clear")
    def clear_jobs() -> Any:
        if not csrf_is_valid():
            return jsonify({"error": "درخواست امنیتی معتبر نیست."}), 403
        return jsonify({"deleted": store.clear_deletable()})

    @application.get("/api/jobs/<job_id>")
    def job_status(job_id: str) -> Any:
        job = store.get(job_id)
        if not job:
            return jsonify({"error": "کار پیدا نشد."}), 404
        return jsonify(store.public(job))

    @application.get("/api/jobs/<job_id>/events")
    def job_events(job_id: str) -> Any:
        if not store.get(job_id):
            return jsonify({"error": "کار پیدا نشد."}), 404

        def events() -> Iterable[str]:
            last_updated = None
            last_sent = 0.0
            while True:
                job = store.get(job_id)
                if not job:
                    break
                now = time.time()
                if job["updated_at"] != last_updated or now - last_sent >= 2:
                    payload = json.dumps(store.public(job), ensure_ascii=False)
                    yield f"event: progress\ndata: {payload}\n\n"
                    last_updated = job["updated_at"]
                    last_sent = now
                else:
                    yield ": keep-alive\n\n"
                if job["status"] in STREAM_END_STATES:
                    break
                time.sleep(0.5)

        response = Response(stream_with_context(events()), mimetype="text/event-stream")
        response.headers["Cache-Control"] = "no-cache"
        response.headers["X-Accel-Buffering"] = "no"
        return response

    @application.post("/api/jobs/<job_id>/cancel")
    def cancel_job(job_id: str) -> Any:
        if not csrf_is_valid():
            return jsonify({"error": "درخواست امنیتی معتبر نیست."}), 403
        if not store.get(job_id):
            return jsonify({"error": "کار پیدا نشد."}), 404
        if not store.request_cancel(job_id):
            return jsonify({"error": "این کار قابل لغو نیست."}), 409
        return jsonify({"status": "cancelling"})

    @application.post("/api/jobs/<job_id>/pause")
    def pause_job(job_id: str) -> Any:
        if not csrf_is_valid():
            return jsonify({"error": "درخواست امنیتی معتبر نیست."}), 403
        if not store.get(job_id):
            return jsonify({"error": "کار پیدا نشد."}), 404
        if not store.request_pause(job_id):
            return jsonify({"error": "این کار قابل توقف نیست."}), 409
        job = store.get(job_id)
        return jsonify({"status": job["status"] if job else "pausing"})

    @application.post("/api/jobs/<job_id>/resume")
    def resume_job(job_id: str) -> Any:
        if not csrf_is_valid():
            return jsonify({"error": "درخواست امنیتی معتبر نیست."}), 403
        if not store.get(job_id):
            return jsonify({"error": "کار پیدا نشد."}), 404
        if not store.resume(job_id):
            return jsonify({"error": "این کار قابل ادامه نیست."}), 409
        if runner is not None:
            runner.enqueue(job_id)
        return jsonify({"status": "queued"})

    @application.delete("/api/jobs/<job_id>")
    def delete_job(job_id: str) -> Any:
        if not csrf_is_valid():
            return jsonify({"error": "درخواست امنیتی معتبر نیست."}), 403
        if not store.get(job_id):
            return jsonify({"error": "کار پیدا نشد."}), 404
        if not store.delete_job(job_id):
            return jsonify({"error": "ابتدا کار در حال اجرا را متوقف کنید."}), 409
        return jsonify({"status": "deleted"})

    @application.get("/api/jobs/<job_id>/download")
    def download_job(job_id: str) -> Any:
        job = store.get(job_id)
        if not job:
            return jsonify({"error": "کار پیدا نشد."}), 404
        output_path = Path(job.get("output_path") or "")
        if job["status"] != "completed" or not output_path.is_file():
            return jsonify({"error": "فایل خروجی هنوز آماده نیست."}), 409
        return send_file(
            output_path, as_attachment=True, download_name=job["output_name"]
        )

    @application.errorhandler(413)
    def too_large(_: Any) -> Any:
        return jsonify({"error": "حجم ورودی بیشتر از حد مجاز تنظیم‌شده است."}), 413

    return application


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Local hybrid MT + LLM document translator")
    parser.add_argument(
        "--config",
        default=os.environ.get("TRANSLATOR_CONFIG", "translator.config.json"),
    )
    parser.add_argument(
        "--doctor", action="store_true", help="Check configuration and dependencies"
    )
    parser.add_argument(
        "--probe-model",
        action="store_true",
        help="With --doctor, also load the selected MT/reviewer models and run a tiny inference",
    )
    parser.add_argument("--host", help="Override the configured bind address")
    parser.add_argument("--port", type=int, help="Override the configured port")
    parser.add_argument("--debug", action="store_true", help="Enable Flask debug mode")
    args = parser.parse_args(argv)
    config = load_config(args.config)
    if args.doctor:
        report = build_doctor_report(config, probe_model=args.probe_model)
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return 0 if report["status"] == "ok" else 1
    runtime = config.get("runtime", {})
    host = args.host or str(runtime.get("host", "127.0.0.1"))
    port = args.port or int(runtime.get("port", 5000))
    if not is_loopback_host(host):
        raise TranslatorError(
            "This local-only release refuses non-loopback binding. Use 127.0.0.1 or localhost."
        )
    web_app = create_app(config, start_worker=True)
    print(f"Local translator: http://{host}:{port}")
    web_app.run(
        host=host, port=port, debug=args.debug, threaded=True, use_reloader=False
    )
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except TranslatorError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        raise SystemExit(2)
