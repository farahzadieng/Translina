#!/usr/bin/env python3
"""Local LLM document translator.

Run:
    python app.py --config translator.config.json

Required packages:
    flask llama-cpp-python lxml python-docx pymupdf

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


class TranslatorError(RuntimeError):
    """A user-facing translation failure."""


class JobCancelled(TranslatorError):
    """Raised at a safe checkpoint when a user cancels a job."""


class ModelCapacityError(TranslatorError):
    """The rendered request cannot fit the model or available memory."""


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
        or "/" in expanded
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
    if not model_key:
        return False
    model = config.get("models", {}).get(model_key, {})
    path = model.get("path")
    return bool(path and Path(path).is_file())


def choose_profile(
    config: dict[str, Any], total_memory_bytes: int | None = None
) -> tuple[str, dict[str, Any]]:
    memory = (
        total_memory_bytes
        if total_memory_bytes is not None
        else globals()["total_memory_bytes"]()
    )
    threshold = (
        float(config.get("runtime", {}).get("low_memory_threshold_gb", 10)) * 1024**3
    )
    requested = str(config.get("runtime", {}).get("memory_profile", "auto")).lower()
    desired = (
        requested
        if requested in {"low", "high"}
        else ("low" if memory <= threshold else "high")
    )
    profiles = config.get("profiles", {})
    profile = copy.deepcopy(profiles.get(desired, {}))
    required = {profile.get("translator"), profile.get("reviewer")}
    required.discard(None)
    if desired == "high" and (
        not required or not all(_model_exists(config, key) for key in required)
    ):
        desired = "low"
        profile = copy.deepcopy(profiles.get("low", {}))
    if not profile:
        raise TranslatorError(f"Runtime profile '{desired}' is not configured.")
    if not _model_exists(config, profile.get("translator")):
        raise TranslatorError(
            "No usable translator model was found for the selected profile."
        )
    if not _model_exists(config, profile.get("reviewer")):
        profile["reviewer"] = profile["translator"]
    return desired, profile


def direction_for(language: str) -> str:
    if language not in LANGUAGES:
        raise TranslatorError(f"Unsupported language: {language}")
    return "rtl" if language in RTL_LANGUAGES else "ltr"


def is_loopback_host(host: str) -> bool:
    normalized = host.strip().lower().strip("[]")
    return normalized in {"localhost", "127.0.0.1", "::1"}


def detect_runtime_backend() -> str:
    """Report what the installed llama.cpp build can actually offload to."""
    try:
        from llama_cpp import llama_cpp as llama_backend
    except (ImportError, OSError):
        return "unavailable"
    try:
        supports_offload = bool(llama_backend.llama_supports_gpu_offload())
    except (AttributeError, OSError):
        supports_offload = False
    if not supports_offload:
        return "cpu"
    return "metal" if sys.platform == "darwin" else "gpu-offload"


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


def _model_manifest_entry(config: dict[str, Any], model_key: str) -> dict[str, Any]:
    path = Path(str(config.get("models", {}).get(model_key, {}).get("path", "")))
    entry: dict[str, Any] = {"key": model_key, "path": str(path)}
    try:
        stat = path.stat()
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
    translator = str(profile.get("translator", ""))
    reviewer = str(profile.get("reviewer", translator))
    unit_hashes = {
        unit.unit_id: hashlib.sha256(
            json.dumps(unit.segments, ensure_ascii=False, separators=(",", ":")).encode(
                "utf-8"
            )
        ).hexdigest()
        for unit in units
    }
    return {
        "version": 1,
        "route": {
            "translator": _model_manifest_entry(config, translator),
            "reviewer": _model_manifest_entry(config, reviewer),
            "context_size": int(profile.get("context_size", 0)),
            "input_budget": int(input_budget),
        },
        "units": unit_hashes,
    }


def reusable_checkpoint_translations(
    state: dict[str, Any], manifest: dict[str, Any]
) -> tuple[dict[str, list[str]], dict[str, list[str]]]:
    previous = state.get("manifest")
    if not isinstance(previous, dict) or previous.get("route") != manifest.get("route"):
        return {}, {}
    previous_units = previous.get("units", {})
    current_units = manifest.get("units", {})
    if not isinstance(previous_units, dict) or not isinstance(current_units, dict):
        return {}, {}
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
            unit_id: value
            for unit_id, value in values.items()
            if unit_id in valid_ids and isinstance(value, list)
        }

    return filtered("drafts"), filtered("finals")


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


def quality_flags(source: str, target: str) -> list[str]:
    flags: list[str] = []
    if not target.strip():
        return ["empty"]
    source_numbers = re.findall(r"(?<!\w)[+-]?(?:\d[\d,./:-]*\d|\d)(?!\w)", source)
    target_numbers = re.findall(r"(?<!\w)[+-]?(?:\d[\d,./:-]*\d|\d)(?!\w)", target)
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
    ratio = len(target.strip()) / max(1, len(source.strip()))
    if len(source.strip()) > 80 and (ratio < 0.25 or ratio > 4.0):
        flags.append("length")
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
    """Own exactly one llama.cpp model at a time."""

    def __init__(self, config: dict[str, Any], profile: dict[str, Any]):
        self.config = config
        self.profile = profile
        self._model: Any = None
        self._model_key: str | None = None
        self._lock = threading.RLock()

    def _load(self, model_key: str) -> Any:
        with self._lock:
            if self._model is not None and self._model_key == model_key:
                return self._model
            self.unload()
            try:
                from llama_cpp import Llama
            except ImportError as exc:
                raise TranslatorError(
                    "llama-cpp-python is not installed. Run 'pip install llama-cpp-python'."
                ) from exc
            model_config = self.config.get("models", {}).get(model_key)
            if not model_config:
                raise TranslatorError(f"Model '{model_key}' is not configured.")
            model_path = Path(str(model_config.get("path", "")))
            if not model_path.is_file():
                raise TranslatorError(f"Model file not found: {model_path}")
            runtime = self.config.get("runtime", {})
            threads_value = runtime.get("threads", "auto")
            threads = (
                max(1, (os.cpu_count() or 4) - 1)
                if threads_value == "auto"
                else int(threads_value)
            )
            gpu_value = runtime.get("gpu_layers", "auto")
            if gpu_value == "auto":
                gpu_layers = (
                    -1 if detect_runtime_backend() in {"metal", "gpu-offload"} else 0
                )
            else:
                gpu_layers = int(gpu_value)
            kwargs: dict[str, Any] = {
                "model_path": str(model_path),
                "n_ctx": int(self.profile.get("context_size", 4096)),
                "n_batch": int(self.profile.get("batch_size", 128)),
                "n_threads": threads,
                "n_gpu_layers": gpu_layers,
                "verbose": bool(runtime.get("verbose_model", False)),
            }
            if model_config.get("chat_format"):
                kwargs["chat_format"] = model_config["chat_format"]
            try:
                self._model = Llama(**kwargs)
            except Exception as exc:
                raise TranslatorError(
                    f"Could not load model '{model_key}': {exc}"
                ) from exc
            self._model_key = model_key
            return self._model

    def count_tokens(self, model_key: str, text: str) -> int:
        model = self._load(model_key)
        try:
            return len(model.tokenize(text.encode("utf-8"), add_bos=False))
        except TypeError:
            return len(model.tokenize(text.encode("utf-8")))

    def generate(
        self,
        model_key: str,
        messages: list[dict[str, str]],
        max_tokens: int,
        temperature: float,
    ) -> str:
        model = self._load(model_key)
        with self._lock:
            kwargs = {
                "messages": messages,
                "max_tokens": int(max_tokens),
                "temperature": float(temperature),
                "top_p": 0.9,
                "repeat_penalty": 1.05,
            }
            try:
                try:
                    response = model.create_chat_completion(
                        **kwargs, response_format={"type": "json_object"}
                    )
                except (TypeError, ValueError):
                    response = model.create_chat_completion(**kwargs)
                choice = response["choices"][0]
                if str(choice.get("finish_reason", "")).lower() == "length":
                    raise ModelCapacityError(
                        "Model generation reached the output-token limit."
                    )
                return str(choice["message"]["content"] or "")
            except ModelCapacityError:
                raise
            except Exception as exc:
                message = str(exc).lower()
                capacity_markers = (
                    "context window",
                    "context is full",
                    "exceed context",
                    "kv cache",
                    "out of memory",
                    "failed to allocate",
                    "could not allocate",
                    " n_ctx",
                    "oom",
                )
                error_type = (
                    ModelCapacityError
                    if any(marker in message for marker in capacity_markers)
                    else TranslatorError
                )
                raise error_type(
                    f"Model generation failed ({model_key}): {exc}"
                ) from exc

    def unload(self) -> None:
        with self._lock:
            model, self._model = self._model, None
            self._model_key = None
            if model is not None:
                try:
                    close = getattr(model, "close", None)
                    if callable(close):
                        close()
                finally:
                    del model
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
    def __init__(
        self,
        config: dict[str, Any],
        profile: dict[str, Any],
        models: ModelManager,
        progress: Any | None = None,
        cancelled: Any | None = None,
        checkpoint: Any | None = None,
    ):
        self.config = config
        self.profile = profile
        self.models = models
        self.progress = progress or (lambda *args, **kwargs: None)
        self.cancelled = cancelled or (lambda: False)
        self.checkpoint = checkpoint or (lambda *args, **kwargs: None)
        self.options = config.get("translation", {})

    def _check_cancelled(self) -> None:
        if self.cancelled():
            raise JobCancelled("Translation cancelled.")

    def _json_call(
        self,
        model_key: str,
        messages: list[dict[str, str]],
        max_tokens: int,
        temperature: float,
    ) -> dict[str, Any]:
        retries = int(self.options.get("max_retries", 2))
        last_error: Exception | None = None
        working_messages = list(messages)
        for attempt in range(retries + 1):
            self._check_cancelled()
            try:
                rendered = json.dumps(working_messages, ensure_ascii=False)
                input_tokens = self.models.count_tokens(model_key, rendered)
                context_size = int(self.profile.get("context_size", 4096))
                safety = int(self.options.get("context_safety_tokens", 512))
                safe_output = context_size - safety - input_tokens
                if safe_output < 64:
                    raise ModelCapacityError(
                        f"Rendered request uses {input_tokens} tokens and leaves no safe output space."
                    )
                raw = self.models.generate(
                    model_key,
                    working_messages,
                    min(int(max_tokens), safe_output),
                    temperature,
                )
                return extract_json_object(raw)
            except ModelCapacityError:
                raise
            except TranslatorError as exc:
                last_error = exc
                if attempt >= retries:
                    break
                working_messages = list(messages) + [
                    {
                        "role": "user",
                        "content": "Return only one valid JSON object matching the requested schema.",
                    }
                ]
        raise TranslatorError(
            f"The model returned invalid structured output: {last_error}"
        )

    def _analyze(
        self,
        units: Sequence[TranslationUnit],
        source_language: str,
        target_language: str,
        tone: str,
    ) -> dict[str, Any]:
        model_key = self.profile["translator"]
        sample = "\n\n".join(unit.text for unit in units if unit.text.strip())
        reserved = min(1024, int(self.options.get("reserved_output_tokens", 1536)))
        context_size = int(self.profile.get("context_size", 4096))
        safety = int(self.options.get("context_safety_tokens", 512))
        fixed_analysis_prompt = (
            "You are a senior translation analyst. Return JSON with keys domain, audience, summary, "
            "style, and terms. terms is an array of objects with source and target. Do not translate the document."
        )
        fixed_tokens = self.models.count_tokens(model_key, fixed_analysis_prompt) + 128
        sample_budget = max(64, context_size - reserved - safety - fixed_tokens)
        sample = truncate_to_token_budget(
            sample,
            lambda value: self.models.count_tokens(model_key, value),
            sample_budget,
        )
        prompt = {
            "source_language": LANGUAGES[source_language],
            "target_language": LANGUAGES[target_language],
            "requested_tone": tone,
            "document_sample": sample,
        }
        messages = [
            {
                "role": "system",
                "content": (fixed_analysis_prompt),
            },
            {"role": "user", "content": json.dumps(prompt, ensure_ascii=False)},
        ]
        return self._json_call(
            model_key,
            messages,
            max_tokens=reserved,
            temperature=0.05,
        )

    def _batch_budget(self, model_key: str, fixed_context: str) -> int:
        context_size = int(self.profile.get("context_size", 4096))
        reserved = int(self.options.get("reserved_output_tokens", 1536))
        safety = int(self.options.get("context_safety_tokens", 512))
        fixed = self.models.count_tokens(model_key, fixed_context)
        return context_input_budget(
            context_size=context_size,
            fixed_tokens=fixed,
            reserved_output_tokens=reserved,
            safety_tokens=safety,
        )

    def _translate_batch(
        self,
        model_key: str,
        batch: Sequence[TranslationUnit],
        source_language: str,
        target_language: str,
        tone: str,
        analysis: dict[str, Any],
        previous: list[dict[str, str]],
    ) -> dict[str, list[str]]:
        previous = fit_previous_context(
            previous,
            lambda text: self.models.count_tokens(model_key, text),
            max(128, int(self.profile.get("context_size", 4096) * 0.12)),
        )
        schema = {
            unit.unit_id: ["one translated string per input segment"] for unit in batch
        }
        payload = {
            "source_language": LANGUAGES[source_language],
            "target_language": LANGUAGES[target_language],
            "tone": TONE_GUIDANCE.get(tone, tone),
            "document_profile": analysis,
            "previous_context": previous,
            "units": [
                {
                    "id": unit.unit_id,
                    "segments": unit.segments,
                    "protected_segment_indices": unit.metadata.get(
                        "protected_segment_indices", []
                    ),
                }
                for unit in batch
            ],
            "output_schema": {"translations": schema},
        }
        system = (
            "You are an elite professional translator. Transfer every meaning accurately, then write as a native "
            "author in the target language. Never translate literally when native syntax requires restructuring. "
            "Do not add, omit, explain, censor, or summarize. Preserve numbers, URLs, names, negation, ambiguity, "
            "and formatting boundaries. Segments are formatting boundaries inside complete paragraphs: use the full "
            "paragraph as context, and return exactly the same number of segments for every id. Return JSON only."
        )
        try:
            data = self._json_call(
                model_key,
                [
                    {"role": "system", "content": system},
                    {
                        "role": "user",
                        "content": json.dumps(payload, ensure_ascii=False),
                    },
                ],
                max_tokens=int(self.options.get("reserved_output_tokens", 1536)),
                temperature=float(self.options.get("temperature", 0.15)),
            )
        except ModelCapacityError:
            if len(batch) <= 1:
                raise
            midpoint = len(batch) // 2
            left_units = list(batch[:midpoint])
            right_units = list(batch[midpoint:])
            left = self._translate_batch(
                model_key,
                left_units,
                source_language,
                target_language,
                tone,
                analysis,
                previous,
            )
            continued_context = list(previous) + [
                {"source": unit.text, "target": "".join(left[unit.unit_id])}
                for unit in left_units
            ]
            right = self._translate_batch(
                model_key,
                right_units,
                source_language,
                target_language,
                tone,
                analysis,
                continued_context,
            )
            return {**left, **right}
        translations = data.get("translations", {})
        if not isinstance(translations, dict):
            raise TranslatorError("Translation response has no translations object.")
        integrity_retries = int(self.options.get("max_retries", 2))
        for attempt in range(integrity_retries + 1):
            try:
                return {
                    unit.unit_id: normalize_translation_segments(
                        unit, translations.get(unit.unit_id)
                    )
                    for unit in batch
                }
            except SegmentIntegrityError:
                if attempt >= integrity_retries:
                    raise
                correction = dict(payload)
                correction["correction"] = (
                    "Your previous response changed protected formatting boundaries. "
                    "Return exactly one output string for every input segment and keep protected indices separate."
                )
                data = self._json_call(
                    model_key,
                    [
                        {"role": "system", "content": system},
                        {
                            "role": "user",
                            "content": json.dumps(correction, ensure_ascii=False),
                        },
                    ],
                    max_tokens=int(self.options.get("reserved_output_tokens", 1536)),
                    temperature=0.0,
                )
                translations = data.get("translations", {})
                if not isinstance(translations, dict):
                    raise TranslatorError(
                        "Translation correction has no translations object."
                    )
        raise SegmentIntegrityError("Protected segment validation failed.")

    def _refresh_analysis(
        self,
        model_key: str,
        analysis: dict[str, Any],
        batch: Sequence[TranslationUnit],
        drafts: dict[str, list[str]],
    ) -> dict[str, Any]:
        recent = [
            {"source": unit.text, "translation": "".join(drafts[unit.unit_id])}
            for unit in batch
        ]
        payload = {
            "current_summary": analysis.get("rolling_summary")
            or analysis.get("summary", ""),
            "known_terms": analysis.get("terms", []),
            "recent_units": recent,
        }
        data = self._json_call(
            model_key,
            [
                {
                    "role": "system",
                    "content": (
                        "Maintain compact translation memory for a long document. Return JSON with summary and terms. "
                        "The summary must retain entities, relationships, argument state, character voices, unresolved "
                        "references, and decisions needed later. terms contains only stable source/target pairs."
                    ),
                },
                {"role": "user", "content": json.dumps(payload, ensure_ascii=False)},
            ],
            max_tokens=min(640, int(self.options.get("reserved_output_tokens", 1536))),
            temperature=0.05,
        )
        updated = dict(analysis)
        if data.get("summary"):
            updated["rolling_summary"] = str(data["summary"])[:4000]
        if isinstance(data.get("terms"), list):
            combined: list[Any] = []
            seen: set[str] = set()
            for term in [*analysis.get("terms", []), *data["terms"]]:
                if not isinstance(term, dict):
                    continue
                key = json.dumps(term, ensure_ascii=False, sort_keys=True)
                if key not in seen:
                    seen.add(key)
                    combined.append(term)
                if len(combined) >= 120:
                    break
            updated["terms"] = combined
        return updated

    def _edit_batch(
        self,
        model_key: str,
        batch: Sequence[TranslationUnit],
        drafts: dict[str, list[str]],
        source_language: str,
        target_language: str,
        tone: str,
        analysis: dict[str, Any],
    ) -> dict[str, list[str]]:
        payload = {
            "source_language": LANGUAGES[source_language],
            "target_language": LANGUAGES[target_language],
            "tone": TONE_GUIDANCE.get(tone, tone),
            "profile": analysis,
            "units": [
                {
                    "id": unit.unit_id,
                    "source": unit.segments,
                    "draft": drafts[unit.unit_id],
                    "protected_segment_indices": unit.metadata.get(
                        "protected_segment_indices", []
                    ),
                }
                for unit in batch
            ],
        }
        system = (
            "You are a bilingual senior editor and fidelity checker. Compare source and draft, then return a final, "
            "publication-ready target-language version. Improve native syntax, rhetoric, terminology, dialogue voice, "
            "and flow as appropriate, while preserving every fact, number, URL, name, negation, intensity, and intended "
            'ambiguity. Return JSON {"translations": {id: [segments]}} with exactly the original segment counts.'
        )
        try:
            data = self._json_call(
                model_key,
                [
                    {"role": "system", "content": system},
                    {
                        "role": "user",
                        "content": json.dumps(payload, ensure_ascii=False),
                    },
                ],
                max_tokens=int(self.options.get("reserved_output_tokens", 1536)),
                temperature=float(self.options.get("editor_temperature", 0.1)),
            )
        except ModelCapacityError:
            if len(batch) <= 1:
                raise
            midpoint = len(batch) // 2
            left = self._edit_batch(
                model_key,
                batch[:midpoint],
                drafts,
                source_language,
                target_language,
                tone,
                analysis,
            )
            right = self._edit_batch(
                model_key,
                batch[midpoint:],
                drafts,
                source_language,
                target_language,
                tone,
                analysis,
            )
            return {**left, **right}
        translations = data.get("translations", {})
        if not isinstance(translations, dict):
            raise TranslatorError("Editor response has no translations object.")
        integrity_retries = int(self.options.get("max_retries", 2))
        for attempt in range(integrity_retries + 1):
            try:
                result: dict[str, list[str]] = {}
                for unit in batch:
                    value = translations.get(unit.unit_id, drafts[unit.unit_id])
                    result[unit.unit_id] = normalize_translation_segments(unit, value)
                return result
            except SegmentIntegrityError:
                if attempt >= integrity_retries:
                    raise
                correction = dict(payload)
                correction["correction"] = (
                    "Keep every protected formatting segment separate and return exactly "
                    "one output string per source segment for every unit."
                )
                data = self._json_call(
                    model_key,
                    [
                        {"role": "system", "content": system},
                        {
                            "role": "user",
                            "content": json.dumps(correction, ensure_ascii=False),
                        },
                    ],
                    max_tokens=int(self.options.get("reserved_output_tokens", 1536)),
                    temperature=0.0,
                )
                translations = data.get("translations", {})
                if not isinstance(translations, dict):
                    raise TranslatorError(
                        "Editor correction has no translations object."
                    )
        raise SegmentIntegrityError("Protected editor segment validation failed.")

    def _fidelity_batch(
        self,
        model_key: str,
        batch: Sequence[TranslationUnit],
        finals: dict[str, list[str]],
        source_language: str,
        target_language: str,
        analysis: dict[str, Any],
    ) -> dict[str, list[str]]:
        payload = {
            "source_language": LANGUAGES[source_language],
            "target_language": LANGUAGES[target_language],
            "document_terms": analysis.get("terms", []),
            "units": [
                {
                    "id": unit.unit_id,
                    "source": unit.text,
                    "translation": "".join(finals[unit.unit_id]),
                }
                for unit in batch
            ],
        }
        system = (
            "Act only as a bilingual translation auditor, not as an editor. For every unit compare source and target "
            "for omissions, additions, mistranslation, names/entities, dates/numbers, negation, modality, intensity, "
            'terminology, ambiguity, register, and character voice. Return JSON {"verdicts": {id: '
            '{"ok": boolean, "issues": [{"severity": "critical|major|minor", '
            '"code": string, "detail": string}]}}}. Use ok=true only when meaning and intended effect are preserved.'
        )
        try:
            data = self._json_call(
                model_key,
                [
                    {"role": "system", "content": system},
                    {
                        "role": "user",
                        "content": json.dumps(payload, ensure_ascii=False),
                    },
                ],
                max_tokens=min(
                    768, int(self.options.get("reserved_output_tokens", 1536))
                ),
                temperature=0.0,
            )
        except ModelCapacityError:
            if len(batch) <= 1:
                raise
            midpoint = len(batch) // 2
            return {
                **self._fidelity_batch(
                    model_key,
                    batch[:midpoint],
                    finals,
                    source_language,
                    target_language,
                    analysis,
                ),
                **self._fidelity_batch(
                    model_key,
                    batch[midpoint:],
                    finals,
                    source_language,
                    target_language,
                    analysis,
                ),
            }
        verdicts = data.get("verdicts", {})
        result: dict[str, list[str]] = {}
        for unit in batch:
            verdict = verdicts.get(unit.unit_id) if isinstance(verdicts, dict) else None
            if not isinstance(verdict, dict):
                result[unit.unit_id] = ["verdict: missing fidelity verdict"]
                continue
            issues: list[str] = []
            raw_issues = verdict.get("issues", [])
            if isinstance(raw_issues, list):
                for issue in raw_issues:
                    if isinstance(issue, dict):
                        severity = str(issue.get("severity", "major")).lower()
                        if severity not in {"critical", "major"}:
                            continue
                        code = str(issue.get("code", "fidelity"))
                        detail = str(issue.get("detail", "meaning differs"))
                        issues.append(f"{code}: {detail}")
                    elif isinstance(issue, str):
                        issues.append(issue)
            if verdict.get("ok") is not True and not issues:
                issues.append("fidelity: auditor rejected the translation")
            result[unit.unit_id] = issues
        return result

    def _repair_unit(
        self,
        model_key: str,
        unit: TranslationUnit,
        candidate: list[str],
        flags: list[str],
        source_language: str,
        target_language: str,
    ) -> list[str]:
        payload = {
            "source_language": LANGUAGES[source_language],
            "target_language": LANGUAGES[target_language],
            "source": unit.segments,
            "candidate": candidate,
            "failed_checks": flags,
        }
        system = (
            "Repair the translation only where needed. Restore every missing or changed number, URL, email, name, "
            'or meaning while keeping natural target-language prose. Return JSON {"translations": {"unit": [segments]}}.'
        )
        data = self._json_call(
            model_key,
            [
                {"role": "system", "content": system},
                {"role": "user", "content": json.dumps(payload, ensure_ascii=False)},
            ],
            max_tokens=int(self.options.get("reserved_output_tokens", 1536)),
            temperature=0.05,
        )
        return normalize_translation_segments(
            unit, data.get("translations", {}).get("unit", candidate)
        )

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
        translator_key = self.profile["translator"]
        reviewer_key = self.profile.get("reviewer", translator_key)
        self.progress("analysis", 5, "در حال تحلیل موضوع، مخاطب و لحن سند")
        analysis = state.get("analysis") or self._analyze(
            nonempty, source_language, target_language, tone
        )
        if custom_instruction.strip():
            analysis["user_instruction"] = custom_instruction.strip()
        self.checkpoint("analysis", {"analysis": analysis})

        fixed_context = json.dumps(analysis, ensure_ascii=False)
        # The editor sends source and draft together, so source units use at most half
        # the otherwise available input budget from the beginning.
        budget = max(64, self._batch_budget(translator_key, fixed_context) // 2)
        working_units, split_mapping = prepare_units_for_budget(
            nonempty,
            lambda text: self.models.count_tokens(translator_key, text),
            budget,
        )
        batches = adaptive_batches(
            working_units,
            lambda text: self.models.count_tokens(translator_key, text),
            budget,
            int(self.profile.get("max_units_per_request", 6)),
        )
        manifest = build_checkpoint_manifest(
            self.config, self.profile, working_units, budget
        )
        drafts, finals = reusable_checkpoint_translations(state, manifest)
        self.checkpoint(
            "manifest",
            {
                "analysis": analysis,
                "manifest": manifest,
                "drafts": drafts,
                "finals": finals,
            },
        )
        previous: list[dict[str, str]] = []
        refresh_every = max(1, int(self.options.get("summary_refresh_batches", 6)))
        for index, batch in enumerate(batches):
            self._check_cancelled()
            pending = [unit for unit in batch if unit.unit_id not in drafts]
            if pending:
                translated = self._translate_batch(
                    translator_key,
                    pending,
                    source_language,
                    target_language,
                    tone,
                    analysis,
                    previous,
                )
                drafts.update(translated)
                self.checkpoint(
                    "draft",
                    {"analysis": analysis, "manifest": manifest, "drafts": drafts},
                )
            for unit in batch:
                previous.append(
                    {"source": unit.text, "target": "".join(drafts[unit.unit_id])}
                )
            previous = previous[-int(self.options.get("previous_units", 3)) :]
            if (
                pending
                and (index + 1) % refresh_every == 0
                and index + 1 < len(batches)
            ):
                try:
                    analysis = self._refresh_analysis(
                        translator_key, analysis, batch, drafts
                    )
                    self.checkpoint(
                        "context",
                        {"analysis": analysis, "manifest": manifest, "drafts": drafts},
                    )
                except TranslatorError:
                    # A compact-memory refresh is helpful but must not discard completed translation work.
                    pass
            self.progress(
                "translation",
                15 + int(45 * (index + 1) / max(1, len(batches))),
                f"ترجمه اولیه: بخش {index + 1} از {len(batches)}",
            )

        if reviewer_key != translator_key:
            self.models.unload()
        fixed_context = json.dumps(analysis, ensure_ascii=False)
        edit_budget = max(64, self._batch_budget(reviewer_key, fixed_context) // 2)
        edit_batches = adaptive_batches(
            working_units,
            lambda text: self.models.count_tokens(reviewer_key, text),
            edit_budget,
            int(self.profile.get("max_units_per_request", 6)),
        )
        for index, batch in enumerate(edit_batches):
            self._check_cancelled()
            pending = [unit for unit in batch if unit.unit_id not in finals]
            if pending:
                edited = self._edit_batch(
                    reviewer_key,
                    pending,
                    drafts,
                    source_language,
                    target_language,
                    tone,
                    analysis,
                )
                finals.update(edited)
                self.checkpoint(
                    "editing",
                    {
                        "analysis": analysis,
                        "manifest": manifest,
                        "drafts": drafts,
                        "finals": finals,
                    },
                )
            self.progress(
                "editing",
                60 + int(20 * (index + 1) / max(1, len(edit_batches))),
                f"ویرایش زبان مقصد: بخش {index + 1} از {len(edit_batches)}",
            )

        fidelity: dict[str, list[str]] = {}
        for index, batch in enumerate(edit_batches):
            self._check_cancelled()
            fidelity.update(
                self._fidelity_batch(
                    reviewer_key,
                    batch,
                    finals,
                    source_language,
                    target_language,
                    analysis,
                )
            )
            self.progress(
                "quality",
                80 + int(7 * (index + 1) / max(1, len(edit_batches))),
                f"داوری وفاداری: بخش {index + 1} از {len(edit_batches)}",
            )

        def collect_issues(
            candidates: Sequence[TranslationUnit],
        ) -> dict[str, list[str]]:
            issues: dict[str, list[str]] = {}
            for candidate in candidates:
                target = "".join(finals[candidate.unit_id])
                combined = [
                    *quality_flags(candidate.text, target),
                    *glossary_flags(candidate.text, target, analysis),
                    *fidelity.get(candidate.unit_id, []),
                ]
                if combined:
                    issues[candidate.unit_id] = list(dict.fromkeys(combined))
            return issues

        units_by_id = {unit.unit_id: unit for unit in working_units}
        unresolved = collect_issues(working_units)
        repair_attempts = max(1, int(self.options.get("quality_repair_attempts", 2)))
        for attempt in range(repair_attempts):
            if not unresolved:
                break
            repaired_units: list[TranslationUnit] = []
            for index, (unit_id, flags) in enumerate(list(unresolved.items())):
                self._check_cancelled()
                unit = units_by_id[unit_id]
                finals[unit_id] = self._repair_unit(
                    reviewer_key,
                    unit,
                    finals[unit_id],
                    flags,
                    source_language,
                    target_language,
                )
                repaired_units.append(unit)
                self.progress(
                    "quality",
                    87 + int(3 * (index + 1) / max(1, len(unresolved))),
                    f"اصلاح کیفیت: تلاش {attempt + 1}، مورد {index + 1} از {len(unresolved)}",
                )
            fidelity = self._fidelity_batch(
                reviewer_key,
                repaired_units,
                finals,
                source_language,
                target_language,
                analysis,
            )
            unresolved = collect_issues(repaired_units)
            self.checkpoint(
                "quality_repair",
                {
                    "analysis": analysis,
                    "manifest": manifest,
                    "drafts": drafts,
                    "finals": finals,
                },
            )
        if unresolved:
            details = "; ".join(
                f"{unit_id}: {', '.join(flags[:3])}"
                for unit_id, flags in list(unresolved.items())[:5]
            )
            raise TranslatorError(
                "Quality control could not resolve critical fidelity issues: " + details
            )
        self.checkpoint(
            "quality",
            {
                "analysis": analysis,
                "manifest": manifest,
                "drafts": drafts,
                "finals": finals,
            },
        )
        self.progress("quality", 90, "کنترل وفاداری و یکپارچگی تکمیل شد")
        self.models.unload()
        reassembled = reassemble_split_units(
            nonempty, working_units, split_mapping, finals
        )
        for unit in units:
            reassembled.setdefault(unit.unit_id, list(unit.segments))
        return reassembled


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
        family = fallback if direction == "rtl" else (source_font or fallback)
        family = family.replace("'", "")
        margin = "0 0 .7em 0" if include_block_spacing else "0"
        escaped = html.escape(text).replace("\n", "<br>")
        return (
            f"<div style=\"font-family:'{family}',sans-serif;font-size:{size}pt;color:{color};"
            f'direction:{direction};text-align:{align};line-height:1.3;margin:{margin}">'
            f"{escaped}</div>"
        )

    @staticmethod
    def _html_fits(page: Any, rect: Any, markup: str, min_scale: float) -> bool:
        try:
            result = page.insert_htmlbox(rect, markup, scale_low=min_scale)
        except (RuntimeError, ValueError):
            return False
        return bool(result and result[0] >= 0)

    def _trial_group_fits(
        self, width: float, height: float, markup: str, min_scale: float
    ) -> bool:
        trial = self.fitz.open()
        try:
            page = trial.new_page(width=width, height=height)
            rect = self.fitz.Rect(36, 36, width - 36, height - 36)
            return self._html_fits(page, rect, markup, min_scale)
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
    ) -> None:
        width, height = source_page.rect.width, source_page.rect.height
        readable_scale = max(0.85, min_scale)
        fragments: list[tuple[TranslationUnit, str]] = []
        pending = list(entries)
        while pending:
            unit, text = pending.pop(0)
            markup = self._entry_html(unit, text, target_language, fallback)
            if self._trial_group_fits(width, height, markup, readable_scale):
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
                width, height, markup, readable_scale
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
            if not self._html_fits(page, rect, markup, readable_scale):
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
                if not self._html_fits(trial_page, rect, markup, min_scale):
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
                    if not self._html_fits(page, rect, markup, min_scale):
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
                    created_at REAL NOT NULL,
                    updated_at REAL NOT NULL
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
                    custom_instruction, output_path, output_name, created_at, updated_at
                ) VALUES (?, 'queued', 'queued', 0, ?, '', ?, ?, ?, ?, ?, ?, ?, ?, '', '', ?, ?)
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
                ),
            )
        return job_id

    def get(self, job_id: str) -> dict[str, Any] | None:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT * FROM jobs WHERE id = ?", (job_id,)
            ).fetchone()
        return dict(row) if row else None

    def public(self, job: dict[str, Any]) -> dict[str, Any]:
        return {key: job.get(key) for key in JOB_PUBLIC_FIELDS}

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

    def request_cancel(self, job_id: str) -> bool:
        job = self.get(job_id)
        if not job or job["status"] in TERMINAL_JOB_STATES:
            return False
        self.update(job_id, cancel_requested=1, message="درخواست لغو ثبت شد")
        return True

    def is_cancelled(self, job_id: str) -> bool:
        job = self.get(job_id)
        return bool(job and job.get("cancel_requested"))

    def recoverable(self) -> list[str]:
        with self._connect() as connection:
            rows = connection.execute(
                "SELECT id FROM jobs WHERE status IN ('queued', 'running') ORDER BY created_at"
            ).fetchall()
            connection.execute(
                "UPDATE jobs SET status='queued', stage='queued', message='بازیابی پس از راه‌اندازی مجدد' "
                "WHERE status='running'"
            )
        return [str(row["id"]) for row in rows]

    def active_count(self) -> int:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT COUNT(*) AS count FROM jobs WHERE status IN ('queued', 'running')"
            ).fetchone()
        return int(row["count"] if row else 0)

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
        if not job or job["status"] in TERMINAL_JOB_STATES:
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
        try:
            profile_name, profile = choose_profile(self.config)
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

            model_manager = ModelManager(self.config, profile)
            pipeline = TranslationPipeline(
                self.config,
                profile,
                model_manager,
                progress=report,
                cancelled=lambda: self.store.is_cancelled(job_id),
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
            except TranslatorError:
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
                model_manager = ModelManager(self.config, low)
                pipeline = TranslationPipeline(
                    self.config,
                    low,
                    model_manager,
                    progress=report,
                    cancelled=lambda: self.store.is_cancelled(job_id),
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
            self.store.update(
                job_id,
                status="completed",
                stage="completed",
                progress=100,
                message="ترجمه آماده دانلود است",
                output_path=str(output_path),
                output_name=filename,
                error="",
            )
        except JobCancelled:
            self.store.update(
                job_id, status="cancelled", stage="cancelled", message="ترجمه لغو شد"
            )
        except Exception as exc:  # noqa: BLE001 - worker boundary must persist every failure
            self.store.update(
                job_id,
                status="failed",
                stage="failed",
                message="پردازش متوقف شد",
                error=str(exc),
            )
        finally:
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
        "llama_cpp": importlib.util.find_spec("llama_cpp") is not None,
        "lxml": importlib.util.find_spec("lxml") is not None,
        "python_docx": importlib.util.find_spec("docx") is not None,
        "pymupdf": importlib.util.find_spec("pymupdf") is not None,
    }


def build_doctor_report(
    config: dict[str, Any], probe_model: bool = False
) -> dict[str, Any]:
    """Build one truthful readiness report for both the CLI and web endpoint."""
    dependencies = _dependency_status()
    backend = detect_runtime_backend()
    models = {
        key: {
            "path": value.get("path"),
            "exists": Path(str(value.get("path", ""))).is_file(),
        }
        for key, value in config.get("models", {}).items()
    }
    errors: list[str] = []
    required = ("flask", "llama_cpp", "lxml", "python_docx", "pymupdf")
    missing = [name for name in required if not dependencies.get(name)]
    if missing:
        errors.append("Missing Python packages: " + ", ".join(missing))

    try:
        profile_name, profile = choose_profile(config)
    except TranslatorError as exc:
        profile_name, profile = "unavailable", {}
        errors.append(str(exc))

    probe: dict[str, Any] = {"requested": bool(probe_model), "ok": None}
    if probe_model:
        if profile_name == "unavailable" or backend == "unavailable":
            probe.update(
                {"ok": False, "error": "Runtime is not ready for a model probe."}
            )
        else:
            manager = ModelManager(config, profile)
            model_key = str(profile.get("translator", ""))
            try:
                prompt = 'Reply with a JSON object containing only: {"ok": true}'
                token_count = manager.count_tokens(model_key, prompt)
                response = manager.generate(
                    model_key,
                    [{"role": "user", "content": prompt}],
                    max_tokens=32,
                    temperature=0.0,
                )
                if not response.strip():
                    raise TranslatorError("The model probe returned an empty response.")
                probe.update(
                    {"ok": True, "model": model_key, "prompt_tokens": token_count}
                )
            except Exception as exc:  # noqa: BLE001 - diagnostic boundary
                probe.update({"ok": False, "model": model_key, "error": str(exc)})
                errors.append(f"Model probe failed: {exc}")
            finally:
                manager.unload()

    healthy = not errors and backend != "unavailable"
    return {
        "status": "ok" if healthy else "incomplete",
        "platform": platform.platform(),
        "memory_gb": round(total_memory_bytes() / 1024**3, 1),
        "backend": backend,
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
    :root{--ink:#192522;--muted:#63716d;--paper:#f4f1e8;--card:#fffdf7;--line:#d9d4c6;--accent:#0f766e;--accent2:#115e59;--danger:#b42318}
    *{box-sizing:border-box}body{margin:0;background:radial-gradient(circle at 15% 0,#d7ebe2 0,transparent 35%),var(--paper);color:var(--ink);font-family:Tahoma,"Noto Sans Arabic",sans-serif;min-height:100vh}
    main{width:min(920px,calc(100% - 28px));margin:42px auto}.eyebrow{color:var(--accent);font-weight:700;font-size:.82rem;letter-spacing:.08em}h1{font-size:clamp(2rem,5vw,3.5rem);margin:.3rem 0 .7rem;line-height:1.15}header p{color:var(--muted);max-width:680px;line-height:1.9}
    .card{background:rgba(255,253,247,.94);border:1px solid var(--line);border-radius:22px;padding:24px;box-shadow:0 18px 55px rgba(25,37,34,.09);margin-top:25px}.grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:16px}.full{grid-column:1/-1}
    label{display:block;font-weight:700;font-size:.9rem;margin-bottom:8px}select,textarea,input[type=file],input[type=text]{width:100%;border:1px solid var(--line);border-radius:12px;background:#fff;padding:12px;color:var(--ink);font:inherit}textarea{min-height:160px;resize:vertical;line-height:1.8}select:focus,textarea:focus,input:focus{outline:3px solid rgba(15,118,110,.14);border-color:var(--accent)}
    .switch{display:flex;gap:8px;margin-bottom:12px}.switch button{background:#ece8dc;color:var(--muted);border:0;padding:8px 14px;border-radius:999px;cursor:pointer}.switch button.active{background:var(--ink);color:white}.actions{display:flex;gap:10px;align-items:center;margin-top:18px}button.primary,a.download{background:var(--accent);color:#fff;border:0;border-radius:12px;padding:12px 20px;font:inherit;font-weight:700;cursor:pointer;text-decoration:none}button.primary:hover,a.download:hover{background:var(--accent2)}button.secondary{background:transparent;border:1px solid var(--line);border-radius:12px;padding:11px 18px;color:var(--danger);cursor:pointer}.hidden{display:none!important}
    #statusCard{overflow:hidden}.statusline{display:flex;justify-content:space-between;gap:16px;align-items:center}.percent{font-size:2rem;font-weight:800;color:var(--accent)}.track{height:12px;background:#e5e1d6;border-radius:999px;margin:18px 0;overflow:hidden}.bar{height:100%;width:0;background:linear-gradient(90deg,var(--accent),#2dd4bf);transition:width .35s ease}.muted{color:var(--muted);font-size:.9rem}.error{color:var(--danger);white-space:pre-wrap;line-height:1.7}.footer{margin:22px 4px;color:var(--muted);font-size:.8rem}.ltr{direction:ltr;text-align:left}
    @media(max-width:700px){main{margin:24px auto}.card{padding:18px}.grid{grid-template-columns:1fr}.full{grid-column:auto}.statusline{align-items:flex-start}.actions{flex-wrap:wrap}}
  </style>
</head>
<body><main>
  <header><div class="eyebrow">LOCAL · PRIVATE · ADAPTIVE</div><h1>مترجم محلی اسناد</h1><p>ترجمهٔ طبیعی و چندمرحله‌ای با مدل‌های GGUF شما؛ همراه با حفظ ساختار Word، بازسازی PDF و نمایش زندهٔ پیشرفت.</p></header>
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
  <section class="card hidden" id="statusCard"><div class="statusline"><div><strong id="stage">آماده‌سازی</strong><div class="muted" id="message"></div></div><div class="percent"><span id="percent">0</span>٪</div></div><div class="track"><div class="bar" id="bar"></div></div><div id="error" class="error"></div><div class="actions"><button id="cancel" class="secondary" type="button">لغو</button><a id="download" class="download hidden">دانلود فایل ترجمه‌شده</a></div></section>
  <div class="footer">فایل‌ها و مدل‌ها روی همین دستگاه پردازش می‌شوند.</div>
</main><script>
const form=document.getElementById('jobForm'), inputType=document.getElementById('inputType'), textInput=document.getElementById('textInput'), fileInput=document.getElementById('fileInput'), csrfToken=document.querySelector('meta[name="csrf-token"]').content;let currentJob=null, events=null;
document.querySelectorAll('.switch button').forEach(btn=>btn.onclick=()=>{document.querySelectorAll('.switch button').forEach(x=>x.classList.remove('active'));btn.classList.add('active');inputType.value=btn.dataset.type;const isText=btn.dataset.type==='text';textInput.classList.toggle('hidden',!isText);fileInput.classList.toggle('hidden',isText);document.getElementById('file').accept=btn.dataset.type==='docx'?'.docx':btn.dataset.type==='pdf'?'.pdf':'';});
function showState(job){document.getElementById('statusCard').classList.remove('hidden');document.getElementById('stage').textContent=job.stage||job.status;document.getElementById('message').textContent=job.message||'';document.getElementById('percent').textContent=job.progress||0;document.getElementById('bar').style.width=(job.progress||0)+'%';document.getElementById('error').textContent=job.error||'';const done=job.status==='completed';document.getElementById('download').classList.toggle('hidden',!done);document.getElementById('cancel').classList.toggle('hidden',['completed','failed','cancelled'].includes(job.status));if(done)document.getElementById('download').href='/api/jobs/'+job.id+'/download';}
function connectJob(jobId){currentJob=jobId;localStorage.setItem('translatorJobId',jobId);if(events)events.close();events=new EventSource('/api/jobs/'+jobId+'/events');events.addEventListener('progress',e=>{const job=JSON.parse(e.data);showState(job);if(['completed','failed','cancelled'].includes(job.status))events.close()});events.onerror=()=>{document.getElementById('message').textContent='ارتباط وضعیت موقتاً قطع شد؛ اتصال مجدد خودکار انجام می‌شود.'};}
form.onsubmit=async e=>{e.preventDefault();document.getElementById('formMessage').textContent='';document.getElementById('error').textContent='';const res=await fetch('/api/jobs',{method:'POST',body:new FormData(form)});const data=await res.json();if(!res.ok){document.getElementById('formMessage').textContent=data.error||'خطا در ثبت کار';return}document.getElementById('statusCard').classList.remove('hidden');connectJob(data.job_id);};
document.getElementById('cancel').onclick=async()=>{if(currentJob)await fetch('/api/jobs/'+currentJob+'/cancel',{method:'POST',headers:{'X-CSRF-Token':csrfToken}})};
const savedJob=localStorage.getItem('translatorJobId');if(savedJob){fetch('/api/jobs/'+savedJob).then(async res=>{if(!res.ok){localStorage.removeItem('translatorJobId');return}const job=await res.json();showState(job);if(!['completed','failed','cancelled'].includes(job.status))connectJob(savedJob);else currentJob=savedJob;});}
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
        return Response(
            INDEX_HTML.replace("__CSRF_TOKEN__", html.escape(token, quote=True)),
            mimetype="text/html",
        )

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
            while True:
                job = store.get(job_id)
                if not job:
                    break
                if job["updated_at"] != last_updated:
                    payload = json.dumps(store.public(job), ensure_ascii=False)
                    yield f"event: progress\ndata: {payload}\n\n"
                    last_updated = job["updated_at"]
                else:
                    yield ": keep-alive\n\n"
                if job["status"] in TERMINAL_JOB_STATES:
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
    parser = argparse.ArgumentParser(description="Local LLM document translator")
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
        help="With --doctor, also load the selected GGUF and run a tiny inference",
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
