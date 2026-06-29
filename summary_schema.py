"""Schema loader and validator for email summary schemas.

Loads JSON schema configs from the ``schema/`` directory, validates them
against structural requirements (strict JSON Schema, required top-level keys),
and provides path-traversal protection when resolving schema names.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from config import project_root_from_module

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

SCHEMA_NAME_RE = re.compile(r"^[a-zA-Z0-9][a-zA-Z0-9_-]*$")
SUMMARY_SCHEMA_DIR = project_root_from_module() / "schema"

# ---------------------------------------------------------------------------
# Exceptions
# ---------------------------------------------------------------------------


class SummarySchemaError(RuntimeError):
    """Raised when a summary schema cannot be loaded or validated."""


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class SummarySchemaSpec:
    """Loaded and validated summary schema specification."""

    name: str
    description: str
    system_prompt: str
    json_schema: dict[str, Any]
    path: Path = field(default_factory=Path)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def load_summary_schema(schema_name: str = "general") -> SummarySchemaSpec:
    """Load and validate a summary schema by *schema_name*.

    Parameters
    ----------
    schema_name:
        Alphanumeric identifier (e.g. ``"general"``). Must match ``SCHEMA_NAME_RE``;
        path separators, dots, absolute paths, URL-encoded sequences, and Unicode
        homoglyphs are rejected.

    Returns
    -------
    SummarySchemaSpec with validated fields.

    Raises
    ------
    SummarySchemaError
        On any validation failure (path traversal, missing file, bad JSON, etc.).
    """
    _validate_schema_name(schema_name)

    schema_path = SUMMARY_SCHEMA_DIR / f"{schema_name}.json"

    if not schema_path.is_file():
        raise SummarySchemaError(f"Schema file not found: {schema_path}")

    try:
        raw = json.loads(schema_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise SummarySchemaError(f"Invalid JSON in {schema_path}: {exc}") from exc

    return _validate_schema_content(raw, schema_name, schema_path)


def build_internal_response_schema(selected_schema: dict[str, Any]) -> dict[str, Any]:
    """Wrap *selected_schema* in an envelope for internal LLM responses.

    The resulting schema adds a ``status`` enum (``ok`` | ``wrong_type``),
    a nullable ``reason`` string, and wraps the user schema as a nullable
    ``result`` object.
    """
    return {
        "type": "object",
        "properties": {
            "status": {
                "type": "string",
                "enum": ["ok", "wrong_type"],
            },
            "reason": {"type": ["string", "null"]},
            "result": {
                "type": ["object", "null"],
                "properties": selected_schema.get("properties", {}),
                "required": selected_schema.get("required", []),
                "additionalProperties": False,
            },
        },
        "required": ["status", "reason", "result"],
        "additionalProperties": False,
    }


# ---------------------------------------------------------------------------
# Internal validation helpers
# ---------------------------------------------------------------------------


def _validate_schema_name(name: str) -> None:
    """Reject schema names that could leak files outside the schema directory."""
    if not name or not SCHEMA_NAME_RE.match(name):
        raise SummarySchemaError(
            f"Invalid schema name '{name}': must be alphanumeric with optional hyphens/underscores"
        )

    # Catch path traversal attempts that slipped past regex
    dangerous_chars = ("..", "/", "\\", "%2e", "%2f", "%5c")
    lower = name.lower()
    for seq in dangerous_chars:
        if seq in lower:
            raise SummarySchemaError(
                f"Rejected schema name containing path-traversal sequence: '{seq}'"
            )


TOP_LEVEL_KEYS = {"name", "description", "system_prompt", "json_schema"}


def _validate_schema_content(
    data: dict, expected_name: str, schema_path: Path
) -> SummarySchemaSpec:
    """Ensure *data* has all required keys and a strict JSON Schema payload."""
    missing = TOP_LEVEL_KEYS - set(data.keys())
    if missing:
        raise SummarySchemaError(
            f"Schema missing required top-level keys: {sorted(missing)}"
        )

    if data["name"] != expected_name:
        raise SummarySchemaError(
            f"Name mismatch: file is '{expected_name}' but schema declares '{data['name']}'"
        )

    js = data["json_schema"]
    _assert_strict_json_schema(js)

    return SummarySchemaSpec(
        name=data["name"],
        description=data["description"],
        system_prompt=data["system_prompt"],
        json_schema=js,
        path=schema_path,
    )


def _assert_strict_json_schema(schema: Any, path: str = "root") -> None:
    """Recursively verify that *schema* is a strict JSON Schema object.

    Rules enforced:
    - Every node that is a dict must have ``"type"`` (or ``"oneOf"``/``"items"``).
    - Root and every sub-object must have ``"type": "object"`` (no root anyOf/oneOf/allOf).
    - Every object node must have ``"additionalProperties": false``.
    - Every object node whose properties are defined must list them all in ``"required"``.
    """
    if not isinstance(schema, dict):
        raise SummarySchemaError(f"{path}: expected an object, got {type(schema).__name__}")

    # Root-level compound keywords are forbidden
    if path == "root":
        for kw in ("anyOf", "oneOf", "allOf"):
            if kw in schema:
                raise SummarySchemaError(
                    f"Root-level '{kw}' is not permitted in strict schemas"
                )

    schema_type = schema.get("type")

    # Objects must be fully strict
    if schema_type == "object" or (
        schema_type is None
        and "properties" in schema
        and "type" not in schema
    ):
        props = schema.get("properties", {})
        required = schema.get("required", [])

        if set(required) != set(props.keys()):
            raise SummarySchemaError(
                f"{path}: 'required' ({sorted(required)}) does not match "
                f"'properties' ({sorted(props.keys())})"
            )

        if schema.get("additionalProperties") is not False:
            raise SummarySchemaError(
                f"{path}: objects must declare 'additionalProperties': false"
            )

        for prop_name, prop_schema in props.items():
            _assert_strict_json_schema(prop_schema, f"{path}.{prop_name}")

    elif schema_type == "array":
        items = schema.get("items")
        if items is None:
            raise SummarySchemaError(f"{path}: array must define 'items'")
        _assert_strict_json_schema(items, f"{path}.items")

    elif isinstance(schema_type, str) or (isinstance(schema_type, list)):
        pass  # leaf type – string, boolean, number, integer, or union like ["string","null"]

    else:
        raise SummarySchemaError(f"{path}: unrecognised schema structure")
