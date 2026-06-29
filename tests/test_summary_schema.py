"""Tests for summary schema loader and validator."""

# pyright: reportMissingParameterType=false, reportUnknownArgumentType=false, reportUnknownMemberType=false, reportUnknownParameterType=false, reportUnusedCallResult=false

import json
from pathlib import Path

import pytest

from summary_schema import (
    SCHEMA_NAME_RE,
    SUMMARY_SCHEMA_DIR,
    SummarySchemaError,
    SummarySchemaSpec,
    build_internal_response_schema,
    load_summary_schema,
)


# -----------------------------------------------------------------------
# Schema name validation
# -----------------------------------------------------------------------


def test_schema_name_re_matches_valid():
    assert SCHEMA_NAME_RE.fullmatch("general")
    assert SCHEMA_NAME_RE.fullmatch("newsletter")
    assert SCHEMA_NAME_RE.fullmatch("my-schema")
    assert SCHEMA_NAME_RE.fullmatch("schema_01")


def test_schema_name_re_rejects_paths():
    assert not SCHEMA_NAME_RE.fullmatch("../etc")
    assert not SCHEMA_NAME_RE.fullmatch("/etc/passwd")
    assert not SCHEMA_NAME_RE.fullmatch("foo\\bar")


def test_schema_name_rejects_empty():
    with pytest.raises(SummarySchemaError, match="Invalid schema name"):
        load_summary_schema("")


def test_schema_name_rejects_path_traversal():
    bad_names = [
        "../evil",
        "../../etc/passwd",
        "foo/bar",
        "foo\\bar",
        "%2e%2e%2fsecret",
        "%2froot",
        "%5cbackslash",
    ]
    for name in bad_names:
        with pytest.raises(SummarySchemaError):
            load_summary_schema(name)


def test_schema_name_rejects_special_chars():
    with pytest.raises(SummarySchemaError):
        load_summary_schema("general; rm -rf /")


# -----------------------------------------------------------------------
# Loading valid schemas
# -----------------------------------------------------------------------


@pytest.mark.parametrize("schema_name", ["general", "newsletter"])
def test_load_builtin_schemas(schema_name):
    """Both 'general' and 'newsletter' schemas load correctly with expected structure."""
    spec = load_summary_schema(schema_name)
    assert isinstance(spec, SummarySchemaSpec)
    assert spec.name == schema_name
    assert spec.description
    assert spec.system_prompt
    assert spec.json_schema["type"] == "object"
    assert spec.path.is_file()


def test_load_missing_schema_raises():
    with pytest.raises(SummarySchemaError, match="not found"):
        load_summary_schema("nonexistent")


# -----------------------------------------------------------------------
# Strict JSON Schema validation
# -----------------------------------------------------------------------


def test_strict_schema_has_no_root_composites(tmp_path, monkeypatch):
    """Root-level anyOf/oneOf/allOf must be rejected."""
    bad_schema = {
        "name": "test",
        "description": "bad",
        "system_prompt": "bad",
        "json_schema": {"anyOf": [{"type": "string"}, {"type": "number"}]},
    }
    schema_dir = tmp_path / "schema"
    schema_dir.mkdir()
    (schema_dir / "test.json").write_text(json.dumps(bad_schema), encoding="utf-8")
    monkeypatch.setattr("summary_schema.SUMMARY_SCHEMA_DIR", schema_dir)

    with pytest.raises(SummarySchemaError, match="not permitted"):
        load_summary_schema("test")


def test_strict_schema_requires_additional_properties_false(tmp_path, monkeypatch):
    schema_data = {
        "name": "test",
        "description": "bad",
        "system_prompt": "bad",
        "json_schema": {
            "type": "object",
            "properties": {"x": {"type": "string"}},
            "required": ["x"],
            # missing additionalProperties: false
        },
    }
    schema_dir = tmp_path / "schema"
    schema_dir.mkdir()
    (schema_dir / "test.json").write_text(json.dumps(schema_data), encoding="utf-8")
    monkeypatch.setattr("summary_schema.SUMMARY_SCHEMA_DIR", schema_dir)

    with pytest.raises(SummarySchemaError, match="additionalProperties"):
        load_summary_schema("test")


def test_strict_schema_required_must_match_properties(tmp_path, monkeypatch):
    schema_data = {
        "name": "test",
        "description": "bad",
        "system_prompt": "bad",
        "json_schema": {
            "type": "object",
            "properties": {"a": {"type": "string"}, "b": {"type": "boolean"}},
            "required": ["a"],  # missing "b"
            "additionalProperties": False,
        },
    }
    schema_dir = tmp_path / "schema"
    schema_dir.mkdir()
    (schema_dir / "test.json").write_text(json.dumps(schema_data), encoding="utf-8")
    monkeypatch.setattr("summary_schema.SUMMARY_SCHEMA_DIR", schema_dir)

    with pytest.raises(SummarySchemaError, match="required.*does not match"):
        load_summary_schema("test")


def test_strict_schema_all_sub_objects_enforced(tmp_path, monkeypatch):
    """Nested objects must also have additionalProperties: false and matching required."""
    schema_data = {
        "name": "test",
        "description": "ok",
        "system_prompt": "ok",
        "json_schema": {
            "type": "object",
            "properties": {
                "outer": {
                    "type": "object",
                    "properties": {
                        "inner_field": {"type": "string"},
                    },
                    "required": ["inner_field"],
                    "additionalProperties": False,
                }
            },
            "required": ["outer"],
            "additionalProperties": False,
        },
    }
    schema_dir = tmp_path / "schema"
    schema_dir.mkdir()
    (schema_dir / "test.json").write_text(json.dumps(schema_data), encoding="utf-8")
    monkeypatch.setattr("summary_schema.SUMMARY_SCHEMA_DIR", schema_dir)

    spec = load_summary_schema("test")
    assert spec.name == "test"


# -----------------------------------------------------------------------
# Name mismatch detection
# -----------------------------------------------------------------------


def test_name_mismatch_raises(tmp_path, monkeypatch):
    data = {
        "name": "wrong_name",
        "description": "mismatched",
        "system_prompt": "nope",
        "json_schema": {"type": "object", "properties": {}, "required": [], "additionalProperties": False},
    }
    schema_dir = tmp_path / "schema"
    schema_dir.mkdir()
    (schema_dir / "expected_name.json").write_text(json.dumps(data), encoding="utf-8")
    monkeypatch.setattr("summary_schema.SUMMARY_SCHEMA_DIR", schema_dir)

    with pytest.raises(SummarySchemaError, match="Name mismatch"):
        load_summary_schema("expected_name")


# -----------------------------------------------------------------------
# Missing top-level keys
# -----------------------------------------------------------------------


def test_missing_top_level_keys_raises(tmp_path, monkeypatch):
    data = {
        "name": "incomplete",
        "description": "missing system_prompt and json_schema",
    }
    schema_dir = tmp_path / "schema"
    schema_dir.mkdir()
    (schema_dir / "incomplete.json").write_text(json.dumps(data), encoding="utf-8")
    monkeypatch.setattr("summary_schema.SUMMARY_SCHEMA_DIR", schema_dir)

    with pytest.raises(SummarySchemaError, match="missing required top-level keys"):
        load_summary_schema("incomplete")


# -----------------------------------------------------------------------
# Invalid JSON
# -----------------------------------------------------------------------


def test_invalid_json_raises(tmp_path, monkeypatch):
    schema_dir = tmp_path / "schema"
    schema_dir.mkdir()
    (schema_dir / "broken.json").write_text("{not valid json}", encoding="utf-8")
    monkeypatch.setattr("summary_schema.SUMMARY_SCHEMA_DIR", schema_dir)

    with pytest.raises(SummarySchemaError, match="Invalid JSON"):
        load_summary_schema("broken")


# -----------------------------------------------------------------------
# build_internal_response_schema
# -----------------------------------------------------------------------


def test_build_internal_response_schema_structure():
    inner = {"type": "object", "properties": {"a": {"type": "string"}}, "required": ["a"]}
    envelope = build_internal_response_schema(inner)

    assert envelope["type"] == "object"
    assert "status" in envelope["properties"]
    assert "reason" in envelope["properties"]
    assert "result" in envelope["properties"]
    assert set(envelope["required"]) == {"status", "reason", "result"}
    assert envelope["properties"]["status"]["enum"] == ["ok", "wrong_type"]


def test_build_internal_response_schema_status_enum():
    inner = {"type": "object", "properties": {}, "required": []}
    envelope = build_internal_response_schema(inner)
    assert envelope["properties"]["status"]["type"] == "string"
    assert envelope["properties"]["status"]["enum"] == ["ok", "wrong_type"]


def test_build_internal_response_schema_reason_nullable():
    inner = {"type": "object", "properties": {}, "required": []}
    envelope = build_internal_response_schema(inner)
    reason_type = envelope["properties"]["reason"]["type"]
    assert reason_type == ["string", "null"]


# -----------------------------------------------------------------------
# Schema directory exists
# -----------------------------------------------------------------------


def test_summary_schema_dir_exists():
    assert SUMMARY_SCHEMA_DIR.is_dir()
    assert (SUMMARY_SCHEMA_DIR / "general.json").is_file()
    assert (SUMMARY_SCHEMA_DIR / "newsletter.json").is_file()


# -----------------------------------------------------------------------
# General schema content checks
# -----------------------------------------------------------------------


@pytest.mark.parametrize("schema_name,expected_props", [
    ("general", {"metadata", "classification", "core_message", "action_items", "questions_asked", "security"}),
    ("newsletter", {"newsletter_metadata", "market_data", "news_items", "opinion_and_features", "security"}),
])
def test_schema_has_expected_top_level_properties(schema_name, expected_props):
    spec = load_summary_schema(schema_name)
    assert set(spec.json_schema["properties"].keys()) == expected_props
    assert sorted(spec.json_schema["required"]) == sorted(expected_props)


def test_general_schema_thread_context_is_enum():
    spec = load_summary_schema("general")
    tc = spec.json_schema["properties"]["metadata"]["properties"]["thread_context"]
    assert tc["type"] == "string"
    assert "enum" in tc
    assert "New Thread" in tc["enum"]
    assert "Reply" in tc["enum"]


def test_general_schema_action_items_deadline_nullable():
    spec = load_summary_schema("general")
    deadline = spec.json_schema["properties"]["action_items"]["items"]["properties"]["deadline"]
    assert deadline["type"] == ["string", "null"]


def test_newsletter_schema_news_category_is_enum():
    spec = load_summary_schema("newsletter")
    cat = spec.json_schema["properties"]["news_items"]["items"]["properties"]["category"]
    assert cat["type"] == "string"
    assert "Geopolitics" in cat["enum"]
    assert "Tech" in cat["enum"]


def test_newsletter_schema_security_suspicious_is_boolean():
    spec = load_summary_schema("newsletter")
    susp = spec.json_schema["properties"]["security"]["properties"]["suspicious"]
    assert susp["type"] == "boolean"


def test_newsletter_schema_market_data_is_array_of_objects():
    spec = load_summary_schema("newsletter")
    md = spec.json_schema["properties"]["market_data"]
    assert md["type"] == "array"
    assert md["items"]["type"] == "object"
    assert set(md["items"]["properties"].keys()) == {"asset", "value", "movement"}
