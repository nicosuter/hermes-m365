# Hermes Agent core compatibility

This plugin requests strict JSON Schema decoding for email summaries by calling:

```python
await ctx.llm.acomplete_structured(
    ...,
    json_schema=schema,
    strict=True,
)
```

## Required Hermes Agent source support

The installed Hermes Agent core must expose a `strict: bool = False` keyword on both `PluginLlm.complete_structured()` and `PluginLlm.acomplete_structured()` in `agent/plugin_llm.py`.

The core must pass that value through `_json_response_format(...)` and set:

```python
{
    "response_format": {
        "type": "json_schema",
        "json_schema": {
            "name": "plugin_structured_output",
            "schema": json_schema,
            "strict": strict,
        },
    }
}
```

Without this core API change, the plugin call fails with `TypeError: ... got an unexpected keyword argument 'strict'`. Merely validating generated JSON afterward is insufficient: strict mode must reach the provider request so compatible providers constrain enums, required properties, and other schema rules during decoding.

Hermes installations that do not yet provide this keyword must update or patch the core before using this plugin version.
