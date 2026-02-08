# Typing Findings Log

Document issues discovered during type migration that require investigation
or cannot be fixed with simple annotations.

## Format

```markdown
### [YYYY-MM-DD] Finding Title
**File**: path/to/file.py:123
**Category**: [design-issue | api-mismatch | missing-stubs | unfixable]
**Severity**: [low | medium | high]

Description of the issue and why it couldn't be fixed.
```

---

## Findings

### [2026-02-06] python-pptx stubs use `typing.Self` causing `Unknown` under pyright pythonVersion=3.9
**File**: `pptx-stubs/dml/color.pyi` (RGBColor.from_string) / `clean_slides/constants.py`
**Category**: missing-stubs
**Severity**: medium

`python-pptx-stubs` annotates some APIs using `typing.Self` (Python 3.11+). With this project checked using pyright `pythonVersion: 3.9`, those return types can degrade to `Unknown` (e.g. `RGBColor.from_string`), creating cascading `reportUnknownMemberType` errors.

Preferred approach: isolate at boundaries with typed helpers/wrappers that return concrete types (e.g. parse hex + construct `RGBColor(r,g,b)`), or add a local stub override that imports `Self` from `typing_extensions`.

### [2026-02-06] lxml element type is `_Element` (pyright `reportPrivateUsage` conflict)
**File**: `clean_slides/xml_helpers.py`
**Category**: api-mismatch
**Severity**: low

`lxml-stubs` models element instances as `lxml.etree._Element` (leading underscore). Pyright's `reportPrivateUsage` flags using `_Element` in our annotations, but `_Element` is the concrete runtime type returned by `etree.Element(...)` and required for correct typing of `.append()/.find()/.remove()`.

Mitigation: add per-file directive `# pyright: reportPrivateUsage=false` in `xml_helpers.py` and use `_Element` types there. Alternative would be falling back to `Any` or incorrect `ElementBase` annotations.

### [2026-02-06] pdf2image typing incomplete; local stub added
**File**: `typings/pdf2image/__init__.pyi`, `clean_slides/screenshot.py`
**Category**: missing-stubs
**Severity**: low

`pdf2image.convert_from_path` has incomplete typing (partial `Unknown` types) which triggers strict pyright errors. Added a small local stub under `typings/` and configured `pyrightconfig.json` with `stubPath: "typings"`.

### [2026-02-06] Pyright strict treats `list[Unknown]` as “partially unknown” even for `len()`/`list()`
**File**: `clean_slides/editor.py` (`_normalize_content`)
**Category**: design-issue
**Severity**: low

When working with dynamic JSON/YAML structures, `isinstance(x, (list, tuple))` can narrow `x` to `list[Unknown] | tuple[Unknown, ...]`. In strict mode this triggers `reportUnknownArgumentType` for otherwise-safe operations like `len(x)` and `list(x)`.

Mitigation: introduce a small `TypeGuard` like `_is_sequence(value) -> TypeGuard[Sequence[object]]` to force a concrete element type, or parse dynamic structures into typed models earlier to avoid `Unknown` leaking into general-purpose code.

### [2026-02-06] lxml `etree.SubElement(..., **attrs)` is hard to type-check in strict mode
**File**: `clean_slides/renderer.py`
**Category**: api-mismatch
**Severity**: low

`lxml-stubs` defines `SubElement(parent, tag, attrib=None, nsmap=None, **extra)`.

When passing a plain `dict[str, str]` via `**attrs`, pyright must assume it *might* contain keys like `attrib` or `nsmap`, and will complain that a `str` value is not compatible with those parameter types (`reportArgumentType`).

Mitigation: prefer `etree.SubElement(..., attrib=attrs_dict)` (or a `TypedDict` + `Unpack`) rather than `**attrs_dict`.

### [2026-02-06] python-pptx PEP695 generic stubs can leave `Unknown` type parameters (Chart/GraphicFrame)
**File**: `clean_slides/inspect_pptx.py` (`inspect_chart`)
**Category**: api-mismatch
**Severity**: low

`python-pptx-stubs` models `GraphicFrame` and `Chart` using PEP695 type parameters. When narrowing a `BaseShape` with `isinstance(shape, GraphicFrame)`, the type parameters can remain `Unknown`, which strict pyright reports (e.g. `chart: Chart[Unknown, Unknown]` → `reportUnknownVariableType`).

Mitigation: introduce a small `Protocol` boundary for the subset of chart APIs we use (e.g. `_chartSpace`, `has_legend`) and narrow to that protocol using a `TypeGuard`, so code doesn’t depend on concrete generic parameters.

(Add entries as issues are discovered)
