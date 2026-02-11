"""YAML spec-oriented CLI commands (validate/preview/verify)."""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Protocol

from .cli_project import apply_config, expand_inputs
from .placeholder import fill_placeholders
from .solver import ConstraintSolver
from .spec_pipeline import load_yaml, parse_spec, preview_spec, validate_spec
from .text_metrics import TextMetrics


class InputArgs(Protocol):
    input: list[str]


class VerifyArgs(InputArgs, Protocol):
    detail: bool
    json: str | None
    config: str | None


class ValidateArgs(InputArgs, Protocol):
    config: str | None


def cmd_validate(args: ValidateArgs) -> int:
    """Validate schema for YAML files."""
    apply_config(args.config)

    input_files = expand_inputs(args.input)
    if not input_files:
        print("No input files found", file=sys.stderr)
        return 1

    all_valid = True
    for path in input_files:
        data = load_yaml(str(path))
        errors, warnings = validate_spec(data)
        if errors:
            all_valid = False
            print(f"{path}:\n  - " + "\n  - ".join(errors))
        elif warnings:
            print(f"{path}:\n  - " + "\n  - ".join(warnings))
        else:
            print(f"{path}: OK")
    return 0 if all_valid else 1


def cmd_preview(args: InputArgs) -> int:
    """Preview table structure without generating PPTX."""
    input_files = expand_inputs(args.input)
    if not input_files:
        print("No input files found", file=sys.stderr)
        return 1

    for path in input_files:
        data = load_yaml(str(path))
        print(preview_spec(data))
    return 0


def cmd_verify(args: VerifyArgs) -> int:
    """Run layout solver + report without generating PPTX."""
    apply_config(args.config)

    input_files = expand_inputs(args.input)
    if not input_files:
        print("No input files found", file=sys.stderr)
        return 1

    metrics = TextMetrics()
    solver = ConstraintSolver(metrics)

    for path in input_files:
        data = load_yaml(str(path))
        errors, warnings = validate_spec(data)
        if errors:
            print(f"{path}:\n  - " + "\n  - ".join(errors), file=sys.stderr)
            return 1
        if warnings:
            print(f"{path}:\n  - " + "\n  - ".join(warnings))

        spec, area, options, placeholders = parse_spec(data)
        if placeholders:
            spec = fill_placeholders(spec)

        _, report = solver.solve(spec, area, options)
        print(f"{path}: {report.to_text(detail=args.detail)}")

        if args.json:
            output_path = Path(args.json)
            output_path.write_text(json.dumps(report.to_json(), indent=2))

    return 0
