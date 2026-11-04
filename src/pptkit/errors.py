"""Uniform JSON envelope + exit codes shared by every subcommand.

Every command answers with the same shape so an agent can branch on `ok`
without learning per-command quirks:

    {"ok": true, "command": "read", "file": "a.pptx",
     "data": {...}, "warnings": [...], "error": null}

Exit codes are part of the contract:

    0  success
    1  usage / argument error
    2  file error (missing, unreadable, not a .pptx)
    3  target not found / ambiguous reference
    4  write or validation failure
"""

from __future__ import annotations

import json
from typing import Any


EXIT_OK = 0
EXIT_USAGE = 1
EXIT_FILE = 2
EXIT_NOT_FOUND = 3
EXIT_WRITE = 4


class PptError(Exception):
    """Base error carrying the exit code the CLI should return."""

    exit_code = EXIT_USAGE
    kind = "error"

    def __init__(self, message: str, *, detail: Any = None):
        super().__init__(message)
        self.message = message
        self.detail = detail


class UsageError(PptError):
    exit_code = EXIT_USAGE
    kind = "usage"


class FileError(PptError):
    exit_code = EXIT_FILE
    kind = "file"


class NotFoundError(PptError):
    exit_code = EXIT_NOT_FOUND
    kind = "not_found"


class WriteError(PptError):
    exit_code = EXIT_WRITE
    kind = "write"


def envelope(
    command: str,
    *,
    ok: bool,
    file: str | None = None,
    data: Any = None,
    warnings: list[str] | None = None,
    error: dict | None = None,
) -> dict:
    return {
        "ok": ok,
        "command": command,
        "file": file,
        "data": data,
        "warnings": warnings or [],
        "error": error,
    }


def dump(payload: dict) -> str:
    return json.dumps(payload, ensure_ascii=False, indent=2)
