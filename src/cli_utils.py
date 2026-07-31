"""CLI shared utilities — structured error output and JSON summaries."""

from __future__ import annotations

import json
import sys
import time
from typing import Any


def fatal_error(module: str, error: Exception, **context):
    """Print structured JSON error to stderr and exit with code 1.

    Every module's main() catches Exception and calls this, so the agent
    always receives machine-parseable error output regardless of what broke.
    """
    payload: dict[str, Any] = {
        "module": module,
        "error_type": type(error).__name__,
        "error_message": str(error),
    }
    payload.update(context)
    print(json.dumps(payload, ensure_ascii=False, default=str), file=sys.stderr)
    sys.exit(1)


class RunTimer:
    """Wall-clock timer for run summaries."""

    def __init__(self):
        self._start = time.monotonic()

    @property
    def elapsed(self) -> float:
        return time.monotonic() - self._start


def print_json_summary(data: dict[str, Any]) -> None:
    """Print a structured JSON run summary to stdout.

    Called by every module at the end of a successful or partial run
    when --json flag is set.  Agent parses this for automated follow-up.
    """
    print(json.dumps(data, ensure_ascii=False, default=str))
