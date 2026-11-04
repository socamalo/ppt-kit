"""Saving: write to a temp file, re-open it to prove it is valid, then swap in.

On Windows an antivirus / search indexer will occasionally hold a freshly
written file open for a few hundred milliseconds, which makes the final
``os.replace`` fail with a sharing violation. That is a transient lock, not a
bad file, so the swap retries with backoff instead of handing the agent a
spurious failure.
"""

from __future__ import annotations

import os
import shutil
import time
import uuid
from typing import Callable, TypeVar

from pptx import Presentation

from .errors import WriteError

T = TypeVar("T")
RETRIES = 6
BASE_DELAY = 0.12


def _retry(fn: Callable[[], T], what: str) -> T:
    last: Exception | None = None
    for attempt in range(RETRIES):
        try:
            return fn()
        except OSError as e:  # sharing violation, permission, busy
            last = e
            time.sleep(BASE_DELAY * (attempt + 1))
        except Exception as e:  # noqa: BLE001 - not a lock, stop trying
            raise WriteError(f"{what}: {e}") from e
    raise WriteError(
        f"{what} (still failing after {RETRIES} attempts; "
        f"the file is probably open in PowerPoint): {last}"
    )


def save(deck, out: str | None = None, *, backup: bool = False) -> str:
    target = os.path.abspath(out or deck.path)
    directory = os.path.dirname(target) or "."
    if not os.path.isdir(directory):
        raise WriteError(f"output directory does not exist: {directory}")

    tmp = os.path.join(
        directory, f".{os.path.basename(target)}.{uuid.uuid4().hex[:8]}.pptk-tmp.pptx"
    )
    try:
        _retry(lambda: deck.prs.save(tmp), "could not write the file")
    except WriteError:
        _cleanup(tmp)
        raise

    # a file we cannot re-open is not a file we ship
    try:
        _retry(lambda: Presentation(tmp), "the written file cannot be re-opened; original left untouched")
    except WriteError:
        _cleanup(tmp)
        raise

    if backup and os.path.exists(target):
        try:
            shutil.copy2(target, target + ".bak")
        except Exception:
            pass

    try:
        _retry(lambda: os.replace(tmp, target), "could not replace the target file")
    except WriteError as e:
        _cleanup(tmp)
        raise WriteError(e.message) from e
    return target


def _cleanup(path: str) -> None:
    try:
        if os.path.exists(path):
            os.remove(path)
    except Exception:
        pass
