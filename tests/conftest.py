from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent))

from build_decks import build_chaos_deck, build_simple_deck  # noqa: E402


@pytest.fixture
def chaos(tmp_path):
    path = tmp_path / "chaos.pptx"
    build_chaos_deck(str(path))
    return path


@pytest.fixture
def simple(tmp_path):
    path = tmp_path / "simple.pptx"
    build_simple_deck(str(path))
    return path
