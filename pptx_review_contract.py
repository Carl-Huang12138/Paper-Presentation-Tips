#!/usr/bin/env python3
"""Repository compatibility entry; implementation lives in the portable skill."""
from pathlib import Path as _Path
_implementation = _Path(__file__).resolve().parent / "skills/academic-pptx-review/scripts/pptx_review_contract.py"
__file__ = str(_implementation)
exec(compile(_implementation.read_bytes(), __file__, "exec"), globals())
