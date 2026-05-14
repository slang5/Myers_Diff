"""
Core data types for the Myers diff implementation.

This module keeps the algorithm data model in one place so the search,
backtracking, and formatting code can stay small and readable.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal, Optional, TypeAlias


# Edit labels used by the diff algorithm.
DiffOperation: TypeAlias = Literal["equal", "delete", "insert"]

# Labels used by the derived change-block view.
ChangeOperation: TypeAlias = Literal["equal", "delete", "insert", "replace"]

# Small aliases that make the implementation signatures easier to read.
LineNumber: TypeAlias = int
Diagonal: TypeAlias = int


# One line of input with its original 1-based line number.
@dataclass(frozen=True, slots=True)
class DiffLine:
	number: LineNumber
	text: str


# One edit produced by the diff.
@dataclass(frozen=True, slots=True)
class DiffEdit:
	type: DiffOperation
	old_line: Optional[DiffLine]
	new_line: Optional[DiffLine]


# A higher-level grouped change used to make replacements easier to read.
@dataclass(frozen=True, slots=True)
class DiffChange:
	type: ChangeOperation
	old_lines: list[DiffLine]
	new_lines: list[DiffLine]


# One saved snapshot of the Myers frontier at a given edit depth.
@dataclass(frozen=True, slots=True)
class AlgorithmState:
	v: dict[Diagonal, int]
	d: int


# Final diff result returned by the public API.
@dataclass(frozen=True, slots=True)
class DiffOutput:
	edits: list[DiffEdit]
	edit_distance: int
	# The grouped view is derived from the raw edits so later renderers can reuse it.
	change_blocks: DiffChanges = field(default_factory=list)


# Helpful type aliases for the rest of the implementation.
Frontier: TypeAlias = dict[Diagonal, int]
Trace: TypeAlias = list[AlgorithmState]
DiffLines: TypeAlias = list[DiffLine]
DiffEdits: TypeAlias = list[DiffEdit]
DiffChanges: TypeAlias = list[DiffChange]

