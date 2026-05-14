"""
Core Myers diff implementation.

This module follows the plan in Logic.md:
- split_lines: turn text into 1-based DiffLine objects
- find_shortest_edit: build the Myers frontier trace
- backtrack: rebuild the edit script from the trace
- diff_replace_block: run a second, token-based diff inside replace blocks

Example usage:

    old_text = "Hello,  world!\nABC"
    new_text = "Hello world?\nABC"

    result = compute_diff(old_text, new_text)
    print(result.edit_distance)
    print(format_diff(result))
    print(format_patch(result))
    print([block.type for block in result.change_blocks])

    token_diffs = build_replace_token_diffs(result)
    for item in token_diffs:
        print([(edit.type, edit.old_line.text if edit.old_line else None,
                edit.new_line.text if edit.new_line else None)
               for edit in item.token_diff.edits])
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from Myers_Class_Type import (
    AlgorithmState,
    DiffChange,
    DiffEdit,
    DiffEdits,
    DiffLine,
    DiffLines,
    DiffChanges,
    DiffOutput,
    Frontier,
    Trace,
)

__all__ = [
    "split_lines",
    "find_shortest_edit",
    "backtrack",
    "compute_diff",
    "diff_lines",
    "build_change_blocks",
    "diff_replace_block",
    "build_replace_token_diffs",
    "ReplaceTokenDiff",
    "format_diff",
    "format_patch",
]


# Split text into words, whitespace, and single punctuation marks.
# This keeps spaces visible and makes punctuation changes easier to read.
TOKEN_PATTERN = re.compile(r"\s+|[\w]+|[^\w\s]")


@dataclass(frozen=True, slots=True)
class ReplaceTokenDiff:
    """A replace block together with its inner token-level diff."""

    block: DiffChange
    token_diff: DiffOutput


def split_lines(text: str) -> DiffLines:
    """Split text into 1-based line records using '\n' exactly."""
    if not isinstance(text, str):
        raise TypeError(f"Expected string input, got {type(text).__name__}")

    # Match the TypeScript behavior: the empty string produces no lines.
    if text == "":
        return []

    lines: DiffLines = []
    for number, line_text in enumerate(text.split("\n"), start=1):
        lines.append(DiffLine(number=number, text=line_text))
    return lines


def find_shortest_edit(old_lines: DiffLines, new_lines: DiffLines) -> Trace:
    """Build the Myers frontier trace for the two line sequences."""
    old_count = len(old_lines)
    new_count = len(new_lines)

    if old_count == 0 and new_count == 0:
        return []

    max_depth = old_count + new_count
    frontier: Frontier = {0: 0}
    trace: Trace = []

    for depth in range(max_depth + 1):
        # Store the frontier before exploring this depth so backtracking can replay it later.
        trace.append(AlgorithmState(v=frontier.copy(), d=depth))

        for diagonal in range(-depth, depth + 1, 2):
            left_x = frontier.get(diagonal - 1, 0)
            right_x = frontier.get(diagonal + 1, 0)

            # Choose the path that reaches the furthest x coordinate.
            if diagonal == -depth or (diagonal != depth and left_x < right_x):
                x = right_x
            else:
                x = left_x + 1

            y = x - diagonal

            # Follow the snake: consume equal lines on the same diagonal.
            while x < old_count and y < new_count and old_lines[x].text == new_lines[y].text:
                x += 1
                y += 1

            frontier[diagonal] = x

            # Once both sequences are fully consumed, the trace is complete.
            if x >= old_count and y >= new_count:
                return trace

    raise RuntimeError("Impossible: no path found")


def backtrack(old_lines: DiffLines, new_lines: DiffLines, trace: Trace) -> DiffEdits:
    """Rebuild the edit script from the saved frontier snapshots."""
    x = len(old_lines)
    y = len(new_lines)
    edits: DiffEdits = []

    # Walk the trace backwards so we can reconstruct the path from end to start.
    for state in reversed(trace):
        frontier = state.v
        depth = state.d
        diagonal = x - y

        left_x = frontier.get(diagonal - 1, 0)
        right_x = frontier.get(diagonal + 1, 0)

        # Use the same choice rule as the forward pass.
        if diagonal == -depth or (diagonal != depth and left_x < right_x):
            previous_diagonal = diagonal + 1
        else:
            previous_diagonal = diagonal - 1

        previous_x = frontier.get(previous_diagonal, 0)
        previous_y = previous_x - previous_diagonal

        # Rebuild the equal lines that were consumed by the snake.
        while x > previous_x and y > previous_y:
            x -= 1
            y -= 1
            edits.append(DiffEdit(type="equal", old_line=old_lines[x], new_line=new_lines[y]))

        # Then record the actual edit that moved between diagonals.
        if depth > 0:
            if x == previous_x:
                edits.append(DiffEdit(type="insert", old_line=None, new_line=new_lines[previous_y]))
            elif y == previous_y:
                edits.append(DiffEdit(type="delete", old_line=old_lines[previous_x], new_line=None))

            x = previous_x
            y = previous_y

    edits.reverse()
    return edits


def build_change_blocks(edits: DiffEdits) -> DiffChanges:
    """Group raw edits into human-friendly blocks.

    This keeps the Myers core intact while making complex changes easier to read:
    - consecutive equal lines are grouped into one equal block
    - consecutive deletes followed by inserts become a replace block
    - pure deletes and pure inserts remain separate blocks
    """

    blocks: DiffChanges = []
    equal_old_lines: DiffLines = []
    equal_new_lines: DiffLines = []
    change_old_lines: DiffLines = []
    change_new_lines: DiffLines = []

    def flush_equal_block() -> None:
        nonlocal equal_old_lines, equal_new_lines

        if equal_old_lines or equal_new_lines:
            blocks.append(
                DiffChange(
                    type="equal",
                    old_lines=equal_old_lines,
                    new_lines=equal_new_lines,
                )
            )
            equal_old_lines = []
            equal_new_lines = []

    def flush_change_block() -> None:
        nonlocal change_old_lines, change_new_lines

        if not change_old_lines and not change_new_lines:
            return

        if change_old_lines and change_new_lines:
            change_type = "replace"
        elif change_old_lines:
            change_type = "delete"
        else:
            change_type = "insert"

        blocks.append(
            DiffChange(
                type=change_type,
                old_lines=change_old_lines,
                new_lines=change_new_lines,
            )
        )
        change_old_lines = []
        change_new_lines = []

    for edit in edits:
        if edit.type == "equal":
            # Close any pending change block before starting a new equal block.
            flush_change_block()
            if edit.old_line is None or edit.new_line is None:
                raise ValueError("equal edits must contain both old and new lines")
            equal_old_lines.append(edit.old_line)
            equal_new_lines.append(edit.new_line)
        elif edit.type == "delete":
            # Close the equal block when the diff enters a changed region.
            flush_equal_block()
            if edit.old_line is None:
                raise ValueError("delete edits must contain an old line")
            change_old_lines.append(edit.old_line)
        elif edit.type == "insert":
            # Close the equal block when the diff enters a changed region.
            flush_equal_block()
            if edit.new_line is None:
                raise ValueError("insert edits must contain a new line")
            change_new_lines.append(edit.new_line)
        else:
            raise ValueError(f"Unsupported edit type: {edit.type}")

    # Flush whatever was left in the current block at the end of the script.
    flush_equal_block()
    flush_change_block()
    return blocks


def _join_lines(lines: DiffLines) -> str:
    """Join a block of lines back into one text string."""
    return "\n".join(line.text for line in lines)


def _tokenize_text(text: str) -> DiffLines:
    """Split text into token records for the inner diff.

    The pattern keeps whitespace intact and splits punctuation into separate
    tokens when possible, which is useful for Word output.
    """

    tokens: DiffLines = []
    for number, token_text in enumerate(TOKEN_PATTERN.findall(text), start=1):
        tokens.append(DiffLine(number=number, text=token_text))
    return tokens


def diff_replace_block(block: DiffChange) -> DiffOutput:
    """Run the secondary token diff for a single replace block."""

    if block.type != "replace":
        raise ValueError("diff_replace_block() expects a replace block")

    # Rebuild the old and new block text exactly, then tokenize it.
    old_text = _join_lines(block.old_lines)
    new_text = _join_lines(block.new_lines)

    old_tokens = _tokenize_text(old_text)
    new_tokens = _tokenize_text(new_text)

    # Reuse the same Myers engine on the token stream.
    if not old_tokens and not new_tokens:
        return DiffOutput(edits=[], edit_distance=0, change_blocks=[])

    trace = find_shortest_edit(old_tokens, new_tokens)
    edits = backtrack(old_tokens, new_tokens, trace)
    token_blocks = build_change_blocks(edits)
    return DiffOutput(edits=edits, edit_distance=len(trace) - 1, change_blocks=token_blocks)


def build_replace_token_diffs(result: DiffOutput) -> list[ReplaceTokenDiff]:
    """Build token-level diffs for every replace block in a line-level result."""

    token_diffs: list[ReplaceTokenDiff] = []

    for block in result.change_blocks:
        # Only replace blocks need the inner diff.
        if block.type != "replace":
            continue

        token_diffs.append(
            ReplaceTokenDiff(
                block=block,
                token_diff=diff_replace_block(block),
            )
        )

    return token_diffs


def _build_aligned_lines(edits: DiffEdits) -> tuple[list[DiffLine | None], list[DiffLine | None]]:
    """Align old and new lines by edit position.

    We keep None values for missing sides so blank lines remain real content and
    do not get confused with a missing line in the patch formatter.
    """

    old_lines: list[DiffLine | None] = []
    new_lines: list[DiffLine | None] = []

    for edit in edits:
        if edit.type == "equal":
            if edit.old_line is None or edit.new_line is None:
                raise ValueError("equal edits must contain both old and new lines")
            old_lines.append(edit.old_line)
            new_lines.append(edit.new_line)
        elif edit.type == "delete":
            if edit.old_line is None:
                raise ValueError("delete edits must contain an old line")
            old_lines.append(edit.old_line)
            new_lines.append(None)
        elif edit.type == "insert":
            if edit.new_line is None:
                raise ValueError("insert edits must contain a new line")
            old_lines.append(None)
            new_lines.append(edit.new_line)
        else:
            raise ValueError(f"Unsupported edit type: {edit.type}")

    return old_lines, new_lines


def format_diff(result: DiffOutput) -> str:
    """Render the raw Myers edits in a compact, human-readable form."""
    lines: list[str] = []

    for edit in result.edits:
        # Keep the raw edit script visible: this is the closest view of the algorithm.
        if edit.type == "equal":
            lines.append(f"  {edit.old_line.text if edit.old_line else ''}")
        elif edit.type == "delete":
            lines.append(f"- {edit.old_line.text if edit.old_line else ''}")
        elif edit.type == "insert":
            lines.append(f"+ {edit.new_line.text if edit.new_line else ''}")
        else:
            raise ValueError(f"Unsupported edit type: {edit.type}")

    return "\n".join(lines)


def format_patch(
    result: DiffOutput,
    old_file_name: str = "old",
    new_file_name: str = "new",
    old_timestamp: str = "",
    new_timestamp: str = "",
) -> str:
    """Render the diff as a unified patch with small context hunks."""
    lines: list[str] = []
    lines.append(f"--- {old_file_name}{'\t' + old_timestamp if old_timestamp else ''}")
    lines.append(f"+++ {new_file_name}{'\t' + new_timestamp if new_timestamp else ''}")

    old_lines, new_lines = _build_aligned_lines(result.edits)
    total = max(len(old_lines), len(new_lines))

    # First identify the changed spans; later we expand them with context.
    change_regions: list[tuple[int, int]] = []
    region_start: int | None = None

    for index in range(total):
        old_line = old_lines[index] if index < len(old_lines) else None
        new_line = new_lines[index] if index < len(new_lines) else None

        is_changed = (
            old_line is None
            or new_line is None
            or old_line.text != new_line.text
        )

        if is_changed:
            if region_start is None:
                region_start = index
        elif region_start is not None:
            change_regions.append((region_start, index - 1))
            region_start = None

    if region_start is not None:
        change_regions.append((region_start, total - 1))

    for start, end in change_regions:
        context_before = 3
        context_after = 3
        hunk_start = max(0, start - context_before)
        hunk_end = min(total - 1, end + context_after)

        # Unified diff line numbers are 1-based and count only the visible lines on each side.
        old_start = 1 + sum(1 for line in old_lines[:hunk_start] if line is not None)
        new_start = 1 + sum(1 for line in new_lines[:hunk_start] if line is not None)
        old_count = sum(1 for line in old_lines[hunk_start : hunk_end + 1] if line is not None)
        new_count = sum(1 for line in new_lines[hunk_start : hunk_end + 1] if line is not None)

        lines.append(f"@@ -{old_start},{old_count} +{new_start},{new_count} @@")

        for index in range(hunk_start, hunk_end + 1):
            old_line = old_lines[index] if index < len(old_lines) else None
            new_line = new_lines[index] if index < len(new_lines) else None

            if old_line is not None and new_line is not None:
                if old_line.text == new_line.text:
                    lines.append(f" {old_line.text}")
                else:
                    lines.append(f"-{old_line.text}")
                    lines.append(f"+{new_line.text}")
            elif old_line is None and new_line is not None:
                lines.append(f"+{new_line.text}")
            elif old_line is not None and new_line is None:
                lines.append(f"-{old_line.text}")

    return "\n".join(lines)


def compute_diff(old_text: str, new_text: str) -> DiffOutput:
    """Compute the Myers diff for two text strings."""
    # Normalize the inputs into DiffLine records first.
    old_lines = split_lines(old_text)
    new_lines = split_lines(new_text)

    # Keep the empty-input behavior explicit and cheap.
    if not old_lines and not new_lines:
        return DiffOutput(edits=[], edit_distance=0)

    trace = find_shortest_edit(old_lines, new_lines)
    edits = backtrack(old_lines, new_lines, trace)

    # Build the derived block view once so replace regions are ready for later rendering.
    change_blocks = build_change_blocks(edits)
    return DiffOutput(edits=edits, edit_distance=len(trace) - 1, change_blocks=change_blocks)


def diff_lines(old_lines: DiffLines, new_lines: DiffLines) -> DiffOutput:
    """Compute the Myers diff for already prepared line arrays."""
    # Keep the empty-input behavior aligned with compute_diff().
    if not old_lines and not new_lines:
        return DiffOutput(edits=[], edit_distance=0)

    trace = find_shortest_edit(old_lines, new_lines)
    edits = backtrack(old_lines, new_lines, trace)

    # Return the grouped change blocks alongside the raw edits so callers do not recompute them.
    change_blocks = build_change_blocks(edits)
    return DiffOutput(edits=edits, edit_distance=len(trace) - 1, change_blocks=change_blocks)