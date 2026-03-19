"""Progressive hole filling for CST CAD model simplification.

Uses AddToHistory for each pick and remove operation, exactly matching
how CST records manual operations in ModelHistory.json. This is REQUIRED
for changes to persist in the model.

Algorithm:
1. Find cone/cylinder seed faces from SAT
2. For each seed, pick faces via AddToHistory, then RemoveSelectedFaces
3. On failure, expand with topologically adjacent faces (filtered by bbox)
4. Retry up to MAX_EXPAND times
5. Skip seeds already consumed by previous holes

Key CST 2025 discovery:
- RunScript alone does NOT persist changes
- AddToHistory is REQUIRED for each pick and for the remove
- Each AddToHistory entry is a separate VBA snippet (no Sub Main wrapper)
- The VBA code inside AddToHistory uses escaped quotes ("" for embedded ")

Validates: Requirements 4.1, 4.2, 4.3, 4.4, 5.1, 5.3
"""

import logging
import os
import tempfile
from typing import Dict, List, Set, Tuple

from code.cst_connection import CSTConnection, CSTConnectionError
from code.models import FillResult, SessionSummary, SimplificationCandidate

logger = logging.getLogger(__name__)

MAX_EXPAND = 5
BBOX_SIZE_RATIO_LIMIT = 20.0
BBOX_OVERLAP_MARGIN = 2.0


def _bbox_volume(bb: Tuple) -> float:
    return (bb[3]-bb[0]) * (bb[4]-bb[1]) * (bb[5]-bb[2])


def _bbox_overlaps(bb1: Tuple, bb2: Tuple, margin: float = BBOX_OVERLAP_MARGIN) -> bool:
    """Check if two bboxes overlap or are within margin of each other."""
    return not (bb1[3] + margin < bb2[0] or bb2[3] + margin < bb1[0] or
                bb1[4] + margin < bb2[1] or bb2[4] + margin < bb1[1] or
                bb1[5] + margin < bb2[2] or bb2[5] + margin < bb1[2])


def _bbox_size_ratio(bb_neighbor: Tuple, bb_seed: Tuple) -> float:
    v_n = max(_bbox_volume(bb_neighbor), 1e-12)
    v_s = max(_bbox_volume(bb_seed), 1e-12)
    return v_n / v_s


class Simplifier:
    """Progressive hole filler for CST models.

    Uses AddToHistory for all model-modifying operations, matching
    the exact format CST uses when recording manual operations.

    Args:
        connection: An active CSTConnection with an open project.
    """

    def __init__(self, connection: CSTConnection) -> None:
        self._conn = connection
        self._out_path = os.path.join(tempfile.gettempdir(), "cst_fill.txt")
        self._out_vba = self._out_path.replace("\\", "\\\\")

    def _run_vba(self, code: str) -> str:
        """Execute VBA via RunScript and read output file."""
        result = self._conn.execute_vba(code, output_file=self._out_path)
        return result or ""

    def _add_to_history(self, caption: str, vba_code: str) -> str:
        """Execute an operation via AddToHistory so it persists.

        This is the ONLY way to make model changes stick in CST 2025.
        The vba_code is the raw VBA that goes into the history entry
        (no Sub Main wrapper needed inside AddToHistory).

        The VBA code passed to AddToHistory must use doubled quotes
        for any embedded string literals.

        Args:
            caption: History entry caption (shown in CST history list)
            vba_code: VBA code to execute and record in history

        Returns:
            Output from the execution
        """
        # Escape the vba_code for embedding inside a VBA string literal:
        # - Replace \ with \\
        # - Replace " with ""  (VBA string escaping)
        # - Replace newlines with " & vbCrLf & "
        escaped_code = vba_code.replace("\\", "\\\\")
        escaped_code = escaped_code.replace('"', '""')
        # Split into lines and join with VBA string concatenation
        lines = escaped_code.split("\n")
        vba_str = '" & vbCrLf & "'.join(lines)

        wrapper = (
            'Sub Main\n'
            f'  Open "{self._out_vba}" For Output As #1\n'
            '  On Error Resume Next\n'
            f'  AddToHistory "{caption}", "{vba_str}"\n'
            '  If Err.Number <> 0 Then\n'
            '    Print #1, "HIST_ERR: " & Err.Description\n'
            '    Err.Clear\n'
            '  Else\n'
            '    Print #1, "OK"\n'
            '  End If\n'
            '  Close #1\n'
            'End Sub\n'
        )
        return self._run_vba(wrapper)

    def _try_fill_hole(self, shape_name: str, face_ids: List[int],
                       hole_index: int) -> Tuple[bool, str]:
        """Try to fill a hole using AddToHistory for each step.

        Matches the exact sequence CST records for manual operations:
        1. Pick each face via AddToHistory
        2. RemoveSelectedFaces via AddToHistory
        3. ClearAllPicks via AddToHistory

        If RemoveSelectedFaces fails, the picks are still in history
        but the model is unchanged (CST handles this gracefully).

        Returns:
            (success, message) tuple
        """
        messages = []

        # Step 1: Pick each face via AddToHistory
        for fid in face_ids:
            pick_code = f'Pick.PickFaceFromId "{shape_name}", "{fid}"'
            result = self._add_to_history("pick face", pick_code)
            messages.append(f"pick {fid}: {result}")
            if "HIST_ERR" in (result or ""):
                # Clean up picks and abort
                self._add_to_history("clear picks", "Pick.ClearAllPicks")
                return False, f"Pick failed for face {fid}: {result}"

        # Step 2: RemoveSelectedFaces via AddToHistory
        remove_code = (
            f'With LocalModification \n'
            f'     .Reset \n'
            f'     .Name "{shape_name}"\n'
            f'     .RemoveSelectedFaces \n'
            f'End With'
        )
        caption = f"remove features of shape: {shape_name}"
        result = self._add_to_history(caption, remove_code)
        messages.append(f"remove: {result}")

        if "HIST_ERR" in (result or "") or "OK" not in (result or ""):
            # Remove failed — clear picks
            self._add_to_history("clear picks", "Pick.ClearAllPicks")
            return False, f"Remove failed: {result}"

        # Step 3: Clear picks via AddToHistory
        self._add_to_history("clear picks", "Pick.ClearAllPicks")

        return True, "; ".join(messages)

    def _expand_faces(
        self,
        current_ids: Set[int],
        seed_bbox: Tuple,
        adjacency: Dict[int, Set[int]],
        bboxes: Dict[int, Tuple],
        consumed: Set[int],
    ) -> Set[int]:
        """Expand face set with adjacent faces filtered by bbox."""
        new_ids = set(current_ids)
        for fid in list(current_ids):
            for neighbor in adjacency.get(fid, set()):
                if neighbor in new_ids or neighbor in consumed:
                    continue
                nb = bboxes.get(neighbor)
                if nb is None:
                    continue
                if not _bbox_overlaps(nb, seed_bbox):
                    continue
                if _bbox_size_ratio(nb, seed_bbox) > BBOX_SIZE_RATIO_LIMIT:
                    continue
                new_ids.add(neighbor)
        return new_ids

    def _highlight_faces(self, shape_name: str, face_ids: List[int],
                         zoom_to_bbox: Tuple = None) -> None:
        """Pick faces in CST GUI to highlight them, then zoom.

        Uses RunScript (not AddToHistory) since highlighting is
        a view operation, not a model change.
        """
        pick_lines = "\n".join(
            f'  Pick.PickFaceFromId "{shape_name}", "{fid}"'
            for fid in face_ids
        )

        zoom_steps = 0
        if zoom_to_bbox and zoom_to_bbox != (0, 0, 0, 0, 0, 0):
            dx = zoom_to_bbox[3] - zoom_to_bbox[0]
            dy = zoom_to_bbox[4] - zoom_to_bbox[1]
            dz = zoom_to_bbox[5] - zoom_to_bbox[2]
            max_dim = max(dx, dy, dz, 0.001)
            if max_dim < 1.0:
                zoom_steps = 15
            elif max_dim < 3.0:
                zoom_steps = 10
            elif max_dim < 10.0:
                zoom_steps = 5

        zoom_lines = "\n".join(['  SendKeys "+"'] * zoom_steps) if zoom_steps else ""

        code = (
            'Sub Main\n'
            '  Plot.ZoomToStructure\n'
            '  Pick.ClearAllPicks\n'
            f'{pick_lines}\n'
            f'{zoom_lines}\n'
            'End Sub\n'
        )
        try:
            self._conn.execute_vba(code)
        except CSTConnectionError:
            pass

    def fill_progressive(
        self,
        shape_name: str,
        seeds: List[int],
        adjacency: Dict[int, Set[int]],
        bboxes: Dict[int, Tuple],
        face_types: Dict[int, str],
        interactive: bool = True,
        seed_groups: List[dict] = None,
    ) -> SessionSummary:
        """Progressive hole removal across all seed faces.

        When seed_groups is provided, iterates over hole groups. Each group
        contains 'seeds' (cone face IDs) and 'loop_faces' (ALL face IDs
        forming the hole, including non-cone faces). All loop_faces are
        picked together for RemoveSelectedFaces.

        When seed_groups is not provided, falls back to treating each seed
        individually (legacy behavior).

        For each group:
        1. Highlight all group seeds in GUI and prompt user (if interactive)
        2. Try RemoveSelectedFaces with all group seeds
        3. On failure, expand with adjacent faces and retry
        4. Skip groups already consumed by previous holes

        Args:
            shape_name: CST shape name (e.g. "component1:solid1")
            seeds: List of cone/cylinder face IDs (used as fallback)
            adjacency: Face adjacency graph
            bboxes: Per-face bounding boxes
            face_types: pid -> surface type string
            interactive: If True, prompt user before each hole
            seed_groups: List of groups, each group is a list of seed IDs
                belonging to the same hole.

        Returns:
            SessionSummary with counts.
        """
        summary = SessionSummary()
        consumed: Set[int] = set()
        hole_count = 0

        # Build iteration list: groups if available, else individual seeds
        if seed_groups:
            groups = seed_groups
        else:
            groups = [{"seeds": [s], "loop_faces": [s]} for s in sorted(seeds)]

        for group in groups:
            group_seeds = group["seeds"]
            loop_faces = group["loop_faces"]

            # Skip if any seed in this group was already consumed
            if any(s in consumed for s in group_seeds):
                logger.info("Group %s already consumed, skipping", group_seeds)
                continue

            # Compute union bbox for the loop faces
            group_bb = None
            for fid in loop_faces:
                bb = bboxes.get(fid)
                if bb is None:
                    continue
                if group_bb is None:
                    group_bb = list(bb)
                else:
                    for i in range(3):
                        group_bb[i] = min(group_bb[i], bb[i])
                    for i in range(3, 6):
                        group_bb[i] = max(group_bb[i], bb[i])
            group_bb_tuple = tuple(group_bb) if group_bb else (0, 0, 0, 0, 0, 0)

            # Pick ALL loop faces (not just seeds) for RemoveSelectedFaces
            current_ids: Set[int] = set(loop_faces)

            if interactive:
                self._highlight_faces(shape_name, sorted(current_ids),
                                      zoom_to_bbox=group_bb_tuple)
                bb_info = ""
                if group_bb_tuple != (0, 0, 0, 0, 0, 0):
                    cx = (group_bb_tuple[0] + group_bb_tuple[3]) / 2
                    cy = (group_bb_tuple[1] + group_bb_tuple[4]) / 2
                    cz = (group_bb_tuple[2] + group_bb_tuple[5]) / 2
                    dx = group_bb_tuple[3] - group_bb_tuple[0]
                    dy = group_bb_tuple[4] - group_bb_tuple[1]
                    dz = group_bb_tuple[5] - group_bb_tuple[2]
                    bb_info = (f"  Location: ({cx:.1f}, {cy:.1f}, {cz:.1f}) "
                               f"size: {dx:.1f}x{dy:.1f}x{dz:.1f} mm")
                print(f"\n  Hole: seeds={sorted(group_seeds)}, "
                      f"all loop faces={sorted(loop_faces)} "
                      f"({len(loop_faces)} face(s))")
                if bb_info:
                    print(bb_info)
                action = self._prompt_action()
                if action == "q":
                    break
                elif action == "n":
                    summary.skipped += 1
                    continue

            success = False
            for attempt in range(MAX_EXPAND + 1):
                sorted_ids = sorted(current_ids)
                logger.info("Group seeds=%s attempt %d: faces=%s",
                            group_seeds, attempt, sorted_ids)

                if interactive:
                    try:
                        self._highlight_faces(shape_name, sorted_ids,
                                              zoom_to_bbox=group_bb_tuple)
                    except CSTConnectionError:
                        pass

                hole_count += 1
                try:
                    ok, msg = self._try_fill_hole(shape_name, sorted_ids, hole_count)
                except CSTConnectionError as exc:
                    ok = False
                    msg = f"CST error: {exc}"
                logger.info("  Result: %s", msg)

                if ok:
                    consumed.update(current_ids)
                    summary.filled += 1
                    print(f"  Hole #{hole_count} removed: faces {sorted_ids}")
                    success = True
                    break
                else:
                    hole_count -= 1
                    if attempt < MAX_EXPAND:
                        old_count = len(current_ids)
                        current_ids = self._expand_faces(
                            current_ids, group_bb_tuple, adjacency, bboxes,
                            consumed
                        )
                        if len(current_ids) == old_count:
                            logger.info("  No new faces to expand, giving up")
                            break
                        logger.info("  Expanded: %d -> %d faces",
                                    old_count, len(current_ids))
                    else:
                        logger.info("  Failed after %d expansion attempts",
                                    MAX_EXPAND)

            if not success:
                summary.failed += 1
                print(f"  Could not remove hole at group {sorted(group_seeds)}")

        self._display_summary(summary)
        return summary

    def run_sequential_workflow(
        self,
        shape_name: str,
        seeds: List[int],
        adjacency: Dict[int, Set[int]],
        bboxes: Dict[int, Tuple],
        face_types: Dict[int, str],
        seed_groups: List[dict] = None,
    ) -> SessionSummary:
        """Interactive sequential workflow."""
        return self.fill_progressive(
            shape_name, seeds, adjacency, bboxes, face_types,
            interactive=True, seed_groups=seed_groups,
        )

    def run_auto_workflow(
        self,
        shape_name: str,
        seeds: List[int],
        adjacency: Dict[int, Set[int]],
        bboxes: Dict[int, Tuple],
        face_types: Dict[int, str],
        seed_groups: List[dict] = None,
    ) -> SessionSummary:
        """Non-interactive workflow — fills all holes without prompting."""
        return self.fill_progressive(
            shape_name, seeds, adjacency, bboxes, face_types,
            interactive=False, seed_groups=seed_groups,
        )

    # ------------------------------------------------------------------
    # Undo
    # ------------------------------------------------------------------

    def undo_fill(self) -> bool:
        """CST 2025 has no programmatic undo. User must Ctrl+Z."""
        print("  NOTE: Automatic undo is not available via CST 2025 COM API.")
        print("  Please press Ctrl+Z in the CST window to undo, then come back.")
        input("  Press Enter after undoing in CST...")
        return True

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _prompt_action() -> str:
        """Prompt for y / n / q."""
        while True:
            choice = input("  Fill this hole? (y/n/q): ").strip().lower()
            if choice in ("y", "n", "q"):
                return choice
            print("  Invalid input. Please enter 'y', 'n', or 'q'.")

    @staticmethod
    def _display_summary(summary: SessionSummary) -> None:
        """Print session summary."""
        print("\n--- Session Summary ---")
        print(f"  Filled:  {summary.filled}")
        print(f"  Skipped: {summary.skipped}")
        print(f"  Failed:  {summary.failed}")
        print("----------------------")
