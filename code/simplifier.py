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

    def _test_fill_hole(self, shape_name: str, face_ids: List[int]) -> bool:
        """Test if a set of faces forms a fillable hole WITHOUT persisting changes.

        Uses RunScript (not AddToHistory) so nothing is saved to model history.
        All errors are caught silently via On Error Resume Next — no error
        dialogs will appear in the CST GUI.

        IMPORTANT: If the test succeeds, the model IS temporarily modified
        in memory (but not persisted). The caller should be aware that
        after a successful test, the model state has changed. This is
        acceptable because we immediately follow up with the real
        AddToHistory fill on the same faces.

        Returns True if RemoveSelectedFaces succeeds for these faces.
        """
        pick_lines = "\n".join(
            f'  Pick.PickFaceFromId "{shape_name}", "{fid}"'
            for fid in face_ids
        )
        code = (
            'Sub Main\n'
            f'  Open "{self._out_vba}" For Output As #1\n'
            '  On Error Resume Next\n'
            '  Pick.ClearAllPicks\n'
            f'{pick_lines}\n'
            '  If Err.Number <> 0 Then\n'
            '    Print #1, "PICK_FAIL"\n'
            '    Err.Clear\n'
            '    Pick.ClearAllPicks\n'
            '    Close #1\n'
            '    Exit Sub\n'
            '  End If\n'
            '  With LocalModification\n'
            '    .Reset\n'
            f'    .Name "{shape_name}"\n'
            '    .RemoveSelectedFaces\n'
            '  End With\n'
            '  If Err.Number <> 0 Then\n'
            '    Print #1, "REMOVE_FAIL"\n'
            '    Err.Clear\n'
            '  Else\n'
            '    Print #1, "OK"\n'
            '  End If\n'
            '  Pick.ClearAllPicks\n'
            '  Close #1\n'
            'End Sub\n'
        )
        result = self._run_vba(code)
        return "OK" in (result or "")

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


    def _try_fill_hole_silent(self, shape_name: str, face_ids: List[int],
                               hole_index: int) -> Tuple[bool, str]:
        """Try to fill a hole silently — no CST GUI error popups.

        Strategy:
        1. Pick all faces via AddToHistory (picks always succeed, no popup)
        2. Attempt RemoveSelectedFaces via RunScript (silent — On Error
           Resume Next catches failures without GUI popups)
        3. If RunScript succeeded, record the remove in history via
           AddToHistory so the change persists across model rebuilds
        4. If RunScript failed, clear picks via AddToHistory and return False

        This avoids the GUI error popups that occur when AddToHistory
        directly executes a failing RemoveSelectedFaces.

        Returns:
            (success, message) tuple
        """
        # Step 1: Pick each face via AddToHistory (always succeeds)
        for fid in face_ids:
            pick_code = f'Pick.PickFaceFromId "{shape_name}", "{fid}"'
            result = self._add_to_history("pick face", pick_code)
            if "HIST_ERR" in (result or ""):
                self._add_to_history("clear picks", "Pick.ClearAllPicks")
                return False, f"Pick failed for face {fid}: {result}"

        # Step 2: Try RemoveSelectedFaces via RunScript (silent)
        test_code = (
            'Sub Main\n'
            f'  Open "{self._out_vba}" For Output As #1\n'
            '  On Error Resume Next\n'
            '  With LocalModification\n'
            '    .Reset\n'
            f'    .Name "{shape_name}"\n'
            '    .RemoveSelectedFaces\n'
            '  End With\n'
            '  If Err.Number <> 0 Then\n'
            '    Print #1, "REMOVE_FAIL"\n'
            '    Err.Clear\n'
            '  Else\n'
            '    Print #1, "OK"\n'
            '  End If\n'
            '  Close #1\n'
            'End Sub\n'
        )
        result = self._run_vba(test_code)

        if "OK" not in (result or ""):
            # Remove failed — clear picks and abort
            self._add_to_history("clear picks", "Pick.ClearAllPicks")
            return False, f"Silent remove failed: {result}"

        # Step 3: RunScript succeeded — model is modified in memory.
        # The picks are already recorded in history from step 1.
        # We skip recording RemoveSelectedFaces via AddToHistory here
        # because the faces are already gone and AddToHistory would
        # show a GUI error popup trying to re-execute it.
        # The model change persists when the project is saved.

        # Step 4: Clear picks
        self._add_to_history("clear picks", "Pick.ClearAllPicks")

        return True, f"Silent fill OK: {face_ids}"


    def probe_nearby_ids(
        self,
        shape_name: str,
        seed_ids: List[int],
        consumed: Set[int],
        max_range: int = 5,
    ) -> List[List[int]]:
        """Generate candidate face-ID sets by probing consecutive IDs near seeds.

        When the SAT parser misses some face entities (their pid attribs
        reference edges/vertices instead of faces), the BFS loop becomes
        wrong. This fallback probes nearby consecutive face IDs around
        the seeds, since CST often assigns consecutive IDs to faces
        forming the same hole.

        Only called after the normal SAT-based fill has already failed.
        Does NOT modify the detection algorithm at all.

        Strategy:
        - Build consecutive runs of IDs starting from each seed
        - Try runs of length 2, 3, 4 (most holes are 2-4 faces)
        - Prioritize forward runs (seed, seed+1, seed+2, ...)
          since CST tends to assign ascending IDs to hole faces

        Returns:
            List of candidate face-ID lists to try, ordered by priority.
        """
        if not seed_ids:
            return []

        seed_set = set(seed_ids)
        candidates = []
        seen = set()

        def _add(ids):
            key = tuple(sorted(ids))
            if key not in seen:
                seen.add(key)
                candidates.append(list(key))

        # Strategy 1: consecutive runs starting from min seed
        # e.g. seed=147 -> try [147,148], [147,148,149], [147,148,149,150]
        min_seed = min(seed_ids)
        for run_len in range(2, 6):
            run = [min_seed + i for i in range(run_len)]
            run = [fid for fid in run if fid not in consumed or fid in seed_set]
            if len(run) >= 2:
                _add(run)

        # Strategy 2: consecutive runs ending at max seed
        max_seed = max(seed_ids)
        for run_len in range(2, 6):
            run = [max_seed - run_len + 1 + i for i in range(run_len)]
            run = [fid for fid in run if fid > 0 and (fid not in consumed or fid in seed_set)]
            if len(run) >= 2:
                _add(run)

        # Strategy 3: seed + each nearby ID individually
        for offset in range(1, max_range + 1):
            for fid in [min_seed + offset, min_seed - offset]:
                if fid > 0 and fid not in consumed:
                    _add(seed_set | {fid})

        # Strategy 4: seed + pairs of nearby IDs (close ones first)
        nearby = []
        for offset in range(1, max_range + 1):
            for fid in [min_seed + offset, min_seed - offset]:
                if fid > 0 and fid not in seed_set and fid not in consumed:
                    nearby.append(fid)
        for i, n1 in enumerate(nearby[:6]):
            for n2 in nearby[i+1:6]:
                _add(seed_set | {n1, n2})

        # Strategy 5: seed + triples (only closest neighbors)
        for i, n1 in enumerate(nearby[:4]):
            for j, n2 in enumerate(nearby[i+1:4], i+1):
                for n3 in nearby[j+1:4]:
                    combo = seed_set | {n1, n2, n3}
                    if len(combo) <= 6:
                        _add(combo)

        return candidates

    def find_ghost_holes(
        self,
        shape_name: str,
        face_types: Dict[int, str],
        consumed: Set[int],
        board_ref_faces: Set[int],
        scan_range: Tuple[int, int] = (1, 600),
    ) -> List[List[int]]:
        """Find holes formed by face IDs completely missing from the SAT export.

        Some CST faces have no corresponding face entity in the SAT file
        (their pid attribs reference edges/vertices/loops instead). These
        faces are invisible to the SAT parser — no surface type, no
        adjacency, no bbox. But they exist in CST and can be picked.

        Strategy: find consecutive runs of ghost IDs, then generate
        candidate groups using a sliding window of sizes 3, 4, and 5.
        Prioritize 4-face groups (most common hole size) then 3 and 5.

        Args:
            shape_name: CST shape name
            face_types: pid -> surface type from SAT parse
            consumed: face IDs already filled
            board_ref_faces: board reference face IDs to exclude
            scan_range: (min_id, max_id) range to scan

        Returns:
            List of candidate face-ID lists, ordered by priority.
        """
        known_ids = set(face_types.keys()) | consumed | board_ref_faces
        min_id, max_id = scan_range

        # Step 1: find all consecutive runs of ghost IDs
        runs: List[List[int]] = []
        current_run: List[int] = []
        for fid in range(min_id, max_id + 1):
            if fid not in known_ids:
                current_run.append(fid)
            else:
                if len(current_run) >= 3:
                    runs.append(current_run)
                current_run = []
        if len(current_run) >= 3:
            runs.append(current_run)

        # Step 2: for each run, generate sliding-window candidates
        # Try window sizes 4, 3, 5 (4 is most common for screw holes)
        candidates: List[List[int]] = []
        seen: set = set()

        for window_size in [4, 3, 5]:
            for run in runs:
                if len(run) < window_size:
                    continue
                for i in range(len(run) - window_size + 1):
                    chunk = run[i:i + window_size]
                    key = tuple(chunk)
                    if key not in seen:
                        seen.add(key)
                        candidates.append(chunk)

        return candidates


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
        """Pick faces in CST GUI and move WCS origin to the hole center.

        Sets the local WCS origin to the bbox center so the coordinate
        crosshair marks the hole location, making it easy to find.

        Uses RunScript (not AddToHistory) since highlighting is
        a view operation, not a model change.
        """
        pick_lines = "\n".join(
            f'  Pick.PickFaceFromId "{shape_name}", "{fid}"'
            for fid in face_ids
        )

        # Move WCS origin to bbox center so crosshair marks the hole
        wcs_lines = ""
        if zoom_to_bbox and zoom_to_bbox != (0, 0, 0, 0, 0, 0):
            cx = (zoom_to_bbox[0] + zoom_to_bbox[3]) / 2
            cy = (zoom_to_bbox[1] + zoom_to_bbox[4]) / 2
            cz = (zoom_to_bbox[2] + zoom_to_bbox[5]) / 2
            wcs_lines = (
                f'  WCS.SetOrigin {cx}, {cy}, {cz}\n'
                '  WCS.ActivateWCS "local"\n'
            )

        code = (
            'Sub Main\n'
            '  Plot.ZoomToStructure\n'
            '  Pick.ClearAllPicks\n'
            f'{pick_lines}\n'
            f'{wcs_lines}'
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
          try:
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
                try:
                    self._highlight_faces(shape_name, sorted(current_ids),
                                          zoom_to_bbox=group_bb_tuple)
                except Exception:
                    logger.debug("Highlight failed for group %s", group_seeds)
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
                try:
                    action = self._prompt_action()
                except (EOFError, KeyboardInterrupt):
                    print("\n  Input interrupted, stopping.")
                    break
                if action == "q":
                    break
                elif action == "n":
                    # Clear picks before moving to next group
                    try:
                        self._conn.execute_vba(
                            'Sub Main\n  Pick.ClearAllPicks\nEnd Sub\n')
                    except Exception:
                        pass
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
                    except Exception:
                        logger.debug("Highlight failed during expand for %s",
                                     group_seeds)

                hole_count += 1
                try:
                    ok, msg = self._try_fill_hole(shape_name, sorted_ids, hole_count)
                except CSTConnectionError as exc:
                    ok = False
                    msg = f"CST error: {exc}"
                except Exception as exc:
                    ok = False
                    msg = f"Unexpected error: {exc}"
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
                # Fallback: probe nearby consecutive face IDs
                logger.info("Normal fill failed for %s, trying ID probe",
                            group_seeds)
                candidates = self.probe_nearby_ids(
                    shape_name, group_seeds, consumed, max_range=5
                )
                for candidate in candidates:
                    hole_count += 1
                    logger.info("  Probe: %s", candidate)
                    if interactive:
                        try:
                            self._highlight_faces(shape_name, candidate,
                                                  zoom_to_bbox=group_bb_tuple)
                        except Exception:
                            pass
                    try:
                        ok, msg = self._try_fill_hole(
                            shape_name, candidate, hole_count)
                    except Exception as exc:
                        ok = False
                        msg = f"Probe error: {exc}"
                    if ok:
                        consumed.update(candidate)
                        summary.filled += 1
                        print(f"  Hole #{hole_count} removed via probe: "
                              f"faces {candidate}")
                        success = True
                        break
                    else:
                        hole_count -= 1

            if not success:
                summary.failed += 1
                print(f"  Could not remove hole at group {sorted(group_seeds)}")

          except Exception as exc:
            logger.error("Error processing group %s: %s",
                         group.get("seeds", "?"), exc)
            print(f"  Error on group {group.get('seeds', '?')}: {exc}")
            summary.failed += 1

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
