"""Run full simplification pipeline on Sunray_MB_v3 with interactive mode.

v5 adds WCS-based hole location indicator on top of v4.
When highlighting a hole, the local WCS origin is moved to the hole's
bbox center so the coordinate crosshair marks the hole location in the
CST GUI, making small holes easy to find.

For SAT-detected holes, the bbox comes from the SAT T-field.
For ghost faces, the bbox is estimated from nearby skipped SAT group
faces (if available). If no bbox is available, only pick-highlight is
used (no WCS repositioning).

Changes from v4:
- SAT fill: logs WCS origin coordinates for each group
- Ghost face scan: estimates bbox from nearby skipped SAT group faces
  and passes it to _highlight_faces for WCS positioning
- WCS reset to global in finally block (already in v4)

Run: python -m code.run_sunray_v5
"""

import os
import sys
import logging
import traceback
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from code.cst_connection import CSTConnection, CSTConnectionError
from code.feature_detector import FeatureDetector
from code.simplifier import Simplifier

PROJECT = r"D:\Users\sunze\Desktop\kiro\cst_simplifier\cst_model\Sunray_MB_v3.cst"
OUT = r"D:\Users\sunze\Desktop\kiro\debug_output.txt"


class RunLog:
    """Real-time logger that writes to both console and file."""

    def __init__(self, path: str):
        self._f = open(path, "w", encoding="utf-8")
        self._write(f"=== Run started: {datetime.now().isoformat()} ===\n")

    def log(self, msg: str):
        print(msg)
        self._write(msg + "\n")

    def _write(self, text: str):
        self._f.write(text)
        self._f.flush()

    def close(self):
        self._write(f"\n=== Run ended: {datetime.now().isoformat()} ===\n")
        self._f.close()


def _union_bbox(face_ids, bboxes):
    bb = None
    for fid in face_ids:
        fb = bboxes.get(fid)
        if fb is None:
            continue
        if bb is None:
            bb = list(fb)
        else:
            for i in range(3):
                bb[i] = min(bb[i], fb[i])
            for i in range(3, 6):
                bb[i] = max(bb[i], fb[i])
    return tuple(bb) if bb else (0, 0, 0, 0, 0, 0)


def main():
    log = RunLog(OUT)

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )

    conn = CSTConnection()
    try:
        log.log("Connecting to CST...")
        conn.connect()
        conn.open_project(PROJECT)
        log.log(f"Opened project: {PROJECT}")

        # --- Detection phase ---
        log.log("\n--- DETECTION PHASE ---")
        detector = FeatureDetector(conn)
        solid_data = detector.detect_seeds()

        if not solid_data:
            log.log("No hole candidates found.")
            return

        for data in solid_data:
            shape = data["shape_name"]
            seeds = data["seeds"]
            groups = data.get("seed_groups", [])
            face_types = data["face_types"]

            total_faces = len(face_types)
            cone_count = sum(1 for st in face_types.values() if "cone" in st)
            plane_count = sum(1 for st in face_types.values() if "plane" in st)

            log.log(f"\nSolid: {shape}")
            log.log(f"  Total faces parsed: {total_faces}")
            log.log(f"  Cone: {cone_count}, Plane: {plane_count}")
            log.log(f"  Seeds: {len(seeds)}, Groups: {len(groups)}")

            for i, g in enumerate(groups):
                log.log(f"  Group {i+1}: seeds={g['seeds']}, "
                         f"loop={g['loop_faces']} ({len(g['loop_faces'])} faces)")

        # --- Fill phase (SAT-detected holes) ---
        log.log("\n--- FILL PHASE (SAT-detected) ---")
        simplifier = Simplifier(conn)

        for data in solid_data:
            shape = data["shape_name"]
            seeds = data["seeds"]
            groups = data.get("seed_groups", [])

            log.log(f"\nProcessing {shape} ({len(seeds)} seeds, "
                     f"{len(groups)} groups)...")

            summary, consumed, skipped_group_faces, skipped_group_bboxes = \
                _run_sat_fill(
                    simplifier, shape, seeds, data["adjacency"],
                    data["bboxes"], data["face_types"], groups, log,
                )

            log.log(f"\n  --- SAT result: filled={summary.filled}, "
                     f"skipped={summary.skipped}, failed={summary.failed}")

            # --- Ghost face scan phase ---
            log.log(f"\n--- GHOST FACE SCAN ---")
            ghost_summary = _run_ghost_face_scan(
                simplifier, conn, shape, data["face_types"], data["bboxes"],
                data["adjacency"], consumed, skipped_group_faces,
                skipped_group_bboxes, log,
            )
            log.log(f"\n  --- Ghost result: filled={ghost_summary.filled}, "
                     f"skipped={ghost_summary.skipped}, failed={ghost_summary.failed}")

    except CSTConnectionError as exc:
        log.log(f"\nCST ERROR: {exc}")
        log.log(traceback.format_exc())
    except Exception as exc:
        log.log(f"\nUNEXPECTED ERROR: {exc}")
        log.log(traceback.format_exc())
    finally:
        # Reset WCS back to global origin
        try:
            conn.execute_vba(
                'Sub Main\n'
                '  WCS.ActivateWCS "global"\n'
                'End Sub\n'
            )
        except Exception:
            pass
        conn.close()
        log.close()
        print(f"\nLog saved to {OUT}")


# ======================================================================
# SAT-detected fill
# ======================================================================

def _run_sat_fill(simplifier, shape_name, seeds, adjacency,
                  bboxes, face_types, seed_groups, log):
    """Run interactive fill for SAT-detected holes.

    Returns (summary, consumed_set, skipped_group_faces, skipped_group_bboxes).
    skipped_group_bboxes maps face_id -> bbox for faces in skipped/failed groups,
    so ghost face scan can use them for WCS positioning.
    """
    from code.models import SessionSummary

    summary = SessionSummary()
    consumed = set()
    hole_count = 0

    groups = seed_groups if seed_groups else [
        {"seeds": [s], "loop_faces": [s]} for s in sorted(seeds)
    ]

    # Track bboxes for skipped/failed groups (for ghost face WCS)
    group_bboxes = {}  # group_index -> bbox tuple

    for gi, group in enumerate(groups):
      try:
        group_seeds = group["seeds"]
        loop_faces = group["loop_faces"]

        log.log(f"\n  [Group {gi+1}/{len(groups)}] seeds={sorted(group_seeds)}, "
                 f"loop={sorted(loop_faces)} ({len(loop_faces)} faces)")

        if any(s in consumed for s in group_seeds):
            log.log(f"    -> SKIPPED (already consumed)")
            continue

        group_bb = _union_bbox(loop_faces, bboxes)
        current_ids = set(loop_faces)

        # Store bbox for this group (used later for ghost face WCS)
        group_bboxes[gi] = group_bb

        # Highlight with WCS crosshair at hole center
        try:
            simplifier._highlight_faces(shape_name, sorted(current_ids),
                                        zoom_to_bbox=group_bb)
        except Exception:
            pass

        # Log WCS position
        if group_bb != (0, 0, 0, 0, 0, 0):
            cx = (group_bb[0] + group_bb[3]) / 2
            cy = (group_bb[1] + group_bb[4]) / 2
            cz = (group_bb[2] + group_bb[5]) / 2
            log.log(f"    WCS origin: ({cx:.2f}, {cy:.2f}, {cz:.2f})")

        try:
            action = simplifier._prompt_action()
        except (EOFError, KeyboardInterrupt):
            log.log(f"    -> INPUT INTERRUPTED")
            break

        log.log(f"    User action: {action}")
        if action == "q":
            break
        elif action == "n":
            try:
                simplifier._conn.execute_vba(
                    'Sub Main\n  Pick.ClearAllPicks\nEnd Sub\n')
            except Exception:
                pass
            summary.skipped += 1
            continue

        # Try fill with expansion
        success = False
        for attempt in range(6):
            sorted_ids = sorted(current_ids)
            log.log(f"    Attempt {attempt}: {sorted_ids}")

            if attempt > 0:
                try:
                    simplifier._highlight_faces(shape_name, sorted_ids,
                                                zoom_to_bbox=group_bb)
                except Exception:
                    pass

            hole_count += 1
            try:
                ok, msg = simplifier._try_fill_hole(
                    shape_name, sorted_ids, hole_count)
            except Exception as exc:
                ok = False
                msg = str(exc)

            log.log(f"    Result: ok={ok}, msg={msg}")

            if ok:
                consumed.update(current_ids)
                summary.filled += 1
                log.log(f"    -> FILLED #{hole_count}")
                success = True
                break
            else:
                hole_count -= 1
                if attempt < 5:
                    old = len(current_ids)
                    current_ids = simplifier._expand_faces(
                        current_ids, group_bb, adjacency, bboxes, consumed)
                    if len(current_ids) == old:
                        log.log(f"    No expansion possible")
                        break
                    log.log(f"    Expanded: {old} -> {len(current_ids)}")

        if not success:
            # Probe fallback (consecutive ID probing)
            log.log(f"    Trying probe fallback...")
            candidates = simplifier.probe_nearby_ids(
                shape_name, group_seeds, consumed, max_range=5)
            log.log(f"    {len(candidates)} probe candidates")

            for ci, candidate in enumerate(candidates):
                hole_count += 1
                log.log(f"    Probe {ci}: {candidate}")
                try:
                    ok, msg = simplifier._try_fill_hole(
                        shape_name, candidate, hole_count)
                except Exception as exc:
                    ok = False
                    msg = str(exc)
                log.log(f"    Probe result: ok={ok}")
                if ok:
                    consumed.update(candidate)
                    summary.filled += 1
                    log.log(f"    -> FILLED via probe #{hole_count}")
                    success = True
                    break
                else:
                    hole_count -= 1

        if not success:
            summary.failed += 1
            log.log(f"    -> FAILED")

      except Exception as exc:
        log.log(f"    -> ERROR: {exc}")
        log.log(traceback.format_exc())
        summary.failed += 1

    log.log(f"\n--- SAT Fill Summary ---")
    log.log(f"  Filled={summary.filled}, Skipped={summary.skipped}, "
             f"Failed={summary.failed}")

    # Collect face IDs and bboxes from groups that were skipped or failed
    skipped_group_faces = set()
    skipped_group_bboxes = {}  # face_id -> bbox (from the group's union bbox)
    for gi, group in enumerate(groups):
        group_seeds = group["seeds"]
        if not any(s in consumed for s in group_seeds):
            bb = group_bboxes.get(gi, (0, 0, 0, 0, 0, 0))
            for fid in group["loop_faces"]:
                skipped_group_faces.add(fid)
                if bb != (0, 0, 0, 0, 0, 0):
                    skipped_group_bboxes[fid] = bb
            for fid in group["seeds"]:
                skipped_group_faces.add(fid)
                if bb != (0, 0, 0, 0, 0, 0):
                    skipped_group_bboxes[fid] = bb

    return summary, consumed, skipped_group_faces, skipped_group_bboxes


# ======================================================================
# Ghost face scan — finds holes invisible to the SAT parser
# ======================================================================

def _find_ghost_face_ids(face_types, consumed, bboxes):
    """Find face IDs missing from the SAT export."""
    plane_faces = [pid for pid, st in face_types.items() if "plane" in st]
    board_ref = set()
    if plane_faces:
        def _area(pid):
            bb = bboxes.get(pid)
            if not bb:
                return 0.0
            dims = sorted([bb[3]-bb[0], bb[4]-bb[1], bb[5]-bb[2]], reverse=True)
            return dims[0] * dims[1]
        by_area = sorted(plane_faces, key=_area, reverse=True)
        board_ref = set(by_area[:2])

    known = set(face_types.keys()) | consumed | board_ref
    scan_max = max(face_types.keys()) + 50 if face_types else 200

    return [fid for fid in range(1, scan_max + 1) if fid not in known]


def _get_board_ref_faces(face_types, bboxes):
    """Get the 2 board reference face IDs (largest plane faces)."""
    plane_faces = [pid for pid, st in face_types.items() if "plane" in st]
    if not plane_faces:
        return set()
    def _area(pid):
        bb = bboxes.get(pid)
        if not bb:
            return 0.0
        dims = sorted([bb[3]-bb[0], bb[4]-bb[1], bb[5]-bb[2]], reverse=True)
        return dims[0] * dims[1]
    by_area = sorted(plane_faces, key=_area, reverse=True)
    return set(by_area[:2])


def _estimate_ghost_bbox(fid, skipped_group_bboxes):
    """Estimate a bbox for a ghost face from nearby skipped SAT group faces.

    Ghost faces have no SAT bbox. But if a nearby face ID belongs to a
    skipped SAT group, we can use that group's bbox as an approximation
    (the ghost face is likely part of the same hole).

    Returns bbox tuple or None.
    """
    # Direct match
    if fid in skipped_group_bboxes:
        return skipped_group_bboxes[fid]
    # Check nearby IDs (within ±5)
    for offset in range(1, 6):
        for nearby in [fid + offset, fid - offset]:
            if nearby in skipped_group_bboxes:
                return skipped_group_bboxes[nearby]
    return None


def _try_ghost_fill_windows(simplifier, shape_name, seed_fid,
                             consumed, ghost_consumed,
                             skipped_group_faces, log):
    """Try consecutive ID windows around a ghost face seed.

    Uses _try_fill_hole_silent: RunScript tests silently (no GUI error
    popups), and only persists via AddToHistory if the test passes.

    First tries windows that include skipped SAT group faces (these are
    faces from groups that failed/were skipped in the SAT fill phase).
    Then falls back to plain consecutive windows.

    Returns (success, face_list).
    """
    # Strategy 1: Try windows that include nearby skipped SAT group faces.
    nearby_skipped = sorted(f for f in skipped_group_faces
                            if abs(f - seed_fid) <= 5
                            and f not in consumed
                            and f not in ghost_consumed)
    if nearby_skipped:
        combined_base = set(nearby_skipped) | {seed_fid}
        cmin, cmax = min(combined_base), max(combined_base)
        for fid in range(cmin, cmax + 1):
            if fid not in consumed and fid not in ghost_consumed:
                combined_base.add(fid)

        for extend in range(0, 3):
            combined = set(combined_base)
            for fid in range(cmin - extend, cmax + extend + 1):
                if fid <= 0:
                    continue
                if fid not in consumed and fid not in ghost_consumed:
                    combined.add(fid)
            combined_list = sorted(combined)
            if len(combined_list) < 2 or seed_fid not in combined_list:
                continue
            log.log(f"      Combined (skipped SAT + ghost): {combined_list}")
            try:
                ok, msg = simplifier._try_fill_hole_silent(
                    shape_name, combined_list, 0)
            except Exception as exc:
                ok = False
                msg = str(exc)
            log.log(f"      Fill: ok={ok}, msg={msg}")
            if ok:
                return True, combined_list

    # Strategy 2: Plain consecutive ID windows (sizes 2-5)
    for window in range(2, 6):
        for start in range(seed_fid - window + 1, seed_fid + 1):
            if start <= 0:
                continue
            window_ids = list(range(start, start + window))
            window_ids = [f for f in window_ids
                          if f not in consumed and f not in ghost_consumed]
            if len(window_ids) < 2 or seed_fid not in window_ids:
                continue
            log.log(f"      Window: {window_ids}")
            try:
                ok, msg = simplifier._try_fill_hole_silent(
                    shape_name, window_ids, 0)
            except Exception as exc:
                ok = False
                msg = str(exc)
            log.log(f"      Fill: ok={ok}, msg={msg}")
            if ok:
                return True, window_ids

    return False, []


def _run_ghost_face_scan(simplifier, conn, shape_name, face_types, bboxes,
                          adjacency, consumed, skipped_group_faces,
                          skipped_group_bboxes, log):
    """Scan ghost faces one by one, asking user to confirm each.

    v5: uses skipped_group_bboxes to estimate ghost face location and
    position the WCS crosshair, making ghost faces easier to find.
    """
    from code.models import SessionSummary

    summary = SessionSummary()
    ghost_ids = _find_ghost_face_ids(face_types, consumed, bboxes)

    log.log(f"  Ghost face IDs found: {len(ghost_ids)}")
    if not ghost_ids:
        log.log("  No ghost faces to scan.")
        return summary

    log.log(f"  Ghost IDs: {ghost_ids}")

    board_ref_faces = _get_board_ref_faces(face_types, bboxes)
    log.log(f"  Board ref faces: {sorted(board_ref_faces)}")

    ghost_consumed = set()
    hole_count = 0

    for fid in ghost_ids:
        if fid in ghost_consumed:
            continue

        log.log(f"\n  [Ghost face {fid}]")

        # Check if this face still exists by trying to pick it
        try:
            import tempfile
            _out_path = os.path.join(tempfile.gettempdir(), "cst_fill.txt")
            _out_vba = _out_path.replace("\\", "\\\\")
            check_code = (
                'Sub Main\n'
                f'  Open "{_out_vba}" For Output As #1\n'
                '  On Error Resume Next\n'
                '  Pick.ClearAllPicks\n'
                f'  Pick.PickFaceFromId "{shape_name}", "{fid}"\n'
                '  If Err.Number <> 0 Then\n'
                '    Print #1, "GONE"\n'
                '    Err.Clear\n'
                '  Else\n'
                '    Print #1, "EXISTS"\n'
                '  End If\n'
                '  Pick.ClearAllPicks\n'
                '  Close #1\n'
                'End Sub\n'
            )
            check_result = conn.execute_vba(check_code, output_file=_out_path)
            if "GONE" in (check_result or ""):
                log.log(f"    -> SKIPPED (face no longer exists)")
                ghost_consumed.add(fid)
                continue
        except Exception:
            pass

        # Estimate bbox from nearby skipped SAT group faces (for WCS)
        ghost_bb = _estimate_ghost_bbox(fid, skipped_group_bboxes)

        # Highlight this face with WCS crosshair if bbox available
        try:
            simplifier._highlight_faces(shape_name, [fid],
                                        zoom_to_bbox=ghost_bb)
        except Exception as exc:
            log.log(f"    Highlight error: {exc}")

        # Log WCS position if available
        if ghost_bb and ghost_bb != (0, 0, 0, 0, 0, 0):
            cx = (ghost_bb[0] + ghost_bb[3]) / 2
            cy = (ghost_bb[1] + ghost_bb[4]) / 2
            cz = (ghost_bb[2] + ghost_bb[5]) / 2
            log.log(f"    WCS origin (estimated): ({cx:.2f}, {cy:.2f}, {cz:.2f})")

        # Ask user
        print(f"\n  Ghost face {fid} (not in SAT export)")
        try:
            while True:
                choice = input("  Is this face part of a hole? (y/n/q): ").strip().lower()
                if choice in ("y", "n", "q"):
                    break
                print("  Please enter y, n, or q.")
        except (EOFError, KeyboardInterrupt):
            log.log(f"    -> INPUT INTERRUPTED")
            break

        log.log(f"    User: {choice}")

        if choice == "q":
            log.log(f"    -> QUIT")
            break
        elif choice == "n":
            try:
                simplifier._conn.execute_vba(
                    'Sub Main\n  Pick.ClearAllPicks\nEnd Sub\n')
            except Exception:
                pass
            summary.skipped += 1
            continue

        # User said yes — clear picks and try consecutive windows
        try:
            simplifier._conn.execute_vba(
                'Sub Main\n  Pick.ClearAllPicks\nEnd Sub\n')
        except Exception:
            pass

        log.log(f"    Trying consecutive ID windows...")
        success, hole_faces = _try_ghost_fill_windows(
            simplifier, shape_name, fid, consumed, ghost_consumed,
            skipped_group_faces, log)

        if success and hole_faces:
            ghost_consumed.update(hole_faces)
            hole_count += 1
            summary.filled += 1
            log.log(f"    -> FILLED ghost hole #{hole_count}: {hole_faces}")
        else:
            summary.failed += 1
            log.log(f"    -> FAILED (no window worked for face {fid})")

    log.log(f"\n--- Ghost Scan Summary ---")
    log.log(f"  Filled={summary.filled}, Skipped={summary.skipped}, "
             f"Failed={summary.failed}")

    return summary


if __name__ == "__main__":
    main()
