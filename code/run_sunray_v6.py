"""Generic CST CAD simplifier — works with any .cst model.

v6: prompts user for the .cst model path at startup instead of
hardcoding Sunray_MB_v3. All features from v5 are included:
- SAT-based hole detection + fill with expansion + probe fallback
- Ghost face scan with combined strategy
- WCS crosshair at hole center for easy location
- Real-time logging to console + file

The log file is saved next to the .cst file as simplifier_log.txt.

Run: python -m code.run_sunray_v6
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


def _get_project_path() -> str:
    """Ask user for the .cst model path. Supports command-line arg too."""
    # Check command-line argument first
    if len(sys.argv) > 1:
        path = sys.argv[1].strip().strip('"').strip("'")
    else:
        print("=" * 60)
        print("  CST CAD Simplifier")
        print("=" * 60)
        path = input("\n  Enter path to .cst model: ").strip().strip('"').strip("'")

    if not path:
        print("  Error: no path provided.")
        sys.exit(1)

    # Normalize path
    path = os.path.abspath(path)

    if not path.lower().endswith(".cst"):
        print(f"  Error: expected a .cst file, got: {path}")
        sys.exit(1)

    if not os.path.exists(path):
        print(f"  Error: file not found: {path}")
        sys.exit(1)

    return path


def main():
    project_path = _get_project_path()

    # Place log file next to the .cst file
    log_dir = os.path.dirname(project_path)
    model_name = os.path.splitext(os.path.basename(project_path))[0]
    log_path = os.path.join(log_dir, f"{model_name}_simplifier_log.txt")

    log = RunLog(log_path)

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )

    conn = CSTConnection()
    try:
        log.log("Connecting to CST...")
        conn.connect()
        conn.open_project(project_path)
        log.log(f"Opened project: {project_path}")

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

            # Ask user if all holes are already filled
            print(f"\n  SAT-detected fill complete. "
                  f"Filled={summary.filled}, Failed={summary.failed}")
            try:
                while True:
                    done = input("  Are all holes filled? (y/n): ").strip().lower()
                    if done in ("y", "n"):
                        break
                    print("  Please enter y or n.")
            except (EOFError, KeyboardInterrupt):
                done = "y"

            log.log(f"  All holes filled? {done}")
            if done == "y":
                log.log("  Skipping ghost face scan (user confirmed all done).")
                continue

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
        print(f"\nLog saved to {log_path}")


# ======================================================================
# SAT-detected fill
# ======================================================================

def _run_sat_fill(simplifier, shape_name, seeds, adjacency,
                  bboxes, face_types, seed_groups, log):
    """Run interactive fill for SAT-detected holes.

    Returns (summary, consumed_set, skipped_group_faces, skipped_group_bboxes).
    """
    from code.models import SessionSummary

    summary = SessionSummary()
    consumed = set()
    hole_count = 0

    groups = seed_groups if seed_groups else [
        {"seeds": [s], "loop_faces": [s]} for s in sorted(seeds)
    ]

    group_bboxes = {}

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
        group_bboxes[gi] = group_bb

        try:
            simplifier._highlight_faces(shape_name, sorted(current_ids),
                                        zoom_to_bbox=group_bb)
        except Exception:
            pass

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

    skipped_group_faces = set()
    skipped_group_bboxes = {}
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
# Ghost face scan
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
    """Estimate bbox for a ghost face from nearby skipped SAT group faces."""
    if fid in skipped_group_bboxes:
        return skipped_group_bboxes[fid]
    for offset in range(1, 6):
        for nearby in [fid + offset, fid - offset]:
            if nearby in skipped_group_bboxes:
                return skipped_group_bboxes[nearby]
    return None


def _try_ghost_fill_windows(simplifier, shape_name, seed_fid,
                             consumed, ghost_consumed,
                             skipped_group_faces, log):
    """Try consecutive ID windows around a ghost face seed."""
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
    """Scan ghost faces one by one, asking user to confirm each."""
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

        ghost_bb = _estimate_ghost_bbox(fid, skipped_group_bboxes)

        try:
            simplifier._highlight_faces(shape_name, [fid],
                                        zoom_to_bbox=ghost_bb)
        except Exception as exc:
            log.log(f"    Highlight error: {exc}")

        if ghost_bb and ghost_bb != (0, 0, 0, 0, 0, 0):
            cx = (ghost_bb[0] + ghost_bb[3]) / 2
            cy = (ghost_bb[1] + ghost_bb[4]) / 2
            cz = (ghost_bb[2] + ghost_bb[5]) / 2
            log.log(f"    WCS origin (estimated): ({cx:.2f}, {cy:.2f}, {cz:.2f})")

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

            # Ask if all holes are now filled
            print(f"\n  Ghost hole filled: {hole_faces}")
            try:
                while True:
                    done = input("  Are all holes filled? (y/n): ").strip().lower()
                    if done in ("y", "n"):
                        break
                    print("  Please enter y or n.")
            except (EOFError, KeyboardInterrupt):
                done = "y"

            log.log(f"    All holes filled? {done}")
            if done == "y":
                log.log("    Stopping ghost scan (user confirmed all done).")
                break
        else:
            summary.failed += 1
            log.log(f"    -> FAILED (no window worked for face {fid})")

    log.log(f"\n--- Ghost Scan Summary ---")
    log.log(f"  Filled={summary.filled}, Skipped={summary.skipped}, "
             f"Failed={summary.failed}")

    return summary


if __name__ == "__main__":
    main()
