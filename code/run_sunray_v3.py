"""Run full simplification pipeline on Sunray_MB_v3 with interactive mode.

Detects holes, filters board edges, then fills holes via AddToHistory.
All output is logged to debug_output.txt in real-time (flushed after
every write) so that if the script crashes, the log is still complete.

Run: python -m code.run_sunray_v3
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
    """Real-time logger that writes to both console and file.

    Every message is flushed immediately so the log survives crashes.
    """

    def __init__(self, path: str):
        self._f = open(path, "w", encoding="utf-8")
        self._write(f"=== Run started: {datetime.now().isoformat()} ===\n")

    def log(self, msg: str):
        """Log a message to both console and file."""
        print(msg)
        self._write(msg + "\n")

    def _write(self, text: str):
        self._f.write(text)
        self._f.flush()

    def close(self):
        self._write(f"\n=== Run ended: {datetime.now().isoformat()} ===\n")
        self._f.close()


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

        # Log detection summary
        for data in solid_data:
            shape = data["shape_name"]
            seeds = data["seeds"]
            groups = data.get("seed_groups", [])
            face_types = data["face_types"]

            total_faces = len(face_types)
            cone_count = sum(1 for st in face_types.values() if "cone" in st)
            plane_count = sum(1 for st in face_types.values() if "plane" in st)
            unknown_count = sum(1 for st in face_types.values()
                                if st == "unknown" or st is None)

            log.log(f"\nSolid: {shape}")
            log.log(f"  Total faces parsed: {total_faces}")
            log.log(f"  Cone faces: {cone_count}, Plane faces: {plane_count}, "
                     f"Unknown: {unknown_count}")
            log.log(f"  Seeds after filter: {len(seeds)}")
            log.log(f"  Seeds: {seeds}")
            log.log(f"  Hole groups: {len(groups)}")

            for i, g in enumerate(groups):
                gs = g["seeds"]
                lf = g["loop_faces"]
                # Compute bbox center for location info
                bb = None
                for fid in lf:
                    fb = data["bboxes"].get(fid)
                    if fb is None:
                        continue
                    if bb is None:
                        bb = list(fb)
                    else:
                        for j in range(3):
                            bb[j] = min(bb[j], fb[j])
                        for j in range(3, 6):
                            bb[j] = max(bb[j], fb[j])
                loc = ""
                if bb:
                    cx = (bb[0] + bb[3]) / 2
                    cy = (bb[1] + bb[4]) / 2
                    cz = (bb[2] + bb[5]) / 2
                    dx = bb[3] - bb[0]
                    dy = bb[4] - bb[1]
                    dz = bb[5] - bb[2]
                    loc = (f" @ ({cx:.1f},{cy:.1f},{cz:.1f}) "
                           f"size {dx:.1f}x{dy:.1f}x{dz:.1f}")
                log.log(f"  Group {i+1}: seeds={gs}, "
                         f"loop={lf} ({len(lf)} faces){loc}")

        # --- Fill phase ---
        log.log("\n--- FILL PHASE ---")
        simplifier = Simplifier(conn)

        for data in solid_data:
            shape = data["shape_name"]
            seeds = data["seeds"]
            groups = data.get("seed_groups", [])

            log.log(f"\nProcessing {shape} ({len(seeds)} seeds, "
                     f"{len(groups)} groups)...")

            summary, fill_log = _run_interactive_with_log(
                simplifier, shape, seeds, data["adjacency"],
                data["bboxes"], data["face_types"], groups, log,
            )

            log.log(f"\n  --- Result for {shape} ---")
            log.log(f"  Filled:  {summary.filled}")
            log.log(f"  Skipped: {summary.skipped}")
            log.log(f"  Failed:  {summary.failed}")

    except CSTConnectionError as exc:
        log.log(f"\nCST ERROR: {exc}")
        log.log(traceback.format_exc())
    except Exception as exc:
        log.log(f"\nUNEXPECTED ERROR: {exc}")
        log.log(traceback.format_exc())
    finally:
        conn.close()
        log.close()
        print(f"\nLog saved to {OUT}")


def _run_interactive_with_log(simplifier, shape_name, seeds, adjacency,
                               bboxes, face_types, seed_groups, log):
    """Run interactive fill with detailed per-group logging.

    This replaces the normal simplifier.run_sequential_workflow call
    so we can log every step to the output file.
    """
    from code.models import SessionSummary

    summary = SessionSummary()
    consumed = set()
    hole_count = 0
    fill_log = []

    groups = seed_groups if seed_groups else [
        {"seeds": [s], "loop_faces": [s]} for s in sorted(seeds)
    ]

    for gi, group in enumerate(groups):
      try:
        group_seeds = group["seeds"]
        loop_faces = group["loop_faces"]

        log.log(f"\n  [Group {gi+1}/{len(groups)}] seeds={sorted(group_seeds)}, "
                 f"loop={sorted(loop_faces)} ({len(loop_faces)} faces)")

        # Skip if consumed
        if any(s in consumed for s in group_seeds):
            log.log(f"    -> SKIPPED (already consumed)")
            continue

        # Compute bbox
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

        if group_bb_tuple != (0, 0, 0, 0, 0, 0):
            cx = (group_bb_tuple[0] + group_bb_tuple[3]) / 2
            cy = (group_bb_tuple[1] + group_bb_tuple[4]) / 2
            cz = (group_bb_tuple[2] + group_bb_tuple[5]) / 2
            dx = group_bb_tuple[3] - group_bb_tuple[0]
            dy = group_bb_tuple[4] - group_bb_tuple[1]
            dz = group_bb_tuple[5] - group_bb_tuple[2]
            log.log(f"    Location: ({cx:.1f}, {cy:.1f}, {cz:.1f}) "
                     f"size: {dx:.1f}x{dy:.1f}x{dz:.1f} mm")

        current_ids = set(loop_faces)

        # Highlight and prompt
        try:
            simplifier._highlight_faces(shape_name, sorted(current_ids),
                                        zoom_to_bbox=group_bb_tuple)
        except Exception as exc:
            log.log(f"    Highlight error: {exc}")

        try:
            action = simplifier._prompt_action()
        except (EOFError, KeyboardInterrupt):
            log.log(f"    -> INPUT INTERRUPTED, stopping")
            break

        log.log(f"    User action: {action}")

        if action == "q":
            log.log(f"    -> QUIT requested")
            break
        elif action == "n":
            try:
                simplifier._conn.execute_vba(
                    'Sub Main\n  Pick.ClearAllPicks\nEnd Sub\n')
            except Exception:
                pass
            summary.skipped += 1
            log.log(f"    -> SKIPPED by user")
            continue

        # Try fill with expansion
        success = False
        for attempt in range(6):  # MAX_EXPAND + 1
            sorted_ids = sorted(current_ids)
            log.log(f"    Attempt {attempt}: {len(sorted_ids)} faces = {sorted_ids}")

            if attempt > 0:
                try:
                    simplifier._highlight_faces(shape_name, sorted_ids,
                                                zoom_to_bbox=group_bb_tuple)
                except Exception:
                    pass

            hole_count += 1
            try:
                ok, msg = simplifier._try_fill_hole(
                    shape_name, sorted_ids, hole_count)
            except Exception as exc:
                ok = False
                msg = f"Exception: {exc}"

            log.log(f"    Fill result: ok={ok}, msg={msg}")

            if ok:
                consumed.update(current_ids)
                summary.filled += 1
                log.log(f"    -> FILLED (hole #{hole_count})")
                success = True
                break
            else:
                hole_count -= 1
                if attempt < 5:
                    old_count = len(current_ids)
                    current_ids = simplifier._expand_faces(
                        current_ids, group_bb_tuple, adjacency, bboxes,
                        consumed
                    )
                    if len(current_ids) == old_count:
                        log.log(f"    No new faces to expand, giving up")
                        break
                    log.log(f"    Expanded: {old_count} -> {len(current_ids)} faces")

        if not success:
            # --- Fallback: probe nearby consecutive face IDs ---
            log.log(f"    Normal fill failed. Trying consecutive ID probe...")
            candidates = simplifier.probe_nearby_ids(
                shape_name, group_seeds, consumed, max_range=5
            )
            log.log(f"    Generated {len(candidates)} probe candidates")

            for ci, candidate in enumerate(candidates):
                hole_count += 1
                log.log(f"    Probe {ci}: {candidate}")
                try:
                    simplifier._highlight_faces(shape_name, candidate,
                                                zoom_to_bbox=group_bb_tuple)
                except Exception:
                    pass
                try:
                    ok, msg = simplifier._try_fill_hole(
                        shape_name, candidate, hole_count)
                except Exception as exc:
                    ok = False
                    msg = f"Exception: {exc}"
                log.log(f"    Probe result: ok={ok}, msg={msg}")
                if ok:
                    consumed.update(candidate)
                    summary.filled += 1
                    log.log(f"    -> FILLED via probe (hole #{hole_count})")
                    success = True
                    break
                else:
                    hole_count -= 1

        if not success:
            summary.failed += 1
            log.log(f"    -> FAILED (all attempts exhausted)")

      except Exception as exc:
        log.log(f"    -> ERROR: {exc}")
        log.log(traceback.format_exc())
        summary.failed += 1

    # Print summary
    log.log(f"\n--- Session Summary ---")
    log.log(f"  Filled:  {summary.filled}")
    log.log(f"  Skipped: {summary.skipped}")
    log.log(f"  Failed:  {summary.failed}")
    log.log(f"----------------------")

    return summary, fill_log


if __name__ == "__main__":
    main()
