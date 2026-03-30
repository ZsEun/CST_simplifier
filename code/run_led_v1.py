"""Shield can simplifier — detect side walls, find dimples, fill them.

v1: Full pipeline for shield can components (cover/frame).
1. Connect to CST, open model
2. Export SAT, parse faces/adjacency/bboxes
3. Find top face → discover side walls via adjacency
4. For each wall, find dimple faces using local UV projection
5. Highlight dimple group, ask user y/n/q, fill via AddToHistory
6. "Are all holes filled?" after each wall
7. WCS crosshair at dimple center, reset to global on exit

Run: python -m code.run_led_v1
  or: python -m code.run_led_v1 "path/to/model.cst"
"""

import os
import sys
import logging
import traceback
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from code.cst_connection import CSTConnection, CSTConnectionError
from code.feature_detector import FeatureDetector, SATParser
from code.wall_detector import WallDetector
from code.simplifier import Simplifier


class RunLog:
    def __init__(self, path):
        self._f = open(path, "w", encoding="utf-8")
        self._write(f"=== Run started: {datetime.now().isoformat()} ===\n")

    def log(self, msg):
        print(msg)
        self._write(msg + "\n")

    def _write(self, text):
        self._f.write(text)
        self._f.flush()

    def close(self):
        self._write(f"\n=== Run ended: {datetime.now().isoformat()} ===\n")
        self._f.close()


def _get_project_path():
    if len(sys.argv) > 1:
        path = sys.argv[1].strip().strip('"').strip("'")
    else:
        print("=" * 60)
        print("  Shield Can Simplifier")
        print("=" * 60)
        path = input("\n  Enter path to .cst model: ").strip().strip('"').strip("'")
    if not path:
        print("  Error: no path provided.")
        sys.exit(1)
    path = os.path.abspath(path)
    if not path.lower().endswith(".cst"):
        print(f"  Error: expected .cst file, got: {path}")
        sys.exit(1)
    if not os.path.exists(path):
        print(f"  Error: file not found: {path}")
        sys.exit(1)
    return path


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


def _ask_all_done():
    """Ask user if all holes are filled. Returns True if yes."""
    try:
        while True:
            choice = input("  Are all holes filled? (y/n): ").strip().lower()
            if choice in ("y", "n"):
                return choice == "y"
            print("  Please enter y or n.")
    except (EOFError, KeyboardInterrupt):
        return True


def main():
    project_path = _get_project_path()
    log_dir = os.path.dirname(project_path)
    model_name = os.path.splitext(os.path.basename(project_path))[0]
    log_path = os.path.join(log_dir, f"{model_name}_simplifier_log.txt")

    log = RunLog(log_path)
    logging.basicConfig(level=logging.INFO,
                        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s")

    conn = CSTConnection()
    try:
        log.log("Connecting to CST...")
        conn.connect()
        conn.open_project(project_path)
        log.log(f"Opened: {project_path}")

        # --- Detection phase ---
        log.log("\n--- DETECTION PHASE ---")
        detector = FeatureDetector(conn)
        solid_data = detector.detect_seeds()

        if not solid_data:
            log.log("No solids found.")
            return

        # Show available components
        log.log(f"\nFound {len(solid_data)} component(s):")
        for i, data in enumerate(solid_data):
            log.log(f"  {i+1}. {data['shape_name']} "
                     f"({len(data['face_types'])} faces)")

        # Let user pick which component
        if len(solid_data) == 1:
            chosen = solid_data[0]
        else:
            try:
                idx = int(input(f"\n  Which component? (1-{len(solid_data)}): ")) - 1
                chosen = solid_data[idx]
            except (ValueError, IndexError, EOFError):
                log.log("  Invalid choice, using first component.")
                chosen = solid_data[0]

        shape_name = chosen["shape_name"]
        log.log(f"\nProcessing: {shape_name}")

        # Re-parse SAT for full face data with geometry
        parts = shape_name.split(":")
        sat_path = detector._export_sat(parts[0], parts[1])
        parser = SATParser(sat_path)
        face_data = parser.parse()
        adjacency = parser.build_adjacency()
        bboxes = parser.get_bounding_boxes()
        log.log(f"  Parsed {len(face_data)} faces, "
                 f"adjacency for {len(adjacency)} faces")

        # --- Wall detection ---
        log.log("\n--- WALL DETECTION ---")
        wall_det = WallDetector()
        top_pid, top_normal = wall_det.find_top_face(face_data, bboxes)
        log.log(f"  Top face: pid={top_pid}, normal={top_normal}")

        walls = wall_det.discover_side_walls(
            top_pid, top_normal, face_data, adjacency, bboxes)
        log.log(f"  Side walls: {len(walls)}")
        for i, w in enumerate(walls):
            log.log(f"    Wall {i+1}: pid={w.face_pid}, "
                     f"normal=({w.normal[0]:.3f},{w.normal[1]:.3f},{w.normal[2]:.3f})")

        # --- Per-wall dimple detection and fill ---
        log.log("\n--- DIMPLE DETECTION & FILL ---")
        simplifier = Simplifier(conn)
        total_filled = 0
        total_skipped = 0
        total_failed = 0

        for wi, wall in enumerate(walls):
            dimple_faces = wall_det.find_dimple_faces(
                wall, face_data, adjacency, bboxes, top_pid, walls)

            if not dimple_faces:
                continue

            log.log(f"\n  Wall {wi+1}/{len(walls)}: pid={wall.face_pid}, "
                     f"{len(dimple_faces)} dimple faces: {dimple_faces}")

            # Set WCS local coordinate at dimple group center
            bb = _union_bbox(dimple_faces, bboxes)
            wcx = (bb[0] + bb[3]) / 2
            wcy = (bb[1] + bb[4]) / 2
            wcz = (bb[2] + bb[5]) / 2
            try:
                conn.execute_vba(
                    'Sub Main\n'
                    f'  WCS.SetOrigin {wcx}, {wcy}, {wcz}\n'
                    '  WCS.ActivateWCS "local"\n'
                    'End Sub\n')
            except Exception:
                pass
            log.log(f"    WCS at ({wcx:.2f}, {wcy:.2f}, {wcz:.2f})")

            # Highlight dimple faces and ask user
            try:
                simplifier._highlight_faces(shape_name, dimple_faces,
                                            zoom_to_bbox=bb)
            except Exception:
                pass

            try:
                action = simplifier._prompt_action()
            except (EOFError, KeyboardInterrupt):
                log.log(f"    -> INPUT INTERRUPTED")
                break

            log.log(f"    User: {action}")
            if action == "q":
                break
            elif action == "n":
                try:
                    conn.execute_vba('Sub Main\n  Pick.ClearAllPicks\nEnd Sub\n')
                except Exception:
                    pass
                total_skipped += 1
                continue

            # Try fill
            log.log(f"    Filling {len(dimple_faces)} faces...")
            try:
                ok, msg = simplifier._try_fill_hole(
                    shape_name, dimple_faces, wi + 1)
            except Exception as exc:
                ok = False
                msg = str(exc)

            log.log(f"    Result: ok={ok}, msg={msg}")

            if ok:
                total_filled += 1
                log.log(f"    -> FILLED")
            else:
                total_failed += 1
                log.log(f"    -> FAILED")

        # --- Summary ---
        log.log(f"\n--- SUMMARY ---")
        log.log(f"  Filled: {total_filled}")
        log.log(f"  Skipped: {total_skipped}")
        log.log(f"  Failed: {total_failed}")

    except CSTConnectionError as exc:
        log.log(f"\nCST ERROR: {exc}")
        log.log(traceback.format_exc())
    except Exception as exc:
        log.log(f"\nERROR: {exc}")
        log.log(traceback.format_exc())
    finally:
        try:
            conn.execute_vba(
                'Sub Main\n'
                '  WCS.ActivateWCS "global"\n'
                '  Pick.ClearAllPicks\n'
                'End Sub\n')
        except Exception:
            pass
        conn.close()
        log.close()
        print(f"\nLog saved to {log_path}")


if __name__ == "__main__":
    main()
