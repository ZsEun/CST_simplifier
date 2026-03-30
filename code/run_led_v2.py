"""Shield can simplifier v2 — full workflow.

For each side wall:
1. Set WCS aligned with wall (origin at center, normal = W axis)
2. Highlight wall face
3. Find dimple faces using UV projection (golden algorithm)
4. If no dimples found → skip to next wall automatically
5. If dimples found → highlight them, ask user y/n/q to fill

Run: python -m code.run_led_v2
  or: python -m code.run_led_v2 "path/to/model.cst"
"""

import os
import sys
import logging
import tempfile
import traceback
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from code.cst_connection import CSTConnection, CSTConnectionError
from code.feature_detector import FeatureDetector, SATParser
from code.wall_detector import WallDetector
from code.simplifier import Simplifier

_out_path = os.path.join(tempfile.gettempdir(), "cst_fill.txt")
_out_vba = _out_path.replace("\\", "\\\\")


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
    return r"D:\Users\sunze\Desktop\kiro\cst_simplifier\cst_model\Sunray_LED_v3.cst"


def _set_wcs_on_wall(conn, wall):
    """Set WCS origin at wall center, aligned with wall normal."""
    bb = wall.bbox
    cx = (bb[0]+bb[3])/2; cy = (bb[1]+bb[4])/2; cz = (bb[2]+bb[5])/2
    wn = wall.normal
    dx = bb[3]-bb[0]; dy = bb[4]-bb[1]; dz = bb[5]-bb[2]
    candidates = [(dx, (1,0,0)), (dy, (0,1,0)), (dz, (0,0,1))]
    candidates.sort(key=lambda c: c[0], reverse=True)
    u_axis = None
    for extent, axis in candidates:
        dot = abs(axis[0]*wn[0]+axis[1]*wn[1]+axis[2]*wn[2])
        if dot < 0.7:
            u_axis = axis
            break
    if u_axis is None:
        u_axis = (1,0,0) if abs(wn[0]) < 0.9 else (0,1,0)

    code = (
        'Sub Main\n'
        f'  Open "{_out_vba}" For Output As #1\n'
        '  On Error Resume Next\n'
        f'  WCS.SetOrigin {cx}, {cy}, {cz}\n'
        f'  WCS.SetNormal {wn[0]}, {wn[1]}, {wn[2]}\n'
        f'  WCS.SetUVector {u_axis[0]}, {u_axis[1]}, {u_axis[2]}\n'
        '  WCS.ActivateWCS "local"\n'
        '  If Err.Number <> 0 Then\n'
        '    Print #1, "WCS_ERR: " & Err.Description\n'
        '    Err.Clear\n'
        '  Else\n'
        '    Print #1, "WCS_OK"\n'
        '  End If\n'
        '  Close #1\n'
        'End Sub\n'
    )
    conn.execute_vba(code, output_file=_out_path)


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

        # Detection
        log.log("\n--- DETECTION ---")
        detector = FeatureDetector(conn)
        solid_data = detector.detect_seeds()
        if not solid_data:
            log.log("No solids found.")
            return

        # Component selection
        log.log(f"\nFound {len(solid_data)} component(s):")
        for i, data in enumerate(solid_data):
            log.log(f"  {i+1}. {data['shape_name']}")
        if len(solid_data) == 1:
            chosen = solid_data[0]
        else:
            try:
                idx = int(input(f"  Which? (1-{len(solid_data)}): ")) - 1
                chosen = solid_data[idx]
            except (ValueError, IndexError, EOFError):
                chosen = solid_data[0]

        shape_name = chosen["shape_name"]
        log.log(f"\nProcessing: {shape_name}")

        # Parse SAT
        parts = shape_name.split(":")
        sat_path = detector._export_sat(parts[0], parts[1])
        parser = SATParser(sat_path)
        face_data = parser.parse()
        adjacency = parser.build_adjacency()
        bboxes = parser.get_bounding_boxes()
        log.log(f"  {len(face_data)} faces, {len(adjacency)} with adjacency")

        # Wall detection
        log.log("\n--- WALL DETECTION ---")
        wall_det = WallDetector()
        top_pid, top_normal = wall_det.find_top_face(face_data, bboxes)
        walls = wall_det.discover_side_walls(
            top_pid, top_normal, face_data, adjacency, bboxes)
        log.log(f"  Top face: pid={top_pid}, {len(walls)} candidate walls")

        # Highlight all walls for user verification
        wall_pids = [w.face_pid for w in walls]
        log.log(f"  All wall pids: {wall_pids}")
        try:
            simplifier = Simplifier(conn)
            simplifier._highlight_faces(shape_name, wall_pids)
        except Exception:
            pass
        input("  All walls highlighted. Press Enter to continue...")
        try:
            conn.execute_vba('Sub Main\n  Pick.ClearAllPicks\nEnd Sub\n')
        except Exception:
            pass

        # Per-wall: set WCS → find dimples → highlight → fill
        log.log("\n--- DIMPLE DETECTION & FILL ---")
        simplifier = Simplifier(conn)
        total_filled = 0
        total_skipped = 0
        total_failed = 0
        walls_with_dimples = 0
        consumed_faces = set()  # track faces already filled

        # Sort walls by bbox area ascending — small walls first claim their dimples,
        # large walls get what's left via consumed tracking
        def _wall_area(w):
            bb = w.bbox
            dx = bb[3]-bb[0]; dy = bb[4]-bb[1]; dz = bb[5]-bb[2]
            dims = sorted([dx, dy, dz], reverse=True)
            return dims[0] * dims[1]
        walls_sorted = sorted(walls, key=_wall_area)

        for wi, wall in enumerate(walls_sorted):
            # Find dimples on this wall
            dimple_faces = wall_det.find_dimple_faces(
                wall, face_data, adjacency, bboxes, top_pid, walls)

            if not dimple_faces:
                continue  # no dimples → skip silently

            # Remove already-consumed faces
            dimple_faces = [f for f in dimple_faces if f not in consumed_faces]
            if not dimple_faces:
                continue  # all dimples already filled

            walls_with_dimples += 1
            bb = wall.bbox
            dx = bb[3]-bb[0]; dy = bb[4]-bb[1]; dz = bb[5]-bb[2]
            log.log(f"\n  Wall {wi+1}/{len(walls_sorted)}: pid={wall.face_pid}, "
                     f"normal=({wall.normal[0]:.2f},{wall.normal[1]:.2f},{wall.normal[2]:.2f}), "
                     f"size=({dx:.2f},{dy:.2f},{dz:.2f})")
            log.log(f"    Dimples: {dimple_faces}")

            # Set WCS aligned with this wall
            try:
                _set_wcs_on_wall(conn, wall)
            except Exception:
                pass

            # Highlight dimple faces
            try:
                simplifier._highlight_faces(shape_name, dimple_faces)
            except Exception:
                pass

            # Ask user
            try:
                action = input(f"  Remove these {len(dimple_faces)} dimple faces? (y/n/q): ").strip().lower()
            except (EOFError, KeyboardInterrupt):
                log.log("    -> INTERRUPTED")
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

            # Fill
            log.log(f"    Filling {len(dimple_faces)} faces...")
            try:
                ok, msg = simplifier._try_fill_hole_silent(
                    shape_name, dimple_faces, walls_with_dimples)
            except Exception as exc:
                ok = False
                msg = str(exc)

            log.log(f"    Result: ok={ok}, msg={msg}")
            if ok:
                total_filled += 1
                consumed_faces.update(dimple_faces)
                log.log(f"    -> FILLED")
            else:
                total_failed += 1
                log.log(f"    -> FAILED")

        # Summary
        log.log(f"\n--- SUMMARY ---")
        log.log(f"  Walls with dimples: {walls_with_dimples}")
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
                'Sub Main\n  WCS.ActivateWCS "global"\n  Pick.ClearAllPicks\nEnd Sub\n')
        except Exception:
            pass
        conn.close()
        log.close()
        print(f"\nLog saved to {log_path}")


if __name__ == "__main__":
    main()
