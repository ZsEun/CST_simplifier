"""Shield can FRAME simplifier — find side walls by normal, detect dimples, fill.

For the frame, wall detection is simpler than the cover:
- Find bottom face (largest plane face)
- Side walls = all plane faces with normal perpendicular to bottom face normal
- Then use the same find_dimple_faces algorithm as the cover

Run: python -m code.run_frame_v1
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
from code.wall_detector import WallDetector, WallInfo, _dot, _normalize
from code.simplifier import Simplifier

PROJECT = r"D:\Users\sunze\Desktop\kiro\cst_simplifier\cst_model\Sunray_LED_v4.cst"
OUT = r"D:\Users\sunze\Desktop\kiro\debug_output.txt"
_out_path = os.path.join(tempfile.gettempdir(), "cst_fill.txt")
_out_vba = _out_path.replace("\\", "\\\\")


class RunLog:
    def __init__(self, path):
        self._f = open(path, "w", encoding="utf-8")
        self._write(f"=== Run started: {datetime.now().isoformat()} ===\n")
    def log(self, msg):
        print(msg); self._write(msg + "\n")
    def _write(self, text):
        self._f.write(text); self._f.flush()
    def close(self):
        self._write(f"\n=== Run ended: {datetime.now().isoformat()} ===\n")
        self._f.close()


def _set_wcs_on_wall(conn, wall):
    wn = wall.normal; bb = wall.bbox
    cx=(bb[0]+bb[3])/2; cy=(bb[1]+bb[4])/2; cz=(bb[2]+bb[5])/2
    dx=bb[3]-bb[0]; dy=bb[4]-bb[1]; dz=bb[5]-bb[2]
    cands = [(dx,(1,0,0)),(dy,(0,1,0)),(dz,(0,0,1))]
    cands.sort(key=lambda c: c[0], reverse=True)
    u = None
    for _, ax in cands:
        if abs(ax[0]*wn[0]+ax[1]*wn[1]+ax[2]*wn[2]) < 0.7:
            u = ax; break
    if not u:
        u = (1,0,0) if abs(wn[0])<0.9 else (0,1,0)
    code = (
        'Sub Main\n'
        f'  Open "{_out_vba}" For Output As #1\n'
        '  On Error Resume Next\n'
        f'  WCS.SetOrigin {cx}, {cy}, {cz}\n'
        f'  WCS.SetNormal {wn[0]}, {wn[1]}, {wn[2]}\n'
        f'  WCS.SetUVector {u[0]}, {u[1]}, {u[2]}\n'
        '  WCS.ActivateWCS "local"\n'
        '  Close #1\n'
        'End Sub\n')
    conn.execute_vba(code, output_file=_out_path)


def _union_bbox(face_ids, bboxes):
    bb = None
    for fid in face_ids:
        fb = bboxes.get(fid)
        if fb is None: continue
        if bb is None: bb = list(fb)
        else:
            for i in range(3): bb[i] = min(bb[i], fb[i])
            for i in range(3,6): bb[i] = max(bb[i], fb[i])
    return tuple(bb) if bb else (0,0,0,0,0,0)


def main():
    log = RunLog(OUT)
    logging.basicConfig(level=logging.INFO)
    conn = CSTConnection()
    try:
        conn.connect(); conn.open_project(PROJECT)
        log.log(f"Opened: {PROJECT}")

        # Detect
        det = FeatureDetector(conn); solid_data = det.detect_seeds()
        frame_shape = None
        for data in solid_data:
            if "FRAM" in data["shape_name"].upper():
                frame_shape = data["shape_name"]; break
        if not frame_shape:
            log.log("No frame!"); return
        log.log(f"\nFrame: {frame_shape}")

        # Parse SAT
        parts = frame_shape.split(":")
        sat = det._export_sat(parts[0], parts[1])
        parser = SATParser(sat)
        face_data = parser.parse()
        adjacency = parser.build_adjacency()
        bboxes = parser.get_bounding_boxes()
        log.log(f"  {len(face_data)} faces, {len(adjacency)} with adjacency")

        # Find bottom face
        wd = WallDetector()
        bot_pid, bot_n = wd.find_top_face(face_data, bboxes)
        log.log(f"\nBottom face: pid={bot_pid}, normal={bot_n}")

        # Find side walls: plane faces with normal ⊥ bottom
        walls = []
        for pid, info in face_data.items():
            if pid == bot_pid: continue
            if info["surface_type"] != "plane-surface": continue
            geom = info.get("geometry", {})
            n = geom.get("normal")
            if n is None: continue
            n = _normalize(n)
            if n == (0.0, 0.0, 0.0): continue
            if abs(_dot(n, bot_n)) <= 0.05:
                bb = bboxes.get(pid, (0,0,0,0,0,0))
                walls.append(WallInfo(face_pid=pid, normal=n, bbox=bb))

        log.log(f"\nSide walls: {len(walls)}")

        # Sort walls small-first
        def _wall_area(w):
            b = w.bbox
            dims = sorted([b[3]-b[0], b[4]-b[1], b[5]-b[2]], reverse=True)
            return dims[0] * dims[1]
        walls.sort(key=_wall_area)

        # Per-wall dimple detection and fill
        log.log("\n--- DIMPLE DETECTION & FILL ---")
        simplifier = Simplifier(conn)
        consumed = set()
        filled = 0; skipped = 0; failed = 0

        for wi, wall in enumerate(walls):
            dimples = wd.find_dimple_faces(
                wall, face_data, adjacency, bboxes, bot_pid, walls)
            dimples = [d for d in dimples if d not in consumed]
            if not dimples:
                continue

            bb = wall.bbox
            dx=bb[3]-bb[0]; dy=bb[4]-bb[1]; dz=bb[5]-bb[2]
            log.log(f"\n  Wall {wi+1}/{len(walls)}: pid={wall.face_pid}, "
                     f"normal=({wall.normal[0]:.2f},{wall.normal[1]:.2f},{wall.normal[2]:.2f}), "
                     f"size=({dx:.2f},{dy:.2f},{dz:.2f})")
            log.log(f"    Dimples: {dimples}")

            try: _set_wcs_on_wall(conn, wall)
            except: pass

            try: simplifier._highlight_faces(frame_shape, dimples)
            except: pass

            try:
                action = input(f"  Remove {len(dimples)} faces? (y/n/q): ").strip().lower()
            except (EOFError, KeyboardInterrupt):
                log.log("    INTERRUPTED"); break

            log.log(f"    User: {action}")
            if action == "q": break
            if action == "n":
                try: conn.execute_vba('Sub Main\n  Pick.ClearAllPicks\nEnd Sub\n')
                except: pass
                skipped += 1; continue

            log.log(f"    Filling {len(dimples)} faces...")
            try:
                ok, msg = simplifier._try_fill_hole_silent(
                    frame_shape, dimples, wi+1)
            except Exception as exc:
                ok = False; msg = str(exc)

            log.log(f"    Result: ok={ok}, msg={msg}")
            if ok:
                filled += 1; consumed.update(dimples)
                log.log(f"    -> FILLED")
            else:
                failed += 1; log.log(f"    -> FAILED")

        log.log(f"\n--- SUMMARY ---")
        log.log(f"  Filled: {filled}, Skipped: {skipped}, Failed: {failed}")

    except Exception as exc:
        log.log(f"\nERROR: {exc}")
        log.log(traceback.format_exc())
    finally:
        try:
            conn.execute_vba(
                'Sub Main\n  WCS.ActivateWCS "global"\n  Pick.ClearAllPicks\nEnd Sub\n')
        except: pass
        conn.close(); log.close()
        print(f"\nLog: {OUT}")


if __name__ == "__main__":
    main()
