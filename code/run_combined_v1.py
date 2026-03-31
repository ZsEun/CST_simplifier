"""Combined simplifier — auto-classifies components and applies the right algorithm.

Classifies by name:
- "BOARD" → PCB board simplifier (cone seed detection + fill)
- "COVER" → Shield can cover simplifier (adjacency wall detection + UVW dimple)
- "FRAM"  → Shield can frame simplifier (normal-based wall detection + UVW dimple)
- Other   → Skipped

Run: python -m code.run_combined_v1
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

PROJECT = r"D:\Users\sunze\Desktop\kiro\cst_simplifier\cst_model\Sunray_LED_v5.cst"
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
    if not u: u = (1,0,0) if abs(wn[0])<0.9 else (0,1,0)
    conn.execute_vba(
        'Sub Main\n'
        f'  Open "{_out_vba}" For Output As #1\n'
        '  On Error Resume Next\n'
        f'  WCS.SetOrigin {cx}, {cy}, {cz}\n'
        f'  WCS.SetNormal {wn[0]}, {wn[1]}, {wn[2]}\n'
        f'  WCS.SetUVector {u[0]}, {u[1]}, {u[2]}\n'
        '  WCS.ActivateWCS "local"\n'
        '  Close #1\nEnd Sub\n', output_file=_out_path)


def _classify(name):
    """Classify component by name. Returns 'board', 'cover', 'frame', or None."""
    upper = name.upper()
    if "BOARD" in upper: return "board"
    if "COVER" in upper: return "cover"
    if "FRAM" in upper: return "frame"
    return None


def _wall_area(w):
    b = w.bbox
    dims = sorted([b[3]-b[0], b[4]-b[1], b[5]-b[2]], reverse=True)
    return dims[0] * dims[1]


def _run_shield_can(conn, log, shape_name, detector, simplifier, mode):
    """Run shield can simplifier (cover or frame).
    mode='cover' uses adjacency wall detection.
    mode='frame' uses direct normal wall detection.
    """
    parts = shape_name.split(":")
    sat = detector._export_sat(parts[0], parts[1])
    parser = SATParser(sat)
    face_data = parser.parse()
    adjacency = parser.build_adjacency()
    bboxes = parser.get_bounding_boxes()
    log.log(f"    {len(face_data)} faces, {len(adjacency)} with adjacency")

    wd = WallDetector()
    ref_pid, ref_n = wd.find_top_face(face_data, bboxes)
    log.log(f"    Reference face: pid={ref_pid}, normal={ref_n}")

    # Wall detection
    if mode == "cover":
        walls = wd.discover_side_walls(ref_pid, ref_n, face_data, adjacency, bboxes)
    else:  # frame
        walls = []
        for pid, info in face_data.items():
            if pid == ref_pid: continue
            if info["surface_type"] != "plane-surface": continue
            geom = info.get("geometry", {})
            n = geom.get("normal")
            if n is None: continue
            n = _normalize(n)
            if n == (0.0, 0.0, 0.0): continue
            if abs(_dot(n, ref_n)) <= 0.05:
                bb = bboxes.get(pid, (0,0,0,0,0,0))
                walls.append(WallInfo(face_pid=pid, normal=n, bbox=bb))

    log.log(f"    Side walls: {len(walls)}")
    walls.sort(key=_wall_area)

    consumed = set()
    filled = 0; skipped = 0; failed = 0

    for wi, wall in enumerate(walls):
        dimples = wd.find_dimple_faces(wall, face_data, adjacency, bboxes, ref_pid, walls)
        dimples = [d for d in dimples if d not in consumed]
        if not dimples: continue

        log.log(f"\n    Wall {wi+1}/{len(walls)}: pid={wall.face_pid}, "
                 f"{len(dimples)} dimples")

        try: _set_wcs_on_wall(conn, wall)
        except: pass
        try: simplifier._highlight_faces(shape_name, dimples)
        except: pass

        try:
            action = input(f"  Remove {len(dimples)} faces? (y/n/q): ").strip().lower()
        except (EOFError, KeyboardInterrupt):
            break
        log.log(f"    User: {action}")
        if action == "q": break
        if action == "n":
            try: conn.execute_vba('Sub Main\n  Pick.ClearAllPicks\nEnd Sub\n')
            except: pass
            skipped += 1; continue

        try:
            ok, msg = simplifier._try_fill_hole_silent(shape_name, dimples, wi+1)
        except Exception as exc:
            ok = False; msg = str(exc)
        log.log(f"    Result: ok={ok}")
        if ok: filled += 1; consumed.update(dimples)
        else: failed += 1

    return filled, skipped, failed


def _run_pcb_board(conn, log, shape_name, data, simplifier):
    """Run PCB board simplifier (from run_sunray_v6 logic)."""
    seeds = data["seeds"]
    groups = data.get("seed_groups", [])
    adjacency = data["adjacency"]
    bboxes = data["bboxes"]

    log.log(f"    Seeds: {len(seeds)}, Groups: {len(groups)}")

    consumed = set()
    filled = 0; skipped = 0; failed = 0

    for gi, group in enumerate(groups):
        group_seeds = group["seeds"]
        loop_faces = group["loop_faces"]
        if any(s in consumed for s in group_seeds): continue

        bb = None
        for fid in loop_faces:
            fb = bboxes.get(fid)
            if fb is None: continue
            if bb is None: bb = list(fb)
            else:
                for i in range(3): bb[i] = min(bb[i], fb[i])
                for i in range(3,6): bb[i] = max(bb[i], fb[i])
        bb_t = tuple(bb) if bb else (0,0,0,0,0,0)

        try: simplifier._highlight_faces(shape_name, sorted(set(loop_faces)), zoom_to_bbox=bb_t)
        except: pass

        try:
            action = simplifier._prompt_action()
        except (EOFError, KeyboardInterrupt):
            break
        log.log(f"    Group {gi+1}: seeds={group_seeds}, user={action}")
        if action == "q": break
        if action == "n":
            try: conn.execute_vba('Sub Main\n  Pick.ClearAllPicks\nEnd Sub\n')
            except: pass
            skipped += 1; continue

        sorted_ids = sorted(set(loop_faces))
        try:
            ok, msg = simplifier._try_fill_hole(shape_name, sorted_ids, gi+1)
        except Exception as exc:
            ok = False; msg = str(exc)
        log.log(f"    Result: ok={ok}")
        if ok: filled += 1; consumed.update(loop_faces)
        else: failed += 1

    return filled, skipped, failed


def main():
    log = RunLog(OUT)
    logging.basicConfig(level=logging.INFO)
    conn = CSTConnection()
    try:
        conn.connect(); conn.open_project(PROJECT)
        log.log(f"Opened: {PROJECT}")

        det = FeatureDetector(conn)
        solid_data = det.detect_seeds()
        log.log(f"\n{len(solid_data)} components found:")

        # Classify components
        classified = []
        for data in solid_data:
            name = data["shape_name"]
            ctype = _classify(name)
            classified.append((data, ctype))
            log.log(f"  {name} -> {ctype or 'SKIP'}")

        # Confirm with user — highlight each component separately
        simplifier = Simplifier(conn)
        to_process = []
        for data, ctype in classified:
            if ctype is None:
                log.log(f"\n  SKIP: {data['shape_name']} (unknown type)")
                continue

            shape = data["shape_name"]
            # Highlight the component by picking its first face
            face_ids = list(data["face_types"].keys())
            if face_ids:
                try:
                    simplifier._highlight_faces(shape, face_ids[:5])
                except Exception:
                    pass

            log.log(f"\n  [{ctype.upper()}] {shape}")
            try:
                ok = input(f"  Process this as {ctype.upper()}? (y/n): ").strip().lower()
            except (EOFError, KeyboardInterrupt):
                ok = "n"

            try:
                conn.execute_vba('Sub Main\n  Pick.ClearAllPicks\nEnd Sub\n')
            except Exception:
                pass

            if ok == "y":
                to_process.append((data, ctype))
                log.log(f"    -> CONFIRMED")
            else:
                log.log(f"    -> SKIPPED")

        if not to_process:
            log.log("\nNo components to process."); return
        total_f = 0; total_s = 0; total_fail = 0

        for data, ctype in to_process:
            shape = data["shape_name"]
            log.log(f"\n{'='*60}")
            log.log(f"Processing [{ctype.upper()}]: {shape}")

            if ctype == "board":
                f, s, fail = _run_pcb_board(conn, log, shape, data, simplifier)
            elif ctype in ("cover", "frame"):
                f, s, fail = _run_shield_can(conn, log, shape, det, simplifier, ctype)
            else:
                continue

            log.log(f"  Result: filled={f}, skipped={s}, failed={fail}")
            total_f += f; total_s += s; total_fail += fail

        log.log(f"\n{'='*60}")
        log.log(f"TOTAL: filled={total_f}, skipped={total_s}, failed={total_fail}")

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
