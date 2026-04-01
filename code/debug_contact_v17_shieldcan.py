"""Debug v17: Shield can cover-frame bridge algorithm.

Algorithm:
1. Compute bbox of cover and frame → find dominant plane (xy) and stack axis (z)
2. Determine stack direction: cover on top → +z, cover on bottom → -z
3. Find frame's top face: plane face with normal along stack axis,
   bbox spans ≥90% of frame bbox in the plane axes, close to frame's zmax
4. Find matching cover face: parallel to frame top face, closest in z
5. Highlight both, compute gap, extrude to bridge

Run: python -m code.debug_contact_v17_shieldcan
"""

import os
import sys
import math
import tempfile
import traceback

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from code.cst_connection import CSTConnection
from code.feature_detector import FeatureDetector, SATParser

PROJECT = r"D:\Users\sunze\Desktop\kiro\cst_simplifier\cst_model\Sunray_LED_v5.cst"
OUT = r"D:\Users\sunze\Desktop\kiro\debug_output.txt"
_out_path = os.path.join(tempfile.gettempdir(), "cst_fill.txt")
_out_vba = _out_path.replace("\\", "\\\\")

SHAPE_COVER = "SUNRAY_PD_LED_SHIELDING_COVER_2:SUNRAY_PD_LED_SHIELDING_COVER_2"
SHAPE_FRAME = "SUNRAY_PD_LED_SHIELDING_FRAM_11:SUNRAY_PD_LED_SHIELDING_FRAM_11"


def _normalize(v):
    mag = math.sqrt(v[0]**2 + v[1]**2 + v[2]**2)
    if mag < 1e-12:
        return (0.0, 0.0, 0.0)
    return (v[0]/mag, v[1]/mag, v[2]/mag)

def _dot(a, b):
    return a[0]*b[0] + a[1]*b[1] + a[2]*b[2]

def _bbox_center(bb):
    return ((bb[0]+bb[3])/2, (bb[1]+bb[4])/2, (bb[2]+bb[5])/2)

def run_vba(conn, code):
    try:
        return conn.execute_vba(code, output_file=_out_path)
    except Exception as exc:
        return f"EXCEPTION: {exc}"

def clear_picks(conn):
    try: conn.execute_vba('Sub Main\n  Pick.ClearAllPicks\nEnd Sub\n')
    except: pass


def compute_union_bbox(bboxes):
    """Compute union bbox from dict of face bboxes."""
    mins = [float('inf')]*3
    maxs = [float('-inf')]*3
    for bb in bboxes.values():
        for i in range(3):
            mins[i] = min(mins[i], bb[i])
            maxs[i] = max(maxs[i], bb[i+3])
    return (mins[0], mins[1], mins[2], maxs[0], maxs[1], maxs[2])


def main():
    f = open(OUT, "w", encoding="utf-8")
    def log(msg):
        print(msg)
        f.write(msg + "\n")
        f.flush()

    conn = CSTConnection()
    try:
        conn.connect()
        conn.open_project(PROJECT)
        log(f"Opened: {PROJECT}")

        det = FeatureDetector(conn)

        # Parse both components
        log(f"\n=== Parsing components ===")
        parts_c = SHAPE_COVER.split(":")
        sat_c = det._export_sat(parts_c[0], parts_c[1])
        parser_c = SATParser(sat_c)
        faces_c = parser_c.parse()
        bboxes_c = parser_c.get_bounding_boxes()

        parts_f = SHAPE_FRAME.split(":")
        sat_f = det._export_sat(parts_f[0], parts_f[1])
        parser_f = SATParser(sat_f)
        faces_f = parser_f.parse()
        bboxes_f = parser_f.get_bounding_boxes()

        log(f"  Cover: {len(faces_c)} faces")
        log(f"  Frame: {len(faces_f)} faces")

        # ── STEP 1: Compute bboxes, find dominant plane ──
        log(f"\n{'='*50}")
        log(f"STEP 1: Component bboxes and dominant plane")

        bbox_c = compute_union_bbox(bboxes_c)
        bbox_f = compute_union_bbox(bboxes_f)
        dims_c = (bbox_c[3]-bbox_c[0], bbox_c[4]-bbox_c[1], bbox_c[5]-bbox_c[2])
        dims_f = (bbox_f[3]-bbox_f[0], bbox_f[4]-bbox_f[1], bbox_f[5]-bbox_f[2])

        log(f"  Cover bbox: ({bbox_c[0]:.2f},{bbox_c[1]:.2f},{bbox_c[2]:.2f}) - "
            f"({bbox_c[3]:.2f},{bbox_c[4]:.2f},{bbox_c[5]:.2f})")
        log(f"  Cover dims: X={dims_c[0]:.2f}, Y={dims_c[1]:.2f}, Z={dims_c[2]:.2f}")
        log(f"  Frame bbox: ({bbox_f[0]:.2f},{bbox_f[1]:.2f},{bbox_f[2]:.2f}) - "
            f"({bbox_f[3]:.2f},{bbox_f[4]:.2f},{bbox_f[5]:.2f})")
        log(f"  Frame dims: X={dims_f[0]:.2f}, Y={dims_f[1]:.2f}, Z={dims_f[2]:.2f}")

        # Find the thin axis (stack axis) — the one with smallest extent
        # Use frame dims to determine
        axis_names = ["X", "Y", "Z"]
        stack_axis = dims_f.index(min(dims_f))  # 0=X, 1=Y, 2=Z
        plane_axes = [i for i in range(3) if i != stack_axis]
        log(f"\n  Stack axis: {axis_names[stack_axis]} (thinnest dimension)")
        log(f"  Plane axes: {axis_names[plane_axes[0]]}, {axis_names[plane_axes[1]]}")

        # ── STEP 2: Determine stack direction ──
        log(f"\n{'='*50}")
        log(f"STEP 2: Stack direction")

        # Cover center vs frame center along stack axis
        center_c = _bbox_center(bbox_c)
        center_f = _bbox_center(bbox_f)
        cover_pos = center_c[stack_axis]
        frame_pos = center_f[stack_axis]

        if cover_pos > frame_pos:
            stack_dir = +1  # cover is above frame → +axis direction
            log(f"  Cover center {axis_names[stack_axis]}={cover_pos:.2f} > "
                f"Frame center {axis_names[stack_axis]}={frame_pos:.2f}")
            log(f"  → Cover is on top, structure goes +{axis_names[stack_axis]}")
        else:
            stack_dir = -1
            log(f"  Cover center {axis_names[stack_axis]}={cover_pos:.2f} < "
                f"Frame center {axis_names[stack_axis]}={frame_pos:.2f}")
            log(f"  → Cover is below, structure goes -{axis_names[stack_axis]}")

        # ── STEP 3: Find frame's top face ──
        log(f"\n{'='*50}")
        log(f"STEP 3: Find frame's top face")

        # Requirements:
        # a. bbox spans ≥90% of frame bbox in both plane axes
        # b. normal is along stack axis (|normal[stack_axis]| > 0.9)
        # c. close to frame's max (if +dir) or min (if -dir) along stack axis
        frame_span_0 = dims_f[plane_axes[0]]
        frame_span_1 = dims_f[plane_axes[1]]
        frame_top_pos = bbox_f[stack_axis + 3] if stack_dir > 0 else bbox_f[stack_axis]

        log(f"  Frame span in {axis_names[plane_axes[0]]}: {frame_span_0:.2f}")
        log(f"  Frame span in {axis_names[plane_axes[1]]}: {frame_span_1:.2f}")
        log(f"  Frame top position ({axis_names[stack_axis]}): {frame_top_pos:.4f}")

        candidates = []
        for pid, info in faces_f.items():
            if info["surface_type"] != "plane-surface":
                continue
            bb = bboxes_f.get(pid)
            if bb is None:
                continue
            geom = info.get("geometry", {})
            n = _normalize(geom.get("normal", (0,0,0)))

            # Check normal along stack axis
            if abs(n[stack_axis]) < 0.9:
                continue

            # Check span in plane axes (≥90% of frame bbox)
            face_span_0 = bb[plane_axes[0]+3] - bb[plane_axes[0]]
            face_span_1 = bb[plane_axes[1]+3] - bb[plane_axes[1]]
            ratio_0 = face_span_0 / frame_span_0 if frame_span_0 > 0 else 0
            ratio_1 = face_span_1 / frame_span_1 if frame_span_1 > 0 else 0

            if ratio_0 < 0.9 or ratio_1 < 0.9:
                continue

            # Distance from frame top
            face_center = _bbox_center(bb)
            dist_to_top = abs(face_center[stack_axis] - frame_top_pos)

            candidates.append({
                "pid": pid, "normal": n, "bbox": bb,
                "span_ratio": (ratio_0, ratio_1),
                "dist_to_top": dist_to_top,
                "center": face_center,
            })

        candidates.sort(key=lambda x: x["dist_to_top"])

        log(f"\n  Candidates (plane faces spanning ≥90% of frame, normal along {axis_names[stack_axis]}):")
        for c in candidates:
            log(f"    face {c['pid']}: dist_to_top={c['dist_to_top']:.4f}, "
                f"span_ratio=({c['span_ratio'][0]:.2f},{c['span_ratio'][1]:.2f}), "
                f"center=({c['center'][0]:.2f},{c['center'][1]:.2f},{c['center'][2]:.2f})")

        if not candidates:
            log("  No frame top face found!")
            return

        frame_top = candidates[0]
        log(f"\n  → Frame top face: {frame_top['pid']} "
            f"(dist_to_top={frame_top['dist_to_top']:.4f})")

        # Highlight frame top face
        conn.execute_vba(
            'Sub Main\n  Pick.ClearAllPicks\n'
            f'  Pick.PickFaceFromId "{SHAPE_FRAME}", "{frame_top["pid"]}"\n'
            '  Plot.ZoomToStructure\nEnd Sub\n'
        )
        ok = input("  Frame top face highlighted. Correct? (y/n): ").strip().lower()
        if ok != "y":
            log("  Wrong face. Stopping.")
            return
        clear_picks(conn)

        # ── STEP 4: Find matching cover face ──
        log(f"\n{'='*50}")
        log(f"STEP 4: Find matching cover face")

        frame_top_n = frame_top["normal"]
        frame_top_center = frame_top["center"]

        cover_candidates = []
        for pid, info in faces_c.items():
            if info["surface_type"] != "plane-surface":
                continue
            bb = bboxes_c.get(pid)
            if bb is None:
                continue
            geom = info.get("geometry", {})
            n = _normalize(geom.get("normal", (0,0,0)))

            # Must be parallel to frame top face
            if abs(_dot(n, frame_top_n)) < 0.95:
                continue

            # Distance along stack axis
            fc = _bbox_center(bb)
            dist = abs(fc[stack_axis] - frame_top_center[stack_axis])

            cover_candidates.append({
                "pid": pid, "normal": n, "bbox": bb,
                "dist": dist, "center": fc,
            })

        cover_candidates.sort(key=lambda x: x["dist"])

        log(f"  Parallel cover faces (sorted by distance to frame top):")
        for c in cover_candidates[:10]:
            log(f"    face {c['pid']}: dist={c['dist']:.4f}, "
                f"center=({c['center'][0]:.2f},{c['center'][1]:.2f},{c['center'][2]:.2f})")

        if not cover_candidates:
            log("  No matching cover face found!")
            return

        cover_mate = cover_candidates[0]
        log(f"\n  → Cover mating face: {cover_mate['pid']} (dist={cover_mate['dist']:.4f})")

        # Highlight cover mating face
        conn.execute_vba(
            'Sub Main\n  Pick.ClearAllPicks\n'
            f'  Pick.PickFaceFromId "{SHAPE_COVER}", "{cover_mate["pid"]}"\n'
            '  Plot.ZoomToStructure\nEnd Sub\n'
        )
        ok = input("  Cover mating face highlighted. Correct? (y/n): ").strip().lower()
        if ok != "y":
            log("  Wrong face. Stopping.")
            return
        clear_picks(conn)

        # ── STEP 5: Compute gap and extrude ──
        log(f"\n{'='*50}")
        log(f"STEP 5: Bridge")

        gap = abs(cover_mate["center"][stack_axis] - frame_top_center[stack_axis])
        log(f"  Gap: {gap:.4f} mm along {axis_names[stack_axis]}")

        # Show both faces
        conn.execute_vba(
            'Sub Main\n  Pick.ClearAllPicks\n'
            f'  Pick.PickFaceFromId "{SHAPE_FRAME}", "{frame_top["pid"]}"\n'
            f'  Pick.PickFaceFromId "{SHAPE_COVER}", "{cover_mate["pid"]}"\n'
            '  Plot.ZoomToStructure\nEnd Sub\n'
        )

        ok = input(f"  Extrude frame face {frame_top['pid']} by {gap:.4f} mm? (y/n): ").strip().lower()
        if ok != "y":
            log("  Cancelled.")
            return

        clear_picks(conn)

        # Pick and extrude
        pick_vba = f'Pick.PickFaceFromId "{SHAPE_FRAME}", "{frame_top["pid"]}"'
        pick_escaped = pick_vba.replace('"', '""')
        result = run_vba(conn, (
            'Sub Main\n'
            f'  Open "{_out_vba}" For Output As #1\n'
            '  On Error Resume Next\n'
            f'  AddToHistory "pick face", "{pick_escaped}"\n'
            '  If Err.Number <> 0 Then\n'
            '    Print #1, "PICK_FAIL: " & Err.Description\n'
            '    Err.Clear\n'
            '  Else\n'
            '    Print #1, "PICK_OK"\n'
            '  End If\n'
            '  Close #1\nEnd Sub\n'
        ))
        log(f"  Pick: {result}")

        extrude_vba = (
            f'With Extrude\n'
            f'  .Reset\n'
            f'  .Name "bridge_{int(__import__("time").time()) % 100000}"\n'
            f'  .Component "{parts_f[0]}"\n'
            f'  .Material "PEC"\n'
            f'  .Mode "Picks"\n'
            f'  .Height "{gap}"\n'
            f'  .Twist "0"\n'
            f'  .Taper "0"\n'
            f'  .UsePicksForHeight "False"\n'
            f'  .DeleteBaseFaceSolid "False"\n'
            f'  .ClearPickedFace "True"\n'
            f'  .Create\n'
            f'End With'
        )
        extrude_escaped = extrude_vba.replace("\\", "\\\\").replace('"', '""')
        lines = extrude_escaped.split("\n")
        vba_str = '" & vbCrLf & "'.join(lines)

        result = run_vba(conn, (
            'Sub Main\n'
            f'  Open "{_out_vba}" For Output As #1\n'
            '  On Error Resume Next\n'
            f'  AddToHistory "extrude face", "{vba_str}"\n'
            '  If Err.Number <> 0 Then\n'
            '    Print #1, "EXTRUDE_FAIL: " & Err.Description\n'
            '    Err.Clear\n'
            '  Else\n'
            '    Print #1, "EXTRUDE_OK"\n'
            '  End If\n'
            '  Close #1\nEnd Sub\n'
        ))
        log(f"  Extrude: {result}")

        if "EXTRUDE_OK" in (result or ""):
            log(f"\n>>> SUCCESS — bridge created!")
        else:
            log(f"\n>>> FAILED")

    except Exception as exc:
        log(f"\nERROR: {exc}")
        log(traceback.format_exc())
    finally:
        clear_picks(conn)
        try: conn.execute_vba('Sub Main\n  WCS.ActivateWCS "global"\nEnd Sub\n')
        except: pass
        conn.close()
        f.close()
        print(f"\nLog: {OUT}")


if __name__ == "__main__":
    main()
