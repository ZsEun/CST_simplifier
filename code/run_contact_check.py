"""Combined contact checker + bridge creator.

Full workflow:
1. List components, user picks two
2. Highlight each for confirmation
3. Check overlap via copy + Solid.Intersect + volume check
4. If no overlap: find reference face, mating face, compute gap
5. Extrude mating face to close the gap (with user confirmation)

Run: python -m code.run_contact_check
"""

import os
import sys
import math
import tempfile
import traceback
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from code.cst_connection import CSTConnection
from code.feature_detector import FeatureDetector, SATParser

PROJECT = r"D:\Users\sunze\Desktop\kiro\cst_simplifier\cst_model\Sunray_LED_v7.cst"
OUT = r"D:\Users\sunze\Desktop\kiro\debug_output.txt"
_out_path = os.path.join(tempfile.gettempdir(), "cst_fill.txt")
_out_vba = _out_path.replace("\\", "\\\\")


# ── Helpers ──────────────────────────────────────────────────────────

def _normalize(v):
    mag = math.sqrt(v[0]**2 + v[1]**2 + v[2]**2)
    if mag < 1e-12:
        return (0.0, 0.0, 0.0)
    return (v[0]/mag, v[1]/mag, v[2]/mag)

def _dot(a, b):
    return a[0]*b[0] + a[1]*b[1] + a[2]*b[2]

def _bbox_area(bb):
    dx = bb[3] - bb[0]; dy = bb[4] - bb[1]; dz = bb[5] - bb[2]
    dims = sorted([dx, dy, dz], reverse=True)
    return dims[0] * dims[1]

def _bbox_center(bb):
    return ((bb[0]+bb[3])/2, (bb[1]+bb[4])/2, (bb[2]+bb[5])/2)

def run_vba(conn, code):
    try:
        return conn.execute_vba(code, output_file=_out_path)
    except Exception as exc:
        return f"EXCEPTION: {exc}"


# ── Component listing ────────────────────────────────────────────────

def list_components(conn):
    list_code = (
        'Sub Main\n'
        f'  Open "{_out_vba}" For Output As #1\n'
        '  Dim rt As Object\n'
        '  Set rt = Resulttree\n'
        '  Dim child As String\n'
        '  child = rt.GetFirstChildName("Components")\n'
        '  Do While child <> ""\n'
        '    Dim subChild As String\n'
        '    subChild = rt.GetFirstChildName(child)\n'
        '    Do While subChild <> ""\n'
        '      Print #1, subChild\n'
        '      subChild = rt.GetNextItemName(subChild)\n'
        '    Loop\n'
        '    child = rt.GetNextItemName(child)\n'
        '  Loop\n'
        '  Close #1\n'
        'End Sub\n'
    )
    result = conn.execute_vba(list_code, output_file=_out_path)
    components = []
    if result:
        for line in result.split("\n"):
            line = line.strip()
            if not line:
                continue
            parts = line.replace("\\", "/").split("/")
            if len(parts) >= 3:
                comp = parts[-2]
                solid = parts[-1]
                components.append((comp, solid, f"{comp}:{solid}"))
    return components


# ── Highlight ────────────────────────────────────────────────────────

def highlight_component(conn, shape_name):
    parts = shape_name.split(":")
    det = FeatureDetector(conn)
    sat_path = det._export_sat(parts[0], parts[1])
    if not sat_path:
        return
    try:
        parser = SATParser(sat_path)
        faces = parser.parse()
        face_ids = sorted(faces.keys())[:5]
        if not face_ids:
            return
        pick_lines = "\n".join(
            f'  Pick.PickFaceFromId "{shape_name}", "{fid}"'
            for fid in face_ids
        )
        conn.execute_vba(
            'Sub Main\n  Pick.ClearAllPicks\n'
            f'{pick_lines}\n  Plot.ZoomToStructure\nEnd Sub\n'
        )
    except Exception:
        pass
    finally:
        try: os.remove(sat_path)
        except OSError: pass

def clear_picks(conn):
    try: conn.execute_vba('Sub Main\n  Pick.ClearAllPicks\nEnd Sub\n')
    except: pass


# ── Contact detection (copy + intersect + volume) ────────────────────

def copy_shape(conn, shape_name, log):
    before = set(s[2] for s in list_components(conn))
    code = (
        'Sub Main\n'
        f'  Open "{_out_vba}" For Output As #1\n'
        '  On Error Resume Next\n'
        '  With Transform\n'
        '    .Reset\n'
        f'    .Name "{shape_name}"\n'
        '    .Vector "0", "0", "0"\n'
        '    .UsePickedPoints "False"\n'
        '    .InvertPickedPoints "False"\n'
        '    .MultipleObjects "True"\n'
        '    .GroupObjects "False"\n'
        '    .Repetitions "1"\n'
        '    .MultipleSelection "False"\n'
        '    .Transform "Shape", "Translate"\n'
        '  End With\n'
        '  If Err.Number <> 0 Then\n'
        '    Print #1, "COPY_FAIL: " & Err.Description\n'
        '    Err.Clear\n'
        '  Else\n'
        '    Print #1, "COPY_OK"\n'
        '  End If\n'
        '  Close #1\n'
        'End Sub\n'
    )
    result = run_vba(conn, code)
    if "COPY_OK" not in (result or ""):
        return None
    after = set(s[2] for s in list_components(conn))
    new = after - before
    return new.pop() if new else None

def delete_shape(conn, shape_name):
    run_vba(conn, (
        'Sub Main\n'
        f'  Open "{_out_vba}" For Output As #1\n'
        '  On Error Resume Next\n'
        f'  Solid.Delete "{shape_name}"\n'
        '  Print #1, "DEL_OK"\n'
        '  Close #1\nEnd Sub\n'
    ))

def check_contact(conn, shape_a, shape_b, log):
    """Returns 'overlap', 'no_overlap', or 'error'."""
    log("  Creating temp copies...")
    copy_a = copy_shape(conn, shape_a, log)
    if not copy_a:
        log("  Failed to copy A")
        return "error"
    copy_b = copy_shape(conn, shape_b, log)
    if not copy_b:
        delete_shape(conn, copy_a)
        log("  Failed to copy B")
        return "error"

    log(f"  Copies: {copy_a}, {copy_b}")
    log(f"  Intersecting...")
    result = run_vba(conn, (
        'Sub Main\n'
        f'  Open "{_out_vba}" For Output As #1\n'
        '  On Error Resume Next\n'
        f'  Solid.Intersect "{copy_a}", "{copy_b}"\n'
        '  If Err.Number <> 0 Then\n'
        '    Print #1, "INTERSECT_ERR"\n'
        '    Err.Clear\n'
        '  Else\n'
        '    Print #1, "INTERSECT_OK"\n'
        '  End If\n'
        '  Close #1\nEnd Sub\n'
    ))
    log(f"  Intersect: {result}")

    current = set(s[2] for s in list_components(conn))
    overlap = False
    if "INTERSECT_OK" in (result or "") and copy_a in current:
        # Check volume
        parts = copy_a.split(":")
        det = FeatureDetector(conn)
        sat_path = det._export_sat(parts[0], parts[1])
        if sat_path:
            try:
                parser = SATParser(sat_path)
                faces = parser.parse()
                bboxes = parser.get_bounding_boxes()
                if faces and bboxes:
                    mins = [float('inf')]*3; maxs = [float('-inf')]*3
                    for bb in bboxes.values():
                        for i in range(3): mins[i] = min(mins[i], bb[i])
                        for i in range(3): maxs[i] = max(maxs[i], bb[i+3])
                    vol = max(0, maxs[0]-mins[0]) * max(0, maxs[1]-mins[1]) * max(0, maxs[2]-mins[2])
                    if vol > 1e-9:
                        overlap = True
            except: pass
            finally:
                try: os.remove(sat_path)
                except: pass

    # Cleanup
    if copy_a in current: delete_shape(conn, copy_a)
    if copy_b in current: delete_shape(conn, copy_b)
    remaining = set(s[2] for s in list_components(conn))
    if copy_a in remaining: delete_shape(conn, copy_a)
    if copy_b in remaining: delete_shape(conn, copy_b)

    return "overlap" if overlap else "no_overlap"


# ── Gap characterization + bridge ────────────────────────────────────

def find_largest_plane(face_data, bboxes):
    best = (-1, (0,0,1), 0.0, (0,0,0,0,0,0))
    for pid, info in face_data.items():
        if info["surface_type"] != "plane-surface":
            continue
        bb = bboxes.get(pid)
        if bb is None:
            continue
        area = _bbox_area(bb)
        if area > best[2]:
            n = _normalize(info.get("geometry", {}).get("normal", (0,0,1)))
            best = (pid, n, area, bb)
    return best

def find_closest_parallel(ref_normal, ref_bbox, face_data, bboxes):
    ref_center = _bbox_center(ref_bbox)
    best = None
    for pid, info in face_data.items():
        if info["surface_type"] != "plane-surface":
            continue
        geom = info.get("geometry", {})
        n = geom.get("normal")
        if n is None:
            continue
        n = _normalize(n)
        if abs(_dot(n, ref_normal)) < 0.95:
            continue
        bb = bboxes.get(pid)
        if bb is None:
            continue
        fc = _bbox_center(bb)
        diff = (fc[0]-ref_center[0], fc[1]-ref_center[1], fc[2]-ref_center[2])
        dist = abs(_dot(diff, ref_normal))
        if best is None or dist < best[2]:
            best = (pid, n, dist, bb)
    return best

def characterize_and_bridge(conn, shape_a, shape_b, log):
    """Find reference face, mating face, compute gap, extrude to bridge."""
    det = FeatureDetector(conn)

    # Parse both
    parts_a = shape_a.split(":")
    sat_a = det._export_sat(parts_a[0], parts_a[1])
    parser_a = SATParser(sat_a)
    faces_a = parser_a.parse()
    bboxes_a = parser_a.get_bounding_boxes()

    parts_b = shape_b.split(":")
    sat_b = det._export_sat(parts_b[0], parts_b[1])
    parser_b = SATParser(sat_b)
    faces_b = parser_b.parse()
    bboxes_b = parser_b.get_bounding_boxes()

    # Find largest plane face in each
    pid_a, n_a, area_a, bb_a = find_largest_plane(faces_a, bboxes_a)
    pid_b, n_b, area_b, bb_b = find_largest_plane(faces_b, bboxes_b)
    log(f"  Largest face A: pid={pid_a}, area={area_a:.1f}")
    log(f"  Largest face B: pid={pid_b}, area={area_b:.1f}")

    # Reference = bigger face
    if area_a >= area_b:
        ref_pid, ref_n, ref_bb, ref_shape = pid_a, n_a, bb_a, shape_a
        other_faces, other_bboxes, other_shape = faces_b, bboxes_b, shape_b
        other_comp = parts_b[0]
    else:
        ref_pid, ref_n, ref_bb, ref_shape = pid_b, n_b, bb_b, shape_b
        other_faces, other_bboxes, other_shape = faces_a, bboxes_a, shape_a
        other_comp = parts_a[0]

    log(f"  Reference: face {ref_pid} on {ref_shape}")

    # Find closest parallel face on other component
    result = find_closest_parallel(ref_n, ref_bb, other_faces, other_bboxes)
    if result is None:
        log("  No parallel mating face found.")
        return False

    mate_pid, mate_n, mate_dist, mate_bb = result
    log(f"  Mating: face {mate_pid} on {other_shape}, gap={mate_dist:.4f} mm")

    # Compute extrusion distance (always positive)
    ref_center = _bbox_center(ref_bb)
    mate_center = _bbox_center(mate_bb)
    ref_pos = _dot(ref_center, ref_n)
    mate_pos = _dot(mate_center, ref_n)
    extrude_dist = abs(mate_pos - ref_pos)
    log(f"  Extrude distance: {extrude_dist:.4f} mm")

    # Highlight both faces
    conn.execute_vba(
        'Sub Main\n  Pick.ClearAllPicks\n'
        f'  Pick.PickFaceFromId "{ref_shape}", "{ref_pid}"\n'
        f'  Pick.PickFaceFromId "{other_shape}", "{mate_pid}"\n'
        '  Plot.ZoomToStructure\nEnd Sub\n'
    )

    log(f"\n  Reference face (large): {ref_pid} on {ref_shape}")
    log(f"  Mating face (to extrude): {mate_pid} on {other_shape}")
    log(f"  Gap: {extrude_dist:.4f} mm")

    ok = input(f"\n  Create bridge by extruding face {mate_pid} by {extrude_dist:.4f} mm? (y/n): ").strip().lower()
    if ok != "y":
        log("  Cancelled.")
        clear_picks(conn)
        return False

    # Pick mating face and extrude via AddToHistory
    clear_picks(conn)

    # Pick face
    pick_vba = f'Pick.PickFaceFromId "{other_shape}", "{mate_pid}"'
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
    if "PICK_OK" not in (result or ""):
        log("  Pick failed.")
        return False

    # Extrude via AddToHistory
    extrude_vba = (
        f'With Extrude\n'
        f'  .Reset\n'
        f'  .Name "bridge_1"\n'
        f'  .Component "{other_comp}"\n'
        f'  .Material "PEC"\n'
        f'  .Mode "Picks"\n'
        f'  .Height "{extrude_dist}"\n'
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

    clear_picks(conn)

    if "EXTRUDE_OK" in (result or ""):
        log(f"\n  >>> Bridge created! Face {mate_pid} extruded by {extrude_dist:.4f} mm")
        return True
    else:
        log(f"\n  >>> Extrude failed: {result}")
        return False


# ── Main ─────────────────────────────────────────────────────────────

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
        log(f"Time: {datetime.now().isoformat()}")

        # Step 1: List components
        components = list_components(conn)
        log(f"\nComponents:")
        for i, (comp, solid, shape) in enumerate(components):
            log(f"  [{i+1}] {shape}")

        if len(components) < 2:
            log("Need at least 2 components.")
            return

        # Step 2: User picks two with visual confirmation
        a_idx = int(input("\nComponent A (number): ")) - 1
        shape_a = components[a_idx][2]
        log(f"\nSelected A: {shape_a}")
        highlight_component(conn, shape_a)
        ok = input("  Is this correct? (y/n): ").strip().lower()
        clear_picks(conn)
        if ok != "y":
            return

        b_idx = int(input("Component B (number): ")) - 1
        shape_b = components[b_idx][2]
        log(f"Selected B: {shape_b}")
        highlight_component(conn, shape_b)
        ok = input("  Is this correct? (y/n): ").strip().lower()
        clear_picks(conn)
        if ok != "y":
            return

        # Step 3: Check contact
        log(f"\n=== Step 1: Contact Check ===")
        status = check_contact(conn, shape_a, shape_b, log)
        log(f"  Result: {status}")

        if status == "overlap":
            log(f"\n>>> Components OVERLAP — no bridge needed.")
            return

        if status == "error":
            log(f"\n>>> Error during contact check.")
            return

        # Step 4: Gap characterization + bridge
        log(f"\n=== Step 2: Gap Characterization + Bridge ===")
        success = characterize_and_bridge(conn, shape_a, shape_b, log)

        if success:
            log(f"\n>>> Done! Bridge created between {shape_a} and {shape_b}.")
        else:
            log(f"\n>>> Bridge creation cancelled or failed.")

    except Exception as exc:
        log(f"\nERROR: {exc}")
        log(traceback.format_exc())
    finally:
        clear_picks(conn)
        try:
            conn.execute_vba(
                'Sub Main\n  WCS.ActivateWCS "global"\nEnd Sub\n')
        except: pass
        conn.close()
        f.close()
        print(f"\nLog: {OUT}")


if __name__ == "__main__":
    main()
