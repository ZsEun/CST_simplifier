"""Debug PCB edge v2: Find PCB, set UVW, validate flatness, list valid PCBs.

Continues from v1:
1. Find all components with PCB keywords
2. For each: find longest straight edge, set UVW, check flatness (4x ratio)
3. List valid PCBs, user picks one

Run: python -m code.debug_pcb_edge_v2
"""

import os, sys, math, re, tempfile, traceback
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from code.cst_connection import CSTConnection
from code.feature_detector import FeatureDetector, SATParser

PROJECT = r"D:\Users\sunze\Desktop\kiro\cst_simplifier\cst_model\Sunray_LED_v7.cst"
OUT = r"D:\Users\sunze\Desktop\kiro\debug_output.txt"
_out_path = os.path.join(tempfile.gettempdir(), "cst_fill.txt")
_out_vba = _out_path.replace("\\", "\\\\")

PCB_KEYWORDS = ["BOARD", "PCB", "MB", "MAIN_BOARD", "MAINBOARD"]

def _normalize(v):
    mag = math.sqrt(v[0]**2+v[1]**2+v[2]**2)
    if mag < 1e-12: return (0,0,0)
    return (v[0]/mag, v[1]/mag, v[2]/mag)
def _sub(a,b): return (a[0]-b[0],a[1]-b[1],a[2]-b[2])
def _length(v): return math.sqrt(v[0]**2+v[1]**2+v[2]**2)
def _dot(a,b): return a[0]*b[0]+a[1]*b[1]+a[2]*b[2]
def _cross(a,b): return (a[1]*b[2]-a[2]*b[1],a[2]*b[0]-a[0]*b[2],a[0]*b[1]-a[1]*b[0])
def is_pcb(name): return any(kw in name.upper() for kw in PCB_KEYWORDS)

def run_vba(conn, code):
    try: return conn.execute_vba(code, output_file=_out_path)
    except Exception as exc: return f"EXCEPTION: {exc}"

def list_components(conn):
    """List all solids recursively, handling any nesting depth."""
    # Use a recursive VBA approach: walk the entire Components tree
    # and collect all leaf nodes (solids)
    result = conn.execute_vba(
        'Sub Main\n'
        f'  Open "{_out_vba}" For Output As #1\n'
        '  Dim rt As Object\n'
        '  Set rt = Resulttree\n'
        '  Call WalkTree(rt, "Components", 1)\n'
        '  Close #1\n'
        'End Sub\n'
        '\n'
        'Sub WalkTree(rt As Object, path As String, depth As Integer)\n'
        '  If depth > 10 Then Exit Sub\n'
        '  Dim child As String\n'
        '  child = rt.GetFirstChildName(path)\n'
        '  Do While child <> ""\n'
        '    Dim subChild As String\n'
        '    subChild = rt.GetFirstChildName(child)\n'
        '    If subChild = "" Then\n'
        '      Print #1, child\n'
        '    Else\n'
        '      Call WalkTree(rt, child, depth + 1)\n'
        '    End If\n'
        '    child = rt.GetNextItemName(child)\n'
        '  Loop\n'
        'End Sub\n',
        output_file=_out_path)
    comps = []
    if result:
        for line in result.split("\n"):
            line = line.strip()
            if not line: continue
            # line is full path like: Components\A\B\solidname
            # or Components\A\solidname
            # Remove "Components\" prefix
            path = line.replace("\\", "/")
            if path.startswith("Components/"):
                path = path[len("Components/"):]
            parts = path.split("/")
            if len(parts) >= 2:
                # Last part = solid name, everything before = component path
                solid = parts[-1]
                comp_path = "/".join(parts[:-1])
                shape = f"{comp_path}:{solid}"
                comps.append((comp_path, solid, shape))
    return comps

def compute_union_bbox(bboxes):
    mins = [float('inf')]*3; maxs = [float('-inf')]*3
    for bb in bboxes.values():
        for i in range(3): mins[i] = min(mins[i], bb[i])
        for i in range(3): maxs[i] = max(maxs[i], bb[i+3])
    return (mins[0],mins[1],mins[2],maxs[0],maxs[1],maxs[2])

def _bbox_center(bb):
    return ((bb[0]+bb[3])/2,(bb[1]+bb[4])/2,(bb[2]+bb[5])/2)

def find_straight_edges(sat_path):
    """Parse SAT to find all straight edges with endpoint coordinates."""
    with open(sat_path, "r", errors="replace") as f: raw = f.read()
    all_lines = raw.split("\n")
    entity_starts = (
        "body ", "lump ", "shell ", "face ", "loop ",
        "coedge ", "edge ", "vertex ", "point ",
        "integer_attrib", "name_attrib", "rgb_color",
        "simple-snl", "cstishape",
        "cone-surface ", "plane-surface ", "spline-surface ",
        "torus-surface ", "sphere-surface ", "ellipse-curve ",
        "straight-curve ", "intcurve-curve ", "pcurve ",
        "cone ", "null",
    )
    hdr = 0
    for i, line in enumerate(all_lines):
        s = line.strip()
        if any(s.startswith(t) for t in entity_starts): hdr = i; break
    entities = []
    for line in all_lines[hdr:]:
        stripped = line.strip()
        if not stripped: continue
        if line.startswith("\t"):
            if entities: entities[-1] = entities[-1] + " " + stripped
            continue
        if entities and not any(stripped.startswith(t) for t in entity_starts):
            if entities and not entities[-1].rstrip().endswith("#"):
                entities[-1] = entities[-1] + " " + stripped; continue
        entities.append(line)

    def get_refs(line): return [int(m) for m in re.findall(r'\$(-?\d+)', line)]
    def etype(idx):
        if 0 <= idx < len(entities): return entities[idx].split()[0] if entities[idx].strip() else ""
        return ""

    point_coords = {}
    for i, ent in enumerate(entities):
        if ent.strip().startswith("point "):
            floats = []
            for n in re.findall(r'([-\d.eE+]+)', ent):
                try: floats.append(float(n))
                except: continue
            if len(floats) >= 3: point_coords[i] = (floats[-3], floats[-2], floats[-1])

    vertex_point = {}
    for i, ent in enumerate(entities):
        if ent.strip().startswith("vertex "):
            for ref in get_refs(ent):
                if 0 <= ref < len(entities) and etype(ref) == "point":
                    vertex_point[i] = ref; break

    edges = []
    for i, ent in enumerate(entities):
        if not ent.strip().startswith("edge "): continue
        refs = get_refs(ent)
        verts = [r for r in refs if 0 <= r < len(entities) and etype(r) == "vertex"]
        if len(verts) < 2: continue
        has_straight = any(0 <= r < len(entities) and etype(r) == "straight-curve" for r in refs)
        if not has_straight: continue
        p1_idx = vertex_point.get(verts[0]); p2_idx = vertex_point.get(verts[1])
        if p1_idx is None or p2_idx is None: continue
        p1 = point_coords.get(p1_idx); p2 = point_coords.get(p2_idx)
        if p1 is None or p2 is None: continue
        l = _length(_sub(p2, p1))
        if l > 1e-6: edges.append((p1, p2, l))
    edges.sort(key=lambda e: e[2], reverse=True)
    return edges

def build_uvw(u_dir, board_normal):
    """Build orthonormal UVW: U along edge, W along board normal, V = W x U."""
    w = _normalize(board_normal)
    v = _normalize(_cross(w, u_dir))
    u = _normalize(_cross(v, w))
    return u, v, w

def project_bbox(bb, u, v, w, origin=(0,0,0)):
    """Project bbox corners into UVW relative to origin."""
    corners = [(bb[0],bb[1],bb[2]),(bb[3],bb[1],bb[2]),(bb[0],bb[4],bb[2]),(bb[3],bb[4],bb[2]),
               (bb[0],bb[1],bb[5]),(bb[3],bb[1],bb[5]),(bb[0],bb[4],bb[5]),(bb[3],bb[4],bb[5])]
    us=[_dot(_sub(c,origin),u) for c in corners]
    vs=[_dot(_sub(c,origin),v) for c in corners]
    ws=[_dot(_sub(c,origin),w) for c in corners]
    return min(us),max(us),min(vs),max(vs),min(ws),max(ws)


def analyze_pcb_candidate(conn, det, shape, log):
    """Analyze a PCB candidate: find edge, set UVW, check flatness.
    Returns dict with UVW info or None if not a valid PCB.
    """
    parts = shape.split(":")
    sat_path = det._export_sat(parts[0], parts[1])
    if not sat_path:
        log(f"    SAT export failed"); return None

    # Parse faces and bboxes
    parser = SATParser(sat_path)
    faces = parser.parse()
    bboxes = parser.get_bounding_boxes()
    if not bboxes:
        log(f"    No bboxes"); return None

    # Find longest straight edge
    edges = find_straight_edges(sat_path)
    if not edges:
        log(f"    No straight edges"); return None

    p1, p2, edge_len = edges[0]
    u_dir = _normalize(_sub(p2, p1))

    # Find board normal from largest plane face
    best_n = None; best_a = 0
    for pid, info in faces.items():
        if info["surface_type"] != "plane-surface": continue
        bb = bboxes.get(pid)
        if not bb: continue
        dims = sorted([bb[3]-bb[0], bb[4]-bb[1], bb[5]-bb[2]], reverse=True)
        a = dims[0] * dims[1]
        if a > best_a:
            best_a = a
            n = info.get("geometry", {}).get("normal")
            if n: best_n = _normalize(n)

    if not best_n:
        log(f"    No plane face found"); return None

    # Build UVW
    u_axis, v_axis, w_axis = build_uvw(u_dir, best_n)

    # Project bbox into UVW
    union_bb = compute_union_bbox(bboxes)
    uvw = project_bbox(union_bb, u_axis, v_axis, w_axis)
    u_span = uvw[1] - uvw[0]
    v_span = uvw[3] - uvw[2]
    w_span = uvw[5] - uvw[4]  # thickness

    # Check flatness: both U and V spans must be >= 4x W span
    if w_span < 1e-6:
        log(f"    Zero thickness"); return None

    u_ratio = u_span / w_span
    v_ratio = v_span / w_span

    log(f"    Edge: {edge_len:.2f} mm, dir=({u_dir[0]:.3f},{u_dir[1]:.3f},{u_dir[2]:.3f})")
    log(f"    UVW spans: U={u_span:.2f}, V={v_span:.2f}, W={w_span:.2f}")
    log(f"    Ratios: U/W={u_ratio:.1f}x, V/W={v_ratio:.1f}x")

    if u_ratio < 4 or v_ratio < 4:
        log(f"    → NOT a PCB (need ≥4x, got {u_ratio:.1f}x and {v_ratio:.1f}x)")
        return None

    log(f"    → VALID PCB")
    return {
        "shape": shape,
        "sat_path": sat_path,
        "faces": faces,
        "bboxes": bboxes,
        "edge": (p1, p2, edge_len),
        "u_axis": u_axis, "v_axis": v_axis, "w_axis": w_axis,
        "uvw": uvw,
        "u_span": u_span, "v_span": v_span, "w_span": w_span,
        "union_bb": union_bb,
    }


def main():
    f = open(OUT, "w", encoding="utf-8")
    def log(msg): print(msg); f.write(msg+"\n"); f.flush()

    conn = CSTConnection()
    try:
        conn.connect(); conn.open_project(PROJECT)
        log(f"Opened: {PROJECT}")

        det = FeatureDetector(conn)
        components = list_components(conn)
        log(f"\nComponents:")
        for i,(c,s,sh) in enumerate(components): log(f"  [{i+1}] {sh}")

        # Find all PCB keyword matches
        log(f"\n=== Step 1: Find PCB candidates ===")
        pcb_cands = [(c,s,sh) for c,s,sh in components if is_pcb(c) or is_pcb(s)]
        log(f"  Keyword matches: {len(pcb_cands)}")

        # Analyze each candidate
        valid_pcbs = []
        for c, s, sh in pcb_cands:
            log(f"\n  Analyzing: {sh}")
            result = analyze_pcb_candidate(conn, det, sh, log)
            if result:
                valid_pcbs.append(result)

        log(f"\n=== Valid PCBs: {len(valid_pcbs)} ===")
        if not valid_pcbs:
            log("  No valid PCB found (need keyword match + flatness ≥4x).")
            return

        for i, pcb in enumerate(valid_pcbs):
            log(f"  [{i+1}] {pcb['shape']}: "
                f"U={pcb['u_span']:.1f}, V={pcb['v_span']:.1f}, W(thickness)={pcb['w_span']:.2f}")

        # User picks
        if len(valid_pcbs) == 1:
            selected = valid_pcbs[0]
        else:
            idx = int(input(f"\n  Select PCB (1-{len(valid_pcbs)}): ")) - 1
            selected = valid_pcbs[idx]

        log(f"\n  → Selected: {selected['shape']}")

        # Show UVW in CST GUI
        p1, p2, edge_len = selected["edge"]
        u = selected["u_axis"]; v = selected["v_axis"]; w = selected["w_axis"]
        center = _bbox_center(selected["union_bb"])

        # First highlight the edge
        mid = ((p1[0]+p2[0])/2, (p1[1]+p2[1])/2, (p1[2]+p2[2])/2)
        run_vba(conn, (
            'Sub Main\n'
            '  Pick.ClearAllPicks\n'
            f'  Pick.PickEdgeFromPoint "{selected["shape"]}", {mid[0]}, {mid[1]}, {mid[2]}\n'
            'End Sub\n'))

        # Show our computed UVW in CST GUI using WCS
        # Origin at PCB center, Normal=W, UVector=U
        conn.execute_vba(
            'Sub Main\n'
            '  Pick.ClearAllPicks\n'
            f'  WCS.SetOrigin {center[0]}, {center[1]}, {center[2]}\n'
            f'  WCS.SetNormal {w[0]}, {w[1]}, {w[2]}\n'
            f'  WCS.SetUVector {u[0]}, {u[1]}, {u[2]}\n'
            '  WCS.ActivateWCS "local"\n'
            '  Plot.ZoomToStructure\n'
            'End Sub\n')

        # Use our computed UVW for all math
        # CST's WCS.Get* methods don't work via COM, so we use our own axes
        log(f"\n  WCS aligned with selected edge (visualization)")
        log(f"  Using computed UVW for math:")
        log(f"    U: ({u[0]:.4f},{u[1]:.4f},{u[2]:.4f})")
        log(f"    V: ({v[0]:.4f},{v[1]:.4f},{v[2]:.4f})")
        log(f"    W: ({w[0]:.4f},{w[1]:.4f},{w[2]:.4f})")
        log(f"  Thickness: {selected['w_span']:.4f} mm")

        log(f"\n  WCS aligned with selected edge")
        log(f"  Computed UVW:")
        log(f"    U: ({u[0]:.4f},{u[1]:.4f},{u[2]:.4f})")
        log(f"    V: ({v[0]:.4f},{v[1]:.4f},{v[2]:.4f})")
        log(f"    W: ({w[0]:.4f},{w[1]:.4f},{w[2]:.4f})")
        log(f"  Thickness: {selected['w_span']:.4f} mm")

        ok = input(f"  Is '{selected['shape']}' the correct PCB? (y/n): ").strip().lower()
        if ok != "y":
            log("  PCB selection rejected. Stopping.")
            return

        # ── Find top and bottom faces ──
        # Use global coordinates projected onto our UVW axes
        # "Top" = highest W value, "Bottom" = lowest W value
        log(f"\n=== Find top and bottom faces ===")
        faces = selected["faces"]
        bboxes = selected["bboxes"]
        pcb_u_span = selected["u_span"]
        pcb_v_span = selected["v_span"]

        top_face = None
        bottom_face = None

        for pid, info in faces.items():
            if info["surface_type"] != "plane-surface":
                continue
            n = info.get("geometry", {}).get("normal")
            if not n: continue
            n = _normalize(n)
            if abs(_dot(n, w)) < 0.9: continue
            bb = bboxes.get(pid)
            if not bb: continue
            fuvw = project_bbox(bb, u, v, w, center)
            fu_span = fuvw[1] - fuvw[0]
            fv_span = fuvw[3] - fuvw[2]
            # Recompute PCB spans relative to center
            pcb_local = project_bbox(selected["union_bb"], u, v, w, center)
            pcb_u_local = pcb_local[1] - pcb_local[0]
            pcb_v_local = pcb_local[3] - pcb_local[2]
            if pcb_u_local > 0 and fu_span / pcb_u_local < 0.8: continue
            if pcb_v_local > 0 and fv_span / pcb_v_local < 0.8: continue
            fw = (fuvw[4] + fuvw[5]) / 2

            log(f"  Candidate: pid={pid}, W={fw:.4f}, "
                f"U_ratio={fu_span/pcb_u_local:.2f}, V_ratio={fv_span/pcb_v_local:.2f}")

            if top_face is None or fw > top_face[1]:
                if top_face and (not bottom_face or top_face[1] < bottom_face[1]):
                    bottom_face = top_face
                top_face = (pid, fw)
            elif not bottom_face or fw < bottom_face[1]:
                bottom_face = (pid, fw)

        if not top_face:
            log("  Face A NOT FOUND"); return
        if not bottom_face:
            log("  Face B NOT FOUND"); return

        thickness = abs(top_face[1] - bottom_face[1])
        # We found two faces — one at higher W, one at lower W
        # Don't assume which is top/bottom — ask the user
        face_high = top_face   # higher W value
        face_low = bottom_face  # lower W value

        log(f"\n  Face A (higher W): pid={face_high[0]}, W={face_high[1]:.4f}")
        log(f"  Face B (lower W): pid={face_low[0]}, W={face_low[1]:.4f}")
        log(f"  Thickness: {thickness:.4f} mm")

        # Highlight face A (higher W)
        log(f"\n  Highlighting Face A (pid={face_high[0]}, higher W)...")
        log(f"  Highlighting Face B (pid={face_low[0]}, lower W)...")

        # No confirmation needed — just assign
        top_pid = face_high[0]; top_w = face_high[1]
        bot_pid = face_low[0]; bot_w = face_low[1]

        log(f"\n  Face A: pid={top_pid}, W={top_w:.4f}")
        log(f"  Face B: pid={bot_pid}, W={bot_w:.4f}")
        log(f"  Thickness: {thickness:.4f} mm")

        # ── Step 4: Find nearby components for Face A (higher W) ──
        threshold = thickness / 4
        face_a_w = face_high[1]  # higher W value
        face_b_w = face_low[1]   # lower W value

        log(f"\n=== Step 4: Find nearby components ===")
        log(f"  Face A (higher W): pid={face_high[0]}, W={face_a_w:.4f}")
        log(f"  Face B (lower W): pid={face_low[0]}, W={face_b_w:.4f}")
        log(f"  Threshold (1/4 thickness): {threshold:.4f} mm")

        pcb_shape_name = selected["shape"]

        # Check Face A side: components with W_min > face_a_w and W_min < face_a_w + threshold
        log(f"\n  --- Components near Face A (W > {face_a_w:.4f}) ---")
        near_a = []
        for comp, solid, shape in components:
            if shape == pcb_shape_name: continue
            c_parts = shape.split(":")
            try:
                c_sat = det._export_sat(c_parts[0], c_parts[1])
            except Exception as exc:
                log(f"    {shape}: SAT FAILED")
                continue
            if not c_sat:
                log(f"    {shape}: SAT RETURNED NONE")
                continue
            c_sp = SATParser(c_sat); c_sp.parse(); c_bbs = c_sp.get_bounding_boxes()
            if not c_bbs: continue
            c_union = compute_union_bbox(c_bbs)
            c_uvw = project_bbox(c_union, u, v, w, center)
            c_w_min = c_uvw[4]
            c_w_max = c_uvw[5]

            # Component must be entirely above face A (W_min >= face_a_w)
            # and close (W_min within threshold of face_a_w)
            gap = c_w_min - face_a_w
            entirely_above = c_w_min >= face_a_w - 0.01  # small tolerance
            within_threshold = gap <= threshold

            log(f"    {shape}: W=[{c_w_min:.4f}, {c_w_max:.4f}], "
                f"gap={gap:.4f}, above={entirely_above}, near={within_threshold}")

            if entirely_above and within_threshold:
                near_a.append((shape, c_uvw, gap))
                log(f"      → CANDIDATE")

        # Check Face B side: components with W_max < face_b_w and W_max > face_b_w - threshold
        log(f"\n  --- Components near Face B (W < {face_b_w:.4f}) ---")
        near_b = []
        for comp, solid, shape in components:
            if shape == pcb_shape_name: continue
            c_parts = shape.split(":")
            try:
                c_sat = det._export_sat(c_parts[0], c_parts[1])
            except Exception as exc:
                continue
            if not c_sat: continue
            c_sp = SATParser(c_sat); c_sp.parse(); c_bbs = c_sp.get_bounding_boxes()
            if not c_bbs: continue
            c_union = compute_union_bbox(c_bbs)
            c_uvw = project_bbox(c_union, u, v, w, center)
            c_w_min = c_uvw[4]
            c_w_max = c_uvw[5]

            gap = face_b_w - c_w_max
            entirely_below = c_w_max <= face_b_w + 0.01
            within_threshold = gap <= threshold

            log(f"    {shape}: W=[{c_w_min:.4f}, {c_w_max:.4f}], "
                f"gap={gap:.4f}, below={entirely_below}, near={within_threshold}")

            if entirely_below and within_threshold:
                near_b.append((shape, c_uvw, gap))
                log(f"      → CANDIDATE")

        log(f"\n=== Summary ===")
        log(f"  Near Face A ({len(near_a)}):")
        for shape, uvw, gap in near_a:
            log(f"    {shape} (gap={gap:.4f} mm)")
        log(f"  Near Face B ({len(near_b)}):")
        for shape, uvw, gap in near_b:
            log(f"    {shape} (gap={gap:.4f} mm)")

        # ── Steps 5-6: Check gaps and bridge ──
        import time as _time

        def copy_shape_local(shape_name):
            before = set(s[2] for s in list_components(conn))
            r = run_vba(conn, (
                'Sub Main\n'
                f'  Open "{_out_vba}" For Output As #1\n  On Error Resume Next\n'
                '  With Transform\n    .Reset\n'
                f'    .Name "{shape_name}"\n    .Vector "0", "0", "0"\n'
                '    .UsePickedPoints "False"\n    .InvertPickedPoints "False"\n'
                '    .MultipleObjects "True"\n    .GroupObjects "False"\n'
                '    .Repetitions "1"\n    .MultipleSelection "False"\n'
                '    .Transform "Shape", "Translate"\n  End With\n'
                '  If Err.Number <> 0 Then\n    Print #1, "FAIL"\n    Err.Clear\n'
                '  Else\n    Print #1, "OK"\n  End If\n  Close #1\nEnd Sub\n'))
            if "OK" not in (r or ""): return None
            after = set(s[2] for s in list_components(conn))
            new = after - before
            return new.pop() if new else None

        def delete_shape_local(shape_name):
            run_vba(conn, (
                'Sub Main\n'
                f'  Open "{_out_vba}" For Output As #1\n  On Error Resume Next\n'
                f'  Solid.Delete "{shape_name}"\n  Print #1, "OK"\n  Close #1\nEnd Sub\n'))

        def check_overlap_local(shape_a, shape_b):
            ca = copy_shape_local(shape_a)
            if not ca: return None
            cb = copy_shape_local(shape_b)
            if not cb: delete_shape_local(ca); return None
            r = run_vba(conn, (
                'Sub Main\n'
                f'  Open "{_out_vba}" For Output As #1\n  On Error Resume Next\n'
                f'  Solid.Intersect "{ca}", "{cb}"\n'
                '  If Err.Number <> 0 Then\n    Print #1, "ERR"\n    Err.Clear\n'
                '  Else\n    Print #1, "OK"\n  End If\n  Close #1\nEnd Sub\n'))
            cur = set(s[2] for s in list_components(conn))
            overlap = False
            if "OK" in (r or "") and ca in cur:
                p = ca.split(":")
                sat = det._export_sat(p[0], p[1])
                if sat:
                    try:
                        sp2 = SATParser(sat); sp2.parse(); bbs2 = sp2.get_bounding_boxes()
                        if bbs2:
                            mins2=[float('inf')]*3; maxs2=[float('-inf')]*3
                            for bb2 in bbs2.values():
                                for i in range(3): mins2[i]=min(mins2[i],bb2[i])
                                for i in range(3): maxs2[i]=max(maxs2[i],bb2[i+3])
                            vol=max(0,maxs2[0]-mins2[0])*max(0,maxs2[1]-mins2[1])*max(0,maxs2[2]-mins2[2])
                            overlap = vol > 1e-9
                    except: pass
                    finally:
                        try: os.remove(sat)
                        except: pass
            if ca in cur: delete_shape_local(ca)
            if cb in cur: delete_shape_local(cb)
            cur2 = set(s[2] for s in list_components(conn))
            if ca in cur2: delete_shape_local(ca)
            if cb in cur2: delete_shape_local(cb)
            return overlap

        def highlight_comp_local(shape_name):
            p = shape_name.split(":")
            sat = det._export_sat(p[0], p[1])
            if not sat: return
            try:
                sp2 = SATParser(sat); fs2 = sp2.parse()
                fids = sorted(fs2.keys())[:5]
                if not fids: return
                picks = "\n".join(f'  Pick.PickFaceFromId "{shape_name}", "{fid}"' for fid in fids)
                conn.execute_vba(f'Sub Main\n  Pick.ClearAllPicks\n{picks}\n  Plot.ZoomToStructure\nEnd Sub\n')
            except: pass
            finally:
                try: os.remove(sat)
                except: pass

        def bridge_to_pcb(comp_shape, pcb_face_w):
            p = comp_shape.split(":")
            sat = det._export_sat(p[0], p[1])
            if not sat: log(f"    SAT export failed"); return False
            sp2 = SATParser(sat); fs2 = sp2.parse(); bbs2 = sp2.get_bounding_boxes()

            best_pid = -1; best_dist = float('inf')
            for pid2, info2 in fs2.items():
                if info2["surface_type"] != "plane-surface": continue
                n2 = info2.get("geometry", {}).get("normal")
                if not n2: continue
                n2 = _normalize(n2)
                if abs(_dot(n2, w)) < 0.9: continue
                bb2 = bbs2.get(pid2)
                if not bb2: continue
                fc = _bbox_center(bb2)
                face_w_val = _dot(_sub(fc, center), w)
                dist = abs(face_w_val - pcb_face_w)
                if dist < best_dist: best_dist = dist; best_pid = pid2

            if best_pid < 0: log(f"    No parallel face found"); return False
            log(f"    Mating face: pid={best_pid}, gap={best_dist:.4f} mm")

            conn.execute_vba(
                'Sub Main\n  Pick.ClearAllPicks\n'
                f'  Pick.PickFaceFromId "{comp_shape}", "{best_pid}"\n'
                '  Plot.ZoomToStructure\nEnd Sub\n')
            ok = input(f"    Extrude face {best_pid} by {best_dist:.4f} mm? (y/n): ").strip().lower()
            conn.execute_vba('Sub Main\n  Pick.ClearAllPicks\nEnd Sub\n')
            if ok != "y": return False

            pick_vba = f'Pick.PickFaceFromId "{comp_shape}", "{best_pid}"'
            pick_esc = pick_vba.replace('"', '""')
            r = run_vba(conn, (
                'Sub Main\n'
                f'  Open "{_out_vba}" For Output As #1\n  On Error Resume Next\n'
                f'  AddToHistory "pick face", "{pick_esc}"\n'
                '  If Err.Number <> 0 Then\n    Print #1, "PICK_FAIL"\n    Err.Clear\n'
                '  Else\n    Print #1, "PICK_OK"\n  End If\n  Close #1\nEnd Sub\n'))
            log(f"    Pick: {r}")
            if "PICK_OK" not in (r or ""): return False

            bname = f"bridge_{int(_time.time()) % 100000}"
            ext_vba = (
                f'With Extrude\n  .Reset\n  .Name "{bname}"\n'
                f'  .Component "{p[0]}"\n  .Material "PEC"\n  .Mode "Picks"\n'
                f'  .Height "{best_dist}"\n  .Twist "0"\n  .Taper "0"\n'
                f'  .UsePicksForHeight "False"\n  .DeleteBaseFaceSolid "False"\n'
                f'  .ClearPickedFace "True"\n  .Create\nEnd With')
            esc = ext_vba.replace("\\", "\\\\").replace('"', '""')
            vba_str = '" & vbCrLf & "'.join(esc.split("\n"))
            r = run_vba(conn, (
                'Sub Main\n'
                f'  Open "{_out_vba}" For Output As #1\n  On Error Resume Next\n'
                f'  AddToHistory "extrude face", "{vba_str}"\n'
                '  If Err.Number <> 0 Then\n    Print #1, "EXTRUDE_FAIL: " & Err.Description\n    Err.Clear\n'
                '  Else\n    Print #1, "EXTRUDE_OK"\n  End If\n  Close #1\nEnd Sub\n'))
            log(f"    Extrude: {r}")
            conn.execute_vba('Sub Main\n  Pick.ClearAllPicks\nEnd Sub\n')
            return "EXTRUDE_OK" in (r or "")

        # Process each side
        for side_name, near_list, pcb_w in [("Face A", near_a, face_a_w), ("Face B", near_b, face_b_w)]:
            if not near_list: continue
            log(f"\n=== Step 5: Process {side_name} ({len(near_list)} candidates) ===")
            for comp_shape, comp_uvw, gap_est in near_list:
                log(f"\n  --- {comp_shape} (est. gap={gap_est:.4f} mm) ---")

                # Find the mating face first (closest parallel face to PCB)
                p = comp_shape.split(":")
                sat = det._export_sat(p[0], p[1])
                if not sat:
                    log(f"    SAT export failed, skipping.")
                    continue
                sp_comp = SATParser(sat); fs_comp = sp_comp.parse(); bbs_comp = sp_comp.get_bounding_boxes()

                mate_pid = -1; mate_dist = float('inf'); mate_bb = None
                for pid2, info2 in fs_comp.items():
                    if info2["surface_type"] != "plane-surface": continue
                    n2 = info2.get("geometry", {}).get("normal")
                    if not n2: continue
                    n2 = _normalize(n2)
                    if abs(_dot(n2, w)) < 0.9: continue
                    bb2 = bbs_comp.get(pid2)
                    if not bb2: continue
                    fc = _bbox_center(bb2)
                    face_w_val = _dot(_sub(fc, center), w)
                    dist = abs(face_w_val - pcb_w)
                    if dist < mate_dist: mate_dist = dist; mate_pid = pid2; mate_bb = bb2

                if mate_pid < 0:
                    log(f"    No parallel face found, skipping.")
                    continue

                # Highlight the mating face and move WCS to its center
                mate_center = _bbox_center(mate_bb)
                conn.execute_vba(
                    'Sub Main\n  Pick.ClearAllPicks\n'
                    f'  Pick.PickFaceFromId "{comp_shape}", "{mate_pid}"\n'
                    f'  WCS.SetOrigin {mate_center[0]}, {mate_center[1]}, {mate_center[2]}\n'
                    f'  WCS.SetNormal {w[0]}, {w[1]}, {w[2]}\n'
                    f'  WCS.SetUVector {u[0]}, {u[1]}, {u[2]}\n'
                    '  WCS.ActivateWCS "local"\n'
                    '  Plot.ZoomToStructure\nEnd Sub\n')

                log(f"    Mating face: pid={mate_pid}, gap={mate_dist:.4f} mm")
                log(f"    WCS moved to face center for visibility")

                log(f"    Checking overlap with PCB...")
                overlap = check_overlap_local(comp_shape, pcb_shape_name)
                if overlap is True:
                    log(f"    → Already in contact. Skipping.")
                    conn.execute_vba('Sub Main\n  Pick.ClearAllPicks\nEnd Sub\n')
                    continue
                elif overlap is None:
                    log(f"    → Error checking overlap. Skipping.")
                    conn.execute_vba('Sub Main\n  Pick.ClearAllPicks\nEnd Sub\n')
                    continue

                log(f"    → Gap confirmed.")
                ok = input(f"    Bridge '{comp_shape}' with the PCB? (y/n/q): ").strip().lower()
                conn.execute_vba('Sub Main\n  Pick.ClearAllPicks\nEnd Sub\n')
                if ok == "q": break
                if ok != "y": continue

                # Bridge: pick + extrude (reuse the mating face we already found)
                pick_vba = f'Pick.PickFaceFromId "{comp_shape}", "{mate_pid}"'
                pick_esc = pick_vba.replace('"', '""')
                r = run_vba(conn, (
                    'Sub Main\n'
                    f'  Open "{_out_vba}" For Output As #1\n  On Error Resume Next\n'
                    f'  AddToHistory "pick face", "{pick_esc}"\n'
                    '  If Err.Number <> 0 Then\n    Print #1, "PICK_FAIL"\n    Err.Clear\n'
                    '  Else\n    Print #1, "PICK_OK"\n  End If\n  Close #1\nEnd Sub\n'))
                log(f"    Pick: {r}")
                if "PICK_OK" not in (r or ""): continue

                bname = f"bridge_{int(_time.time()) % 100000}"
                ext_vba = (
                    f'With Extrude\n  .Reset\n  .Name "{bname}"\n'
                    f'  .Component "{p[0]}"\n  .Material "PEC"\n  .Mode "Picks"\n'
                    f'  .Height "{mate_dist}"\n  .Twist "0"\n  .Taper "0"\n'
                    f'  .UsePicksForHeight "False"\n  .DeleteBaseFaceSolid "False"\n'
                    f'  .ClearPickedFace "True"\n  .Create\nEnd With')
                esc = ext_vba.replace("\\", "\\\\").replace('"', '""')
                vba_str = '" & vbCrLf & "'.join(esc.split("\n"))
                r = run_vba(conn, (
                    'Sub Main\n'
                    f'  Open "{_out_vba}" For Output As #1\n  On Error Resume Next\n'
                    f'  AddToHistory "extrude face", "{vba_str}"\n'
                    '  If Err.Number <> 0 Then\n    Print #1, "EXTRUDE_FAIL: " & Err.Description\n    Err.Clear\n'
                    '  Else\n    Print #1, "EXTRUDE_OK"\n  End If\n  Close #1\nEnd Sub\n'))
                log(f"    Extrude: {r}")
                conn.execute_vba('Sub Main\n  Pick.ClearAllPicks\nEnd Sub\n')
                log(f"    {'>>> Bridge created!' if 'EXTRUDE_OK' in (r or '') else '>>> Bridge failed.'}")

            # Restore WCS to PCB center after processing this side
            conn.execute_vba(
                'Sub Main\n'
                f'  WCS.SetOrigin {center[0]}, {center[1]}, {center[2]}\n'
                f'  WCS.SetNormal {w[0]}, {w[1]}, {w[2]}\n'
                f'  WCS.SetUVector {u[0]}, {u[1]}, {u[2]}\n'
                '  WCS.ActivateWCS "local"\nEnd Sub\n')

        log(f"\n=== Done ===")

    except Exception as exc:
        log(f"\nERROR: {exc}"); log(traceback.format_exc())
    finally:
        conn.close(); f.close()
        print(f"\nLog: {OUT}")


if __name__ == "__main__":
    main()
