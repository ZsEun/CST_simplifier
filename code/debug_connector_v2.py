"""Connector replacement v2: Replace connector with block, bridge to FPC.

Workflow:
1. User provides connector name, confirm with SelectTreeItem
2. Analyze geometry: longest edge → UVW, flat plane, thickness
3. Reset WCS to global, create replacement Brick
4. Find FPC: keyword search + W-axis proximity check
5. New overlap check: find closest parallel FPC face, check if within block W range
6. If gap: extrude block face toward FPC
7. Delete original connector

Run: python -m code.debug_connector_v2
"""

import os, sys, math, re, time, tempfile, traceback
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from code.cst_connection import CSTConnection
from code.feature_detector import FeatureDetector, SATParser

PROJECT = r"D:\Users\sunze\Desktop\kiro\cst_simplifier\cst_model\Sunray_metal_v3.cst"
OUT = r"D:\Users\sunze\Desktop\kiro\debug_output.txt"
_out_path = os.path.join(tempfile.gettempdir(), "cst_fill.txt")
_out_vba = _out_path.replace("\\", "\\\\")

def _normalize(v):
    mag = math.sqrt(v[0]**2+v[1]**2+v[2]**2)
    if mag < 1e-12: return (0,0,0)
    return (v[0]/mag, v[1]/mag, v[2]/mag)
def _dot(a,b): return a[0]*b[0]+a[1]*b[1]+a[2]*b[2]
def _cross(a,b): return (a[1]*b[2]-a[2]*b[1],a[2]*b[0]-a[0]*b[2],a[0]*b[1]-a[1]*b[0])
def _sub(a,b): return (a[0]-b[0],a[1]-b[1],a[2]-b[2])
def _length(v): return math.sqrt(v[0]**2+v[1]**2+v[2]**2)
def _bbox_center(bb): return ((bb[0]+bb[3])/2,(bb[1]+bb[4])/2,(bb[2]+bb[5])/2)

def run_vba(conn, code):
    try: return conn.execute_vba(code, output_file=_out_path)
    except Exception as exc: return f"EXCEPTION: {exc}"

def list_components(conn):
    result = conn.execute_vba(
        'Sub Main\n'
        f'  Open "{_out_vba}" For Output As #1\n'
        '  Dim rt As Object\n  Set rt = Resulttree\n'
        '  Call WalkTree(rt, "Components", 1)\n'
        '  Close #1\nEnd Sub\n\n'
        'Sub WalkTree(rt As Object, path As String, depth As Integer)\n'
        '  If depth > 10 Then Exit Sub\n'
        '  Dim child As String\n  child = rt.GetFirstChildName(path)\n'
        '  Do While child <> ""\n    Dim subChild As String\n'
        '    subChild = rt.GetFirstChildName(child)\n'
        '    If subChild = "" Then\n      Print #1, child\n'
        '    Else\n      Call WalkTree(rt, child, depth + 1)\n'
        '    End If\n    child = rt.GetNextItemName(child)\n'
        '  Loop\nEnd Sub\n',
        output_file=_out_path)
    comps = []
    if result:
        for line in result.split("\n"):
            line = line.strip()
            if not line: continue
            path = line.replace("\\", "/")
            if path.startswith("Components/"): path = path[len("Components/"):]
            parts = path.split("/")
            if len(parts) >= 2:
                solid = parts[-1]; comp_path = "/".join(parts[:-1])
                comps.append({"comp": comp_path, "solid": solid,
                    "shape": f"{comp_path}:{solid}", "raw": line, "parts": parts})
    return comps

def compute_union_bbox(bboxes):
    mins = [float('inf')]*3; maxs = [float('-inf')]*3
    for bb in bboxes.values():
        for i in range(3): mins[i] = min(mins[i], bb[i])
        for i in range(3): maxs[i] = max(maxs[i], bb[i+3])
    return (mins[0],mins[1],mins[2],maxs[0],maxs[1],maxs[2])

def find_straight_edges(sat_path):
    with open(sat_path, "r", errors="replace") as f: raw = f.read()
    all_lines = raw.split("\n")
    entity_starts = ("body ","lump ","shell ","face ","loop ","coedge ","edge ","vertex ","point ",
        "integer_attrib","name_attrib","rgb_color","simple-snl","cstishape",
        "cone-surface ","plane-surface ","spline-surface ","torus-surface ","sphere-surface ",
        "ellipse-curve ","straight-curve ","intcurve-curve ","pcurve ","cone ","null")
    hdr = 0
    for i, line in enumerate(all_lines):
        s = line.strip()
        if any(s.startswith(t) for t in entity_starts): hdr = i; break
    entities = []
    for line in all_lines[hdr:]:
        stripped = line.strip()
        if not stripped: continue
        if line.startswith("\t"):
            if entities: entities[-1] = entities[-1] + " " + stripped; continue
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
                if 0 <= ref < len(entities) and etype(ref) == "point": vertex_point[i] = ref; break
    edges = []
    for i, ent in enumerate(entities):
        if not ent.strip().startswith("edge "): continue
        refs = get_refs(ent)
        verts = [r for r in refs if 0 <= r < len(entities) and etype(r) == "vertex"]
        if len(verts) < 2: continue
        if not any(0 <= r < len(entities) and etype(r) == "straight-curve" for r in refs): continue
        p1_idx = vertex_point.get(verts[0]); p2_idx = vertex_point.get(verts[1])
        if p1_idx is None or p2_idx is None: continue
        p1 = point_coords.get(p1_idx); p2 = point_coords.get(p2_idx)
        if p1 is None or p2 is None: continue
        l = _length(_sub(p2, p1))
        if l > 1e-6: edges.append((p1, p2, l))
    edges.sort(key=lambda e: e[2], reverse=True)
    return edges

def build_uvw(u_dir, board_normal):
    w = _normalize(board_normal)
    v = _normalize(_cross(w, u_dir))
    u = _normalize(_cross(v, w))
    return u, v, w

def project_bbox(bb, u, v, w, origin=(0,0,0)):
    corners = [(bb[0],bb[1],bb[2]),(bb[3],bb[1],bb[2]),(bb[0],bb[4],bb[2]),(bb[3],bb[4],bb[2]),
               (bb[0],bb[1],bb[5]),(bb[3],bb[1],bb[5]),(bb[0],bb[4],bb[5]),(bb[3],bb[4],bb[5])]
    us=[_dot(_sub(c,origin),u) for c in corners]
    vs=[_dot(_sub(c,origin),v) for c in corners]
    ws=[_dot(_sub(c,origin),w) for c in corners]
    return min(us),max(us),min(vs),max(vs),min(ws),max(ws)


def main():
    f = open(OUT, "w", encoding="utf-8")
    def log(msg): print(msg); f.write(msg+"\n"); f.flush()

    conn = CSTConnection()
    try:
        conn.connect(); conn.open_project(PROJECT)
        log(f"Opened: {PROJECT}")
        det = FeatureDetector(conn)
        components = list_components(conn)

        # ── Step 1: Select connector ──
        log(f"\n=== Step 1: Select connector ===")
        connector_name = input("Enter connector component name: ").strip()
        if not connector_name: log("No name."); return

        matches = [c for c in components if connector_name.upper() in c["shape"].upper()]
        if not matches: log(f"No match for '{connector_name}'"); return
        if len(matches) > 1:
            for i, m in enumerate(matches): log(f"  [{i+1}] {m['shape']}")
            selected = matches[int(input(f"Enter selection (1-{len(matches)}): ")) - 1]
        else:
            selected = matches[0]
        log(f"Selected: {selected['shape']}")
        conn.execute_vba(f'Sub Main\n  SelectTreeItem("{selected["raw"]}")\n  Plot.ZoomToStructure\nEnd Sub\n')
        if input("Correct connector? (y/n): ").strip().lower() != "y": return

        # ── Step 2: Analyze geometry ──
        log(f"\n=== Step 2: Analyze geometry ===")
        parts = selected["shape"].split(":")
        sat_path = det._export_sat(parts[0], parts[1])
        if not sat_path: log("SAT failed."); return
        parser = SATParser(sat_path); faces = parser.parse(); bboxes = parser.get_bounding_boxes()
        union_bb = compute_union_bbox(bboxes); center = _bbox_center(union_bb)

        edges = find_straight_edges(sat_path)
        if not edges: log("No edges."); return
        u_raw = _normalize(_sub(edges[0][1], edges[0][0]))

        best_n = None; best_a = 0
        for pid, info in faces.items():
            if info["surface_type"] != "plane-surface": continue
            bb = bboxes.get(pid)
            if not bb: continue
            dims = sorted([bb[3]-bb[0], bb[4]-bb[1], bb[5]-bb[2]], reverse=True)
            a = dims[0]*dims[1]
            if a > best_a: best_a = a; n = info.get("geometry",{}).get("normal"); best_n = _normalize(n) if n else best_n
        if not best_n: log("No normal."); return

        u_axis, v_axis, w_axis = build_uvw(u_raw, best_n)
        uvw = project_bbox(union_bb, u_axis, v_axis, w_axis, center)
        w_span = uvw[5]-uvw[4]
        w_top = uvw[5]; w_bot = uvw[4]
        log(f"  UVW spans: U={uvw[1]-uvw[0]:.2f}, V={uvw[3]-uvw[2]:.2f}, W={w_span:.2f}")
        log(f"  W range: [{w_bot:.4f}, {w_top:.4f}]")

        # ── Step 3: Create replacement block ──
        log(f"\n=== Step 3: Create block ===")
        conn.execute_vba('Sub Main\n  WCS.ActivateWCS "global"\nEnd Sub\n')
        block_name = f"connector_block_{int(time.time()) % 100000}"
        comp_name = parts[0]
        brick_vba = (
            f'With Brick\n  .Reset\n  .Name "{block_name}"\n  .Component "{comp_name}"\n'
            f'  .Material "PEC"\n  .Xrange "{union_bb[0]}", "{union_bb[3]}"\n'
            f'  .Yrange "{union_bb[1]}", "{union_bb[4]}"\n  .Zrange "{union_bb[2]}", "{union_bb[5]}"\n'
            f'  .Create\nEnd With')
        esc = brick_vba.replace("\\","\\\\").replace('"','""')
        vba_str = '" & vbCrLf & "'.join(esc.split("\n"))
        r = run_vba(conn, (
            'Sub Main\n'
            f'  Open "{_out_vba}" For Output As #1\n  On Error Resume Next\n'
            f'  AddToHistory "create connector block", "{vba_str}"\n'
            '  If Err.Number <> 0 Then\n    Print #1, "FAIL: " & Err.Description\n    Err.Clear\n'
            '  Else\n    Print #1, "OK"\n  End If\n  Close #1\nEnd Sub\n'))
        log(f"  Brick: {r}")
        if "OK" not in (r or ""): log("Failed."); return
        new_shape = f"{comp_name}:{block_name}"
        log(f"  New block: {new_shape}")
        bridge_names = []  # track bridges for merging later

        # ── Step 4: Find FPC and check contact ──
        log(f"\n=== Step 4: Find FPC ===")
        FPC_KW = ["FPC", "FLEX", "FPCA"]
        fpc_candidates = []
        for c in components:
            if c["shape"] == selected["shape"] or c["shape"] == new_shape: continue
            if any(kw in c["solid"].upper() for kw in FPC_KW):
                fpc_candidates.append(c)
        log(f"  FPC keyword matches: {len(fpc_candidates)}")

        fpc_selected = None
        for fc in fpc_candidates:
            conn.execute_vba(f'Sub Main\n  SelectTreeItem("{fc["raw"]}")\n  Plot.ZoomToStructure\nEnd Sub\n')
            ok = input(f"  Is '{fc['solid']}' the FPC? (y/n): ").strip().lower()
            if ok == "y": fpc_selected = fc; break

        if not fpc_selected:
            fpc_name = input("  Enter FPC name manually (or Enter to skip): ").strip()
            if fpc_name:
                fm = [c for c in components if fpc_name.upper() in c["shape"].upper()]
                if fm:
                    if len(fm) == 1: fpc_selected = fm[0]
                    else:
                        for i, m in enumerate(fm[:10]): log(f"    [{i+1}] {m['solid']}")
                        fpc_selected = fm[int(input(f"    Enter selection: ")) - 1]

        if fpc_selected:
            log(f"  FPC: {fpc_selected['solid']}")

            # ── New overlap check ──
            log(f"\n=== Step 4b: Check FPC contact (W-axis method) ===")

            # Get block's top/bottom W positions (already have from step 2)
            log(f"  Block W range: [{w_bot:.4f}, {w_top:.4f}]")

            # Export FPC SAT, find parallel faces
            fpc_parts = fpc_selected["shape"].split(":")
            fpc_sat = det._export_sat(fpc_parts[0], fpc_parts[1])
            if fpc_sat:
                fpc_sp = SATParser(fpc_sat); fpc_faces = fpc_sp.parse(); fpc_bboxes = fpc_sp.get_bounding_boxes()

                # Find FPC plane faces parallel to block's W axis
                closest_fpc_face = None; closest_dist = float('inf')
                for pid, info in fpc_faces.items():
                    if info["surface_type"] != "plane-surface": continue
                    n = info.get("geometry",{}).get("normal")
                    if not n: continue
                    n = _normalize(n)
                    if abs(_dot(n, w_axis)) < 0.9: continue
                    bb = fpc_bboxes.get(pid)
                    if not bb: continue
                    fc = _bbox_center(bb)
                    face_w = _dot(_sub(fc, center), w_axis)
                    # Distance to block center
                    block_w_center = (w_top + w_bot) / 2
                    dist = abs(face_w - block_w_center)
                    if dist < closest_dist:
                        closest_dist = dist; closest_fpc_face = (pid, face_w, bb)

                if closest_fpc_face:
                    fpc_pid, fpc_w, fpc_bb = closest_fpc_face
                    log(f"  Closest FPC face: pid={fpc_pid}, W={fpc_w:.4f}")
                    log(f"  Block W range: [{w_bot:.4f}, {w_top:.4f}]")

                    if w_bot <= fpc_w <= w_top:
                        log(f"  FPC face is WITHIN block W range → in contact, no bridge needed.")
                    else:
                        # Gap exists — need to extend block
                        if fpc_w > w_top:
                            gap = fpc_w - w_top
                            log(f"  FPC face is ABOVE block (gap={gap:.4f} mm)")
                        else:
                            gap = w_bot - fpc_w
                            log(f"  FPC face is BELOW block (gap={gap:.4f} mm)")

                        log(f"  Bridging block to FPC (gap={gap:.4f} mm)")
                        # Find the block face closest to FPC and extrude
                        block_sat = det._export_sat(comp_name, block_name)
                        if block_sat:
                            bsp = SATParser(block_sat); bf = bsp.parse(); bbx = bsp.get_bounding_boxes()
                            best_pid = -1; best_d = float('inf')
                            for pid2, info2 in bf.items():
                                if info2["surface_type"] != "plane-surface": continue
                                n2 = info2.get("geometry",{}).get("normal")
                                if not n2: continue
                                n2 = _normalize(n2)
                                if abs(_dot(n2, w_axis)) < 0.9: continue
                                bb2 = bbx.get(pid2)
                                if not bb2: continue
                                fc2 = _bbox_center(bb2)
                                fw2 = _dot(_sub(fc2, center), w_axis)
                                d2 = abs(fw2 - fpc_w)
                                if d2 < best_d: best_d = d2; best_pid = pid2

                            if best_pid >= 0:
                                conn.execute_vba('Sub Main\n  WCS.ActivateWCS "global"\nEnd Sub\n')
                                pick_vba = f'Pick.PickFaceFromId "{new_shape}", "{best_pid}"'
                                pick_esc = pick_vba.replace('"', '""')
                                run_vba(conn, (
                                    'Sub Main\n'
                                    f'  Open "{_out_vba}" For Output As #1\n  On Error Resume Next\n'
                                    f'  AddToHistory "pick face", "{pick_esc}"\n'
                                    '  If Err.Number <> 0 Then\n    Print #1, "FAIL"\n    Err.Clear\n'
                                    '  Else\n    Print #1, "OK"\n  End If\n  Close #1\nEnd Sub\n'))

                                br_name = f"fpc_bridge_{int(time.time()) % 100000}"
                                ext_vba = (
                                    f'With Extrude\n  .Reset\n  .Name "{br_name}"\n'
                                    f'  .Component "{comp_name}"\n  .Material "PEC"\n  .Mode "Picks"\n'
                                    f'  .Height "{gap}"\n  .Twist "0"\n  .Taper "0"\n'
                                    f'  .UsePicksForHeight "False"\n  .DeleteBaseFaceSolid "False"\n'
                                    f'  .ClearPickedFace "True"\n  .Create\nEnd With')
                                esc2 = ext_vba.replace("\\","\\\\").replace('"','""')
                                vs2 = '" & vbCrLf & "'.join(esc2.split("\n"))
                                r = run_vba(conn, (
                                    'Sub Main\n'
                                    f'  Open "{_out_vba}" For Output As #1\n  On Error Resume Next\n'
                                    f'  AddToHistory "bridge to FPC", "{vs2}"\n'
                                    '  If Err.Number <> 0 Then\n    Print #1, "FAIL: " & Err.Description\n    Err.Clear\n'
                                    '  Else\n    Print #1, "OK"\n  End If\n  Close #1\nEnd Sub\n'))
                                log(f"  FPC bridge: {r}")
                                if "OK" in (r or ""):
                                    bridge_names.append(f"{comp_name}:{br_name}")
                                conn.execute_vba('Sub Main\n  Pick.ClearAllPicks\nEnd Sub\n')
                else:
                    log(f"  No parallel FPC face found.")
        else:
            log("  Skipping FPC.")

        # ── Step 4c: Find PCB and check contact ──
        log(f"\n=== Step 4c: Find PCB ===")
        block_plane_area = (uvw[1]-uvw[0]) * (uvw[3]-uvw[2])  # U_span * V_span
        log(f"  Block plane area (U×V): {block_plane_area:.4f}")
        PCB_KW = ["BOARD", "PCB", "MB", "MAIN_BOARD"]

        # Count keyword matches first (cheap) before expensive SAT exports
        pcb_kw_matches = [c for c in components
            if c["shape"] != selected["shape"] and c["shape"] != new_shape
            and any(kw in c["solid"].upper() for kw in PCB_KW)]
        log(f"  PCB keyword matches: {len(pcb_kw_matches)}")

        skip_auto = False
        if len(pcb_kw_matches) > 100:
            choice = input(f"  {len(pcb_kw_matches)} PCB keyword matches found. Auto-detect may be slow. Wait for auto-detect? (y/n): ").strip().lower()
            if choice != "y":
                skip_auto = True

        pcb_candidates = []
        if not skip_auto:
            for c in pcb_kw_matches:
                c_parts_pcb = c["shape"].split(":")
                try:
                    c_sat = det._export_sat(c_parts_pcb[0], c_parts_pcb[1])
                    if not c_sat: continue
                    c_sp = SATParser(c_sat); c_faces = c_sp.parse(); c_bbs = c_sp.get_bounding_boxes()
                    if not c_bbs: continue
                    # Filter 1: largest plane face area >= 10x block's plane face
                    best_face_area = 0; best_face_normal = None
                    for cpid, cinfo in c_faces.items():
                        if cinfo["surface_type"] != "plane-surface": continue
                        cbb = c_bbs.get(cpid)
                        if not cbb: continue
                        cdims = sorted([cbb[3]-cbb[0], cbb[4]-cbb[1], cbb[5]-cbb[2]], reverse=True)
                        a = cdims[0]*cdims[1]
                        if a > best_face_area:
                            best_face_area = a
                            cn = cinfo.get("geometry",{}).get("normal")
                            best_face_normal = _normalize(cn) if cn else None
                    if best_face_area < block_plane_area * 10:
                        continue
                    # Filter 2: largest face normal parallel to W axis (|dot| >= 0.9)
                    if not best_face_normal or abs(_dot(best_face_normal, w_axis)) < 0.9:
                        continue
                    # Filter 3: W distance < connector thickness
                    c_bb = compute_union_bbox(c_bbs)
                    c_uvw = project_bbox(c_bb, u_axis, v_axis, w_axis, center)
                    w_dist = min(abs(c_uvw[4] - w_bot), abs(c_uvw[4] - w_top),
                                 abs(c_uvw[5] - w_bot), abs(c_uvw[5] - w_top))
                    if w_dist >= w_span:
                        continue
                    pcb_candidates.append((c, w_dist))
                    log(f"  PCB candidate: {c['solid']} (W dist={w_dist:.4f}, face area={best_face_area:.2f})")
                except: continue

        pcb_selected = None
        if pcb_candidates:
            pcb_candidates.sort(key=lambda x: x[1])
            for pc, pd in pcb_candidates:
                conn.execute_vba(f'Sub Main\n  SelectTreeItem("{pc["raw"]}")\n  Plot.ZoomToStructure\nEnd Sub\n')
                ok = input(f"  Is '{pc['solid']}' the PCB? (y/n): ").strip().lower()
                if ok == "y": pcb_selected = pc; break

        if not pcb_selected:
            pcb_name = input("  Enter PCB name manually (or Enter to skip): ").strip()
            if pcb_name:
                pm = [c for c in components if pcb_name.upper() in c["shape"].upper()]
                if pm:
                    if len(pm) == 1: pcb_selected = pm[0]
                    else:
                        for i, m in enumerate(pm[:10]): log(f"    [{i+1}] {m['solid']}")
                        pcb_selected = pm[int(input(f"    Enter selection: ")) - 1]

        if pcb_selected:
            log(f"  PCB: {pcb_selected['solid']}")

            # Same W-axis overlap check as FPC
            log(f"\n=== Step 4d: Check PCB contact (W-axis method) ===")
            pcb_parts = pcb_selected["shape"].split(":")
            pcb_sat = det._export_sat(pcb_parts[0], pcb_parts[1])
            if pcb_sat:
                pcb_sp = SATParser(pcb_sat); pcb_faces = pcb_sp.parse(); pcb_bboxes = pcb_sp.get_bounding_boxes()

                closest_pcb_face = None; closest_pcb_dist = float('inf')
                for pid, info in pcb_faces.items():
                    if info["surface_type"] != "plane-surface": continue
                    n = info.get("geometry",{}).get("normal")
                    if not n: continue
                    n = _normalize(n)
                    if abs(_dot(n, w_axis)) < 0.9: continue
                    bb = pcb_bboxes.get(pid)
                    if not bb: continue
                    fc = _bbox_center(bb)
                    face_w = _dot(_sub(fc, center), w_axis)
                    block_w_center = (w_top + w_bot) / 2
                    dist = abs(face_w - block_w_center)
                    if dist < closest_pcb_dist:
                        closest_pcb_dist = dist; closest_pcb_face = (pid, face_w, bb)

                if closest_pcb_face:
                    pcb_pid, pcb_w, pcb_bb = closest_pcb_face
                    log(f"  Closest PCB face: pid={pcb_pid}, W={pcb_w:.4f}")
                    log(f"  Block W range: [{w_bot:.4f}, {w_top:.4f}]")

                    if w_bot <= pcb_w <= w_top:
                        log(f"  PCB face is WITHIN block W range → in contact.")
                    else:
                        if pcb_w > w_top:
                            gap = pcb_w - w_top
                            log(f"  PCB face is ABOVE block (gap={gap:.4f} mm)")
                        else:
                            gap = w_bot - pcb_w
                            log(f"  PCB face is BELOW block (gap={gap:.4f} mm)")

                        log(f"  Bridging block to PCB (gap={gap:.4f} mm)")
                        block_sat = det._export_sat(comp_name, block_name)
                        if block_sat:
                            bsp = SATParser(block_sat); bf = bsp.parse(); bbx = bsp.get_bounding_boxes()
                            best_pid = -1; best_d = float('inf')
                            for pid2, info2 in bf.items():
                                if info2["surface_type"] != "plane-surface": continue
                                n2 = info2.get("geometry",{}).get("normal")
                                if not n2: continue
                                n2 = _normalize(n2)
                                if abs(_dot(n2, w_axis)) < 0.9: continue
                                bb2 = bbx.get(pid2)
                                if not bb2: continue
                                fc2 = _bbox_center(bb2)
                                fw2 = _dot(_sub(fc2, center), w_axis)
                                d2 = abs(fw2 - pcb_w)
                                if d2 < best_d: best_d = d2; best_pid = pid2

                            if best_pid >= 0:
                                conn.execute_vba('Sub Main\n  WCS.ActivateWCS "global"\nEnd Sub\n')
                                pick_vba = f'Pick.PickFaceFromId "{new_shape}", "{best_pid}"'
                                pick_esc = pick_vba.replace('"', '""')
                                run_vba(conn, (
                                    'Sub Main\n'
                                    f'  Open "{_out_vba}" For Output As #1\n  On Error Resume Next\n'
                                    f'  AddToHistory "pick face", "{pick_esc}"\n'
                                    '  If Err.Number <> 0 Then\n    Print #1, "FAIL"\n    Err.Clear\n'
                                    '  Else\n    Print #1, "OK"\n  End If\n  Close #1\nEnd Sub\n'))

                                br_name = f"pcb_bridge_{int(time.time()) % 100000}"
                                ext_vba = (
                                    f'With Extrude\n  .Reset\n  .Name "{br_name}"\n'
                                    f'  .Component "{comp_name}"\n  .Material "PEC"\n  .Mode "Picks"\n'
                                    f'  .Height "{gap}"\n  .Twist "0"\n  .Taper "0"\n'
                                    f'  .UsePicksForHeight "False"\n  .DeleteBaseFaceSolid "False"\n'
                                    f'  .ClearPickedFace "True"\n  .Create\nEnd With')
                                esc3 = ext_vba.replace("\\","\\\\").replace('"','""')
                                vs3 = '" & vbCrLf & "'.join(esc3.split("\n"))
                                r = run_vba(conn, (
                                    'Sub Main\n'
                                    f'  Open "{_out_vba}" For Output As #1\n  On Error Resume Next\n'
                                    f'  AddToHistory "bridge to PCB", "{vs3}"\n'
                                    '  If Err.Number <> 0 Then\n    Print #1, "FAIL: " & Err.Description\n    Err.Clear\n'
                                    '  Else\n    Print #1, "OK"\n  End If\n  Close #1\nEnd Sub\n'))
                                log(f"  PCB bridge: {r}")
                                if "OK" in (r or ""):
                                    bridge_names.append(f"{comp_name}:{br_name}")
                                conn.execute_vba('Sub Main\n  Pick.ClearAllPicks\nEnd Sub\n')
                else:
                    log(f"  No parallel PCB face found.")
        else:
            log("  Skipping PCB.")

        # ── Step 4e: Merge block and bridges ──
        if bridge_names:
            log(f"\n=== Step 4e: Merge block with bridges ===")
            for br_shape in bridge_names:
                add_vba = f'Solid.Add "{new_shape}", "{br_shape}"'
                add_esc = add_vba.replace('"', '""')
                r = run_vba(conn, (
                    'Sub Main\n'
                    f'  Open "{_out_vba}" For Output As #1\n  On Error Resume Next\n'
                    f'  AddToHistory "merge bridge", "{add_esc}"\n'
                    '  If Err.Number <> 0 Then\n    Print #1, "FAIL: " & Err.Description\n    Err.Clear\n'
                    '  Else\n    Print #1, "OK"\n  End If\n  Close #1\nEnd Sub\n'))
                log(f"  Merge {br_shape}: {r}")

        # ── Step 5: Delete original connector ──
        log(f"\n=== Step 5: Delete original ===")
        ok = input(f"  Delete original connector '{selected['solid']}'? (y/n): ").strip().lower()
        if ok == "y":
            r = run_vba(conn, (
                'Sub Main\n'
                f'  Open "{_out_vba}" For Output As #1\n  On Error Resume Next\n'
                f'  Solid.Delete "{selected["shape"]}"\n'
                '  If Err.Number <> 0 Then\n    Print #1, "FAIL"\n    Err.Clear\n'
                '  Else\n    Print #1, "OK"\n  End If\n  Close #1\nEnd Sub\n'))
            log(f"  Delete: {r}")

        # ── Step 6: Clean up overlapping components ──
        log(f"\n=== Step 6: Clean up overlapping components ===")

        # Get merged block's current bbox
        components = list_components(conn)  # refresh
        block_sat = det._export_sat(comp_name, block_name)
        if block_sat:
            bsp = SATParser(block_sat); bsp.parse(); block_bbs = bsp.get_bounding_boxes()
            if block_bbs:
                block_bb = compute_union_bbox(block_bbs)
                log(f"  Block bbox: ({block_bb[0]:.2f},{block_bb[1]:.2f},{block_bb[2]:.2f}) - ({block_bb[3]:.2f},{block_bb[4]:.2f},{block_bb[5]:.2f})")

                def bbox_within(inner, outer, tol=0.01):
                    """Check if inner bbox is fully within outer bbox."""
                    return (inner[0] >= outer[0] - tol and inner[1] >= outer[1] - tol and inner[2] >= outer[2] - tol and
                            inner[3] <= outer[3] + tol and inner[4] <= outer[4] + tol and inner[5] <= outer[5] + tol)

                def get_level(comp_dict):
                    """Get tree depth level from parts."""
                    return len(comp_dict["parts"])

                def find_overlapping(search_comps, ref_bb, exclude_shapes):
                    """Find components whose bbox is within ref_bb. Shows progress."""
                    found = []
                    total = len(search_comps)
                    for i, c in enumerate(search_comps):
                        if c["shape"] in exclude_shapes: continue
                        if total > 20 and (i+1) % 10 == 0:
                            log(f"    Scanning... {i+1}/{total}")
                        cp = c["shape"].split(":")
                        try:
                            cs = det._export_sat(cp[0], cp[1])
                            if not cs: continue
                            csp = SATParser(cs); csp.parse(); cbs = csp.get_bounding_boxes()
                            if not cbs: continue
                            cbb = compute_union_bbox(cbs)
                            if bbox_within(cbb, ref_bb):
                                found.append(c)
                        except: continue
                    return found

                def search_and_delete(search_comps, ref_bb, exclude_shapes, label=""):
                    """Check count, optionally ask user, then find and delete."""
                    n = len(search_comps)
                    log(f"  {label}{n} candidates to scan...")
                    if n > 100:
                        choice = input(f"  {n} components to scan (may be slow). Auto-detect? (y/n): ").strip().lower()
                        if choice != "y":
                            return  # user will provide names manually in the loop
                    found = find_overlapping(search_comps, ref_bb, exclude_shapes)
                    delete_found_components(found)

                def delete_found_components(found_list):
                    """Ask user to confirm and delete found components."""
                    if not found_list:
                        log(f"  No overlapping components found.")
                        return
                    # Select all found in CST
                    for fc in found_list:
                        conn.execute_vba(f'Sub Main\n  SelectTreeItem("{fc["raw"]}")\nEnd Sub\n')
                    log(f"  Found {len(found_list)} overlapping components:")
                    for fc in found_list:
                        log(f"    {fc['solid']}")
                    ok = input(f"  Delete these {len(found_list)} components? (y/n): ").strip().lower()
                    if ok == "y":
                        for fc in found_list:
                            r = run_vba(conn, (
                                'Sub Main\n'
                                f'  Open "{_out_vba}" For Output As #1\n  On Error Resume Next\n'
                                f'  Solid.Delete "{fc["shape"]}"\n'
                                '  If Err.Number <> 0 Then\n    Print #1, "FAIL"\n    Err.Clear\n'
                                '  Else\n    Print #1, "OK"\n  End If\n  Close #1\nEnd Sub\n'))
                            ok_del = "OK" in (r or "")
                            if ok_del:
                                try:
                                    conn.execute_vba(
                                        'Sub Main\n  On Error Resume Next\n'
                                        f'  AddToHistory "deleted: {fc["solid"]}", "\'shape was: {fc["shape"]}"\n'
                                        'End Sub\n')
                                except: pass
                            log(f"    {fc['solid']}: {'OK' if ok_del else 'FAILED'}")

                # Determine block's tree level
                block_level = len(comp_name.split("/"))
                log(f"  Block level: {block_level}")

                # Search level n-2 to n
                exclude = {new_shape, selected["shape"]}
                search_comps = [c for c in components
                    if c["shape"] not in exclude
                    and block_level - 1 <= len(c["comp"].split("/")) <= block_level]
                search_and_delete(search_comps, block_bb, exclude, f"Levels {block_level-1}-{block_level}: ")

                # Hide block, ask user for more
                conn.execute_vba(f'Sub Main\n  SelectTreeItem("Components\\{comp_name.replace("/", chr(92))}\\{block_name}")\nEnd Sub\n')
                # Make block invisible by selecting it (CST shows others transparent)

                while True:
                    more = input("  Any more components to delete? (y/n): ").strip().lower()
                    if more != "y":
                        break
                    extra_name = input("  Enter component name to search around: ").strip()
                    if not extra_name:
                        break
                    # Find the component
                    components = list_components(conn)  # refresh
                    em = [c for c in components if extra_name.upper() in c["shape"].upper()]
                    if not em:
                        log(f"  No match for '{extra_name}'")
                        continue
                    target = em[0]
                    target_level = len(target["comp"].split("/"))
                    log(f"  Target: {target['shape']} (level {target_level})")
                    # Search level m-1 to m, including the target itself
                    search_comps2 = [c for c in components
                        if c["shape"] not in exclude
                        and target_level - 1 <= len(c["comp"].split("/")) <= target_level]
                    search_and_delete(search_comps2, block_bb, exclude, f"Levels {target_level-1}-{target_level}: ")
                    # Also check if the target itself should be deleted (may not be within block bbox)
                    components = list_components(conn)  # refresh after deletions
                    still_exists = any(c["shape"] == target["shape"] for c in components)
                    if still_exists:
                        ok2 = input(f"  '{target['solid']}' still exists. Delete it? (y/n): ").strip().lower()
                        if ok2 == "y":
                            r = run_vba(conn, (
                                'Sub Main\n'
                                f'  Open "{_out_vba}" For Output As #1\n  On Error Resume Next\n'
                                f'  Solid.Delete "{target["shape"]}"\n'
                                '  If Err.Number <> 0 Then\n    Print #1, "FAIL"\n    Err.Clear\n'
                                '  Else\n    Print #1, "OK"\n  End If\n  Close #1\nEnd Sub\n'))
                            ok_del = "OK" in (r or "")
                            if ok_del:
                                try:
                                    conn.execute_vba(
                                        'Sub Main\n  On Error Resume Next\n'
                                        f'  AddToHistory "deleted: {target["solid"]}", "\'shape was: {target["shape"]}"\n'
                                        'End Sub\n')
                                except: pass
                            log(f"    {target['solid']}: {'OK' if ok_del else 'FAILED'}")

        log(f"\n=== Done ===")

    except Exception as exc:
        log(f"\nERROR: {exc}"); log(traceback.format_exc())
    finally:
        try: conn.execute_vba('Sub Main\n  Pick.ClearAllPicks\n  WCS.ActivateWCS "global"\nEnd Sub\n')
        except: pass
        conn.close(); f.close()
        print(f"\nLog: {OUT}")

if __name__ == "__main__":
    main()
