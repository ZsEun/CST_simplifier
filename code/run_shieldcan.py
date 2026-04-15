"""Shield can simplifier — removes dimples/holes from cover and frame walls.

Uses keyword-based classification + interactive dialog for component selection,
then runs the existing wall detection + dimple fill algorithm.

Run: python -m code.run_shieldcan
     or via GUI button 2

Set PROJECT before calling main() for GUI mode.
"""

import os, sys, math, tempfile, traceback, logging
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from code.cst_connection import CSTConnection
from code.feature_detector import FeatureDetector, SATParser
from code.wall_detector import WallDetector, WallInfo, _dot, _normalize, group_walls_by_normal, split_subgroups_by_uv_overlap
from code.simplifier import Simplifier
from code.shield_can_dialog import classify_shield_components, pair_cover_frame_by_bbox
import math

PROJECT = None
_out_path = os.path.join(tempfile.gettempdir(), "cst_fill.txt")
_out_vba = _out_path.replace("\\", "\\\\")


def _get_project_path():
    global PROJECT
    if PROJECT:
        return PROJECT
    if len(sys.argv) > 1:
        return os.path.abspath(sys.argv[1].strip().strip('"').strip("'"))
    path = input("Enter the CST project path: ").strip().strip('"').strip("'")
    return os.path.abspath(path)


def _run_one_pass(conn, det, wd, simplifier, shape_name, parts, mode, pass_num):
    """Run one pass of wall detection + dimple removal.
    Returns number of faces consumed.
    """
    sat = det._export_sat(parts[0], parts[1])
    if not sat:
        print(f"  SAT export failed for {shape_name}")
        return 0

    parser = SATParser(sat)
    face_data = parser.parse()
    adjacency = parser.build_adjacency()
    bboxes = parser.get_bounding_boxes()

    ref_pid, ref_n = wd.find_top_face(face_data, bboxes)
    corner_fillets = None

    if mode == "cover":
        walls, corner_fillets = wd.discover_side_walls_validated(ref_pid, ref_n, face_data, adjacency, bboxes)
    else:
        # Frame mode: all plane faces perpendicular to reference normal
        # with W-height filter (wall W span must be > 0.5 × component height)
        w_axis = ref_n
        overall_bb = None
        for pid_bb, bb in bboxes.items():
            if bb == (0,0,0,0,0,0) or bb == (0.0,0.0,0.0,0.0,0.0,0.0): continue
            if overall_bb is None:
                overall_bb = list(bb)
            else:
                for i in range(3): overall_bb[i] = min(overall_bb[i], bb[i])
                for i in range(3, 6): overall_bb[i] = max(overall_bb[i], bb[i])
        if overall_bb:
            corners = [(overall_bb[0],overall_bb[1],overall_bb[2]),(overall_bb[3],overall_bb[1],overall_bb[2]),
                        (overall_bb[0],overall_bb[4],overall_bb[2]),(overall_bb[3],overall_bb[4],overall_bb[2]),
                        (overall_bb[0],overall_bb[1],overall_bb[5]),(overall_bb[3],overall_bb[1],overall_bb[5]),
                        (overall_bb[0],overall_bb[4],overall_bb[5]),(overall_bb[3],overall_bb[4],overall_bb[5])]
            ws = [c[0]*w_axis[0]+c[1]*w_axis[1]+c[2]*w_axis[2] for c in corners]
            component_w_height = max(ws) - min(ws)
        else:
            component_w_height = 1.0
        min_wall_w_span = component_w_height * 0.5

        candidate_walls = []
        for pid, info in face_data.items():
            if pid == ref_pid: continue
            if info["surface_type"] != "plane-surface": continue
            geom = info.get("geometry", {})
            n = geom.get("normal")
            if n is None: continue
            n = _normalize(n)
            if n == (0.0, 0.0, 0.0): continue
            if abs(_dot(n, ref_n)) <= 0.05:
                bb = bboxes.get(pid, (0, 0, 0, 0, 0, 0))
                candidate_walls.append(WallInfo(face_pid=pid, normal=n, bbox=bb))

        walls = []
        for w in candidate_walls:
            bb = w.bbox
            if bb == (0,0,0,0,0,0): continue
            corners = [(bb[0],bb[1],bb[2]),(bb[3],bb[1],bb[2]),
                        (bb[0],bb[4],bb[2]),(bb[3],bb[4],bb[2]),
                        (bb[0],bb[1],bb[5]),(bb[3],bb[1],bb[5]),
                        (bb[0],bb[4],bb[5]),(bb[3],bb[4],bb[5])]
            ws = [c[0]*w_axis[0]+c[1]*w_axis[1]+c[2]*w_axis[2] for c in corners]
            if (max(ws) - min(ws)) >= min_wall_w_span:
                walls.append(w)

    def _wall_area(w):
        b = w.bbox
        dims = sorted([b[3]-b[0], b[4]-b[1], b[5]-b[2]], reverse=True)
        return dims[0] * dims[1]
    walls.sort(key=_wall_area)

    print(f"  Found {len(walls)} walls")

    # Copper thickness
    plane_faces = []
    for pid, info in face_data.items():
        if info["surface_type"] != "plane-surface": continue
        bb = bboxes.get(pid)
        if bb is None: continue
        dims = sorted([bb[3]-bb[0], bb[4]-bb[1], bb[5]-bb[2]], reverse=True)
        n = info.get("geometry", {}).get("normal")
        if n: plane_faces.append((pid, dims[0]*dims[1], _normalize(n), bb))
    plane_faces.sort(key=lambda x: x[1], reverse=True)
    pf1 = plane_faces[0]
    pf2 = next((pf for pf in plane_faces[1:] if abs(_dot(pf1[2], pf[2])) > 0.95), None)
    if pf2:
        pn = pf1[2]
        c1 = tuple((pf1[3][i]+pf1[3][i+3])/2 for i in range(3))
        c2 = tuple((pf2[3][i]+pf2[3][i+3])/2 for i in range(3))
        copper_thickness = abs(sum((c1[i]-c2[i])*pn[i] for i in range(3)))
    else:
        copper_thickness = 0.15
    dist_threshold = copper_thickness * 1.5

    # Sub-grouping
    wall_groups = group_walls_by_normal(walls)

    def _split_by_w_distance(walls_list, threshold):
        if len(walls_list) <= 2: return [walls_list]
        remaining = list(walls_list)
        result_groups = []
        while remaining:
            root = max(remaining, key=_wall_area)
            wn = root.normal
            rc = tuple((root.bbox[i]+root.bbox[i+3])/2 for i in range(3))
            root_w = sum(rc[i]*wn[i] for i in range(3))
            kept, next_remaining = [], []
            for w in remaining:
                wc = tuple((w.bbox[i]+w.bbox[i+3])/2 for i in range(3))
                if abs(sum(wc[i]*wn[i] for i in range(3)) - root_w) <= threshold:
                    kept.append(w)
                else:
                    next_remaining.append(w)
            if len(kept) > 2 and next_remaining:
                result_groups.extend(_split_by_w_distance(kept, threshold))
            elif len(kept) > 2:
                result_groups.append(kept); remaining = []; continue
            else:
                result_groups.append(kept)
            remaining = next_remaining
        return result_groups

    all_sub_groups = []
    for ng in wall_groups:
        all_sub_groups.extend(_split_by_w_distance(ng, dist_threshold))
    all_sub_groups = split_subgroups_by_uv_overlap(all_sub_groups, ref_n)

    consumed = set()
    for gi, group in enumerate(all_sub_groups):
        group_pids = [w.face_pid for w in group]
        all_dimples = set()
        for wall in group:
            all_dimples.update(wd.find_dimple_faces(wall, face_data, adjacency, bboxes, ref_pid, walls,
                                                     corner_fillets=corner_fillets))
        dimples = sorted(d for d in all_dimples if d not in consumed)
        if not dimples: continue

        bb = group[0].bbox
        wn = group[0].normal
        cx=(bb[0]+bb[3])/2; cy=(bb[1]+bb[4])/2; cz=(bb[2]+bb[5])/2
        try:
            conn.execute_vba('Sub Main\n'
                f'  WCS.SetOrigin {cx}, {cy}, {cz}\n'
                f'  WCS.SetNormal {wn[0]}, {wn[1]}, {wn[2]}\n'
                '  WCS.ActivateWCS "local"\nEnd Sub\n')
        except: pass
        try:
            simplifier._highlight_faces(shape_name, dimples)
        except: pass

        label = "Verification: " if pass_num == 2 else ""
        print(f"  {label}Group {gi+1}/{len(all_sub_groups)} (walls {group_pids}): {len(dimples)} dimples")
        action = input(f"  Remove {len(dimples)} faces? (y/n/q): ").strip().lower()
        if action == "q":
            raise RuntimeError("User quit")
        if action != "y":
            try: conn.execute_vba('Sub Main\n  Pick.ClearAllPicks\nEnd Sub\n')
            except: pass
            continue

        ok, msg = simplifier._try_fill_hole_silent(shape_name, dimples, gi+1)
        if ok: consumed.update(dimples)
        print(f"    {'OK' if ok else 'FAILED'}")

    return len(consumed)


def _process_component(conn, det, shape_name, mode, simplifier):
    """Process a single shield can component with two-pass workflow."""
    print(f"\nProcessing [{mode.upper()}]: {shape_name}")
    parts = shape_name.split(":")
    tree_path = f"Components\\{parts[0].replace('/', chr(92))}\\{parts[1]}"
    try:
        conn.execute_vba(f'Sub Main\n  SelectTreeItem("{tree_path}")\n  Plot.ZoomToStructure\nEnd Sub\n')
    except: pass
    print(f"  Now working on: {parts[1]}")

    wd = WallDetector()

    # Pass 1: interactive
    print(f"\n  --- Pass 1: Simplification ---")
    c1 = _run_one_pass(conn, det, wd, simplifier, shape_name, parts, mode, pass_num=1)
    print(f"  Pass 1 done: {c1} faces removed")

    # Pass 2: verification (re-export SAT, re-detect)
    print(f"\n  --- Pass 2: Verification ---")
    c2 = _run_one_pass(conn, det, wd, simplifier, shape_name, parts, mode, pass_num=2)
    if c2 == 0:
        print(f"  Verification passed — no remaining dimples!")
    else:
        print(f"  Pass 2: {c2} additional faces removed")

    print(f"  Done: {c1 + c2} total faces removed from {shape_name}")


def _select_in_cst(conn, shape_name):
    """Select a component in CST GUI. Must be called from the worker thread."""
    parts = shape_name.split(":")
    tree_path = f"Components\\{parts[0].replace('/', chr(92))}\\{parts[1]}"
    try:
        conn.execute_vba(f'Sub Main\n  SelectTreeItem("{tree_path}")\n  Plot.ZoomToStructure\nEnd Sub\n')
    except:
        pass


def main():
    project_path = _get_project_path()

    conn = CSTConnection()
    try:
        conn.connect()
        conn.open_project(project_path)
        print(f"Opened: {project_path}")

        det = FeatureDetector(conn)
        all_solids = det._enumerate_solids()
        print(f"Total solids: {len(all_solids)}")

        # Classify by keyword
        classified = classify_shield_components(all_solids)
        n_cover = len(classified["cover"])
        n_frame = len(classified["frame"])
        n_one = len(classified["one_piece"])
        print(f"Keyword classified: {n_cover} covers, {n_frame} frames, {n_one} one-piece")

        # Pair covers and frames by bbox overlap
        classified = pair_cover_frame_by_bbox(classified, det)

        # Show classifier dialog (GUI mode) or terminal confirmation
        # Check if we're in GUI mode (input is monkey-patched)
        import builtins
        gui_mode = hasattr(builtins, '_shield_can_gui_root')

        if gui_mode:
            import importlib
            import code.shield_can_dialog as _scd
            importlib.reload(_scd)
            from code.shield_can_dialog import show_classifier_dialog
            root = builtins._shield_can_gui_root
            confirmed = show_classifier_dialog(
                root, classified,
                cst_select_fn=lambda s: _select_in_cst(conn, s))
            if confirmed is None:
                print("Cancelled.")
                return
        else:
            # Terminal mode: list components and ask for confirmation
            print("\n--- Shield Can Components ---")
            for group in ["cover", "frame", "one_piece"]:
                items = classified[group]
                if items:
                    label = {"cover": "COVER", "frame": "FRAME", "one_piece": "ONE PIECE"}[group]
                    print(f"\n  {label}:")
                    for i, e in enumerate(items):
                        print(f"    [{i+1}] {e['solid']}")

            ok = input("\nProceed with these components? (y/n): ").strip().lower()
            if ok != "y":
                # Let user manually specify
                print("  You can manually add components.")
                for group in ["cover", "frame"]:
                    name = input(f"  Enter the {group} component name (or Enter to skip): ").strip()
                    if name:
                        matches = [(c, s) for c, s in all_solids if name.upper() in f"{c}:{s}".upper()]
                        if matches:
                            comp, solid = matches[0]
                            classified[group] = [{"comp": comp, "solid": solid, "shape": f"{comp}:{solid}"}]
                        else:
                            print(f"    No match for '{name}'")
            confirmed = classified

        # Process confirmed components
        simplifier = Simplifier(conn)

        for entry in confirmed.get("cover", []):
            shape = entry["shape"]
            # Resolve shape if added manually (comp might be empty)
            if not entry.get("comp"):
                matches = [(c, s) for c, s in all_solids if entry["solid"].upper() in f"{c}:{s}".upper()]
                if matches:
                    shape = f"{matches[0][0]}:{matches[0][1]}"
                else:
                    print(f"  Could not find component: {entry['solid']}")
                    continue
            _process_component(conn, det, shape, "cover", simplifier)

        for entry in confirmed.get("frame", []):
            shape = entry["shape"]
            if not entry.get("comp"):
                matches = [(c, s) for c, s in all_solids if entry["solid"].upper() in f"{c}:{s}".upper()]
                if matches:
                    shape = f"{matches[0][0]}:{matches[0][1]}"
                else:
                    print(f"  Could not find component: {entry['solid']}")
                    continue
            _process_component(conn, det, shape, "frame", simplifier)

        # One piece: skip for now
        if confirmed.get("one_piece"):
            print(f"\n  Skipping {len(confirmed['one_piece'])} one-piece components (not implemented yet)")

        print("\n=== Done ===")

    except RuntimeError as exc:
        if "User quit" in str(exc):
            print("\nUser quit.")
        else:
            print(f"\nERROR: {exc}")
            traceback.print_exc()
    except Exception as exc:
        print(f"\nERROR: {exc}")
        traceback.print_exc()
    finally:
        try:
            conn.execute_vba('Sub Main\n  WCS.ActivateWCS "global"\n  Pick.ClearAllPicks\nEnd Sub\n')
        except:
            pass
        conn.close()


if __name__ == "__main__":
    main()
