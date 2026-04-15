"""Shield can component classifier dialog.

Shows a 3-column dialog for classifying shield can components into
Cover, Frame, and One Piece groups. Each component has buttons to
select in CST, move between groups, or remove.

Usage:
    # From worker thread (GUI mode):
    result = show_classifier_dialog(root, classified, cst_select_fn)
    # result = {"cover": [...], "frame": [...], "one_piece": [...]} or None

    # Classify components by keyword:
    classified = classify_shield_components(all_solids)
"""

import os
import tkinter as tk
from tkinter import ttk
import threading


def classify_shield_components(all_solids):
    """Classify solids into cover/frame/one_piece by keyword matching.

    Args:
        all_solids: list of (comp_path, solid_name) tuples

    Returns:
        dict with keys "cover", "frame", "one_piece", each a list of
        {"comp": comp_path, "solid": solid_name, "shape": "comp:solid"}
    """
    result = {"cover": [], "frame": [], "one_piece": []}

    for comp, solid in all_solids:
        name_upper = solid.upper()
        path_upper = comp.upper()
        full = f"{path_upper}/{name_upper}"

        is_shield = "SHIELD" in full
        is_cover = "COVER" in name_upper
        is_frame = "FRAM" in name_upper

        entry = {"comp": comp, "solid": solid, "shape": f"{comp}:{solid}"}

        if is_cover and is_shield:
            result["cover"].append(entry)
        elif is_frame and is_shield:
            result["frame"].append(entry)
        elif is_shield:
            result["one_piece"].append(entry)

    return result


def pair_cover_frame_by_bbox(classified, det):
    """Pair covers and frames by bbox overlap in UV plane.

    For each cover, export SAT to get its bbox. Find a frame with similar
    UV footprint (>50% overlap in both U and V). Unpaired components
    move to one_piece.

    Args:
        classified: dict from classify_shield_components
        det: FeatureDetector instance (for SAT export)

    Returns:
        Updated classified dict with paired covers/frames and unpaired → one_piece
    """
    from code.feature_detector import SATParser

    def _get_component_bbox(entry):
        """Export SAT and compute overall bbox for a component."""
        try:
            sat = det._export_sat(entry["comp"], entry["solid"])
            if not sat:
                return None
            parser = SATParser(sat)
            face_data = parser.parse()
            bboxes = parser.get_bounding_boxes()
            overall = None
            for pid, bb in bboxes.items():
                if bb == (0,0,0,0,0,0) or bb == (0.0,0.0,0.0,0.0,0.0,0.0):
                    continue
                if overall is None:
                    overall = list(bb)
                else:
                    for i in range(3):
                        overall[i] = min(overall[i], bb[i])
                    for i in range(3, 6):
                        overall[i] = max(overall[i], bb[i])
            return tuple(overall) if overall else None
        except Exception:
            return None

    def _uv_overlap_ratio(bb1, bb2):
        """Compute UV overlap ratio and W proximity. W = thinnest axis."""
        dims1 = [(bb1[i+3]-bb1[i], i) for i in range(3)]
        dims1.sort(key=lambda x: x[0])
        w_idx = dims1[0][1]
        uv_indices = [d[1] for d in dims1[1:]]
        ratios = []
        for ax in uv_indices:
            lo1, hi1 = bb1[ax], bb1[ax+3]
            lo2, hi2 = bb2[ax], bb2[ax+3]
            overlap = max(0, min(hi1, hi2) - max(lo1, lo2))
            min_span = min(hi1-lo1, hi2-lo2)
            ratios.append(overlap / min_span if min_span > 0 else 0)
        # W proximity: check if W ranges touch or overlap
        w_lo1, w_hi1 = bb1[w_idx], bb1[w_idx+3]
        w_lo2, w_hi2 = bb2[w_idx], bb2[w_idx+3]
        w_gap = max(0, max(w_lo1, w_lo2) - min(w_hi1, w_hi2))
        w_span = max(w_hi1-w_lo1, w_hi2-w_lo2)
        w_close = w_gap <= w_span * 2 if w_span > 0 else w_gap < 1.0
        return tuple(ratios), w_close, w_gap

    print("  Computing bboxes for cover/frame pairing...")
    cover_bboxes = {}
    for entry in classified.get("cover", []):
        bb = _get_component_bbox(entry)
        if bb:
            cover_bboxes[entry["shape"]] = bb
    frame_bboxes = {}
    for entry in classified.get("frame", []):
        bb = _get_component_bbox(entry)
        if bb:
            frame_bboxes[entry["shape"]] = bb

    paired_covers = set()
    paired_frames = set()
    for cover_entry in classified.get("cover", []):
        cover_bb = cover_bboxes.get(cover_entry["shape"])
        if cover_bb is None:
            continue
        best_frame = None
        best_score = 0
        for frame_entry in classified.get("frame", []):
            if frame_entry["shape"] in paired_frames:
                continue
            frame_bb = frame_bboxes.get(frame_entry["shape"])
            if frame_bb is None:
                continue
            (u_ratio, v_ratio), w_close, w_gap = _uv_overlap_ratio(cover_bb, frame_bb)
            if not w_close:
                continue  # too far apart in W — not a pair
            score = min(u_ratio, v_ratio)
            if score > best_score and score > 0.5:
                best_score = score
                best_frame = frame_entry
        if best_frame:
            paired_covers.add(cover_entry["shape"])
            paired_frames.add(best_frame["shape"])
            print(f"    Paired: {cover_entry['solid']} <-> {best_frame['solid']} ({best_score:.0%})")

    result = {"cover": [], "frame": [], "one_piece": list(classified.get("one_piece", []))}
    for entry in classified.get("cover", []):
        if entry["shape"] in paired_covers:
            result["cover"].append(entry)
        else:
            result["one_piece"].append(entry)
            print(f"    Unpaired cover → one_piece: {entry['solid']}")
    for entry in classified.get("frame", []):
        if entry["shape"] in paired_frames:
            result["frame"].append(entry)
        else:
            result["one_piece"].append(entry)
            print(f"    Unpaired frame → one_piece: {entry['solid']}")

    print(f"  Pairing: {len(result['cover'])} covers, {len(result['frame'])} frames, "
          f"{len(result['one_piece'])} one-piece")
    return result


def show_classifier_dialog(root, classified, cst_select_fn=None):
    """Show the classifier dialog and return confirmed lists.

    Must be called from a worker thread. Creates the dialog on the main
    thread via root.after() and blocks until the user clicks OK or Cancel.

    CST select requests are queued by the dialog and executed on the
    worker thread (which owns the COM connection).
    """
    result = [None]
    dialog_done = threading.Event()
    select_queue = []

    def _show():
        dialog = ShieldCanClassifierDialog(root, classified, select_queue=select_queue)
        def _on_close():
            result[0] = dialog.result
            dialog_done.set()
        dialog.bind("<Destroy>", lambda e: _on_close() if e.widget == dialog else None)

    root.after(0, _show)

    # Worker thread loop: process CST select requests while waiting for dialog
    while not dialog_done.is_set():
        if select_queue:
            shape = select_queue.pop(0)
            if cst_select_fn:
                try:
                    cst_select_fn(shape)
                except Exception:
                    pass
        dialog_done.wait(timeout=0.1)

    return result[0]


# _do_cst_select is no longer needed — CST calls happen on worker thread


class ShieldCanClassifierDialog(tk.Toplevel):
    """3-column dialog for shield can component classification."""

    GROUPS = ["cover", "frame", "one_piece"]
    LABELS = {"cover": "Shield Can Cover", "frame": "Shield Can Frame",
              "one_piece": "One Piece (skip)"}

    def __init__(self, parent, classified, select_queue=None):
        super().__init__(parent)
        self.title("Shield Can Component Classifier")
        self.transient(parent)
        self.grab_set()
        self.resizable(True, True)
        self.geometry("900x500")

        self._select_queue = select_queue  # list to append shape names for CST select
        self.result = None
        # Track last selected item per group (survives focus loss)
        self._last_sel = {g: None for g in self.GROUPS}

        # Internal data: group_name -> list of entry dicts
        self.data = {g: list(classified.get(g, [])) for g in self.GROUPS}

        # --- Header ---
        hdr = tk.Frame(self, padx=10, pady=5)
        hdr.pack(fill="x")
        tk.Label(hdr, text="Two Piece Shield Can", font=("Arial", 11, "bold")).pack(side="left", padx=20)
        tk.Label(hdr, text="One Piece Shield Can", font=("Arial", 11, "bold")).pack(side="right", padx=20)

        # --- 3 columns ---
        cols_frame = tk.Frame(self, padx=5, pady=5)
        cols_frame.pack(fill="both", expand=True)

        self.listboxes = {}
        self.move_vars = {}

        for i, group in enumerate(self.GROUPS):
            col = tk.LabelFrame(cols_frame, text=self.LABELS[group], padx=5, pady=5)
            col.pack(side="left", fill="both", expand=True, padx=3)

            # Scrollable listbox
            lb_frame = tk.Frame(col)
            lb_frame.pack(fill="both", expand=True)
            sb = tk.Scrollbar(lb_frame)
            sb.pack(side="right", fill="y")
            lb = tk.Listbox(lb_frame, yscrollcommand=sb.set, font=("Consolas", 9),
                            selectmode="single", width=30, height=15,
                            exportselection=False)  # Keep selection when focus lost
            lb.pack(fill="both", expand=True)
            sb.config(command=lb.yview)
            lb.bind("<<ListboxSelect>>", lambda e, g=group: self._on_lb_select(g))
            self.listboxes[group] = lb

            # Buttons for selected item
            btn_frame = tk.Frame(col)
            btn_frame.pack(fill="x", pady=3)
            tk.Button(btn_frame, text="Select in CST",
                      command=lambda g=group: self._select_in_cst(g),
                      width=12).pack(side="left", padx=2)

            # Move dropdown
            move_var = tk.StringVar(value="Move to...")
            move_menu = ttk.Combobox(btn_frame, textvariable=move_var,
                                     values=[self.LABELS[g2] for g2 in self.GROUPS if g2 != group],
                                     state="readonly", width=14)
            move_menu.pack(side="left", padx=2)
            move_menu.bind("<<ComboboxSelected>>",
                           lambda e, g=group, mv=move_var: self._move_item(g, mv))
            self.move_vars[group] = move_var

            tk.Button(btn_frame, text="Remove",
                      command=lambda g=group: self._remove_item(g),
                      width=8, fg="red").pack(side="left", padx=2)

        # --- Add component manually ---
        add_frame = tk.Frame(self, padx=10, pady=5)
        add_frame.pack(fill="x")
        tk.Label(add_frame, text="Add component:").pack(side="left")
        self.add_entry = tk.Entry(add_frame, width=30)
        self.add_entry.pack(side="left", padx=5)
        self.add_group_var = tk.StringVar(value="cover")
        ttk.Combobox(add_frame, textvariable=self.add_group_var,
                     values=self.GROUPS, state="readonly", width=10).pack(side="left", padx=2)
        tk.Button(add_frame, text="Add", command=self._add_manual, width=6).pack(side="left", padx=2)

        # --- OK / Cancel ---
        btn_bottom = tk.Frame(self, padx=10, pady=10)
        btn_bottom.pack(fill="x")
        tk.Button(btn_bottom, text="OK", command=self._ok, width=12,
                  bg="#4CAF50", fg="white").pack(side="left", padx=10)
        tk.Button(btn_bottom, text="Cancel", command=self._cancel, width=12).pack(side="left", padx=10)

        # Populate listboxes
        self._refresh_all()

        # Center dialog
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{max(0,x)}+{max(0,y)}")

        self.protocol("WM_DELETE_WINDOW", self._cancel)

    def _on_lb_select(self, group):
        """Track listbox selection so it survives focus loss."""
        lb = self.listboxes[group]
        sel = lb.curselection()
        if sel:
            self._last_sel[group] = sel[0]

    def _get_selected(self, group):
        """Get selected item, using tracked index if listbox lost focus."""
        lb = self.listboxes[group]
        sel = lb.curselection()
        idx = sel[0] if sel else self._last_sel.get(group)
        if idx is None or idx >= len(self.data[group]):
            return None, None
        return idx, self.data[group][idx]

    def _refresh_all(self):
        for group in self.GROUPS:
            lb = self.listboxes[group]
            lb.delete(0, "end")
            for entry in self.data[group]:
                lb.insert("end", entry["solid"])
            # Restore selection
            saved = self._last_sel.get(group)
            if saved is not None and saved < len(self.data[group]):
                lb.selection_set(saved)

    def _select_in_cst(self, group):
        idx, entry = self._get_selected(group)
        if entry is None:
            return
        if self._select_queue is not None:
            self._select_queue.append(entry["shape"])

    def _move_item(self, from_group, move_var):
        idx, entry = self._get_selected(from_group)
        if entry is None:
            move_var.set("Move to...")
            return
        target_label = move_var.get()
        if target_label == "Move to...":
            return
        target_group = None
        for g in self.GROUPS:
            if self.LABELS[g] == target_label:
                target_group = g
                break
        if target_group is None or target_group == from_group:
            move_var.set("Move to...")
            return
        self.data[from_group].pop(idx)
        self.data[target_group].append(entry)
        self._last_sel[from_group] = None
        self._refresh_all()
        move_var.set("Move to...")

    def _remove_item(self, group):
        idx, entry = self._get_selected(group)
        if entry is None:
            return
        self.data[group].pop(idx)
        self._last_sel[group] = None
        self._refresh_all()

    def _add_manual(self):
        name = self.add_entry.get().strip()
        if not name:
            return
        group = self.add_group_var.get()
        entry = {"comp": "", "solid": name, "shape": name}
        self.data[group].append(entry)
        self._refresh_all()
        self.add_entry.delete(0, "end")

    def _ok(self):
        self.result = dict(self.data)
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()
