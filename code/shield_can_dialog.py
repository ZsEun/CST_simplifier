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
