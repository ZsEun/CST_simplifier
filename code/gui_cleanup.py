"""GUI for component cleanup — delete plastic/unnecessary components.

Separate tool from the main simplifier GUI.
- Editable keyword lists (delete + exclude)
- Browse CST project
- Scan and show matched components
- Delete/Skip/Quit with SelectTreeItem highlighting
- Save keywords for future use

Run: python -m code.gui_cleanup
"""

import os, sys, re, threading, tempfile, traceback
import tkinter as tk
from tkinter import filedialog, messagebox

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

_out_path = os.path.join(tempfile.gettempdir(), "cst_fill.txt")
_out_vba = _out_path.replace("\\", "\\\\")

DEFAULT_DELETE_KW = ["COVER"]
DEFAULT_AUTO_KW = ["SCREW", "RUBBER", "MYLAR", "ADH", "PLASTIC"]
DEFAULT_EXCLUDE_KW = ["SHIELDING", "SPRING"]
KW_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cleanup_keywords.txt")


def load_keywords():
    delete_kw = list(DEFAULT_DELETE_KW)
    auto_kw = list(DEFAULT_AUTO_KW)
    exclude_kw = list(DEFAULT_EXCLUDE_KW)
    if os.path.isfile(KW_FILE):
        section = "delete"
        with open(KW_FILE) as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    if "exclude" in line.lower(): section = "exclude"
                    elif "auto" in line.lower(): section = "auto"
                    elif "delete" in line.lower(): section = "delete"
                    continue
                if section == "delete": delete_kw.append(line)
                elif section == "auto": auto_kw.append(line)
                else: exclude_kw.append(line)
        delete_kw = list(dict.fromkeys(delete_kw))
        auto_kw = list(dict.fromkeys(auto_kw))
        exclude_kw = list(dict.fromkeys(exclude_kw))
    return delete_kw, auto_kw, exclude_kw


def save_keywords(delete_kw, auto_kw, exclude_kw):
    with open(KW_FILE, "w") as f:
        f.write("# Delete keywords (ask user)\n")
        for kw in delete_kw: f.write(kw + "\n")
        f.write("\n# Auto-delete keywords (no confirmation)\n")
        for kw in auto_kw: f.write(kw + "\n")
        f.write("\n# Exclude keywords\n")
        for kw in exclude_kw: f.write(kw + "\n")


class GUIWriter:
    def __init__(self, text_widget, root):
        self.text = text_widget; self.root = root
    def write(self, msg):
        if msg: self.root.after(0, self._append, msg)
    def _append(self, msg):
        self.text.config(state="normal")
        self.text.insert("end", msg)
        self.text.see("end")
        self.text.config(state="disabled")
    def flush(self): pass


def gui_input(root, prompt):
    result = [None]; event = threading.Event()
    def _ask():
        dialog = tk.Toplevel(root); dialog.title("Confirm")
        dialog.transient(root); dialog.grab_set(); dialog.resizable(False, False)
        tk.Label(dialog, text=prompt, padx=20, pady=10, wraplength=500).pack()
        btn_frame = tk.Frame(dialog, padx=10, pady=10); btn_frame.pack()
        def _yes(): result[0]="y"; dialog.destroy(); event.set()
        def _no(): result[0]="n"; dialog.destroy(); event.set()
        def _quit(): result[0]="q"; dialog.destroy(); event.set()
        tk.Button(btn_frame, text="Yes", command=_yes, width=10, bg="#4CAF50", fg="white").pack(side="left", padx=5)
        tk.Button(btn_frame, text="No", command=_no, width=10).pack(side="left", padx=5)
        tk.Button(btn_frame, text="Quit", command=_quit, width=10, bg="#f44336", fg="white").pack(side="left", padx=5)
        dialog.update_idletasks()
        x = root.winfo_x() + (root.winfo_width() - dialog.winfo_width()) // 2
        y = root.winfo_y() + (root.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{x}+{y}")
        dialog.protocol("WM_DELETE_WINDOW", _quit)
    root.after(0, _ask); event.wait(); return result[0]


def gui_text_input(root, prompt, options=None):
    """Show a dialog with clickable option buttons for exclude keywords."""
    result = [None]; event = threading.Event()
    def _ask():
        dialog = tk.Toplevel(root); dialog.title("Add Exclude Keyword")
        dialog.transient(root); dialog.grab_set(); dialog.resizable(False, False)
        tk.Label(dialog, text=prompt, padx=20, pady=10, wraplength=500).pack()

        if options:
            tk.Label(dialog, text="Click a keyword to exclude:").pack()
            bf = tk.Frame(dialog, padx=10, pady=5); bf.pack()
            for opt in options:
                def _pick(o=opt): result[0]=o; dialog.destroy(); event.set()
                tk.Button(bf, text=opt, command=_pick, width=15).pack(side="left", padx=3, pady=3)

        ef = tk.Frame(dialog, padx=10, pady=5); ef.pack()
        tk.Label(ef, text="Or type custom:").pack(side="left")
        entry = tk.Entry(ef, width=20); entry.pack(side="left", padx=5)
        def _ok(): result[0]=entry.get().strip(); dialog.destroy(); event.set()
        tk.Button(ef, text="OK", command=_ok, width=8).pack(side="left")

        def _skip(): result[0]=""; dialog.destroy(); event.set()
        tk.Button(dialog, text="Skip", command=_skip, width=10).pack(pady=5)

        dialog.update_idletasks()
        x = root.winfo_x() + (root.winfo_width() - dialog.winfo_width()) // 2
        y = root.winfo_y() + (root.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{x}+{y}")
        dialog.protocol("WM_DELETE_WINDOW", _skip)
    root.after(0, _ask); event.wait(); return result[0]



class CleanupApp:
    def __init__(self, root):
        self.root = root
        root.title("CST Component Cleanup")
        root.geometry("800x600")

        # Project path
        pf = tk.Frame(root, padx=10, pady=5); pf.pack(fill="x")
        tk.Label(pf, text="CST Project:").pack(side="left")
        self.project_path = tk.StringVar(value="")
        tk.Entry(pf, textvariable=self.project_path, width=50).pack(side="left", padx=5)
        tk.Button(pf, text="Browse", command=self._browse).pack(side="left")

        # Keywords frame
        kf = tk.Frame(root, padx=10, pady=5); kf.pack(fill="x")

        # Delete keywords
        tk.Label(kf, text="Delete keywords:").grid(row=0, column=0, sticky="w")
        self.delete_kw_var = tk.StringVar()
        self.delete_kw_entry = tk.Entry(kf, textvariable=self.delete_kw_var, width=40)
        self.delete_kw_entry.grid(row=0, column=1, padx=5)

        # Exclude keywords
        tk.Label(kf, text="Exclude keywords:").grid(row=1, column=0, sticky="w")
        self.exclude_kw_var = tk.StringVar()
        self.exclude_kw_entry = tk.Entry(kf, textvariable=self.exclude_kw_var, width=40)
        self.exclude_kw_entry.grid(row=1, column=1, padx=5)

        # Auto-delete keywords
        tk.Label(kf, text="Auto-delete (no ask):").grid(row=2, column=0, sticky="w")
        self.auto_kw_var = tk.StringVar()
        self.auto_kw_entry = tk.Entry(kf, textvariable=self.auto_kw_var, width=40)
        self.auto_kw_entry.grid(row=2, column=1, padx=5)

        # Load saved keywords
        dk, ak, ek = load_keywords()
        self.delete_kw_var.set(", ".join(dk))
        self.auto_kw_var.set(", ".join(ak))
        self.exclude_kw_var.set(", ".join(ek))

        # Buttons
        bf = tk.Frame(root, padx=10, pady=5); bf.pack(fill="x")
        self.btn_scan = tk.Button(bf, text="Scan & Clean", command=self._run, width=20, height=1)
        self.btn_scan.pack(side="left", padx=5)
        self.btn_save = tk.Button(bf, text="Save Keywords", command=self._save_kw, width=15)
        self.btn_save.pack(side="left", padx=5)
        self.btn_export = tk.Button(bf, text="Export Keywords", command=self._export_kw, width=15)
        self.btn_export.pack(side="left", padx=5)
        self.btn_import = tk.Button(bf, text="Import Keywords", command=self._import_kw, width=15)
        self.btn_import.pack(side="left", padx=5)

        # Log
        lf = tk.Frame(root, padx=10, pady=5); lf.pack(fill="both", expand=True)
        sb = tk.Scrollbar(lf); sb.pack(side="right", fill="y")
        self.log_text = tk.Text(lf, state="disabled", wrap="word",
                                yscrollcommand=sb.set, font=("Consolas", 9))
        self.log_text.pack(fill="both", expand=True)
        sb.config(command=self.log_text.yview)

        # Status
        self.status = tk.StringVar(value="Ready.")
        tk.Label(root, textvariable=self.status, padx=10, pady=3,
                 anchor="w", fg="gray", relief="sunken").pack(fill="x")

        self.gui_writer = GUIWriter(self.log_text, root)

    def _browse(self):
        p = filedialog.askopenfilename(title="Select CST Project",
            filetypes=[("CST files", "*.cst"), ("All", "*.*")])
        if p: self.project_path.set(p)

    def _save_kw(self):
        dk = [k.strip() for k in self.delete_kw_var.get().split(",") if k.strip()]
        ak = [k.strip() for k in self.auto_kw_var.get().split(",") if k.strip()]
        ek = [k.strip() for k in self.exclude_kw_var.get().split(",") if k.strip()]
        save_keywords(dk, ak, ek)
        messagebox.showinfo("Saved", f"Keywords saved to {KW_FILE}")

    def _export_kw(self):
        dk = [k.strip() for k in self.delete_kw_var.get().split(",") if k.strip()]
        ak = [k.strip() for k in self.auto_kw_var.get().split(",") if k.strip()]
        ek = [k.strip() for k in self.exclude_kw_var.get().split(",") if k.strip()]
        path = filedialog.asksaveasfilename(
            title="Export Keywords",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All", "*.*")],
            initialfile="cleanup_keywords.txt")
        if path:
            with open(path, "w") as f:
                f.write("# Delete keywords (ask user)\n")
                for kw in dk: f.write(kw + "\n")
                f.write("\n# Auto-delete keywords (no confirmation)\n")
                for kw in ak: f.write(kw + "\n")
                f.write("\n# Exclude keywords\n")
                for kw in ek: f.write(kw + "\n")
            messagebox.showinfo("Exported", f"Keywords exported to {path}")

    def _import_kw(self):
        path = filedialog.askopenfilename(
            title="Import Keywords",
            filetypes=[("Text files", "*.txt"), ("All", "*.*")])
        if not path: return
        dk = []; ak = []; ek = []; section = "delete"
        with open(path) as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    if "exclude" in line.lower(): section = "exclude"
                    elif "auto" in line.lower(): section = "auto"
                    elif "delete" in line.lower(): section = "delete"
                    continue
                if section == "delete": dk.append(line)
                elif section == "auto": ak.append(line)
                else: ek.append(line)
        if dk: self.delete_kw_var.set(", ".join(dk))
        if ak: self.auto_kw_var.set(", ".join(ak))
        if ek: self.exclude_kw_var.set(", ".join(ek))
        messagebox.showinfo("Imported", f"Loaded {len(dk)} delete + {len(ak)} auto + {len(ek)} exclude keywords")

    def _clear_log(self):
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")

    def _run(self):
        project = self.project_path.get().strip()
        if not project or not os.path.isfile(project):
            messagebox.showerror("Error", "Select a valid CST project."); return
        self._clear_log()
        self.btn_scan.config(state="disabled")
        self.status.set("Running cleanup...")
        t = threading.Thread(target=self._run_thread, args=(project,), daemon=True)
        t.start()

    def _run_thread(self, project):
        old_stdout = sys.stdout; sys.stdout = self.gui_writer
        import builtins; old_input = builtins.input
        builtins.input = lambda prompt="": gui_input(self.root, prompt)

        try:
            self._do_cleanup(project)
            self.root.after(0, lambda: self.status.set("Done."))
        except Exception as exc:
            print(f"\nERROR: {exc}"); print(traceback.format_exc())
            self.root.after(0, lambda: self.status.set(f"Error: {exc}"))
        finally:
            sys.stdout = old_stdout; builtins.input = old_input
            self.root.after(0, lambda: self.btn_scan.config(state="normal"))

    def _do_cleanup(self, project):
        from code.cst_connection import CSTConnection
        from code.feature_detector import FeatureDetector, SATParser

        # Parse keywords from GUI
        delete_kw = [k.strip().upper() for k in self.delete_kw_var.get().split(",") if k.strip()]
        auto_kw = [k.strip().upper() for k in self.auto_kw_var.get().split(",") if k.strip()]
        exclude_kw = [k.strip().upper() for k in self.exclude_kw_var.get().split(",") if k.strip()]

        conn = CSTConnection(); conn.connect(); conn.open_project(project)
        print(f"Opened: {project}")
        det = FeatureDetector(conn)

        try:
            # List components recursively
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

            components = []
            if result:
                for line in result.split("\n"):
                    line = line.strip()
                    if not line: continue
                    path = line.replace("\\", "/")
                    if path.startswith("Components/"): path = path[len("Components/"):]
                    parts = path.split("/")
                    if len(parts) >= 2:
                        solid = parts[-1]; comp_path = "/".join(parts[:-1])
                        components.append({"comp": comp_path, "solid": solid,
                            "shape": f"{comp_path}:{solid}", "raw": line, "parts": parts})

            print(f"Total components: {len(components)}")

            # Filter — match only against the SOLID NAME, not the full path
            def matches(comp_dict):
                name = comp_dict["solid"].upper()  # only check solid name
                for ek in exclude_kw:
                    if ek in name: return None, None
                for ak in auto_kw:
                    if ak in name: return ak, "auto"
                for dk in delete_kw:
                    if dk in name: return dk, "ask"
                return None, None

            matched = []
            for c in components:
                kw, mode = matches(c)
                if kw:
                    c["keyword"] = kw
                    c["mode"] = mode  # "auto" or "ask"
                    matched.append(c)

            print(f"Matched: {len(matched)} (delete: {delete_kw}, auto: {auto_kw}, exclude: {exclude_kw})")

            # Group similar — improved: if xxx exists alongside xxx_1, xxx_2, group them all
            groups = {}
            for c in matched:
                parent = c["comp"]
                solid = c["solid"]
                # Strip trailing _N to get base name
                m = re.match(r'^(.+?)_(\d+)$', solid)
                base = m.group(1) if m else solid
                key = f"{parent}:{base}"
                if key not in groups:
                    groups[key] = {"parent": parent, "base": base, "items": [], "keyword": c["keyword"]}
                groups[key]["items"].append(c)

            # Second pass: merge groups where one base is a prefix of another
            # e.g. base="SCREW_ST17_L5" and base="SCREW_ST17_L5_10" should merge
            # because SCREW_ST17_L5_10 is the original and _1,_2 are copies
            keys = list(groups.keys())
            merged = set()
            for k1 in keys:
                if k1 in merged: continue
                g1 = groups[k1]
                for k2 in keys:
                    if k2 == k1 or k2 in merged: continue
                    g2 = groups[k2]
                    if g1["parent"] != g2["parent"]: continue
                    # Check if g1.base == g2.base + "_N" pattern
                    # i.e. g2 is the parent of g1's items
                    if g1["base"] == g2["base"]:
                        continue  # same group
                    # If g1.base starts with g2.base + "_" → merge g1 into g2
                    if g1["base"].startswith(g2["base"] + "_"):
                        g2["items"].extend(g1["items"])
                        merged.add(k1)
                        break
                    # If g2.base starts with g1.base + "_" → merge g2 into g1
                    if g2["base"].startswith(g1["base"] + "_"):
                        g1["items"].extend(g2["items"])
                        merged.add(k2)

            for k in merged:
                del groups[k]

            groups = list(groups.values())

            # Sort by Zmax (get bbox for first item in each group)
            for g in groups:
                first = g["items"][0]; p = first["shape"].split(":")
                g["zmax"] = -float('inf')
                try:
                    sat = det._export_sat(p[0], p[1])
                    if sat:
                        sp = SATParser(sat); sp.parse(); bbs = sp.get_bounding_boxes()
                        if bbs:
                            g["zmax"] = max(bb[5] for bb in bbs.values())
                except: pass
            groups.sort(key=lambda g: g["zmax"], reverse=True)

            print(f"Groups: {len(groups)} (sorted by Zmax)")

            # Process each group
            deleted = 0
            for gi, group in enumerate(groups):
                items = group["items"]; names = [it["solid"] for it in items]
                print(f"\n[{gi+1}/{len(groups)}] [{group['keyword']}] {group['base']} ({len(items)} items)")

                # Check if all items in group are auto-delete
                all_auto = all(it.get("mode") == "auto" for it in items)

                if all_auto:
                    # Auto-delete without asking
                    print(f"  AUTO-DELETE ({group['keyword']})")
                    action = "y"
                else:
                    # SelectTreeItem to highlight
                    parent_tree_path = items[0]["raw"].rsplit("\\", 1)[0] if "\\" in items[0]["raw"] else items[0]["raw"]
                    conn.execute_vba(
                        'Sub Main\n'
                        f'  SelectTreeItem("{parent_tree_path}")\n'
                        '  Plot.ZoomToStructure\nEnd Sub\n')

                    if len(items) == 1:
                        prompt = f"Delete '{items[0]['solid']}'?"
                    else:
                        prompt = f"Delete all {len(items)} items ({', '.join(names[:5])}{'...' if len(names)>5 else ''})?"

                    action = input(prompt)

                if action == "q":
                    print("Quit."); break
                elif action == "y":
                    for item in items:
                        shape_name = item["shape"]
                        # Step 1: Delete via RunScript (silent, no popup)
                        r = conn.execute_vba(
                            'Sub Main\n'
                            f'  Open "{_out_vba}" For Output As #1\n  On Error Resume Next\n'
                            f'  Solid.Delete "{shape_name}"\n'
                            '  If Err.Number <> 0 Then\n    Print #1, "FAIL: " & Err.Description\n    Err.Clear\n'
                            '  Else\n    Print #1, "OK"\n  End If\n  Close #1\nEnd Sub\n',
                            output_file=_out_path)
                        ok = "OK" in (r or "")
                        if ok:
                            deleted += 1
                            # Step 2: Record in history (comment only, no re-execution)
                            try:
                                conn.execute_vba(
                                    'Sub Main\n  On Error Resume Next\n'
                                    f'  AddToHistory "deleted: {item["solid"]}", "\'shape was: {shape_name}"\n'
                                    'End Sub\n')
                            except: pass
                        print(f"  {item['solid']}: {'OK' if ok else 'FAILED'}")

                    # After deleting all items, clean up empty parent folders
                    # Check from the deepest parent upward
                    parent_raw = items[0]["raw"].replace("\\", "/")
                    # Remove "Components/" prefix and solid name to get parent path
                    if parent_raw.startswith("Components/"):
                        parent_raw = parent_raw[len("Components/"):]
                    parent_parts = parent_raw.split("/")[:-1]  # remove solid name

                    # Walk up the tree, deleting empty folders
                    while parent_parts:
                        folder_path = "Components\\" + "\\".join(parent_parts)
                        # Check if folder has any children
                        check_r = conn.execute_vba(
                            'Sub Main\n'
                            f'  Open "{_out_vba}" For Output As #1\n'
                            '  Dim rt As Object\n  Set rt = Resulttree\n'
                            f'  Dim child As String\n  child = rt.GetFirstChildName("{folder_path}")\n'
                            '  If child = "" Then\n    Print #1, "EMPTY"\n'
                            '  Else\n    Print #1, "HAS_CHILDREN"\n  End If\n'
                            '  Close #1\nEnd Sub\n',
                            output_file=_out_path)
                        if "EMPTY" in (check_r or ""):
                            # Delete empty component folder
                            comp_name = "/".join(parent_parts)
                            del_r = conn.execute_vba(
                                'Sub Main\n'
                                f'  Open "{_out_vba}" For Output As #1\n  On Error Resume Next\n'
                                f'  Component.Delete "{comp_name}"\n'
                                '  If Err.Number <> 0 Then\n    Print #1, "COMP_DEL_FAIL"\n    Err.Clear\n'
                                '  Else\n    Print #1, "COMP_DEL_OK"\n  End If\n  Close #1\nEnd Sub\n',
                                output_file=_out_path)
                            if "COMP_DEL_OK" in (del_r or ""):
                                print(f"  Deleted empty folder: {comp_name}")
                            parent_parts.pop()  # go up one level
                        else:
                            break  # folder not empty, stop
                else:
                    print("  Skipped.")
                    # Ask for exclude keyword
                    name_parts = re.split(r'[_\-]', group["base"])
                    name_parts = [p for p in name_parts if len(p) > 2]
                    if name_parts:
                        excl = gui_text_input(self.root,
                            f"Skipped '{group['base']}'. Add exclude keyword?",
                            options=name_parts)
                        if excl:
                            excl_upper = excl.upper()
                            exclude_kw.append(excl_upper)
                            print(f"  Added exclude: {excl_upper}")
                            # Update GUI
                            self.root.after(0, lambda: self.exclude_kw_var.set(", ".join(exclude_kw)))
                            # Re-filter remaining groups
                            new_groups = []
                            for rg in groups[gi+1:]:
                                skip = False
                                for item in rg["items"]:
                                    if excl_upper in item["shape"].upper():
                                        skip = True; break
                                if not skip: new_groups.append(rg)
                            removed = len(groups[gi+1:]) - len(new_groups)
                            if removed > 0:
                                print(f"  Excluded {removed} more groups")
                                groups[gi+1:] = new_groups

            print(f"\nTotal deleted: {deleted}")

            # Ask to save
            save = input("Save keywords for future use?")
            if save == "y":
                save_keywords(delete_kw, auto_kw, exclude_kw)
                print(f"Saved to {KW_FILE}")

        finally:
            try: conn.execute_vba('Sub Main\n  Pick.ClearAllPicks\n  WCS.ActivateWCS "global"\nEnd Sub\n')
            except: pass
            conn.close()


def main():
    root = tk.Tk()
    app = CleanupApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
