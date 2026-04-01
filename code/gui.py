"""Simple GUI launcher for CST CAD Simplifier tools.

Three buttons:
1. Fill holes on PCB board
2. Simplify dimples/holes on shield can (cover + frame)
3. Check connection between shield can cover and frame

Output log shown in GUI text area. User input via popup dialogs.

Run: python -m code.gui
"""

import os
import sys
import io
import threading
import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


class GUIWriter:
    """Redirects print() output to a tkinter Text widget."""

    def __init__(self, text_widget, root):
        self.text = text_widget
        self.root = root

    def write(self, msg):
        if msg:
            self.root.after(0, self._append, msg)

    def _append(self, msg):
        self.text.config(state="normal")
        self.text.insert("end", msg)
        self.text.see("end")
        self.text.config(state="disabled")

    def flush(self):
        pass


def gui_input(root, prompt):
    """Replacement for input() that shows a dialog with Yes/No/Quit buttons."""
    result = [None]
    event = threading.Event()

    def _ask():
        dialog = tk.Toplevel(root)
        dialog.title("Confirm")
        dialog.transient(root)
        dialog.grab_set()
        dialog.resizable(False, False)

        tk.Label(dialog, text=prompt, padx=20, pady=10, wraplength=400).pack()

        btn_frame = tk.Frame(dialog, padx=10, pady=10)
        btn_frame.pack()

        def _yes():
            result[0] = "y"
            dialog.destroy()
            event.set()

        def _no():
            result[0] = "n"
            dialog.destroy()
            event.set()

        def _quit():
            result[0] = "q"
            dialog.destroy()
            event.set()

        tk.Button(btn_frame, text="Yes", command=_yes, width=10, bg="#4CAF50", fg="white").pack(side="left", padx=5)
        tk.Button(btn_frame, text="No", command=_no, width=10).pack(side="left", padx=5)
        tk.Button(btn_frame, text="Quit", command=_quit, width=10, bg="#f44336", fg="white").pack(side="left", padx=5)

        # Center dialog on parent
        dialog.update_idletasks()
        x = root.winfo_x() + (root.winfo_width() - dialog.winfo_width()) // 2
        y = root.winfo_y() + (root.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{x}+{y}")

        dialog.protocol("WM_DELETE_WINDOW", _quit)

    root.after(0, _ask)
    event.wait()
    return result[0]


class App:
    def __init__(self, root):
        self.root = root
        root.title("CST CAD Simplifier")
        root.geometry("700x500")

        # Project path
        self.project_path = tk.StringVar(value="")

        path_frame = tk.Frame(root, padx=10, pady=5)
        path_frame.pack(fill="x")
        tk.Label(path_frame, text="CST Project:").pack(side="left")
        tk.Entry(path_frame, textvariable=self.project_path, width=50).pack(side="left", padx=5)
        tk.Button(path_frame, text="Browse", command=self._browse).pack(side="left")

        # Buttons
        btn_frame = tk.Frame(root, padx=10, pady=5)
        btn_frame.pack(fill="x")

        self.btn1 = tk.Button(
            btn_frame, text="1. Fill Holes on PCB Board",
            command=lambda: self._run_tool("pcb"), width=30, height=1,
        )
        self.btn1.pack(side="left", padx=3)

        self.btn2 = tk.Button(
            btn_frame, text="2. Simplify Shield Can",
            command=lambda: self._run_tool("shieldcan"), width=30, height=1,
        )
        self.btn2.pack(side="left", padx=3)

        self.btn3 = tk.Button(
            btn_frame, text="3. Bridge Cover-Frame Gap",
            command=lambda: self._run_tool("bridge"), width=30, height=1,
        )
        self.btn3.pack(side="left", padx=3)

        # Log area
        log_frame = tk.Frame(root, padx=10, pady=5)
        log_frame.pack(fill="both", expand=True)

        scrollbar = tk.Scrollbar(log_frame)
        scrollbar.pack(side="right", fill="y")

        self.log_text = tk.Text(log_frame, state="disabled", wrap="word",
                                yscrollcommand=scrollbar.set, font=("Consolas", 9))
        self.log_text.pack(fill="both", expand=True)
        scrollbar.config(command=self.log_text.yview)

        # Status bar
        self.status = tk.StringVar(value="Ready. Select a CST project file to begin.")
        tk.Label(root, textvariable=self.status, padx=10, pady=3,
                 anchor="w", fg="gray", relief="sunken").pack(fill="x")

        # Redirect stdout
        self.gui_writer = GUIWriter(self.log_text, root)

    def _browse(self):
        path = filedialog.askopenfilename(
            title="Select CST Project",
            filetypes=[("CST files", "*.cst"), ("All files", "*.*")],
        )
        if path:
            self.project_path.set(path)

    def _clear_log(self):
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")

    def _set_buttons_state(self, state):
        self.btn1.config(state=state)
        self.btn2.config(state=state)
        self.btn3.config(state=state)

    def _run_tool(self, tool):
        project = self.project_path.get().strip()
        if not project or not os.path.isfile(project):
            messagebox.showerror("Error", "Please select a valid CST project file.")
            return

        self._clear_log()
        self._set_buttons_state("disabled")
        self.status.set(f"Running {tool}...")

        thread = threading.Thread(target=self._run_in_thread, args=(tool, project), daemon=True)
        thread.start()

    def _run_in_thread(self, tool, project):
        # Redirect stdout to GUI
        old_stdout = sys.stdout
        sys.stdout = self.gui_writer

        # Monkey-patch builtins.input to use GUI dialog
        import builtins
        old_input = builtins.input
        builtins.input = lambda prompt="": gui_input(self.root, prompt)

        try:
            if tool == "pcb":
                self._run_pcb(project)
            elif tool == "shieldcan":
                self._run_shieldcan(project)
            elif tool == "bridge":
                self._run_bridge(project)
            self.root.after(0, lambda: self.status.set("Done. Ready for next tool."))
        except Exception as exc:
            print(f"\nERROR: {exc}")
            import traceback
            print(traceback.format_exc())
            self.root.after(0, lambda: self.status.set(f"Error: {exc}"))
        finally:
            sys.stdout = old_stdout
            builtins.input = old_input
            self.root.after(0, lambda: self._set_buttons_state("normal"))

    def _connect(self, project):
        from code.cst_connection import CSTConnection
        conn = CSTConnection()
        conn.connect()
        conn.open_project(project)
        return conn

    def _run_pcb(self, project):
        """Run PCB board hole filler."""
        from code.feature_detector import FeatureDetector
        from code.simplifier import Simplifier

        conn = self._connect(project)
        try:
            det = FeatureDetector(conn)
            solid_data = det.detect_seeds()
            simplifier = Simplifier(conn)

            found = False
            for data in solid_data:
                name = data["shape_name"].upper()
                if "BOARD" not in name:
                    continue
                found = True
                print(f"\nProcessing PCB: {data['shape_name']}")
                simplifier.fill_progressive(
                    data["shape_name"], data["seeds"], data["adjacency"],
                    data["bboxes"], data["face_types"],
                    interactive=True, seed_groups=data.get("seed_groups", []),
                )

            if not found:
                print("No PCB board component found (name must contain 'BOARD').")
        finally:
            try: conn.execute_vba('Sub Main\n  WCS.ActivateWCS "global"\n  Pick.ClearAllPicks\nEnd Sub\n')
            except: pass
            conn.close()

    def _run_shieldcan(self, project):
        """Run shield can simplifier (cover + frame)."""
        from code.feature_detector import FeatureDetector, SATParser
        from code.wall_detector import WallDetector, WallInfo, _dot, _normalize
        from code.simplifier import Simplifier

        conn = self._connect(project)
        try:
            det = FeatureDetector(conn)
            solid_data = det.detect_seeds()
            simplifier = Simplifier(conn)

            for data in solid_data:
                name = data["shape_name"].upper()
                if "COVER" in name:
                    mode = "cover"
                elif "FRAM" in name:
                    mode = "frame"
                else:
                    continue

                shape_name = data["shape_name"]
                print(f"\nProcessing [{mode.upper()}]: {shape_name}")

                parts = shape_name.split(":")
                sat = det._export_sat(parts[0], parts[1])
                parser = SATParser(sat)
                face_data = parser.parse()
                adjacency = parser.build_adjacency()
                bboxes = parser.get_bounding_boxes()

                wd = WallDetector()
                ref_pid, ref_n = wd.find_top_face(face_data, bboxes)

                if mode == "cover":
                    walls = wd.discover_side_walls(ref_pid, ref_n, face_data, adjacency, bboxes)
                else:
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

                def _wall_area(w):
                    b = w.bbox
                    dims = sorted([b[3]-b[0], b[4]-b[1], b[5]-b[2]], reverse=True)
                    return dims[0] * dims[1]
                walls.sort(key=_wall_area)

                consumed = set()
                for wi, wall in enumerate(walls):
                    dimples = wd.find_dimple_faces(wall, face_data, adjacency, bboxes, ref_pid, walls)
                    dimples = [d for d in dimples if d not in consumed]
                    if not dimples: continue

                    wn = wall.normal; bb = wall.bbox
                    cx=(bb[0]+bb[3])/2; cy=(bb[1]+bb[4])/2; cz=(bb[2]+bb[5])/2
                    try:
                        conn.execute_vba(
                            'Sub Main\n'
                            f'  WCS.SetOrigin {cx}, {cy}, {cz}\n'
                            f'  WCS.SetNormal {wn[0]}, {wn[1]}, {wn[2]}\n'
                            '  WCS.ActivateWCS "local"\nEnd Sub\n')
                    except: pass
                    try: simplifier._highlight_faces(shape_name, dimples)
                    except: pass

                    print(f"  Wall {wi+1}/{len(walls)}: {len(dimples)} dimples")
                    action = input(f"Remove {len(dimples)} faces? (y/n/q): ").strip().lower()
                    if action == "q": break
                    if action != "y":
                        try: conn.execute_vba('Sub Main\n  Pick.ClearAllPicks\nEnd Sub\n')
                        except: pass
                        continue

                    ok, msg = simplifier._try_fill_hole_silent(shape_name, dimples, wi+1)
                    if ok: consumed.update(dimples)
                    print(f"    {'OK' if ok else 'FAILED'}")

        finally:
            try: conn.execute_vba('Sub Main\n  WCS.ActivateWCS "global"\n  Pick.ClearAllPicks\nEnd Sub\n')
            except: pass
            conn.close()

    def _run_bridge(self, project):
        """Run shield can cover-frame bridge."""
        import code.debug_contact_v17_shieldcan as bridge_mod
        bridge_mod.PROJECT = project
        bridge_mod.main()


def main():
    root = tk.Tk()
    app = App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
