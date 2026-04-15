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


class UserQuitException(Exception):
    """Raised when user clicks Quit in a dialog."""
    pass


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
    """Replacement for input() — shows text entry for name prompts, Yes/No/Quit for confirmations."""
    # Detect if this is a text input prompt (contains "Enter" or "name" or "component")
    is_text_prompt = any(kw in prompt.lower() for kw in ["enter ", "type ", "provide "])

    result = [None]
    event = threading.Event()

    def _ask():
        dialog = tk.Toplevel(root)
        dialog.title("Input" if is_text_prompt else "Confirm")
        dialog.transient(root)
        dialog.grab_set()
        dialog.resizable(False, False)

        tk.Label(dialog, text=prompt, padx=20, pady=10, wraplength=400).pack()

        if is_text_prompt:
            entry = tk.Entry(dialog, width=40)
            entry.pack(padx=20, pady=5)
            entry.focus_set()

            def _ok():
                result[0] = entry.get().strip()
                dialog.destroy()
                event.set()

            def _skip():
                result[0] = ""
                dialog.destroy()
                event.set()

            bf = tk.Frame(dialog, padx=10, pady=10); bf.pack()
            tk.Button(bf, text="OK", command=_ok, width=10, bg="#4CAF50", fg="white").pack(side="left", padx=5)
            tk.Button(bf, text="Skip", command=_skip, width=10).pack(side="left", padx=5)
            entry.bind("<Return>", lambda e: _ok())
        else:
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

        dialog.update_idletasks()
        x = root.winfo_x() + (root.winfo_width() - dialog.winfo_width()) // 2
        y = root.winfo_y() + (root.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{x}+{y}")
        dialog.protocol("WM_DELETE_WINDOW", lambda: (_quit() if not is_text_prompt else _skip()))

    root.after(0, _ask)
    event.wait()
    if result[0] == "q":
        raise UserQuitException("User clicked Quit")
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

        # Second row of buttons
        btn_frame2 = tk.Frame(root, padx=10, pady=2)
        btn_frame2.pack(fill="x")

        self.btn4 = tk.Button(
            btn_frame2, text="4. Bridge Grounding for PCB",
            command=lambda: self._run_tool("pcb_bridge"), width=30, height=1,
        )
        self.btn4.pack(side="left", padx=3)

        self.btn5 = tk.Button(
            btn_frame2, text="5. Replace Connector",
            command=lambda: self._run_tool("connector"), width=30, height=1,
        )
        self.btn5.pack(side="left", padx=3)

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

    def _save_log(self, log_path):
        """Save the current log text to a file."""
        try:
            self.log_text.config(state="normal")
            content = self.log_text.get("1.0", "end").strip()
            self.log_text.config(state="disabled")
            if content:
                with open(log_path, "w", encoding="utf-8") as f:
                    f.write(content)
                print(f"\nLog saved: {log_path}")
        except Exception as exc:
            print(f"\nFailed to save log: {exc}")

    def _set_buttons_state(self, state):
        self.btn1.config(state=state)
        self.btn2.config(state=state)
        self.btn3.config(state=state)
        self.btn4.config(state=state)
        self.btn5.config(state=state)

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
            elif tool == "pcb_bridge":
                self._run_pcb_bridge(project)
            elif tool == "connector":
                self._run_connector(project)
            self.root.after(0, lambda: self.status.set("Done. Ready for next tool."))
        except UserQuitException:
            print("\nUser quit.")
            self.root.after(0, lambda: self.status.set("Stopped by user. Ready for next tool."))
        except RuntimeError as exc:
            if "User quit" in str(exc):
                print("\nUser quit.")
                self.root.after(0, lambda: self.status.set("Stopped by user. Ready for next tool."))
            else:
                print(f"\nERROR: {exc}")
                import traceback
                print(traceback.format_exc())
                self.root.after(0, lambda: self.status.set(f"Error: {exc}"))
        except Exception as exc:
            print(f"\nERROR: {exc}")
            import traceback
            print(traceback.format_exc())
            self.root.after(0, lambda: self.status.set(f"Error: {exc}"))
        finally:
            # Save log file next to the CST project
            try:
                from datetime import datetime
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                log_dir = os.path.dirname(project)
                log_path = os.path.join(log_dir, f"cst_simplifier_{tool}_{timestamp}.log")
                self.root.after(0, self._save_log, log_path)
            except Exception:
                pass
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
        import importlib
        import code.run_sunray_v6 as pcb_mod
        importlib.reload(pcb_mod)
        pcb_mod.PROJECT = project
        pcb_mod.main()

    def _run_shieldcan(self, project):
        """Run shield can simplifier (cover + frame)."""
        import importlib
        import builtins
        import code.run_shieldcan as sc_mod
        importlib.reload(sc_mod)
        sc_mod.PROJECT = project
        # Pass root widget so the dialog can be created on the main thread
        builtins._shield_can_gui_root = self.root
        try:
            sc_mod.main()
        finally:
            if hasattr(builtins, '_shield_can_gui_root'):
                del builtins._shield_can_gui_root

    def _run_bridge(self, project):
        """Run shield can cover-frame bridge."""
        import code.debug_contact_v17_shieldcan as bridge_mod
        bridge_mod.PROJECT = project
        bridge_mod.main()

    def _run_pcb_bridge(self, project):
        """Run PCB grounding bridge."""
        import code.debug_pcb_edge_v2 as pcb_mod
        pcb_mod.PROJECT = project
        pcb_mod.main()

    def _run_connector(self, project):
        """Run connector replacement."""
        import importlib
        import code.debug_connector_v2 as conn_mod
        importlib.reload(conn_mod)
        conn_mod.PROJECT = project
        conn_mod.main()


def main():
    root = tk.Tk()
    app = App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
