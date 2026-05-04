"""Standalone runner for CMA Simulation Setup.

Automates the full CMA configuration workflow on a CST model:
1. Assign PEC material to all components
2. Set simulation frequency range
3. Create E-field and H-field monitors
4. Generate mesh and verify
5. Set boundary conditions to open (add space)
6. Configure Integral Equation Solver for CMA

Default test model: cst_simplifier/cst_model/EM_Sunray_v2.cst

Run:
  python -m code.run_cma_setup
  python -m code.run_cma_setup "path/to/model.cst"
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Can be set externally (e.g. by GUI) before calling main()
PROJECT = None

# Default test model
DEFAULT_MODEL = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "cst_simplifier", "cst_model", "EM_Sunray_v2.cst",
)


def _get_project_path() -> str:
    """Resolve the .cst model path from PROJECT variable, CLI arg, or prompt."""
    global PROJECT

    if PROJECT:
        path = PROJECT
    elif len(sys.argv) > 1:
        path = sys.argv[1].strip().strip('"').strip("'")
    else:
        print("=" * 60)
        print("  CMA Simulation Setup")
        print("=" * 60)
        default_hint = ""
        if os.path.isfile(DEFAULT_MODEL):
            default_hint = f" [default: {DEFAULT_MODEL}]"
        path = input(f"\n  Enter path to .cst model{default_hint}: ").strip().strip('"').strip("'")

        if not path and os.path.isfile(DEFAULT_MODEL):
            path = DEFAULT_MODEL

    if not path:
        print("  Error: no path provided.")
        sys.exit(1)

    path = os.path.abspath(path)

    if not path.lower().endswith(".cst"):
        print(f"  Error: expected a .cst file, got: {path}")
        sys.exit(1)

    if not os.path.exists(path):
        print(f"  Error: file not found: {path}")
        sys.exit(1)

    return path


def main():
    project_path = _get_project_path()

    import code.cma_setup as cma_mod
    cma_mod.PROJECT = project_path
    cma_mod.main()


if __name__ == "__main__":
    main()
