"""Eigenmode Simulation Setup for CST Studio Suite 2025.

Automates the creation of per-component eigenmode simulation projects from
shield can components in a CST model:

1. Select RF technologies to determine frequency range
2. Identify shield can components (cover, frame, one-piece)
3. Confirm components with user via CST GUI selection
4. For each confirmed component:
   - Export as SAB from source project
   - Create new CST project
   - Import SAB geometry
   - Assign PEC material to all solids
   - Set electric boundary conditions (all 6 faces)
   - Configure frequency range
   - Switch to eigenmode solver (30 modes, CPU acceleration)
   - Generate mesh
   - Save project

All CST modifications use AddToHistory for persistence where appropriate.
Direct VBA is used for ChangeSolverType, Mesh.Update, and Solver parallelization.

Validates: Requirements 1–15
"""

import logging
import os
import re
import tempfile
import time
from dataclasses import dataclass, field
from typing import List, Tuple, Optional

logger = logging.getLogger(__name__)

# Module-level project path — set by GUI before calling main()
PROJECT = None


# ---------------------------------------------------------------------------
# Exception
# ---------------------------------------------------------------------------

class UserQuitException(Exception):
    """Raised when user enters 'q' or clicks Quit in a GUI dialog."""
    pass


# ---------------------------------------------------------------------------
# Data models
# ---------------------------------------------------------------------------

@dataclass
class WorkflowStep:
    """Result of a single workflow step."""

    name: str       # e.g. "SAB Export", "Electric Boundaries"
    success: bool   # True if step completed without error
    message: str    # Human-readable result description
    critical: bool  # If True and failed, component is skipped


@dataclass
class RFTechnology:
    """An RF technology with its frequency band."""

    name: str           # e.g. "WiFi 2G", "WiFi 5G", "BLE/BT", "LoRa"
    fmin_ghz: float     # Lower frequency bound in GHz
    fmax_ghz: float     # Upper frequency bound in GHz


@dataclass
class FrequencyRange:
    """Combined simulation frequency range."""

    fmin_ghz: float         # Minimum frequency (min of all selected lower bounds)
    fmax_ghz: float         # Maximum frequency (max of all selected upper bounds)
    technologies: List[str]  # Names of selected technologies


@dataclass
class ShieldCanComponent:
    """A shield can component identified for eigenmode simulation."""

    comp_path: str      # Component path in CST tree (e.g. "SHIELDING/COVER")
    solid_name: str     # Solid name (e.g. "SHIELDING_COVER_1")
    shape_name: str     # Full shape reference "comp_path:solid_name"
    category: str       # "cover", "frame", or "one_piece"


@dataclass
class ComponentResult:
    """Result of processing one shield can component."""

    component: ShieldCanComponent
    steps: List[WorkflowStep]
    success: bool       # True if all critical steps passed
    project_path: str   # Path to saved eigenmode project (empty if failed)


@dataclass
class EigenmodeWorkflowSummary:
    """Summary of the complete eigenmode setup workflow."""

    frequency_range: FrequencyRange
    total_components: int
    results: List[ComponentResult]
    total_time: float

    @property
    def successful(self) -> int:
        return sum(1 for r in self.results if r.success)

    @property
    def failed(self) -> int:
        return sum(1 for r in self.results if not r.success)


# ---------------------------------------------------------------------------
# RF Technology lookup table
# ---------------------------------------------------------------------------

RF_TECHNOLOGIES = {
    "WiFi 2G": RFTechnology("WiFi 2G", 1.5, 7.5),
    "WiFi 5G": RFTechnology("WiFi 5G", 4.0, 15.0),
    "BLE/BT":  RFTechnology("BLE/BT", 1.5, 7.5),
    "LoRa":    RFTechnology("LoRa", 0.7, 3.0),
}


# ---------------------------------------------------------------------------
# Input validation helpers
# ---------------------------------------------------------------------------

def prompt_frequency(prompt_text: str, default: float = None) -> float:
    """Prompt user for a frequency value in GHz.

    Validates numeric input, re-prompts on invalid. Uses *default* when the
    user provides empty input and a default is set. Raises
    ``UserQuitException`` on ``'q'`` input.
    """
    while True:
        raw = input(prompt_text).strip()

        if raw.lower() == "q":
            raise UserQuitException("User quit.")

        if raw == "" and default is not None:
            return default

        try:
            value = float(raw)
        except ValueError:
            print("Invalid input. Please enter a number.")
            continue

        if value <= 0:
            print("Frequency must be positive.")
            continue

        return value


def prompt_integer(prompt_text: str, default: int = None, min_val: int = 1) -> int:
    """Prompt user for an integer value."""
    while True:
        raw = input(prompt_text).strip()

        if raw.lower() == "q":
            raise UserQuitException("User quit.")

        if raw == "" and default is not None:
            return default

        try:
            value = int(raw)
        except ValueError:
            print("Please enter a whole number.")
            continue

        if value < min_val:
            print(f"Value must be at least {min_val}.")
            continue

        return value


def prompt_confirm(prompt_text: str) -> bool:
    """Prompt user for y/n confirmation.

    Returns True for 'y'/'Y', False for 'n'/'N'.
    Raises UserQuitException on 'q' input.
    """
    while True:
        raw = input(prompt_text).strip()

        if raw.lower() == "q":
            raise UserQuitException("User quit.")

        if raw.lower() == "y":
            return True

        if raw.lower() == "n":
            return False

        print("Please enter 'y' or 'n'.")


# ---------------------------------------------------------------------------
# RF Technology selection
# ---------------------------------------------------------------------------

def compute_combined_frequency_range(technologies: List[RFTechnology]) -> Tuple[float, float]:
    """Compute combined frequency range from a list of RF technologies.

    Returns (fmin, fmax) where fmin = min of all lower bounds and
    fmax = max of all upper bounds.
    """
    fmin = min(t.fmin_ghz for t in technologies)
    fmax = max(t.fmax_ghz for t in technologies)
    return fmin, fmax


def select_rf_technologies() -> FrequencyRange:
    """Multi-select RF technologies and compute combined frequency range.

    Presents a numbered list of technologies plus "Other" for custom entry.
    User enters comma-separated numbers. Computes combined range as
    min(lower bounds), max(upper bounds).

    Returns:
        FrequencyRange with computed min/max and list of technology names.
    """
    tech_list = list(RF_TECHNOLOGIES.values())

    while True:
        print("\n  Select RF technologies (comma-separated numbers):")
        for i, tech in enumerate(tech_list, 1):
            print(f"    {i}. {tech.name} ({tech.fmin_ghz}–{tech.fmax_ghz} GHz)")
        print(f"    {len(tech_list) + 1}. Other (custom frequency range)")

        raw = input("\n  Enter selection (e.g. 1,2): ").strip()

        if raw.lower() == "q":
            raise UserQuitException("User quit.")

        if not raw:
            print("  Please select at least one technology.")
            continue

        # Parse selections
        try:
            selections = [int(x.strip()) for x in raw.split(",")]
        except ValueError:
            print("  Invalid input. Enter numbers separated by commas.")
            continue

        selected_techs: List[RFTechnology] = []
        tech_names: List[str] = []
        needs_custom = False

        for sel in selections:
            if sel < 1 or sel > len(tech_list) + 1:
                print(f"  Invalid selection: {sel}")
                break
            if sel == len(tech_list) + 1:
                needs_custom = True
            else:
                tech = tech_list[sel - 1]
                if tech.name not in tech_names:
                    selected_techs.append(tech)
                    tech_names.append(tech.name)
        else:
            # All selections valid
            if needs_custom:
                print("\n  Custom frequency range:")
                custom_fmin = prompt_frequency("    Enter minimum frequency (GHz): ")
                custom_fmax = prompt_frequency("    Enter maximum frequency (GHz): ")
                if custom_fmin >= custom_fmax:
                    print("  Minimum must be less than maximum. Try again.")
                    continue
                selected_techs.append(RFTechnology("Other", custom_fmin, custom_fmax))
                tech_names.append("Other")

            if not selected_techs:
                print("  Please select at least one technology.")
                continue

            # Compute combined range
            fmin, fmax = compute_combined_frequency_range(selected_techs)

            print(f"\n  Combined frequency range: {fmin}–{fmax} GHz")
            print(f"  Technologies: {', '.join(tech_names)}")

            if prompt_confirm("  Confirm? (y/n): "):
                return FrequencyRange(fmin_ghz=fmin, fmax_ghz=fmax, technologies=tech_names)
            # else loop again

        # If we broke out of the for loop (invalid selection), continue outer loop
        continue


# ---------------------------------------------------------------------------
# Shield can identification
# ---------------------------------------------------------------------------

def classify_shield_can(comp_path: str, solid_name: str) -> Optional[str]:
    """Classify a single component as a shield can category.

    Returns "cover", "frame", "one_piece", or None if not a shield can.
    """
    full_path = f"{comp_path}/{solid_name}".upper()

    if "SHIELD" not in full_path:
        return None

    name_upper = solid_name.upper()
    if "COVER" in name_upper:
        return "cover"
    elif "FRAM" in name_upper:
        return "frame"
    else:
        return "one_piece"


def identify_shield_cans(conn) -> List[ShieldCanComponent]:
    """Enumerate all solids and identify shield can components.

    Uses FeatureDetector._enumerate_solids() to get all (comp, solid) pairs,
    then classifies by keywords.

    Args:
        conn: A connected CSTConnection instance.

    Returns:
        List of ShieldCanComponent objects.
    """
    from code.feature_detector import FeatureDetector

    detector = FeatureDetector(conn)
    solids = detector._enumerate_solids()

    components: List[ShieldCanComponent] = []
    for comp_path, solid_name in solids:
        category = classify_shield_can(comp_path, solid_name)
        if category is not None:
            shape_name = f"{comp_path}:{solid_name}"
            components.append(ShieldCanComponent(
                comp_path=comp_path,
                solid_name=solid_name,
                shape_name=shape_name,
                category=category,
            ))

    print(f"  Found {len(components)} shield can components:")
    covers = sum(1 for c in components if c.category == "cover")
    frames = sum(1 for c in components if c.category == "frame")
    one_piece = sum(1 for c in components if c.category == "one_piece")
    print(f"    Covers: {covers}, Frames: {frames}, One-piece: {one_piece}")

    return components


# ---------------------------------------------------------------------------
# Component confirmation
# ---------------------------------------------------------------------------

def confirm_components(conn, components: List[ShieldCanComponent]) -> List[ShieldCanComponent]:
    """Confirm shield can components with user via CST GUI selection.

    For each component, selects it in the CST GUI and prompts the user
    to include (y), skip (n), or quit (q).

    Args:
        conn: A connected CSTConnection instance.
        components: List of ShieldCanComponent to confirm.

    Returns:
        List of confirmed ShieldCanComponent objects.
    """
    if not components:
        print("  No shield can components to confirm.")
        return []

    confirmed: List[ShieldCanComponent] = []
    total = len(components)

    print(f"\n  Confirming {total} shield can components...")
    print("  Each component will be selected in CST for visual verification.")
    print("  Enter 'y' to include, 'n' to skip, 'q' to stop.\n")

    for i, comp in enumerate(components, 1):
        # Select in CST GUI
        out_path = os.path.join(tempfile.gettempdir(), "cst_eigen_select.txt")
        try:
            macro = build_select_component_vba(comp.shape_name, out_path)
            conn.execute_vba(macro, output_file=out_path)
        except Exception as exc:
            logger.warning("Could not select '%s' in CST: %s", comp.shape_name, exc)
        finally:
            try:
                os.remove(out_path)
            except OSError:
                pass

        print(f"  Component {i}/{total}: {comp.solid_name} [{comp.category}]")
        print(f"    Path: {comp.comp_path}")

        raw = input(f"    Include this component? (y/n/q): ").strip().lower()

        if raw == "q":
            print("  Stopping confirmation loop.")
            break
        elif raw == "y":
            confirmed.append(comp)
            print(f"    → Included")
        else:
            print(f"    → Skipped")

    print(f"\n  Confirmed {len(confirmed)} of {total} components for processing.")
    return confirmed


# ---------------------------------------------------------------------------
# VBA macro builders
# ---------------------------------------------------------------------------

def _escape_vba_path(path: str) -> str:
    """Escape a file path for embedding in a VBA string literal."""
    return path.replace("\\", "\\\\")


def _sanitize_filename(name: str) -> str:
    """Sanitize a component name for use in filenames.

    Replaces path separators and invalid filename characters with underscores.
    """
    # Replace common path separators and invalid chars
    sanitized = re.sub(r'[\\/:*?"<>|]', '_', name)
    # Replace multiple underscores with single
    sanitized = re.sub(r'_+', '_', sanitized)
    # Strip leading/trailing underscores
    sanitized = sanitized.strip('_')
    return sanitized if sanitized else "component"


def get_sab_path(solid_name: str, source_dir: str) -> str:
    """Generate the SAB export file path for a component."""
    sanitized = _sanitize_filename(solid_name)
    return os.path.join(source_dir, f"{sanitized}.sab")


def get_eigenmode_project_path(solid_name: str, source_dir: str) -> str:
    """Generate the eigenmode project file path for a component."""
    sanitized = _sanitize_filename(solid_name)
    return os.path.join(source_dir, f"Eigenmode_{sanitized}.cst")


def build_select_component_vba(shape_name: str, output_path: str) -> str:
    """Build VBA macro to select a component in CST tree for visual confirmation.

    The shape_name is in "comp_path:solid_name" format where comp_path may
    contain forward slashes. The CST tree uses backslash separators:
    "Components\\comp\\subcomp\\solid".
    """
    esc_path = _escape_vba_path(output_path)
    # Convert "comp_path:solid_name" to "Components\comp_path\solid_name"
    # with forward slashes in comp_path converted to backslashes
    parts = shape_name.split(":")
    if len(parts) == 2:
        comp_path = parts[0].replace("/", "\\")
        tree_path = f"Components\\{comp_path}\\{parts[1]}"
    else:
        tree_path = f"Components\\{shape_name.replace('/', chr(92))}"

    return (
        "Sub Main\n"
        "  Dim outPath As String\n"
        f'  outPath = "{esc_path}"\n'
        "  Open outPath For Output As #1\n"
        "  On Error Resume Next\n"
        f'  SelectTreeItem("{tree_path}")\n'
        "  If Err.Number <> 0 Then\n"
        '    Print #1, "FAIL: " & Err.Description\n'
        "    Err.Clear\n"
        "  Else\n"
        '    Print #1, "OK"\n'
        "  End If\n"
        "  Close #1\n"
        "End Sub\n"
    )


def build_sab_export_vba(shape_name: str, sab_path: str, output_path: str) -> str:
    """Build VBA macro to export a single component as .sab file.

    Uses CST SAT object: SAT.Reset + .FileName + .Write "shape_name"
    """
    esc_path = _escape_vba_path(output_path)
    esc_sab = _escape_vba_path(sab_path)

    return (
        "Sub Main\n"
        "  Dim outPath As String\n"
        f'  outPath = "{esc_path}"\n'
        "  Open outPath For Output As #1\n"
        "  On Error Resume Next\n"
        "  With SAT\n"
        "    .Reset\n"
        f'    .FileName "{esc_sab}"\n'
        f'    .Write "{shape_name}"\n'
        "  End With\n"
        "  If Err.Number <> 0 Then\n"
        '    Print #1, "FAIL: " & Err.Description\n'
        "    Err.Clear\n"
        "  Else\n"
        '    Print #1, "OK"\n'
        "  End If\n"
        "  Close #1\n"
        "End Sub\n"
    )


def build_sab_import_vba(sab_path: str, output_path: str) -> str:
    """Build VBA macro to import a .sab file into the current project.

    Uses CST SAT object: SAT.Reset + .FileName + .Id "1" + .Read
    """
    esc_path = _escape_vba_path(output_path)
    esc_sab = _escape_vba_path(sab_path)

    return (
        "Sub Main\n"
        "  Dim outPath As String\n"
        f'  outPath = "{esc_path}"\n'
        "  Open outPath For Output As #1\n"
        "  On Error Resume Next\n"
        "  With SAT\n"
        "    .Reset\n"
        f'    .FileName "{esc_sab}"\n'
        '    .Id "1"\n'
        "    .Read\n"
        "  End With\n"
        "  If Err.Number <> 0 Then\n"
        '    Print #1, "FAIL: " & Err.Description\n'
        "    Err.Clear\n"
        "  Else\n"
        '    Print #1, "OK"\n'
        "  End If\n"
        "  Close #1\n"
        "End Sub\n"
    )


def build_pec_material_vba(shape_name: str, output_path: str) -> str:
    """Build VBA macro to assign PEC material to a shape via AddToHistory.

    The shape_name must be in "component:solid" format.
    """
    esc_path = _escape_vba_path(output_path)
    return (
        "Sub Main\n"
        "  Dim outPath As String\n"
        f'  outPath = "{esc_path}"\n'
        "  Open outPath For Output As #1\n"
        "  On Error Resume Next\n"
        f'  AddToHistory "assign PEC: {shape_name}", '
        f'"Solid.ChangeMaterial ""{shape_name}"", ""PEC"""\n'
        "  If Err.Number <> 0 Then\n"
        '    Print #1, "FAIL: " & Err.Description\n'
        "    Err.Clear\n"
        "  Else\n"
        '    Print #1, "OK"\n'
        "  End If\n"
        "  Close #1\n"
        "End Sub\n"
    )


def build_electric_boundary_vba(output_path: str) -> str:
    """Build VBA macro to set all 6 boundary faces to "electric" via AddToHistory.

    Sets Xmin, Xmax, Ymin, Ymax, Zmin, Zmax all to "electric" for a
    closed metallic cavity simulation.
    """
    esc_path = _escape_vba_path(output_path)
    return (
        "Sub Main\n"
        "  Dim outPath As String\n"
        f'  outPath = "{esc_path}"\n'
        "  Open outPath For Output As #1\n"
        "  On Error Resume Next\n"
        "  Dim sCode As String\n"
        '  sCode = "With Boundary" & vbCrLf\n'
        '  sCode = sCode & "  .Xmin ""electric""" & vbCrLf\n'
        '  sCode = sCode & "  .Xmax ""electric""" & vbCrLf\n'
        '  sCode = sCode & "  .Ymin ""electric""" & vbCrLf\n'
        '  sCode = sCode & "  .Ymax ""electric""" & vbCrLf\n'
        '  sCode = sCode & "  .Zmin ""electric""" & vbCrLf\n'
        '  sCode = sCode & "  .Zmax ""electric""" & vbCrLf\n'
        '  sCode = sCode & "End With"\n'
        '  AddToHistory "set boundary conditions", sCode\n'
        "  If Err.Number <> 0 Then\n"
        '    Print #1, "FAIL: " & Err.Description\n'
        "    Err.Clear\n"
        "  Else\n"
        '    Print #1, "OK"\n'
        "  End If\n"
        "  Close #1\n"
        "End Sub\n"
    )


def build_frequency_range_vba(fmin_ghz: float, fmax_ghz: float, output_path: str) -> str:
    """Build VBA macro to set the simulation frequency range via AddToHistory."""
    esc_path = _escape_vba_path(output_path)
    return (
        "Sub Main\n"
        "  Dim outPath As String\n"
        f'  outPath = "{esc_path}"\n'
        "  Open outPath For Output As #1\n"
        "  On Error Resume Next\n"
        f'  AddToHistory "set frequency range", '
        f'"Solver.FrequencyRange ""{fmin_ghz}"", ""{fmax_ghz}"""\n'
        "  If Err.Number <> 0 Then\n"
        '    Print #1, "FAIL: " & Err.Description\n'
        "    Err.Clear\n"
        "  Else\n"
        '    Print #1, "OK"\n'
        "  End If\n"
        "  Close #1\n"
        "End Sub\n"
    )


def build_eigenmode_solver_vba(num_modes: int, freq_target_ghz: float, output_path: str) -> str:
    """Build VBA macro to switch to eigenmode solver and configure it.

    - ChangeSolverType "HF Eigenmode" as direct VBA (NOT AddToHistory)
    - EigenmodeSolver configuration via AddToHistory
    - Solver.UseParallelization + MaximumNumberOfCPUDevices as direct VBA
    """
    esc_path = _escape_vba_path(output_path)
    return (
        "Sub Main\n"
        "  Dim outPath As String\n"
        f'  outPath = "{esc_path}"\n'
        "  Open outPath For Output As #1\n"
        "  On Error Resume Next\n"
        # -- Switch to eigenmode solver (direct, NOT AddToHistory) --
        '  ChangeSolverType "HF Eigenmode"\n'
        "  If Err.Number <> 0 Then\n"
        '    Print #1, "FAIL: ChangeSolverType: " & Err.Description\n'
        "    Err.Clear\n"
        "    Close #1\n"
        "    Exit Sub\n"
        "  End If\n"
        # -- EigenmodeSolver configuration via AddToHistory --
        "  Dim sCode As String\n"
        '  sCode = "With EigenmodeSolver" & vbCrLf\n'
        f'  sCode = sCode & "  .SetNumberOfModes ""{num_modes}""" & vbCrLf\n'
        f'  sCode = sCode & "  .SetFrequencyTarget ""True"", ""{freq_target_ghz}""" & vbCrLf\n'
        '  sCode = sCode & "  .SetUseParallelization ""True""" & vbCrLf\n'
        '  sCode = sCode & "  .SetMaxNumberOfThreads ""1024""" & vbCrLf\n'
        '  sCode = sCode & "  .MaximumNumberOfCPUDevices ""8""" & vbCrLf\n'
        '  sCode = sCode & "End With"\n'
        '  AddToHistory "configure eigenmode solver", sCode\n'
        "  If Err.Number <> 0 Then\n"
        '    Print #1, "FAIL: EigenmodeSolver config: " & Err.Description\n'
        "    Err.Clear\n"
        "  Else\n"
        '    Print #1, "OK"\n'
        "  End If\n"
        "  Close #1\n"
        "End Sub\n"
    )


def build_mesh_vba(output_path: str) -> str:
    """Build VBA macro to generate mesh for eigenmode solver.

    The eigenmode solver uses a hexahedral mesh. We set the mesh creator,
    update the mesh, and switch to mesh view mode.
    """
    esc_path = _escape_vba_path(output_path)
    return (
        "Sub Main\n"
        "  Dim outPath As String\n"
        f'  outPath = "{esc_path}"\n'
        "  Open outPath For Output As #1\n"
        "  On Error Resume Next\n"
        '  Mesh.SetCreator "High Frequency"\n'
        "  Mesh.Update\n"
        "  If Err.Number <> 0 Then\n"
        '    Print #1, "FAIL: Mesh.Update: " & Err.Description\n'
        "    Err.Clear\n"
        "  Else\n"
        '    Mesh.ViewMeshMode "True"\n'
        "    If Err.Number <> 0 Then\n"
        "      Err.Clear\n"
        "    End If\n"
        '    Print #1, "OK"\n'
        "  End If\n"
        "  Close #1\n"
        "End Sub\n"
    )


def build_save_project_vba(save_path: str, output_path: str) -> str:
    """Build VBA macro to save the project with SaveAs."""
    esc_path = _escape_vba_path(output_path)
    esc_save = _escape_vba_path(save_path)
    return (
        "Sub Main\n"
        "  Dim outPath As String\n"
        f'  outPath = "{esc_path}"\n'
        "  Open outPath For Output As #1\n"
        "  On Error Resume Next\n"
        f'  SaveAs "{esc_save}", "False"\n'
        "  If Err.Number <> 0 Then\n"
        '    Print #1, "FAIL: " & Err.Description\n'
        "    Err.Clear\n"
        "  Else\n"
        '    Print #1, "OK"\n'
        "  End If\n"
        "  Close #1\n"
        "End Sub\n"
    )


# ---------------------------------------------------------------------------
# Per-component processing
# ---------------------------------------------------------------------------

def _execute_vba_step(conn, macro: str, output_path: str, step_name: str) -> Tuple[bool, str]:
    """Execute a VBA macro and check the output file for OK/FAIL.

    Returns:
        (success, message) tuple.
    """
    try:
        result = conn.execute_vba(macro, output_file=output_path)
        if result and result.startswith("FAIL:"):
            return False, f"{step_name}: {result}"
        return True, f"{step_name}: OK"
    except Exception as exc:
        return False, f"{step_name}: Exception — {exc}"
    finally:
        try:
            os.remove(output_path)
        except OSError:
            pass


def process_component(conn, app, component: ShieldCanComponent,
                      freq_range: FrequencyRange,
                      source_dir: str, source_path: str) -> ComponentResult:
    """Execute the full per-component eigenmode workflow.

    Steps:
    1. SAB export from source project (critical)
    2. Create new project via app.NewMWS() (critical)
    3. SAB import into new project (critical)
    4. PEC material assignment for all solids (non-critical)
    5. Electric boundary conditions (non-critical)
    6. Frequency range configuration (non-critical)
    7. Solver type change + eigenmode config (non-critical)
    8. Mesh generation (non-critical)
    9. Project save (non-critical)

    After processing, switches back to the source project.
    """
    from code.cst_connection import call_method, CSTConnectionError
    from code.feature_detector import FeatureDetector

    steps: List[WorkflowStep] = []
    sab_path = get_sab_path(component.solid_name, source_dir)
    project_path = get_eigenmode_project_path(component.solid_name, source_dir)
    new_project = None

    print(f"\n    Processing: {component.solid_name} [{component.category}]")

    # --- Step 1: SAB Export (critical) ---
    out_path = os.path.join(tempfile.gettempdir(), "cst_eigen_export.txt")
    macro = build_sab_export_vba(component.shape_name, sab_path, out_path)
    success, msg = _execute_vba_step(conn, macro, out_path, "SAB Export")
    steps.append(WorkflowStep("SAB Export", success, msg, critical=True))

    if not success:
        print(f"      ✗ SAB Export FAILED — skipping component")
        logger.warning("SAB export failed for '%s': %s", component.shape_name, msg)
        return ComponentResult(component, steps, success=False, project_path="")

    # Verify SAB file exists
    if not os.path.isfile(sab_path):
        steps[-1] = WorkflowStep("SAB Export", False, "SAB file not created", critical=True)
        print(f"      ✗ SAB file not created — skipping component")
        return ComponentResult(component, steps, success=False, project_path="")

    print(f"      ✓ SAB exported: {os.path.basename(sab_path)}")

    # --- Step 2: Create new project (critical) ---
    try:
        call_method(app, "NewMWS")
        new_project = app.Active3D()
        if new_project is None:
            raise CSTConnectionError("Active3D() returned None after NewMWS()")
        # Switch connection to new project
        old_project = conn._project
        conn._project = new_project
        steps.append(WorkflowStep("New Project", True, "Created new MWS project", critical=True))
        print(f"      ✓ New project created")
    except Exception as exc:
        steps.append(WorkflowStep("New Project", False, f"Failed: {exc}", critical=True))
        print(f"      ✗ New project creation FAILED — skipping component")
        logger.warning("New project creation failed: %s", exc)
        # Clean up SAB file
        try:
            os.remove(sab_path)
        except OSError:
            pass
        return ComponentResult(component, steps, success=False, project_path="")

    # --- Step 3: SAB Import (critical) ---
    out_path = os.path.join(tempfile.gettempdir(), "cst_eigen_import.txt")
    macro = build_sab_import_vba(sab_path, out_path)
    success, msg = _execute_vba_step(conn, macro, out_path, "SAB Import")
    steps.append(WorkflowStep("SAB Import", success, msg, critical=True))

    if not success:
        print(f"      ✗ SAB Import FAILED — skipping component")
        logger.warning("SAB import failed for '%s': %s", component.shape_name, msg)
        # Switch back to source project
        try:
            call_method(app, "OpenFile", source_path)
            conn._project = app.Active3D()
        except Exception:
            conn._project = old_project
        # Clean up SAB file
        try:
            os.remove(sab_path)
        except OSError:
            pass
        return ComponentResult(component, steps, success=False, project_path="")

    print(f"      ✓ SAB imported")

    # --- Step 4: PEC Material Assignment (non-critical) ---
    # Enumerate solids in the new project and assign PEC to all
    try:
        detector = FeatureDetector(conn)
        new_solids = detector._enumerate_solids()
        pec_ok = True
        for comp_path, solid_name in new_solids:
            shape = f"{comp_path}:{solid_name}"
            out_path = os.path.join(tempfile.gettempdir(), "cst_eigen_pec.txt")
            macro = build_pec_material_vba(shape, out_path)
            s, m = _execute_vba_step(conn, macro, out_path, f"PEC: {shape}")
            if not s:
                pec_ok = False
                logger.warning("PEC assignment failed for '%s': %s", shape, m)
        if pec_ok and new_solids:
            steps.append(WorkflowStep("PEC Material", True,
                                      f"Assigned PEC to {len(new_solids)} solids", critical=False))
            print(f"      ✓ PEC material assigned to {len(new_solids)} solids")
        elif not new_solids:
            steps.append(WorkflowStep("PEC Material", True,
                                      "No solids found (SAB may use default)", critical=False))
            print(f"      ⚠ No solids found in new project")
        else:
            steps.append(WorkflowStep("PEC Material", False,
                                      "Some PEC assignments failed", critical=False))
            print(f"      ⚠ Some PEC assignments failed")
    except Exception as exc:
        steps.append(WorkflowStep("PEC Material", False, str(exc), critical=False))
        print(f"      ⚠ PEC material: {exc}")

    # --- Step 5: Electric Boundary Conditions (non-critical) ---
    out_path = os.path.join(tempfile.gettempdir(), "cst_eigen_boundary.txt")
    macro = build_electric_boundary_vba(out_path)
    success, msg = _execute_vba_step(conn, macro, out_path, "Electric Boundaries")
    steps.append(WorkflowStep("Electric Boundaries", success, msg, critical=False))
    if success:
        print(f"      ✓ Electric boundaries set (all 6 faces)")
    else:
        print(f"      ⚠ Boundary config: {msg}")

    # --- Step 6: Frequency Range (non-critical) ---
    out_path = os.path.join(tempfile.gettempdir(), "cst_eigen_freq.txt")
    macro = build_frequency_range_vba(freq_range.fmin_ghz, freq_range.fmax_ghz, out_path)
    success, msg = _execute_vba_step(conn, macro, out_path, "Frequency Range")
    steps.append(WorkflowStep("Frequency Range", success, msg, critical=False))
    if success:
        print(f"      ✓ Frequency range: {freq_range.fmin_ghz}–{freq_range.fmax_ghz} GHz")
    else:
        print(f"      ⚠ Frequency range: {msg}")

    # --- Step 7: Eigenmode Solver (non-critical) ---
    # Use fmin as frequency target
    out_path = os.path.join(tempfile.gettempdir(), "cst_eigen_solver.txt")
    macro = build_eigenmode_solver_vba(30, freq_range.fmin_ghz, out_path)
    success, msg = _execute_vba_step(conn, macro, out_path, "Eigenmode Solver")
    steps.append(WorkflowStep("Eigenmode Solver", success, msg, critical=False))
    if success:
        print(f"      ✓ Eigenmode solver: 30 modes, target {freq_range.fmin_ghz} GHz, 8 CPUs")
    else:
        print(f"      ⚠ Eigenmode solver: {msg}")

    # --- Step 8: Mesh Generation (non-critical) ---
    out_path = os.path.join(tempfile.gettempdir(), "cst_eigen_mesh.txt")
    macro = build_mesh_vba(out_path)
    success, msg = _execute_vba_step(conn, macro, out_path, "Mesh Generation")
    steps.append(WorkflowStep("Mesh Generation", success, msg, critical=False))
    if success:
        print(f"      ✓ Mesh generated")
    else:
        print(f"      ⚠ Mesh generation: {msg}")

    # --- Step 9: Save Project (non-critical) ---
    out_path = os.path.join(tempfile.gettempdir(), "cst_eigen_save.txt")
    macro = build_save_project_vba(project_path, out_path)
    success, msg = _execute_vba_step(conn, macro, out_path, "Save Project")
    steps.append(WorkflowStep("Save Project", success, msg, critical=False))
    if success:
        print(f"      ✓ Saved: {os.path.basename(project_path)}")
    else:
        print(f"      ⚠ Save failed: {msg}")

    # --- Cleanup: switch back to source project ---
    try:
        call_method(app, "OpenFile", source_path)
        conn._project = app.Active3D()
    except Exception as exc:
        logger.warning("Failed to switch back to source project: %s", exc)
        conn._project = old_project

    # Clean up SAB file
    try:
        os.remove(sab_path)
    except OSError:
        pass

    # Determine overall success (all critical steps passed)
    all_critical_ok = all(s.success for s in steps if s.critical)
    return ComponentResult(component, steps, success=all_critical_ok, project_path=project_path)


# ---------------------------------------------------------------------------
# EigenmodeSetup class — workflow orchestration
# ---------------------------------------------------------------------------

class EigenmodeSetup:
    """Orchestrates the eigenmode simulation setup workflow."""

    def __init__(self, connection):
        """Initialize with a live CST connection.

        Args:
            connection: A connected CSTConnection instance with an open project.
        """
        self._conn = connection
        self._app = connection.app

    def run(self) -> EigenmodeWorkflowSummary:
        """Execute the complete eigenmode setup workflow.

        Phases:
        1. RF technology selection → FrequencyRange
        2. Shield can identification → component list
        3. Component confirmation → confirmed list
        4. Per-component processing loop

        Returns:
            EigenmodeWorkflowSummary with all results.
        """
        start = time.time()

        # Phase 1: RF Technology Selection
        print("\n--- Phase 1: RF Technology Selection ---")
        freq_range = select_rf_technologies()

        # Phase 2: Shield Can Identification
        print("\n--- Phase 2: Shield Can Identification ---")
        components = identify_shield_cans(self._conn)

        if not components:
            print("  No shield can components found in the model.")
            elapsed = time.time() - start
            return EigenmodeWorkflowSummary(
                frequency_range=freq_range,
                total_components=0,
                results=[],
                total_time=elapsed,
            )

        # Phase 3: Component Confirmation
        print("\n--- Phase 3: Component Confirmation ---")
        confirmed = confirm_components(self._conn, components)

        if not confirmed:
            print("  No components confirmed for processing.")
            elapsed = time.time() - start
            return EigenmodeWorkflowSummary(
                frequency_range=freq_range,
                total_components=0,
                results=[],
                total_time=elapsed,
            )

        # Phase 4: Per-Component Processing
        print("\n--- Phase 4: Per-Component Processing ---")
        print(f"  Processing {len(confirmed)} components...")

        # Determine source directory (same dir as source project)
        source_path = os.path.abspath(PROJECT or "")
        source_dir = os.path.dirname(source_path)

        results: List[ComponentResult] = []
        for i, comp in enumerate(confirmed, 1):
            print(f"\n  [{i}/{len(confirmed)}] {comp.solid_name}")
            result = process_component(
                self._conn, self._app, comp, freq_range, source_dir, source_path,
            )
            results.append(result)

            # Print per-component summary
            status = "✓ SUCCESS" if result.success else "✗ FAILED"
            print(f"    Result: {status}")

        # Final summary
        elapsed = time.time() - start
        summary = EigenmodeWorkflowSummary(
            frequency_range=freq_range,
            total_components=len(confirmed),
            results=results,
            total_time=elapsed,
        )
        self._print_summary(summary)
        return summary

    @staticmethod
    def _print_summary(summary: EigenmodeWorkflowSummary) -> None:
        """Print a human-readable workflow summary."""
        print("\n" + "=" * 60)
        print("  Eigenmode Setup Summary")
        print("=" * 60)
        print(f"\n  Frequency range: {summary.frequency_range.fmin_ghz}–"
              f"{summary.frequency_range.fmax_ghz} GHz")
        print(f"  Technologies: {', '.join(summary.frequency_range.technologies)}")
        print(f"\n  Components processed: {summary.total_components}")
        print(f"  Successful: {summary.successful}")
        print(f"  Failed: {summary.failed}")

        if summary.results:
            print("\n  Per-component results:")
            for r in summary.results:
                status = "✓" if r.success else "✗"
                print(f"    {status} {r.component.solid_name} [{r.component.category}]")
                if r.success and r.project_path:
                    print(f"      → {os.path.basename(r.project_path)}")
                for step in r.steps:
                    if not step.success:
                        print(f"        ⚠ {step.name}: {step.message}")

        print(f"\n  Total time: {summary.total_time:.1f}s")
        print("=" * 60)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main(project_path: str = None):
    """Entry point for eigenmode setup workflow.

    Connects to CST, opens the project, runs the full eigenmode setup
    workflow, and prints a summary.

    Args:
        project_path: Path to the .cst project file. Falls back to
            the module-level PROJECT variable, then prompts the user.
    """
    from code.cst_connection import CSTConnection, CSTConnectionError

    global PROJECT

    # Resolve project path
    path = project_path or PROJECT
    if not path:
        path = input("  Enter path to .cst model: ").strip().strip('"').strip("'")
    if not path:
        print("  Error: no project path provided.")
        return

    path = os.path.abspath(path)
    PROJECT = path

    print("=" * 60)
    print("  Eigenmode Simulation Setup")
    print("=" * 60)
    print(f"  Project: {path}")

    conn = CSTConnection()
    try:
        print("\n  Connecting to CST...")
        conn.connect()
        conn.open_project(path)
        print(f"  Project opened: {os.path.basename(path)}")

        setup = EigenmodeSetup(conn)
        setup.run()

    except CSTConnectionError as exc:
        print(f"\n  CRITICAL ERROR: {exc}")
        logger.error("CST connection error: %s", exc)
        return
    except UserQuitException:
        print("\n  Workflow stopped by user.")
        return
    finally:
        conn.close()
        print("\n  CST connection closed.")


if __name__ == "__main__":
    main()
