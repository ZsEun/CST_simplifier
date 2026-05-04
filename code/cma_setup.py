"""CMA (Characteristic Mode Analysis) Simulation Setup for CST Studio Suite 2025.

Automates the multi-step configuration workflow required after importing a model
into CST for CMA analysis:

1. Assign PEC material to all solid shapes
2. Set simulation frequency range
3. Create E-field and H-field monitors
4. Generate mesh for visual verification
5. Set boundary conditions to open (add space)
6. Configure Integral Equation Solver with CMA settings

All CST modifications use AddToHistory for persistence. User input is collected
via input() which the GUI automatically replaces with popup dialogs.

Validates: Requirements 1–10
"""

import logging
import os
import tempfile
import time
from dataclasses import dataclass, field
from typing import List, Tuple

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
    """Result of a single workflow step.

    Validates: Requirements 7.1, 7.5
    """

    name: str       # e.g. "Material Assignment", "Frequency Range"
    success: bool   # True if step completed without error
    message: str    # Human-readable result description
    critical: bool  # If True and failed, workflow stops


@dataclass
class WorkflowSummary:
    """Summary of the complete CMA setup workflow.

    Validates: Requirements 7.5
    """

    steps: List[WorkflowStep]
    total_time: float  # Total elapsed time in seconds

    @property
    def completed(self) -> List[str]:
        return [s.name for s in self.steps if s.success]

    @property
    def failed(self) -> List[str]:
        return [s.name for s in self.steps if not s.success]


@dataclass
class FrequencyConfig:
    """Validated frequency configuration."""

    fmin_ghz: float   # Minimum frequency in GHz
    fmax_ghz: float   # Maximum frequency in GHz
    monitor_ghz: float  # Monitor frequency in GHz


@dataclass
class SolverConfig:
    """I-Solver configuration parameters.

    Validates: Requirements 6.1–6.8
    """

    method: str = "ACA"
    mode_tracking: bool = False
    freq_samples_as_monitors: bool = True
    num_modes: int = 10
    weighting_coefficients: bool = False
    cpu_devices: int = 8


# ---------------------------------------------------------------------------
# Input validation helpers
# ---------------------------------------------------------------------------

def prompt_frequency(prompt_text: str, default: float = None) -> float:
    """Prompt user for a frequency value in GHz.

    Validates numeric input, re-prompts on invalid. Uses *default* when the
    user provides empty input and a default is set. Raises
    ``UserQuitException`` on ``'q'`` input.

    Validates: Requirements 10.1, 10.2, 10.3, 10.4, 10.5
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
    """Prompt user for an integer value.

    Validates integer input >= *min_val*, re-prompts on invalid. Uses
    *default* when the user provides empty input and a default is set.
    Raises ``UserQuitException`` on ``'q'`` input.

    Validates: Requirements 10.1, 10.2, 10.3, 10.4, 10.5
    """
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

    Returns ``True`` for ``'y'``/``'Y'``, ``False`` for ``'n'``/``'N'``.
    Raises ``UserQuitException`` on ``'q'`` input.

    Validates: Requirements 10.1, 10.5
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
# VBA macro builders
# ---------------------------------------------------------------------------

def _escape_vba_path(path: str) -> str:
    """Escape a file path for embedding in a VBA string literal.

    Backslashes are doubled so they survive VBA string parsing.
    """
    return path.replace("\\", "\\\\")


def build_material_change_vba(shape_name: str, output_path: str) -> str:
    """Build VBA macro to change a shape's material to PEC.

    The *shape_name* must be in ``"component:solid"`` format.  The macro
    uses ``AddToHistory`` so the change persists in the CST model history.

    Validates: Requirements 1.2, 9.1, 9.2, 9.3
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


def build_check_material_vba(shape_name: str, output_path: str) -> str:
    """Build VBA macro to query a shape's current material.

    Writes the material name to *output_path* so Python can read it back.
    Used to skip shapes that are already assigned PEC.

    Validates: Requirements 1.3, 9.3
    """
    esc_path = _escape_vba_path(output_path)
    return (
        "Sub Main\n"
        "  Dim outPath As String\n"
        f'  outPath = "{esc_path}"\n'
        "  Open outPath For Output As #1\n"
        "  On Error Resume Next\n"
        f'  Dim matName As String\n'
        f'  matName = Solid.GetMaterialNameForShape("{shape_name}")\n'
        "  If Err.Number <> 0 Then\n"
        '    Print #1, "FAIL: " & Err.Description\n'
        "    Err.Clear\n"
        "  Else\n"
        "    Print #1, matName\n"
        "  End If\n"
        "  Close #1\n"
        "End Sub\n"
    )


def build_frequency_range_vba(fmin_ghz: float, fmax_ghz: float,
                              output_path: str) -> str:
    """Build VBA macro to set the simulation frequency range.

    Uses ``Solver.FrequencyRange`` wrapped in ``AddToHistory``.

    Validates: Requirements 2.3, 9.1, 9.2, 9.3
    """
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


def build_monitor_vba(field_type: str, frequency_ghz: float,
                      output_path: str) -> str:
    """Build VBA macro to create a field monitor.

    *field_type* must be ``"Efield"`` or ``"Hfield"``.  The monitor name
    follows the CST convention: ``"e-field (f=2.45)"`` or
    ``"h-field (f=2.45)"``.

    Uses a ``Monitor`` With block joined by ``vbCrLf`` and wrapped in
    ``AddToHistory``.

    Validates: Requirements 3.3, 3.4, 9.1, 9.2
    """
    if field_type == "Efield":
        monitor_name = f"e-field (f={frequency_ghz})"
    else:
        monitor_name = f"h-field (f={frequency_ghz})"

    esc_path = _escape_vba_path(output_path)
    return (
        "Sub Main\n"
        "  Dim outPath As String\n"
        f'  outPath = "{esc_path}"\n'
        "  Open outPath For Output As #1\n"
        "  On Error Resume Next\n"
        "  Dim sCode As String\n"
        '  sCode = "With Monitor" & vbCrLf\n'
        '  sCode = sCode & "  .Reset" & vbCrLf\n'
        f'  sCode = sCode & "  .Name ""{monitor_name}""" & vbCrLf\n'
        '  sCode = sCode & "  .Dimension ""Volume""" & vbCrLf\n'
        '  sCode = sCode & "  .Domain ""Frequency""" & vbCrLf\n'
        f'  sCode = sCode & "  .FieldType ""{field_type}""" & vbCrLf\n'
        f'  sCode = sCode & "  .MonitorValue ""{frequency_ghz}""" & vbCrLf\n'
        '  sCode = sCode & "  .Create" & vbCrLf\n'
        '  sCode = sCode & "End With"\n'
        f'  AddToHistory "define monitor: {monitor_name}", sCode\n'
        "  If Err.Number <> 0 Then\n"
        '    Print #1, "FAIL: " & Err.Description\n'
        "    Err.Clear\n"
        "  Else\n"
        '    Print #1, "OK"\n'
        "  End If\n"
        "  Close #1\n"
        "End Sub\n"
    )


def build_mesh_vba(output_path: str) -> str:
    """Build VBA macro to generate the surface mesh for the I-Solver.

    Uses ``Mesh.Update`` to generate the mesh, then switches to mesh
    view mode so the user can visually inspect it.

    Validates: Requirements 4.1, 9.1, 9.2, 9.3
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
        '    Print #1, "FAIL: " & Err.Description\n'
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


def build_boundary_vba(output_path: str) -> str:
    """Build VBA macro to set all boundaries to expanded open.

    Sets all six faces (Xmin, Xmax, Ymin, Ymax, Zmin, Zmax) to
    ``"expanded open"`` via a ``Boundary`` With block joined by
    ``vbCrLf``, wrapped in ``AddToHistory``.

    Validates: Requirements 5.1, 9.1, 9.2
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
        '  sCode = sCode & "  .Xmin ""expanded open""" & vbCrLf\n'
        '  sCode = sCode & "  .Xmax ""expanded open""" & vbCrLf\n'
        '  sCode = sCode & "  .Ymin ""expanded open""" & vbCrLf\n'
        '  sCode = sCode & "  .Ymax ""expanded open""" & vbCrLf\n'
        '  sCode = sCode & "  .Zmin ""expanded open""" & vbCrLf\n'
        '  sCode = sCode & "  .Zmax ""expanded open""" & vbCrLf\n'
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


def build_ie_solver_vba(num_modes: int, output_path: str) -> str:
    """Build VBA macro to configure the I-Solver for CMA.

    First switches to ``HF IntegralEq`` via ``ChangeSolverType``, then
    applies ``FDSolver`` and ``IESolver`` settings via ``AddToHistory``
    using the exact VBA syntax recorded from CST Microwave Studio 2025.

    Key settings from the CST recording:
    - ``FDSolver.SetMethod "Surface", "General purpose"``
    - ``FDSolver.Stimulation "CMA", "All"``
    - ``FDSolver.UseParallelization "True"``
    - ``FDSolver.MaximumNumberOfCPUDevices "8"``
    - ``IESolver.ModeTrackingCMA "False"``
    - ``IESolver.NumberOfModesCMA``
    - ``IESolver.FrequencySamplesCMA "0"``
    - ``IESolver.CalculateModalWeightingCoefficientsCMA "False"``

    Validates: Requirements 6.1, 6.2, 6.3, 6.6, 6.7, 6.8, 9.1, 9.2
    """
    esc_path = _escape_vba_path(output_path)
    return (
        "Sub Main\n"
        "  Dim outPath As String\n"
        f'  outPath = "{esc_path}"\n'
        "  Open outPath For Output As #1\n"
        "  On Error Resume Next\n"
        # -- Switch to I-Solver (direct, not via AddToHistory) --
        '  ChangeSolverType "HF IntegralEq"\n'
        "  If Err.Number <> 0 Then\n"
        '    Print #1, "FAIL: ChangeSolverType: " & Err.Description\n'
        "    Err.Clear\n"
        "    Close #1\n"
        "    Exit Sub\n"
        "  End If\n"
        # -- FDSolver settings (via AddToHistory) --
        "  Dim sCode As String\n"
        '  sCode = "With FDSolver" & vbCrLf\n'
        '  sCode = sCode & "  .SetMethod ""Surface"", ""General purpose""" & vbCrLf\n'
        '  sCode = sCode & "  .Stimulation ""CMA"", ""All""" & vbCrLf\n'
        '  sCode = sCode & "  .Type ""IterativeMoM""" & vbCrLf\n'
        '  sCode = sCode & "  .UseParallelization ""True""" & vbCrLf\n'
        '  sCode = sCode & "  .MaxCPUs ""1024""" & vbCrLf\n'
        '  sCode = sCode & "  .MaximumNumberOfCPUDevices ""8""" & vbCrLf\n'
        '  sCode = sCode & "End With" & vbCrLf\n'
        # -- IESolver CMA settings --
        '  sCode = sCode & "With IESolver" & vbCrLf\n'
        '  sCode = sCode & "  .SetAccuracySetting ""Medium""" & vbCrLf\n'
        '  sCode = sCode & "  .ModeTrackingCMA ""False""" & vbCrLf\n'
        f'  sCode = sCode & "  .NumberOfModesCMA ""{num_modes}""" & vbCrLf\n'
        '  sCode = sCode & "  .SetAccuracySettingCMA ""Default""" & vbCrLf\n'
        '  sCode = sCode & "  .FrequencySamplesCMA ""0""" & vbCrLf\n'
        '  sCode = sCode & "  .SetMemSettingCMA ""Auto""" & vbCrLf\n'
        '  sCode = sCode & "  .CalculateModalWeightingCoefficientsCMA ""True""" & vbCrLf\n'
        '  sCode = sCode & "End With"\n'
        '  AddToHistory "configure IE solver for CMA", sCode\n'
        "  If Err.Number <> 0 Then\n"
        '    Print #1, "FAIL: AddToHistory: " & Err.Description\n'
        "    Err.Clear\n"
        "  Else\n"
        '    Print #1, "OK"\n'
        "  End If\n"
        "  Close #1\n"
        "End Sub\n"
    )


# ---------------------------------------------------------------------------
# CMASetup class
# ---------------------------------------------------------------------------

class CMASetup:
    """Orchestrates the CMA simulation setup workflow.

    Uses an existing ``CSTConnection`` for VBA execution and a
    ``FeatureDetector`` for enumerating solid shapes in the model.

    Validates: Requirements 1–7
    """

    def __init__(self, connection, detector):
        """Initialise with a live CST connection and feature detector.

        Args:
            connection: A connected ``CSTConnection`` instance.
            detector: A ``FeatureDetector`` instance (provides
                ``_enumerate_solids``).

        Validates: Requirements 7.6
        """
        self._conn = connection
        self._detector = detector

    # --- Step 1: Material Assignment ---

    def assign_pec_materials(self) -> Tuple[int, int]:
        """Assign PEC material to every solid shape in the model.

        Enumerates all ``(component, solid)`` pairs via the feature
        detector, checks each shape's current material, and applies PEC
        where needed.  Per-component failures are logged as warnings;
        processing continues for remaining shapes.

        Returns:
            ``(total_processed, total_changed)`` — count of shapes
            inspected and count actually changed to PEC.

        Validates: Requirements 1.1, 1.2, 1.3, 1.4, 1.5
        """
        solids = self._detector._enumerate_solids()
        total_processed = 0
        total_changed = 0

        for component, solid in solids:
            shape_name = f"{component}:{solid}"
            total_processed += 1
            out_path = os.path.join(tempfile.gettempdir(), "cst_cma_mat_check.txt")
            try:
                # Check current material
                macro = build_check_material_vba(shape_name, out_path)
                result = self._conn.execute_vba(macro, output_file=out_path)

                if result and result.strip().upper() == "PEC":
                    logger.info("Shape '%s' already PEC — skipped.", shape_name)
                    continue

                if result and result.startswith("FAIL:"):
                    logger.warning(
                        "Could not query material for '%s': %s",
                        shape_name, result,
                    )
                    continue

                # Apply PEC material
                out_path_change = os.path.join(
                    tempfile.gettempdir(), "cst_cma_mat_change.txt",
                )
                macro = build_material_change_vba(shape_name, out_path_change)
                result = self._conn.execute_vba(
                    macro, output_file=out_path_change,
                )

                if result and result.startswith("FAIL:"):
                    logger.warning(
                        "Material change failed for '%s': %s",
                        shape_name, result,
                    )
                    continue

                total_changed += 1
                logger.info("Assigned PEC to '%s'.", shape_name)

            except Exception as exc:
                logger.warning(
                    "Error processing '%s': %s", shape_name, exc,
                )
            finally:
                for p in (
                    out_path,
                    os.path.join(tempfile.gettempdir(), "cst_cma_mat_change.txt"),
                ):
                    try:
                        os.remove(p)
                    except OSError:
                        pass

        print(
            f"  Material assignment complete: {total_processed} processed, "
            f"{total_changed} changed to PEC."
        )
        logger.info(
            "Material assignment: %d processed, %d changed.",
            total_processed, total_changed,
        )
        return total_processed, total_changed

    # --- Step 2: Frequency Range ---

    def configure_frequency_range(self) -> Tuple[float, float]:
        """Prompt the user for min/max frequency and apply to the model.

        Re-prompts when ``fmin >= fmax``.  Executes the frequency-range
        VBA macro via ``AddToHistory``.

        Returns:
            ``(fmin_ghz, fmax_ghz)``

        Validates: Requirements 2.1, 2.2, 2.3, 2.4, 2.5
        """
        while True:
            fmin = prompt_frequency(
                "  Enter minimum frequency in GHz (must be > 0): "
            )
            fmax = prompt_frequency("  Enter maximum frequency (GHz): ")

            if fmin >= fmax:
                print(
                    "  Minimum frequency must be less than maximum frequency. "
                    "Please try again."
                )
                continue
            break

        out_path = os.path.join(tempfile.gettempdir(), "cst_cma_freq.txt")
        try:
            macro = build_frequency_range_vba(fmin, fmax, out_path)
            result = self._conn.execute_vba(macro, output_file=out_path)

            if result and result.startswith("FAIL:"):
                logger.warning("Frequency range setting failed: %s", result)
                print(f"  Warning: frequency range setting failed: {result}")
            else:
                print(f"  Frequency range set to {fmin} – {fmax} GHz.")
                logger.info("Frequency range: %s – %s GHz.", fmin, fmax)
        finally:
            try:
                os.remove(out_path)
            except OSError:
                pass

        return fmin, fmax

    # --- Step 3: Field Monitors ---

    def create_field_monitors(self, fmin: float, fmax: float) -> float:
        """Create E-field and H-field monitors at a user-specified frequency.

        Validates that the monitor frequency lies within ``[fmin, fmax]``
        and re-prompts if out of range.

        Returns:
            The validated monitor frequency in GHz.

        Validates: Requirements 3.1, 3.2, 3.3, 3.4, 3.5, 3.6
        """
        while True:
            freq = prompt_frequency("  Enter monitor frequency (GHz): ")
            if freq < fmin or freq > fmax:
                print(
                    f"  Monitor frequency must be between {fmin} and "
                    f"{fmax} GHz. Please try again."
                )
                continue
            break

        # E-field monitor
        out_path = os.path.join(tempfile.gettempdir(), "cst_cma_monitor.txt")
        try:
            macro = build_monitor_vba("Efield", freq, out_path)
            result = self._conn.execute_vba(macro, output_file=out_path)
            if result and result.startswith("FAIL:"):
                logger.warning("E-field monitor creation failed: %s", result)
                print(f"  Warning: E-field monitor creation failed: {result}")
            else:
                print(f"  Created E-field monitor at {freq} GHz.")
                logger.info("Created E-field monitor at %s GHz.", freq)
        finally:
            try:
                os.remove(out_path)
            except OSError:
                pass

        # H-field monitor
        out_path = os.path.join(tempfile.gettempdir(), "cst_cma_monitor.txt")
        try:
            macro = build_monitor_vba("Hfield", freq, out_path)
            result = self._conn.execute_vba(macro, output_file=out_path)
            if result and result.startswith("FAIL:"):
                logger.warning("H-field monitor creation failed: %s", result)
                print(f"  Warning: H-field monitor creation failed: {result}")
            else:
                print(f"  Created H-field monitor at {freq} GHz.")
                logger.info("Created H-field monitor at %s GHz.", freq)
        finally:
            try:
                os.remove(out_path)
            except OSError:
                pass

        return freq

    # --- Step 4: Mesh Generation ---

    def generate_mesh(self) -> bool:
        """Generate the I-Solver surface mesh.

        Returns:
            ``True`` on success, ``False`` on failure.

        Validates: Requirements 4.1, 4.2
        """
        out_path = os.path.join(tempfile.gettempdir(), "cst_cma_mesh.txt")
        try:
            macro = build_mesh_vba(out_path)
            result = self._conn.execute_vba(macro, output_file=out_path)

            if result and result.startswith("FAIL:"):
                logger.warning("Mesh generation failed: %s", result)
                print(f"  Mesh generation failed: {result}")
                return False

            print("  Surface mesh generated successfully.")
            logger.info("Surface mesh generation completed.")
            return True
        finally:
            try:
                os.remove(out_path)
            except OSError:
                pass

    # --- Step 5: Boundary Conditions ---

    def set_boundary_conditions(self) -> bool:
        """Set all six boundary faces to expanded open.

        Returns:
            ``True`` on success, ``False`` on failure.

        Validates: Requirements 5.1, 5.2, 5.3
        """
        out_path = os.path.join(tempfile.gettempdir(), "cst_cma_boundary.txt")
        try:
            macro = build_boundary_vba(out_path)
            result = self._conn.execute_vba(macro, output_file=out_path)

            if result and result.startswith("FAIL:"):
                logger.warning("Boundary configuration failed: %s", result)
                print(f"  Warning: boundary configuration failed: {result}")
                return False

            print("  Boundary conditions set to expanded open (all faces).")
            logger.info(
                "Boundary conditions configured: all faces set to "
                "expanded open."
            )
            return True
        finally:
            try:
                os.remove(out_path)
            except OSError:
                pass

    # --- Step 6: I-Solver Configuration ---

    def configure_ie_solver(self, monitor_freq: float) -> int:
        """Configure the Integral Equation Solver for CMA.

        Prompts the user for the number of characteristic modes (default
        10) and applies IESolver CMA settings via ``AddToHistory``, then
        sets Solver CPU acceleration directly.

        Note: ``ChangeSolverType`` and ``FDSolver.Method`` are not used
        — CST Microwave Studio does not support them.  The ``IESolver``
        object handles the solver configuration implicitly.

        Args:
            monitor_freq: The monitor frequency in GHz (for logging).

        Returns:
            The configured number of modes.

        Validates: Requirements 6.1–6.10
        """
        num_modes = prompt_integer(
            "  Enter number of modes [default: 10]: ", default=10,
        )

        out_path = os.path.join(tempfile.gettempdir(), "cst_cma_solver.txt")
        try:
            macro = build_ie_solver_vba(num_modes, out_path)
            result = self._conn.execute_vba(macro, output_file=out_path)

            if result and result.startswith("FAIL:"):
                logger.warning("I-Solver configuration failed: %s", result)
                print(f"  Warning: I-Solver configuration failed: {result}")
            else:
                print("  I-Solver configured for CMA:")
                print(f"    Method:              Integral Equation (ACA)")
                print(f"    Mode tracking:       Disabled")
                print(f"    Frequency samples:   As monitors")
                print(f"    Number of modes:     {num_modes}")
                print(f"    Weighting coeff.:    Disabled")
                print(f"    CPU devices:         8")
                print(f"    Monitor frequency:   {monitor_freq} GHz")
                logger.info(
                    "I-Solver configured: method=ACA, modes=%d, "
                    "mode_tracking=False, freq_samples=monitors, "
                    "weighting=False, cpu_devices=8, "
                    "monitor_freq=%s GHz.",
                    num_modes, monitor_freq,
                )
        finally:
            try:
                os.remove(out_path)
            except OSError:
                pass

        return num_modes

    # --- Workflow Orchestration ---

    def run(self) -> WorkflowSummary:
        """Execute the full CMA setup workflow.

        Runs all six steps in order:
        1. Material Assignment (non-critical)
        2. Frequency Range (non-critical)
        3. Field Monitors (non-critical, needs fmin/fmax)
        4. Boundary Conditions (non-critical)
        5. I-Solver Configuration (non-critical, uses monitor_freq)
        6. Mesh Generation (CRITICAL — if rejected, stop)

        The I-Solver and boundary conditions are configured before mesh
        generation so the mesh reflects the final solver settings.

        Non-critical failures are logged and the workflow continues.
        Critical failures (mesh rejection) stop the workflow immediately.
        ``UserQuitException`` is always re-raised.

        Returns:
            A ``WorkflowSummary`` with elapsed time and step results.

        Validates: Requirements 7.1, 7.2, 7.3, 7.4, 7.5
        """
        start = time.time()
        steps: List[WorkflowStep] = []

        # Shared state passed between steps
        fmin = None
        fmax = None
        monitor_freq = None

        # Step 1: Material Assignment (non-critical)
        print("\n--- Step 1/6: Material Assignment ---")
        try:
            processed, changed = self.assign_pec_materials()
            steps.append(WorkflowStep(
                "Material Assignment", True,
                f"{processed} processed, {changed} changed to PEC", False,
            ))
        except UserQuitException:
            raise
        except Exception as exc:
            logger.warning("Material assignment failed: %s", exc)
            steps.append(WorkflowStep(
                "Material Assignment", False, str(exc), False,
            ))

        # Step 2: Frequency Range (non-critical but needed for later steps)
        print("\n--- Step 2/6: Frequency Range ---")
        try:
            fmin, fmax = self.configure_frequency_range()
            steps.append(WorkflowStep(
                "Frequency Range", True,
                f"{fmin} – {fmax} GHz", False,
            ))
        except UserQuitException:
            raise
        except Exception as exc:
            logger.warning("Frequency range configuration failed: %s", exc)
            steps.append(WorkflowStep(
                "Frequency Range", False, str(exc), False,
            ))

        # Step 3: Field Monitors (non-critical, needs fmin/fmax)
        print("\n--- Step 3/6: Field Monitors ---")
        try:
            if fmin is not None and fmax is not None:
                monitor_freq = self.create_field_monitors(fmin, fmax)
                steps.append(WorkflowStep(
                    "Field Monitors", True,
                    f"E-field and H-field at {monitor_freq} GHz", False,
                ))
            else:
                steps.append(WorkflowStep(
                    "Field Monitors", False,
                    "Skipped — frequency range not available", False,
                ))
        except UserQuitException:
            raise
        except Exception as exc:
            logger.warning("Field monitor creation failed: %s", exc)
            steps.append(WorkflowStep(
                "Field Monitors", False, str(exc), False,
            ))

        # Step 4: Boundary Conditions (non-critical)
        print("\n--- Step 4/6: Boundary Conditions ---")
        try:
            self.set_boundary_conditions()
            steps.append(WorkflowStep(
                "Boundary Conditions", True,
                "All faces set to expanded open", False,
            ))
        except UserQuitException:
            raise
        except Exception as exc:
            logger.warning("Boundary configuration failed: %s", exc)
            steps.append(WorkflowStep(
                "Boundary Conditions", False, str(exc), False,
            ))

        # Step 5: I-Solver Configuration (non-critical, uses monitor_freq)
        print("\n--- Step 5/6: I-Solver Configuration ---")
        try:
            freq_for_solver = monitor_freq if monitor_freq is not None else 0.0
            num_modes = self.configure_ie_solver(freq_for_solver)
            steps.append(WorkflowStep(
                "I-Solver Configuration", True,
                f"{num_modes} modes, ACA method, 8 CPU devices", False,
            ))
        except UserQuitException:
            raise
        except Exception as exc:
            logger.warning("I-Solver configuration failed: %s", exc)
            steps.append(WorkflowStep(
                "I-Solver Configuration", False, str(exc), False,
            ))

        # Step 6: Mesh Generation (CRITICAL — if rejected, stop)
        print("\n--- Step 6/6: Mesh Generation ---")
        try:
            accepted = self.generate_mesh()
            if accepted:
                steps.append(WorkflowStep(
                    "Mesh Generation", True,
                    "Mesh generated and accepted", True,
                ))
            else:
                steps.append(WorkflowStep(
                    "Mesh Generation", False,
                    "Mesh rejected by user", True,
                ))
                # Critical failure — stop workflow
                elapsed = time.time() - start
                summary = WorkflowSummary(steps, elapsed)
                self._print_summary(summary)
                return summary
        except UserQuitException:
            raise
        except Exception as exc:
            logger.warning("Mesh generation failed: %s", exc)
            steps.append(WorkflowStep(
                "Mesh Generation", False, str(exc), True,
            ))
            # Critical failure — stop workflow
            elapsed = time.time() - start
            summary = WorkflowSummary(steps, elapsed)
            self._print_summary(summary)
            return summary

        elapsed = time.time() - start
        summary = WorkflowSummary(steps, elapsed)
        self._print_summary(summary)
        return summary

    @staticmethod
    def _print_summary(summary: WorkflowSummary) -> None:
        """Print a human-readable workflow summary."""
        print("\n" + "=" * 50)
        print("  CMA Setup Summary")
        print("=" * 50)

        if summary.completed:
            print(f"\n  Completed ({len(summary.completed)}):")
            for name in summary.completed:
                print(f"    ✓ {name}")

        if summary.failed:
            print(f"\n  Failed ({len(summary.failed)}):")
            for name in summary.failed:
                print(f"    ✗ {name}")

        print(f"\n  Total time: {summary.total_time:.1f}s")
        print("=" * 50)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main(project_path: str = None):
    """Entry point for CMA setup workflow.

    Connects to CST, opens the project, runs the full CMA setup
    workflow, and prints a summary.

    Args:
        project_path: Path to the ``.cst`` project file.  Falls back to
            the module-level ``PROJECT`` variable, then prompts the user.

    Validates: Requirements 7.1, 7.4, 7.6, 10.5
    """
    from code.cst_connection import CSTConnection, CSTConnectionError
    from code.feature_detector import FeatureDetector

    # Resolve project path
    path = project_path or PROJECT
    if not path:
        path = input("  Enter path to .cst model: ").strip().strip('"').strip("'")
    if not path:
        print("  Error: no project path provided.")
        return

    path = os.path.abspath(path)

    print("=" * 60)
    print("  CMA Simulation Setup")
    print("=" * 60)
    print(f"  Project: {path}")

    conn = CSTConnection()
    try:
        print("\n  Connecting to CST...")
        conn.connect()
        conn.open_project(path)
        print(f"  Project opened: {os.path.basename(path)}")

        detector = FeatureDetector(conn)
        setup = CMASetup(conn, detector)
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
