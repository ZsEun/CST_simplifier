"""CST Studio Suite 2025 COM connection management.

Provides CSTConnection class for connecting to CST Studio Suite 2025
via COM automation, opening projects, executing VBA macros via temp
.bas files, and managing COM reference lifecycle.

Key CST 2025 API notes:
- OpenFile() must be called via raw _oleobj_.Invoke (returns int, confuses pywin32)
- Active3D() returns the project COM object
- RunVBA() does NOT exist; use RunScript() with a temp .bas file
- RunScript() must also be called via raw _oleobj_.Invoke

Validates: Requirements 6.1, 6.2, 6.3, 6.4, 6.5, 7.1, 7.2, 7.3
"""

import logging
import os
import tempfile

logger = logging.getLogger(__name__)

CST_PROG_ID = "CSTStudio.Application"


class CSTConnectionError(Exception):
    """Raised when a CST COM connection or operation fails."""


def call_method(com_obj, method_name, *args):
    """Call a COM method by name using raw oleobj dispatch.

    CST 2025's COM methods sometimes return unexpected types that
    confuse pywin32's automatic dispatch. This helper uses raw
    IDispatch.Invoke to bypass that issue.

    Args:
        com_obj: A pywin32 COM object.
        method_name: Name of the method to call.
        *args: Positional arguments to pass.

    Returns:
        The raw return value from the COM call.
    """
    import pythoncom

    oleobj = com_obj._oleobj_
    disp_id = oleobj.GetIDsOfNames(method_name)
    if isinstance(disp_id, tuple):
        disp_id = disp_id[0]
    return oleobj.Invoke(disp_id, 0, pythoncom.DISPATCH_METHOD, 0, *args)


class CSTConnection:
    """Manages the COM connection to CST Studio Suite 2025.

    Connects to a running CST instance, opens .cst project files,
    executes VBA macros via temp .bas files + RunScript, and ensures
    proper COM reference cleanup.

    Usage::

        with CSTConnection() as conn:
            conn.connect()
            conn.open_project(r"C:\\path\\to\\model.cst")
            result = conn.execute_vba('Sub Main\\n  MsgBox "hello"\\nEnd Sub')
    """

    def __init__(self):
        self._app = None
        self._project = None

    @property
    def app(self):
        """The CST application COM object."""
        return self._app

    @property
    def project(self):
        """The currently open CST project COM object."""
        return self._project

    def connect(self):
        """Connect to a running CST instance or launch a new one.

        Validates: Requirements 6.1, 6.2
        """
        import win32com.client

        try:
            self._app = win32com.client.GetActiveObject(CST_PROG_ID)
            logger.info("Connected to running CST Studio Suite instance.")
            return
        except Exception:
            logger.info("No running CST instance found. Launching new instance...")

        try:
            self._app = win32com.client.Dispatch(CST_PROG_ID)
            logger.info("Launched new CST Studio Suite instance.")
        except Exception as exc:
            raise CSTConnectionError(
                f"Failed to connect to or launch CST Studio Suite. "
                f"Ensure CST 2025 is installed and COM is registered. "
                f"Details: {exc}"
            ) from exc

    def open_project(self, path: str):
        """Open a CST project file.

        Uses raw _oleobj_.Invoke for OpenFile (CST 2025 returns an int
        that confuses pywin32), then Active3D() to get the project object.

        Validates: Requirements 6.3, 6.4
        """
        if self._app is None:
            raise CSTConnectionError("Not connected to CST. Call connect() first.")

        abs_path = os.path.abspath(path)

        if not abs_path.lower().endswith(".cst"):
            raise CSTConnectionError(
                f"Invalid project file: '{abs_path}'. Expected .cst extension."
            )
        if not os.path.isfile(abs_path):
            raise CSTConnectionError(f"Project file not found: '{abs_path}'.")

        try:
            # Must use raw invoke — OpenFile returns int, confusing pywin32
            call_method(self._app, "OpenFile", abs_path)

            self._project = self._app.Active3D()
            if self._project is None:
                raise CSTConnectionError(
                    f"Opened '{abs_path}' but could not retrieve the active project."
                )
            logger.info("Opened project: %s", abs_path)
        except CSTConnectionError:
            raise
        except Exception as exc:
            raise CSTConnectionError(
                f"Failed to open project '{abs_path}'. Details: {exc}"
            ) from exc

    def execute_vba(self, macro_code: str, output_file: str = None) -> str:
        """Execute a VBA macro in the active CST project.

        Writes the macro to a temp .bas file and runs it via RunScript.
        CST 2025 does NOT have a RunVBA method.

        If output_file is provided, reads and returns its contents after
        execution (for VBA macros that write results to a file).

        Args:
            macro_code: VBA code string (must contain Sub Main...End Sub).
            output_file: Optional path to a file the VBA writes output to.

        Returns:
            Contents of output_file if provided, else empty string.

        Validates: Requirements 7.1, 7.2, 7.3
        """
        if self._project is None:
            raise CSTConnectionError("No project is open. Call open_project() first.")

        tmp_path = os.path.join(tempfile.gettempdir(), "cst_macro.bas")
        try:
            with open(tmp_path, "w") as f:
                f.write(macro_code)

            call_method(self._project, "RunScript", tmp_path)

            if output_file and os.path.isfile(output_file):
                with open(output_file, "r") as f:
                    return f.read().strip()
            return ""
        except CSTConnectionError:
            raise
        except Exception as exc:
            raise CSTConnectionError(
                f"VBA execution failed. Details: {exc}"
            ) from exc
        finally:
            try:
                os.remove(tmp_path)
            except OSError:
                pass

    def close(self):
        """Release all COM references. Safe to call multiple times.

        Validates: Requirements 6.5
        """
        if self._project is not None:
            self._project = None
            logger.info("Released CST project COM reference.")
        if self._app is not None:
            self._app = None
            logger.info("Released CST application COM reference.")

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
        return False
