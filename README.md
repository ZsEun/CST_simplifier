# CST CAD Model Simplifier

Automates detection and filling of screw holes in STP-imported CAD models within CST Studio Suite 2025. Connects via COM automation, exports SAT geometry, parses topology to find cone-surface hole features, filters out board-edge fillets, groups hole faces, and removes them using AddToHistory-based VBA commands.

## Requirements

- Windows (COM automation required)
- CST Studio Suite 2025 (installed and COM registered)
- Python 3.10+
- pywin32

## Installation

```bash
pip install -r code/requirements.txt
```

## Project Structure

```
code/
    __init__.py          - Package init
    __main__.py          - Entry point for "python -m code"
    models.py            - Data classes (FaceInfo, SessionSummary, etc.)
    cst_connection.py    - COM connection to CST Studio Suite 2025
    feature_detector.py  - SAT parser + hole detection + board edge filter
    simplifier.py        - Progressive hole filling via AddToHistory
    main.py              - CLI entry point with argparse
    run_sunray_v3.py     - Standalone runner for Sunray model (SAT-based fill)
    run_sunray_v4.py     - Full pipeline with ghost face scan phase
    requirements.txt     - Python dependencies
    tests/               - Test suite
```

## Usage

1. Open CST Studio Suite 2025 and load your `.cst` project.

2. Generic CLI:

```bash
python -m code --project "C:\path\to\your\model.cst"
```

Add `--auto` for non-interactive mode (fills all holes without asking):

```bash
python -m code --project "C:\path\to\your\model.cst" --auto
```

3. For the Sunray model (interactive, with ghost face scan):

```bash
python -m code.run_sunray_v4
```

## How It Works

### Detection Pipeline

1. Connects to running CST instance via `win32com` COM automation
2. Enumerates all solids by walking the CST result tree via VBA macros
3. Exports ACIS SAT geometry for each solid
4. Parses SAT file to extract face entities, surface types, pid mappings, and topology-based adjacency
5. Identifies cone-surface faces as screw hole seeds
6. Filters out board-edge fillets using the board edge filter algorithm
7. Groups remaining seeds into hole groups via BFS adjacency walk

### Board Edge Filter

Distinguishes board-edge fillets from actual screw holes:

1. Finds the 2 largest plane faces (PCB top/bottom) as board reference
2. For each cone seed, BFS-walks adjacency (excluding board ref faces) to build the feature loop
3. Computes the union bounding box of the loop
4. If the loop spans >= 50% of the board in both in-plane axes, it's a board edge fillet (filtered out)

### Fill Algorithm

For each detected hole group:

1. Picks all loop faces via `AddToHistory` (required for CST persistence)
2. Calls `LocalModification.RemoveSelectedFaces` via `AddToHistory`
3. On failure, expands face set with adjacent faces and retries (up to 5 attempts)
4. Falls back to consecutive ID probing if expansion fails
5. Skips groups whose faces were already consumed by previous fills

### Ghost Face Scan (v4)

Some CST faces have no corresponding entity in the SAT export ("ghost faces"). After the SAT-based fill completes, v4 adds a second phase:

1. Identifies all face IDs missing from the SAT export
2. For each ghost face, checks if it still exists (may have been consumed by a previous fill)
3. Highlights the face and asks the user to confirm it's part of a hole
4. Tries consecutive ID windows (sizes 2-5) around the face, using silent fills (no GUI error popups)
5. For ghost faces near skipped SAT groups, combines the skipped group's seed faces with the ghost face for a more complete fill

## CST 2025 COM API Notes

### What works
- `win32com.client.GetActiveObject("CSTStudio.Application")`
- `app._oleobj_.Invoke(...)` via `call_method()` for OpenFile, RunScript
- `Pick.PickFaceFromId` (face ID must be a string)
- `LocalModification.RemoveSelectedFaces`
- `AddToHistory` (required for model changes to persist)
- `SAT.Reset/.FileName/.Write` for ACIS SAT export

### What does NOT work
- `proj.RunVBA(code)` (does not exist)
- `Solid.GetNumberOfFaces()`, `Solid.GetFaceType()`, etc.
- `Modeler.Undo`, `proj.Undo()`, `app.Undo()`
- RunScript alone does NOT persist model changes

## License

MIT
