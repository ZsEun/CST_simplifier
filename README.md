# CST CAD Model Simplifier

Automates detection and removal of holes and dimples in STP-imported CAD models within CST Studio Suite 2025. Two tools:

1. **PCB Board Simplifier** (`run_sunray_v6.py`) — removes screw holes from flat PCB boards
2. **Shield Can Cover Simplifier** (`run_led_v2.py`) — removes dimples/holes from shield can cover side walls
3. **Shield Can Frame Simplifier** (`run_frame_v1.py`) — removes dimples/holes from shield can frame side walls
4. **Shield Can Contact Bridge** (`debug_contact_v17_shieldcan.py`) — detects gap between shield can cover and frame, creates a bridge by extruding the frame's top face to close the gap

Both connect via COM automation, export SAT geometry, parse topology, and fill features using `AddToHistory` + `RemoveSelectedFaces`.

## Requirements

- Windows (COM automation required)
- CST Studio Suite 2025
- Python 3.10+
- pywin32

## Installation

```bash
pip install -r code/requirements.txt
```

## Quick Start

### GUI (recommended)

```bash
python -m code.gui
```

Simple GUI with three buttons:
1. Fill Holes on PCB Board
2. Simplify Shield Can (Cover + Frame dimples)
3. Bridge Shield Can Cover-Frame Gap

Browse to your .cst file, click a button. Prompts appear as Yes/No/Quit dialogs. Output log shown in the GUI.

### Combined CLI (handles multi-component models)

```bash
python -m code.run_combined_v1
```

Auto-classifies components by name:
- "BOARD" → PCB screw hole removal
- "COVER" → Shield can cover dimple removal
- "FRAM" → Shield can frame dimple removal
- Other → Skipped

Highlights each component, asks for confirmation, then processes with the right algorithm.

### Individual tools

```bash
python -m code.run_sunray_v6      # PCB board only
python -m code.run_led_v2         # Shield can cover only
python -m code.run_frame_v1       # Shield can frame only
python -m code.debug_contact_v17_shieldcan  # Shield can cover-frame bridge
```

## Project Structure

```
code/
    cst_connection.py    - COM connection to CST 2025
    feature_detector.py  - SAT parser + hole detection + board edge filter
    simplifier.py        - Progressive hole filling via AddToHistory
    wall_detector.py     - Shield can wall + dimple detection
    models.py            - Data classes
    run_combined_v1.py   - Combined simplifier (auto-classifies components)
    run_sunray_v6.py     - PCB simplifier (standalone)
    run_led_v2.py        - Shield can cover simplifier (standalone)
    run_frame_v1.py      - Shield can frame simplifier (standalone)
    run_contact_check.py - Contact checker (generic, experimental)
    debug_contact_v17_shieldcan.py - Shield can cover-frame bridge (recommended)
    gui.py               - GUI launcher with Yes/No/Quit buttons
    run_led_v1.py        - Earlier shield can cover version
```


## PCB Board Simplifier Algorithm

1. Export SAT, parse face types/adjacency/bboxes
2. Find cone-surface seed faces (screw hole walls)
3. Filter out board-edge fillets (span ≥50% of board in both in-plane axes)
4. Group seeds into holes via BFS adjacency walk
5. For each hole: highlight, ask y/n/q, fill via AddToHistory
6. Progressive expansion on failure, consecutive ID probe fallback
7. Ghost face scan for faces missing from SAT export

## Shield Can Cover Simplifier Algorithm

### 1. Wall Detection (Cover)
- Find top face (largest plane face by bbox area)
- Walk adjacency: top face → curved corner faces (torus/spline) → perpendicular plane faces
- These perpendicular planes are the side walls (works for any angle)

## Shield Can Frame Simplifier Algorithm

### 1. Wall Detection (Frame)
- Find bottom face (largest plane face by bbox area)
- Side walls = all plane faces whose normal is perpendicular to bottom face normal (|dot| ≤ 0.05)
- Simpler than cover because frame walls don't need adjacency walk

### 2. Dimple Detection (per wall)
For each wall, find dimple faces using local UVW coordinate projection:
- **W axis** = wall normal direction
- **U, V axes** = in-plane directions (computed via cross product)
- Project wall bbox into UV → get wall's UV range
- For each face in the model:
  - **UV containment**: face UV footprint must be within wall's UV range
  - **W proximity**: face must be close to wall in normal direction (|face_W - wall_W| < 2 × face's max UV span)
  - **UV span**: face must be small relative to wall (< 50% of wall span)
  - **Normal filter**: reject plane/cone faces with normal perpendicular to wall (structural edges)
  - **Exclude**: top face, all wall faces, wall's direct adjacency neighbors
- **Zero-bbox expansion**: add adjacency neighbors with zero bboxes (spline surfaces whose bbox couldn't be extracted)

### 3. Fill
- Sort walls by area ascending (small walls claim dimples first)
- Track consumed faces to avoid duplicate fills
- For each wall with dimples: set WCS aligned with wall, highlight, ask user, fill via AddToHistory
- Use silent fill (`_try_fill_hole_silent`) to avoid GUI error popups

## Shield Can Contact Bridge Algorithm

Bridges the gap between shield can cover and frame. Specifically designed for the case where the cover sits on top of the frame with a small gap between their mating rims.

### Algorithm
1. **Determine orientation**: Compute bboxes of both components. The thin axis is the stack axis (typically Z). The two larger axes form the plane (typically XY).
2. **Determine stack direction**: Compare cover and frame centers along the stack axis. Cover above frame → +Z direction.
3. **Find frame top face**: Among all plane faces on the frame, find the one that:
   - Has normal along the stack axis (|normal[stack_axis]| > 0.9)
   - Spans ≥90% of the frame bbox in both plane axes
   - Is closest to the frame's max position along the stack axis
4. **Find matching cover face**: Among all plane faces on the cover, find the one parallel to the frame top face and closest to it along the stack axis.
5. **Bridge**: Extrude the frame top face along the stack axis by the gap distance to close the gap. Uses `AddToHistory` for persistence.

### Limitations
- Only works for shield can cover + frame geometry (flat mating rims)
- Assumes cover and frame are oriented along a principal axis (X, Y, or Z)
- Does not handle angled or curved mating surfaces

## CST 2025 COM API Notes

### What works
- `win32com.client.GetActiveObject("CSTStudio.Application")`
- `Pick.PickFaceFromId`, `LocalModification.RemoveSelectedFaces`
- `AddToHistory` (required for persistence)
- `SAT.Write` for ACIS SAT export
- `WCS.SetOrigin`, `WCS.SetNormal`, `WCS.SetUVector`, `WCS.ActivateWCS`

### What does NOT work
- `Solid.GetNumberOfFaces()`, `Solid.GetFaceType()`, etc.
- All `Plot.Zoom*`, `Plot.Pan`, `Plot.SetCamera` methods
- `Pick.GetPickedFaceBoundingBox`, `Solid.GetFaceBoundingBox`
- RunScript alone does NOT persist model changes

## License

MIT
