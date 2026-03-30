# CST CAD Model Simplifier

Automates detection and removal of holes and dimples in STP-imported CAD models within CST Studio Suite 2025. Two tools:

1. **PCB Board Simplifier** (`run_sunray_v6.py`) — removes screw holes from flat PCB boards
2. **Shield Can Simplifier** (`run_led_v2.py`) — removes dimples/holes from shield can side walls

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

### PCB Board (screw holes)

```bash
python -m code.run_sunray_v6
```

### Shield Can (dimples on side walls)

```bash
python -m code.run_led_v2
```

Both prompt for the `.cst` model path and guide you through interactive filling.

## Project Structure

```
code/
    cst_connection.py    - COM connection to CST 2025
    feature_detector.py  - SAT parser + hole detection + board edge filter
    simplifier.py        - Progressive hole filling via AddToHistory
    wall_detector.py     - Shield can wall + dimple detection
    models.py            - Data classes
    run_sunray_v6.py     - PCB simplifier (latest)
    run_led_v2.py        - Shield can simplifier (latest)
    run_sunray_v3-v5.py  - Earlier PCB versions
    run_led_v1.py        - Earlier shield can version
```


## PCB Board Simplifier Algorithm

1. Export SAT, parse face types/adjacency/bboxes
2. Find cone-surface seed faces (screw hole walls)
3. Filter out board-edge fillets (span ≥50% of board in both in-plane axes)
4. Group seeds into holes via BFS adjacency walk
5. For each hole: highlight, ask y/n/q, fill via AddToHistory
6. Progressive expansion on failure, consecutive ID probe fallback
7. Ghost face scan for faces missing from SAT export

## Shield Can Simplifier Algorithm

### 1. Wall Detection
- Find top face (largest plane face by bbox area)
- Walk adjacency: top face → curved corner faces (torus/spline) → perpendicular plane faces
- These perpendicular planes are the side walls (works for any angle, not just axis-aligned)

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
