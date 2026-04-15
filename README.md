# CST CAD Model Simplifier

Automates detection and removal of holes and dimples in STP-imported CAD models within CST Studio Suite 2025. Tools:

1. **PCB Board Simplifier** (`run_sunray_v6.py`) — removes screw holes from flat PCB boards
2. **Shield Can Simplifier** (`run_shieldcan.py`) — removes dimples/holes from shield can cover and frame side walls, with bbox-based cover/frame pairing, two-pass verification workflow
3. **Shield Can Contact Bridge** (`debug_contact_v17_shieldcan.py`) — detects gap between shield can cover and frame, creates a bridge by extruding the frame's top face to close the gap
4. **PCB Grounding Bridge** (`debug_pcb_edge_v2.py`) — detects components near the PCB with gaps, bridges them to the PCB by extruding the closest parallel face
5. **Component Cleanup** (`gui_cleanup.py`) — identifies and deletes plastic/unnecessary components by keyword matching, with auto-delete, exclude lists, and keyword import/export
7. **Connector Replacement** (`debug_connector_v2.py`) — replaces connector components with PEC blocks, bridges to FPC and PCB, merges bridges, and cleans up overlapping components

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
4. Bridge Grounding for PCB
5. Replace Connector

Browse to your .cst file, click a button. Prompts appear as Yes/No/Quit dialogs (Quit stops the entire tool immediately). Output log shown in the GUI. Each run saves a timestamped log file next to the .cst project.

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
    wall_detector.py     - Shield can wall + dimple detection (W-touch, aspect ratio, UV overlap)
    models.py            - Data classes
    shield_can_dialog.py - 3-column classifier dialog for cover/frame/one-piece
    component_cache.py   - Shared component cache (PCB/FPC names across tool runs)
    run_combined_v1.py   - Combined simplifier (auto-classifies components)
    run_sunray_v6.py     - PCB simplifier (standalone)
    run_shieldcan.py     - Shield can simplifier (cover + frame, two-pass workflow)
    run_contact_check.py - Contact checker (generic, experimental)
    debug_shieldcan_walls.py - Shield can debug script (interactive wall + dimple testing)
    debug_contact_v17_shieldcan.py - Shield can cover-frame bridge
    debug_pcb_edge_v2.py     - PCB grounding bridge
    debug_connector_v2.py    - Connector replacement with FPC/PCB bridging
    gui_cleanup.py       - Component cleanup GUI (separate tool)
    gui.py               - GUI launcher (5 buttons)
```


## PCB Board Simplifier Algorithm

Validated on Sunray_metal_v4_fun1.

1. User provides the PCB component name directly (no auto-scan)
2. Fuzzy match against all solids, list multiple matches, SelectTreeItem to confirm
3. Export SAT, parse face types/adjacency/bboxes
4. Find cone-surface seed faces (screw hole walls)
5. Filter out board-edge fillets (span ≥50% of board in both in-plane axes)
6. Group seeds into holes via BFS adjacency walk
7. For each hole: highlight, move WCS crosshair to hole center, ask y/n/q, fill via silent RunScript
8. Progressive expansion on failure, consecutive ID probe fallback
9. Ghost face scan for faces missing from SAT export
10. All fill operations use RunScript with On Error Resume Next (no CST GUI error popups)

## Shield Can Cover Simplifier Algorithm

**Note**: This is the best achievable version using rule-based SAT geometry filtering. Some edge cases (corner fillets near wall intersections) may still be incorrectly detected. A computer vision-based approach is planned for future improvement.

### 1. Wall Detection (Cover) — Validated with W-touch + Aspect Ratio
- Find reference face (largest plane face by bbox area) → W axis = reference normal
- Walk SAT adjacency: reference face → curved corner fillet faces → perpendicular plane faces
- **W-touch validation**: each candidate wall's bbox must physically touch its connecting fillet's bbox in the W direction (within 0.1mm tolerance). Rejects dimple faces that share SAT edges with fillets but aren't physically adjacent.
- **Aspect ratio validation**: in wall-local coords (W=wall normal, U=ref normal, V=W×U), V span must be ≥ U span. Real walls are longer than tall; small structural faces are not.

### 2. Wall Sub-Grouping
- **Copper thickness measurement**: find two largest parallel plane faces, measure distance → copper thickness
- **Group by normal**: all walls with same normal direction (dot > 0.99) go together
- **Recursive W-distance splitting**: within each normal-group, find root face (largest area), split by W distance from root (threshold = 1.5× copper thickness). Recurse if sub-group has >2 faces.
- **UV overlap splitting**: for sub-groups still >2 faces, split by UV bbox overlap. Walls at different positions along the same side don't overlap in UV → separate sub-groups.

### 3. Dimple Detection (per wall sub-group)
For each wall in the sub-group, find dimple faces using local UVW coordinate projection:
- **W axis** = wall normal direction
- **U axis** = short edge of wall, **V axis** = long edge (swapped based on bbox projection)
- **Exclude set**: reference face + all wall faces + corner fillet faces
- For each non-excluded face:
  - **W proximity**: face center's W distance from wall < 2× face's max UV span
  - **UV containment**: face UV footprint within wall's UV range (0.1mm margin)
  - **V span check**: face V span < 50% of wall V span (only check V, skip U — dimples can span full wall width)
  - **Normal filter**: reject plane/cone faces perpendicular to wall (|dot| < 0.3)
- **Zero-bbox expansion**: add adjacency neighbors with zero bboxes (spline surfaces), skip faces adjacent to corner fillets
- Merge dimples from all walls in the sub-group, remove already-consumed faces

### 4. Two-Pass Workflow
- **Pass 1**: interactive — export SAT, detect walls, detect dimples, ask user confirmation per sub-group
- **Pass 2**: verification — re-export SAT (fresh geometry after pass 1), re-detect from scratch, ask user confirmation for any remaining dimples
- If pass 2 finds nothing → verification passed

### 5. Fill
- Pick all dimple faces via AddToHistory, then RemoveSelectedFaces via AddToHistory
- All operations persist in CST history list for user revert (Ctrl+Z)

## Shield Can Frame Simplifier Algorithm

### 1. Wall Detection (Frame) — W-Height Filtered
- Find reference face (largest plane face by bbox area) → W axis = reference normal
- Compute component overall bbox → project onto W axis → component W height
- Find all plane faces perpendicular to reference normal (|dot| ≤ 0.05)
- **W-height filter**: wall W span must be ≥ 50% of component W height. Rejects tiny structural faces at corners.

### 2. Wall Sub-Grouping (same as cover)
- Copper thickness measurement from two largest parallel plane faces
- Group by normal (dot > 0.99)
- Recursive W-distance splitting (threshold = 1.5× copper thickness)
- UV overlap splitting for sub-groups with >2 faces

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

## Shield Can Component Pairing

Before processing, covers and frames are paired by bbox overlap to identify two-piece vs one-piece shield cans.

### Algorithm
1. Classify components by keyword: SHIELD+COVER → cover, SHIELD+FRAM → frame, SHIELD alone → one-piece
2. For each cover and frame, export SAT and compute overall bbox (union of all face bboxes)
3. For each cover, find the best matching frame by:
   - UV overlap > 50% in both axes (project bboxes onto the two larger dimensions)
   - W proximity: W gap ≤ 2× thicker component's W span (cover and frame must be stacked)
4. Paired → two-piece (cover + frame processed separately)
5. Unpaired → moved to one-piece category
6. Show classifier dialog for user to review/move/remove before processing

## Two-Pass Verification Workflow

Both cover and frame simplification use a two-pass workflow:
- **Pass 1**: Export SAT → detect walls → detect dimples → ask user confirmation → fill
- **Pass 2**: Re-export SAT (fresh geometry after pass 1) → re-detect from scratch → ask user confirmation for remaining dimples
- If pass 2 finds nothing → "Verification passed"

## Component Cleanup Tool

Separate GUI tool for removing plastic, rubber, screws, and other non-metal components from imported CAD models.

```bash
python -m code.gui_cleanup
```

### Features
- Three keyword categories: Delete (ask user), Auto-delete (no confirmation), Exclude (never delete)
- Default keywords: Delete=COVER, Auto-delete=SCREW/RUBBER/MYLAR/ADH/PLASTIC, Exclude=SHIELDING/SPRING
- Groups similar components (xxx, xxx_1, xxx_2) and processes as a batch
- Sorts by Zmax (topmost/outermost components first)
- SelectTreeItem highlights components in CST navigation tree
- Clickable buttons for adding exclude keywords (no typing needed)
- Import/Export/Save keyword lists for reuse across projects
- Auto-deletes empty parent folders after removing all children

## PCB Grounding Bridge Algorithm

Bridges gaps between the PCB and nearby components (e.g. shield can frames, connectors) to ensure electrical grounding contact.

### Algorithm
1. **Find PCB**: Search components by keywords (BOARD, PCB, MB, MAIN_BOARD). Validate flatness: both plane dimensions must be ≥4x the thickness.
2. **Establish UVW**: Find the longest straight edge from SAT topology. Build orthonormal UVW: U=edge direction, W=board normal (from largest plane face), V=W×U. Set WCS in CST for visualization.
3. **Find top/bottom faces**: Two large plane faces with normals along W, spanning ≥80% of PCB in U and V.
4. **Find nearby components**: For each side of the PCB, find components whose W_min is within 1/4 thickness of the PCB face, entirely on one side (not straddling).
5. **Check gaps**: For each nearby component, use copy+Solid.Intersect+volume check to confirm no overlap (gap exists).
6. **Bridge**: Find the closest parallel face on the component, highlight it with WCS at its center, ask user to confirm, then extrude toward the PCB face via AddToHistory.
7. **Repeat for both sides** of the PCB.

### Key Features
- Recursive component tree walker (handles any nesting depth)
- UVW coordinate system aligned with PCB edge (works for non-axis-aligned boards)
- WCS crosshair at mating face center for easy visual identification
- WCS restored to PCB center after bridging

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

## Connector Replacement Algorithm

Replaces connector components with PEC blocks and bridges them to FPC and PCB for electrical contact.

```bash
python -m code.debug_connector_v2
```

### Algorithm
1. **Select connector**: User provides name, fuzzy match, SelectTreeItem to confirm
2. **Analyze geometry**: Find longest straight edge → U axis. Find largest plane face → W axis (board normal). Build UVW coordinate system.
3. **Create replacement block**: PEC Brick with same global bbox as original connector via AddToHistory
4. **Find FPC**: Search by keywords (FPC, FLEX, FPCA), check W-axis proximity, user confirms. If no auto-match, manual name entry.
5. **Check FPC contact**: Find closest parallel FPC face, check if within block's W range. If gap: auto-extrude block face toward FPC.
6. **Find PCB**: Search by keywords (BOARD, PCB, MB) with 3 filters:
   - Largest plane face area >= 10x block's plane face area
   - Largest face normal parallel to W axis (|dot| >= 0.9)
   - W distance < connector thickness
7. **Check PCB contact**: Same W-axis method as FPC, auto-bridge if gap.
8. **Merge bridges**: `Solid.Add` merges FPC/PCB bridges into the main block.
9. **Delete original**: Remove original connector after user confirmation.
10. **Clean up overlapping components**: Find components at tree levels n-1 to n whose bbox is within the merged block's bbox. Delete with user confirmation. Repeat with manual input until no more components remain.

### Key Features
- UVW coordinate system handles non-axis-aligned connectors
- Three PCB filters eliminate false positives from small keyword-matching components
- Auto-bridge without confirmation (gaps are always small and unambiguous)
- >100 candidate handling: ask user to auto-detect or manually input
- Progress logging during long scans
- Solid.Add merges bridges into a single block

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
