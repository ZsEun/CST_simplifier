CST CAD Model Simplifier
========================

Automates detection and filling of screw holes in STP-imported CAD models
within CST Studio Suite 2025. Connects via COM automation, exports SAT
geometry, parses topology to find cone-surface hole features, filters out
board-edge fillets, groups hole faces, and removes them using
AddToHistory-based VBA commands.


REQUIREMENTS
------------
- Windows (COM automation required)
- CST Studio Suite 2025 (must be installed and COM registered)
- Python 3.10+
- pywin32  (pip install pywin32)


INSTALLATION
------------
    pip install pywin32


PROJECT STRUCTURE
-----------------
code/
    __init__.py          - Package init
    __main__.py          - Entry point for "python -m code"
    models.py            - Data classes (FaceInfo, SessionSummary, etc.)
    cst_connection.py    - COM connection to CST Studio Suite 2025
    feature_detector.py  - SAT parser + hole detection + board edge filter
    simplifier.py        - Progressive hole filling via AddToHistory
    main.py              - CLI entry point with argparse
    run_sunray_v3.py     - Standalone runner for Sunray_MB_v3 model
    requirements.txt     - Python dependencies


HOW TO RUN
----------
1. Open CST Studio Suite 2025 and load your .cst project.

2. Generic CLI usage:

       python -m code --project "C:\path\to\your\model.cst"

   Add --auto for non-interactive mode (fills all holes without asking):

       python -m code --project "C:\path\to\your\model.cst" --auto

3. For the Sunray_MB_v3 model specifically:

       python -m code.run_sunray_v3

   This runs in interactive mode, asking y/n/q for each detected hole.


CODE LOGIC
----------

1. CONNECT TO CST
   - Uses win32com to connect to a running CST instance via
     GetActiveObject("CSTStudio.Application")
   - Opens the .cst project file using raw _oleobj_.Invoke (required
     because CST's OpenFile returns an int that confuses pywin32)
   - Gets the project COM object via app.Active3D()

2. ENUMERATE SOLIDS
   - Runs a VBA macro via RunScript that walks the Resulttree to list
     all component/solid pairs
   - VBA writes results to a temp file, Python reads it back

3. EXPORT SAT
   - For each solid, exports an ACIS SAT file using CST's SAT.Write VBA
   - SAT file contains full B-rep topology with face entities

4. PARSE SAT FILE
   - Extracts face entities and their pid attributes (pid = CST face ID)
   - Identifies surface types: cone-surface, plane-surface, etc.
   - Builds topology-based adjacency graph:
     face -> loop -> coedge chain -> edge -> partner coedge -> neighbor face
   - Handles seam edges (360-degree surfaces like full cones) via partner
     coedge detection
   - Extracts per-face bounding boxes from the T-field in face entity lines

5. DETECT SEED FACES
   - Seeds = all cone-surface faces (screw hole walls are cone geometry)
   - No radius filtering; board edge filter is the only gate

6. BOARD EDGE FILTER (removes board-edge cones, keeps hole cones)
   Algorithm:
   a. Find the 2 plane faces with the largest bounding box area
      (these are the PCB top and bottom)
   b. Get the board normal direction from the plane surface normal
      (e.g., Z axis)
   c. The 2 remaining axes are the in-plane axes (e.g., X and Y)
   d. For each cone seed, BFS-walk adjacency excluding the 2 board
      reference faces to build the connected feature loop
   e. Compute the union bounding box of all faces in the loop
   f. Compare loop extent vs board extent on the 2 in-plane axes only
   g. If the loop spans >= 50% of the board in BOTH in-plane axes,
      it is a board edge fillet -> filter out
      Otherwise it is a hole -> keep

7. GROUP SEEDS BY HOLE
   - BFS from each seed (excluding board ref faces) to find all connected
     faces forming the hole
   - Seeds sharing the same connected loop are grouped together
   - Each group contains:
     * seeds: the cone face IDs (used for identification)
     * loop_faces: ALL face IDs forming the hole (cones + planes between
       them). All loop_faces must be picked for RemoveSelectedFaces to work.

8. FILL HOLES (Interactive or Auto)
   For each hole group:
   a. Highlight all loop faces in CST GUI (pick + zoom)
   b. In interactive mode: ask user y/n/q
   c. Pick each face via AddToHistory (required for persistence)
   d. Call LocalModification.RemoveSelectedFaces via AddToHistory
   e. Clear picks via AddToHistory
   f. On failure: expand face set with adjacent faces and retry
      (up to 5 expansion attempts)
   g. Skip groups whose faces were already consumed by previous fills


CST 2025 COM API NOTES
-----------------------
What works:
- win32com.client.GetActiveObject("CSTStudio.Application")
- app._oleobj_.Invoke(...) via call_method() for OpenFile, RunScript
- app.Active3D() to get project COM object
- VBA file I/O (Open ... For Output As #1) for data exchange
- Pick.PickFaceFromId "component:shape", "faceId" (face ID must be STRING)
- LocalModification.Reset/.Name/.RemoveSelectedFaces
- AddToHistory "caption", "vba_code" (REQUIRED for changes to persist)
- SAT.Reset/.FileName/.Write for ACIS SAT export

What does NOT work:
- proj.RunVBA(code) -- does not exist
- Solid.GetNameOfAllShapes(), Solid.GetNumberOfFaces(), etc. -- all fail
- Modeler.Undo, proj.Undo(), app.Undo() -- none work via COM
- RunScript alone does NOT persist model changes (AddToHistory is required)


SAT PARSING NOTES
-----------------
- SAT pid attribute = CST face ID directly (no offset needed)
- SAT face entity ORDER does not match CST face ID order
- Face entity T-field contains accurate bounding box:
  face ... T min_x min_y min_z max_x max_y max_z F #
- Topology: face -> loop -> coedge -> edge, partner coedges for adjacency
