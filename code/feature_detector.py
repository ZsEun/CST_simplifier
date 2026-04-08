"""Geometry-based hole detection for STP-imported CST models.

CST 2025 SAT export discovery:
- SAT `pid` attribute = CST face ID directly (no offset, confirmed)
- SAT face entity ORDER does NOT match CST face ID order
- We parse SAT topology (face→loop→coedge→edge) for adjacency
- Bounding boxes extracted from face entity T-field (reliable for seam edges)

Detection strategy:
1. Enumerate solids via Resulttree
2. Export SAT, parse face entities with pid mapping
3. Build topology-based adjacency graph
4. Extract per-face bounding boxes
5. Return seed faces (cone/cylinder) for progressive fill

Validates: Requirements 1.1, 1.2, 7.1, 7.4, 8.1
"""

import logging
import os
import re
import tempfile
from typing import Dict, List, Optional, Set, Tuple

from code.cst_connection import CSTConnection, CSTConnectionError
from code.models import FaceInfo, SimplificationCandidate

logger = logging.getLogger(__name__)


# ======================================================================
# SAT (ACIS) file parser
# ======================================================================

class SATParser:
    """Parse an ACIS SAT file to extract face info, adjacency, and bboxes.

    Key discovery: SAT `pid` attribute = CST face ID directly (no offset).
    """

    def __init__(self, sat_path: str):
        self._path = sat_path
        self._entities: List[str] = []
        self._face_indices: List[int] = []
        self._face_to_pid: Dict[int, int] = {}
        self._pid_to_ent: Dict[int, int] = {}

    def parse(self) -> Dict[int, dict]:
        """Parse SAT and return face info keyed by CST face ID (pid).

        Returns:
            Dict mapping CST face ID (from pid) to face info dict:
            - 'surface_type': str
            - 'geometry': dict with center, axis, radius (for cone/cylinder)
            - 'face_entity_idx': int
            - 'sat_order': int
        """
        with open(self._path, "r", errors="replace") as f:
            raw = f.read()

        all_lines = raw.split("\n")
        self._find_header(all_lines)
        self._find_faces()
        self._face_to_pid = self._extract_pids(self._face_indices)
        self._pid_to_ent = {pid: ent for ent, pid in self._face_to_pid.items()}

        faces: Dict[int, dict] = {}
        for sat_order, ent_idx in enumerate(self._face_indices):
            surface_type, geometry = self._get_surface_info(ent_idx)
            pid = self._face_to_pid.get(ent_idx)
            if pid is None:
                logger.warning("Face at entity %d has no pid, skipping", ent_idx)
                continue
            faces[pid] = {
                "surface_type": surface_type or "unknown",
                "geometry": geometry,
                "face_entity_idx": ent_idx,
                "sat_order": sat_order,
            }
        return faces

    def build_adjacency(self) -> Dict[int, Set[int]]:
        """Build face adjacency graph from SAT topology.

        Parses face → loop → coedge chain, then uses two strategies:
        1. edge → faces: two faces sharing the same edge are adjacent
        2. partner coedges: each coedge has a partner on the adjacent face;
           follow partner → loop → face to find the neighbor.

        Strategy 2 is critical for seam edges (360° surfaces like full cones)
        where the coedge chain is self-referencing (next=self, prev=self).
        Without it, such faces get 0 neighbors.

        SAT coedge $-refs order: $attr ... $next $prev $partner $edge $loop
        We identify partner as: a coedge ref that is NOT next, NOT prev,
        NOT the coedge itself, and NOT already the edge or loop ref.

        Must call parse() first.

        Returns:
            Dict mapping CST face ID (pid) → set of adjacent CST face IDs.
        """
        # face → loops
        face_loops: Dict[int, List[int]] = {}
        for fi in self._face_indices:
            refs = self._get_refs(self._entities[fi])
            loops = [r for r in refs
                     if 0 <= r < len(self._entities) and self._etype(r) == "loop"]
            face_loops[fi] = loops

        # loop → face (reverse map)
        loop_to_face: Dict[int, int] = {}
        for fi, loops in face_loops.items():
            for li in loops:
                loop_to_face[li] = fi

        # loop → coedges (circular linked list walk via next pointer)
        # SAT coedge refs: ..., $next, $prev, $partner, $edge, $loop
        # We walk via "next" only (first coedge ref that isn't self).
        # For seam edges, next=self so the chain has just 1 coedge.
        loop_coedges: Dict[int, List[int]] = {}
        for loops in face_loops.values():
            for li in loops:
                if li in loop_coedges:
                    continue
                loop_line = self._entities[li]
                refs = self._get_refs(loop_line)
                first_coedge = None
                for ref in refs:
                    if 0 <= ref < len(self._entities) and self._etype(ref) == "coedge":
                        first_coedge = ref
                        break
                if first_coedge is None:
                    loop_coedges[li] = []
                    continue
                coedges = [first_coedge]
                visited = {first_coedge}
                current = first_coedge
                while True:
                    ce_refs = self._get_refs(self._entities[current])
                    # Find next coedge: first coedge ref that isn't self
                    # and hasn't been visited
                    next_ce = None
                    for ref in ce_refs:
                        if (0 <= ref < len(self._entities)
                                and self._etype(ref) == "coedge"
                                and ref not in visited):
                            next_ce = ref
                            break
                    if next_ce is None or next_ce == first_coedge:
                        break
                    coedges.append(next_ce)
                    visited.add(next_ce)
                    current = next_ce
                loop_coedges[li] = coedges

        # Collect all coedges and map coedge → edge
        all_coedges: Set[int] = set()
        for ces in loop_coedges.values():
            all_coedges.update(ces)
        coedge_edge: Dict[int, int] = {}
        for ci in all_coedges:
            for ref in self._get_refs(self._entities[ci]):
                if 0 <= ref < len(self._entities) and self._etype(ref) == "edge":
                    coedge_edge[ci] = ref
                    break

        # edge → set of face entity indices
        edge_faces: Dict[int, Set[int]] = {}
        for fi, loops in face_loops.items():
            for li in loops:
                for ci in loop_coedges.get(li, []):
                    ei = coedge_edge.get(ci)
                    if ei is not None:
                        edge_faces.setdefault(ei, set()).add(fi)

        # --- Partner coedge adjacency (handles seam edges) ---
        # For each coedge, find its partner coedge. The partner lives on
        # the adjacent face's loop. SAT coedge format:
        #   coedge $attr ... $next $prev $partner $edge $loop
        # All coedge-type refs: next, prev, partner. We identify partner
        # as any coedge ref that is NOT in the same loop's coedge list.
        coedge_to_loop: Dict[int, int] = {}
        for li, ces in loop_coedges.items():
            for ci in ces:
                coedge_to_loop[ci] = li

        for fi, loops in face_loops.items():
            for li in loops:
                for ci in loop_coedges.get(li, []):
                    ce_refs = self._get_refs(self._entities[ci])
                    # Find partner: a coedge ref NOT in this loop
                    for ref in ce_refs:
                        if (0 <= ref < len(self._entities)
                                and self._etype(ref) == "coedge"
                                and ref != ci):
                            # Check if this coedge is in a different loop
                            partner_loop = coedge_to_loop.get(ref)
                            if partner_loop is not None and partner_loop != li:
                                # Partner is in a different loop → different face
                                partner_face = loop_to_face.get(partner_loop)
                                if partner_face is not None:
                                    ei = coedge_edge.get(ci)
                                    if ei is not None:
                                        edge_faces.setdefault(ei, set()).add(fi)
                                        edge_faces[ei].add(partner_face)
                            elif partner_loop is None and ref not in all_coedges:
                                # Partner coedge wasn't found by chain walk
                                # (happens with seam edges). Scan its loop ref
                                # to find which face it belongs to.
                                partner_refs = self._get_refs(self._entities[ref])
                                for pr in partner_refs:
                                    if (0 <= pr < len(self._entities)
                                            and self._etype(pr) == "loop"
                                            and pr in loop_to_face):
                                        partner_face = loop_to_face[pr]
                                        if partner_face != fi:
                                            ei = coedge_edge.get(ci)
                                            if ei is not None:
                                                edge_faces.setdefault(ei, set()).add(fi)
                                                edge_faces[ei].add(partner_face)
                                        break

        # Build adjacency by pid
        adjacency: Dict[int, Set[int]] = {}
        for face_ents in edge_faces.values():
            pids = [self._face_to_pid[fe] for fe in face_ents if fe in self._face_to_pid]
            for p1 in pids:
                for p2 in pids:
                    if p1 != p2:
                        adjacency.setdefault(p1, set()).add(p2)

        return adjacency

    def get_bounding_boxes(self) -> Dict[int, Tuple[float, float, float, float, float, float]]:
        """Extract bounding box per face from the SAT face entity T-field.

        The face entity line contains an embedded bounding box after the
        'T' marker:  face ... T min_x min_y min_z max_x max_y max_z F #

        This is far more reliable than vertex-based bbox computation,
        especially for seam-edge 360° faces (full cones/cylinders) where
        the coedge chain yields only a single vertex.

        Must call parse() first.

        Returns:
            Dict mapping CST face ID (pid) → (min_x, min_y, min_z, max_x, max_y, max_z).
        """
        bboxes: Dict[int, Tuple[float, float, float, float, float, float]] = {}
        for fi in self._face_indices:
            pid = self._face_to_pid.get(fi)
            if pid is None:
                continue
            m = re.search(
                r'\bT\s+([-\d.eE+]+)\s+([-\d.eE+]+)\s+([-\d.eE+]+)\s+'
                r'([-\d.eE+]+)\s+([-\d.eE+]+)\s+([-\d.eE+]+)\s+F',
                self._entities[fi],
            )
            if m:
                bboxes[pid] = tuple(float(m.group(i)) for i in range(1, 7))
        return bboxes

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _find_header(self, lines: List[str]):
        """Find where entity data begins and store entities.

        Handles multi-line entities (e.g. spline-surface with NURBS data
        on continuation lines starting with tab). Continuation lines are
        joined with their parent entity so that $N references resolve
        to the correct entity index.
        """
        hdr = 0
        entity_starts = (
            "body ", "lump ", "shell ", "face ", "loop ",
            "coedge ", "edge ", "vertex ", "point ",
            "integer_attrib", "name_attrib", "rgb_color",
            "simple-snl", "cstishape",
            "cone-surface ", "plane-surface ", "spline-surface ",
            "torus-surface ", "sphere-surface ", "ellipse-curve ",
            "straight-curve ", "intcurve-curve ", "pcurve ",
            "cone ", "null",
        )
        for i, line in enumerate(lines):
            s = line.strip()
            if any(s.startswith(t) for t in entity_starts):
                hdr = i
                break

        # Join continuation lines with their parent entity
        entities = []
        for line in lines[hdr:]:
            stripped = line.strip()
            if not stripped:
                continue
            # A continuation line starts with tab, or is a data line that
            # doesn't start with a known entity keyword
            if line.startswith("\t"):
                if entities:
                    entities[-1] = entities[-1] + " " + stripped
                continue
            if entities and not any(stripped.startswith(t) for t in entity_starts):
                # Check if previous entity is complete (ends with #)
                if entities and not entities[-1].rstrip().endswith("#"):
                    entities[-1] = entities[-1] + " " + stripped
                    continue
            entities.append(line)

        self._entities = entities

    def _find_faces(self):
        """Find all face entity indices."""
        self._face_indices = [
            i for i, line in enumerate(self._entities)
            if line.startswith("face ")
        ]

    def _get_refs(self, line: str) -> List[int]:
        """Extract all $N references from an entity line."""
        return [int(m) for m in re.findall(r"\$(-?\d+)", line)]

    def _etype(self, idx: int) -> str:
        """Get entity type (first word) at index."""
        if 0 <= idx < len(self._entities):
            line = self._entities[idx]
            return line.split()[0] if line.strip() else ""
        return ""

    def _extract_pids(self, face_indices: List[int]) -> Dict[int, int]:
        """Extract pid attribute for each face entity.

        Four-pass strategy:
        (1) scan pid attribs for direct face refs
        (2) walk attribute chain from face's first ref
        (3) walk ALL refs from face (not just first) looking for attribs
        (4) reverse search: for each unassigned pid attrib, walk its
            ref chain looking for entities that share refs with orphan faces
        """
        face_to_pid: Dict[int, int] = {}
        face_set = set(face_indices)

        # Pass 1: pid attribs referencing a face directly
        for i, line in enumerate(self._entities):
            m = re.search(r'@3 pid (\d+)', line)
            if m:
                pid_val = int(m.group(1))
                for ref in self._get_refs(line):
                    if ref in face_set:
                        face_to_pid[ref] = pid_val
                        break

        # Pass 2: walk attribute chain from face's first ref
        for fi in face_indices:
            if fi in face_to_pid:
                continue
            refs = self._get_refs(self._entities[fi])
            if refs and refs[0] >= 0:
                attr_idx = refs[0]
                visited: set = set()
                while 0 <= attr_idx < len(self._entities) and attr_idx not in visited:
                    visited.add(attr_idx)
                    attr_line = self._entities[attr_idx]
                    m = re.search(r'@3 pid (\d+)', attr_line)
                    if m:
                        face_to_pid[fi] = int(m.group(1))
                        break
                    attr_refs = self._get_refs(attr_line)
                    next_attr = -1
                    for ar in attr_refs:
                        if ar >= 0 and ar != fi and ar not in visited:
                            if 0 <= ar < len(self._entities) and "attrib" in self._etype(ar):
                                next_attr = ar
                                break
                    attr_idx = next_attr if next_attr >= 0 else -1

        # Pass 3: for faces still missing, walk ALL refs (not just first)
        # and follow any attrib chains found
        for fi in face_indices:
            if fi in face_to_pid:
                continue
            refs = self._get_refs(self._entities[fi])
            for ref in refs:
                if ref < 0 or ref >= len(self._entities):
                    continue
                # Check if this ref itself is a pid attrib
                ref_line = self._entities[ref]
                m = re.search(r'@3 pid (\d+)', ref_line)
                if m:
                    face_to_pid[fi] = int(m.group(1))
                    break
                # Walk attrib chain from this ref
                if "attrib" in self._etype(ref):
                    attr_idx = ref
                    visited = {fi}
                    while 0 <= attr_idx < len(self._entities) and attr_idx not in visited:
                        visited.add(attr_idx)
                        attr_line = self._entities[attr_idx]
                        m = re.search(r'@3 pid (\d+)', attr_line)
                        if m:
                            face_to_pid[fi] = int(m.group(1))
                            break
                        attr_refs = self._get_refs(attr_line)
                        next_attr = -1
                        for ar in attr_refs:
                            if ar >= 0 and ar not in visited:
                                if 0 <= ar < len(self._entities) and "attrib" in self._etype(ar):
                                    next_attr = ar
                                    break
                        attr_idx = next_attr if next_attr >= 0 else -1
                    if fi in face_to_pid:
                        break

        # Pass 4: for STILL missing faces, try to match by finding pid
        # attribs whose ref chain leads to an entity that is also
        # referenced by the orphan face (shared loop, edge, etc.)
        missing = [fi for fi in face_indices if fi not in face_to_pid]
        if missing:
            assigned_pids = set(face_to_pid.values())
            # Collect all pid attribs not yet assigned
            unassigned_pid_attribs = []
            for i, line in enumerate(self._entities):
                m = re.search(r'@3 pid (\d+)', line)
                if m:
                    pid_val = int(m.group(1))
                    if pid_val not in assigned_pids:
                        # Only consider pid attribs that reference a face
                        refs = self._get_refs(line)
                        for ref in refs:
                            if 0 <= ref < len(self._entities) and self._etype(ref) == "face":
                                unassigned_pid_attribs.append((i, pid_val, ref))
                                break

            missing_set = set(missing)
            for _, pid_val, face_ref in unassigned_pid_attribs:
                if face_ref in missing_set and face_ref not in face_to_pid:
                    face_to_pid[face_ref] = pid_val

        return face_to_pid

    def _get_surface_info(self, face_idx: int) -> Tuple[Optional[str], dict]:
        """Get surface type and geometry for a face entity.

        Strategy:
        1. Check direct $refs from the face entity for *-surface types
        2. Check refs-of-refs (one level deeper) — handles cases where
           the surface is referenced through a loop or other entity
        3. Scan the face line text for surface keywords as last resort
        """
        face_line = self._entities[face_idx]
        refs = self._get_refs(face_line)

        # Pass 1: direct refs
        for ref in refs:
            if 0 <= ref < len(self._entities):
                first_word = self._etype(ref)
                if first_word.endswith("-surface"):
                    return self._parse_surface_entity(ref, first_word)

        # Pass 2: one level deeper — check refs of each ref
        # This handles cases where the face entity doesn't directly
        # reference its surface (seen in some CST SAT exports where
        # the last ref is an attrib instead of a surface)
        for ref in refs:
            if 0 <= ref < len(self._entities):
                sub_refs = self._get_refs(self._entities[ref])
                for sr in sub_refs:
                    if 0 <= sr < len(self._entities):
                        first_word = self._etype(sr)
                        if first_word.endswith("-surface"):
                            return self._parse_surface_entity(sr, first_word)

        # Pass 3: scan nearby entities. In some SAT files the surface
        # entity is at face_idx+1 or face_idx+2 (positional convention)
        for offset in [1, 2, 3]:
            idx = face_idx + offset
            if 0 <= idx < len(self._entities):
                first_word = self._etype(idx)
                if first_word.endswith("-surface"):
                    return self._parse_surface_entity(idx, first_word)
                # Stop if we hit another face or non-related entity
                if first_word in ("face", "body", "lump", "shell"):
                    break

        return None, {}

    def _parse_surface_entity(self, idx: int, surface_type: str) -> Tuple[str, dict]:
        """Parse a surface entity and return (type, geometry)."""
        geometry = {}
        if surface_type == "cone-surface":
            geometry = self._parse_cone_surface(self._entities[idx])
        elif surface_type == "plane-surface":
            geometry = self._parse_plane_surface(self._entities[idx])
        return surface_type, geometry

    def _parse_cone_surface(self, line: str) -> dict:
        """Parse cone-surface entity for geometry."""
        cleaned = re.sub(r'\$-?\d+', ' ', line)
        cleaned = re.sub(r'^cone-surface\s*', '', cleaned)
        cleaned = re.sub(r'[#IFT]', ' ', cleaned)
        cleaned = re.sub(r'\b(reversed|forward_v|forward)\b', ' ', cleaned)
        nums = []
        for t in cleaned.split():
            try:
                nums.append(float(t))
            except ValueError:
                continue
        if len(nums) < 10:
            return {}
        off = 2
        return {
            "center": (nums[off], nums[off+1], nums[off+2]),
            "axis": (nums[off+3], nums[off+4], nums[off+5]),
            "cos_half_angle": nums[off+6],
            "sin_half_angle": nums[off+7],
            "radius": next((n for n in nums[off+8:] if n > 0), 0.0),
        }

    def _parse_plane_surface(self, line: str) -> dict:
        """Parse plane-surface entity for normal and position."""
        cleaned = re.sub(r'\$-?\d+', ' ', line)
        cleaned = re.sub(r'^plane-surface\s*', '', cleaned)
        cleaned = re.sub(r'[#IFT]', ' ', cleaned)
        cleaned = re.sub(r'\b(reversed|forward_v|forward)\b', ' ', cleaned)
        nums = []
        for t in cleaned.split():
            try:
                nums.append(float(t))
            except ValueError:
                continue
        if len(nums) >= 8:
            off = 2
            return {
                "center": (nums[off], nums[off+1], nums[off+2]),
                "normal": (nums[off+3], nums[off+4], nums[off+5]),
            }
        return {}


# ======================================================================
# Feature detector
# ======================================================================

class FeatureDetector:
    """Detects hole-like features in a CST model.

    Uses SAT export + pid mapping + topology-based adjacency.

    Args:
        connection: An active CSTConnection with an open project.
    """

    def __init__(self, connection: CSTConnection) -> None:
        self._conn = connection

    def _enumerate_solids(self) -> List[Tuple[str, str]]:
        """List every (component_name, solid_name) pair in the model.

        Uses recursive tree walker (depth limit 10) to handle any nesting depth.
        """
        out_path = os.path.join(tempfile.gettempdir(), "cst_solids.txt")
        vba_path = out_path.replace("\\", "\\\\")
        macro = (
            'Sub Main\n'
            f'  Open "{vba_path}" For Output As #1\n'
            '  Dim rt As Object\n'
            '  Set rt = Resulttree\n'
            '  Call WalkTree(rt, "Components", 1)\n'
            '  Close #1\n'
            'End Sub\n'
            '\n'
            'Sub WalkTree(rt As Object, path As String, depth As Integer)\n'
            '  If depth > 10 Then Exit Sub\n'
            '  Dim child As String\n'
            '  child = rt.GetFirstChildName(path)\n'
            '  Do While child <> ""\n'
            '    Dim subChild As String\n'
            '    subChild = rt.GetFirstChildName(child)\n'
            '    If subChild = "" Then\n'
            '      Print #1, child\n'
            '    Else\n'
            '      Call WalkTree(rt, child, depth + 1)\n'
            '    End If\n'
            '    child = rt.GetNextItemName(child)\n'
            '  Loop\n'
            'End Sub\n'
        )
        result = self._conn.execute_vba(macro, output_file=out_path)
        solids: List[Tuple[str, str]] = []
        if result:
            for line in result.split("\n"):
                line = line.strip()
                if not line:
                    continue
                path = line.replace("\\", "/")
                if path.startswith("Components/"):
                    path = path[len("Components/"):]
                parts = path.split("/")
                if len(parts) >= 2:
                    solid = parts[-1]
                    comp_path = "/".join(parts[:-1])
                    solids.append((comp_path, solid))
        try:
            os.remove(out_path)
        except OSError:
            pass
        logger.info("Enumerated %d solid shapes.", len(solids))
        return solids


    def _export_sat(self, component: str, solid: str) -> Optional[str]:
        """Export solid to SAT file and return path, or None on failure."""
        sat_path = os.path.join(tempfile.gettempdir(), "cst_export.sat")
        sat_vba = sat_path.replace("\\", "\\\\")
        full_name = f"{component}:{solid}"
        macro = (
            'Sub Main\n'
            '  With SAT\n'
            '    .Reset\n'
            f'    .FileName "{sat_vba}"\n'
            f'    .Write "{full_name}"\n'
            '  End With\n'
            'End Sub\n'
        )
        try:
            self._conn.execute_vba(macro)
            if os.path.isfile(sat_path):
                return sat_path
        except Exception as exc:
            logger.warning("SAT export failed for %s: %s", full_name, exc)
        return None

    def analyze_solid(
        self, component: str, solid: str
    ) -> Optional[Tuple[Dict[int, dict], Dict[int, Set[int]],
                         Dict[int, Tuple[float, float, float, float, float, float]]]]:
        """Export SAT and return (faces, adjacency, bboxes) for a solid.

        Returns None if export/parse fails.
        """
        sat_path = self._export_sat(component, solid)
        if not sat_path:
            return None
        try:
            parser = SATParser(sat_path)
            faces = parser.parse()
            adjacency = parser.build_adjacency()
            bboxes = parser.get_bounding_boxes()
            return faces, adjacency, bboxes
        except Exception as exc:
            logger.warning("SAT parse failed for %s:%s: %s", component, solid, exc)
            return None
        finally:
            try:
                os.remove(sat_path)
            except OSError:
                pass

    def _filter_edge_fillets(
        self,
        seeds: List[int],
        faces: Dict[int, dict],
        face_types: Dict[int, str],
        adjacency: Dict[int, Set[int]],
        bboxes: Dict[int, Tuple],
        board_edge_ratio: float = 0.5,
    ) -> List[int]:
        """Filter out PCB board-edge cone/cylinder faces using board-normal-aware
        loop-size comparison.

        Algorithm:
        1. Find the board reference faces — the 2 plane faces with the
           largest bounding box area (PCB top/bottom).
        2. Determine the board normal axis from the reference face's
           surface normal. The 2 remaining axes are the in-plane axes.
        3. For each cone seed, BFS-walk adjacency (excluding board ref
           faces) to build the connected "feature loop".
        4. Compute the union bounding box of the feature loop.
        5. Compare loop extent vs board extent on the 2 in-plane axes
           only (skip the board normal axis). If the loop spans ≥
           board_edge_ratio of the board in BOTH in-plane axes →
           board edge → filter out. Otherwise → hole → keep.

        Args:
            seeds: Cone/cylinder face IDs to filter.
            faces: Full face info dict from SAT parse.
            face_types: pid → surface type string.
            adjacency: Face adjacency graph.
            bboxes: Per-face bounding boxes (from T-field).
            board_edge_ratio: Threshold ratio (0-1).

        Returns:
            Filtered list of seed face IDs (board-edge cones removed).
        """
        # --- Step 1: Find board reference faces (largest plane faces) ---
        plane_faces = [pid for pid, st in face_types.items() if "plane" in st]
        if not plane_faces:
            return seeds

        def _bbox_area(pid: int) -> float:
            bb = bboxes.get(pid)
            if not bb:
                return 0.0
            dx = bb[3] - bb[0]
            dy = bb[4] - bb[1]
            dz = bb[5] - bb[2]
            dims = sorted([dx, dy, dz], reverse=True)
            return dims[0] * dims[1]

        plane_by_area = sorted(plane_faces, key=_bbox_area, reverse=True)
        board_ref_faces = set(plane_by_area[:2])

        # Board reference bbox = union of the top 2 plane face bboxes
        board_bb = None
        for pid in board_ref_faces:
            bb = bboxes.get(pid)
            if bb is None:
                continue
            if board_bb is None:
                board_bb = list(bb)
            else:
                for i in range(3):
                    board_bb[i] = min(board_bb[i], bb[i])
                for i in range(3, 6):
                    board_bb[i] = max(board_bb[i], bb[i])

        if board_bb is None:
            return seeds

        board_extents = [board_bb[3] - board_bb[0],
                         board_bb[4] - board_bb[1],
                         board_bb[5] - board_bb[2]]

        # --- Step 2: Determine board normal axis ---
        board_normal_axis = None  # 0=X, 1=Y, 2=Z
        for ref_pid in sorted(board_ref_faces):
            ref_geo = faces.get(ref_pid, {}).get("geometry", {})
            normal = ref_geo.get("normal")
            if normal:
                abs_n = [abs(normal[0]), abs(normal[1]), abs(normal[2])]
                board_normal_axis = abs_n.index(max(abs_n))
                break

        # Fallback: if no normal found, use the axis with smallest board extent
        # (the board thickness direction)
        if board_normal_axis is None:
            board_normal_axis = board_extents.index(min(board_extents))

        in_plane_axes = [i for i in range(3) if i != board_normal_axis]
        axis_names = ["X", "Y", "Z"]
        logger.info("  Board ref faces: %s, normal axis: %s, in-plane: %s",
                     sorted(board_ref_faces), axis_names[board_normal_axis],
                     [axis_names[i] for i in in_plane_axes])
        logger.info("  Board extents: %.1f x %.1f x %.1f",
                     board_extents[0], board_extents[1], board_extents[2])

        # --- Step 3-5: For each seed, build feature loop and compare ---
        filtered = []
        for pid in seeds:
            # BFS from seed, excluding board reference faces
            loop = set()
            queue = [pid]
            while queue:
                current = queue.pop()
                if current in loop or current in board_ref_faces:
                    continue
                loop.add(current)
                for neighbor in adjacency.get(current, set()):
                    if neighbor not in loop and neighbor not in board_ref_faces:
                        queue.append(neighbor)

            # Compute union bbox of the loop
            loop_bb = None
            for fid in loop:
                bb = bboxes.get(fid)
                if bb is None:
                    continue
                if loop_bb is None:
                    loop_bb = list(bb)
                else:
                    for i in range(3):
                        loop_bb[i] = min(loop_bb[i], bb[i])
                    for i in range(3, 6):
                        loop_bb[i] = max(loop_bb[i], bb[i])

            if loop_bb is None:
                filtered.append(pid)
                continue

            loop_extents = [loop_bb[3] - loop_bb[0],
                            loop_bb[4] - loop_bb[1],
                            loop_bb[5] - loop_bb[2]]

            # Compare only on in-plane axes (skip board normal axis)
            in_plane_large = 0
            for ax in in_plane_axes:
                if board_extents[ax] > 0 and loop_extents[ax] / board_extents[ax] >= board_edge_ratio:
                    in_plane_large += 1

            logger.info("  Seed %d: loop=%s, extents=%.1f x %.1f x %.1f, "
                        "in_plane_large=%d",
                        pid, sorted(loop),
                        loop_extents[0], loop_extents[1], loop_extents[2],
                        in_plane_large)

            if in_plane_large >= 2:
                logger.info("    -> FILTERED (board edge)")
                continue

            filtered.append(pid)

        return filtered


    def _group_seeds_by_loop(
        self,
        seeds: List[int],
        face_types: Dict[int, str],
        adjacency: Dict[int, Set[int]],
        bboxes: Dict[int, Tuple],
    ) -> List[dict]:
        """Group seed faces by hole and return full loop faces for each.

        Uses BFS from each seed, walking adjacency excluding board reference
        faces, to build the full connected loop. ALL faces in the loop
        (not just cone seeds) must be picked for RemoveSelectedFaces to work.

        Args:
            seeds: Filtered cone/cylinder face IDs.
            face_types: pid → surface type string.
            adjacency: Face adjacency graph.
            bboxes: Per-face bounding boxes.

        Returns:
            List of dicts, each with:
            - 'seeds': sorted list of cone seed face IDs in this hole
            - 'loop_faces': sorted list of ALL face IDs forming the hole
              (includes seeds + non-cone faces like planes between cones)
        """
        if not seeds:
            return []

        # Find board reference faces (same logic as _filter_edge_fillets)
        plane_faces = [pid for pid, st in face_types.items() if "plane" in st]
        board_ref_faces: Set[int] = set()
        if plane_faces:
            def _bbox_area(pid: int) -> float:
                bb = bboxes.get(pid)
                if not bb:
                    return 0.0
                dx = bb[3] - bb[0]
                dy = bb[4] - bb[1]
                dz = bb[5] - bb[2]
                dims = sorted([dx, dy, dz], reverse=True)
                return dims[0] * dims[1]
            plane_by_area = sorted(plane_faces, key=_bbox_area, reverse=True)
            board_ref_faces = set(plane_by_area[:2])

        groups: List[dict] = []
        assigned: Set[int] = set()

        for seed in sorted(seeds):
            if seed in assigned:
                continue

            # BFS from this seed, excluding board ref faces
            loop = set()
            queue = [seed]
            while queue:
                current = queue.pop()
                if current in loop or current in board_ref_faces:
                    continue
                loop.add(current)
                for neighbor in adjacency.get(current, set()):
                    if neighbor not in loop and neighbor not in board_ref_faces:
                        queue.append(neighbor)

            # Find which seeds are in this loop
            group_seeds = sorted(s for s in seeds if s in loop)
            groups.append({
                "seeds": group_seeds,
                "loop_faces": sorted(loop),
            })
            for s in group_seeds:
                assigned.add(s)

        logger.info("  Grouped %d seeds into %d hole groups", len(seeds), len(groups))
        for i, g in enumerate(groups):
            logger.info("    Group %d: seeds=%s, loop=%s",
                        i, g["seeds"], g["loop_faces"])

        return groups

    def detect_seeds(self) -> List[dict]:
        """Detect cone/cylinder seed faces across all solids.

        Pipeline:
        1. Find all cone-surface faces (screw hole walls)
        2. Filter out board-edge cones using the board edge algorithm
        3. Group remaining seeds by hole

        Returns a list of dicts, one per solid, each containing:
        - 'component': str
        - 'solid': str
        - 'shape_name': str (component:solid)
        - 'seeds': List[int] (CST face IDs of cone/cylinder faces)
        - 'seed_groups': List[dict] (seeds grouped by hole)
        - 'face_types': Dict[int, str] (pid → surface type)
        - 'adjacency': Dict[int, Set[int]]
        - 'bboxes': Dict[int, tuple]
        """
        solids = self._enumerate_solids()
        if not solids:
            logger.info("No solids found.")
            return []

        results = []
        for component, solid in solids:
            logger.info("Analyzing %s:%s ...", component, solid)
            analysis = self.analyze_solid(component, solid)
            if not analysis:
                continue

            faces, adjacency, bboxes = analysis
            face_types = {pid: info["surface_type"] for pid, info in faces.items()}

            # Seeds: all cone-surface faces (screw hole walls)
            seeds = [pid for pid, st in face_types.items() if "cone" in st]

            # Filter out PCB board-edge cones using loop-size-vs-board-size
            seeds = self._filter_edge_fillets(
                seeds, faces, face_types, adjacency, bboxes
            )

            if seeds:
                # Group seeds by hole (seeds sharing the same loop)
                seed_groups = self._group_seeds_by_loop(
                    seeds, face_types, adjacency, bboxes
                )

                results.append({
                    "component": component,
                    "solid": solid,
                    "shape_name": f"{component}:{solid}",
                    "seeds": sorted(seeds),
                    "seed_groups": seed_groups,
                    "face_types": face_types,
                    "adjacency": adjacency,
                    "bboxes": bboxes,
                })
                logger.info("  Found %d seed faces in %d groups",
                            len(seeds), len(seed_groups))

        return results
