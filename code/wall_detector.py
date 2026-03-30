"""Adjacency-based side wall discovery for shield can components.

Uses the existing SAT adjacency graph and face normals to find side walls:
1. Find top face (largest plane face)
2. Find curved faces adjacent to top face (rounded corners connecting top to walls)
3. Find plane faces adjacent to those curved faces
4. Filter by normal perpendicular to top face normal → side walls

No SAT topology walking needed — uses only data already provided by FeatureDetector.

Implements Requirements 1-3 of the shield-can-simplifier spec.
"""

import math
import logging
from dataclasses import dataclass, field
from typing import Dict, List, Set, Tuple

logger = logging.getLogger(__name__)


@dataclass
class WallInfo:
    """A discovered side wall face."""
    face_pid: int
    normal: Tuple[float, float, float]
    bbox: Tuple[float, float, float, float, float, float] = (0, 0, 0, 0, 0, 0)


def _normalize(v: Tuple[float, float, float]) -> Tuple[float, float, float]:
    mag = math.sqrt(v[0]**2 + v[1]**2 + v[2]**2)
    if mag < 1e-12:
        return (0.0, 0.0, 0.0)
    return (v[0]/mag, v[1]/mag, v[2]/mag)


def _dot(a: Tuple[float, float, float],
         b: Tuple[float, float, float]) -> float:
    return a[0]*b[0] + a[1]*b[1] + a[2]*b[2]


class WallDetector:
    """Discovers side walls of a shield can component via adjacency.

    Uses only data from FeatureDetector: face types, normals, adjacency, bboxes.
    No SAT file parsing needed.
    """

    def find_top_face(self, face_data: Dict[int, dict],
                      bboxes: Dict[int, Tuple]) -> Tuple[int, Tuple[float, float, float]]:
        """Find the largest plane face (top cover) and its normal.

        Returns: (top_face_pid, top_face_normal)
        """
        best_pid = -1
        best_area = 0.0
        best_normal = (0.0, 0.0, 1.0)

        for pid, info in face_data.items():
            if info["surface_type"] != "plane-surface":
                continue
            bb = bboxes.get(pid)
            if bb is None:
                continue
            dx = bb[3] - bb[0]
            dy = bb[4] - bb[1]
            dz = bb[5] - bb[2]
            dims = sorted([dx, dy, dz], reverse=True)
            area = dims[0] * dims[1]
            if area > best_area:
                best_area = area
                best_pid = pid
                geom = info.get("geometry", {})
                normal = geom.get("normal", (0.0, 0.0, 1.0))
                best_normal = _normalize(normal)

        logger.info("Top face: pid=%d, area=%.1f, normal=%s",
                     best_pid, best_area, best_normal)
        return best_pid, best_normal


    def discover_side_walls(
        self,
        top_face_pid: int,
        top_normal: Tuple[float, float, float],
        face_data: Dict[int, dict],
        adjacency: Dict[int, Set[int]],
        bboxes: Dict[int, Tuple],
        perpendicular_threshold: float = 0.05,
    ) -> List[WallInfo]:
        """Find side wall faces via adjacency through curved corner faces.

        Algorithm:
        1. Find all non-plane faces adjacent to the top face (curved corners)
        2. Find all plane faces adjacent to those curved faces
        3. Keep only those whose normal is perpendicular to top normal

        Args:
            top_face_pid: pid of the top face
            top_normal: normal vector of the top face
            face_data: pid → {surface_type, geometry}
            adjacency: pid → set of adjacent pids
            bboxes: pid → bbox tuple
            perpendicular_threshold: max |dot(wall_normal, top_normal)| to
                consider perpendicular (0 = perfectly perpendicular)

        Returns:
            List of WallInfo for discovered side walls.
        """
        # Step 1: find curved faces adjacent to top face
        top_neighbors = adjacency.get(top_face_pid, set())
        curved_faces = set()
        for npid in top_neighbors:
            info = face_data.get(npid)
            if info is None:
                continue
            stype = info["surface_type"]
            # Non-plane faces are the curved corners (torus, spline, etc.)
            if "plane" not in stype:
                curved_faces.add(npid)

        logger.info("Step 1: %d curved faces adjacent to top face %d: %s",
                     len(curved_faces), top_face_pid, sorted(curved_faces))

        # Step 2: find plane faces adjacent to the curved corner faces
        candidate_walls = set()
        for cfid in curved_faces:
            for npid in adjacency.get(cfid, set()):
                if npid == top_face_pid:
                    continue
                info = face_data.get(npid)
                if info is None:
                    continue
                if info["surface_type"] == "plane-surface":
                    candidate_walls.add(npid)

        logger.info("Step 2: %d candidate plane faces adjacent to curved corners: %s",
                     len(candidate_walls), sorted(candidate_walls))

        # Step 3: filter by normal perpendicular to top face
        walls = []
        for pid in sorted(candidate_walls):
            info = face_data.get(pid)
            geom = info.get("geometry", {})
            normal = geom.get("normal")
            if normal is None:
                continue
            normal = _normalize(normal)
            dot = abs(_dot(normal, top_normal))
            if dot <= perpendicular_threshold:
                bb = bboxes.get(pid, (0, 0, 0, 0, 0, 0))
                walls.append(WallInfo(face_pid=pid, normal=normal, bbox=bb))
                logger.info("  Wall: pid=%d, normal=%s, dot=%.3f", pid, normal, dot)
            else:
                logger.info("  Rejected: pid=%d, normal=%s, dot=%.3f (not perpendicular)",
                             pid, normal, dot)

        logger.info("Discovered %d side walls", len(walls))
        return walls


    def find_dimple_faces(
        self,
        wall: WallInfo,
        face_data: Dict[int, dict],
        adjacency: Dict[int, Set[int]],
        bboxes: Dict[int, Tuple],
        top_face_pid: int,
        all_walls: List[WallInfo],
        max_span_ratio: float = 0.5,
        normal_threshold: float = 0.3,
        margin: float = 0.1,
    ) -> List[int]:
        """Find dimple/hole faces on a side wall using local UV projection.

        GOLDEN ALGORITHM (tested on wall 81 and wall 102):
        1. Build local UVW coordinate system from wall normal
        2. Project wall bbox into UV to get wall's UV range
        3. Find faces whose UV projection is within the wall's UV range
        4. Filter: exclude top face, all wall faces, wall's direct neighbors
        5. Filter: UV span must be < max_span_ratio of wall span (small faces only)
        6. Filter: for plane/cone faces, normal must NOT be perpendicular to wall
           (perpendicular = structural edge, parallel = dimple face)

        Args:
            wall: the side wall to find dimples on
            face_data: pid → {surface_type, geometry}
            adjacency: pid → set of adjacent pids
            bboxes: pid → bbox tuple
            top_face_pid: pid of the top face
            all_walls: list of all discovered walls
            max_span_ratio: max UV span ratio relative to wall (default 0.5)
            normal_threshold: min |dot| to keep plane/cone faces (default 0.3)
            margin: UV containment margin (default 0.1)

        Returns:
            List of face pids that are dimple/hole faces on this wall.
        """
        wn = wall.normal
        wb = wall.bbox

        # Build local UVW axes (W = wall normal)
        nx, ny, nz = wn
        ref = (1, 0, 0) if abs(nx) < 0.9 else (0, 1, 0)
        ux = ny * ref[2] - nz * ref[1]
        uy = nz * ref[0] - nx * ref[2]
        uz = nx * ref[1] - ny * ref[0]
        mag = math.sqrt(ux*ux + uy*uy + uz*uz)
        u_axis = (ux/mag, uy/mag, uz/mag)
        v_axis = (ny*u_axis[2] - nz*u_axis[1],
                  nz*u_axis[0] - nx*u_axis[2],
                  nx*u_axis[1] - ny*u_axis[0])

        # Project wall bbox into UV
        def _project_uv(bb):
            corners = [
                (bb[0], bb[1], bb[2]), (bb[3], bb[1], bb[2]),
                (bb[0], bb[4], bb[2]), (bb[3], bb[4], bb[2]),
                (bb[0], bb[1], bb[5]), (bb[3], bb[1], bb[5]),
                (bb[0], bb[4], bb[5]), (bb[3], bb[4], bb[5]),
            ]
            us = [c[0]*u_axis[0]+c[1]*u_axis[1]+c[2]*u_axis[2] for c in corners]
            vs = [c[0]*v_axis[0]+c[1]*v_axis[1]+c[2]*v_axis[2] for c in corners]
            return min(us), max(us), min(vs), max(vs)

        wu_min, wu_max, wv_min, wv_max = _project_uv(wb)
        wall_u_span = wu_max - wu_min
        wall_v_span = wv_max - wv_min

        # Compute wall's W position (distance along normal from origin)
        # The wall center projected onto W axis
        wall_center = ((wb[0]+wb[3])/2, (wb[1]+wb[4])/2, (wb[2]+wb[5])/2)
        wall_w = wall_center[0]*wn[0] + wall_center[1]*wn[1] + wall_center[2]*wn[2]

        # Exclude: top face, all walls, wall's direct neighbors
        exclude = {top_face_pid} | {w.face_pid for w in all_walls}
        exclude.update(adjacency.get(wall.face_pid, set()))

        result = []
        for pid, info in face_data.items():
            if pid in exclude:
                continue
            bb = bboxes.get(pid)
            if bb is None:
                continue

            fu_min, fu_max, fv_min, fv_max = _project_uv(bb)

            # Check W distance: face must be close to the wall in the normal direction
            # Use the face's own largest UV span as tolerance
            face_center = ((bb[0]+bb[3])/2, (bb[1]+bb[4])/2, (bb[2]+bb[5])/2)
            face_w = face_center[0]*wn[0] + face_center[1]*wn[1] + face_center[2]*wn[2]
            face_max_uv_span = max(fu_max - fu_min, fv_max - fv_min)
            if abs(face_w - wall_w) > 2 * face_max_uv_span:
                continue

            # Check UV containment
            if not (fu_min >= wu_min - margin and fu_max <= wu_max + margin and
                    fv_min >= wv_min - margin and fv_max <= wv_max + margin):
                continue

            # Check UV span is small
            fu_span = fu_max - fu_min
            fv_span = fv_max - fv_min
            if wall_u_span > 0 and fu_span > wall_u_span * max_span_ratio:
                continue
            if wall_v_span > 0 and fv_span > wall_v_span * max_span_ratio:
                continue

            # Normal filter: reject faces perpendicular to wall
            stype = info.get("surface_type", "")
            geom = info.get("geometry", {})
            if stype == "plane-surface":
                fn = geom.get("normal")
                if fn and abs(_dot(fn, wn)) < normal_threshold:
                    continue
            elif stype == "cone-surface":
                axis = geom.get("axis")
                if axis and abs(_dot(axis, wn)) < normal_threshold:
                    continue

            result.append(pid)

        # Expand: add adjacency neighbors with zero bboxes (spline surfaces
        # whose bbox couldn't be extracted). These are curved transition
        # surfaces of dimples that are adjacent to already-found dimple faces.
        result_set = set(result)
        expanded = set()
        for pid in result:
            for neighbor in adjacency.get(pid, set()):
                if neighbor in result_set or neighbor in exclude or neighbor in expanded:
                    continue
                nbb = bboxes.get(neighbor, (0,0,0,0,0,0))
                if nbb == (0, 0, 0, 0, 0, 0) or nbb == (0.0, 0.0, 0.0, 0.0, 0.0, 0.0):
                    expanded.add(neighbor)
        result_set.update(expanded)

        logger.info("Wall %d: found %d dimple faces (+%d zero-bbox neighbors): %s",
                     wall.face_pid, len(result), len(expanded), sorted(result_set))
        return sorted(result_set)


    def group_seeds_per_wall(
        self,
        walls: List[WallInfo],
        face_data: Dict[int, dict],
        adjacency: Dict[int, Set[int]],
        bboxes: Dict[int, Tuple],
        top_face_pid: int,
        bbox_overlap_margin: float = 2.0,
    ) -> Dict[int, List[dict]]:
        """Group cone seeds by which wall they belong to.

        For each wall, find cone seeds whose bbox overlaps with the wall's
        bbox region, then BFS-group them into hole/dimple groups (excluding
        the wall face and top face from the BFS walk).

        This is analogous to the PCB simplifier's seed grouping, where each
        wall face plays the role of the board reference face.

        Args:
            walls: discovered side walls
            face_data: pid → {surface_type, geometry}
            adjacency: pid → set of adjacent pids
            bboxes: pid → bbox tuple
            top_face_pid: pid of the top face (excluded from BFS)
            bbox_overlap_margin: margin for bbox overlap check

        Returns:
            Dict mapping wall_face_pid → list of seed groups.
            Each group is {"seeds": [cone pids], "loop_faces": [all pids]}.
        """
        # Collect all cone seed pids
        cone_pids = set()
        for pid, info in face_data.items():
            stype = info.get("surface_type", "")
            if "cone" in stype:
                cone_pids.add(pid)

        # Collect all wall face pids (to exclude from BFS)
        wall_pids = {w.face_pid for w in walls}

        # For each wall, find nearby cone seeds and group them
        result: Dict[int, List[dict]] = {}
        assigned_seeds: Set[int] = set()

        for wall in walls:
            wb = wall.bbox
            nearby_cones = []
            for cpid in cone_pids:
                if cpid in assigned_seeds:
                    continue
                cb = bboxes.get(cpid)
                if cb is None:
                    continue
                # Check bbox overlap with margin
                m = bbox_overlap_margin
                if (cb[3]+m >= wb[0] and wb[3]+m >= cb[0] and
                    cb[4]+m >= wb[1] and wb[4]+m >= cb[1] and
                    cb[5]+m >= wb[2] and wb[5]+m >= cb[2]):
                    nearby_cones.append(cpid)

            if not nearby_cones:
                continue

            # BFS-group the nearby cones into hole groups
            # Exclude: top face, all wall faces, already-assigned seeds
            exclude = {top_face_pid} | wall_pids
            groups = self._group_seeds_bfs(
                nearby_cones, adjacency, exclude, assigned_seeds)

            if groups:
                result[wall.face_pid] = groups
                for g in groups:
                    assigned_seeds.update(g["seeds"])

            logger.info("Wall pid=%d: %d nearby cones, %d groups",
                         wall.face_pid, len(nearby_cones), len(groups))

        return result

    def _group_seeds_bfs(
        self,
        seed_pids: List[int],
        adjacency: Dict[int, Set[int]],
        exclude: Set[int],
        already_assigned: Set[int],
    ) -> List[dict]:
        """Group seed faces into connected components via BFS.

        Each group contains seeds that are connected through adjacency
        (excluding the wall/top faces). Non-cone faces encountered during
        BFS are included in loop_faces.

        Returns:
            List of {"seeds": [...], "loop_faces": [...]}
        """
        visited: Set[int] = set()
        groups = []

        for seed in sorted(seed_pids):
            if seed in visited or seed in already_assigned:
                continue

            # BFS from this seed
            queue = [seed]
            visited.add(seed)
            group_seeds = []
            loop_faces = []

            while queue:
                current = queue.pop(0)
                loop_faces.append(current)
                if current in set(seed_pids):
                    group_seeds.append(current)

                for neighbor in adjacency.get(current, set()):
                    if neighbor in visited or neighbor in exclude or neighbor in already_assigned:
                        continue
                    visited.add(neighbor)
                    queue.append(neighbor)

            if group_seeds:
                groups.append({
                    "seeds": sorted(group_seeds),
                    "loop_faces": sorted(loop_faces),
                })

        return groups
