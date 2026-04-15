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


    def discover_side_walls_validated(
        self,
        top_face_pid: int,
        top_normal: Tuple[float, float, float],
        face_data: Dict[int, dict],
        adjacency: Dict[int, Set[int]],
        bboxes: Dict[int, Tuple],
        perpendicular_threshold: float = 0.05,
        w_touch_tolerance: float = 0.1,
    ) -> List[WallInfo]:
        """Find side walls with W-touch and aspect ratio validation.

        Improved algorithm over discover_side_walls:
        1. Find curved faces adjacent to top face (corner fillets)
        2. Find plane faces adjacent to fillets, perpendicular to top normal
        3. W-touch validation: wall bbox must touch fillet bbox in W direction
           (W = top face normal). Rejects dimple faces that share SAT edges
           with fillets but aren't physically adjacent in W.
        4. Aspect ratio validation: in wall-local coords (W=wall normal,
           U=ref normal, V=W×U), V span must be >= U span. Real walls are
           longer than tall; small structural faces are not.

        Args:
            top_face_pid: pid of the top/reference face
            top_normal: normal of the top face (used as W axis)
            face_data: pid → {surface_type, geometry}
            adjacency: pid → set of adjacent pids
            bboxes: pid → bbox tuple
            perpendicular_threshold: max |dot| for perpendicular check
            w_touch_tolerance: tolerance in mm for W-touch check

        Returns:
            List of WallInfo for validated side walls.
        """
        w_axis = top_normal

        def _w_range(bbox):
            if bbox is None:
                return None
            if bbox == (0,0,0,0,0,0) or bbox == (0.0,0.0,0.0,0.0,0.0,0.0):
                return None
            corners = [
                (bbox[0],bbox[1],bbox[2]),(bbox[3],bbox[1],bbox[2]),
                (bbox[0],bbox[4],bbox[2]),(bbox[3],bbox[4],bbox[2]),
                (bbox[0],bbox[1],bbox[5]),(bbox[3],bbox[1],bbox[5]),
                (bbox[0],bbox[4],bbox[5]),(bbox[3],bbox[4],bbox[5]),
            ]
            ws = [c[0]*w_axis[0]+c[1]*w_axis[1]+c[2]*w_axis[2] for c in corners]
            return (min(ws), max(ws))

        def _w_touches(r1, r2):
            if r1 is None or r2 is None:
                return False
            return (r1[1] + w_touch_tolerance >= r2[0] and
                    r2[1] + w_touch_tolerance >= r1[0])

        # Step 1: curved faces adjacent to top face
        top_neighbors = adjacency.get(top_face_pid, set())
        curved_faces = set()
        for npid in top_neighbors:
            info = face_data.get(npid)
            if info and "plane" not in info["surface_type"]:
                curved_faces.add(npid)

        # Step 2: candidate walls + track connecting fillets
        candidate_set = set()
        wall_to_fillets: Dict[int, Set[int]] = {}
        for cfid in curved_faces:
            for npid in adjacency.get(cfid, set()):
                if npid == top_face_pid:
                    continue
                info = face_data.get(npid)
                if info is None or info["surface_type"] != "plane-surface":
                    continue
                geom = info.get("geometry", {})
                normal = geom.get("normal")
                if normal is None:
                    continue
                normal = _normalize(normal)
                if abs(_dot(normal, top_normal)) > perpendicular_threshold:
                    continue
                candidate_set.add(npid)
                if npid not in wall_to_fillets:
                    wall_to_fillets[npid] = set()
                wall_to_fillets[npid].add(cfid)

        logger.info("Step 2: %d candidates before validation", len(candidate_set))

        # Step 3: W-touch validation
        w_passed = []
        for pid in sorted(candidate_set):
            wall_wr = _w_range(bboxes.get(pid))
            fillets = wall_to_fillets.get(pid, set())
            touches = False
            for cfid in fillets:
                cf_wr = _w_range(bboxes.get(cfid))
                if _w_touches(wall_wr, cf_wr):
                    touches = True
                    break
            if touches:
                info = face_data.get(pid)
                normal = _normalize(info.get("geometry", {}).get("normal", (0,0,0)))
                bb = bboxes.get(pid, (0,0,0,0,0,0))
                w_passed.append(WallInfo(face_pid=pid, normal=normal, bbox=bb))
            else:
                logger.info("  W-touch rejected: pid=%d", pid)

        # Step 4: Aspect ratio validation (V span >= U span)
        walls = []
        for w in w_passed:
            wn = w.normal
            u_ax = top_normal
            # V = W_wall × U
            vx = wn[1]*u_ax[2] - wn[2]*u_ax[1]
            vy = wn[2]*u_ax[0] - wn[0]*u_ax[2]
            vz = wn[0]*u_ax[1] - wn[1]*u_ax[0]
            vmag = math.sqrt(vx*vx + vy*vy + vz*vz)
            if vmag < 1e-12:
                walls.append(w)
                continue
            v_ax = (vx/vmag, vy/vmag, vz/vmag)

            bb = w.bbox
            corners = [
                (bb[0],bb[1],bb[2]),(bb[3],bb[1],bb[2]),
                (bb[0],bb[4],bb[2]),(bb[3],bb[4],bb[2]),
                (bb[0],bb[1],bb[5]),(bb[3],bb[1],bb[5]),
                (bb[0],bb[4],bb[5]),(bb[3],bb[4],bb[5]),
            ]
            us = [c[0]*u_ax[0]+c[1]*u_ax[1]+c[2]*u_ax[2] for c in corners]
            vs = [c[0]*v_ax[0]+c[1]*v_ax[1]+c[2]*v_ax[2] for c in corners]
            u_span = max(us) - min(us)
            v_span = max(vs) - min(vs)

            if v_span >= u_span:
                walls.append(w)
            else:
                logger.info("  Aspect rejected: pid=%d, U=%.3f, V=%.3f",
                            w.face_pid, u_span, v_span)

        logger.info("Validated %d side walls (from %d candidates)",
                     len(walls), len(candidate_set))
        return walls, curved_faces


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
        corner_fillets: Set[int] = None,
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
            corner_fillets: set of corner fillet face pids (used to filter
                zero-bbox expansion — skip faces adjacent to fillets)

        Returns:
            List of face pids that are dimple/hole faces on this wall.
        """
        wn = wall.normal
        wb = wall.bbox

        # Build local UVW axes (W = wall normal, U = short edge, V = long edge)
        nx, ny, nz = wn
        ref = (1, 0, 0) if abs(nx) < 0.9 else (0, 1, 0)
        ux = ny * ref[2] - nz * ref[1]
        uy = nz * ref[0] - nx * ref[2]
        uz = nx * ref[1] - ny * ref[0]
        mag = math.sqrt(ux*ux + uy*uy + uz*uz)
        tmp_u = (ux/mag, uy/mag, uz/mag)
        tmp_v = (ny*tmp_u[2] - nz*tmp_u[1],
                  nz*tmp_u[0] - nx*tmp_u[2],
                  nx*tmp_u[1] - ny*tmp_u[0])

        # Project wall bbox into temp UV to find which is short/long
        def _project_uv(bb):
            corners = [
                (bb[0], bb[1], bb[2]), (bb[3], bb[1], bb[2]),
                (bb[0], bb[4], bb[2]), (bb[3], bb[4], bb[2]),
                (bb[0], bb[1], bb[5]), (bb[3], bb[1], bb[5]),
                (bb[0], bb[4], bb[5]), (bb[3], bb[4], bb[5]),
            ]
            us = [c[0]*tmp_u[0]+c[1]*tmp_u[1]+c[2]*tmp_u[2] for c in corners]
            vs = [c[0]*tmp_v[0]+c[1]*tmp_v[1]+c[2]*tmp_v[2] for c in corners]
            return min(us), max(us), min(vs), max(vs)

        tmp_u_min, tmp_u_max, tmp_v_min, tmp_v_max = _project_uv(wb)
        tmp_u_span = tmp_u_max - tmp_u_min
        tmp_v_span = tmp_v_max - tmp_v_min

        # Swap so U = short edge, V = long edge
        if tmp_u_span > tmp_v_span:
            tmp_u, tmp_v = tmp_v, tmp_u

        u_axis = tmp_u
        v_axis = tmp_v

        # Redefine _project_uv with final axes
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

        # Exclude: top face, all walls, and corner fillets
        exclude = {top_face_pid} | {w.face_pid for w in all_walls}
        if corner_fillets:
            exclude |= corner_fillets

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

            # Check UV span: only check V (long edge) — dimples can span full wall width
            fu_span = fu_max - fu_min
            fv_span = fv_max - fv_min
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

            # U-midpoint filter: dimple faces sit in the middle of the wall
            # along U (short edge = copper thickness direction).
            # Corner fillets are at the edges (top/bottom of wall in U).
            # Reject faces whose U center is in the outer 10% margins.
            if wall_u_span > 0:
                face_u_center = (fu_min + fu_max) / 2
                wall_u_mid = (wu_min + wu_max) / 2
                u_offset = abs(face_u_center - wall_u_mid)
                if u_offset > wall_u_span * 0.4:
                    continue

            result.append(pid)

        # Expand: add adjacency neighbors with zero bboxes (spline surfaces
        # whose bbox couldn't be extracted). These are curved transition
        # surfaces of dimples that are adjacent to already-found dimple faces.
        # Skip faces adjacent to corner fillets — those are structural surfaces.
        fillet_set = corner_fillets or set()
        result_set = set(result)
        expanded = set()
        for pid in result:
            for neighbor in adjacency.get(pid, set()):
                if neighbor in result_set or neighbor in exclude or neighbor in expanded:
                    continue
                nbb = bboxes.get(neighbor, (0,0,0,0,0,0))
                if nbb == (0, 0, 0, 0, 0, 0) or nbb == (0.0, 0.0, 0.0, 0.0, 0.0, 0.0):
                    # Skip if adjacent to any corner fillet
                    if fillet_set and (adjacency.get(neighbor, set()) & fillet_set):
                        continue
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

    def find_dimples_merged_group(
        self,
        walls_group: List[WallInfo],
        face_data: Dict[int, dict],
        adjacency: Dict[int, Set[int]],
        bboxes: Dict[int, Tuple],
        top_face_pid: int,
        all_walls: List[WallInfo],
        max_span_ratio: float = 0.5,
        normal_threshold: float = 0.3,
        margin: float = 0.1,
    ) -> List[int]:
        """Find dimple faces for a group of same-normal walls using merged bbox.

        Key differences from find_dimple_faces:
        - Uses union bbox of all walls in the group
        - Exclude set: only {ref face, all wall faces} — NOT wall neighbors
        - This allows dimple faces adjacent to the wall to be found
        """
        wn = walls_group[0].normal

        # Merge bboxes
        merged_bb = None
        for wall in walls_group:
            wb = wall.bbox
            if merged_bb is None:
                merged_bb = list(wb)
            else:
                for i in range(3):
                    merged_bb[i] = min(merged_bb[i], wb[i])
                for i in range(3, 6):
                    merged_bb[i] = max(merged_bb[i], wb[i])
        merged_bb = tuple(merged_bb)

        # Build UVW (U=short edge, V=long edge)
        nx, ny, nz = wn
        ref = (1, 0, 0) if abs(nx) < 0.9 else (0, 1, 0)
        ux = ny * ref[2] - nz * ref[1]
        uy = nz * ref[0] - nx * ref[2]
        uz = nx * ref[1] - ny * ref[0]
        mag = math.sqrt(ux*ux + uy*uy + uz*uz)
        tmp_u = (ux/mag, uy/mag, uz/mag)
        tmp_v = (ny*tmp_u[2] - nz*tmp_u[1],
                 nz*tmp_u[0] - nx*tmp_u[2],
                 nx*tmp_u[1] - ny*tmp_u[0])

        def _project(bb, u_ax, v_ax):
            corners = [
                (bb[0],bb[1],bb[2]),(bb[3],bb[1],bb[2]),
                (bb[0],bb[4],bb[2]),(bb[3],bb[4],bb[2]),
                (bb[0],bb[1],bb[5]),(bb[3],bb[1],bb[5]),
                (bb[0],bb[4],bb[5]),(bb[3],bb[4],bb[5]),
            ]
            us = [c[0]*u_ax[0]+c[1]*u_ax[1]+c[2]*u_ax[2] for c in corners]
            vs = [c[0]*v_ax[0]+c[1]*v_ax[1]+c[2]*v_ax[2] for c in corners]
            return min(us), max(us), min(vs), max(vs)

        tu_min, tu_max, tv_min, tv_max = _project(merged_bb, tmp_u, tmp_v)
        if (tu_max - tu_min) > (tv_max - tv_min):
            tmp_u, tmp_v = tmp_v, tmp_u
        u_axis, v_axis = tmp_u, tmp_v

        def _project_uv(bb):
            return _project(bb, u_axis, v_axis)

        wu_min, wu_max, wv_min, wv_max = _project_uv(merged_bb)
        wall_v_span = wv_max - wv_min

        wall_center = ((merged_bb[0]+merged_bb[3])/2, (merged_bb[1]+merged_bb[4])/2,
                        (merged_bb[2]+merged_bb[5])/2)
        wall_w = wall_center[0]*wn[0] + wall_center[1]*wn[1] + wall_center[2]*wn[2]

        # Exclude: only ref face + all wall faces (NOT wall neighbors)
        exclude = {top_face_pid} | {w.face_pid for w in all_walls}

        result = []
        for pid, info in face_data.items():
            if pid in exclude:
                continue
            bb = bboxes.get(pid)
            if bb is None:
                continue
            fu_min, fu_max, fv_min, fv_max = _project_uv(bb)
            face_center = ((bb[0]+bb[3])/2, (bb[1]+bb[4])/2, (bb[2]+bb[5])/2)
            face_w = face_center[0]*wn[0] + face_center[1]*wn[1] + face_center[2]*wn[2]
            face_max_uv_span = max(fu_max - fu_min, fv_max - fv_min)
            if abs(face_w - wall_w) > 2 * face_max_uv_span:
                continue
            if not (fu_min >= wu_min - margin and fu_max <= wu_max + margin and
                    fv_min >= wv_min - margin and fv_max <= wv_max + margin):
                continue
            fv_span = fv_max - fv_min
            if wall_v_span > 0 and fv_span > wall_v_span * max_span_ratio:
                continue
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

        # Zero-bbox expansion
        result_set = set(result)
        expanded = set()
        for pid in result:
            for neighbor in adjacency.get(pid, set()):
                if neighbor in result_set or neighbor in exclude or neighbor in expanded:
                    continue
                nbb = bboxes.get(neighbor, (0,0,0,0,0,0))
                if nbb == (0,0,0,0,0,0) or nbb == (0.0,0.0,0.0,0.0,0.0,0.0):
                    expanded.add(neighbor)
        result_set.update(expanded)

        logger.info("Merged walls %s: found %d dimple faces (+%d zero-bbox): %s",
                     [w.face_pid for w in walls_group], len(result), len(expanded),
                     sorted(result_set))
        return sorted(result_set)


def group_walls_by_normal(walls: List[WallInfo], threshold: float = 0.99) -> List[List[WallInfo]]:
    """Group walls with similar normals (dot > threshold).

    Returns list of groups. Each group is a list of WallInfo with similar normals.
    Single walls (no similar partner) are returned as groups of 1.
    """
    assigned = set()
    groups = []
    for i, w1 in enumerate(walls):
        if w1.face_pid in assigned:
            continue
        group = [w1]
        assigned.add(w1.face_pid)
        for j, w2 in enumerate(walls):
            if j <= i or w2.face_pid in assigned:
                continue
            if abs(_dot(w1.normal, w2.normal)) > threshold:
                group.append(w2)
                assigned.add(w2.face_pid)
        groups.append(group)
    return groups


def assign_dimples_to_nearest_wall(
    walls: List[WallInfo],
    wall_dimples: Dict[int, List[int]],
    bboxes: Dict[int, Tuple],
) -> Dict[int, List[int]]:
    """Reassign dimple faces to the nearest wall by W distance.

    When multiple walls have the same normal (e.g., two halves of the same
    side), dimples from one wall may pass the spatial filters for the other.
    This function checks each dimple's W distance to all walls with the same
    normal and assigns it to the closest one.

    Args:
        walls: list of WallInfo objects
        wall_dimples: dict mapping wall.face_pid → list of dimple face IDs
        bboxes: per-face bounding boxes

    Returns:
        Updated wall_dimples dict with reassigned faces.
    """
    # Build wall W positions
    wall_w_positions = {}
    for wall in walls:
        wb = wall.bbox
        wn = wall.normal
        wc = ((wb[0]+wb[3])/2, (wb[1]+wb[4])/2, (wb[2]+wb[5])/2)
        wall_w_positions[wall.face_pid] = wc[0]*wn[0] + wc[1]*wn[1] + wc[2]*wn[2]

    # Group walls by similar normal (dot > 0.99)
    wall_groups = []
    assigned_walls = set()
    for i, w1 in enumerate(walls):
        if w1.face_pid in assigned_walls:
            continue
        group = [w1]
        assigned_walls.add(w1.face_pid)
        for j, w2 in enumerate(walls):
            if j <= i or w2.face_pid in assigned_walls:
                continue
            if abs(_dot(w1.normal, w2.normal)) > 0.99:
                group.append(w2)
                assigned_walls.add(w2.face_pid)
        if len(group) > 1:
            wall_groups.append(group)

    # For each group of same-normal walls, reassign dimples
    result = dict(wall_dimples)
    for group in wall_groups:
        group_pids = [w.face_pid for w in group]
        # Collect all dimples from all walls in this group
        all_dimples = set()
        for pid in group_pids:
            all_dimples.update(result.get(pid, []))

        if not all_dimples:
            continue

        # Reassign each dimple to the nearest wall
        new_assignment = {pid: [] for pid in group_pids}
        for fid in all_dimples:
            bb = bboxes.get(fid)
            if bb is None:
                # Can't compute W distance — assign to first wall that had it
                for pid in group_pids:
                    if fid in result.get(pid, []):
                        new_assignment[pid].append(fid)
                        break
                continue

            # Compute face W position using the group's shared normal
            wn = group[0].normal
            fc = ((bb[0]+bb[3])/2, (bb[1]+bb[4])/2, (bb[2]+bb[5])/2)
            face_w = fc[0]*wn[0] + fc[1]*wn[1] + fc[2]*wn[2]

            # Find nearest wall
            best_pid = group_pids[0]
            best_dist = float('inf')
            for pid in group_pids:
                dist = abs(face_w - wall_w_positions[pid])
                if dist < best_dist:
                    best_dist = dist
                    best_pid = pid
            new_assignment[best_pid].append(fid)

        for pid in group_pids:
            result[pid] = sorted(new_assignment[pid])

    return result


def find_dimples_for_wall_group(
    walls_group: List[WallInfo],
    face_data,
    adjacency,
    bboxes,
    ref_pid: int,
    all_walls: List[WallInfo],
    max_span_ratio: float = 0.5,
    normal_threshold: float = 0.3,
    margin: float = 0.1,
) -> List[int]:
    """Find dimple faces for a group of same-normal walls using merged bbox.

    Key differences from find_dimple_faces:
    - Uses union bbox of all walls in the group
    - Exclude set: only {ref face, all wall faces} — NOT wall neighbors
    - U = short edge, V = long edge; only V span is checked
    """
    wn = walls_group[0].normal

    # Merge bboxes
    merged_bb = None
    for wall in walls_group:
        wb = wall.bbox
        if merged_bb is None:
            merged_bb = list(wb)
        else:
            for i in range(3):
                merged_bb[i] = min(merged_bb[i], wb[i])
            for i in range(3, 6):
                merged_bb[i] = max(merged_bb[i], wb[i])
    merged_bb = tuple(merged_bb)

    nx, ny, nz = wn
    ref = (1, 0, 0) if abs(nx) < 0.9 else (0, 1, 0)
    ux = ny * ref[2] - nz * ref[1]
    uy = nz * ref[0] - nx * ref[2]
    uz = nx * ref[1] - ny * ref[0]
    mag = math.sqrt(ux*ux + uy*uy + uz*uz)
    tmp_u = (ux/mag, uy/mag, uz/mag)
    tmp_v = (ny*tmp_u[2] - nz*tmp_u[1],
             nz*tmp_u[0] - nx*tmp_u[2],
             nx*tmp_u[1] - ny*tmp_u[0])

    def _project(bb, u_ax, v_ax):
        corners = [
            (bb[0],bb[1],bb[2]),(bb[3],bb[1],bb[2]),
            (bb[0],bb[4],bb[2]),(bb[3],bb[4],bb[2]),
            (bb[0],bb[1],bb[5]),(bb[3],bb[1],bb[5]),
            (bb[0],bb[4],bb[5]),(bb[3],bb[4],bb[5]),
        ]
        us = [c[0]*u_ax[0]+c[1]*u_ax[1]+c[2]*u_ax[2] for c in corners]
        vs = [c[0]*v_ax[0]+c[1]*v_ax[1]+c[2]*v_ax[2] for c in corners]
        return min(us), max(us), min(vs), max(vs)

    tu_min, tu_max, tv_min, tv_max = _project(merged_bb, tmp_u, tmp_v)
    if (tu_max - tu_min) > (tv_max - tv_min):
        tmp_u, tmp_v = tmp_v, tmp_u
    u_axis, v_axis = tmp_u, tmp_v

    def _project_uv(bb):
        return _project(bb, u_axis, v_axis)

    wu_min, wu_max, wv_min, wv_max = _project_uv(merged_bb)
    wall_v_span = wv_max - wv_min

    wall_center = ((merged_bb[0]+merged_bb[3])/2, (merged_bb[1]+merged_bb[4])/2, (merged_bb[2]+merged_bb[5])/2)
    wall_w = wall_center[0]*wn[0] + wall_center[1]*wn[1] + wall_center[2]*wn[2]

    exclude = {ref_pid} | {w.face_pid for w in all_walls}

    result = []
    for pid, info in face_data.items():
        if pid in exclude:
            continue
        bb = bboxes.get(pid)
        if bb is None:
            continue
        fu_min, fu_max, fv_min, fv_max = _project_uv(bb)
        face_center = ((bb[0]+bb[3])/2, (bb[1]+bb[4])/2, (bb[2]+bb[5])/2)
        face_w = face_center[0]*wn[0] + face_center[1]*wn[1] + face_center[2]*wn[2]
        face_max_uv_span = max(fu_max - fu_min, fv_max - fv_min)
        if abs(face_w - wall_w) > 2 * face_max_uv_span:
            continue
        if not (fu_min >= wu_min - margin and fu_max <= wu_max + margin and
                fv_min >= wv_min - margin and fv_max <= wv_max + margin):
            continue
        fv_span = fv_max - fv_min
        if wall_v_span > 0 and fv_span > wall_v_span * max_span_ratio:
            continue
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

    # Zero-bbox expansion
    result_set = set(result)
    expanded = set()
    for pid in result:
        for neighbor in adjacency.get(pid, set()):
            if neighbor in result_set or neighbor in exclude or neighbor in expanded:
                continue
            nbb = bboxes.get(neighbor, (0,0,0,0,0,0))
            if nbb == (0,0,0,0,0,0) or nbb == (0.0,0.0,0.0,0.0,0.0,0.0):
                expanded.add(neighbor)
    result_set.update(expanded)

    return sorted(result_set)


def group_walls_by_normal(walls: List[WallInfo], threshold: float = 0.99) -> List[List[WallInfo]]:
    """Group walls with the same normal direction.

    All walls with dot > threshold go into the same group.
    No size or area filtering — just normal direction.
    """
    groups = []
    assigned = set()
    for i, w1 in enumerate(walls):
        if w1.face_pid in assigned:
            continue
        group = [w1]
        assigned.add(w1.face_pid)
        for j, w2 in enumerate(walls):
            if j <= i or w2.face_pid in assigned:
                continue
            if abs(_dot(w1.normal, w2.normal)) > threshold:
                group.append(w2)
                assigned.add(w2.face_pid)
        groups.append(group)
    return groups


def split_subgroups_by_uv_overlap(
    sub_groups: List[List[WallInfo]],
    ref_normal: Tuple[float, float, float],
    margin: float = 0.5,
) -> List[List[WallInfo]]:
    """Split sub-groups with >2 walls by UV bbox overlap.

    For sub-groups where walls have the same normal and similar W position
    but are at different locations along the side, their UV ranges won't
    overlap. This splits them into separate sub-groups.

    Local coords per wall group: W = wall normal, U = ref normal, V = W × U.

    Args:
        sub_groups: list of wall sub-groups from W-distance splitting
        ref_normal: reference face normal (used as U axis)
        margin: UV overlap margin in mm

    Returns:
        List of sub-groups, each with walls whose UV ranges overlap.
    """
    def _wall_area(w):
        b = w.bbox
        dims = sorted([b[3]-b[0], b[4]-b[1], b[5]-b[2]], reverse=True)
        return dims[0] * dims[1]

    result = []
    for sg in sub_groups:
        if len(sg) <= 2:
            result.append(sg)
            continue

        wn = sg[0].normal
        u_ax = ref_normal
        vx = wn[1]*u_ax[2] - wn[2]*u_ax[1]
        vy = wn[2]*u_ax[0] - wn[0]*u_ax[2]
        vz = wn[0]*u_ax[1] - wn[1]*u_ax[0]
        vmag = math.sqrt(vx*vx + vy*vy + vz*vz)
        if vmag < 1e-12:
            result.append(sg)
            continue
        v_ax = (vx/vmag, vy/vmag, vz/vmag)

        def _uv_range(w):
            bb = w.bbox
            corners = [
                (bb[0],bb[1],bb[2]),(bb[3],bb[1],bb[2]),
                (bb[0],bb[4],bb[2]),(bb[3],bb[4],bb[2]),
                (bb[0],bb[1],bb[5]),(bb[3],bb[1],bb[5]),
                (bb[0],bb[4],bb[5]),(bb[3],bb[4],bb[5]),
            ]
            us = [c[0]*u_ax[0]+c[1]*u_ax[1]+c[2]*u_ax[2] for c in corners]
            vs = [c[0]*v_ax[0]+c[1]*v_ax[1]+c[2]*v_ax[2] for c in corners]
            return (min(us), max(us), min(vs), max(vs))

        def _overlaps(r1, r2):
            return not (r1[1]+margin < r2[0] or r2[1]+margin < r1[0] or
                        r1[3]+margin < r2[2] or r2[3]+margin < r1[2])

        uv_ranges = {w.face_pid: _uv_range(w) for w in sg}
        clustered = set()
        for w in sorted(sg, key=_wall_area, reverse=True):
            if w.face_pid in clustered:
                continue
            cluster = [w]
            clustered.add(w.face_pid)
            for w2 in sg:
                if w2.face_pid in clustered:
                    continue
                if any(_overlaps(uv_ranges[cw.face_pid], uv_ranges[w2.face_pid]) for cw in cluster):
                    cluster.append(w2)
                    clustered.add(w2.face_pid)
            result.append(cluster)

    return result
