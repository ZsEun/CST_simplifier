"""Unit tests for FeatureDetector pure-Python helpers.

Tests parsing of VBA output lines and face grouping logic
without requiring a running CST instance.
"""

import math
from unittest.mock import MagicMock

from code.feature_detector import FeatureDetector
from code.models import FaceInfo, SimplificationCandidate


def _make_face(
    face_id=0,
    solid="solid1",
    component="comp1",
    surface_type="Cylinder",
    cx=0.0, cy=0.0, cz=0.0,
    ax=0.0, ay=0.0, az=1.0,
    radius=1.0,
    bbox_min=(0.0, 0.0, 0.0),
    bbox_max=(2.0, 2.0, 5.0),
):
    return FaceInfo(
        face_id=face_id,
        solid_name=solid,
        component=component,
        surface_type=surface_type,
        center_x=cx, center_y=cy, center_z=cz,
        axis_x=ax, axis_y=ay, axis_z=az,
        radius=radius,
        bbox_min=bbox_min,
        bbox_max=bbox_max,
    )


class TestParseFaceLine:
    """Tests for FeatureDetector._parse_face_line."""

    def test_valid_cylindrical_line(self):
        line = "3|Cylinder|1.0|2.0|3.0|0.0|0.0|0.0|2.0|2.0|6.0"
        face = FeatureDetector._parse_face_line(line, "comp1", "solid1")
        assert face is not None
        assert face.face_id == 3
        assert face.surface_type == "Cylinder"
        assert face.center_x == 1.0
        assert face.component == "comp1"
        assert face.solid_name == "solid1"
        assert face.bbox_min == (0.0, 0.0, 0.0)
        assert face.bbox_max == (2.0, 2.0, 6.0)
        # Two similar extents (2.0, 2.0) → cylindrical, radius = 2.0/2 = 1.0
        assert face.radius > 0

    def test_malformed_line_returns_none(self):
        face = FeatureDetector._parse_face_line("bad|data", "c", "s")
        assert face is None

    def test_empty_line_returns_none(self):
        face = FeatureDetector._parse_face_line("", "c", "s")
        assert face is None

    def test_non_numeric_values_returns_none(self):
        line = "0|Plane|abc|0|0|0|0|0|1|1|1"
        face = FeatureDetector._parse_face_line(line, "c", "s")
        assert face is None


class TestFilterInnerFaces:
    """Tests for FeatureDetector._filter_inner_faces."""

    def test_keeps_cylindrical_faces(self):
        faces = [
            _make_face(surface_type="Cylinder", radius=1.5),
            _make_face(surface_type="Plane", radius=0.0),
        ]
        result = FeatureDetector._filter_inner_faces(faces)
        assert len(result) == 1
        assert result[0].surface_type == "Cylinder"

    def test_keeps_faces_with_nonzero_radius(self):
        faces = [_make_face(surface_type="Unknown", radius=0.5)]
        result = FeatureDetector._filter_inner_faces(faces)
        assert len(result) == 1

    def test_excludes_planar_zero_radius(self):
        faces = [_make_face(surface_type="Plane", radius=0.0)]
        result = FeatureDetector._filter_inner_faces(faces)
        assert len(result) == 0


class TestGroupFacesIntoHoles:
    """Tests for FeatureDetector._group_faces_into_holes."""

    def test_two_coaxial_cylindrical_faces_grouped(self):
        detector = FeatureDetector(MagicMock())
        f1 = _make_face(face_id=0, cx=5.0, cy=5.0, cz=1.0,
                        ax=0.0, ay=0.0, az=1.0, radius=1.0,
                        surface_type="Cylinder")
        f2 = _make_face(face_id=1, cx=5.0, cy=5.0, cz=3.0,
                        ax=0.0, ay=0.0, az=1.0, radius=1.0,
                        surface_type="Cylinder")
        groups = detector._group_faces_into_holes([f1, f2])
        assert len(groups) == 1
        assert len(groups[0]) == 2

    def test_faces_with_different_axes_separate_groups(self):
        detector = FeatureDetector(MagicMock())
        f1 = _make_face(face_id=0, ax=0.0, ay=0.0, az=1.0,
                        cx=0.0, cy=0.0, cz=0.0,
                        surface_type="Cylinder")
        f2 = _make_face(face_id=1, ax=1.0, ay=0.0, az=0.0,
                        cx=10.0, cy=10.0, cz=0.0,
                        surface_type="Cylinder")
        groups = detector._group_faces_into_holes([f1, f2])
        assert len(groups) == 2

    def test_no_inner_faces_returns_empty(self):
        detector = FeatureDetector(MagicMock())
        faces = [_make_face(surface_type="Plane", radius=0.0)]
        groups = detector._group_faces_into_holes(faces)
        assert groups == []

    def test_empty_input_returns_empty(self):
        detector = FeatureDetector(MagicMock())
        assert detector._group_faces_into_holes([]) == []


class TestClassifyHole:
    """Tests for FeatureDetector._classify_hole."""

    def test_single_cylinder_is_cylindrical(self):
        faces = [_make_face(surface_type="Cylinder", radius=1.5)]
        assert FeatureDetector._classify_hole(faces) == "cylindrical"

    def test_multiple_same_radius_cylinders_is_cylindrical(self):
        faces = [
            _make_face(face_id=0, surface_type="Cylinder", radius=1.5),
            _make_face(face_id=1, surface_type="Cylinder", radius=1.5),
        ]
        assert FeatureDetector._classify_hole(faces) == "cylindrical"

    def test_cylinders_within_tolerance_is_cylindrical(self):
        faces = [
            _make_face(face_id=0, surface_type="Cylinder", radius=1.50),
            _make_face(face_id=1, surface_type="Cylinder", radius=1.55),
        ]
        assert FeatureDetector._classify_hole(faces) == "cylindrical"

    def test_varying_radii_thread_faces_is_threaded(self):
        faces = [
            _make_face(face_id=0, surface_type="Cylinder", radius=1.0),
            _make_face(face_id=1, surface_type="Cone", radius=1.5),
            _make_face(face_id=2, surface_type="BSpline", radius=0.8),
        ]
        assert FeatureDetector._classify_hole(faces) == "threaded"

    def test_mixed_non_thread_faces_is_irregular(self):
        faces = [
            _make_face(face_id=0, surface_type="Plane", radius=1.0),
            _make_face(face_id=1, surface_type="Plane", radius=2.0),
        ]
        assert FeatureDetector._classify_hole(faces) == "irregular"

    def test_empty_group_is_irregular(self):
        assert FeatureDetector._classify_hole([]) == "irregular"


class TestComputeBoundingRadius:
    """Tests for FeatureDetector._compute_bounding_radius."""

    def test_single_face_returns_its_radius(self):
        faces = [_make_face(radius=2.5)]
        assert FeatureDetector._compute_bounding_radius(faces) == 2.5

    def test_multiple_faces_returns_max_radius(self):
        faces = [
            _make_face(face_id=0, radius=1.0),
            _make_face(face_id=1, radius=3.0),
            _make_face(face_id=2, radius=2.0),
        ]
        assert FeatureDetector._compute_bounding_radius(faces) == 3.0

    def test_empty_group_returns_zero(self):
        assert FeatureDetector._compute_bounding_radius([]) == 0.0

    def test_zero_radius_falls_back_to_bbox(self):
        faces = [
            _make_face(
                radius=0.0,
                bbox_min=(0.0, 0.0, 0.0),
                bbox_max=(4.0, 4.0, 10.0),
            ),
        ]
        result = FeatureDetector._compute_bounding_radius(faces)
        # Sorted extents: [4.0, 4.0, 10.0], second / 2 = 2.0
        assert result == 2.0


class TestComputeDepth:
    """Tests for FeatureDetector._compute_depth."""

    def test_faces_along_z_axis(self):
        faces = [
            _make_face(
                face_id=0, cx=0.0, cy=0.0, cz=1.0,
                ax=0.0, ay=0.0, az=1.0,
                bbox_min=(0.0, 0.0, 0.0), bbox_max=(2.0, 2.0, 5.0),
            ),
            _make_face(
                face_id=1, cx=0.0, cy=0.0, cz=4.0,
                ax=0.0, ay=0.0, az=1.0,
                bbox_min=(0.0, 0.0, 2.0), bbox_max=(2.0, 2.0, 8.0),
            ),
        ]
        depth = FeatureDetector._compute_depth(faces)
        # center span = 4.0 - 1.0 = 3.0, bbox depth = max(5.0, 6.0) = 6.0
        assert depth == 6.0

    def test_single_face_uses_bbox(self):
        faces = [
            _make_face(
                ax=0.0, ay=0.0, az=1.0,
                bbox_min=(0.0, 0.0, 0.0), bbox_max=(2.0, 2.0, 7.0),
            ),
        ]
        depth = FeatureDetector._compute_depth(faces)
        assert depth == 7.0

    def test_empty_group_returns_zero(self):
        assert FeatureDetector._compute_depth([]) == 0.0

    def test_no_valid_axis_falls_back_to_bbox_max(self):
        faces = [
            _make_face(
                ax=0.0, ay=0.0, az=0.0,
                bbox_min=(0.0, 0.0, 0.0), bbox_max=(3.0, 4.0, 5.0),
            ),
        ]
        depth = FeatureDetector._compute_depth(faces)
        assert depth == 5.0


class TestSelectPrimaryFace:
    """Tests for FeatureDetector._select_primary_face."""

    def test_selects_largest_radius(self):
        faces = [
            _make_face(face_id=10, radius=1.0),
            _make_face(face_id=20, radius=3.0),
            _make_face(face_id=30, radius=2.0),
        ]
        assert FeatureDetector._select_primary_face(faces) == 20

    def test_single_face_returns_its_id(self):
        faces = [_make_face(face_id=42, radius=1.5)]
        assert FeatureDetector._select_primary_face(faces) == 42

    def test_empty_group_returns_zero(self):
        assert FeatureDetector._select_primary_face([]) == 0


class TestBuildCandidate:
    """Tests for FeatureDetector._build_candidate."""

    def test_builds_cylindrical_candidate(self):
        detector = FeatureDetector(MagicMock())
        faces = [
            _make_face(
                face_id=5, surface_type="Cylinder", radius=1.5,
                cx=1.0, cy=2.0, cz=3.0,
                ax=0.0, ay=0.0, az=1.0,
                bbox_min=(0.0, 0.0, 0.0), bbox_max=(3.0, 3.0, 6.0),
            ),
            _make_face(
                face_id=6, surface_type="Cylinder", radius=1.5,
                cx=1.0, cy=2.0, cz=5.0,
                ax=0.0, ay=0.0, az=1.0,
                bbox_min=(0.0, 0.0, 3.0), bbox_max=(3.0, 3.0, 8.0),
            ),
        ]
        cand = detector._build_candidate(faces, "comp1", "solid1", 0)

        assert isinstance(cand, SimplificationCandidate)
        assert cand.feature_type == "cylindrical"
        assert cand.bounding_radius_mm == 1.5
        assert cand.component == "comp1"
        assert cand.shape_name == "solid1"
        assert cand.face_ids == [5, 6]
        assert cand.primary_face_id in (5, 6)
        assert cand.depth_mm > 0
        assert cand.axis == (0.0, 0.0, 1.0)
        assert cand.centroid == (1.0, 2.0, 4.0)
        assert "cylindrical" in cand.description

    def test_builds_threaded_candidate(self):
        detector = FeatureDetector(MagicMock())
        faces = [
            _make_face(face_id=1, surface_type="Cylinder", radius=1.0,
                       ax=0.0, ay=0.0, az=1.0),
            _make_face(face_id=2, surface_type="Cone", radius=2.0,
                       ax=0.0, ay=0.0, az=1.0),
        ]
        cand = detector._build_candidate(faces, "comp1", "solid1", 1)

        assert cand.feature_type == "threaded"
        assert cand.bounding_radius_mm == 2.0
        assert cand.primary_face_id == 2


class TestDetectHoles:
    """Tests for FeatureDetector.detect_holes — the main detection pipeline."""

    def _make_detector_with_mocks(
        self, solids, faces_by_solid
    ):
        """Create a FeatureDetector with mocked _enumerate_solids and
        _query_face_properties so we don't need a real CST connection."""
        conn = MagicMock()
        detector = FeatureDetector(conn)
        detector._enumerate_solids = MagicMock(return_value=solids)

        def query_faces(component, solid):
            return faces_by_solid.get((component, solid), [])

        detector._query_face_properties = MagicMock(side_effect=query_faces)
        return detector

    def test_returns_empty_when_no_solids(self):
        """Req 1.7: empty list when no hole-like features found."""
        detector = self._make_detector_with_mocks([], {})
        result = detector.detect_holes()
        assert result == []

    def test_returns_empty_when_no_hole_faces(self):
        """Req 1.7: solid with only planar faces → no candidates."""
        solids = [("comp1", "solid1")]
        faces = {
            ("comp1", "solid1"): [
                _make_face(surface_type="Plane", radius=0.0),
            ]
        }
        detector = self._make_detector_with_mocks(solids, faces)
        result = detector.detect_holes()
        assert result == []

    def test_filters_by_max_radius(self):
        """Req 2.3, 2.4: candidates with radius > max_radius excluded."""
        solids = [("comp1", "solid1")]
        faces = {
            ("comp1", "solid1"): [
                # Small hole (r=1.0) — should pass default 3.0 filter
                _make_face(face_id=0, surface_type="Cylinder", radius=1.0,
                           cx=0.0, cy=0.0, cz=1.0,
                           ax=0.0, ay=0.0, az=1.0),
                _make_face(face_id=1, surface_type="Cylinder", radius=1.0,
                           cx=0.0, cy=0.0, cz=3.0,
                           ax=0.0, ay=0.0, az=1.0),
                # Large hole (r=5.0) — should be filtered out
                _make_face(face_id=2, surface_type="Cylinder", radius=5.0,
                           cx=20.0, cy=20.0, cz=1.0,
                           ax=0.0, ay=0.0, az=1.0),
                _make_face(face_id=3, surface_type="Cylinder", radius=5.0,
                           cx=20.0, cy=20.0, cz=3.0,
                           ax=0.0, ay=0.0, az=1.0),
            ]
        }
        detector = self._make_detector_with_mocks(solids, faces)
        result = detector.detect_holes(max_radius_mm=3.0)
        assert all(c.bounding_radius_mm <= 3.0 for c in result)
        assert len(result) == 1
        assert result[0].bounding_radius_mm == 1.0

    def test_includes_candidate_at_exact_max_radius(self):
        """Req 2.4: radius == max_radius is included."""
        solids = [("comp1", "solid1")]
        faces = {
            ("comp1", "solid1"): [
                _make_face(face_id=0, surface_type="Cylinder", radius=3.0,
                           cx=0.0, cy=0.0, cz=0.0,
                           ax=0.0, ay=0.0, az=1.0),
            ]
        }
        detector = self._make_detector_with_mocks(solids, faces)
        result = detector.detect_holes(max_radius_mm=3.0)
        assert len(result) == 1
        assert result[0].bounding_radius_mm == 3.0

    def test_candidates_include_component_name(self):
        """Req 8.2, 8.3: each candidate includes parent component name."""
        solids = [("myComponent", "solid1")]
        faces = {
            ("myComponent", "solid1"): [
                _make_face(face_id=0, surface_type="Cylinder", radius=1.0,
                           component="myComponent",
                           cx=0.0, cy=0.0, cz=0.0,
                           ax=0.0, ay=0.0, az=1.0),
            ]
        }
        detector = self._make_detector_with_mocks(solids, faces)
        result = detector.detect_holes()
        assert len(result) == 1
        assert result[0].component == "myComponent"

    def test_sorted_by_component_then_radius(self):
        """Results sorted by component name ascending, then radius ascending."""
        solids = [("compB", "s1"), ("compA", "s2")]
        faces = {
            ("compB", "s1"): [
                _make_face(face_id=0, surface_type="Cylinder", radius=2.0,
                           component="compB",
                           cx=0.0, cy=0.0, cz=0.0,
                           ax=0.0, ay=0.0, az=1.0),
            ],
            ("compA", "s2"): [
                _make_face(face_id=1, surface_type="Cylinder", radius=1.0,
                           component="compA",
                           cx=10.0, cy=10.0, cz=0.0,
                           ax=0.0, ay=0.0, az=1.0),
            ],
        }
        detector = self._make_detector_with_mocks(solids, faces)
        result = detector.detect_holes(max_radius_mm=5.0)
        assert len(result) == 2
        assert result[0].component == "compA"
        assert result[1].component == "compB"

    def test_multi_component_detection(self):
        """Req 8.1: enumerate all solids across all components."""
        solids = [("comp1", "s1"), ("comp2", "s2")]
        faces = {
            ("comp1", "s1"): [
                _make_face(face_id=0, surface_type="Cylinder", radius=1.0,
                           component="comp1",
                           cx=0.0, cy=0.0, cz=0.0,
                           ax=0.0, ay=0.0, az=1.0),
            ],
            ("comp2", "s2"): [
                _make_face(face_id=1, surface_type="Cylinder", radius=1.5,
                           component="comp2",
                           cx=10.0, cy=10.0, cz=0.0,
                           ax=0.0, ay=0.0, az=1.0),
            ],
        }
        detector = self._make_detector_with_mocks(solids, faces)
        result = detector.detect_holes(max_radius_mm=5.0)
        assert len(result) == 2
        components = {c.component for c in result}
        assert components == {"comp1", "comp2"}
