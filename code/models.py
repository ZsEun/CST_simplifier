"""Core data types and enums for the CST CAD Model Simplifier."""

from dataclasses import dataclass, field
from enum import Enum
from typing import List, Tuple


class HoleType(Enum):
    """Classification of detected hole features.

    Validates: Requirements 1.3, 1.4, 1.5
    """

    CYLINDRICAL = "cylindrical"
    THREADED = "threaded"
    IRREGULAR = "irregular"


@dataclass
class FaceInfo:
    """Properties of a single face queried from the CST geometry kernel.

    Validates: Requirements 1.2, 7.4
    """

    face_id: int
    solid_name: str
    component: str
    surface_type: str
    center_x: float
    center_y: float
    center_z: float
    axis_x: float
    axis_y: float
    axis_z: float
    radius: float
    bbox_min: Tuple[float, float, float]
    bbox_max: Tuple[float, float, float]


@dataclass
class SimplificationCandidate:
    """A detected hole-like feature flagged for potential simplification.

    Validates: Requirements 1.6, 2.3, 2.4, 8.2
    """

    shape_name: str
    component: str
    feature_type: str  # "cylindrical", "threaded", or "irregular"
    bounding_radius_mm: float
    depth_mm: float
    axis: Tuple[float, float, float]
    centroid: Tuple[float, float, float]
    face_ids: List[int] = field(default_factory=list)
    primary_face_id: int = 0
    description: str = ""


@dataclass
class FillResult:
    """Result of a hole fill operation.

    Validates: Requirements 4.3, 4.4
    """

    success: bool
    candidate: SimplificationCandidate
    error_message: str = ""


@dataclass
class SessionSummary:
    """Summary of a simplification session.

    Validates: Requirements 3.10, 4.5
    """

    filled: int = 0
    skipped: int = 0
    undone: int = 0
    failed: int = 0
