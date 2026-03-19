"""CLI entry point for the CST CAD Model Simplifier.

Connects to CST Studio Suite 2025, detects hole seed faces via SAT
export/parsing, builds adjacency graph, and runs progressive hole
removal with interactive confirmation.

Validates: Requirements 9.1, 9.2, 9.3, 9.4, 6.1, 6.5, 1.1, 3.10, 4.5
"""

import argparse
import logging
import sys

from code.cst_connection import CSTConnection, CSTConnectionError
from code.feature_detector import FeatureDetector
from code.simplifier import Simplifier

logger = logging.getLogger(__name__)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="CST CAD Model Simplifier — detect and fill holes in STP-imported models.",
    )
    parser.add_argument(
        "--project", required=True,
        help="Path to the .cst project file.",
    )
    parser.add_argument(
        "--auto", action="store_true",
        help="Non-interactive mode: fill all holes without prompting.",
    )
    return parser


def main(argv: list[str] | None = None) -> None:
    """Run the CST CAD Model Simplifier."""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )

    parser = build_parser()
    args = parser.parse_args(argv)

    connection = CSTConnection()
    try:
        print("Connecting to CST Studio Suite 2025...")
        connection.connect()

        print(f"Opening project: {args.project}")
        connection.open_project(args.project)

        print("Detecting holes...")
        print("  (Exporting SAT, parsing topology, building adjacency...)")
        detector = FeatureDetector(connection)
        solid_data = detector.detect_seeds()

        if not solid_data:
            print("No hole candidates found. Nothing to simplify.")
            return

        total_seeds = sum(len(s["seeds"]) for s in solid_data)
        print(f"Found {total_seeds} seed face(s) across {len(solid_data)} solid(s).\n")

        simplifier = Simplifier(connection)

        for data in solid_data:
            shape = data["shape_name"]
            print(f"\nProcessing {shape} ({len(data['seeds'])} seeds)...")

            if args.auto:
                simplifier.run_auto_workflow(
                    shape, data["seeds"], data["adjacency"],
                    data["bboxes"], data["face_types"],
                    seed_groups=data.get("seed_groups"),
                )
            else:
                simplifier.run_sequential_workflow(
                    shape, data["seeds"], data["adjacency"],
                    data["bboxes"], data["face_types"],
                    seed_groups=data.get("seed_groups"),
                )

    except CSTConnectionError as exc:
        logger.error("CST connection error: %s", exc)
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)
    except KeyboardInterrupt:
        print("\nInterrupted by user.")
    finally:
        connection.close()


if __name__ == "__main__":
    main()
