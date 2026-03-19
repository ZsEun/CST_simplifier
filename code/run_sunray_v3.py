"""Run full simplification pipeline on Sunray_MB_v3 in auto mode.

Detects holes, filters board edges, then fills all holes via AddToHistory.
Logs results to debug_output.txt.

Run: python -m code.run_sunray_v3
"""

import os
import sys
import logging

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from code.cst_connection import CSTConnection, CSTConnectionError
from code.feature_detector import FeatureDetector
from code.simplifier import Simplifier

PROJECT = r"D:\Users\sunze\Desktop\kiro\cst_simplifier\cst_model\Sunray_MB_v3.cst"
OUT = r"D:\Users\sunze\Desktop\kiro\debug_output.txt"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)


def main():
    lines = []

    conn = CSTConnection()
    try:
        conn.connect()
        conn.open_project(PROJECT)

        detector = FeatureDetector(conn)
        solid_data = detector.detect_seeds()

        if not solid_data:
            lines.append("No hole candidates found.")
        else:
            total_seeds = sum(len(s["seeds"]) for s in solid_data)
            total_groups = sum(len(s.get("seed_groups", [])) for s in solid_data)
            lines.append(f"Found {total_seeds} seed(s) in {total_groups} hole group(s) "
                         f"across {len(solid_data)} solid(s).")

            simplifier = Simplifier(conn)

            for data in solid_data:
                shape = data["shape_name"]
                seeds = data["seeds"]
                lines.append(f"\nProcessing {shape} ({len(seeds)} seeds)...")
                lines.append(f"  Seeds: {seeds}")

                summary = simplifier.run_sequential_workflow(
                    shape, seeds, data["adjacency"],
                    data["bboxes"], data["face_types"],
                    seed_groups=data.get("seed_groups"),
                )

                lines.append(f"  Filled:  {summary.filled}")
                lines.append(f"  Skipped: {summary.skipped}")
                lines.append(f"  Failed:  {summary.failed}")

    except CSTConnectionError as exc:
        lines.append(f"ERROR: {exc}")
    except Exception as exc:
        lines.append(f"UNEXPECTED ERROR: {exc}")
    finally:
        conn.close()

    output = "\n".join(lines)
    print(output)
    with open(OUT, "w") as f:
        f.write(output)
    print(f"\nSaved to {OUT}")


if __name__ == "__main__":
    main()
