"""
Master build script — regenerates all artifacts from ficc_data.yaml.

Usage:
    python3 build_all.py           # run all generators
    python3 build_all.py excel     # Excel only
    python3 build_all.py diagrams  # draw.io only
"""

import sys
import subprocess
import os

HERE = os.path.dirname(os.path.abspath(__file__))


def run(label, script):
    print(f"\n{'='*50}")
    print(f"  {label}")
    print(f"{'='*50}")
    result = subprocess.run(
        [sys.executable, os.path.join(HERE, script)],
        cwd=HERE,
    )
    if result.returncode != 0:
        print(f"ERROR: {script} failed (exit {result.returncode})")
        sys.exit(result.returncode)


def main():
    targets = sys.argv[1:] or ["excel", "diagrams"]

    if "excel" in targets:
        run("Building Excel workbook…", "build_excel.py")

    if "diagrams" in targets:
        run("Building draw.io diagrams…", "build_diagrams.py")

    print(f"\n{'='*50}")
    print("  All done.")
    print(f"{'='*50}\n")


if __name__ == "__main__":
    main()
