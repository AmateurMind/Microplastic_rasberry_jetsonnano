"""
Remap merged_dataset labels to 1-class (all → 0)
=================================================
The merged_dataset has labels with class IDs 0-5 (from kueranan).
For Stage 1 YOLO detector we only want class 0: microplastic.

This script rewrites every label file in-place so that ALL class IDs become 0.

Usage:
    python scripts/remap_labels_1class.py
"""

import os
from pathlib import Path

MERGED_DIR = Path(__file__).parent.parent / "merged_dataset"
SPLITS = ["train", "valid", "test"]


def remap_labels():
    total_files = 0
    total_changed = 0

    for split in SPLITS:
        label_dir = MERGED_DIR / split / "labels"
        if not label_dir.exists():
            print(f"[SKIP] {split}/labels not found")
            continue

        label_files = list(label_dir.glob("*.txt"))
        print(f"\n{split}: {len(label_files)} label files")

        for lbl_path in label_files:
            total_files += 1
            lines = lbl_path.read_text().splitlines()
            new_lines = []
            changed = False

            for line in lines:
                parts = line.strip().split()
                if len(parts) >= 5 and parts[0] != "0":
                    parts[0] = "0"
                    changed = True
                new_lines.append(" ".join(parts))

            if changed:
                lbl_path.write_text("\n".join(new_lines) + "\n")
                total_changed += 1

    print(f"\n[DONE] Processed {total_files} files.")
    print(f"       Remapped {total_changed} files (had non-zero class IDs).")
    print(f"       All labels are now class 0 (microplastic).")


if __name__ == "__main__":
    remap_labels()
