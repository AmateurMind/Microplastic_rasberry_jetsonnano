"""Split generated synthetic YOLO data into train/val/test directories."""
import argparse
import os
import random
import shutil


def main():
    parser = argparse.ArgumentParser(description="Split synthetic YOLO data")
    parser.add_argument("--source", required=True, help="Directory containing images/, labels/, and classes.txt")
    parser.add_argument("--output", required=True, help="Destination dataset directory")
    args = parser.parse_args()

    source = os.path.abspath(args.source)
    destination = os.path.abspath(args.output)
    classes_path = os.path.join(source, "classes.txt")
    if not os.path.isfile(classes_path):
        raise SystemExit(f"Missing classes file: {classes_path}")

    with open(classes_path, encoding="utf-8") as handle:
        classes = handle.read().splitlines()
    images_dir = os.path.join(source, "images")
    labels_dir = os.path.join(source, "labels")
    images = sorted(f for f in os.listdir(images_dir) if f.lower().endswith((".png", ".jpg", ".jpeg")))
    if not images:
        raise SystemExit(f"No images found in {images_dir}")

    random.seed(7)
    random.shuffle(images)
    n_train = int(len(images) * 0.70)
    n_val = int(len(images) * 0.15)
    splits = {
        "train": images[:n_train],
        "val": images[n_train:n_train + n_val],
        "test": images[n_train + n_val:],
    }

    for split, files in splits.items():
        split_images = os.path.join(destination, split, "images")
        split_labels = os.path.join(destination, split, "labels")
        os.makedirs(split_images, exist_ok=True)
        os.makedirs(split_labels, exist_ok=True)
        for filename in files:
            base, _ = os.path.splitext(filename)
            label = os.path.join(labels_dir, base + ".txt")
            if not os.path.isfile(label):
                raise SystemExit(f"Missing label for {filename}: {label}")
            shutil.copy2(os.path.join(images_dir, filename), os.path.join(split_images, filename))
            shutil.copy2(label, os.path.join(split_labels, base + ".txt"))
        print(f"{split}: {len(files)} images")

    with open(os.path.join(destination, "classes.txt"), "w", encoding="utf-8") as handle:
        handle.write("\n".join(classes) + "\n")
    yaml_content = f"""# Procedurally generated synthetic microplastic dataset
path: {destination.replace(os.sep, '/')}
train: train/images
val: val/images
test: test/images

nc: {len(classes)}
names: {classes}
"""
    with open(os.path.join(destination, "data.yaml"), "w", encoding="utf-8") as handle:
        handle.write(yaml_content)
    print(f"data.yaml written to {destination}")


if __name__ == "__main__":
    main()
