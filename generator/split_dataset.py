import os, random, shutil

random.seed(7)
SRC = "/home/claude/task3/dataset_raw"
DST = "/home/claude/task3/Custom_Microplastic_Dataset"
CLASSES = open(os.path.join(SRC, "classes.txt")).read().splitlines()

images = sorted(f for f in os.listdir(os.path.join(SRC, "images")) if f.endswith(".png"))
random.shuffle(images)

n = len(images)
n_train = int(n * 0.70)
n_val = int(n * 0.15)
splits = {
    "train": images[:n_train],
    "val": images[n_train:n_train + n_val],
    "test": images[n_train + n_val:],
}

for split, files in splits.items():
    os.makedirs(os.path.join(DST, split, "images"), exist_ok=True)
    os.makedirs(os.path.join(DST, split, "labels"), exist_ok=True)
    for fname in files:
        base = fname[:-4]
        shutil.copy(os.path.join(SRC, "images", fname), os.path.join(DST, split, "images", fname))
        shutil.copy(os.path.join(SRC, "labels", base + ".txt"), os.path.join(DST, split, "labels", base + ".txt"))
    print(f"{split}: {len(files)} images")

# data.yaml (Ultralytics YOLO format)
yaml_content = f"""# Custom Microplastic Detection Dataset (Task 3)
# Synthetic subset generated procedurally; combine with downloaded secondary
# datasets (see dataset_documentation.md) under the same folder convention.
path: {DST}
train: train/images
val: val/images
test: test/images

nc: {len(CLASSES)}
names: {CLASSES}
"""
with open(os.path.join(DST, "data.yaml"), "w") as f:
    f.write(yaml_content)

shutil.copy(os.path.join(SRC, "classes.txt"), os.path.join(DST, "classes.txt"))
print("data.yaml written to", DST)
