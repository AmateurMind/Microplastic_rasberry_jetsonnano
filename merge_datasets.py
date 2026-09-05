import os
import shutil

# Datasets to merge
datasets = ["dataset", "kueranan", "roboflow_in_water"]
output_dir = "merged_dataset"
splits = ["train", "valid", "test"]

# Create output directories
for split in splits:
    os.makedirs(os.path.join(output_dir, split, "images"), exist_ok=True)
    os.makedirs(os.path.join(output_dir, split, "labels"), exist_ok=True)
    
    # YOLO sometimes uses 'val' instead of 'valid', so create that symlink/folder just in case or we map it
    os.makedirs(os.path.join(output_dir, "val", "images"), exist_ok=True)
    os.makedirs(os.path.join(output_dir, "val", "labels"), exist_ok=True)

total_images = 0

for dataset_name in datasets:
    if not os.path.exists(dataset_name):
        print(f"Warning: Dataset directory '{dataset_name}' not found. Skipping.")
        continue
        
    for split in splits:
        # Check both 'valid' and 'val' folder names
        split_names_to_check = [split]
        if split == "valid":
            split_names_to_check.append("val")
            
        for split_dir in split_names_to_check:
            img_dir = os.path.join(dataset_name, split_dir, "images")
            lbl_dir = os.path.join(dataset_name, split_dir, "labels")
            
            if os.path.exists(img_dir) and os.path.exists(lbl_dir):
                images = os.listdir(img_dir)
                for img_name in images:
                    if not img_name.lower().endswith(('.png', '.jpg', '.jpeg')):
                        continue
                        
                    # Prefix with dataset name to prevent name collisions
                    new_img_name = f"{dataset_name}_{img_name}"
                    base_name = os.path.splitext(img_name)[0]
                    lbl_name = base_name + ".txt"
                    new_lbl_name = f"{dataset_name}_{lbl_name}"
                    
                    src_img = os.path.join(img_dir, img_name)
                    dst_img = os.path.join(output_dir, split, "images", new_img_name)
                    
                    src_lbl = os.path.join(lbl_dir, lbl_name)
                    dst_lbl = os.path.join(output_dir, split, "labels", new_lbl_name)
                    
                    # Copy image
                    shutil.copy2(src_img, dst_img)
                    
                    # Read, map class, and write label
                    if os.path.exists(src_lbl):
                        with open(src_lbl, "r") as f_in, open(dst_lbl, "w") as f_out:
                            for line in f_in:
                                parts = line.strip().split()
                                if len(parts) >= 5:
                                    class_id = int(parts[0])
                                    
                                    if dataset_name == "kueranan":
                                        # kueranan has 0: fiber, 1: film, 2: fragment, 3: pallet
                                        # We map to 1: fiber, 2: film, 3: fragment, 4: pallet
                                        parts[0] = str(class_id + 1)
                                    else:
                                        # dataset and roboflow_in_water have 0: microplastic
                                        parts[0] = "0"
                                        
                                    f_out.write(" ".join(parts) + "\n")
                    total_images += 1

print(f"Successfully merged {total_images} images into '{output_dir}'.")

# Generate data.yaml
yaml_content = f"""train: train/images
val: valid/images
test: test/images

nc: 6
names: 
  0: microplastic
  1: fiber
  2: film
  3: fragment
  4: pallet
  5: foam
"""

with open(os.path.join(output_dir, "data.yaml"), "w") as f:
    f.write(yaml_content)

print(f"Created {os.path.join(output_dir, 'data.yaml')}")
