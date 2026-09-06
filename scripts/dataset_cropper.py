import os
import cv2
from pathlib import Path
from tqdm import tqdm

def crop_yolo_dataset(base_dir, out_dir, classes):
    """
    Crops objects from YOLO formatted dataset and saves them into class-specific folders.
    """
    splits = [('train', 'train'), ('valid', 'val'), ('test', 'test')]
    
    for split_in, split_out in splits:
        images_dir = Path(base_dir) / split_in / 'images'
        labels_dir = Path(base_dir) / split_in / 'labels'
        
        if not images_dir.exists():
            print(f"Skipping {split_in} as it does not exist.")
            continue
            
        print(f"Processing {split_in} split...")
        image_files = list(images_dir.glob('*.jpg')) + list(images_dir.glob('*.png'))
        
        for img_path in tqdm(image_files):
            label_path = labels_dir / (img_path.stem + '.txt')
            
            if not label_path.exists():
                continue
                
            img = cv2.imread(str(img_path))
            if img is None:
                continue
                
            h, w, _ = img.shape
            
            with open(label_path, 'r') as f:
                lines = f.readlines()
                
            for i, line in enumerate(lines):
                parts = line.strip().split()
                if len(parts) < 5:
                    continue
                    
                class_id = int(parts[0])
                if class_id >= len(classes):
                    continue
                    
                class_name = classes[class_id]
                
                # YOLO format: class_id center_x center_y width height (normalized)
                cx, cy, bw, bh = map(float, parts[1:5])
                
                # Convert to pixel coordinates
                x1 = int((cx - bw / 2) * w)
                y1 = int((cy - bh / 2) * h)
                x2 = int((cx + bw / 2) * w)
                y2 = int((cy + bh / 2) * h)
                
                # Ensure within bounds
                x1, y1 = max(0, x1), max(0, y1)
                x2, y2 = min(w, x2), min(h, y2)
                
                if x2 <= x1 or y2 <= y1:
                    continue
                    
                cropped_img = img[y1:y2, x1:x2]
                
                # Create output directory: out_dir / split / class_name
                out_class_dir = Path(out_dir) / split_out / class_name
                out_class_dir.mkdir(parents=True, exist_ok=True)
                
                # Save cropped image
                out_filename = f"{img_path.stem}_crop_{i}.jpg"
                out_filepath = out_class_dir / out_filename
                cv2.imwrite(str(out_filepath), cropped_img)
                
if __name__ == "__main__":
    base_yolo_dir = "../merged_dataset"
    output_classification_dir = "../classification_dataset"
    
    # Based on merged_dataset/data.yaml
    class_names = ['microplastic', 'fiber', 'film', 'fragment', 'pallet', 'foam']
    
    # Change current working directory to the script's directory to resolve relative paths correctly
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    
    print(f"Cropping YOLO dataset from {base_yolo_dir} to {output_classification_dir}...")
    crop_yolo_dataset(base_yolo_dir, output_classification_dir, class_names)
    print("Cropping complete!")
