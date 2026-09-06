import os
from ultralytics import YOLO

def train_yolov11():
    # Load a pre-trained YOLOv11 model (e.g., small version)
    # Note: Ultralytics currently maps 'yolo11' via the same interface.
    model = YOLO("yolo11n.pt") 
    
    # Path to dataset configuration
    data_yaml = "merged_dataset/data.yaml" # Update with your dataset path if different
    
    if not os.path.exists(data_yaml):
        print(f"Error: {data_yaml} not found. Please ensure the dataset path is correct.")
        return

    # Train the model
    results = model.train(
        data=data_yaml,
        epochs=50,
        imgsz=640,
        batch=16,
        project="models/yolov11",
        name="yolov11_microplastics",
        device="0" # Use GPU if available, else omit or set to "cpu"
    )
    
    print("YOLOv11 training completed.")

if __name__ == "__main__":
    train_yolov11()
