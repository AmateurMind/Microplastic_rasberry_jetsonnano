from ultralytics import YOLO
import os

def main():
    # Load a pre-trained YOLOv8 model (nano version is best for Jetson Nano)
    model = YOLO("yolov8n.pt")
    
    # Path to your merged dataset configuration
    data_yaml_path = os.path.abspath(os.path.join("merged_dataset", "data.yaml"))
    
    # Train the model
    # adjust epochs and batch size based on your hardware capabilities
    results = model.train(
        data=data_yaml_path,
        epochs=50,          # number of training epochs
        imgsz=640,          # image size
        batch=16,           # batch size
        name="microplastic_detector", # name of the experiment
        device="0"          # Use "0" for GPU, or "cpu" if no GPU is available
    )
    
    print("Training complete! The best model weights are saved in 'runs/detect/microplastic_detector/weights/best.pt'")

if __name__ == "__main__":
    main()
