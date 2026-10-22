import os
from pathlib import Path
from ultralytics import YOLO

def train_yolov11():
    model = YOLO("yolo11n.pt")

    # 1-class yaml: all particles unified as "microplastic"
    # Stage 1 of 2-stage detect+classify pipeline
    root = Path(__file__).parent.parent
    data_yaml = str(root / "merged_dataset" / "data_1class.yaml")

    if not os.path.exists(data_yaml):
        print(f"Error: {data_yaml} not found.")
        return

    results = model.train(
        data=data_yaml,
        epochs=50,
        imgsz=640,
        batch=16,
        project=str(root / "models" / "yolov11"),
        name="microplastic_detector_1class",
        device="0",
        patience=10,
        save=True,
        exist_ok=True,
        single_cls=True,   # force all labels treated as class 0
    )

    best = root / "models" / "yolov11" / "microplastic_detector_1class" / "weights" / "best.pt"
    print(f"\nTraining complete. Best weights: {best}")
    print("Use this model in detect_and_classify.py with --yolo_model", best)

if __name__ == "__main__":
    train_yolov11()
