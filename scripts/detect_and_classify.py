"""
detect_and_classify.py  —  Stage 1 + Stage 2 Unified Inference
===============================================================
Stage 1: YOLO detects microplastic particles (bounding boxes)
Stage 2: SVM/RF classifier identifies the specific type of each particle

Works on:
  - A single image file
  - A folder of images
  - A webcam / video file (real-time)

Usage examples:
    # Single image
    python scripts/detect_and_classify.py --source path/to/image.jpg

    # Folder of images
    python scripts/detect_and_classify.py --source path/to/folder/

    # Webcam (real-time, good for Jetson / Raspberry Pi)
    python scripts/detect_and_classify.py --source 0

    # Video file
    python scripts/detect_and_classify.py --source path/to/video.mp4

    # Choose classifier (default: svm)
    python scripts/detect_and_classify.py --source 0 --classifier rf

    # Custom YOLO model weights
    python scripts/detect_and_classify.py --source 0 --yolo_model runs/detect/best.pt
"""

import os
import cv2
import numpy as np
import joblib
import argparse
from pathlib import Path
from skimage.feature import hog

# Try importing YOLO (ultralytics)
try:
    from ultralytics import YOLO
except ImportError:
    raise ImportError("ultralytics not installed. Run: pip install ultralytics")

# ── Paths ────────────────────────────────────────────────────────────────────
ROOT = Path(__file__).parent.parent
MODELS_DIR = ROOT / "models"
CLASSES_FILE = MODELS_DIR / "classifier_classes.txt"

# Fallback: if retrained classifier doesn't exist, use original
SVM_PATH = MODELS_DIR / "svm_classifier.pkl"
RF_PATH = MODELS_DIR / "rf_classifier.pkl"

# Default fallback YOLO model (replace with your trained best.pt after training)
DEFAULT_YOLO = str(ROOT / "yolo11n.pt")

IMG_SIZE = (128, 128)

# Colors per class (BGR) for display
CLASS_COLORS = {
    "fiber":    (255, 100, 0),
    "film":     (0, 200, 255),
    "fragment": (0, 255, 100),
    "pallet":   (200, 0, 255),
    "foam":     (0, 165, 255),
    "unknown":  (150, 150, 150),
}
# ─────────────────────────────────────────────────────────────────────────────


def load_classifier(clf_type):
    """Load the Stage 2 classifier and class names."""
    path = SVM_PATH if clf_type == "svm" else RF_PATH
    if not path.exists():
        raise FileNotFoundError(
            f"Classifier not found: {path}\n"
            f"Run: python scripts/retrain_classifier.py"
        )
    clf = joblib.load(path)

    # Load class names
    if CLASSES_FILE.exists():
        classes = CLASSES_FILE.read_text().strip().splitlines()
    else:
        # Fallback: sorted train dirs (skipping microplastic)
        train_dir = ROOT / "classification_dataset" / "train"
        classes = sorted([
            d.name for d in train_dir.iterdir()
            if d.is_dir() and d.name != "microplastic"
        ])

    print(f"Classifier classes: {classes}")
    return clf, classes


def extract_features(img_bgr):
    """Extract HOG + Color Histogram from an OpenCV BGR crop."""
    img = cv2.resize(img_bgr, IMG_SIZE)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    hog_feat = hog(gray, orientations=9, pixels_per_cell=(8, 8),
                   cells_per_block=(2, 2), block_norm='L2-Hys', transform_sqrt=True)
    hist = cv2.calcHist([img], [0, 1, 2], None, [8, 8, 8],
                        [0, 256, 0, 256, 0, 256])
    hist = cv2.normalize(hist, hist).flatten()
    return np.hstack((hog_feat, hist))


def classify_crop(crop_bgr, clf, classes):
    """Run Stage 2 classifier on a cropped region. Returns (label, confidence)."""
    if crop_bgr.size == 0:
        return "unknown", 0.0

    feat = extract_features(crop_bgr)
    pred_idx = clf.predict([feat])[0]
    label = classes[pred_idx]

    # Confidence (probability) — SVC needs probability=True, RF has it natively
    try:
        proba = clf.predict_proba([feat])[0]
        confidence = float(proba[pred_idx])
    except Exception:
        confidence = 1.0

    return label, confidence


def draw_result(frame, box, label, confidence):
    """Draw bounding box + label on frame."""
    x1, y1, x2, y2 = map(int, box)
    color = CLASS_COLORS.get(label, CLASS_COLORS["unknown"])

    # Box
    cv2.rectangle(frame, (x1, y1), (x2, y2), color, 2)

    # Label background + text
    text = f"{label} {confidence:.0%}"
    (tw, th), _ = cv2.getTextSize(text, cv2.FONT_HERSHEY_SIMPLEX, 0.55, 1)
    cv2.rectangle(frame, (x1, y1 - th - 6), (x1 + tw + 4, y1), color, -1)
    cv2.putText(frame, text, (x1 + 2, y1 - 3),
                cv2.FONT_HERSHEY_SIMPLEX, 0.55, (255, 255, 255), 1, cv2.LINE_AA)
    return frame


def process_frame(frame, yolo_model, clf, classes, conf_threshold=0.3):
    """Run Stage 1 (YOLO) then Stage 2 (Classifier) on a single frame."""
    results = yolo_model.predict(frame, conf=conf_threshold, verbose=False)
    detections = results[0].boxes

    particle_count = 0
    type_counts = {}

    if detections is not None and len(detections):
        for box in detections.xyxy:
            x1, y1, x2, y2 = map(int, box)

            # Clamp to frame bounds
            x1, y1 = max(0, x1), max(0, y1)
            x2, y2 = min(frame.shape[1], x2), min(frame.shape[0], y2)

            # ── Stage 2: Classify the crop ──
            crop = frame[y1:y2, x1:x2]
            label, confidence = classify_crop(crop, clf, classes)

            # Draw on frame
            draw_result(frame, (x1, y1, x2, y2), label, confidence)

            particle_count += 1
            type_counts[label] = type_counts.get(label, 0) + 1

    # ── Status overlay (top-left) ──
    cv2.putText(frame, f"Particles: {particle_count}", (10, 28),
                cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 0), 2)
    y_offset = 55
    for cls_name, cnt in type_counts.items():
        color = CLASS_COLORS.get(cls_name, (255, 255, 255))
        cv2.putText(frame, f"  {cls_name}: {cnt}", (10, y_offset),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.55, color, 1)
        y_offset += 22

    return frame, particle_count, type_counts


def run_on_images(source_path, yolo_model, clf, classes, conf_threshold, save_dir):
    """Process a single image or a folder of images."""
    source = Path(source_path)
    img_paths = []

    if source.is_file():
        img_paths = [source]
    elif source.is_dir():
        img_paths = sorted(source.glob("*.jpg")) + sorted(source.glob("*.png"))

    print(f"\nProcessing {len(img_paths)} image(s)...\n")
    save_dir.mkdir(parents=True, exist_ok=True)

    for img_path in img_paths:
        frame = cv2.imread(str(img_path))
        if frame is None:
            print(f"[WARN] Cannot read {img_path}, skipping.")
            continue

        frame, count, types = process_frame(frame, yolo_model, clf, classes, conf_threshold)

        out_path = save_dir / img_path.name
        cv2.imwrite(str(out_path), frame)
        print(f"{img_path.name}: {count} particle(s) → {types}  → saved to {out_path}")


def run_realtime(source, yolo_model, clf, classes, conf_threshold):
    """Process webcam or video stream in real-time."""
    # source = 0 for webcam, or a video file path
    try:
        src = int(source)  # webcam index
    except ValueError:
        src = source       # video file path

    cap = cv2.VideoCapture(src)
    if not cap.isOpened():
        raise RuntimeError(f"Cannot open video source: {source}")

    print("Running real-time detection. Press 'q' to quit.\n")
    while True:
        ret, frame = cap.read()
        if not ret:
            break

        frame, count, types = process_frame(frame, yolo_model, clf, classes, conf_threshold)

        cv2.imshow("Microplastic Detector", frame)
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    cap.release()
    cv2.destroyAllWindows()
    print("Stopped.")


def main():
    parser = argparse.ArgumentParser(description="Microplastic Detect + Classify (2-Stage Pipeline)")
    parser.add_argument("--source",     type=str, required=True,
                        help="Image file, image folder, video file, or webcam index (e.g. 0)")
    parser.add_argument("--yolo_model", type=str, default=DEFAULT_YOLO,
                        help="Path to YOLO .pt weights (default: yolo11n.pt pretrained)")
    parser.add_argument("--classifier", type=str, default="svm", choices=["svm", "rf"],
                        help="Which Stage 2 classifier to use: svm (default) or rf")
    parser.add_argument("--conf",       type=float, default=0.3,
                        help="YOLO detection confidence threshold (default: 0.3)")
    parser.add_argument("--save_dir",   type=str, default="runs/detect_classify",
                        help="Output folder for annotated images")
    args = parser.parse_args()

    # ── Load models ──
    print(f"Loading YOLO from: {args.yolo_model}")
    yolo = YOLO(args.yolo_model)

    print(f"Loading {args.classifier.upper()} classifier...")
    clf, classes = load_classifier(args.classifier)

    save_dir = ROOT / args.save_dir

    # ── Route to correct mode ──
    source = args.source
    is_webcam_or_video = (
        source.isdigit() or
        source.endswith((".mp4", ".avi", ".mov", ".mkv"))
    )

    if is_webcam_or_video:
        run_realtime(source, yolo, clf, classes, args.conf)
    else:
        run_on_images(source, yolo, clf, classes, args.conf, save_dir)


if __name__ == "__main__":
    os.chdir(Path(__file__).parent.parent)
    main()
