"""
Stage 2 Classifier Retraining
==============================
Trains SVM and Random Forest classifiers on SPECIFIC plastic types only:
    fiber, film, fragment, pallet

Deliberately EXCLUDES the generic 'microplastic' folder from training because
it adds noise (those images are unlabelled by type, hurting classification accuracy).

Usage:
    python scripts/retrain_classifier.py

Output:
    models/svm_classifier.pkl      - SVM classifier (specific classes only)
    models/rf_classifier.pkl       - Random Forest classifier (specific classes only)
    models/classifier_classes.txt  - Class names in sorted order (for inference)
"""

import os
import cv2
import numpy as np
from sklearn.svm import SVC
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import classification_report, accuracy_score, confusion_matrix
from skimage.feature import hog
import joblib
from pathlib import Path

# ── Config ──────────────────────────────────────────────────────────────────
DATA_DIR = Path(__file__).parent.parent / "classification_dataset"
MODELS_DIR = Path(__file__).parent.parent / "models"
IMG_SIZE = (128, 128)

# Only train on specific sub-types — NOT the generic "microplastic" bucket
SPECIFIC_CLASSES = ["fiber", "film", "fragment", "pallet"]
# ────────────────────────────────────────────────────────────────────────────


def extract_features(image_path):
    """HOG (shape) + Color Histogram (color) concatenated into one feature vector."""
    img = cv2.imread(str(image_path))
    if img is None:
        return None
    img = cv2.resize(img, IMG_SIZE)

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    hog_feat = hog(gray, orientations=9, pixels_per_cell=(8, 8),
                   cells_per_block=(2, 2), block_norm='L2-Hys', transform_sqrt=True)

    hist = cv2.calcHist([img], [0, 1, 2], None, [8, 8, 8],
                        [0, 256, 0, 256, 0, 256])
    hist = cv2.normalize(hist, hist).flatten()

    return np.hstack((hog_feat, hist))


def load_split(split):
    """Load images for a given split, only for SPECIFIC_CLASSES."""
    split_dir = DATA_DIR / split
    X, y = [], []

    for cls_idx, cls_name in enumerate(SPECIFIC_CLASSES):
        cls_dir = split_dir / cls_name
        if not cls_dir.exists():
            print(f"  [WARN] {split}/{cls_name} not found, skipping.")
            continue

        files = list(cls_dir.glob("*.jpg")) + list(cls_dir.glob("*.png"))
        print(f"  {split}/{cls_name}: {len(files)} images")

        for img_path in files:
            feat = extract_features(img_path)
            if feat is not None:
                X.append(feat)
                y.append(cls_idx)

    return np.array(X), np.array(y)


def main():
    MODELS_DIR.mkdir(exist_ok=True)

    print(f"Classes to train: {SPECIFIC_CLASSES}\n")

    # ── Load data ──
    print("Loading train split...")
    X_train, y_train = load_split("train")

    print("\nLoading val split...")
    X_val, y_val = load_split("val")

    print(f"\nTrain samples : {len(X_train)}")
    print(f"Val samples   : {len(X_val)}")
    print(f"Feature size  : {X_train.shape[1]}\n")

    # ── Train SVM ──
    print("=" * 50)
    print("Training SVM (RBF kernel)...")
    svm = SVC(kernel='rbf', C=10.0, gamma='scale', probability=True, random_state=42)
    svm.fit(X_train, y_train)

    svm_preds = svm.predict(X_val)
    svm_acc = accuracy_score(y_val, svm_preds)
    print(f"SVM Val Accuracy: {svm_acc:.4f}")
    print(classification_report(y_val, svm_preds, target_names=SPECIFIC_CLASSES))

    svm_path = MODELS_DIR / "svm_classifier.pkl"
    joblib.dump(svm, svm_path)
    print(f"Saved: {svm_path}")

    # ── Train Random Forest ──
    print("=" * 50)
    print("Training Random Forest (200 trees)...")
    rf = RandomForestClassifier(n_estimators=200, random_state=42, n_jobs=-1)
    rf.fit(X_train, y_train)

    rf_preds = rf.predict(X_val)
    rf_acc = accuracy_score(y_val, rf_preds)
    print(f"Random Forest Val Accuracy: {rf_acc:.4f}")
    print(classification_report(y_val, rf_preds, target_names=SPECIFIC_CLASSES))

    rf_path = MODELS_DIR / "rf_classifier.pkl"
    joblib.dump(rf, rf_path)
    print(f"Saved: {rf_path}")

    # ── Save class names (needed at inference time) ──
    classes_path = MODELS_DIR / "classifier_classes.txt"
    classes_path.write_text("\n".join(SPECIFIC_CLASSES))
    print(f"Saved class list: {classes_path}")

    print("\n[DONE] Best model to use for Stage 2:")
    if svm_acc >= rf_acc:
        print(f"  -> SVM  (val acc: {svm_acc:.4f}) -> models/svm_classifier.pkl")
    else:
        print(f"  -> Random Forest (val acc: {rf_acc:.4f}) -> models/rf_classifier.pkl")


if __name__ == "__main__":
    os.chdir(Path(__file__).parent)
    main()
