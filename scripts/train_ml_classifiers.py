import os
import cv2
import numpy as np
from sklearn.svm import SVC
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import classification_report, accuracy_score
from skimage.feature import hog
import joblib
import argparse
from pathlib import Path

def extract_features(image_path, img_size=(128, 128)):
    """Extract HOG features and Color Histogram from an image."""
    img = cv2.imread(str(image_path))
    if img is None:
        return None
        
    img = cv2.resize(img, img_size)
    
    # 1. HOG Features (Shape)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    hog_features = hog(gray, orientations=9, pixels_per_cell=(8, 8),
                       cells_per_block=(2, 2), block_norm='L2-Hys', transform_sqrt=True)
                       
    # 2. Color Histogram (Color)
    hist = cv2.calcHist([img], [0, 1, 2], None, [8, 8, 8], [0, 256, 0, 256, 0, 256])
    hist = cv2.normalize(hist, hist).flatten()
    
    # Combine features
    return np.hstack((hog_features, hist))

def load_dataset(data_dir, split):
    split_dir = Path(data_dir) / split
    X, y = [], []
    
    classes = [d.name for d in split_dir.iterdir() if d.is_dir()]
    class_to_idx = {cls_name: i for i, cls_name in enumerate(classes)}
    
    print(f"Loading {split} dataset from {split_dir}...")
    for cls_name in classes:
        cls_dir = split_dir / cls_name
        for img_path in cls_dir.glob('*.jpg'):
            features = extract_features(img_path)
            if features is not None:
                X.append(features)
                y.append(class_to_idx[cls_name])
                
    return np.array(X), np.array(y), classes

def train_ml_models(data_dir):
    X_train, y_train, classes = load_dataset(data_dir, 'train')
    X_val, y_val, _ = load_dataset(data_dir, 'val')
    
    print(f"Extracted {X_train.shape[1]} features per image.")
    print(f"Training set: {X_train.shape[0]} samples")
    print(f"Validation set: {X_val.shape[0]} samples")
    
    # 1. Train SVM
    print("\n--- Training Support Vector Machine (SVM) ---")
    svm_clf = SVC(kernel='rbf', C=1.0, random_state=42)
    svm_clf.fit(X_train, y_train)
    
    svm_preds = svm_clf.predict(X_val)
    print(f"SVM Accuracy: {accuracy_score(y_val, svm_preds):.4f}")
    print(classification_report(y_val, svm_preds, target_names=classes))
    
    # Save SVM
    joblib.dump(svm_clf, '../models/svm_model.pkl')
    
    # 2. Train Random Forest
    print("\n--- Training Random Forest ---")
    rf_clf = RandomForestClassifier(n_estimators=100, random_state=42)
    rf_clf.fit(X_train, y_train)
    
    rf_preds = rf_clf.predict(X_val)
    print(f"Random Forest Accuracy: {accuracy_score(y_val, rf_preds):.4f}")
    print(classification_report(y_val, rf_preds, target_names=classes))
    
    # Save Random Forest
    joblib.dump(rf_clf, '../models/rf_model.pkl')
    
    print("\nModels saved to ../models/ (svm_model.pkl, rf_model.pkl)")

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Train ML Classifiers (SVM, Random Forest)')
    parser.add_argument('--data', type=str, default='../classification_dataset',
                        help='Path to dataset')
    args = parser.parse_args()
    
    # Change current working directory to the script's directory
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    
    train_ml_models(args.data)
