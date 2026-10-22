import os
import cv2
import numpy as np
import joblib
from skimage.feature import hog

def extract_features(image_path, img_size=(128, 128)):
    img = cv2.imread(str(image_path))
    if img is None: return None
    img = cv2.resize(img, img_size)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    hog_features = hog(gray, orientations=9, pixels_per_cell=(8, 8),
                       cells_per_block=(2, 2), block_norm='L2-Hys', transform_sqrt=True)
    hist = cv2.calcHist([img], [0, 1, 2], None, [8, 8, 8], [0, 256, 0, 256, 0, 256])
    hist = cv2.normalize(hist, hist).flatten()
    return np.hstack((hog_features, hist))

def test_ml(image_path):
    print("\n--- Testing ML Model (SVM) ---")
    try:
        svm_model = joblib.load('../models/svm_model.pkl')
        features = extract_features(image_path)
        if features is not None:
            classes = sorted(os.listdir('../classification_dataset/train'))
            pred_idx = svm_model.predict([features])[0]
            print(f"SVM Prediction: {classes[pred_idx]}")
        else:
            print("Failed to load image for ML.")
    except Exception as e:
        print(f"Error testing ML: {e}")

def test_rf(image_path):
    print("\n--- Testing ML Model (Random Forest) ---")
    try:
        rf_model = joblib.load('../models/rf_model.pkl')
        features = extract_features(image_path)
        if features is not None:
            classes = sorted(os.listdir('../classification_dataset/train'))
            pred_idx = rf_model.predict([features])[0]
            print(f"Random Forest Prediction: {classes[pred_idx]}")
        else:
            print("Failed to load image for ML.")
    except Exception as e:
        print(f"Error testing Random Forest: {e}")

if __name__ == '__main__':
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    test_image = r'C:\Users\suhai\Downloads\Custom_Microplastic_Dataset\Custom_Microplastic_Dataset\kueranan\valid\images\104_png.rf.ec4057e0614b9812ad467c5bf81847cf.jpg'
    
    print(f"Testing on image: {test_image}")
    test_ml(test_image)
    test_rf(test_image)
