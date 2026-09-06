# Microplastic Detection and Classification: Model Comparison

This document provides a comprehensive comparison of all the object detection and classification models planned for the microplastic monitoring system.

## 1. Object Detection Models

These models are responsible for finding and localizing microplastics (bounding boxes or masks) in an image.

| Model | Architecture Type | Pros for Microplastics | Cons | Best Use Case |
| :--- | :--- | :--- | :--- | :--- |
| **YOLOv8** | One-stage, Anchor-free | Very fast, highly optimized for edge devices, good baseline accuracy. | Can struggle with extremely small, clustered particles compared to two-stage detectors. | Real-time edge deployment (Raspberry Pi/Jetson). |
| **YOLOv11** | One-stage, Anchor-free | Excellent speed/accuracy tradeoff, improved small object detection via refined attention. | Newer architecture, might require specific library versions (Ultralytics). | High-accuracy real-time edge deployment. |
| **Faster R-CNN** | Two-stage (Region Proposal) | Very high accuracy, especially for small and densely packed objects (common in microplastics). | Slow inference speed, computationally heavy, difficult to run on edge devices without heavy optimization. | Cloud-based or Desktop GPU processing where accuracy is paramount. |
| **SSD (Single Shot Detector)** | One-stage, Anchor-based | Faster than Faster R-CNN, handles multiple scales well. | Lower accuracy on small objects compared to YOLO and Faster R-CNN. | Legacy edge devices or when simple deployment is needed. |
| **Mask R-CNN** | Two-stage (Instance Seg.) | Provides pixel-level segmentation (masks) in addition to bounding boxes, excellent for shape analysis (e.g., distinguishing fragments vs. films). | Slowest inference, requires high computational power and polygon annotations instead of just bounding boxes. | Detailed morphological analysis of microplastics in a lab setting. |

## 2. Classification Models (Machine Learning)

These models are used to classify pre-cropped images of microplastics or extracted features into predefined classes (Fibers, Fragments, Films, Foams, Pellets).

| Algorithm | Pros | Cons |
| :--- | :--- | :--- |
| **SVM (Support Vector Machine)** | Effective in high-dimensional spaces, works well with limited data when combined with good feature extractors (HOG, SIFT). | Computationally expensive for large datasets, requires manual feature extraction. |
| **Random Forest** | Robust to overfitting, handles non-linear data well, provides feature importance. | Can be slow to predict with many trees, less effective on raw image pixels without feature extraction. |

## 3. Classification Models (Deep Learning)

Deep learning classifiers automatically extract features and classify the images.

| Model | Pros | Cons |
| :--- | :--- | :--- |
| **Custom CNN** | Lightweight, can be tailored exactly to the dataset, fast inference. | May lack the feature extraction power of pre-trained models on small datasets. |
| **MobileNetV2** | Extremely lightweight, designed for mobile and edge devices, very fast. | Lower accuracy compared to heavier models like ResNet50. |
| **ResNet50** | Very accurate, solves the vanishing gradient problem, widely supported. | Large model size, slower inference on low-power edge devices. |
| **EfficientNet** | Balances depth, width, and resolution for optimal performance; highly accurate and efficient. | Can be slightly slower to train than MobileNet. |
| **Vision Transformer (ViT)** | State-of-the-art accuracy, captures global context of the image excellently. | Requires massive amounts of data to train from scratch, very computationally heavy for inference. |

## Conclusion & Strategy

1.  **Detection**: We will use YOLOv8/YOLOv11 for edge deployment due to their speed. Faster R-CNN and Mask R-CNN will be trained to establish an accuracy benchmark and for detailed lab analysis.
2.  **Classification**: MobileNetV2 and EfficientNet will be prioritized for mobile/edge integration, while ResNet50 and ViT will be evaluated for maximum classification accuracy on desktop GPUs. Machine learning models (SVM, Random Forest) will serve as baseline comparisons using traditional feature extraction methods.
