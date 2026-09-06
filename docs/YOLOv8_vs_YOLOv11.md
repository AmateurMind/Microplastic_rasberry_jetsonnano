# YOLOv8 vs YOLOv11 for Microplastic Detection

Both YOLOv8 and YOLOv11 are state-of-the-art models in the YOLO (You Only Look Once) family, designed for real-time object detection. When applying these models to microplastic detection, several key differences and improvements in YOLOv11 make it a compelling upgrade, although YOLOv8 remains a strong baseline.

## 1. Architecture and Feature Extraction
- **YOLOv8**: Introduced a new backbone and an anchor-free detection head, which improved accuracy and efficiency over previous versions.
- **YOLOv11**: Builds upon the YOLOv8 architecture but introduces more efficient feature extraction modules. It focuses on reducing computational overhead while maintaining or improving Mean Average Precision (mAP). For microplastics, YOLOv11's refined spatial attention mechanisms can better capture small details.

## 2. Detection of Small Objects
- **YOLOv8**: Struggles slightly with very small objects compared to larger ones, which is a common challenge in microplastic detection (e.g., small fibers or fragments).
- **YOLOv11**: Incorporates specific optimizations for small object detection, often utilizing higher-resolution feature maps in the neck of the network. This makes it inherently better at detecting small, barely visible microplastics under a microscope.

## 3. Parameter Efficiency and Speed
- **YOLOv8**: Highly optimized and offers various sizes. Inference is fast, making it suitable for edge devices.
- **YOLOv11**: Achieves a better trade-off between parameters and accuracy. A YOLOv11 model often achieves the same mAP as a larger YOLOv8 model but with fewer parameters, leading to lower latency on edge AI platforms like the Jetson Nano.

## Comparison Summary for Microplastics

| Feature | YOLOv8 | YOLOv11 |
| :--- | :--- | :--- |
| **Small Object Detection** | Good | Excellent |
| **Edge AI Suitability** | High | Very High (Better FPS/mAP ratio) |
| **Training Speed** | Fast | Faster convergence |
| **Architecture** | Anchor-free, CSPDarknet | Refined backbone, enhanced attention |

To compare them in practice, we will evaluate mAP50, Inference Time, and Recall on the microplastics dataset.
