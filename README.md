# Microplastic Detection and Classification System

This repository contains the code and resources for an AI-enabled Microplastic Detection and Monitoring System.

## Overview

The goal of this project is to detect and classify microplastic particles in water samples using Computer Vision and Deep Learning techniques. The system is designed to eventually be deployed on Edge AI platforms (like NVIDIA Jetson Nano or Raspberry Pi) for real-time environmental monitoring.

## Dataset

The dataset is a compilation of multiple sources containing images of microplastics in water. The classes have been standardized into 6 categories:

- `0`: microplastic (general)
- `1`: fiber
- `2`: film
- `3`: fragment
- `4`: pallet
- `5`: foam

### Dataset Preparation

To merge the separate datasets into a single unified dataset for training, run:

```bash
python merge_datasets.py
```

This script automatically maps all labels to the 6 classes above and consolidates the images into the `merged_dataset/` directory.

## Models

### YOLOv8 (Object Detection)

We use YOLOv8 as the primary object detection model to find the precise location and class of microplastics in water samples.

**Training:**
To train the YOLOv8 model on the merged dataset, run:

```bash
python train_yolo.py
```

_Note: Make sure your PyTorch installation supports CUDA if you are running this on a GPU._

Training outputs, including the final model weights (`best.pt`) and evaluation graphs (Confusion Matrix, Precision-Recall curves), are automatically saved in the `runs/detect/microplastic_detector/` folder.

## Environment Setup

conda activate yolo

yolo predict model=best.pt

source='C:\Users\suhai\Downloads\Custom_Microplastic_Dataset\Custom_Microplastic_Dataset\valid\images\125_png.rf.ecf0bcc98949c8e55394452b8626409b.jpg'

## SVM AND RANDOM FOREST

PS C:\Users\suhai\Downloads\Custom_Microplastic_Dataset\Custom_Microplastic_Dataset> python scripts/test_predict.py

> > Testing on image: C:\Users\suhai\Downloads\Custom_Microplastic_Dataset\Custom_Microplastic_Dataset\kueranan\valid\images\104_png.rf.ec4057e0614b9812ad467c5bf81847cf.jpg

--- Testing ML Model (SVM) ---
SVM Prediction: microplastic

--- Testing ML Model (Random Forest) ---
Random Forest Prediction: microplastic

1. Create a Conda environment:

```bash
conda create -n yolo python=3.10
conda activate yolo
```

2. Install GPU-enabled PyTorch:

```bash
conda install pytorch torchvision torchaudio pytorch-cuda=11.8 -c pytorch -c nvidia
```

3. Install dependencies:

```bash
pip install ultralytics "numpy<2.0" "opencv-python<4.10"
```
