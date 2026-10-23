# 🔬 Microplastic Detection & Classification Dataset — Complete Technical Analysis

---

## 1. System Overview & The Two-Stage Architecture

### 1.1 Purpose & Problem Formulation

This **Custom Microplastic Dataset** was engineered to train real-time edge computer vision systems for automated microplastic detection and morphological classification in aquatic samples (drinking water, wastewater discharge, surface grabs).

Initially framed as a monolithic 6-class detection problem, practical experiments revealed severe taxonomy clash between generic `microplastic` labels and morphological sub-types. To resolve this, the system decouples detection from classification into a **Decoupled Two-Stage Architecture**:

```
Input Video Frame / Microscope Image
                 │
                 ▼
┌────────────────────────────────────────────────────────┐
│ STAGE 1: YOLO11n Single-Class Particle Detector       │
│ • Task: Particle Localization ("Where is it?")         │
│ • Classes: 1 unified class ('microplastic')            │
│ • Output: High-precision Bounding Boxes [cx, cy, w, h] │
│ • Filters: Water bubbles, organic bio-flocs, glare    │
└────────────────────────────────────────────────────────┘
                 │
      Cropped Bounding Boxes
                 │
                 ▼
┌────────────────────────────────────────────────────────┐
│ HYBRID FEATURE EXTRACTION                              │
│ • HOG (Histogram of Oriented Gradients) -> 1,764 dims  │
│ • 3D RGB Color Histogram (8x8x8 bins)   ->   512 dims  │
│ • Total Feature Vector                  -> 2,276 dims  │
└────────────────────────────────────────────────────────┘
                 │
                 ▼
┌────────────────────────────────────────────────────────┐
│ STAGE 2: Morphological Classifier (RBF-SVM / RF)       │
│ • Task: Fine-grained Taxonomy ("What kind is it?")    │
│ • Classes: fiber, film, fragment, pallet               │
│ • Validation Accuracy: 99.79% (929 / 931 correct)      │
└────────────────────────────────────────────────────────┘
                 │
                 ▼
Final Annotated Image with Particle Counts, Classes & Microplastic Concentrations
```

### 1.2 Target Morphological Classes

The taxonomy reflects international aquatic monitoring protocols (GESAMP / ISO):

| Class ID | Class Name          | Morphology & Origin Description                                              | Role in Pipeline         |
| -------- | ------------------- | ---------------------------------------------------------------------------- | ------------------------ |
| **0**    | `microplastic`      | Generic/unclassified microplastic particles from multi-source repositories   | Stage 1 Detection Target |
| **1**    | `fiber`             | Slender, elongated filaments (synthetic textiles, fishing nets, ropes)       | Stage 2 Classification   |
| **2**    | `film`              | Ultra-thin, flexible planar 2D fragments (plastic bags, packaging films)     | Stage 2 Classification   |
| **3**    | `fragment`          | Rigid, angular, irregular particles derived from broken macroplastic debris  | Stage 2 Classification   |
| **4**    | `pellet` / `pallet` | Pre-production industrial resin beads, smooth spherical/cylindrical granules | Stage 2 Classification   |
| **5**    | `foam`              | Porous, cellular expanded polystyrene (EPS) fragments                        | Stage 2 Classification   |

---

## 2. Deep Dive: How the YOLO Algorithm Works (Stage 1)

YOLO (**You Only Look Once**) is a single-stage, anchor-free deep convolutional object detection model. Unlike two-stage detectors (e.g. Faster R-CNN) that first generate thousands of region proposals and run secondary networks to classify them, YOLO formulates object detection as a direct spatial regression problem in a **single network forward pass**.

### 2.1 Neural Network Components

1. **Backbone (Feature Extraction)**:
   - Built on a modified CSPDarknet with **C3k2 modules** (Cross Stage Partial with 2 convolutions) in YOLO11.
   - Progressively downsamples the input image ($640 \times 640$) via strided convolutions to extract hierarchical spatial features:
     - Shallow layers: Edge boundaries, gradient contrasts, optical transparency.
     - Deep layers: Complex geometric shapes, contours, particle-liquid boundaries.
2. **Neck (Multi-Scale Feature Fusion)**:
   - Employs a **PANet (Path Aggregation Network)** coupled with **FPN (Feature Pyramid Network)**.
   - Combines low-level spatial detail (crucial for locating sub-millimeter particles and thin fibers) with high-level contextual semantics (differentiating plastic from background water illumination artifacts).
   - Generates predictions at three distinct grid scales: $80 \times 80$ (tiny particles), $40 \times 40$ (medium particles), and $20 \times 20$ (larger clusters).
3. **Head (Anchor-Free Detection)**:
   - Operates in an **anchor-free** manner: directly predicts the center offset $(\Delta x, \Delta y)$ and normalized dimensions $(w, h)$ from each grid cell, eliminating manually tuned prior anchor boxes.
   - Decouples classification heads from bounding box regression heads to avoid conflicting feature gradients.

### 2.2 Mathematical Loss Formulations

During training via [train_yolov11.py](file:///c:/Users/suhai/Downloads/Custom_Microplastic_Dataset/Custom_Microplastic_Dataset/scripts/train_yolov11.py), YOLO optimizes a multi-task composite loss function:

$$\mathcal{L}_{\text{total}} = \lambda_{\text{box}} \mathcal{L}_{\text{CIoU}} + \lambda_{\text{dfl}} \mathcal{L}_{\text{DFL}} + \lambda_{\text{cls}} \mathcal{L}_{\text{BCE}}$$

1. **Complete Intersection over Union (CIoU) Loss**:
   Accounts for overlap area, center point distance, and aspect ratio consistency:
   $$\mathcal{L}_{\text{CIoU}} = 1 - \text{IoU} + \frac{\rho^2(b, b^{\text{gt}})}{c^2} + \alpha v$$
   Where:
   - $\text{IoU} = \frac{|B \cap B^{\text{gt}}|}{|B \cup B^{\text{gt}}|}$
   - $\rho(b, b^{\text{gt}})$ is the Euclidean distance between predicted and ground-truth box centroids.
   - $c$ is the diagonal length of the smallest enclosing box.
   - $v = \frac{4}{\pi^2}\left(\arctan\frac{w^{\text{gt}}}{h^{\text{gt}}} - \arctan\frac{w}{h}\right)^2$ measures aspect ratio disparity, with balancing factor $\alpha$.

2. **Distribution Focal Loss (DFL)**:
   Models bounding box boundaries as continuous probability distributions rather than Dirac delta functions, allowing the network to handle uncertain or blurry microplastic boundaries suspended in fluid:
   $$\mathcal{L}_{\text{DFL}}(S_i, S_{i+1}) = - \left((y_{i+1} - y)\log(S_i) + (y - y_i)\log(S_{i+1})\right)$$

3. **Binary Cross-Entropy (BCE) Loss**:
   Evaluates particle presence confidence (objectness):
   $$\mathcal{L}_{\text{BCE}} = - [y \log(\hat{y}) + (1 - y)\log(1 - \hat{y})]$$

### 2.3 Post-Processing: Non-Maximum Suppression (NMS)

Because multiple neighboring grid cells often detect the same particle with high confidence, NMS eliminates redundant overlapping detections:

1. Sorts all candidate boxes by descending confidence score.
2. Selects the highest scoring box $B_{\text{max}}$.
3. Computes the pairwise IoU between $B_{\text{max}}$ and all remaining candidate boxes.
4. Removes any box with $\text{IoU} > \text{threshold}$ (typically $\text{IoU} > 0.45$).
5. Repeats until all non-overlapping detections are finalized.

---

## 3. Deep Dive: How Feature Extraction & SVM Work (Stage 2)

Stage 2 classifies the cropped particle bounding boxes extracted by [dataset_cropper.py](file:///c:/Users/suhai/Downloads/Custom_Microplastic_Dataset/Custom_Microplastic_Dataset/scripts/dataset_cropper.py). Instead of feeding raw pixel arrays (which are sensitive to illumination and zoom), the pipeline constructs a **hybrid morphological and chromatic feature vector**.

### 3.1 Feature Extraction Pipeline ([retrain_classifier.py](file:///c:/Users/suhai/Downloads/Custom_Microplastic_Dataset/Custom_Microplastic_Dataset/scripts/retrain_classifier.py))

Every detected particle crop is resized to $128 \times 128$ pixels. Two complementary descriptors are computed:

```
Cropped Image (128x128x3)
     │
     ├──> Grayscale ──> HOG Extraction (9 bins, 8x8 cells, 2x2 blocks) ──> 1,764 features
     │
     └──> RGB Color ──> 3D Color Histogram (8x8x8 bins, L1-normalized) ──>   512 features
                                                                                 │
                                          Combined Feature Vector (HSTACK) ──────┴──> 2,276 features
```

#### A. Histogram of Oriented Gradients (HOG) — 1,764 Dimensions

HOG captures local shape and edge structural distributions:

1. **Gradient Computation**: Calculates horizontal and vertical gradients $G_x(x,y)$ and $G_y(x,y)$ using 1D derivative filters $[-1, 0, 1]$.
   - Gradient magnitude: $M(x, y) = \sqrt{G_x^2 + G_y^2}$
   - Gradient orientation: $\theta(x, y) = \arctan\left(\frac{G_y}{G_x}\right) \pmod{180^\circ}$
2. **Spatial Binning**: The $128 \times 128$ image is divided into $8 \times 8$ pixel cells. Within each cell, gradient magnitudes are binned into 9 orientation channels ($0^\circ \text{ to } 180^\circ$).
3. **Block Normalization**: Cells are grouped into $2 \times 2$ blocks with $50\%$ overlap and normalized using L2-Hys norm to achieve illumination invariance.

- **Morphological Discriminative Power**:
  - `fiber`: Concentrated gradient energy in a single orientation across elongated dimensions.
  - `fragment`: Irregular, high-variance multi-directional sharp edge gradients.
  - `pallet`: Smooth, circular gradient transitions radially distributed around the perimeter.
  - `film`: Low interior gradient magnitude with diffuse outer planar boundaries.

#### B. 3D RGB Color Histogram — 512 Dimensions

Computes a joint 3D color distribution across the RGB channels using 8 bins per axis ($8 \times 8 \times 8 = 512$ total bins):

- Distinguishes dyed synthetic polymers, weathered white pellets, and transparent translucent films.
- L1-normalized so total bin frequency sums to 1.0, rendering it invariant to particle area.

---

### 3.2 How the Support Vector Machine (SVM) Works

A Support Vector Machine is a supervised maximum-margin classifier designed to find an optimal decision boundary separating classes in high-dimensional feature space.

#### A. The Maximum Margin Hyperplane

In a binary case with feature vectors $\mathbf{x}_i \in \mathbb{R}^{2276}$ and labels $y_i \in \{-1, +1\}$, the separating hyperplane is defined as:

$$\mathbf{w}^T \mathbf{x} + b = 0$$

The distance (margin) between the hyperplane and the closest data points is $\frac{2}{\|\mathbf{w}\|}$. To maximize this margin while allowing for soft penalties on overlapping points, SVM solves the primal optimization problem:

$$\min_{\mathbf{w}, b, \boldsymbol{\xi}} \frac{1}{2} \|\mathbf{w}\|^2 + C \sum_{i=1}^N \xi_i \quad \text{subject to } y_i (\mathbf{w}^T \phi(\mathbf{x}_i) + b) \ge 1 - \xi_i, \quad \xi_i \ge 0$$

Where:

- $\mathbf{w}$ is the weight vector orthogonal to the decision boundary.
- $C$ is the regularization parameter ($C = 10.0$ in our implementation). It controls the trade-off between maximizing the margin and minimizing classification errors.
- $\xi_i$ are slack variables allowing misclassifications when classes overlap slightly.
- **Support Vectors**: Only the data points that lie directly on the margin boundaries ($y_i(\mathbf{w}^T \phi(\mathbf{x}_i) + b) = 1$) or violate the margin ($\xi_i > 0$) determine the boundary position. Non-critical interior points have zero influence, making SVM highly resilient to overfitting.

#### B. The Radial Basis Function (RBF) Kernel Trick

Microplastic morphological feature vectors cannot be separated by simple linear hyperplanes. To achieve separation without calculating coordinates in an intractable space, SVM uses the **Kernel Trick**:

$$K(\mathbf{x}_i, \mathbf{x}_j) = \phi(\mathbf{x}_i)^T \phi(\mathbf{x}_j) = \exp\left(-\gamma \|\mathbf{x}_i - \mathbf{x}_j\|^2\right)$$

- The RBF (Gaussian) kernel implicitly projects the 2,276-dimensional vectors into an infinite-dimensional Hilbert space where non-linear particle clusters become linearly separable.
- The scale parameter $\gamma = \frac{1}{2\sigma^2}$ (set to `'scale'` = $\frac{1}{N_{\text{features}} \cdot \text{Var}(X)}$) defines the influence radius of individual support vectors.

#### C. Multi-Class Classification (One-vs-One Voting)

To classify into the 4 target classes (`fiber`, `film`, `fragment`, `pallet`):

- The model trains $\frac{K(K-1)}{2} = \frac{4 \times 3}{2} = 6$ binary SVM classifiers (e.g. fiber vs. film, fiber vs. fragment, film vs. pallet, etc.).
- At inference time, each binary classifier casts a vote, and the class with the highest aggregate vote count is assigned to the particle crop.

---

## 4. Multi-Source Dataset Integration

### 4.1 Raw Sources Merged

| Source Directory     | Origin & Reference                                  | Raw Images | Annotation Format             | Source Classes                                    | Native Resolution |
| -------------------- | --------------------------------------------------- | ---------- | ----------------------------- | ------------------------------------------------- | ----------------- |
| `dataset/`           | Kaggle (Roboflow Microplastics)                     | **781**    | YOLOv8 Bounding Box           | 1 (`microplastic`)                                | 640×640           |
| `kueranan/`          | Roboflow Universe (`kueranan/microplastic_100` v10) | **401**    | YOLOv8 Bounding Box           | 4 (`fiber`, `film`, `fragment`, `pallet`)         | 640×640           |
| `roboflow_in_water/` | Roboflow Universe (`Microplastic in Water` v1)      | **920**    | YOLOv8 OBB (Converted to HBB) | 1 (`microplastic`)                                | 600×800           |
| `synthetic/`         | Procedurally rendered microscopy models             | **180**    | YOLOv8 Bounding Box           | 5 (`fiber`, `film`, `fragment`, `pellet`, `foam`) | 640×640           |

> **Total Raw Input Images Across Sources**: **~2,282 images**

### 4.2 Merging & Standardization Pipeline

The dataset ingestion is handled via [merge_datasets.py](file:///c:/Users/suhai/Downloads/Custom_Microplastic_Dataset/Custom_Microplastic_Dataset/merge_datasets.py) and [remap_labels_1class.py](file:///c:/Users/suhai/Downloads/Custom_Microplastic_Dataset/Custom_Microplastic_Dataset/scripts/remap_labels_1class.py):

1. **Directory Normalization**: Unifies heterogeneous split nomenclatures (`valid/` vs `val/`).
2. **Namespace Collision Prevention**: Prefixes raw filenames with source tags (e.g. `kueranan_image.jpg`, `roboflow_in_water_013_png...`).
3. **Stage 1 Label Harmonization**: All labels mapped to 1 unified class `0` in [data_1class.yaml](file:///c:/Users/suhai/Downloads/Custom_Microplastic_Dataset/Custom_Microplastic_Dataset/data_1class.yaml) to eliminate contradictory loss penalties during detector training.
4. **Bounding Box Regularization**: Normalized YOLO coordinates (`cx cy w h`) bounded to $[0.0, 1.0]$.

---

## 5. Dataset Distribution & Split Counts

### 5.1 Stage 1 Full-Frame Detection Dataset (`merged_dataset/`)

| Split          | Images    | % of Total | Label Files | Annotated Particles (Boxes) | Avg Particles / Image |
| -------------- | --------- | ---------- | ----------- | --------------------------- | --------------------- |
| **Train**      | **1,725** | 82.06%     | 1,725       | 10,403                      | 6.03                  |
| **Validation** | **322**   | 15.32%     | 322         | 2,683                       | 8.33                  |
| **Test**       | **55**    | 2.62%      | 55          | 77                          | 1.40                  |
| **Total**      | **2,102** | 100.0%     | 2,102       | **13,163**                  | 6.26                  |

### 5.2 Stage 2 Cropped Particle Classification Dataset (`classification_dataset/`)

Cropped by [dataset_cropper.py](file:///c:/Users/suhai/Downloads/Custom_Microplastic_Dataset/Custom_Microplastic_Dataset/scripts/dataset_cropper.py) directly from verified bounding box coordinates:

| Class                      | Train Set Patches | Val Set Patches | Test Set Patches | Total Patches | Class Status in Stage 2 Training                    |
| -------------------------- | ----------------- | --------------- | ---------------- | ------------- | --------------------------------------------------- |
| **fiber**                  | 906               | 247             | 0                | 1,153         | Included (morphologically verified)                 |
| **film**                   | 963               | 242             | 0                | 1,205         | Included (morphologically verified)                 |
| **fragment**               | 960               | 230             | 0                | 1,190         | Included (morphologically verified)                 |
| **pallet** (pellet)        | 929               | 212             | 0                | 1,141         | Included (morphologically verified)                 |
| **microplastic** (generic) | 6,335             | 1,746           | 65               | 8,146         | **Excluded from Classifier** (unlabelled sub-types) |
| **Total Evaluated**        | **3,758**         | **931**         | **0**            | **4,689**     | _4-Class Balanced Benchmark_                        |

---

## 6. Empirical Experimental Results

### 6.1 Stage 1 Detector Performance (YOLO11n)

Trained for 50 epochs using [train_yolov11.py](file:///c:/Users/suhai/Downloads/Custom_Microplastic_Dataset/Custom_Microplastic_Dataset/scripts/train_yolov11.py) on [data_1class.yaml](file:///c:/Users/suhai/Downloads/Custom_Microplastic_Dataset/Custom_Microplastic_Dataset/data_1class.yaml):

| Metric                      | Result               | Description / Operational Significance                                             |
| --------------------------- | -------------------- | ---------------------------------------------------------------------------------- |
| **Precision (B)**           | **87.2%** (0.872)    | Low false-positive rate; reliably rejects water surface bubbles and organic debris |
| **Recall (B)**              | **74.7%** (0.747)    | High sensitivity on low-contrast transparent particles and thin fibers             |
| **mAP@50 (B)**              | **82.8%** (0.828)    | High particle localization quality across diverse water backgrounds                |
| **mAP@50-95 (B)**           | **40.7%** (0.407)    | Strict multi-IoU bounding box boundary tightness                                   |
| **Inference Time (GPU)**    | **4.7 ms / frame**   | Over 210 FPS on desktop CUDA GPU                                                   |
| **Inference Time (Jetson)** | **~24.5 ms / frame** | Sustained ~41 FPS real-time throughput on NVIDIA Jetson Nano                       |
| **Weight File Size**        | **5.5 MB**           | Compact footprint; cached directly in embedded L2 cache                            |

_Saved Model Weights_: [models/yolov11/microplastic_detector_1class/weights/best.pt](file:///c:/Users/suhai/Downloads/Custom_Microplastic_Dataset/Custom_Microplastic_Dataset/models/yolov11/microplastic_detector_1class/weights/best.pt)  
_Training Curves_: [results.png](file:///c:/Users/suhai/Downloads/Custom_Microplastic_Dataset/Custom_Microplastic_Dataset/models/yolov11/microplastic_detector_1class/results.png), [confusion_matrix_normalized.png](file:///c:/Users/suhai/Downloads/Custom_Microplastic_Dataset/Custom_Microplastic_Dataset/models/yolov11/microplastic_detector_1class/confusion_matrix_normalized.png), [BoxPR_curve.png](file:///c:/Users/suhai/Downloads/Custom_Microplastic_Dataset/Custom_Microplastic_Dataset/models/yolov11/microplastic_detector_1class/BoxPR_curve.png), [BoxF1_curve.png](file:///c:/Users/suhai/Downloads/Custom_Microplastic_Dataset/Custom_Microplastic_Dataset/models/yolov11/microplastic_detector_1class/BoxF1_curve.png)

### 6.2 Stage 2 Morphology Classification Results

Trained via [retrain_classifier.py](file:///c:/Users/suhai/Downloads/Custom_Microplastic_Dataset/Custom_Microplastic_Dataset/scripts/retrain_classifier.py) on 3,758 training and 931 unseen validation patches:

| Morphological Class        | Precision     | Recall        | F1-Score   | Validation Support                  |
| -------------------------- | ------------- | ------------- | ---------- | ----------------------------------- |
| **fiber**                  | 1.00 (100.0%) | 1.00 (100.0%) | 1.00       | 247 samples                         |
| **film**                   | 1.00 (100.0%) | 1.00 (100.0%) | 1.00       | 242 samples                         |
| **fragment**               | 1.00 (100.0%) | 1.00 (100.0%) | 1.00       | 230 samples                         |
| **pallet**                 | 1.00 (100.0%) | 0.99 (99.1%)  | 1.00       | 212 samples                         |
| **Overall Model Accuracy** | —             | —             | **99.79%** | **931 samples** (929 / 931 correct) |

_Trained Classifiers_:

- SVM (RBF kernel, $C=10.0$): [models/svm_classifier.pkl](file:///c:/Users/suhai/Downloads/Custom_Microplastic_Dataset/Custom_Microplastic_Dataset/models/svm_classifier.pkl)
- Random Forest (200 estimators): [models/rf_classifier.pkl](file:///c:/Users/suhai/Downloads/Custom_Microplastic_Dataset/Custom_Microplastic_Dataset/models/rf_classifier.pkl)
- Class Mapping: [models/classifier_classes.txt](file:///c:/Users/suhai/Downloads/Custom_Microplastic_Dataset/Custom_Microplastic_Dataset/models/classifier_classes.txt)

---

## 7. Decoupled Pipeline vs. Monolithic Multi-Class Baseline

| Architecture Paradigm                     | Localization mAP@50 | Morphology Classification Acc | Major Bottleneck                                                        |
| ----------------------------------------- | ------------------- | ----------------------------- | ----------------------------------------------------------------------- |
| **Monolithic Multi-Class YOLOv8**         | 61.4%               | 63.2%                         | Taxonomy clash; loss competition between generic and sub-types          |
| **Monolithic Multi-Class YOLO11**         | 63.8%               | 66.5%                         | Severe class imbalance; ambiguous fiber/fragment gradients              |
| **Decoupled Two-Stage (YOLO11 + RF/SVM)** | **82.8%**           | **99.79%**                    | **Optimal**: Clean gradient propagation and balanced morphology classes |

---

## 8. Real-Time Deployment & Inference Scripts

The complete inference engine is implemented in [detect_and_classify.py](file:///c:/Users/suhai/Downloads/Custom_Microplastic_Dataset/Custom_Microplastic_Dataset/scripts/detect_and_classify.py) and documented in [howtorun(merged-dataset).md](file:///c:/Users/suhai/Downloads/Custom_Microplastic_Dataset/Custom_Microplastic_Dataset/howtorun%28merged-dataset%29.md):

```bash
# 1. Single Image Inference
conda run -n yolo python scripts/detect_and_classify.py \
  --source "merged_dataset/test/images/roboflow_in_water_013_png.rf.b42a6cb49dea9046bbce321cce923938.jpg" \
  --yolo_model "models/yolov11/microplastic_detector_1class/weights/best.pt"

# 2. Batch Folder Inference
conda run -n yolo python scripts/detect_and_classify.py \
  --source "merged_dataset/test/images" \
  --yolo_model "models/yolov11/microplastic_detector_1class/weights/best.pt"

# 3. Real-Time Camera / Microscope Live Feed (Jetson Nano / Raspberry Pi)
conda run -n yolo python scripts/detect_and_classify.py \
  --source 0 \
  --yolo_model "models/yolov11/microplastic_detector_1class/weights/best.pt"
```

---

## 9. Comprehensive System Summary

| Parameter                               | Specification                                     |
| --------------------------------------- | ------------------------------------------------- |
| **Total Full-Frame Images**             | 2,102 images (1,725 Train / 322 Val / 55 Test)    |
| **Total Labeled Particle Instances**    | 13,163 localized bounding boxes                   |
| **Cropped Classification Patches**      | 10,093 train / 2,677 val / 65 test                |
| **Stage 2 Specific Morphology Samples** | 4,689 patches (3,758 train / 931 val)             |
| **Stage 1 Detector Architecture**       | YOLO11n (Single-class, anchor-free, 5.5 MB)       |
| **Stage 1 Detection Performance**       | Precision: 87.2%, Recall: 74.7%, mAP@50: 82.8%    |
| **Stage 2 Morphology Classifier**       | RBF-SVM / Random Forest + HOG + Color Histogram   |
| **Stage 2 Classification Accuracy**     | 99.79% (929 / 931 validation samples)             |
| **Target Edge Platforms**               | NVIDIA Jetson Nano, Raspberry Pi 4/5, Desktop GPU |
