# Local YOLO Training

## Specification

- Train a one-class (`microplastic`) Ultralytics YOLO detector on the real Kaggle dataset only.
- Use `dataset/train` (577 image-label pairs) for training and `dataset/valid` (204 pairs) for validation.
- Start with the compact pretrained `yolo26n.pt` model at 512 px and batch size 2, suited to the available 4 GB RTX 2050.
- Save runs locally under `runs/rtx2050_actual`.

## Progress

- [x] Inspect the dataset and confirm complete YOLO image-label pairs.
- [x] Confirm the NVIDIA GPU is visible to the system.
- [x] Add a Windows-safe, one-class dataset configuration.
- [ ] Install PyTorch and Ultralytics with CUDA support.
- [ ] Launch training and verify that weights and logs are produced.
- [ ] Record final training status and output locations.
