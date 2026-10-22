# 1. Test on any test image:
conda run -n yolo python scripts/detect_and_classify.py `
  --source "merged_dataset/test/images/roboflow_in_water_013_png.rf.b42a6cb49dea9046bbce321cce923938.jpg" `
  --yolo_model "models/yolov11/microplastic_detector_1class/weights/best.pt"

# 2. Batch process an entire folder:
conda run -n yolo python scripts/detect_and_classify.py `
  --source "merged_dataset/test/images" `
  --yolo_model "models/yolov11/microplastic_detector_1class/weights/best.pt"

# 3. Real-time Camera / Edge deployment (Webcam / USB camera):
conda run -n yolo python scripts/detect_and_classify.py `
  --source 0 `
  --yolo_model "models/yolov11/microplastic_detector_1class/weights/best.pt"