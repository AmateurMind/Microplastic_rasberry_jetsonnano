# Synthetic Microplastic Dataset

This standalone dataset was generated procedurally with
`generator/generate_synthetic.py` for five morphology classes:

```text
0 fiber
1 fragment
2 film
3 pellet
4 foam
```

It contains 180 640x640 PNG images with YOLO-format labels, split into:

- `train/`: 125 images
- `val/`: 27 images
- `test/`: 28 images

Use `data.yaml` with Ultralytics YOLO. This dataset is separate from the
downloaded `kueranan/` dataset.
