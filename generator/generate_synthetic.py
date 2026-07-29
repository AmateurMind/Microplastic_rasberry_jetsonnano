"""
Synthetic Microplastic Micrograph Generator
--------------------------------------------
Generates procedurally-created synthetic images that mimic microplastic
particles (fibre, fragment, film, pellet, foam) imaged on a filter-membrane
background under a digital microscope, together with YOLO-format bounding
box annotations. Intended to AUGMENT (not replace) real secondary datasets
in the custom dataset described in Task 3 of the project report.

Classes (based on commonly used morphological categories in the literature,
e.g. GESAMP / Meyers et al. 2022 Nile-Red morphology scheme):
    0 fiber
    1 fragment
    2 film
    3 pellet
    4 foam

Output: PNG image + .txt YOLO label per sample, written to dataset_raw/
"""
import os, random, math
import numpy as np
from PIL import Image, ImageDraw, ImageFilter

random.seed(42)
np.random.seed(42)

IMG_SIZE = 640
CLASSES = ["fiber", "fragment", "film", "pellet", "foam"]
OUT_DIR = "/home/claude/task3/dataset_raw"
os.makedirs(os.path.join(OUT_DIR, "images"), exist_ok=True)
os.makedirs(os.path.join(OUT_DIR, "labels"), exist_ok=True)

PARTICLE_COLORS = [
    (235, 235, 240), (210, 220, 230), (190, 195, 205),
    (225, 210, 190), (200, 180, 160), (170, 180, 190),
]


def make_filter_background(size):
    """Simulated membrane-filter background with mild speckle texture and vignette."""
    base = np.random.normal(loc=228, scale=6, size=(size, size)).clip(180, 255).astype(np.uint8)
    img = Image.fromarray(base).convert("L").convert("RGB")
    # radial vignette to mimic microscope illumination falloff
    yy, xx = np.mgrid[0:size, 0:size]
    cx, cy = size / 2, size / 2
    dist = np.sqrt((xx - cx) ** 2 + (yy - cy) ** 2) / (size / 2)
    vignette = (1 - 0.25 * np.clip(dist, 0, 1)).astype(np.float32)
    arr = np.array(img).astype(np.float32)
    arr *= vignette[..., None]
    img = Image.fromarray(arr.clip(0, 255).astype(np.uint8))
    img = img.filter(ImageFilter.GaussianBlur(radius=0.6))
    return img


def rand_color():
    c = random.choice(PARTICLE_COLORS)
    jitter = tuple(int(max(0, min(255, v + random.randint(-12, 12)))) for v in c)
    return jitter


def draw_fiber(draw, cx, cy, scale):
    length = random.uniform(40, 140) * scale
    width = random.uniform(2, 5) * scale
    angle = random.uniform(0, math.pi)
    curve = random.uniform(-0.4, 0.4)
    n_seg = 12
    pts = []
    for i in range(n_seg + 1):
        t = i / n_seg - 0.5
        x = cx + math.cos(angle) * length * t + math.sin(angle) * curve * length * (t ** 2)
        y = cy + math.sin(angle) * length * t - math.cos(angle) * curve * length * (t ** 2)
        pts.append((x, y))
    color = rand_color()
    draw.line(pts, fill=color, width=int(max(1, width)), joint="curve")
    xs = [p[0] for p in pts]; ys = [p[1] for p in pts]
    pad = width
    return (min(xs) - pad, min(ys) - pad, max(xs) + pad, max(ys) + pad)


def draw_fragment(draw, cx, cy, scale):
    r = random.uniform(10, 28) * scale
    n_pts = random.randint(6, 10)
    pts = []
    for i in range(n_pts):
        a = 2 * math.pi * i / n_pts + random.uniform(-0.15, 0.15)
        rr = r * random.uniform(0.6, 1.15)
        pts.append((cx + rr * math.cos(a), cy + rr * math.sin(a)))
    color = rand_color()
    draw.polygon(pts, fill=color, outline=tuple(max(0, v - 30) for v in color))
    xs = [p[0] for p in pts]; ys = [p[1] for p in pts]
    return (min(xs), min(ys), max(xs), max(ys))


def draw_film(draw, cx, cy, scale):
    w = random.uniform(30, 70) * scale
    h = random.uniform(15, 30) * scale
    angle = random.uniform(0, math.pi)
    n_pts = 10
    pts = []
    for i in range(n_pts):
        t = i / (n_pts - 1)
        lx = (t - 0.5) * w
        ly = math.sin(t * math.pi * random.uniform(1.5, 2.5)) * h * 0.25
        rx = lx * math.cos(angle) - ly * math.sin(angle)
        ry = lx * math.sin(angle) + ly * math.cos(angle)
        pts.append((cx + rx, cy + ry))
    for i in range(n_pts - 1, -1, -1):
        t = i / (n_pts - 1)
        lx = (t - 0.5) * w
        ly = h * 0.5 + math.sin(t * math.pi * random.uniform(1.5, 2.5)) * h * 0.15
        rx = lx * math.cos(angle) - ly * math.sin(angle)
        ry = lx * math.sin(angle) + ly * math.cos(angle)
        pts.append((cx + rx, cy + ry))
    color = rand_color()
    translucent = tuple(min(255, v + 15) for v in color)
    draw.polygon(pts, fill=translucent, outline=tuple(max(0, v - 25) for v in color))
    xs = [p[0] for p in pts]; ys = [p[1] for p in pts]
    return (min(xs), min(ys), max(xs), max(ys))


def draw_pellet(draw, cx, cy, scale):
    r = random.uniform(9, 20) * scale
    color = rand_color()
    bbox = (cx - r, cy - r, cx + r, cy + r)
    draw.ellipse(bbox, fill=color, outline=tuple(max(0, v - 35) for v in color))
    hl = (cx - r * 0.35, cy - r * 0.35, cx - r * 0.05, cy - r * 0.05)
    draw.ellipse(hl, fill=tuple(min(255, v + 25) for v in color))
    return bbox


def draw_foam(draw, cx, cy, scale):
    r = random.uniform(16, 32) * scale
    color = rand_color()
    bbox = (cx - r, cy - r, cx + r, cy + r)
    draw.ellipse(bbox, fill=color, outline=tuple(max(0, v - 30) for v in color))
    n_holes = random.randint(5, 10)
    for _ in range(n_holes):
        a = random.uniform(0, 2 * math.pi)
        rr = random.uniform(0, r * 0.7)
        hx, hy = cx + rr * math.cos(a), cy + rr * math.sin(a)
        hr = random.uniform(1.5, 4) * scale
        bg = rand_color()
        draw.ellipse((hx - hr, hy - hr, hx + hr, hy + hr), fill=tuple(min(255, v + 10) for v in bg))
    return bbox


DRAW_FN = {0: draw_fiber, 1: draw_fragment, 2: draw_film, 3: draw_pellet, 4: draw_foam}


def clamp_box(box, size):
    x1, y1, x2, y2 = box
    x1 = max(0, min(size - 1, x1)); y1 = max(0, min(size - 1, y1))
    x2 = max(0, min(size - 1, x2)); y2 = max(0, min(size - 1, y2))
    if x2 <= x1: x2 = x1 + 1
    if y2 <= y1: y2 = y1 + 1
    return x1, y1, x2, y2


def to_yolo(box, size):
    x1, y1, x2, y2 = box
    cx = (x1 + x2) / 2 / size
    cy = (y1 + y2) / 2 / size
    w = (x2 - x1) / size
    h = (y2 - y1) / size
    return cx, cy, w, h


def generate_image(idx, n_particles_range=(6, 22)):
    img = make_filter_background(IMG_SIZE)
    draw = ImageDraw.Draw(img, "RGB")
    labels = []
    n_particles = random.randint(*n_particles_range)
    # class distribution weighted like typical environmental samples: fibers/fragments dominant
    weights = [0.38, 0.30, 0.10, 0.12, 0.10]
    for _ in range(n_particles):
        cls = random.choices(range(5), weights=weights, k=1)[0]
        margin = 40
        cx = random.uniform(margin, IMG_SIZE - margin)
        cy = random.uniform(margin, IMG_SIZE - margin)
        scale = random.uniform(0.7, 1.3)
        box = DRAW_FN[cls](draw, cx, cy, scale)
        box = clamp_box(box, IMG_SIZE)
        yolo = to_yolo(box, IMG_SIZE)
        labels.append((cls, *yolo))
    # mild global noise + blur to emulate camera sensor noise
    arr = np.array(img).astype(np.int16)
    noise = np.random.normal(0, 4, arr.shape).astype(np.int16)
    arr = np.clip(arr + noise, 0, 255).astype(np.uint8)
    img = Image.fromarray(arr)
    if random.random() < 0.3:
        img = img.filter(ImageFilter.GaussianBlur(radius=random.uniform(0.3, 0.8)))
    fname = f"synth_{idx:04d}"
    img.save(os.path.join(OUT_DIR, "images", fname + ".png"))
    with open(os.path.join(OUT_DIR, "labels", fname + ".txt"), "w") as f:
        for cls, cx, cy, w, h in labels:
            f.write(f"{cls} {cx:.6f} {cy:.6f} {w:.6f} {h:.6f}\n")
    return fname


if __name__ == "__main__":
    N = 180
    names = [generate_image(i) for i in range(N)]
    with open(os.path.join(OUT_DIR, "classes.txt"), "w") as f:
        f.write("\n".join(CLASSES))
    print(f"Generated {len(names)} synthetic images with YOLO labels in {OUT_DIR}")
