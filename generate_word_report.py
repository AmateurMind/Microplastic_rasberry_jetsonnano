import sys
import subprocess
import os

# Install python-docx if not installed
try:
    import docx
except ImportError:
    print("python-docx not found. Installing...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "python-docx"])
    import docx

from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

def main():
    doc = docx.Document()

    # Title
    title = doc.add_heading('YOLOv8 vs YOLOv11 Comparison', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph("Comparison report for predicting microplastics on test image: 104_png.rf.ec4057e0614b9812ad467c5bf81847cf.jpg")

    # 1. Model Architecture & Size
    doc.add_heading('1. Model Architecture & Size', level=1)
    doc.add_paragraph("Based on the terminal output for the YOLOv11n and YOLOv8 model runs:")
    
    p = doc.add_paragraph()
    p.add_run("• YOLOv11n: ").bold = True
    p.add_run("101 layers, 2,583,322 parameters, 6.3 GFLOPs\n")
    p.add_run("• YOLOv8: ").bold = True
    p.add_run("73 layers, 3,006,818 parameters, 8.1 GFLOPs")
    
    p = doc.add_paragraph()
    p.add_run("Takeaway: ").bold = True
    p.add_run("YOLOv11 is more computationally efficient (lower GFLOPs) and has about 14% fewer parameters, despite being a deeper model with more layers.")

    # 2. Prediction Results & Validation
    doc.add_heading('2. Prediction Results & Validation', level=1)
    doc.add_paragraph("The ground truth for the test image contains exactly:")
    
    doc.add_paragraph("1 fiber (Class 0)", style='List Bullet')
    doc.add_paragraph("4 films (Class 1)", style='List Bullet')
    doc.add_paragraph("5 fragments (Class 2)", style='List Bullet')
    doc.add_paragraph("4 pallets (Class 3)", style='List Bullet')
    
    doc.add_paragraph("Predictions for both models:")
    p = doc.add_paragraph()
    p.add_run("• YOLOv11: ").bold = True
    p.add_run("1 fiber, 4 films, 5 fragments, 4 pallets\n")
    p.add_run("• YOLOv8: ").bold = True
    p.add_run("1 fiber, 4 films, 5 fragments, 4 pallets")
    
    p = doc.add_paragraph()
    p.add_run("Validation Result: ").bold = True
    p.add_run("Both models were 100% correct in their predictions. They successfully identified every single object in the image and classified them perfectly.")

    # 3. Inference Speed
    doc.add_heading('3. Inference Speed', level=1)
    p = doc.add_paragraph()
    p.add_run("• YOLOv11n: ").bold = True
    p.add_run("21.2ms inference time\n")
    p.add_run("• YOLOv8: ").bold = True
    p.add_run("11.1ms inference time")
    
    p = doc.add_paragraph()
    p.add_run("Takeaway: ").bold = True
    p.add_run("YOLOv8 was about 2x faster for inference on this single image on the RTX 2050 GPU.")

    # 4. Confidence Score Analysis
    doc.add_heading('4. Confidence Score Analysis', level=1)
    doc.add_paragraph("While both models detected the exact same objects, YOLOv11 was generally more confident in its predictions for larger objects.")
    
    doc.add_paragraph("Objects where YOLOv11 was MORE confident:", style='List Bullet')
    p_more = doc.add_paragraph()
    p_more.add_run("Fiber: 0.86 → 0.89\nTop right fragment: 0.63 → 0.67\nTop middle film: 0.70 → 0.74\nMiddle left film: 0.74 → 0.77\nCenter fragment: 0.60 → 0.68\nBottom left fragment: 0.64 → 0.69\nBottom center film: 0.62 → 0.68\nBottom right fragment: 0.62 → 0.70")
    
    doc.add_paragraph("Objects where YOLOv11 was LESS confident:", style='List Bullet')
    p_less = doc.add_paragraph()
    p_less.add_run("Center pallet: 0.56 → 0.46\nBottom right pallet: 0.59 → 0.55")
    
    p = doc.add_paragraph()
    p.add_run("Takeaway: ").bold = True
    p.add_run("YOLOv11 outputs higher confidence scores for most objects like fibers, films, and fragments, showing its architectural improvements make it more certain about its correct predictions.")

    # Images
    doc.add_heading('5. Prediction Visuals', level=1)
    
    # Paths from the terminal output earlier
    img_v8 = os.path.join("runs", "detect", "predict-8", "104_png.rf.ec4057e0614b9812ad467c5bf81847cf.jpg")
    img_v11 = os.path.join("runs", "detect", "predict-7", "104_png.rf.ec4057e0614b9812ad467c5bf81847cf.jpg")
    
    doc.add_paragraph("YOLOv8 Prediction Result (predict-8):")
    try:
        doc.add_picture(img_v8, width=Inches(5))
    except Exception as e:
        doc.add_paragraph(f"(Could not insert image. Ensure it exists at {img_v8})")

    doc.add_paragraph("YOLOv11 Prediction Result (predict-7):")
    try:
        doc.add_picture(img_v11, width=Inches(5))
    except Exception as e:
        doc.add_paragraph(f"(Could not insert image. Ensure it exists at {img_v11})")

    # Save
    out_path = os.path.join("docs", "YOLOv8_vs_YOLOv11_Comparison_Report.docx")
    
    # Ensure docs dir exists
    if not os.path.exists("docs"):
        os.makedirs("docs")
        
    doc.save(out_path)
    print(f"Success! Saved report to {out_path}")

if __name__ == "__main__":
    main()
