# Copyright (c) 2026 Bhanusri mattey
# Licensed under the Business Source License 1.1
# See LICENSE file in the project root for details.
# Commercial use prohibited until 2030-03-01


from pypdf import PdfReader, PdfWriter, Transformation
from pypdf.generic import NameObject
import pytesseract
import numpy as np
import cv2
from PIL import ImageOps
import pypdfium2 as pdfium
import tempfile
import shutil
import sys
from pathlib import Path
import os
import io
from PIL import Image
from dotenv import load_dotenv
load_dotenv()

if sys.platform == "win32":
    pytesseract.pytesseract.tesseract_cmd = os.environ.get(
        "TESSERACT_PATH",
        r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    )



import os
import shutil
import tempfile

def pdf_to_images(pdf_path, dpi=300):
    pdf = pdfium.PdfDocument(pdf_path)
    images = []

    scale = dpi / 72  # VERY IMPORTANT

    for page in pdf:
        bitmap = page.render(scale=scale)
        pil_image = bitmap.to_pil()
        images.append(pil_image)

    pdf.close()
    return images

def preprocess_image(pil_image):
    img = np.array(pil_image.convert("L"))
    img = cv2.fastNlMeansDenoising(img, h=10)
    img = cv2.adaptiveThreshold(
        img, 255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        31, 10
    )
    return Image.fromarray(img)

def detect_rotation_angle_ocr(pil_image):
    clean = preprocess_image(pil_image)
    
    # Test all rotations instead of trusting OSD alone
    rotations = [0, 90, 180, 270]
    scores = {}
    
    for angle in rotations:
        if angle == 0:
            rotated = clean
        else:
            rotated = clean.rotate(angle, expand=True)
        scores[angle] = ocr_score(rotated)
    
   
    best_angle = max(scores, key=scores.get)
    
    # Only rotate if best is meaningfully better than original
    if best_angle != 0 and scores[best_angle] > scores[0] * 1.05:
        return best_angle
    
    return 0


def detect_flip_direction(pil_image):
    normal = pil_image
    flip_h = ImageOps.mirror(pil_image)   # left-right
    flip_v = ImageOps.flip(pil_image)     # top-bottom

    score_normal = ocr_score(normal)
    if score_normal > 80:
        return "normal"
   
    score_h = ocr_score(flip_h)
    score_v = ocr_score(flip_v)

    scores = {
        "normal": score_normal,
        "horizontal": score_h,
        "vertical": score_v
    }

    best = max(scores, key=scores.get)
    
    return best

def ocr_score(image):
    
    data = pytesseract.image_to_data(
        image,
        output_type=pytesseract.Output.DICT,
        config='--psm 6 --oem 1'
    )

    confidences = [
        c for c in data["conf"]
        if isinstance(c, int) and c > 0
    ]

    return sum(confidences) / len(confidences) if confidences else 0


def auto_correct_pdf_per_page(pdf_path):
    pdf_path = Path(pdf_path)

    reader = PdfReader(pdf_path)
    writer = PdfWriter()

    images = pdf_to_images(pdf_path)

    for i, page in enumerate(reader.pages):
        
        img = images[i]

        # 1️⃣ OCR-based rotation detection
        angle = detect_rotation_angle_ocr(img)

        if angle != 0:
            page.rotate(-angle)
            img = img.rotate(angle, expand=True)

        # 2️⃣ OCR-based flip detection
        flip_mode = detect_flip_direction(img)

        w = float(page.mediabox.width)
        h = float(page.mediabox.height)

        if flip_mode == "horizontal":
            t = Transformation().scale(-1, 1).translate(w, 0)
            page.add_transformation(t)

        elif flip_mode == "vertical":
            t = Transformation().scale(1, -1).translate(0, h)
            page.add_transformation(t)

        writer.add_page(page)

    # 🔒 write safely to temp file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        writer.write(tmp)
        temp_path = Path(tmp.name)

    return str(temp_path)