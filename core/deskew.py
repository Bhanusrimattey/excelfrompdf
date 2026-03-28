# Copyright (c) 2026 Bhanusri mattey
# Licensed under the Business Source License 1.1
# See LICENSE file in the project root for details.
# Commercial use prohibited until 2030-03-01

import cv2
import numpy as np
import math

# ---------- utils ----------
def rotate_expand(img, angle_deg, border=(255,255,255), interp=cv2.INTER_CUBIC):
    """Rotate by -angle_deg with expanded canvas to avoid cropping."""
    h, w = img.shape[:2]
    c = (w/2, h/2)
    M = cv2.getRotationMatrix2D(c, -angle_deg, 1.0)
    cos, sin = abs(M[0,0]), abs(M[0,1])
    new_w = int(h*sin + w*cos)
    new_h = int(h*cos + w*sin)
    M[0,2] += (new_w/2) - c[0]
    M[1,2] += (new_h/2) - c[1]
    return cv2.warpAffine(img, M, (new_w, new_h), flags=interp, borderValue=border)

def binarize_otsu(img_bgr):
    if len(img_bgr.shape) == 3:
        gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)
    else:
        gray = img_bgr
    gray = cv2.medianBlur(gray, 3)
    _, bw = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY+cv2.THRESH_OTSU)
    return bw  # text≈0, bg≈255

def projection_score_rows(bw):
    """Sharpness of horizontal projection (bigger = better)."""
    inv = (255 - bw) // 255  # text=1
    proj = inv.sum(axis=1).astype(np.float32)
    d = np.diff(proj)
    return float(np.dot(d, d))

def golden_max(f, a, b, tol=1e-3):
    """Golden-section search to maximize f on [a,b]."""
    gr = (math.sqrt(5) - 1) / 2
    c = b - gr*(b - a)
    d = a + gr*(b - a)
    fc, fd = f(c), f(d)
    while (b - a) > tol:
        if fc < fd:
            a = c
            c, fc = d, fd
            d = a + gr*(b - a)
            fd = f(d)
        else:
            b = d
            d, fd = c, fc
            c = b - gr*(b - a)
            fc = f(c)
    x = (a + b) / 2
    return x, f(x)

# ---------- key: suppress rulings so text drives the angle ----------
def suppress_rulings(bw):
    """
    Remove long horizontal/vertical lines from a binarized page.
    Input: bw (text black=0, bg white=255)
    Return: bw_text_only (text black, bg white)
    """
    inv = 255 - bw  # text white
    h, w = inv.shape

    # Long kernels; adjust factors if needed for very large pages
    k_h = max(25, w // 25)   # horizontal line length
    k_v = max(25, h // 25)   # vertical   line length
    kernel_h = cv2.getStructuringElement(cv2.MORPH_RECT, (k_h, 1))
    kernel_v = cv2.getStructuringElement(cv2.MORPH_RECT, (1, k_v))

    horiz = cv2.morphologyEx(inv, cv2.MORPH_OPEN, kernel_h, iterations=1)
    vert  = cv2.morphologyEx(inv, cv2.MORPH_OPEN, kernel_v, iterations=1)

    rulings = cv2.bitwise_or(horiz, vert)
    inv_wo = cv2.subtract(inv, rulings)   # remove lines, keep text
    bw_text = 255 - inv_wo                # back to text black
    return bw_text

# ---------- angle estimators ----------
def estimate_angle_projection(img_bgr, coarse_range=6.0, coarse_step=0.25, refine_half=0.5):
    """
    Use text-only mask + projection score.
    """
    bw = binarize_otsu(img_bgr)
    bw_text = suppress_rulings(bw)

    # optional downscale for speed (keep geometry consistent)
    H, W = bw_text.shape
    scale = 1600.0 / max(H, W)
    work = cv2.resize(bw_text, None, fx=scale, fy=scale, interpolation=cv2.INTER_NEAREST) if scale < 1.0 else bw_text

    # coarse scan
    angles = np.arange(-coarse_range, coarse_range + 1e-9, coarse_step)
    scores = [projection_score_rows(rotate_expand(work, a, interp=cv2.INTER_NEAREST)) for a in angles]
    a0 = float(angles[int(np.argmax(scores))])

    # refine with golden-section around the peak
    def f(a):
        return projection_score_rows(rotate_expand(work, float(a), interp=cv2.INTER_NEAREST))
    lo, hi = a0 - refine_half, a0 + refine_half
    best_a, best_s = golden_max(f, lo, hi, tol=1e-3)

    # confidence via prominence in refined window
    grid = np.linspace(lo, hi, 41)
    vals = np.array([f(a) for a in grid], dtype=np.float64)
    base = float(np.median(vals))
    peak = float(vals.max())
    conf = 0.0 if peak <= 0 else max(0.0, min(1.0, (peak - base) / (peak + 1e-6)))
    return float(best_a), conf

def estimate_angle_moments(img_bgr):
    """
    Fallback: orientation from contour moments of text components.
    """
    bw = binarize_otsu(img_bgr)
    bw_text = suppress_rulings(bw)
    inv = 255 - bw_text
    # light open to split blobs
    inv = cv2.morphologyEx(inv, cv2.MORPH_OPEN,
                           cv2.getStructuringElement(cv2.MORPH_RECT, (3,3)))
    cnts, _ = cv2.findContours(inv, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)

    angles = []
    weights = []
    for c in cnts:
        area = cv2.contourArea(c)
        if area < 20:   # ignore tiny specks
            continue
        mu = cv2.moments(c)
        denom = (mu['mu20'] + mu['mu02'])
        if denom <= 1e-6:
            continue
        # orientation of ellipse major axis
        theta = 0.5 * math.atan2(2*mu['mu11'], (mu['mu20'] - mu['mu02']))
        a_deg = math.degrees(theta)
        # normalize to [-90, 90]
        if a_deg > 90: a_deg -= 180
        if a_deg < -90: a_deg += 180
        if abs(a_deg) <= 30:  # ignore garbage
            angles.append(a_deg)
            weights.append(area)
    if not angles:
        return 0.0, 0.0
    a = float(np.average(np.array(angles), weights=np.array(weights)))
    # dispersion → confidence
    mad = float(np.median(np.abs(np.array(angles) - np.median(angles))))
    conf = max(0.0, min(1.0, 1.0 - mad/15.0))
    return a, conf

def estimate_best_angle(img_bgr):
    a1, c1 = estimate_angle_projection(img_bgr)
    a2, c2 = estimate_angle_moments(img_bgr)

    # Evaluate both on full-res text-only projection score
    bw = binarize_otsu(img_bgr)
    work = suppress_rulings(bw)
    s1 = projection_score_rows(rotate_expand(work, a1, interp=cv2.INTER_NEAREST))
    s2 = projection_score_rows(rotate_expand(work, a2, interp=cv2.INTER_NEAREST))

    if s1 >= s2:
        return a1, max(c1, 0.6)
    else:
        return a2, max(c2, 0.6)

# ---------- API ----------
def deskew(image_bgr, min_apply_deg=0.15):
    """
    Returns (rotated, angle_deg, confidence).
    Skips rotation if |angle| < min_apply_deg (saves quality/time).
    """
    angle_deg, conf = estimate_best_angle(image_bgr)
    if abs(angle_deg) < min_apply_deg:
        return image_bgr, angle_deg, conf
    return rotate_expand(image_bgr, angle_deg), angle_deg, conf

def full_process(image_bgr):
    angle_deg, conf = estimate_best_angle(image_bgr)
    rotated = rotate_expand(image_bgr, angle_deg)
    # optional cleanup:
    # rotated = clean_borders(rotated)  # if you use border whitening
    # crop tight content
    gray = cv2.cvtColor(rotated, cv2.COLOR_BGR2GRAY)
    _, th = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV+cv2.THRESH_OTSU)
    ys, xs = np.where(th > 0)
    if len(xs) and len(ys):
        x0, x1 = max(0, xs.min()-8), min(rotated.shape[1], xs.max()+8)
        y0, y1 = max(0, ys.min()-8), min(rotated.shape[0], ys.max()+8)
        cropped = rotated[y0:y1, x0:x1]
    else:
        cropped = rotated
    # margin
    final = cv2.copyMakeBorder(cropped, 20, 20, 20, 20,
                               cv2.BORDER_CONSTANT, value=(255,255,255))
    return final, angle_deg, conf
