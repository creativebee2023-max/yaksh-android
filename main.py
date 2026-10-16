import os
import re
import io
import glob
import math
import threading
from urllib.parse import urljoin, urlparse
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from bs4 import BeautifulSoup
import requests
from PIL import Image, ImageOps, ImageDraw, ImageTk
from rembg import remove, new_session
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# =============================================================================
# THEME & DESIGN CONSTANTS
# =============================================================================
COLOR_BORDER = RGBColor(30, 41, 59)
COLOR_HEADER = RGBColor(15, 23, 42)
COLOR_LABEL_KEY = RGBColor(71, 85, 105)
COLOR_LABEL_VAL = RGBColor(15, 23, 42)
COLOR_CARD_BG = RGBColor(248, 250, 252)
FONT_NAME = "Calibri"

QUALITY_PRESETS = {
    "Maximum Quality (Lossless/HQ)": {"max_dim": 2048, "jpg_quality": 92},
    "High Quality (Balanced - Default)": {"max_dim": 1600, "jpg_quality": 85},
    "Web/Email Compact": {"max_dim": 1200, "jpg_quality": 75},
    "Ultra Small (Draft)": {"max_dim": 900, "jpg_quality": 65}
}


# =============================================================================
# 1. ENHANCED WEB DOWNLOADER WITH SMART FILTERING
# =============================================================================
def sanitize_filename(filename):
    """Sanitize the image filename by removing invalid path characters[cite: 4]."""
    return re.sub(r'[\\/*?:"<>|]', "", filename)


def download_images_from_url(target_url, output_folder, log_func, avoid_icons=True, filter_mode="All Images"):
    """Downloads images from a website URL with smart filtering for icons, models, and garments[cite: 4]."""
    log_func(f"Fetching website: {target_url}[cite: 4]")

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/122.0.0.0 Safari/537.36"
        ),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
    }

    try:
        response = requests.get(target_url, headers=headers, timeout=15)
        response.raise_for_status()
    except Exception as e:
        log_func(f"Error fetching page: {e}[cite: 4]")
        return 0

    soup = BeautifulSoup(response.text, "html.parser")
    img_tags = soup.find_all("img")
    log_func(f"Found {len(img_tags)} total image tags[cite: 4].")

    if not img_tags:
        log_func("No images found on page[cite: 4].")
        return 0

    os.makedirs(output_folder, exist_ok=True)
    downloaded_count = 0
    skipped_count = 0

    icon_keywords = ["icon", "logo", "badge", "sprite", "arrow", "social", "footer", "banner", "pixel", "spinner", "placeholder"]

    for idx, img in enumerate(img_tags, start=1):
        img_url = (
            img.get("src")
            or img.get("data-src")
            or img.get("data-original")
            or img.get("data-lazy")
        )

        if not img_url:
            continue

        img_url = urljoin(target_url, img_url)

        if not img_url.startswith(("http://", "https://")):
            continue

        alt_text = img.get("alt", "").lower()
        url_lower = img_url.lower()

        if avoid_icons and any(kw in url_lower or kw in alt_text for kw in icon_keywords):
            skipped_count += 1
            continue

        try:
            img_data = requests.get(img_url, headers=headers, timeout=10)
            if img_data.status_code != 200:
                continue

            pil_temp = Image.open(io.BytesIO(img_data.content))
            width, height = pil_temp.size

            if avoid_icons and (width < 180 or height < 180):
                skipped_count += 1
                continue

            if filter_mode == "Garment Only (Flat-lay / Swatch)":
                model_keywords = ["model", "lifestyle", "on-model", "wear", "person", "girl", "boy"]
                if any(kw in url_lower or kw in alt_text for kw in model_keywords):
                    skipped_count += 1
                    continue
            elif filter_mode == "Garment with Model":
                model_keywords = ["model", "lifestyle", "on-model", "wear", "person", "girl", "boy"]
                if not any(kw in url_lower or kw in alt_text for kw in model_keywords):
                    if abs(width - height) < 20 and width < 400:
                        skipped_count += 1
                        continue

            parsed_url = urlparse(img_url)
            filename = os.path.basename(parsed_url.path.split('?')[0])

            if not filename or "." not in filename:
                filename = f"image_{idx}.jpg"
            else:
                filename = sanitize_filename(filename)

            filepath = os.path.join(output_folder, filename)

            with open(filepath, "wb") as f:
                f.write(img_data.content)
            
            log_func(f"[{download_count+1}] Saved: {filename} ({width}x{height})[cite: 4]")
            downloaded_count += 1

        except Exception as err:
            log_func(f"[{idx}/{len(img_tags)}] Download error: {err}[cite: 4]")

    log_func(f"\nFinished! Downloaded {downloaded_count} image(s), Skipped {skipped_count} icons/filtered items to '{output_folder}'.[cite: 4]")
    return downloaded_count


# =============================================================================
# 2. ZIGZAG / BORDER PATTERN ENGINE
# =============================================================================
def create_pattern_mask(width, height, tooth_width, tooth_depth, shape_type="Zigzag (Sawtooth)", corner_style="Mitred (Sharp)", corner_size=20, auto_scale=False):
    """Generates a grayscale mask (0-255) with edge shapes and refined corners[cite: 3]."""
    if auto_scale:
        base_dim = min(width, height)
        scale_factor = max(1.0, base_dim / 1000.0)
        tooth_width = int(tooth_width * scale_factor)
        tooth_depth = int(tooth_depth * scale_factor)
        corner_size = int(corner_size * scale_factor)

    tooth_width = max(4, min(tooth_width, width // 2, height // 2))
    tooth_depth = max(1, min(tooth_depth, width // 4, height // 4))
    corner_size = max(0, min(corner_size, width // 3, height // 3))

    mask = Image.new("L", (width, height), 0)
    draw = ImageDraw.Draw(mask)

    if shape_type == "Zigzag (Sawtooth)":
        polygon_points = []
        x = corner_size
        while x < width - corner_size:
            polygon_points.append((x, tooth_depth))
            polygon_points.append((min(x + tooth_width // 2, width - corner_size), 0))
            x += tooth_width
        polygon_points.append((width - corner_size, tooth_depth))

        if corner_style == "Bevelled (Flat Chamfer)":
            polygon_points.append((width - tooth_depth, corner_size))
        elif corner_style == "Mitred (Sharp)":
            polygon_points.append((width, 0))
            polygon_points.append((width - tooth_depth, corner_size))

        y = corner_size
        while y < height - corner_size:
            polygon_points.append((width - tooth_depth, y))
            polygon_points.append((width, min(y + tooth_width // 2, height - corner_size)))
            y += tooth_width
        polygon_points.append((width - tooth_depth, height - corner_size))

        if corner_style == "Bevelled (Flat Chamfer)":
            polygon_points.append((width - corner_size, height - tooth_depth))
        elif corner_style == "Mitred (Sharp)":
            polygon_points.append((width, height))
            polygon_points.append((width - corner_size, height - tooth_depth))

        x = width - corner_size
        while x > corner_size:
            polygon_points.append((x, height - tooth_depth))
            polygon_points.append((max(x - tooth_width // 2, corner_size), height))
            x -= tooth_width
        polygon_points.append((corner_size, height - tooth_depth))

        if corner_style == "Bevelled (Flat Chamfer)":
            polygon_points.append((tooth_depth, height - corner_size))
        elif corner_style == "Mitred (Sharp)":
            polygon_points.append((0, height))
            polygon_points.append((tooth_depth, height - corner_size))

        y = height - corner_size
        while y > corner_size:
            polygon_points.append((tooth_depth, y))
            polygon_points.append((0, max(y - tooth_width // 2, corner_size)))
            y -= tooth_width
        polygon_points.append((tooth_depth, corner_size))

        if corner_style == "Bevelled (Flat Chamfer)":
            polygon_points.append((corner_size, tooth_depth))
        elif corner_style == "Mitred (Sharp)":
            polygon_points.append((0, 0))
            polygon_points.append((corner_size, tooth_depth))

        draw.polygon(polygon_points, fill=255)

        if corner_style == "Rounded (Fillet)" and corner_size > 0:
            corner_mask = Image.new("L", (width, height), 255)
            c_draw = ImageDraw.Draw(corner_mask)
            r = corner_size * 2
            c_draw.pieslice([0, 0, r, r], 180, 270, fill=0)
            c_draw.pieslice([width - r, 0, width, r], 270, 360, fill=0)
            c_draw.pieslice([width - r, height - r, width, height], 0, 90, fill=0)
            c_draw.pieslice([0, height - r, r, height], 90, 180, fill=0)
            mask = Image.eval(Image.composite(mask, Image.new("L", (width, height), 0), corner_mask), lambda p: p)

    elif shape_type == "Square Stamp (Perforated)":
        polygon_points = []
        half_w = tooth_width // 2
        x = corner_size
        while x < width - corner_size:
            polygon_points.append((x, tooth_depth))
            polygon_points.append((min(x + half_w, width - corner_size), tooth_depth))
            polygon_points.append((min(x + half_w, width - corner_size), 0))
            polygon_points.append((min(x + tooth_width, width - corner_size), 0))
            x += tooth_width
        polygon_points.append((width - corner_size, tooth_depth))
        draw.polygon(polygon_points, fill=255)

    return mask


# =============================================================================
# 3. ADVANCED IMAGE PROCESSING & CATALOG GENERATION ENGINE
# =============================================================================
def process_and_optimize_image(img_input, remove_bg=True, tight_crop=True, auto_rotate=False, 
                               bg_model="birefnet-general", alpha_matting=False, apply_border=False, 
                               border_cfg=None, preset_name="High Quality (Balanced - Default)", 
                               format_mode="Auto (JPEG for opaque, PNG for transparent)"):
    """Processes image with high-accuracy AI background removal models, alpha matting, and cropping."""
    preset = QUALITY_PRESETS.get(preset_name, QUALITY_PRESETS["High Quality (Balanced - Default)"])
    max_dim = preset["max_dim"]
    jpg_quality = preset["jpg_quality"]

    if isinstance(img_input, str):
        img = Image.open(img_input)
    elif isinstance(img_input, bytes):
        img = Image.open(io.BytesIO(img_input))
    else:
        img = img_input.copy()

    if auto_rotate:
        try:
            img = ImageOps.exif_transpose(img)
        except Exception:
            pass

    img = img.convert("RGBA")

    # 1. High-Precision Background Removal
    if remove_bg:
        img_bytes = io.BytesIO()
        img.save(img_bytes, format="PNG")
        
        try:
            session = new_session(bg_model)
            no_bg_bytes = remove(
                img_bytes.getvalue(), 
                session=session, 
                alpha_matting=alpha_matting,
                alpha_matting_foreground_threshold=240,
                alpha_matting_background_threshold=10
            )
        except Exception:
            # Fallback to standard remove if model session fails
            no_bg_bytes = remove(img_bytes.getvalue(), alpha_matting=alpha_matting)

        img = Image.open(io.BytesIO(no_bg_bytes)).convert("RGBA")

    # 2. Tight Crop
    if tight_crop:
        bbox = img.getbbox()
        if bbox:
            img = img.crop(bbox)

    # 3. Optional Border Effect (Zigzag/Wave)
    if apply_border and border_cfg:
        w, h = img.size
        mask = create_pattern_mask(
            width=w, height=h,
            tooth_width=border_cfg["tooth_width"],
            tooth_depth=border_cfg["tooth_depth"],
            shape_type=border_cfg["shape_type"],
            corner_style=border_cfg["corner_style"],
            corner_size=border_cfg["corner_size"],
            auto_scale=border_cfg["auto_scale"]
        )
        result = Image.new("RGBA", (w, h), (0, 0, 0, 0))
        result.paste(img, (0, 0), mask=mask)
        img = result

    # 4. Scaling
    orig_w, orig_h = img.size
    if max(orig_w, orig_h) > max_dim:
        scale = max_dim / float(max(orig_w, orig_h))
        new_w = max(1, int(orig_w * scale))
        new_h = max(1, int(orig_h * scale))
        img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)

    output_stream = io.BytesIO()
    use_png = True if ("Always PNG" in format_mode or (img.mode == 'RGBA' and apply_border)) else False

    if use_png:
        img.save(output_stream, format="PNG", optimize=True)
    else:
        if img.mode == 'RGBA':
            background = Image.new("RGB", img.size, (255, 255, 255))
            background.paste(img, mask=img.split()[3])
            img = background
        else:
            img = img.convert("RGB")
        img.save(output_stream, format="JPEG", quality=jpg_quality, optimize=True, progressive=True)

    output_stream.seek(0)
    return output_stream, img.size


def generate_default_labels(identifier):
    filename = os.path.splitext(os.path.basename(identifier))[0]
    clean_name = filename.replace('_', ' ').replace('-', ' ')
    parts = clean_name.split()
    if len(parts) >= 3:
        return [f"Code: {parts[0].upper()}", f"Style: {' '.join(parts[1:-1]).title()}", f"Variant: {parts[-1].title()}"]
    elif len(parts) == 2:
        return [f"Item Code: {parts[0].upper()}", f"Description: {parts[1].title()}"]
    else:
        return [f"Label: {clean_name.title()}", f"ID: {os.path.basename(identifier)}"]


def calculate_grid_positions(rows, cols, slide_width=Inches(13.333), slide_height=Inches(7.5)):
    margin_x = Inches(0.6)
    margin_y_top = Inches(1.2)
    margin_y_bottom = Inches(0.5)
    gap_x = Inches(0.3)
    gap_y = Inches(0.3)

    available_w = slide_width - (2 * margin_x) - ((cols - 1) * gap_x)
    available_h = slide_height - margin_y_top - margin_y_bottom - ((rows - 1) * gap_y)

    card_w = available_w / cols
    card_h = available_h / rows

    configs = []
    for r in range(rows):
        for c in range(cols):
            left = margin_x + c * (card_w + gap_x)
            top = margin_y_top + r * (card_h + gap_y)
            configs.append({"left": left, "top": top, "width": card_w, "height": card_h})
    return configs


def add_slide_border_and_header(slide, slide_width, slide_height, title_text):
    margin = Inches(0.35)
    border_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, margin, margin, slide_width - (2 * margin), slide_height - (2 * margin))
    border_box.fill.background()
    border_box.line.color.rgb = COLOR_BORDER
    border_box.line.width = Pt(1.5)

    header_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.45), Inches(12.133), Inches(0.6))
    tf = header_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title_text.upper()
    p.font.name = FONT_NAME
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = COLOR_HEADER


def add_formatted_multilabel(slide, left, top, width, height, labels, font_size=10):
    lbl_box = slide.shapes.add_textbox(left, top, width, height)
    ltf = lbl_box.text_frame
    ltf.word_wrap = True
    ltf.margin_top = ltf.margin_bottom = ltf.margin_left = ltf.margin_right = Inches(0)

    for i, line in enumerate(labels):
        p = ltf.paragraphs[0] if i == 0 else ltf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        if ":" in line:
            key, val = line.split(":", 1)
            run_key = p.add_run()
            run_key.text = f"{key.strip()}: "
            run_key.font.name, run_key.font.size, run_key.font.bold, run_key.font.color.rgb = FONT_NAME, Pt(font_size), True, COLOR_LABEL_KEY

            run_val = p.add_run()
            run_val.text = val.strip()
            run_val.font.name, run_val.font.size, run_val.font.bold, run_val.font.color.rgb = FONT_NAME, Pt(font_size), False, COLOR_LABEL_VAL
        else:
            run = p.add_run()
            run.text = line
            run.font.name, run.font.size, run.font.color.rgb = FONT_NAME, Pt(font_size), COLOR_LABEL_VAL


def create_catalog_presentation(image_paths, output_file, rows, cols, remove_bg, tight_crop, auto_rotate, bg_model, alpha_matting, apply_border, border_cfg, preset_name, format_mode, log_func, update_progress):
    items_per_slide = rows * cols
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    card_configs = calculate_grid_positions(rows, cols, prs.slide_width, prs.slide_height)
    image_groups = [image_paths[i:i + items_per_slide] for i in range(0, len(image_paths), items_per_slide)]
    total_groups = len(image_groups)

    for slide_idx, group in enumerate(image_groups):
        log_func(f"Building catalog slide {slide_idx + 1} of {total_groups}...")
        slide = prs.slides.add_slide(blank_layout)
        add_slide_border_and_header(slide, prs.slide_width, prs.slide_height, f"GARMENT CATALOG PAGE {slide_idx + 1}")

        for idx, img_path in enumerate(group):
            cfg = card_configs[idx]
            labels = generate_default_labels(img_path)
            font_sz = 9 if items_per_slide > 4 else 11
            label_block_height = Inches(max(0.4, len(labels) * (font_sz * 0.015) + 0.1))

            optimized_stream, (crop_w, crop_h) = process_and_optimize_image(
                img_path, remove_bg=remove_bg, tight_crop=tight_crop, auto_rotate=auto_rotate,
                bg_model=bg_model, alpha_matting=alpha_matting,
                apply_border=apply_border, border_cfg=border_cfg,
                preset_name=preset_name, format_mode=format_mode
            )

            card_bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cfg["left"], cfg["top"], cfg["width"], cfg["height"])
            card_bg.fill.solid()
            card_bg.fill.fore_color.rgb = COLOR_CARD_BG
            card_bg.line.color.rgb = RGBColor(226, 232, 240)

            image_max_w = cfg["width"] - Inches(0.25)
            image_max_h = cfg["height"] - label_block_height - Inches(0.25)

            aspect_ratio = crop_w / crop_h if crop_h > 0 else 1.0
            target_w = image_max_w
            target_h = target_w / aspect_ratio
            if target_h > image_max_h:
                target_h = image_max_h
                target_w = target_h * aspect_ratio

            img_left = cfg["left"] + (cfg["width"] - target_w) / 2
            img_top = cfg["top"] + Inches(0.12) + (image_max_h - target_h) / 2

            slide.shapes.add_picture(optimized_stream, img_left, img_top, target_w, target_h)
            lbl_top = cfg["top"] + Inches(0.12) + image_max_h + Inches(0.05)
            add_formatted_multilabel(slide, cfg["left"] + Inches(0.15), lbl_top, cfg["width"] - Inches(0.3), label_block_height, labels, font_size=font_sz)

        update_progress((slide_idx + 1) / total_groups * 100)

    prs.save(output_file)
    size_mb = os.path.getsize(output_file) / (1024 * 1024)
    log_func(f"Catalog saved successfully: {output_file} ({size_mb:.2f} MB)")


# =============================================================================
# UNIFIED INTERACTIVE GUI DASHBOARD
# =============================================================================
class UnifiedCatalogStudio(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Unified Garment Catalog Studio (Downloader -> RemBG -> Border -> Catalog)")
        self.geometry("900x1000")
        self.resizable(True, True)

        style = ttk.Style(self)
        style.theme_use("clam")

        # Variables
        self.enable_download_var = tk.BooleanVar(value=True)
        self.web_url_var = tk.StringVar(value="https://")
        self.input_folder_var = tk.StringVar(value=os.path.abspath("garment_photos"))
        self.output_folder_var = tk.StringVar(value=os.path.abspath("catalog_output"))

        # Filtering variables
        self.avoid_icons_var = tk.BooleanVar(value=True)
        self.filter_mode_var = tk.StringVar(value="All Images")

        # Processing variables
        self.remove_bg_var = tk.BooleanVar(value=True)
        self.tight_crop_var = tk.BooleanVar(value=True)
        self.auto_rotate_var = tk.BooleanVar(value=False)
        self.bg_model_var = tk.StringVar(value="birefnet-general")  # Superior accuracy model for garments
        self.alpha_matting_var = tk.BooleanVar(value=True)          # High precision edge refinement

        self.enable_border_var = tk.BooleanVar(value=False)
        self.shape_var = tk.StringVar(value="Zigzag (Sawtooth)")
        self.corner_style_var = tk.StringVar(value="Mitred (Sharp)")
        self.corner_size_var = tk.IntVar(value=20)
        self.depth_var = tk.IntVar(value=15)
        self.width_var = tk.IntVar(value=30)
        self.scale_var = tk.BooleanVar(value=True)

        self.rows_var = tk.IntVar(value=2)
        self.cols_var = tk.IntVar(value=2)
        self.preset_var = tk.StringVar(value="High Quality (Balanced - Default)")
        self.format_var = tk.StringVar(value="Auto (JPEG for opaque, PNG for transparent)")

        self._build_ui()

    def _build_ui(self):
        header = tk.Frame(self, bg="#1e293b", height=60)
        header.pack(fill=tk.X)
        tk.Label(header, text="Garment Catalog & Automation Studio", font=("Segoe UI", 15, "bold"), fg="#ffffff", bg="#1e293b").pack(pady=14)

        main_container = ttk.Frame(self, padding=10)
        main_container.pack(fill=tk.BOTH, expand=True)

        # 1. Web Downloader Section with Filtering Options
        dl_frame = ttk.LabelFrame(main_container, text=" Step 1: Web Image Downloader & Smart Filtering ", padding=8)
        dl_frame.pack(fill=tk.X, pady=3)
        ttk.Checkbutton(dl_frame, text="Run image downloader automatically as part of pipeline", variable=self.enable_download_var).pack(anchor=tk.W, pady=2)
        
        url_row = ttk.Frame(dl_frame)
        url_row.pack(fill=tk.X, pady=2)
        ttk.Label(url_row, text="Website URL:").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(url_row, textvariable=self.web_url_var, width=50).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        self.download_only_btn = ttk.Button(url_row, text="📥 Download Images First", command=self._start_download_only_thread)
        self.download_only_btn.pack(side=tk.RIGHT)

        filter_row = ttk.Frame(dl_frame)
        filter_row.pack(fill=tk.X, pady=5)
        ttk.Checkbutton(filter_row, text="Avoid Icons & UI Elements", variable=self.avoid_icons_var).pack(side=tk.LEFT, padx=(0, 15))
        
        ttk.Label(filter_row, text="Image Type:").pack(side=tk.LEFT, padx=(5, 5))
        ttk.Combobox(
            filter_row, 
            textvariable=self.filter_mode_var, 
            values=["All Images", "Garment Only (Flat-lay / Swatch)", "Garment with Model"], 
            state="readonly", 
            width=32
        ).pack(side=tk.LEFT)

        # 2. Input / Output Directories
        io_frame = ttk.LabelFrame(main_container, text=" Step 2: Target Image Folder & Output Catalog Directory ", padding=8)
        io_frame.pack(fill=tk.X, pady=3)

        in_row = ttk.Frame(io_frame)
        in_row.pack(fill=tk.X, pady=2)
        ttk.Label(in_row, text="Image Folder:", width=15).pack(side=tk.LEFT)
        ttk.Entry(in_row, textvariable=self.input_folder_var, width=60).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(in_row, text="Browse...", command=self._browse_input_folder).pack(side=tk.RIGHT)

        out_row = ttk.Frame(io_frame)
        out_row.pack(fill=tk.X, pady=2)
        ttk.Label(out_row, text="Output Directory:", width=15).pack(side=tk.LEFT)
        ttk.Entry(out_row, textvariable=self.output_folder_var, width=60).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(out_row, text="Browse...", command=self._browse_output_folder).pack(side=tk.RIGHT)

        # 3. Processing Options (Remove BG, Precision Model, Alpha Matting, & Crop)
        proc_frame = ttk.LabelFrame(main_container, text=" Step 3: Background Removal, Precision Engine & Cropping ", padding=8)
        proc_frame.pack(fill=tk.X, pady=3)
        
        ttk.Checkbutton(proc_frame, text="AI Background Removal (rembg)", variable=self.remove_bg_var).grid(row=0, column=0, sticky=tk.W, pady=2)
        ttk.Checkbutton(proc_frame, text="Crop Tight to Object Bounds", variable=self.tight_crop_var).grid(row=0, column=1, sticky=tk.W, padx=20, pady=2)
        ttk.Checkbutton(proc_frame, text="🔄 Auto-Rotate Orientation (Optional)", variable=self.auto_rotate_var).grid(row=1, column=0, sticky=tk.W, pady=2)
        ttk.Checkbutton(proc_frame, text="✨ Alpha Matting (High Precision Edges)", variable=self.alpha_matting_var).grid(row=1, column=1, sticky=tk.W, padx=20, pady=2)

        model_row = ttk.Frame(proc_frame)
        model_row.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=5)
        ttk.Label(model_row, text="AI Model Accuracy:").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Combobox(
            model_row,
            textvariable=self.bg_model_var,
            values=["birefnet-general", "u2net", "isnet-general-use", "u2netp"],
            state="readonly",
            width=30
        ).pack(side=tk.LEFT, padx=5)
        ttk.Label(model_row, text="(birefnet-general provides best garment accuracy)", font=("Segoe UI", 8, "italic"), foreground="#64748b").pack(side=tk.LEFT, padx=5)

        # 4. Border / Zigzag Effect Options (Optional)
        border_frame = ttk.LabelFrame(main_container, text=" Step 4: Optional Custom Border Effect (Zigzag/Wave) ", padding=8)
        border_frame.pack(fill=tk.X, pady=3)
        ttk.Checkbutton(border_frame, text="Apply Zigzag / Pattern Border to Images", variable=self.enable_border_var).pack(anchor=tk.W, pady=2)

        b_grid = ttk.Frame(border_frame)
        b_grid.pack(fill=tk.X, pady=2)
        ttk.Label(b_grid, text="Pattern Style:").grid(row=0, column=0, sticky=tk.W, pady=2)
        ttk.Combobox(b_grid, textvariable=self.shape_var, values=["Zigzag (Sawtooth)", "Square Stamp (Perforated)"], state="readonly", width=25).grid(row=0, column=1, sticky=tk.W, padx=5)

        ttk.Label(b_grid, text="Depth (px):").grid(row=0, column=2, sticky=tk.W, padx=10, pady=2)
        ttk.Spinbox(b_grid, from_=2, to=50, textvariable=self.depth_var, width=5).grid(row=0, column=3, sticky=tk.W)

        # 5. Catalog Layout & Quality Settings
        cat_frame = ttk.LabelFrame(main_container, text=" Step 5: Catalog Grid & Presentation Layout ", padding=8)
        cat_frame.pack(fill=tk.X, pady=3)

        c_grid = ttk.Frame(cat_frame)
        c_grid.pack(fill=tk.X, pady=2)
        ttk.Label(c_grid, text="Grid Rows:").grid(row=0, column=0, sticky=tk.W)
        ttk.Spinbox(c_grid, from_=1, to=4, textvariable=self.rows_var, width=5).grid(row=0, column=1, sticky=tk.W, padx=(5, 20))
        ttk.Label(c_grid, text="Grid Columns:").grid(row=0, column=2, sticky=tk.W)
        ttk.Spinbox(c_grid, from_=1, to=4, textvariable=self.cols_var, width=5).grid(row=0, column=3, sticky=tk.W, padx=5)

        ttk.Label(cat_frame, text="Quality Preset:").pack(anchor=tk.W, pady=(5, 2))
        ttk.Combobox(cat_frame, textvariable=self.preset_var, values=list(QUALITY_PRESETS.keys()), state="readonly", width=50).pack(anchor=tk.W)

        # Execution Button & Log Console
        exec_frame = ttk.Frame(main_container)
        exec_frame.pack(fill=tk.X, pady=8)

        self.run_btn = tk.Button(exec_frame, text="RUN COMPLETE AUTOMATION PIPELINE", font=("Segoe UI", 11, "bold"), bg="#0f172a", fg="#ffffff", command=self._start_pipeline_thread, pady=8)
        self.run_btn.pack(fill=tk.X)

        self.progress = ttk.Progressbar(main_container, orient=tk.HORIZONTAL, mode='determinate')
        self.progress.pack(fill=tk.X, pady=3)

        log_frame = ttk.LabelFrame(main_container, text=" Live Execution Logs ", padding=5)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=3)

        self.log_text = tk.Text(log_frame, height=6, font=("Consolas", 9), state=tk.DISABLED, bg="#f8fafc")
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def _browse_input_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.input_folder_var.set(folder)

    def _browse_output_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_folder_var.set(folder)

    def _log(self, message):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def _update_progress(self, val):
        self.progress['value'] = val
        self.update_idletasks()

    def _start_download_only_thread(self):
        url = self.web_url_var.get().strip()
        folder = self.input_folder_var.get().strip()
        if not url or url == "https://":
            messagebox.showerror("Error", "Please enter a valid website URL.")
            return
        if not folder:
            messagebox.showerror("Error", "Please select an image folder path first.")
            return

        self.download_only_btn.config(state=tk.DISABLED)
        threading.Thread(
            target=self._run_download_only, 
            args=(url, folder, self.avoid_icons_var.get(), self.filter_mode_var.get()), 
            daemon=True
        ).start()

    def _run_download_only(self, url, folder, avoid_icons, filter_mode):
        try:
            self._log("==========================================")
            self._log("MANUAL DOWNLOAD WITH FILTERING INITIATED")
            self._log("==========================================")
            count = download_images_from_url(url, folder, self._log, avoid_icons, filter_mode)
            self._log(f"Finished downloading {count} images to '{folder}'.")
            messagebox.showinfo("Download Complete", f"Successfully downloaded {count} filtered images!")
        except Exception as e:
            self._log(f"Download Error: {str(e)}")
            messagebox.showerror("Error", str(e))
        finally:
            self.download_only_btn.config(state=tk.NORMAL)

    def _start_pipeline_thread(self):
        self.run_btn.config(state=tk.DISABLED, text="PROCESSING PIPELINE...")
        threading.Thread(target=self._run_pipeline, daemon=True).start()

    def _run_pipeline(self):
        try:
            in_folder = self.input_folder_var.get().strip()
            out_folder = self.output_folder_var.get().strip()
            os.makedirs(out_folder, exist_ok=True)

            if self.enable_download_var.get():
                url = self.web_url_var.get().strip()
                if url and url != "https://":
                    self._log("==========================================")
                    self._log("STEP 1: DOWNLOADING & FILTERING IMAGES")
                    self._log("==========================================")
                    download_images_from_url(
                        url, in_folder, self._log, 
                        self.avoid_icons_var.get(), self.filter_mode_var.get()
                    )
                else:
                    self._log("Step 1 skipped: Invalid or empty URL provided.")

            self._log("==========================================")
            self._log("STEP 2: GATHERING IMAGES FROM FOLDER")
            self._log("==========================================")
            extensions = ('*.jpg', '*.jpeg', '*.png', '*.webp', '*.bmp')
            image_paths = []
            for ext in extensions:
                image_paths.extend(glob.glob(os.path.join(in_folder, ext)))
                image_paths.extend(glob.glob(os.path.join(in_folder, ext.upper())))
            image_paths = sorted(list(set(image_paths)))

            if not image_paths:
                messagebox.showerror("Error", f"No valid images found in folder path: {in_folder}\nPlease download images first or select a valid folder.")
                self.run_btn.config(state=tk.NORMAL, text="RUN COMPLETE AUTOMATION PIPELINE")
                return

            self._log(f"Found {len(image_paths)} images ready for catalog processing.")

            border_cfg = None
            if self.enable_border_var.get():
                border_cfg = {
                    "tooth_width": self.width_var.get(),
                    "tooth_depth": self.depth_var.get(),
                    "shape_type": self.shape_var.get(),
                    "corner_style": self.corner_style_var.get(),
                    "corner_size": self.corner_size_var.get(),
                    "auto_scale": self.scale_var.get()
                }

            self._log("==========================================")
            self._log("STEPS 3-5: PROCESSING & BUILDING CATALOG")
            self._log("==========================================")
            output_ppt = os.path.join(out_folder, "Garment_Catalog_Master.pptx")
            
            create_catalog_presentation(
                image_paths=image_paths,
                output_file=output_ppt,
                rows=self.rows_var.get(),
                cols=self.cols_var.get(),
                remove_bg=self.remove_bg_var.get(),
                tight_crop=self.tight_crop_var.get(),
                auto_rotate=self.auto_rotate_var.get(),
                bg_model=self.bg_model_var.get(),
                alpha_matting=self.alpha_matting_var.get(),
                apply_border=self.enable_border_var.get(),
                border_cfg=border_cfg,
                preset_name=self.preset_var.get(),
                format_mode=self.format_var.get(),
                log_func=self._log,
                update_progress=self._update_progress
            )

            self._log("==========================================")
            self._log("PIPELINE COMPLETED SUCCESSFULLY!")
            self._log("==========================================")
            messagebox.showinfo("Success", f"Catalog successfully created!\nSaved to:\n{output_ppt}")

        except Exception as e:
            self._log(f"CRITICAL ERROR: {str(e)}")
            messagebox.showerror("Error", str(e))
        finally:
            self.run_btn.config(state=tk.NORMAL, text="RUN COMPLETE AUTOMATION PIPELINE")


if __name__ == "__main__":
    app = UnifiedCatalogStudio()
    app.mainloop()

