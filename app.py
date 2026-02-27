import streamlit as st
import pandas as pd
import qrcode
from io import BytesIO
from PIL import Image, ImageDraw, ImageFont
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import matplotlib.colors as mc
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, A3, A5, A6, B4, B5, letter, legal, landscape
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import stringWidth
from matplotlib.font_manager import FontProperties, fontManager
import textwrap
import tempfile
import os
import numpy as np
import glob

# ──────────────────────────────────────────────
# Page sizes
# ──────────────────────────────────────────────
PAGE_SIZES = {
    "A3 (297×420 mm)": A3,
    "A4 (210×297 mm)": A4,
    "A5 (148×210 mm)": A5,
    "A6 (105×148 mm)": A6,
    "B4 (250×353 mm)": B4,
    "B5 (176×250 mm)": B5,
    "Letter (216×279 mm)": letter,
    "Legal (216×356 mm)": legal,
}

COLORS = [
    "#e74c3c", "#3498db", "#2ecc71", "#f39c12",
    "#9b59b6", "#1abc9c", "#e67e22", "#34495e",
    "#d35400", "#16a085", "#c0392b", "#2980b9",
]

def find_thai_font() -> str | None:
    bundled = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts", "Sarabun-Regular.ttf")
    if os.path.exists(bundled):
        return bundled

    candidates = [
        "/System/Library/Fonts/Supplemental/Ayuthaya.ttf",
        "/System/Library/Fonts/Supplemental/Sathu.ttf",
        "/usr/share/fonts/truetype/thai-tlwg/TlwgTypo.ttf",
        "/usr/share/fonts/truetype/thai-tlwg/Loma.ttf",
        "/usr/share/fonts/truetype/thai-tlwg/Garuda.ttf",
        "/usr/share/fonts/truetype/thai-tlwg/Sarabun.ttf",
        "/usr/share/fonts/truetype/tlwg/TlwgTypo.ttf",
        "/usr/share/fonts/truetype/tlwg/Loma.ttf",
        "/usr/share/fonts/truetype/tlwg/Garuda.ttf",
        "/usr/share/fonts/truetype/tlwg/Garuda-Bold.ttf",
        "/usr/share/fonts/truetype/tlwg/Sarabun.ttf",
        "/usr/share/fonts/truetype/noto/NotoSansThai-Regular.ttf",
        "/usr/share/fonts/opentype/noto/NotoSansThai-Regular.otf",
        "C:/Windows/Fonts/tahoma.ttf",
    ]
    for path in candidates:
        if os.path.exists(path):
            return path

    search_patterns = [
        "/usr/share/fonts/**/Garuda*.ttf",
        "/usr/share/fonts/**/Loma*.ttf",
        "/usr/share/fonts/**/Sarabun*.ttf",
        "/usr/share/fonts/**/TlwgTypo*.ttf",
        "/usr/share/fonts/**/NotoSansThai*.ttf",
        "/usr/share/fonts/**/NotoSansThai*.otf",
        "/usr/share/fonts/**/*Thai*.ttf",
        "/usr/share/fonts/**/*thai*.ttf",
    ]
    for pattern in search_patterns:
        results = glob.glob(pattern, recursive=True)
        if results:
            return results[0]

    return None

_THAI_FONT_PATH = find_thai_font()
_THAI_FONT_NAME = "Helvetica"
if _THAI_FONT_PATH:
    try:
        pdfmetrics.registerFont(TTFont("ThaiFont", _THAI_FONT_PATH))
        _THAI_FONT_NAME = "ThaiFont"
    except Exception:
        pass
    try:
        fontManager.addfont(_THAI_FONT_PATH)
    except Exception:
        pass

def render_thai_text_image(text: str, font_path: str, font_size_pt: float,
                          width_mm: float, align: str = "center",
                          color_hex: str = "#555555") -> tuple:
    render_scale = 8
    pixel_size = max(12, int(font_size_pt * render_scale))
    mm_per_px = (font_size_pt * 0.353) / pixel_size
    target_w_px = max(20, int(width_mm / mm_per_px))

    try:
        font = ImageFont.truetype(font_path, pixel_size)
    except Exception:
        return None, 0, 0

    def measure_px(s):
        bb = font.getbbox(s)
        return bb[2] - bb[0] if bb else 0

    if " " in text:
        tokens = text.split(" ")
        lines, cur = [], ""
        for tok in tokens:
            trial = (cur + " " + tok).strip()
            if measure_px(trial) <= target_w_px:
                cur = trial
            else:
                if cur:
                    lines.append(cur)
                cur = tok
        if cur:
            lines.append(cur)
    else:
        lines, cur = [], ""
        for ch in text:
            trial = cur + ch
            if measure_px(trial) <= target_w_px:
                cur = trial
            else:
                if cur:
                    lines.append(cur)
                cur = ch
        if cur:
            lines.append(cur)
    if not lines:
        lines = [text]

    line_spacing = int(pixel_size * 1.5)
    canvas_w = target_w_px + 16
    canvas_h = line_spacing * len(lines) + pixel_size
    img = Image.new("RGBA", (canvas_w, canvas_h), (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)

    r = int(color_hex[1:3], 16)
    g = int(color_hex[3:5], 16)
    b = int(color_hex[5:7], 16)
    fill = (r, g, b, 255)

    pil_align = {"ซ้าย": "left", "ขวา": "right"}.get(align, "center")

    for i, line in enumerate(lines):
        ly = i * line_spacing
        lw = measure_px(line)
        if pil_align == "left":
            lx = 4
        elif pil_align == "right":
            lx = canvas_w - lw - 4
        else:
            lx = (canvas_w - lw) // 2
        draw.text((lx, ly), line, font=font, fill=fill)

    bbox = img.getbbox()
    if bbox:
        img = img.crop((0, 0, canvas_w, bbox[3] + 4))

    w_mm = img.width * mm_per_px
    h_mm = img.height * mm_per_px
    return np.array(img), w_mm, h_mm

def render_thai_text_pil(text: str, font_path: str, font_size_pt: float,
                         width_mm: float, align: str = "center",
                         color_hex: str = "#000000") -> tuple:
    render_scale = 12
    pixel_size = max(16, int(font_size_pt * render_scale))
    mm_per_px = (font_size_pt * 0.353) / pixel_size
    target_w_px = max(20, int(width_mm / mm_per_px))

    try:
        font = ImageFont.truetype(font_path, pixel_size)
    except Exception:
        return None, 0, 0

    def measure_px(s):
        bb = font.getbbox(s)
        return bb[2] - bb[0] if bb else 0

    if " " in text:
        tokens = text.split(" ")
        lines, cur = [], ""
        for tok in tokens:
            trial = (cur + " " + tok).strip()
            if measure_px(trial) <= target_w_px:
                cur = trial
            else:
                if cur:
                    lines.append(cur)
                cur = tok
        if cur:
            lines.append(cur)
    else:
        lines, cur = [], ""
        for ch in text:
            trial = cur + ch
            if measure_px(trial) <= target_w_px:
                cur = trial
            else:
                if cur:
                    lines.append(cur)
                cur = ch
        if cur:
            lines.append(cur)
    if not lines:
        lines = [text]

    line_spacing = int(pixel_size * 1.5)
    canvas_w = target_w_px + 16
    canvas_h = line_spacing * len(lines) + pixel_size
    img = Image.new("RGBA", (canvas_w, canvas_h), (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)

    r = int(color_hex[1:3], 16)
    g = int(color_hex[3:5], 16)
    b = int(color_hex[5:7], 16)
    fill = (r, g, b, 255)

    pil_align = {"ซ้าย": "left", "ขวา": "right"}.get(align, "center")

    for i, line in enumerate(lines):
        ly = i * line_spacing
        lw = measure_px(line)
        if pil_align == "left":
            lx = 4
        elif pil_align == "right":
            lx = canvas_w - lw - 4
        else:
            lx = (canvas_w - lw) // 2
        draw.text((lx, ly), line, font=font, fill=fill)

    bbox = img.getbbox()
    if bbox:
        img = img.crop((0, 0, canvas_w, bbox[3] + 4))

    bg = Image.new("RGB", img.size, (255, 255, 255))
    bg.paste(img, mask=img.split()[3])

    w_mm = bg.width * mm_per_px
    h_mm = bg.height * mm_per_px
    return bg, w_mm, h_mm

def wrap_text_for_preview(text: str, font_size_pt: float, width_mm: float) -> list:
    char_w_mm = max(0.5, font_size_pt * 0.353 * 0.55)
    chars_per_line = max(3, int(width_mm / char_w_mm))
    if " " in text:
        return textwrap.wrap(text, width=chars_per_line) or [text]
    else:
        return [text[i:i + chars_per_line] for i in range(0, len(text), chars_per_line)] or [text]

def wrap_text_for_pdf(text: str, font_name: str, font_size: float, max_width_pt: float) -> list:
    def measure(s):
        try:
            return stringWidth(s, font_name, font_size)
        except Exception:
            return len(s) * font_size * 0.5

    if " " in text:
        words = text.split()
        lines, current = [], ""
        for word in words:
            candidate = (current + " " + word).strip()
            if measure(candidate) <= max_width_pt:
                current = candidate
            else:
                if current:
                    lines.append(current)
                current = word
        if current:
            lines.append(current)
        return lines or [text]
    else:
        lines, current = [], ""
        for char in text:
            candidate = current + char
            if measure(candidate) <= max_width_pt:
                current = candidate
            else:
                if current:
                    lines.append(current)
                current = char
        if current:
            lines.append(current)
        return lines or [text]

def smart_str(val) -> str:
    if pd.isna(val) or val is None:
        return ""
    if isinstance(val, float) and val == int(val):
        return str(int(val))
    return str(val)

def generate_qr_image(data: str, size_px: int = 300) -> Image.Image:
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10,
        border=2,
    )
    qr.add_data(str(data))
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
    img = img.resize((size_px, size_px), Image.NEAREST)
    return img

def get_image_aspect_ratio(img_path):
    """คำนวณสัดส่วนรูปภาพ (Height / Width)"""
    if isinstance(img_path, str) and os.path.isfile(img_path):
        try:
            with Image.open(img_path) as img:
                return img.height / img.width
        except Exception:
            return 1.0
    return 1.0

def create_page_preview(
    page_w_mm, page_h_mm,
    qr_configs,
    total_pages,
):
    fig_h = 8
    fig_w = fig_h * page_w_mm / page_h_mm
    fig, ax = plt.subplots(1, 1, figsize=(max(4, fig_w), fig_h), dpi=100)

    stack_count = min(total_pages - 1, 4)
    for i in range(stack_count, 0, -1):
        offset = i * 1.8
        shadow = patches.FancyBboxPatch(
            (offset, offset), page_w_mm, page_h_mm,
            boxstyle="round,pad=0",
            linewidth=0.8, edgecolor="#bbb", facecolor="#f0f0f0",
            zorder=-i,
        )
        ax.add_patch(shadow)

    page_rect = patches.FancyBboxPatch(
        (0, 0), page_w_mm, page_h_mm,
        boxstyle="round,pad=0",
        linewidth=2, edgecolor="#333", facecolor="white",
    )
    ax.add_patch(page_rect)

    for cfg in qr_configs:
        x = cfg["x_mm"]
        y = cfg["y_mm"]
        w = cfg["width_mm"]
        h = cfg["height_mm"]
        color = cfg["color"]
        col_name = cfg["col_name"]
        value = smart_str(cfg["value"])
        label_value = smart_str(cfg.get("label_value", value))

        if not value: 
            continue

        is_active = cfg.get("is_active", False)
        border_w = 3.5 if is_active else 1.5
        qr_rect = patches.FancyBboxPatch(
            (x, y), w, h,
            boxstyle="round,pad=0.3",
            linewidth=border_w, edgecolor=color, facecolor="#fafafa",
        )
        ax.add_patch(qr_rect)

        try:
            if isinstance(value, str) and os.path.isfile(value) and value.lower().endswith(('.png', '.jpg', '.jpeg')):
                img = Image.open(value).convert("RGB")
                qr_arr = np.array(img)
            else:
                qr_img = generate_qr_image(value, size_px=200)
                qr_arr = np.array(qr_img)
                
            ax.imshow(
                qr_arr,
                extent=[x + 0.5, x + w - 0.5, y + h - 0.5, y + 0.5],
                aspect="auto", zorder=5, interpolation="nearest",
            )
        except Exception:
            ax.text(x + w / 2, y + h / 2, "IMG/QR", ha="center", va="center",
                    fontsize=max(6, w * 0.3), color="#333", weight="bold")

        badge_text = str(col_name)
        badge_w = max(10, len(badge_text) * 2.2 + 4)
        badge_h = 3.5
        badge_x = x + (w - badge_w) / 2
        badge_y = y - badge_h - 1.5
        if badge_y < -5:
            badge_y = y + h + 1.5

        badge = patches.FancyBboxPatch(
            (badge_x, badge_y), badge_w, badge_h,
            boxstyle="round,pad=0.5",
            linewidth=0, facecolor=color, alpha=0.9, zorder=6,
        )
        ax.add_patch(badge)
        ax.text(badge_x + badge_w / 2, badge_y + badge_h / 2, badge_text,
                ha="center", va="center", fontsize=5.5, color="white", weight="bold", zorder=7)

        if cfg.get("show_label", True):
            label_y_pos = y + h + 2
            if badge_y > y:
                label_y_pos = badge_y + badge_h + 1
            label_font_size = cfg.get("label_font_size", 7)
            label_x_offset = cfg.get("label_x_offset", 0)
            label_width_mm = max(5.0, float(cfg.get("label_width_mm", w)))
            label_x_center = x + w / 2 + label_x_offset
            label_x_left = label_x_center - label_width_mm / 2
            align_label = cfg.get("label_align", "กลาง")

            rendered = False
            if _THAI_FONT_PATH:
                arr, tw, th = render_thai_text_image(
                    label_value, _THAI_FONT_PATH, label_font_size, 
                    label_width_mm, align_label, "#555555")
                if arr is not None and tw > 0 and th > 0:
                    if align_label == "ซ้าย":
                        x0 = label_x_left
                    elif align_label == "ขวา":
                        x0 = label_x_left + label_width_mm - tw
                    else:
                        x0 = label_x_center - tw / 2

                    rgb = mc.to_rgb(color)
                    box_h = max(th, 3)
                    ax.add_patch(patches.Rectangle(
                        (label_x_left, label_y_pos - 0.5), label_width_mm, box_h + 1,
                        linewidth=0.6, edgecolor=color, facecolor=(*rgb, 0.06),
                        linestyle="--", zorder=6,
                    ))

                    ax.imshow(arr, extent=[x0, x0 + tw, label_y_pos + th, label_y_pos],
                              aspect="auto", zorder=7, interpolation="bilinear")
                    rendered = True

            if not rendered:
                lines = wrap_text_for_preview(label_value, label_font_size, label_width_mm)
                line_height_mm = label_font_size * 0.353 * 2.5
                total_h = len(lines) * line_height_mm + line_height_mm * 0.5
                rgb = mc.to_rgb(color)
                ax.add_patch(patches.Rectangle(
                    (label_x_left, label_y_pos - 0.5), label_width_mm, total_h,
                    linewidth=0.6, edgecolor=color, facecolor=(*rgb, 0.06),
                    linestyle="--", zorder=6,
                ))
                if align_label == "ซ้าย":
                    ha_mpl, text_x = "left", label_x_left + 0.5
                elif align_label == "ขวา":
                    ha_mpl, text_x = "right", label_x_left + label_width_mm - 0.5
                else:
                    ha_mpl, text_x = "center", label_x_center
                for li, line in enumerate(lines):
                    line_y = label_y_pos + li * line_height_mm
                    ax.text(text_x, line_y, line,
                            ha=ha_mpl, va="top", fontsize=label_font_size,
                            color="#555", zorder=7)

    pad = 8
    ax.set_xlim(-pad, page_w_mm + pad + stack_count * 2)
    ax.set_ylim(page_h_mm + pad + stack_count * 2, -pad)
    ax.set_aspect("equal")
    ax.axis("off")

    plt.tight_layout()
    return fig

def generate_pdf(
    df_selected,
    col_configs,
    page_size,
    orientation,
    progress_callback=None,
):
    if orientation == "Landscape":
        page_size = landscape(page_size)

    page_w, page_h = page_size
    total_rows = len(df_selected)

    pdf_buf = BytesIO()
    c = canvas.Canvas(pdf_buf, pagesize=page_size)

    for row_idx in range(total_rows):
        if row_idx > 0:
            c.showPage()

        for col_name, cfg in col_configs.items():
            raw = df_selected[col_name].iloc[row_idx]
            if pd.isna(raw) or raw is None or raw == "":
                continue
            
            value = smart_str(raw)
            x_pt = cfg["x_mm"] * mm
            w_pt = cfg["width_mm"] * mm
            h_pt = cfg["height_mm"] * mm
            y_pt = page_h - cfg["y_mm"] * mm - h_pt

            if isinstance(value, str) and os.path.isfile(value) and value.lower().endswith(('.png', '.jpg', '.jpeg')):
                c.drawImage(value, x_pt, y_pt, width=w_pt, height=h_pt)
            else:
                qr_img = generate_qr_image(value, size_px=max(200, int(cfg["width_mm"] * 10)))
                tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
                qr_img.save(tmp, format="PNG")
                tmp.close()
                c.drawImage(tmp.name, x_pt, y_pt, width=w_pt, height=h_pt)
                os.unlink(tmp.name)

            if cfg.get("show_label", True):
                label_col = cfg.get("label_col", col_name)
                label_offset = cfg.get("label_row_offset", 0)
                target_row_idx = row_idx + label_offset
                if label_col in df_selected.columns and 0 <= target_row_idx < len(df_selected):
                    raw_label = df_selected[label_col].iloc[target_row_idx]
                else:
                    raw_label = "" 
                label_value = smart_str(raw_label)
                font_size = cfg.get("label_font_size", 7)
                x_offset_pt = cfg.get("label_x_offset", 0) * mm
                label_width_mm = max(5.0, float(cfg.get("label_width_mm", cfg["width_mm"])))
                label_x_center = x_pt + w_pt / 2 + x_offset_pt
                align_label = cfg.get("label_align", "กลาง")

                rendered_pdf = False
                if _THAI_FONT_PATH:
                    text_img, tw_mm, th_mm = render_thai_text_pil(
                        label_value, _THAI_FONT_PATH, font_size,
                        label_width_mm, align_label, "#000000")
                    if text_img is not None and tw_mm > 0:
                        tw_pt = tw_mm * mm
                        th_pt = th_mm * mm
                        label_x_left = label_x_center - (label_width_mm * mm) / 2

                        if align_label == "ซ้าย":
                            img_x = label_x_left
                        elif align_label == "ขวา":
                            img_x = label_x_left + label_width_mm * mm - tw_pt
                        else:
                            img_x = label_x_center - tw_pt / 2

                        img_y = y_pt - th_pt - 2  

                        tmp_txt = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
                        text_img.save(tmp_txt, format="PNG")
                        tmp_txt.close()
                        c.drawImage(tmp_txt.name, img_x, img_y, tw_pt, th_pt)
                        os.unlink(tmp_txt.name)
                        rendered_pdf = True

                if not rendered_pdf:
                    label_width_pt = label_width_mm * mm
                    label_x_left = label_x_center - label_width_pt / 2
                    c.setFont(_THAI_FONT_NAME, font_size)
                    lines = wrap_text_for_pdf(label_value, _THAI_FONT_NAME, font_size, label_width_pt)
                    line_spacing = font_size * 1.4
                    label_y = y_pt - font_size - 2
                    for line in lines:
                        if align_label == "ซ้าย":
                            c.drawString(label_x_left, label_y, line)
                        elif align_label == "ขวา":
                            c.drawRightString(label_x_left + label_width_pt, label_y, line)
                        else:
                            c.drawCentredString(label_x_center, label_y, line)
                        label_y -= line_spacing

        if progress_callback:
            progress_callback((row_idx + 1) / total_rows)

    c.save()
    pdf_buf.seek(0)
    return pdf_buf, total_rows

# ══════════════════════════════════════════════
# MAIN APP
# ══════════════════════════════════════════════
def main():
    st.set_page_config(
        page_title="QR Code & Image Generator",
        page_icon="🔲",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    st.markdown(
        """
        <style>
        .stApp { background-color: #ffffff !important; color: #333333 !important; }
        header[data-testid="stHeader"] { background-color: #ffffff !important; }
        section[data-testid="stSidebar"] { background-color: #f7f7f7 !important; color: #333 !important; }
        section[data-testid="stSidebar"] * { color: #333 !important; }
        .stMarkdown, .stText, p, span, label, div { color: #333333 !important; }

        .stDeployButton, #MainMenu, footer, header .stActionButton { display: none !important; }

        .main-title { font-size: 2.2rem; font-weight: 700; color: #1a1a1a !important; margin-bottom: 0; }
        .sub-title  { font-size: 1rem; color: #555 !important; margin-top: 0; }
        .step-header { background: linear-gradient(90deg, #4a6cf7 0%, #6a5acd 100%);
                       color: white !important; padding: 10px 20px; border-radius: 10px;
                       font-size: 1.1rem; font-weight: 600; margin: 20px 0 10px 0; }
        .step-header * { color: white !important; }
        .info-box   { background: #f0f4ff; border-left: 4px solid #4a6cf7;
                      padding: 15px; border-radius: 5px; margin: 10px 0; color: #333 !important; }
        .col-badge  { display: inline-block; padding: 4px 12px; border-radius: 20px;
                      color: white !important; font-weight: 600; font-size: 0.9rem; margin: 2px; }

        .stSlider label, .stCheckbox label, .stSelectbox label,
        .stMultiSelect label, .stNumberInput label, .stTextInput label,
        .stFileUploader label { color: #333 !important; }
        .stMetric label { color: #666 !important; }
        .stMetric [data-testid="stMetricValue"] { color: #1a1a1a !important; }

        .streamlit-expanderHeader { color: #333 !important; }
        </style>
        """,
        unsafe_allow_html=True,
    )

    # เก็บ Folder ชั่วคราวไว้ตลอดอายุ Session
    if "temp_dir" not in st.session_state:
        st.session_state.temp_dir = tempfile.mkdtemp()
    if "img_group_count" not in st.session_state:
        st.session_state.img_group_count = 1

    st.markdown('<p class="main-title">🔲 QR Code & Image PDF Generator</p>', unsafe_allow_html=True)
    st.markdown(
        '<p class="sub-title">นำเข้าข้อมูลจาก Excel (สร้าง QR) หรือ อัปโหลดรูปภาพ → จัดตำแหน่งบนหน้ากระดาษ '
        '→ Export PDF (1 แถว = 1 หน้า)</p>',
        unsafe_allow_html=True,
    )
    st.divider()

    with st.sidebar:
        st.header("📄 ตั้งค่าหน้ากระดาษ")

        page_options = list(PAGE_SIZES.keys()) + ["📐 กำหนดเอง (Custom)"]
        page_size_name = st.selectbox("ขนาดกระดาษ", page_options, index=1)

        if page_size_name == "📐 กำหนดเอง (Custom)":
            ccol1, ccol2 = st.columns(2)
            with ccol1:
                custom_w = st.number_input("กว้าง (mm)", value=210.0, min_value=10.0, max_value=5000.0, step=1.0, format="%.1f")
            with ccol2:
                custom_h = st.number_input("สูง (mm)", value=297.0, min_value=10.0, max_value=5000.0, step=1.0, format="%.1f")
            page_size = (custom_w * mm, custom_h * mm)
            page_w_mm = float(custom_w)
            page_h_mm = float(custom_h)
        else:
            page_size = PAGE_SIZES[page_size_name]
            page_w_mm = round(page_size[0] / mm, 1)
            page_h_mm = round(page_size[1] / mm, 1)

        orientation = st.selectbox("แนวกระดาษ", ["Portrait (แนวตั้ง)", "Landscape (แนวนอน)"])
        is_landscape = "Landscape" in orientation
        if is_landscape:
            page_w_mm, page_h_mm = page_h_mm, page_w_mm

        st.divider()
        st.header("🔲 ค่าเริ่มต้น QR / ภาพ")
        default_qr_size = st.number_input("ขนาดเริ่มต้น (mm)", min_value=3, max_value=500, value=30, step=1)
        default_show_label = st.checkbox("แสดงข้อความใต้ QR/ภาพ", value=False)
        default_label_size = st.number_input("ขนาดตัวอักษร (pt)", min_value=2, max_value=72, value=7, step=1)

    # ─── STEP 1 : IMPORT DATA ───
    st.markdown('<div class="step-header">📥 Step 1 — นำเข้าข้อมูล</div>', unsafe_allow_html=True)

    data_source = st.radio(
        "📌 เลือกแหล่งข้อมูล", 
        ["🖼️ อัปโหลดรูปภาพหลายไฟล์ (PNG/JPG)", "📊 นำเข้าไฟล์ Excel (สร้าง QR Code)"], 
        horizontal=True
    )

    if data_source == "🖼️ อัปโหลดรูปภาพหลายไฟล์ (PNG/JPG)":
        if st.button("➕ เพิ่มกลุ่มรูปภาพ (คอลัมน์ใหม่)", use_container_width=False):
            st.session_state.img_group_count += 1
            st.rerun()

        dict_data = {}
        max_len = 0
        total_uploaded = 0
        
        for g in range(1, st.session_state.img_group_count + 1):
            uploaded_images = st.file_uploader(
                f"📂 อัปโหลดกลุ่มรูปภาพที่ {g} (คอลัมน์ {g})", 
                type=["png", "jpg", "jpeg"], 
                accept_multiple_files=True,
                key=f"uploader_{g}"
            )
            
            paths = []
            if uploaded_images:
                for img_file in uploaded_images:
                    temp_path = os.path.join(st.session_state.temp_dir, f"g{g}_{img_file.name}")
                    with open(temp_path, "wb") as f:
                        f.write(img_file.getbuffer())
                    paths.append(temp_path)
                    total_uploaded += 1
                    
            col_name = f"Image_Group_{g}"
            dict_data[col_name] = paths
            max_len = max(max_len, len(paths))
            
        if total_uploaded == 0:
            st.info("👆 กรุณาอัปโหลดรูปภาพอย่างน้อย 1 กลุ่มเพื่อเริ่มต้น")
            st.stop()
            
        # เติมช่องว่าง (None) ให้ทุกคอลัมน์มีจำนวนแถวเท่ากัน
        for k in dict_data.keys():
            dict_data[k] = dict_data[k] + [None] * (max_len - len(dict_data[k]))
            
        df = pd.DataFrame(dict_data)
        
        st.success(f"✅ อัปโหลดรูปภาพสำเร็จ รวม **{total_uploaded:,}** ไฟล์ (สร้างได้สูงสุด {max_len} หน้า)")
        with st.expander("👀 ดูรายชื่อไฟล์รูปภาพ", expanded=False):
            st.dataframe(df.astype(str), height=300)

    else:
        uploaded_file = st.file_uploader("เลือกไฟล์ Excel (.xlsx, .xls)", type=["xlsx", "xls"])

        if uploaded_file is None:
            st.info("👆 กรุณาอัปโหลดไฟล์ Excel เพื่อเริ่มต้น")
            st.stop()

        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names

        if len(sheet_names) > 1:
            selected_sheet = st.selectbox("เลือก Sheet", sheet_names)
        else:
            selected_sheet = sheet_names[0]

        has_header = st.checkbox("แถวแรกเป็นหัวข้อคอลัมน์ (Header)", value=False)

        if has_header:
            df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
        else:
            df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=None)
            alpha_names = []
            for idx in range(len(df.columns)):
                name = ""
                n = idx
                while True:
                    name = chr(ord('A') + n % 26) + name
                    n = n // 26 - 1
                    if n < 0:
                        break
                alpha_names.append(name)
            df.columns = alpha_names

        df.columns = [str(c) for c in df.columns]

        for col in df.columns:
            if df[col].dtype == "float64":
                non_null = df[col].dropna()
                if len(non_null) > 0 and (non_null == non_null.astype(int)).all():
                    df[col] = df[col].astype("Int64")  

        st.success(f"✅ โหลดสำเร็จ — **{len(df):,}** แถว, **{len(df.columns)}** คอลัมน์  (Sheet: {selected_sheet})")
        with st.expander("👀 ดูข้อมูลทั้งหมด", expanded=False):
            st.dataframe(df.astype(str), height=300)

    # ─── STEP 2 : SELECT COLUMNS & RANGE ───
    st.markdown(
        '<div class="step-header">🎯 Step 2 — เลือกคอลัมน์ (ได้หลายคอลัมน์) และช่วงข้อมูล</div>',
        unsafe_allow_html=True,
    )

    col_options = df.columns.tolist()
    
    # เลือกทุก Image_Group อัตโนมัติถ้ามี
    img_cols = [c for c in col_options if c.startswith("Image_Group_")]
    default_selection = img_cols if img_cols else ([col_options[0]] if col_options else [])
    
    selected_cols = st.multiselect(
        "เลือกคอลัมน์ที่ต้องการสร้าง QR Code หรือแสดงรูปภาพ (เลือกได้หลายคอลัมน์)",
        col_options,
        default=default_selection,
        help="แต่ละคอลัมน์ = ข้อมูล 1 ตัวบนแต่ละหน้า สามารถเลื่อนตำแหน่งแยกกันได้อิสระ",
    )

    if not selected_cols:
        st.warning("⚠️ กรุณาเลือกอย่างน้อย 1 คอลัมน์")
        st.stop()

    badges_html = ""
    for i, col in enumerate(selected_cols):
        color = COLORS[i % len(COLORS)]
        badges_html += f'<span class="col-badge" style="background:{color};">{str(col)}</span> '
    st.markdown(f"คอลัมน์ที่เลือก: {badges_html}", unsafe_allow_html=True)

    total_data_rows = len(df)
    rcol1, rcol2 = st.columns(2)
    with rcol1:
        start_row = st.number_input("เริ่มจากแถวที่", min_value=1, max_value=total_data_rows, value=1)
    with rcol2:
        end_row = st.number_input("ถึงแถวที่", min_value=1, max_value=total_data_rows, value=total_data_rows)

    if start_row > end_row:
        st.error("❌ แถวเริ่มต้นต้องน้อยกว่าหรือเท่ากับแถวสิ้นสุด")
        st.stop()

    df_selected = df.iloc[start_row - 1: end_row].dropna(how="all", subset=selected_cols).reset_index(drop=True)
    total_rows = len(df_selected)

    if total_rows == 0:
        st.warning("⚠️ ไม่มีข้อมูลในช่วงที่เลือก")
        st.stop()

    st.markdown(
        f"""
        <div class="info-box">
            📊 <b>สรุป:</b> เลือก <b>{len(selected_cols)}</b> คอลัมน์ ×
            <b>{total_rows:,}</b> แถว → PDF <b>{total_rows:,} หน้า</b>
            (แต่ละหน้ามีข้อมูล {len(selected_cols)} จุด)
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ─── STEP 3 : POSITION EACH QR/IMAGE ───
    st.markdown(
        '<div class="step-header">📐 Step 3 — จัดตำแหน่ง QR / ภาพ บนหน้ากระดาษ</div>',
        unsafe_allow_html=True,
    )

    max_x = int(page_w_mm)
    max_y = int(page_h_mm)

    if "qr_positions" not in st.session_state:
        st.session_state.qr_positions = {}

    for i, col_name in enumerate(selected_cols):
        cn = str(col_name)
        if cn not in st.session_state.qr_positions:
            # คำนวณ Aspect Ratio เพื่อไม่ให้รูปเบี้ยว
            ratio = 1.0
            valid_vals = df_selected[col_name].dropna()
            if len(valid_vals) > 0:
                first_val = valid_vals.iloc[0]
                ratio = get_image_aspect_ratio(first_val)
                
            st.session_state.qr_positions[cn] = {
                "x": 10,
                "y": max(0, min(10 + i * (default_qr_size + 20), max_y - default_qr_size)),
                "width": default_qr_size,
                "ratio": ratio,
                "label": default_show_label,
                "label_col": cn,
                "label_row_offset": 0,
                "label_x_offset": 0,
                "label_font_size": default_label_size,
                "label_width_mm": default_qr_size,
                "label_align": "กลาง",
            }

    def on_col_select():
        cn = str(st.session_state._active_qr)
        pos = st.session_state.qr_positions.get(cn, {"x": 10, "y": 10, "width": default_qr_size, "ratio": 1.0, "label": True})
        st.session_state._edit_x = pos["x"]
        st.session_state._edit_y = pos["y"]
        st.session_state._edit_width = pos["width"]
        st.session_state._edit_ratio = pos.get("ratio", 1.0)
        st.session_state._edit_label = pos["label"]
        st.session_state._edit_label_x_offset = pos.get("label_x_offset", 0)
        st.session_state._edit_label_font_size = pos.get("label_font_size", default_label_size)
        st.session_state._edit_label_width = pos.get("label_width_mm", pos["width"])
        st.session_state._edit_label_align = pos.get("label_align", "กลาง")
        st.session_state._last_active_cn = cn

    def save_back():
        cn = str(st.session_state._active_qr)
        st.session_state.qr_positions[cn] = {
            "x": st.session_state._edit_x,
            "y": st.session_state._edit_y,
            "width": st.session_state._edit_width,
            "ratio": st.session_state.get("_edit_ratio", 1.0),
            "label": st.session_state._edit_label,
            "label_col": st.session_state.get("_edit_label_col", cn),
            "label_row_offset": st.session_state.get("_edit_label_row_offset", 0),
            "label_x_offset": st.session_state.get("_edit_label_x_offset", 0),
            "label_font_size": st.session_state.get("_edit_label_font_size", default_label_size),
            "label_width_mm": st.session_state.get("_edit_label_width", st.session_state._edit_width),
            "label_align": st.session_state.get("_edit_label_align", "กลาง"),
        }

    first_col = str(selected_cols[0])
    if "_active_qr" not in st.session_state:
        st.session_state._active_qr = first_col
    if st.session_state._active_qr not in [str(c) for c in selected_cols]:
        st.session_state._active_qr = first_col

    active_cn = str(st.session_state._active_qr)
    pos = st.session_state.qr_positions.get(active_cn, {"x": 10, "y": 10, "width": default_qr_size, "ratio": 1.0, "label": True})

    need_init = "_edit_x" not in st.session_state
    col_changed = st.session_state.get("_last_active_cn") != active_cn
    if need_init or col_changed:
        st.session_state._edit_x = pos["x"]
        st.session_state._edit_y = pos["y"]
        st.session_state._edit_width = pos["width"]
        st.session_state._edit_ratio = pos.get("ratio", 1.0)
        st.session_state._edit_label = pos["label"]
        st.session_state._edit_label_col = pos.get("label_col", active_cn)
        st.session_state._edit_label_row_offset = pos.get("label_row_offset", 0)
        st.session_state._edit_label_x_offset = pos.get("label_x_offset", 0)
        st.session_state._edit_label_font_size = pos.get("label_font_size", default_label_size)
        st.session_state._edit_label_width = pos.get("label_width_mm", pos["width"])
        st.session_state._edit_label_align = pos.get("label_align", "กลาง")
        st.session_state._last_active_cn = active_cn

    ctrl_col, preview_col = st.columns([1, 2])

    with ctrl_col:
        col_str_list = [str(c) for c in selected_cols]
        if len(selected_cols) > 1:
            st.selectbox(
                "🔲 เลือกคอลัมน์ที่ต้องการปรับตำแหน่ง",
                col_str_list,
                format_func=lambda c: f"คอลัมน์: {c}",
                key="_active_qr",
                on_change=on_col_select,
            )
        else:
            st.session_state._active_qr = first_col

        active_cn = str(st.session_state._active_qr)
        active_idx = col_str_list.index(active_cn) if active_cn in col_str_list else 0
        active_color = COLORS[active_idx % len(COLORS)]
        
        # หาทีค่าที่มีอยู่จริงเพื่อมาแสดงตัวอย่าง
        valid_sample = df_selected[selected_cols[active_idx]].dropna()
        sample_val = smart_str(valid_sample.iloc[0]) if len(valid_sample) > 0 else ""
        
        if len(sample_val) > 30 and ("/" in sample_val or "\\" in sample_val):
            sample_val_display = "[ไฟล์รูปภาพ] " + os.path.basename(sample_val)
        else:
            sample_val_display = sample_val

        st.markdown(
            f'<div style="border-left: 4px solid {active_color}; padding: 10px 14px; '
            f'margin: 10px 0; background: {active_color}11; border-radius: 0 8px 8px 0;">'
            f'<b style="color:{active_color}; font-size:1.1rem;">🔲 {active_cn}</b>'
            f'<br><span style="font-size:0.85rem; color:#888;">ตัวอย่าง: {sample_val_display}</span></div>',
            unsafe_allow_html=True,
        )

        c1, c2 = st.columns(2)
        with c1:
            st.number_input("↔ X ซ้าย-ขวา (mm)", min_value=0, max_value=max_x,
                            step=1, key="_edit_x", on_change=save_back)
        with c2:
            st.number_input("↕ Y บน-ล่าง (mm)", min_value=0, max_value=max_y,
                            step=1, key="_edit_y", on_change=save_back)

        c3, c4 = st.columns(2)
        with c3:
            st.number_input("↔ ความกว้างรูป/QR (mm)", min_value=3, max_value=500,
                            step=1, key="_edit_width", on_change=save_back)
        with c4:
            # คำนวณและแสดงความสูงอัตโนมัติ
            cur_ratio = st.session_state.get("_edit_ratio", 1.0)
            calc_height = st.session_state._edit_width * cur_ratio
            st.markdown(f"<br><span style='color:#555;'>↕ สูงอัตโนมัติ:<br><b>{calc_height:.1f} mm</b></span>", unsafe_allow_html=True)
            
        st.checkbox("แสดงข้อความด้านล่าง", key="_edit_label", on_change=save_back)

        if st.session_state.get("_edit_label", False):
            p_now = st.session_state.qr_positions.get(active_cn, {})
            
            st.session_state["_edit_label_col"] = p_now.get("label_col", active_cn)
            st.session_state["_edit_label_row_offset"] = int(p_now.get("label_row_offset", 0))
            st.session_state["_edit_label_x_offset"] = int(p_now.get("label_x_offset", 0))
            st.session_state["_edit_label_font_size"] = int(p_now.get("label_font_size", default_label_size))
            st.session_state["_edit_label_width"] = int(p_now.get("label_width_mm", p_now.get("width", default_qr_size)))
            st.session_state["_edit_label_align"] = p_now.get("label_align", "กลาง")
            c_lbl1, c_lbl2 = st.columns(2)
            with c_lbl1:
                st.selectbox("📄 เลือกคอลัมน์ข้อความ", options=df.columns.tolist(), key="_edit_label_col", on_change=save_back)
            with c_lbl2:
                st.number_input("↕ เลื่อนแถวข้อมูล", min_value=-500, max_value=500, step=1, key="_edit_label_row_offset", on_change=save_back, help="1 = เลื่อนลง 1 แถว")

            c5, c6 = st.columns(2)
            with c5:
                st.number_input("↔ ขยับซ้าย-ขวาข้อความ (mm)", min_value=-200, max_value=200,
                                step=1, key="_edit_label_x_offset", on_change=save_back)
            with c6:
                st.number_input("🔤 ขนาดตัวอักษร (pt)", min_value=2, max_value=72,
                                step=1, key="_edit_label_font_size", on_change=save_back)
            c7, c8 = st.columns(2)
            with c7:
                st.number_input("📏 ความกว้างกรอบตัวอักษร (mm)", min_value=5, max_value=max_x,
                                step=1, key="_edit_label_width", on_change=save_back)
            with c8:
                align_options = ["ซ้าย", "กลาง", "ขวา"]
                st.radio("📐 จัดตำแหน่งตัวอักษร", options=align_options,
                         key="_edit_label_align", on_change=save_back, horizontal=True)

        save_back()

        if len(selected_cols) > 1:
            st.markdown("---")
            st.markdown("**📋 ตำแหน่งทั้งหมด:**")
            for j, cn in enumerate(selected_cols):
                cn_str = str(cn)
                clr = COLORS[j % len(COLORS)]
                p = st.session_state.qr_positions.get(cn_str, {})
                marker = " 👈" if cn_str == active_cn else ""
                st.markdown(
                    f'<span style="color:{clr}; font-weight:600;">{cn_str}</span> '
                    f'— X:{p.get("x",0)} Y:{p.get("y",0)} กว้าง:{p.get("width",30)}mm{marker}',
                    unsafe_allow_html=True,
                )

    col_configs = {}
    qr_preview_configs = []
    
    # หาสิ่งที่จะแสดงในหน้า Preview (หน้า 1 หรือหน้าที่ข้อมูลไม่ว่าง)
    for i, col_name in enumerate(selected_cols):
        cn_str = str(col_name)
        color = COLORS[i % len(COLORS)]
        p = st.session_state.qr_positions.get(cn_str, {"x": 10, "y": 10, "width": default_qr_size, "ratio": 1.0, "label": True})
        
        # หา value สำหรับโชว์
        valid_vals = df_selected[col_name].dropna()
        sv = smart_str(valid_vals.iloc[0]) if len(valid_vals) > 0 else ""
        
        label_col = p.get("label_col", cn_str)
        label_offset = p.get("label_row_offset", 0)
        sv_label = ""
        if len(df_selected) > 0 and label_col in df_selected.columns:
            preview_idx = 0 + label_offset
            if 0 <= preview_idx < len(df_selected):
                lbl_val = df_selected[label_col].iloc[preview_idx]
                sv_label = smart_str(lbl_val) if pd.notna(lbl_val) else ""
        else:
            sv_label = sv
            
        w_mm = p["width"]
        h_mm = p["width"] * p.get("ratio", 1.0)
        
        col_configs[col_name] = {
            "x_mm": p["x"],
            "y_mm": p["y"],
            "width_mm": w_mm,
            "height_mm": h_mm,
            "show_label": p["label"],
            "label_col": label_col,
            "label_row_offset": label_offset,
            "label_font_size": p.get("label_font_size", default_label_size),
            "label_x_offset": p.get("label_x_offset", 0),
            "label_width_mm": p.get("label_width_mm", w_mm),
            "label_align": p.get("label_align", "กลาง"),
        }

        qr_preview_configs.append({
            "col_name": cn_str,
            "x_mm": p["x"],
            "y_mm": p["y"],
            "width_mm": w_mm,
            "height_mm": h_mm,
            "value": sv,
            "label_value": sv_label,
            "color": color,
            "show_label": p["label"],
            "label_font_size": p.get("label_font_size", default_label_size),
            "label_x_offset": p.get("label_x_offset", 0),
            "label_width_mm": p.get("label_width_mm", w_mm),
            "label_align": p.get("label_align", "กลาง"),
            "is_active": (cn_str == active_cn),
        })

    with preview_col:
        st.markdown("#### 👁️ Preview")

        fig = create_page_preview(
            page_w_mm, page_h_mm,
            qr_preview_configs,
            total_pages=total_rows,
        )
        st.pyplot(fig)
        plt.close(fig)

        st.caption(
            f"📌 หน้าตัวอย่างจากทั้งหมด {total_rows:,} หน้า — "
            f"ทุกหน้ามีตำแหน่งเดียวกัน ข้อมูลต่างกันตามแถว"
        )

    # ─── STEP 4 : EXPORT ───
    st.markdown('<div class="step-header">📤 Step 4 — Export เป็น PDF</div>', unsafe_allow_html=True)

    ecol1, ecol2, ecol3 = st.columns([2, 1, 1])
    with ecol1:
        pdf_filename = st.text_input("ชื่อไฟล์ PDF", value="output_document")
    with ecol2:
        st.metric("ข้อมูล / หน้า", f"{len(selected_cols)}")
    with ecol3:
        st.metric("จำนวนหน้า PDF", f"{total_rows:,}")

    if st.button("🖨️ สร้าง PDF และดาวน์โหลด", type="primary", use_container_width=True):
        progress_bar = st.progress(0, text="กำลังสร้างไฟล์ PDF...")

        def update_progress(pct):
            done = int(pct * total_rows)
            progress_bar.progress(
                min(pct, 1.0),
                text=f"กำลังสร้าง... หน้า {done}/{total_rows:,} ({pct:.0%})",
            )

        orient_str = "Landscape" if is_landscape else "Portrait"

        pdf_buf, n_pages = generate_pdf(
            df_selected=df_selected,
            col_configs=col_configs,
            page_size=page_size,
            orientation=orient_str,
            progress_callback=update_progress,
        )

        progress_bar.progress(1.0, text="✅ เสร็จสิ้น!")
        st.success(f"✅ สร้าง PDF เรียบร้อย — **{n_pages:,} หน้า**, ข้อมูล {len(selected_cols)} จุด/หน้า")

        st.download_button(
            label="📥 ดาวน์โหลด PDF",
            data=pdf_buf,
            file_name=f"{pdf_filename}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

if __name__ == "__main__":
    main()