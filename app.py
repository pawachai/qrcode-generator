import streamlit as st
import pandas as pd
import qrcode
from io import BytesIO
from PIL import Image, ImageDraw, ImageFont
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.offsetbox import OffsetImage, AnnotationBbox
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, A3, A5, A6, B4, B5, letter, legal, landscape
from reportlab.lib.units import mm
import tempfile
import os
import math
import numpy as np

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


def smart_str(val) -> str:
    """Convert value to string, removing .0 from whole numbers."""
    if pd.isna(val):
        return ""
    if isinstance(val, float) and val == int(val):
        return str(int(val))
    return str(val)


def generate_qr_image(data: str, size_px: int = 300) -> Image.Image:
    """Generate a QR code as a PIL Image."""
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


def create_page_preview(
    page_w_mm, page_h_mm,
    qr_configs,
    total_pages,
):
    """Create a matplotlib preview showing 1 page with QR codes positioned."""
    fig_h = 8
    fig_w = fig_h * page_w_mm / page_h_mm
    fig, ax = plt.subplots(1, 1, figsize=(max(4, fig_w), fig_h), dpi=100)

    # Draw stacked pages behind (shadow effect)
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

    # Front page
    page_rect = patches.FancyBboxPatch(
        (0, 0), page_w_mm, page_h_mm,
        boxstyle="round,pad=0",
        linewidth=2, edgecolor="#333", facecolor="white",
    )
    ax.add_patch(page_rect)

    # Draw each QR code
    for cfg in qr_configs:
        x = cfg["x_mm"]
        y = cfg["y_mm"]
        size = cfg["size_mm"]
        color = cfg["color"]
        col_name = cfg["col_name"]
        value = smart_str(cfg["value"])

        # QR code border — thicker if active/selected
        is_active = cfg.get("is_active", False)
        border_w = 3.5 if is_active else 1.5
        qr_rect = patches.FancyBboxPatch(
            (x, y), size, size,
            boxstyle="round,pad=0.3",
            linewidth=border_w, edgecolor=color, facecolor="#fafafa",
        )
        ax.add_patch(qr_rect)

        # Draw actual QR image — fill the entire box
        try:
            qr_img = generate_qr_image(value, size_px=200)
            qr_arr = np.array(qr_img)
            ax.imshow(
                qr_arr,
                extent=[x + 0.5, x + size - 0.5, y + size - 0.5, y + 0.5],
                aspect="auto", zorder=5, interpolation="nearest",
            )
        except Exception:
            ax.text(x + size / 2, y + size / 2, "QR", ha="center", va="center",
                    fontsize=max(6, size * 0.3), color="#333", weight="bold")

        # Column name badge
        badge_text = str(col_name)
        badge_w = max(10, len(badge_text) * 2.2 + 4)
        badge_h = 3.5
        badge_x = x + (size - badge_w) / 2
        badge_y = y - badge_h - 1.5
        if badge_y < -5:
            badge_y = y + size + 1.5

        badge = patches.FancyBboxPatch(
            (badge_x, badge_y), badge_w, badge_h,
            boxstyle="round,pad=0.5",
            linewidth=0, facecolor=color, alpha=0.9, zorder=6,
        )
        ax.add_patch(badge)
        ax.text(badge_x + badge_w / 2, badge_y + badge_h / 2, badge_text,
                ha="center", va="center", fontsize=5.5, color="white", weight="bold", zorder=7)

        # Value text below QR
        if cfg.get("show_label", True):
            label_y_pos = y + size + 2
            if badge_y > y:
                label_y_pos = badge_y + badge_h + 1
            display = value if len(value) <= 28 else value[:25] + "..."
            ax.text(x + size / 2, label_y_pos, display,
                    ha="center", va="top", fontsize=5, color="#555", zorder=7)

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
    """Generate PDF: 1 page per row, each page has QR codes for each column."""
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
            if pd.isna(raw):
                continue
            value = smart_str(raw)

            qr_img = generate_qr_image(value, size_px=max(200, int(cfg["size_mm"] * 10)))

            tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
            qr_img.save(tmp, format="PNG")
            tmp.close()

            x_pt = cfg["x_mm"] * mm
            y_pt = page_h - cfg["y_mm"] * mm - cfg["size_mm"] * mm

            c.drawImage(tmp.name, x_pt, y_pt, cfg["size_mm"] * mm, cfg["size_mm"] * mm)
            os.unlink(tmp.name)

            if cfg.get("show_label", True):
                label_x = x_pt + (cfg["size_mm"] * mm) / 2
                label_y = y_pt - cfg.get("label_font_size", 7) - 2
                c.setFont("Helvetica", cfg.get("label_font_size", 7))
                display = value if len(value) <= 40 else value[:37] + "..."
                c.drawCentredString(label_x, label_y, display)

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
        page_title="QR Code Generator",
        page_icon="🔲",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    st.markdown(
        """
        <style>
        /* ── Force light / white theme ── */
        .stApp { background-color: #ffffff !important; color: #333333 !important; }
        header[data-testid="stHeader"] { background-color: #ffffff !important; }
        section[data-testid="stSidebar"] { background-color: #f7f7f7 !important; color: #333 !important; }
        section[data-testid="stSidebar"] * { color: #333 !important; }
        .stMarkdown, .stText, p, span, label, div { color: #333333 !important; }

        /* Hide Deploy button & Streamlit menu */
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

        /* Inputs & widgets */
        .stSlider label, .stCheckbox label, .stSelectbox label,
        .stMultiSelect label, .stNumberInput label, .stTextInput label,
        .stFileUploader label { color: #333 !important; }
        .stMetric label { color: #666 !important; }
        .stMetric [data-testid="stMetricValue"] { color: #1a1a1a !important; }

        /* Expander */
        .streamlit-expanderHeader { color: #333 !important; }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown('<p class="main-title">🔲 QR Code Generator from Excel</p>', unsafe_allow_html=True)
    st.markdown(
        '<p class="sub-title">นำเข้า Excel → เลือกคอลัมน์ → จัดตำแหน่ง QR อิสระบนหน้ากระดาษ '
        '→ Export PDF (1 แถว = 1 หน้า)</p>',
        unsafe_allow_html=True,
    )
    st.divider()

    # ─── SIDEBAR : PAGE SETTINGS ───
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
        st.header("🔲 ค่าเริ่มต้น QR Code")
        default_qr_size = st.number_input("ขนาดเริ่มต้น QR (mm)", min_value=3, max_value=500, value=30, step=1)
        default_show_label = st.checkbox("แสดงข้อความใต้ QR", value=True)
        default_label_size = st.number_input("ขนาดตัวอักษร (pt)", min_value=2, max_value=72, value=7, step=1)

    # ─── STEP 1 : IMPORT EXCEL ───
    st.markdown('<div class="step-header">📥 Step 1 — นำเข้าไฟล์ Excel</div>', unsafe_allow_html=True)

    uploaded_file = st.file_uploader(
        "เลือกไฟล์ Excel (.xlsx, .xls)",
        type=["xlsx", "xls"],
    )

    if uploaded_file is None:
        st.info("👆 กรุณาอัปโหลดไฟล์ Excel เพื่อเริ่มต้น")
        st.stop()

    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names

    if len(sheet_names) > 1:
        selected_sheet = st.selectbox("เลือก Sheet", sheet_names)
    else:
        selected_sheet = sheet_names[0]

    has_header = st.checkbox(
        "แถวแรกเป็นหัวข้อคอลัมน์ (Header)",
        value=False,
        help="ถ้า Excel ไม่มีหัวข้อคอลัมน์ ให้ปิดตัวเลือกนี้ → แถวแรกจะเป็นข้อมูล",
    )

    if has_header:
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
    else:
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=None)
        # Auto-generate column names: A, B, C, ...
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

    # Ensure all column names are strings
    df.columns = [str(c) for c in df.columns]

    # Clean float columns: 1.0 → 1 if all values are whole numbers
    for col in df.columns:
        if df[col].dtype == "float64":
            non_null = df[col].dropna()
            if len(non_null) > 0 and (non_null == non_null.astype(int)).all():
                df[col] = df[col].astype("Int64")  # nullable int

    st.success(f"✅ โหลดสำเร็จ — **{len(df):,}** แถว, **{len(df.columns)}** คอลัมน์  (Sheet: {selected_sheet})")

    with st.expander("👀 ดูข้อมูลทั้งหมด", expanded=False):
        st.dataframe(df, height=300)

    # ─── STEP 2 : SELECT COLUMNS & RANGE ───
    st.markdown(
        '<div class="step-header">🎯 Step 2 — เลือกคอลัมน์ (ได้หลายคอลัมน์) และช่วงข้อมูล</div>',
        unsafe_allow_html=True,
    )

    col_options = df.columns.tolist()
    selected_cols = st.multiselect(
        "เลือกคอลัมน์ที่ต้องการสร้าง QR Code (เลือกได้หลายคอลัมน์)",
        col_options,
        default=[col_options[0]] if col_options else [],
        help="แต่ละคอลัมน์ = QR 1 ตัวบนแต่ละหน้า สามารถเลื่อนตำแหน่งแยกกันได้อิสระ",
    )

    if not selected_cols:
        st.warning("⚠️ กรุณาเลือกอย่างน้อย 1 คอลัมน์")
        st.stop()

    # Colored badges
    badges_html = ""
    for i, col in enumerate(selected_cols):
        color = COLORS[i % len(COLORS)]
        badges_html += f'<span class="col-badge" style="background:{color};">{str(col)}</span> '
    st.markdown(f"คอลัมน์ที่เลือก: {badges_html}", unsafe_allow_html=True)

    # Row range
    total_data_rows = len(df)
    rcol1, rcol2 = st.columns(2)
    with rcol1:
        start_row = st.number_input("เริ่มจากแถวที่", min_value=1, max_value=total_data_rows, value=1)
    with rcol2:
        end_row = st.number_input("ถึงแถวที่", min_value=1, max_value=total_data_rows, value=total_data_rows)

    if start_row > end_row:
        st.error("❌ แถวเริ่มต้นต้องน้อยกว่าหรือเท่ากับแถวสิ้นสุด")
        st.stop()

    df_selected = df.iloc[start_row - 1: end_row][selected_cols].dropna(how="all").reset_index(drop=True)
    total_rows = len(df_selected)

    if total_rows == 0:
        st.warning("⚠️ ไม่มีข้อมูลในช่วงที่เลือก")
        st.stop()

    st.markdown(
        f"""
        <div class="info-box">
            📊 <b>สรุป:</b> เลือก <b>{len(selected_cols)}</b> คอลัมน์ ×
            <b>{total_rows:,}</b> แถว → PDF <b>{total_rows:,} หน้า</b>
            (แต่ละหน้ามี QR {len(selected_cols)} ตัว)
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ─── STEP 3 : POSITION EACH QR ───
    st.markdown(
        '<div class="step-header">📐 Step 3 — จัดตำแหน่ง QR Code บนหน้ากระดาษ</div>',
        unsafe_allow_html=True,
    )

    max_x = int(page_w_mm)
    max_y = int(page_h_mm)

    # ── Store all positions in a dict inside session_state ──
    if "qr_positions" not in st.session_state:
        st.session_state.qr_positions = {}

    # Initialize defaults for any new columns
    for i, col_name in enumerate(selected_cols):
        cn = str(col_name)
        if cn not in st.session_state.qr_positions:
            st.session_state.qr_positions[cn] = {
                "x": 10,
                "y": max(0, min(10 + i * (default_qr_size + 20), max_y - default_qr_size)),
                "size": default_qr_size,
                "label": default_show_label,
            }

    # ── Callbacks ──
    def on_col_select():
        """When user picks a different QR from dropdown, load its values into edit widgets."""
        cn = str(st.session_state._active_qr)
        pos = st.session_state.qr_positions.get(cn, {"x": 10, "y": 10, "size": default_qr_size, "label": True})
        st.session_state._edit_x = pos["x"]
        st.session_state._edit_y = pos["y"]
        st.session_state._edit_size = pos["size"]
        st.session_state._edit_label = pos["label"]

    def save_back():
        """Save current edit widget values back to the active column's position."""
        cn = str(st.session_state._active_qr)
        st.session_state.qr_positions[cn] = {
            "x": st.session_state._edit_x,
            "y": st.session_state._edit_y,
            "size": st.session_state._edit_size,
            "label": st.session_state._edit_label,
        }

    # ── Init edit widget defaults (first column) ──
    first_col = str(selected_cols[0])
    if "_active_qr" not in st.session_state:
        st.session_state._active_qr = first_col
    # Make sure active is still valid
    if st.session_state._active_qr not in [str(c) for c in selected_cols]:
        st.session_state._active_qr = first_col

    active_cn = str(st.session_state._active_qr)
    pos = st.session_state.qr_positions.get(active_cn, {"x": 10, "y": 10, "size": default_qr_size, "label": True})

    if "_edit_x" not in st.session_state:
        st.session_state._edit_x = pos["x"]
        st.session_state._edit_y = pos["y"]
        st.session_state._edit_size = pos["size"]
        st.session_state._edit_label = pos["label"]

    ctrl_col, preview_col = st.columns([1, 2])

    with ctrl_col:
        # Dropdown to pick which QR to edit
        col_str_list = [str(c) for c in selected_cols]
        if len(selected_cols) > 1:
            st.selectbox(
                "🔲 เลือก QR ที่ต้องการปรับตำแหน่ง",
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
        sample_val = smart_str(df_selected[selected_cols[active_idx]].iloc[0]) if len(df_selected) > 0 else ""

        st.markdown(
            f'<div style="border-left: 4px solid {active_color}; padding: 10px 14px; '
            f'margin: 10px 0; background: {active_color}11; border-radius: 0 8px 8px 0;">'
            f'<b style="color:{active_color}; font-size:1.1rem;">🔲 {active_cn}</b>'
            f'<br><span style="font-size:0.85rem; color:#888;">ตัวอย่างค่า: {sample_val}</span></div>',
            unsafe_allow_html=True,
        )

        # Edit widgets — fixed keys, save on change
        c1, c2 = st.columns(2)
        with c1:
            st.number_input("↔ X ซ้าย-ขวา (mm)", min_value=0, max_value=max_x,
                            step=1, key="_edit_x", on_change=save_back)
        with c2:
            st.number_input("↕ Y บน-ล่าง (mm)", min_value=0, max_value=max_y,
                            step=1, key="_edit_y", on_change=save_back)

        c3, c4 = st.columns(2)
        with c3:
            st.number_input("ขนาด QR (mm)", min_value=3, max_value=500,
                            step=1, key="_edit_size", on_change=save_back)
        with c4:
            st.checkbox("แสดงข้อความใต้ QR", key="_edit_label", on_change=save_back)

        # Also save on every render (in case user just typed)
        save_back()

        # Summary of all QR positions
        if len(selected_cols) > 1:
            st.markdown("---")
            st.markdown("**📋 ตำแหน่ง QR ทั้งหมด:**")
            for j, cn in enumerate(selected_cols):
                cn_str = str(cn)
                clr = COLORS[j % len(COLORS)]
                p = st.session_state.qr_positions.get(cn_str, {})
                marker = " 👈" if cn_str == active_cn else ""
                st.markdown(
                    f'<span style="color:{clr}; font-weight:600;">{cn_str}</span> '
                    f'— X:{p.get("x",0)} Y:{p.get("y",0)} ขนาด:{p.get("size",30)}mm{marker}',
                    unsafe_allow_html=True,
                )

    # ── Build configs from stored positions ──
    col_configs = {}
    qr_preview_configs = []
    for i, col_name in enumerate(selected_cols):
        cn_str = str(col_name)
        color = COLORS[i % len(COLORS)]
        p = st.session_state.qr_positions.get(cn_str, {"x": 10, "y": 10, "size": default_qr_size, "label": True})
        sv = smart_str(df_selected[col_name].iloc[0]) if len(df_selected) > 0 else ""

        col_configs[col_name] = {
            "x_mm": p["x"],
            "y_mm": p["y"],
            "size_mm": p["size"],
            "show_label": p["label"],
            "label_font_size": default_label_size,
        }

        qr_preview_configs.append({
            "col_name": cn_str,
            "x_mm": p["x"],
            "y_mm": p["y"],
            "size_mm": p["size"],
            "value": sv,
            "color": color,
            "show_label": p["label"],
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
            f"📌 หน้าที่ 1 จากทั้งหมด {total_rows:,} หน้า — "
            f"ทุกหน้ามี QR ตำแหน่งเดียวกัน ข้อมูลต่างกันตามแถว"
        )

    # ─── STEP 4 : EXPORT ───
    st.markdown('<div class="step-header">📤 Step 4 — Export เป็น PDF</div>', unsafe_allow_html=True)

    ecol1, ecol2, ecol3 = st.columns([2, 1, 1])
    with ecol1:
        pdf_filename = st.text_input("ชื่อไฟล์ PDF", value="qrcodes_output")
    with ecol2:
        st.metric("QR / หน้า", f"{len(selected_cols)}")
    with ecol3:
        st.metric("จำนวนหน้า PDF", f"{total_rows:,}")

    if st.button("🖨️ สร้าง PDF และดาวน์โหลด", type="primary", use_container_width=True):
        progress_bar = st.progress(0, text="กำลังสร้าง QR Code...")

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
        st.success(f"✅ สร้าง PDF เรียบร้อย — **{n_pages:,} หน้า**, QR {len(selected_cols)} ตัว/หน้า")

        st.download_button(
            label="📥 ดาวน์โหลด PDF",
            data=pdf_buf,
            file_name=f"{pdf_filename}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )


if __name__ == "__main__":
    main()
