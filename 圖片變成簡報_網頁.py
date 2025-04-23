import streamlit as st
from pptx import Presentation
from pptx.util import Cm, Pt
from PIL import Image
import io

# 九宮格排列順序設定：3、8、4、5、1、6、2、7
nine_grid_order = [2, 7, 3, 4, 0, 5, 1, 6]


def insert_image(slide, image_file, center_pos, max_width, max_height, quality):
    img = Image.open(image_file)
    img_ratio = img.width / img.height
    box_ratio = max_width / max_height
    if img_ratio > box_ratio:
        final_width = max_width
        final_height = max_width / img_ratio
    else:
        final_height = max_height
        final_width = max_height * img_ratio
    img_byte_arr = io.BytesIO()
    img.convert("RGB").save(img_byte_arr, format='JPEG', quality=quality)
    img_byte_arr.seek(0)
    pos_x, pos_y = center_pos
    slide.shapes.add_picture(
        img_byte_arr,
        pos_x - final_width / 2,
        pos_y - final_height / 2,
        width=final_width,
        height=final_height
    )


def create_slide_1(prs, images, margin_cm, padding_px, quality):
    for image_file in images:
        blank_slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_slide_layout)
        slide_width = prs.slide_width
        slide_height = prs.slide_height
        margin = Cm(margin_cm)
        px = Pt(padding_px)
        usable_width = slide_width - 2 * margin
        usable_height = slide_height - 2 * margin
        center = (slide_width / 2, slide_height / 2)
        max_width = usable_width
        max_height = usable_height
        insert_image(slide, image_file, center, max_width, max_height, quality)


def create_slide_5(prs, images, margin_cm, padding_px, quality):
    layout_order = ["center", "top", "bottom", "left", "right"]
    blank_slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_slide_layout)
    slide_width = prs.slide_width
    slide_height = prs.slide_height
    margin = Cm(margin_cm)
    px = Pt(padding_px)
    usable_width = slide_width - 2 * margin
    usable_height = slide_height - 2 * margin
    cx, cy = slide_width / 2, slide_height / 2
    positions = {
        "center": (cx, cy),
        "top": (cx, px + cy - usable_height / 3),
        "bottom": (cx, cy - px + usable_height / 3),
        "left": (cx - px - usable_width / 3, cy),
        "right": (px + cx + usable_width / 3, cy)
    }
    max_width = usable_width / 3
    max_height = usable_height / 3
    for i, image_file in enumerate(images):
        if i >= len(layout_order):
            break
        insert_image(slide, image_file, positions[layout_order[i]], max_width, max_height, quality)


def create_slide_8(prs, images, margin_cm, padding_px, quality):
    blank_slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_slide_layout)
    slide_width = prs.slide_width
    slide_height = prs.slide_height
    margin = Cm(margin_cm)
    px = Pt(padding_px)
    usable_width = slide_width - 2 * margin
    usable_height = slide_height - 2 * margin
    cols, rows = 3, 3
    cell_width = usable_width / cols
    cell_height = usable_height / rows
    max_width = cell_width - px
    max_height = cell_height - px
    for idx_in_order, img_idx in enumerate(nine_grid_order):
        if img_idx >= len(images):
            break
        row = idx_in_order // cols
        col = idx_in_order % cols
        pos_x = margin + col * cell_width + cell_width / 2
        pos_y = margin + row * cell_height + cell_height / 2
        insert_image(slide, images[img_idx], (pos_x, pos_y), max_width, max_height, quality)


def main():
    st.set_page_config(page_title="圖片轉 PowerPoint 工具")
    st.title("🖼️ 圖片轉 PowerPoint 工具")

    layout_mode = st.selectbox("選擇排版模式", [
        "1圖一頁（滿版）",
        "5圖一頁（中心、上、下、左、右）",
        "8圖一頁（九宮格依序為38451627）"
    ])
    layout_margin = st.number_input("簡報邊界(cm)", 1)
    layout_padding = st.number_input("圖片間距(px)", 2)
    uploaded = st.file_uploader("拖曳圖片上傳（最多200MB）", type=["png", "jpg", "jpeg", "bmp", "gif"], accept_multiple_files=True)

    compression_option = st.selectbox("圖片壓縮設定：", [
        "建議（輕壓縮）85%",
        "原圖（無壓縮）100%",
        "小檔（高壓縮）65%"
    ])
    quality_dict = {
        "原圖（無壓縮）100%": 100,
        "建議（輕壓縮）85%": 85,
        "小檔（高壓縮）65%": 65
    }
    quality = quality_dict[compression_option]

    if uploaded:
        st.session_state.uploaded_files = uploaded

    if "uploaded_files" in st.session_state and st.session_state.uploaded_files:
        if st.button("🚀 產生PPT"):
            prs = Presentation()
            files = st.session_state.uploaded_files

            if layout_mode.startswith("8圖"):
                for i in range(0, len(files), 8):
                    create_slide_8(prs, files[i:i+8], layout_margin, layout_padding, quality)
            elif layout_mode.startswith("5圖"):
                for i in range(0, len(files), 5):
                    create_slide_5(prs, files[i:i+5], layout_margin, layout_padding, quality)
            elif layout_mode.startswith("1圖"):
                create_slide_1(prs, files, layout_margin, layout_padding, quality)

            pptx_io = io.BytesIO()
            prs.save(pptx_io)
            pptx_io.seek(0)

            st.success("✅ PPT產生成功！")
            st.download_button("📥 下載PPT", pptx_io, file_name="images_to_ppt.pptx")
            st.session_state.clear()


if __name__ == "__main__":
    main()