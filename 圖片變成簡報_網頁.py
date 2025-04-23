import streamlit as st
from pptx import Presentation
from pptx.util import Cm, Pt
from PIL import Image
import io

def create_slide_5(prs, images,inputMargin,inputPadding,quality):
    layout_order = ["center", "top", "bottom", "left", "right"]
    blank_slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_slide_layout)
    slide_width = prs.slide_width
    slide_height = prs.slide_height
    margin = Cm(inputMargin)
    px = Pt(inputPadding)
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
        insert_image(slide, image_file, positions[layout_order[i]], max_width, max_height)

def create_slide_8(prs, images,inputMargin,inputPadding,quality):
    blank_slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_slide_layout)
    slide_width = prs.slide_width
    slide_height = prs.slide_height

    margin = Cm(inputMargin)
    px = Pt(inputPadding)
    usable_width = slide_width - 2 * margin
    usable_height = slide_height - 2 * margin

    cols, rows = 3, 3
    cell_width = usable_width / cols
    cell_height = usable_height / rows

    max_width = cell_width - px
    max_height = cell_height - px

    # 特殊排列順序：index 代表位置，值代表第幾張圖片（1-based）
    layout_index = [3, 8, 4,
                    5, 1, 6,
                    2, 7, None]

    for pos, img_num in enumerate(layout_index):
        if img_num is None or img_num > len(images):
            continue

        row = pos // cols
        col = pos % cols
        pos_x = margin + col * cell_width + cell_width / 2
        pos_y = margin + row * cell_height + cell_height / 2

        img = Image.open(images[img_num - 1])
        img_ratio = img.width / img.height
        box_ratio = max_width / max_height

        if img_ratio > box_ratio:
            final_width = max_width
            final_height = max_width / img_ratio
        else:
            final_height = max_height
            final_width = max_height * img_ratio

        img_byte_arr = io.BytesIO()
        #改存成JPEG並調整壓縮品質以控制簡報大小
        #img.save(img_byte_arr, format='PNG')
        img.convert("RGB").save(img_byte_arr, format='JPEG', quality=quality)       
        img_byte_arr.seek(0)

        slide.shapes.add_picture(
            img_byte_arr,
            pos_x - final_width / 2,
            pos_y - final_height / 2,
            width=final_width,
            height=final_height
        )

def insert_image(slide, image_file, center_pos, max_width, max_height):
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
    img.save(img_byte_arr, format='PNG')
    img_byte_arr.seek(0)
    pos_x, pos_y = center_pos
    slide.shapes.add_picture(
        img_byte_arr,
        pos_x - final_width / 2,
        pos_y - final_height / 2,
        width=final_width,
        height=final_height
    )

def main():
    st.set_page_config(page_title="圖片轉 PowerPoint 工具")
    st.title("🖼️ 圖片轉 PowerPoint 工具")
    layout_mode = st.selectbox("選擇排版模式", ["8圖一頁（九宮格依序為38451627）", "5圖一頁（中心、上、下、左、右）"])
    layout_margin=st.number_input("簡報邊界(cm)",1)#預設為1公分
    layout_padding=st.number_input("圖片間距(px)",2)#預設為2px
    uploaded = st.file_uploader(
        "拖曳圖片上傳（最多200MB）",
        type=["png", "jpg", "jpeg", "bmp", "gif"],
        accept_multiple_files=True
    )

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
                    create_slide_8(prs, files[i:i+8],layout_margin,layout_padding,quality)
            elif layout_mode.startswith("5圖"):
                for i in range(0, len(files), 5):
                    create_slide_5(prs, files[i:i+5],layout_margin,layout_padding,quality)

            pptx_io = io.BytesIO()
            prs.save(pptx_io)
            pptx_io.seek(0)

            st.success("✅ PPT產生成功！")
            st.download_button("📥 下載PPT", pptx_io, file_name="images_to_ppt.pptx")
            st.session_state.clear()  # ⬅️ 自動清除 session 狀態

if __name__ == "__main__":
    main()
