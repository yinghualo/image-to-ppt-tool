import streamlit as st
from pptx import Presentation
from pptx.util import Cm, Pt
from PIL import Image
import io

layout_order = [
    "center",
    "top",
    "bottom",
    "left",
    "right"
]

def create_slide(prs, images):
    blank_slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_slide_layout)
    slide_width = prs.slide_width
    slide_height = prs.slide_height

    margin_cm = 1
    margin = Cm(margin_cm)
    px=Pt(2) #間距
    usable_width = slide_width - 2 * margin
    usable_height = slide_height - 2 * margin

    cx = slide_width / 2
    cy = slide_height / 2

    positions = {
        "center": (cx, cy),
        "top": (cx, px+cy - usable_height / 3),
        "bottom": (cx, cy-px + usable_height / 3),
        "left": (cx-px - usable_width / 3, cy),
        "right": (px+cx + usable_width / 3, cy)
    }#用px多一些間距

    max_width = usable_width / 3
    max_height = usable_height / 3

    for i, image_file in enumerate(images):
        if i >= len(layout_order):
            break

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

        pos_x, pos_y = positions[layout_order[i]]
        #print(pos_x, pos_y)
        slide.shapes.add_picture(img_byte_arr, pos_x - final_width / 2, pos_y - final_height / 2, width=final_width, height=final_height)

def main():
    st.set_page_config(page_title="圖片轉 PowerPoint 工具")
    st.title("✨圖片轉 PowerPoint（5圖一頁排版）")
    st.markdown("➡️請上傳圖片，系統會以每 5 張圖一頁，依序排列生成簡報。  \n➡️排列方式為：中間、上方、下方、左側、右側。")

    uploaded = st.file_uploader(
        "拖曳圖片到這裡（圖片大小上限 200MB，每頁 5 張圖）",
        type=["png", "jpg", "jpeg", "bmp", "gif"],
        accept_multiple_files=True
    )

    if uploaded:
        st.session_state.uploaded_files = uploaded

    if "uploaded_files" in st.session_state and st.session_state.uploaded_files:
        # st.markdown("#### 圖片預覽：")
        # for file in st.session_state.uploaded_files:
        #     st.image(file, caption=file.name, use_container_width=True)

        if st.button("🚀 產生PPT"):
            prs = Presentation()
            files = st.session_state.uploaded_files
            for i in range(0, len(files), 5):
                group = files[i:i+5]
                create_slide(prs, group)

            pptx_io = io.BytesIO()
            prs.save(pptx_io)
            pptx_io.seek(0)

            st.success("✅ PPT產生成功！")
            st.download_button("⬇️ 下載PPT", pptx_io, file_name="images_to_ppt.pptx")

if __name__ == "__main__":
    main()
