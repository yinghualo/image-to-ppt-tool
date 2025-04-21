# Image to PowerPoint Tool（圖片轉簡報工具）

這是一個使用 [Streamlit](https://streamlit.io/) 製作的圖片轉簡報小工具，能夠將每 5 張圖片依照特定位置（中間、上方、下方、左側、右側）自動排列成一張投影片，產生整齊美觀的 PowerPoint 檔案，並提供下載。

## 🧰 功能特色

- 支援 JPG、PNG、GIF、BMP、JPEG 格式的圖片上傳
- 每 5 張圖片生成一張投影片，位置分別為：
  - 第 1 張：中間
  - 第 2 張：上方
  - 第 3 張：下方
  - 第 4 張：左方
  - 第 5 張：右方
- 自動圖片縮放，不變形
- 預覽圖片排序即為簡報順序
- 支援多圖上傳，非 5 的倍數也可正常生成簡報
- 線上下載 `.pptx` 簡報檔案

## 🚀 線上體驗

👉 [點我使用線上版工具](https://image-to-ppt-tool-uipazfkmzaurd5ncqf8df9.streamlit.app/)

## 💻 開發與執行方式

### 本地端執行

1. 安裝依賴套件：
```bash
pip install -r requirements.txt
