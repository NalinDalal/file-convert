to run the app:

> python3 -m venv venv
> source venv/bin/activate
> pip install fpdf pillow moviepy python-docx PyPDF2 openpyxl pandas pdf2image python-pptx ebooklib
> brew install poppler # macOS (needed for pdf2image)
> python3 app.py

## ✅ **Supported File Conversions**

works for only:

| **Input Format** | **Target Format** | **What It Does**                           |
| ---------------- | ----------------- | ------------------------------------------ |
| `.docx`          | `.pdf`            | Converts Word documents to PDF (text only) |
| `.pdf`           | `.docx`           | Extracts PDF text and saves into Word file |
| `.jpg`, `.png`   | `.pdf`            | Converts images into PDF pages             |
| `.mp4`           | `.mp3`            | Extracts audio from video and saves as MP3 |
| `.wav`           | `.mp3`            | Converts WAV audio to MP3                  |
| `.txt`           | `.pdf`            | Converts plain text files into PDF         |

- `.mp3 → .wav`, `.pdf → .png`
- `.docx → .txt`, `.pdf → .jpg`

- `.xlsx → .csv` or vice versa
- `.pptx → .pdf`
- `.pdf → .png` (page-to-image using `pdf2image`)
- `.epub → .pdf`
