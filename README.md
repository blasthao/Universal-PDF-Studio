# Universal PDF Studio (Multifunctional PDF Smart Editing & Processing Software V1.0)

[![Python Version](https://img.shields.io/badge/Python-3.8%2B-blue.svg)](https://www.python.org/)
[![Streamlit](https://img.shields.io/badge/Streamlit-1.35%2B-red.svg)](https://streamlit.io/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

A comprehensive, high-performance web-based PDF toolbox built with **Streamlit**, **PyMuPDF**, and advanced **cryptography**. Designed for seamless cross-format document conversion, intelligent visual page manipulation, ROI cropping, document security, and legal **PKCS#7 digital signatures**.

---

## 🌟 Key Features

1. **🖼️ Image to PDF**
   * Batch convert popular image formats (`JPG`, `JPEG`, `PNG`, `TIFF`, `WEBP`, `BMP`) into a single or multiple PDF documents.
2. **🔄 PDF to File Conversion**
   * Export PDF pages into high-fidelity Word (`.docx`), Excel (`.xlsx` tables), PowerPoint (`.pptx`), HTML, and image packages (`.zip`).
3. **🔀 Intelligent PDF Merging**
   * Visual thumbnail preview for each uploaded file, custom sequence reordering, and pre-rotation (`90°`, `180°`, `270°`) before merging.
4. **✂️ PDF Splitting & ROI Cropping**
   * Split documents by single pages, extract custom page ranges, or use interactive grid multi-selection.
   * **ROI Page Cropping**: Manually adjust cropping regions via intuitive slider handles or auto-trim white margins with side-by-side visual comparison.
5. **🛠️ Advanced Editing & Security**
   * **Compression**: Deep Deflate compression to optimize object streams and reduce file size.
   * **Watermarking & Page Numbers**: Dynamic angle watermarking and custom positioning for page numbers.
   * **Encryption / Decryption**: Lock PDFs with passwords or completely strip password protection in one click.
6. **🛡️ Dual-Mode Signature Center**
   * **Simple Signature**: Overlay transparent image seals/stamps or custom Chinese/English text signatures with color and font-size customization.
   * **PKCS#7 Digital Signatures**: Generate real **RSA-2048 / X.509** digital certificates compliant with eIDAS, ESIGN, and UETA standards. Embeds cryptographic hashes into the PDF structure so that signature details and timestamps are directly viewable in **Windows File Properties -> Digital Signatures**.
7. **🌐 Internationalization & UI Customization**
   * Full multilingual support for **Simplified Chinese**, **English**, and **Japanese**.
   * Adjustable UI font size slider and auto-adaptive dark/light theme support.

---

## 🛠️ Tech Stack & Dependencies

* **Web Framework**: [Streamlit](https://streamlit.io/)
* **PDF Engine**: [PyMuPDF (fitz)](https://github.com/pymupdf/PyMuPDF), [pypdf](https://pypdf.readthedocs.io/)
* **Document Processing**: `pdf2docx`, `python-docx`, `python-pptx`, `openpyxl`, `pdfplumber`, `pandas`, `Pillow`
* **Security & Cryptography**: `cryptography` (RSA & PKCS#7 support)

---

## 🚀 Quick Start

### 1. Clone the Repository
```bash
git clone [https://github.com/your-username/universal-pdf-studio.git](https://github.com/your-username/universal-pdf-studio.git)
cd universal-pdf-studio

```

### 2. Install Dependencies

Ensure you have Python 3.8+ installed, then run:

```bash
pip install -r requirements.txt

```

### 3. Run the Application

```bash
streamlit run app.py

```

Open your browser and navigate to `http://localhost:8501`.

---

## 📦 Standalone Desktop App Packaging (PyInstaller)

If you want to package this project into a standalone Windows `.exe` application:

1. Place your custom application icon as `app.ico` in the root directory.
2. Run the provided build script:
```bash
python build.py
```


3. Find your standalone application package in the `dist/` folder.

---

## ❤️ Acknowledgments & Tribute

* Inspired by and tribute to **[I❤PDF](https://www.ilovepdf.com)** for pioneering accessible, efficient online document utilities worldwide.


## 📄 License

This project is licensed under the [MIT License](https://www.google.com/search?q=LICENSE).
