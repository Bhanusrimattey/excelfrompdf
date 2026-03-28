# excelfrompdf

> Convert PDF tables to Excel — automatically, accurately, and for free.

**excelfrompdf** is an open-source Python/FastAPI web service that extracts tabular data from PDF files and converts it to structured `.xlsx` spreadsheets. It intelligently detects whether a PDF is native (text-based) or scanned, and routes it through the appropriate extraction pipeline automatically.

---



## Features

- **Smart PDF routing** — detects native vs. scanned PDFs and picks the right pipeline automatically
- **Native PDF pipeline** — uses geometric line/edge detection for high-accuracy table extraction from machine-generated PDFs
- **Scanned PDF pipeline** — OCR-powered extraction for image-based or scanned documents
- **Text based table extraction** — Text based table extraction for borderless tables
- **Auto-correction** — per-page orientation and skew correction before extraction
- **Excel output** — clean `.xlsx` output preserving table structure

---

## Project Structure

```
excelfrompdf/
├── connection.py          # Entry point: PDF routing logic
├── core/
│   ├── pdfnormal.py       # Native PDF table extraction (line/edge-based)
│   ├── pdfscanned.py      # Scanned PDF extraction (OCR-based)
│   ├── autocorrect.py     # Per-page PDF orientation correction
│   ├── pdfhelper.py       # Helper for scanned PDF table extraction (line/edge-based)
│   └── pdftext.py         # Native/scanned table extraction (text-based)
├── .env.example           # Environment variable template
├── requirements.txt       # Python dependencies
└── README.md
```

---

## How It Works

`connection.py` is the routing brain of the project:

```
Input PDF
    │
    ├─ Extract text (pdfminer)
    ├─ Extract horizontal & vertical edges
    │
    ├─ Auto-correct orientation (autocorrect)
    │
    ├─ [text present AND sufficient edges?]
    │       YES → pdfnormal pipeline  (native PDF)
    │       NO  → pdfscanned pipeline (scanned/image PDF)
    │
    └─ Output: final.xlsx
```

The routing condition checks both the presence of extractable text and a minimum number of detected table edges (`h_edges > 2` and `v_edges > 2`) to confirm a native, structured PDF before using the faster line-based pipeline.

---

## Quickstart

### 1. Clone the repository

```bash
git clone https://github.com/yourusername/excelfrompdf.git
cd excelfrompdf
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Configure environment

```bash
cp .env.example .env
# Edit .env with your credentials (Google Cloud, etc.)
```

### 4. Run a conversion

```bash
python connection.py
```

By default this converts `test.pdf` → `final.xlsx`. Edit `connection.py` to pass your own file path.

---

## Requirements

- Python 3.10.8
- Dependencies listed in `requirements.txt`
- `pytesseract` + Tesseract OCR installed on your system

Install Python dependencies all at once:

```bash
pip install -r requirements.txt
```

---

## Configuration

Copy `.env.example` to `.env` and fill in the values you need:

```env
# TESSERACT_PATH (required on Windows)
TESSERACT_PATH=path\Tesseract-OCR\tesseract.exe

# Google Cloud Vision (optional — improves scanned PDF accuracy)
# If needed, you can add google vision api for improved accuracy
GOOGLE_APPLICATION_CREDENTIALS=path/to/service-account.json
```

---

## License

This project is licensed under the **Business Source License 1.1 (BUSL 1.1)**.

- ✅ You may view, fork, and self-host for **non-commercial** use
- ✅ Contributions welcome
- ❌ Commercial redistribution or hosting as a competing SaaS is not permitted without a separate commercial license

The license converts to **Apache License 2.0** on **2030-03-01**.

See [LICENSE](./LICENSE) for full terms.

---

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you'd like to change.

1. Fork the repo
2. Create your branch (`git checkout -b feature/my-feature`)
3. Commit your changes (`git commit -m 'Add my feature'`)
4. Push to the branch (`git push origin feature/my-feature`)
5. Open a Pull Request

---

## Acknowledgements

Built with [pdfminer.six](https://github.com/pdfminer/pdfminer.six), [openpyxl](https://openpyxl.readthedocs.io/), [Tesseract OCR](https://github.com/tesseract-ocr/tesseract), [PaddleOCR](https://github.com/PaddlePaddle/PaddleOCR).
