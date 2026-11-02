
# 📄 Smart PDF OCR – Intelligent Document Data Extraction
*Python | Computer Vision | Optical Character Recognition | Data Processing*

---

## 📖 Overview

**Smart PDF OCR – Intelligent Document Data Extraction** is a Python-based document processing project designed to automatically recognize, extract, and structure information from PDF documents using **Optical Character Recognition (OCR)** and **Computer Vision** techniques.

The system implements an automated document processing pipeline combining image preprocessing, document layout detection, multi-configuration OCR, data extraction, data cleaning, and Excel export.

Instead of applying a single OCR strategy to every document, the system introduces an intelligent routing mechanism that automatically distinguishes between **Table-based Documents** and **Text-based Documents**, selecting the appropriate processing pipeline for each document.

The project focuses on extracting structured personal information, including:

- Full Name
- Gender
- Date of Birth
- Current Residential Address
- Identification Number (CCCD/CMND)
- Personal Identification Number / Passport Number

By combining OpenCV, Tesseract OCR, Regex-based extraction, and Pandas, the system transforms unstructured PDF documents into structured and reusable Excel datasets.

---

## ✨ Features

- 🎯 **Smart Document Layout Detection:** Automatically identifies table-based and text-based PDF layouts using OpenCV, morphological operations, and horizontal/vertical line detection.

- 📊 **Table OCR Pipeline:** Detects table structures, identifies target columns, segments individual cells, and extracts dates of birth and identification numbers.

- 📝 **Text-based Document OCR:** Recognizes Vietnamese text across multiple PDF pages and extracts personal information using Regex-based field detection.

- 🧠 **Multi-attempt OCR Optimization:** Executes multiple OCR configurations with different image scaling factors, padding ratios, thresholding methods, and Tesseract Page Segmentation Modes (PSM) to improve recognition results.

- 🔍 **Intelligent Date of Birth Recognition:** Supports multiple date formats, generates candidate dates, validates calendar values, and selects results through a voting mechanism.

- 🔢 **Identification Number Extraction:** Detects and extracts identification numbers with different lengths, supporting 8–12 digit numbers in table documents and 9-digit or 12-digit numbers in text documents.

- ⚙️ **Advanced Image Preprocessing:** Applies grayscale conversion, image resizing, Gaussian Blur, Median Blur, Otsu Thresholding, Adaptive Thresholding, morphological operations, contrast enhancement, and sharpening.

- 📈 **OCR Confidence Scoring:** Evaluates identification number recognition using Tesseract confidence, agreement between OCR attempts, score margin, and identification-number length.

- 🇻🇳 **Vietnamese Text Processing:** Supports Vietnamese OCR, accent normalization, common OCR character corrections, and Vietnamese field recognition.

- 🔗 **Anchor-based ID Matching:** Uses predefined textual anchors to locate identification numbers and combines OCR results from multiple processing strategies.

- 🧹 **Data Cleaning & Deduplication:** Normalizes extracted text, removes unnecessary characters, validates selected fields, and eliminates duplicate records.

- 📑 **Automated Excel Export:** Converts extracted information into structured Excel files using Pandas and OpenPyXL, preserving identification numbers as text to prevent leading-zero loss.

- 📂 **Batch PDF Processing:** Automatically scans the input directory and processes multiple PDF documents, generating a separate Excel output for each file.

---

## 🛠 Technologies Used

| Technology | Purpose |
|---|---|
| Python 3.x | Core programming language |
| OpenCV | Image processing and table layout detection |
| Tesseract OCR | Text and numeric recognition engine |
| Pytesseract | Python interface for Tesseract OCR |
| NumPy | Image matrix operations and pixel projection |
| Pandas | Data extraction, processing, and DataFrame management |
| PDF2Image | Converting PDF pages into images |
| Pillow (PIL) | Image enhancement and preprocessing |
| OpenPyXL | Excel generation and cell formatting |
| Regular Expressions (Regex) | Pattern matching and structured data extraction |
| Poppler | PDF rendering and image conversion |

---

## 🏗 System Architecture

The system is designed around a **Smart OCR Routing Architecture**, consisting of two independent processing pipelines.

```text
                  INPUT PDF DOCUMENTS
                           |
                           v
                    MAIN CONTROLLER
                           |
                           v
                  DOCUMENT LAYOUT DETECTION
                           |
                  +--------+--------+
                  |                 |
                  v                 v
             TABLE LAYOUT       TEXT LAYOUT
                  |                 |
                  v                 v
             PIPELINE A          PIPELINE B
                  |                 |
          TABLE GRID DETECTION   FULL-PAGE OCR
                  |                 |
          COLUMN IDENTIFICATION  ANCHOR DETECTION
                  |                 |
          CELL EXTRACTION        RECORD SPLITTING
                  |                 |
          MULTI-ATTEMPT OCR      FIELD EXTRACTION
                  |                 |
          CONFIDENCE SCORING     ID MATCHING
                  |                 |
                  +--------+--------+
                           |
                           v
                   DATA PROCESSING
                           |
                           v
                    PANDAS DATAFRAME
                           |
                           v
                     EXCEL EXPORT
                           |
                           v
                    OUTPUT DIRECTORY
```

### 🔹 Pipeline A – Table-based OCR

Designed for PDF documents containing structured tables.

**Processing Workflow:**

1. Convert PDF pages into images using PDF2Image.
2. Convert images into grayscale using OpenCV.
3. Detect horizontal and vertical table lines.
4. Identify target columns containing dates of birth and identification numbers.
5. Segment individual table rows and cells.
6. Apply multiple OCR configurations to each target cell.
7. Generate and evaluate candidate recognition results.
8. Calculate OCR confidence scores.
9. Remove duplicate identification numbers.
10. Export structured data into Excel.

**Key Algorithms and Techniques:**

- Otsu Thresholding.
- Morphological Erosion and Dilation.
- Horizontal and Vertical Projection.
- Gaussian and Median Blur.
- Multi-configuration Tesseract OCR.
- Candidate Generation and Validation.
- Voting-based Date Recognition.
- Weighted OCR Confidence Scoring.

### 🔹 Pipeline B – Text-based OCR

Designed for PDF documents containing paragraphs, administrative documents, decisions, and personal information records.

**Processing Workflow:**

1. Convert PDF pages into images.
2. Perform full-page Vietnamese OCR.
3. Perform additional numeric OCR using multiple configurations.
4. Detect predefined identification-number anchors.
5. Combine OCR text from all pages.
6. Split documents into individual personal records using the "Họ và tên" field.
7. Extract personal information using Regex.
8. Retrieve identification numbers directly from record text when available.
9. Match missing identification numbers using the anchor-based queue.
10. Remove duplicate records.
11. Export structured information into Excel.

**Extracted Fields:**

- Full Name.
- Gender.
- Date of Birth.
- Current Residential Address.
- Personal Identification Number / Passport Number.

---

## 🧠 Core Algorithms & Implementation

### 1. Smart Layout Detection

The system determines document structure by analyzing horizontal and vertical lines in the first PDF page.

OpenCV morphological operations are used to isolate table boundaries, while pixel projection identifies potential line coordinates.

The detected layout is classified using the following conditions:

- At least 3 detected vertical lines.
- At least 8 detected horizontal lines.

If both conditions are satisfied, the document is classified as a table-based document. Otherwise, it is processed as a text-based document.

### 2. Multi-attempt OCR Strategy

To improve recognition robustness, the system executes multiple OCR attempts with different preprocessing configurations.

**For Identification Numbers:**

- 8 OCR configurations.
- Different image scaling factors.
- Different padding ratios.
- Multiple PSM modes.
- Optional stroke thinning.

**For Dates of Birth:**

- 8 OCR configurations.
- Otsu and Adaptive Thresholding.
- Gaussian and Median Blur.
- Multiple image scaling factors.
- Different PSM modes.

Each attempt produces a candidate result and an evaluation score. The system then selects the most suitable result based on its scoring mechanism.

### 3. Date of Birth Candidate Parsing

The date extraction module supports:

- `DD/MM/YYYY`
- `DD-MM-YYYY`
- `DD.MM.YYYY`
- Compact numeric formats such as `DDMMYYYY` and `DDMMYY`.

The system generates possible date candidates, checks calendar validity, evaluates year candidates, and removes duplicates.

When multiple OCR attempts produce different results, a voting mechanism prioritizes the date that appears most frequently, using the accumulated score as a tie-breaker.

### 4. OCR Confidence Scoring

The identification-number confidence module evaluates recognition quality using four weighted components:

| Component | Weight |
|---|---:|
| Best Tesseract Confidence | 45% |
| OCR Agreement | 30% |
| Score Margin | 15% |
| Identification Length Bonus | 10% |

The resulting score is classified into three confidence levels:

- 🟢 **HIGH:** ≥ 90%
- 🟡 **MEDIUM:** ≥ 75% and < 90%
- 🔴 **LOW:** < 75%

These scores represent internal OCR quality estimates rather than statistically calibrated probabilities of correctness.

### 5. Anchor-based Identification Matching

For text-based documents, the system uses predefined textual anchors to locate identification numbers.

Instead of searching for arbitrary numeric sequences across the entire document, it searches for identification-number patterns near recognized labels.

The extracted numbers are then combined with record-level OCR results to support identification-number matching.

### 6. Vietnamese Text Normalization

The system implements several text-processing utilities:

- Whitespace normalization.
- Vietnamese accent removal for pattern matching.
- OCR character correction (`O` → `0`, `I` → `1`, `l` → `1`).
- Common Vietnamese OCR error correction.
- Regex-based field extraction.
- Original-text position mapping after accent normalization.

These utilities improve the consistency of text extraction and make it easier to identify information despite common OCR recognition errors.

---

## 📂 Project Structure

```text
Smart-PDF-OCR/
│
├── input/
│   ├── document_01.pdf
│   ├── document_02.pdf
│   └── document_03.pdf
│
├── output/
│   ├── document_01_OCR.xlsx
│   ├── document_02_OCR.xlsx
│   └── document_03_OCR.xlsx
│
├── ocr_debug/
│
├── main.py
│
└── README.md
```

| Directory / File | Description |
|---|---|
| `input/` | Stores PDF documents for OCR processing |
| `output/` | Stores generated Excel files |
| `ocr_debug/` | Intended directory for debugging artifacts |
| `main.py` | Main application source code |
| `README.md` | Project documentation |

---

## ⚙️ Installation & Configuration

### 1. Requirements

- Python 3.x
- Tesseract OCR
- Poppler
- Required Python libraries

### 2. Install Python Dependencies

```bash
pip install numpy pandas opencv-python pdf2image pytesseract Pillow openpyxl
```

### 3. Install Tesseract OCR

Download and install Tesseract OCR:

https://github.com/tesseract-ocr/tesseract

Ensure that the following language data files are available:

- `eng` – English.
- `vie` – Vietnamese.

### 4. Install Poppler

Download Poppler for Windows:

https://github.com/oschwartz10612/poppler-windows/releases

Extract the package and locate the `Library/bin` directory.

### 5. Configure Environment Paths

Update the configuration section in the Python source code according to your local installation:

```python
POPPLER_PATH = r"D:\poppler\Library\bin"

pytesseract.pytesseract.tesseract_cmd = (
    r"C:\Program Files\Tesseract-OCR\tesseract.exe"
)

os.environ["TESSDATA_PREFIX"] = (
    r"C:\Program Files\Tesseract-OCR"
)
```

---

## 🚀 How to Run

### Step 1 – Prepare Input Documents

Place one or more PDF files into the `input/` directory.

### Step 2 – Configure OCR Environment

Ensure that Poppler and Tesseract paths are configured correctly.

### Step 3 – Execute the Application

Run the following command:

```bash
python main.py
```

### Step 4 – Automatic Processing

The application automatically:

1. Scans the input directory.
2. Detects the document layout.
3. Selects the appropriate OCR pipeline.
4. Processes all PDF pages.
5. Extracts and cleans structured information.
6. Generates an Excel file for each PDF document.

### Step 5 – Retrieve Results

Open the `output/` directory to access the generated Excel files.

---

## 📊 Output Format

### Pipeline A – Table OCR

| Column | Description |
|---|---|
| Ngày tháng năm sinh | Extracted date of birth |
| Số CMTND | Extracted identification number |
| Độ tin cậy OCR (%) | Estimated OCR confidence score |

### Pipeline B – Text OCR

| Column | Description |
|---|---|
| Họ và tên | Full name |
| Giới tính | Gender |
| Ngày sinh | Date of birth |
| Nơi ở hiện tại | Current residential address |
| Số định danh/Hộ chiếu | Personal identification / passport number |

**Excel Formatting:** Identification-number columns are explicitly converted to text format to preserve leading zeros and prevent Excel from automatically converting identification numbers into numeric values.

---

## 🔧 Configuration Parameters

The following parameters can be adjusted in the source code:

| Parameter | Default Value | Description |
|---|---:|---|
| `DPI_TABLE` | 320 | Rendering resolution for table documents |
| `DPI_TEXT` | 300 | Rendering resolution for text documents |
| `CMT_COL_INDEX` | 4 | Target column index for identification-number extraction |
| `STT_OFFSET` | 4 | Internal row numbering offset |
| `DOB_MIN_YEAR` | 1900 | Minimum accepted year of birth |
| `DOB_MAX_YEAR` | Current Year | Maximum accepted year of birth |
| `DEBUG` | True | Enables console debugging output |
| `DEBUG_SAVE` | False | Reserved configuration for saving debug images |

---

## ⚠️ Current Limitations

- Document layout detection currently analyzes only the first PDF page.
- Table column identification relies on predefined column positions.
- Text-based identification matching uses document order when direct record-level extraction is unavailable.
- OCR confidence scores are heuristic estimates and have not been statistically calibrated.
- Text-based date extraction does not currently perform full calendar validation.
- Duplicate removal by full name may remove distinct individuals with identical names.
- Low-confidence OCR results are not automatically separated into a dedicated manual-review worksheet.
- The application currently operates through a Python command-line workflow without a graphical user interface.

For sensitive documents, extracted information should be reviewed before being used in official records or administrative processes.

---

## 🔮 Future Improvements

Potential development directions include:

- 📌 Page-level Smart Layout Detection.
- 📌 Automatic table header and column recognition.
- 📌 Coordinate-based identification-number matching.
- 📌 Improved personal information validation.
- 📌 Advanced OCR confidence calibration.
- 📌 Manual OCR review and correction interface.
- 📌 Dedicated low-confidence data reporting.
- 📌 PyQt6 desktop application.
- 📌 Multi-threaded PDF processing.
- 📌 Processing history and performance statistics.
- 📌 Enhanced data protection and access control.

---

## 🔐 Privacy & Data Security

This project processes potentially sensitive personal information, including names, dates of birth, addresses, and identification numbers.

Users should ensure that:

- PDF documents are processed only with appropriate authorization.
- Personal information is not committed to public GitHub repositories.
- Input and output directories are properly protected.
- Sensitive information is excluded from debugging logs.
- Generated Excel files are stored securely.

---

## 👨‍💻 Author & Project Information

**Project:** Smart PDF OCR – Intelligent Document Data Extraction

**Programming Language:** Python

**Domain:** Artificial Intelligence / Computer Vision / OCR / Data Processing

**Core Technologies:** OpenCV, Tesseract OCR, Pytesseract, NumPy, Pandas, Pillow, PDF2Image, OpenPyXL.

**Project Objective:** Automating PDF document recognition and personal information extraction through an intelligent, multi-pipeline OCR processing architecture.

---

<p align="center">
  <b>📄 Smart PDF OCR</b><br>
  <i>Turning Unstructured Documents into Structured Data.</i>
</p>
