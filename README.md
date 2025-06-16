# Microsoft Word QC Automation

This is a web-based quality control (QC) tool that checks Microsoft Word (`.docx`) documents for FDA formatting and content consistency requirements.

The tool allows users to upload a `.docx` file and receive a detailed JSON+HTML report verifying formatting, margins, hyperlinks, and more.

---

## ✅ Features Checked

### Formatting Checks:
1. Font type must be **Times New Roman**
2. Font size must be:
   - **12pt** for normal content
   - **9pt** for table content
3. Table of Contents (TOC) presence and anchor validity
4. Page orientation must be **Portrait**

### Layout & Structure:
5. Margins must be **1 inch** on all sides
   - ❌ Error if less than 0.75in
   - ⚠️ Warning if not exactly 1in
6. Header/Footer must be at least **0.38in** from page edge

### Hyperlink Checks:
7. Internal hyperlinks must point to correct anchors
8. External hyperlinks must be valid and accessible

---
## 🚀 How to Run

1. Install dependencies:
   ```bash
   pip install flask python-docx requests
2. Run
   python app.py
3. Visit local address as outputted by terminal
   
## 📂 Project Structure
```bash
.
├── app.py                  # Main Flask server
├── uploads/                # Uploaded .docx files (auto-created)
├── templates/
│   ├── upload.html         # File upload form
│   └── result.html         # Results page
├── requirements.txt
└── README.md
```

## 🧪 Sample Output

The report displays:

❌ Font and margin violations

⚠️ Warnings for non-critical format issues

✅ Green marks for passed checks
