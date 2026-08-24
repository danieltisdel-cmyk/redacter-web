# Redacter ✦

Document redaction tool. Upload PDF, Word, Excel, PowerPoint, or image files and redact sensitive information with black box or blur.

## Features
- Manual term entry — type exact words to redact
- Auto-detect: Names, Dates, Times, Phone Numbers, Emails, SSNs
- OCR support — finds text inside scanned images embedded in documents
- Black box or blur redaction style
- Supports: PDF, DOCX, XLSX, PPTX, TXT, PNG, JPG, TIFF

## Deploy to Railway
[![Deploy on Railway](https://railway.app/button.svg)](https://railway.app/template)

1. Fork this repo
2. Go to railway.app → New Project → Deploy from GitHub
3. Select this repo → Railway auto-detects Dockerfile
4. Deploy — done.

## Local Development
```bash
pip install -r requirements.txt
brew install tesseract  # macOS
python app.py
# Open http://localhost:5050
```
