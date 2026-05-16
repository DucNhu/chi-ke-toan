# CHi-ke-toan Web App

## New Features & UI Updates (May 2026)

- **Nghi vấn ghép bill** (Suspicious combined bill detection) section is now displayed *above* the "Ảnh chưa được map" section for better visibility.
- **Trạng thái** (Status) column in the review table now uses colored badges for each status:
  - <span style="background:#bbf7d0;color:#166534;padding:2px 8px;border-radius:12px;font-size:12px;">Khớp</span>: Matched successfully
  - <span style="background:#fef08a;color:#a16207;padding:2px 8px;border-radius:12px;font-size:12px;">Không khớp tên</span>: Name mismatch
  - <span style="background:#fecaca;color:#b91c1c;padding:2px 8px;border-radius:12px;font-size:12px;">Không khớp</span>: No suitable bill found or OCR error
  - <span style="background:#fca5a5;color:#7f1d1d;padding:2px 8px;border-radius:12px;font-size:12px;">Nhiều bill giống nhau</span>: Multiple bills with same name and amount
  - <span style="background:#e0e7ff;color:#3730a3;padding:2px 8px;border-radius:12px;font-size:12px;">Đang đợi phân tích</span>: OCR in progress
- **Table headers are now sticky** (frozen) and remain visible while scrolling for all main tables (review, unmatched bills, suspicious bills).
- Tooltip for the status column header has been updated to explain badge meanings.

---

# Original Instructions

Build me a single-page web app (pure HTML + CSS + JS, no framework) that does the following:

## OVERVIEW
A tool to match transaction screenshot images with invoice data from an Excel file, then export each matched pair as a Word document.

## TECH STACK (all via CDN, no npm)
- SheetJS (xlsx.js) — parse Excel file
- Tesseract.js — OCR to extract text from images
- docx.js (docx library via CDN) — generate .docx files
- Vanilla JS, no framework
- TailwindCSS CDN for styling

## FEATURES

### Step 1 — Upload Excel
- Accept .xlsx file
- Parse columns: "tên bill" (bill name), "đơn giá" (price), "ngày xuất hoá đơn" (invoice date)
- Display parsed data in a preview table

### Step 2 — Upload Images
- Accept multiple image files at once (.jpg, .png, .webp)
- For each image, run Tesseract.js OCR (language: vie+eng)
- Extract money amount using these regex patterns:
  (try all patterns, pick the largest confident match)
  - /(\d{1,3}(?:[.,]\d{3})+)\s*(?:đ|VND|vnd|VNĐ|vnđ)/gi
  - /(?:số tiền|amount|tổng tiền|thành tiền|total)[^\d]*(\d{1,3}(?:[.,]\d{3})+)/gi
  - /(\d{1,3}(?:[.,]\d{3})+)/g
  - /(\d{4,})/g  ← fallback: any number >= 4 digits
- Normalize extracted amount: remove dots/commas as thousand separators → integer

### Step 3 — Auto Mapping
- For each image, find the invoice row where "đơn giá" matches the extracted amount
- Matching tolerance: ±1000 VND (in case of OCR rounding error)
- If multiple matches found → flag as "ambiguous", let user pick manually
- If no match found → flag as "unmatched", let user assign manually
- Display result as a matching table:
  | Image preview | Extracted amount | Matched bill name | Status | Action |

### Step 4 — Manual Review
- User can click any row to override the auto-match
- Dropdown shows all available invoice rows to reassign
- Show ✅ confirmed / ⚠️ ambiguous / ❌ unmatched status badges

### Step 5 — Export Word files
- For each confirmed matched pair, generate a .docx file using docx.js
- Each .docx contains:
  - Title: bill name (tên bill)
  - Line: "Đơn giá: {price} VND"
  - Line: "Ngày xuất hoá đơn: {date}"
  - The transaction screenshot image embedded below
- File name: "{tên bill}.docx"
- Add a "Xuất tất cả" button that downloads all .docx files one by one

## UI LAYOUT
- Step-by-step wizard UI (Step 1 → 2 → 3 → 4 → 5)
- Progress bar at top showing current step
- Clean card-based layout
- Show OCR progress bar per image while processing
- Responsive, works on desktop

## ERROR HANDLING
- If OCR fails on an image → mark as unmatched, show error icon
- If Excel columns not found → show friendly error message asking user to check column names
- If docx export fails → show toast notification

## IMPORTANT NOTES
- All processing is client-side only, no backend, no API calls
- Support Vietnamese text throughout (column names, labels)
- Column name matching should be case-insensitive and trim whitespace
- Normalize price comparison: treat "1.500.000" and "1500000" and "1,500,000" as equal