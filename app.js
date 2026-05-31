// Globals
let invoices = []; // {id, billName, priceRaw, price, date, fullName, fundType}
let images = [];   // {id, file, dataUrl, ocrText, extractedAmount, status, matchedInvoiceId, ocrProgress}

const excelInput = document.getElementById('excel-input');
const invoicePreview = document.getElementById('invoice-preview');
const excelError = document.getElementById('excel-error');
const imagesInput = document.getElementById('images-input');
const imagesList = document.getElementById('images-list');
const runMappingBtn = document.getElementById('run-mapping');
const mappingTable = document.getElementById('mapping-table');
const reviewTable = document.getElementById('review-table');
const exportAllBtn = document.getElementById('export-all');
const exportLog = document.getElementById('export-log');
const progressBar = document.getElementById('progress-bar');
const mappingSummary = document.getElementById('mapping-summary');
const unmatchedImagesSection = document.getElementById('unmatched-images-section');
const unmatchedBillsSection = document.getElementById('unmatched-bills-section');
const imagesProgressTitle = document.getElementById('images-progress-title');
const suspiciousBillSection = document.getElementById('suspicious-bill-section');
const imagePopup = document.getElementById('image-popup');
const imagePopupImg = document.getElementById('image-popup-img');
const excelDropzone = document.getElementById('excel-dropzone');
const excelFileName = document.getElementById('excel-file-name');
const imagesDropzone = document.getElementById('images-dropzone');
const imagesFileCount = document.getElementById('images-file-count');

// Image popup: click to close
if (imagePopup) {
  imagePopup.addEventListener('click', () => { imagePopup.style.display = 'none'; });
}

// Delegate click on any img with class .img-popup to open popup
document.addEventListener('click', (e) => {
  const target = e.target.closest('.img-popup');
  if (target && imagePopup && imagePopupImg) {
    imagePopupImg.src = target.src;
    imagePopup.style.display = 'flex';
  }
});

// --- Dropzone helpers ---
function setupDropzone(dropzone, fileInput, onFiles) {
  if (!dropzone || !fileInput) return;

  // Click to open file picker
  dropzone.addEventListener('click', (e) => {
    if (e.target === fileInput) return;
    fileInput.click();
  });

  // Drag & drop visual feedback
  ['dragenter', 'dragover'].forEach(evt => {
    dropzone.addEventListener(evt, (e) => {
      e.preventDefault(); e.stopPropagation();
      dropzone.classList.add('drag-over');
    });
  });
  ['dragleave', 'drop'].forEach(evt => {
    dropzone.addEventListener(evt, (e) => {
      e.preventDefault(); e.stopPropagation();
      dropzone.classList.remove('drag-over');
    });
  });

  // Handle drop
  dropzone.addEventListener('drop', (e) => {
    const files = e.dataTransfer.files;
    if (files.length) {
      fileInput.files = files;
      fileInput.dispatchEvent(new Event('change', { bubbles: true }));
    }
  });

  // Show file name/count after selection
  if (onFiles) {
    fileInput.addEventListener('change', () => onFiles(fileInput.files));
  }
}

setupDropzone(excelDropzone, excelInput, (files) => {
  if (files.length && excelFileName) {
    excelFileName.textContent = '📄 ' + files[0].name;
    excelFileName.classList.remove('hidden');
  }
});

setupDropzone(imagesDropzone, imagesInput, (files) => {
  if (files.length && imagesFileCount) {
    imagesFileCount.textContent = '🖼️ Đã chọn ' + files.length + ' ảnh';
    imagesFileCount.classList.remove('hidden');
  }
});

function setStepProgress(step) {
  const pct = Math.min(100, Math.max(0, (step-1) * 25));
  progressBar.style.width = pct + '%';
}
setStepProgress(1);

// Helpers
function normalizeColName(s) {
  return String(s || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'd')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}
function normalizePriceToInt(s) {
  if (s == null) return null;
  if (typeof s === 'number' && Number.isFinite(s)) {
    return Math.round(s);
  }
  const str = String(s).trim();
  if (!str) return null;
  const numeric = Number(str.replace(/\s+/g, '').replace(/,/g, ''));
  if (Number.isFinite(numeric)) {
    return Math.round(numeric);
  }
  const compact = str.replace(/\s+/g, '').replace(/[.,]/g, '');
  const n = parseInt(compact, 10);
  return Number.isNaN(n) ? null : n;
}

function formatAmount(value) {
  const amount = normalizePriceToInt(value);
  if (amount == null) return '—';
  return amount.toLocaleString('vi-VN');
}

function normalizeMatchText(value) {
  return normalizeColName(value)
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function hasStrongTokenMatch(sourceText, candidateText) {
  const source = normalizeMatchText(sourceText);
  const candidate = normalizeMatchText(candidateText);
  if (!source || !candidate) return false;
  if (source.includes(candidate)) return true;

  const tokens = candidate.split(' ').filter(token => token.length >= 2);
  if (!tokens.length) return false;

  const matchedTokens = tokens.filter(token => source.includes(token)).length;
  const requiredTokens = tokens.length <= 2 ? tokens.length : tokens.length - 1;
  return matchedTokens >= requiredTokens && matchedTokens >= 2;
}

function getInvoiceNameMatchInfo(img, invoice) {
  if (!invoice) {
    return { matched: false, label: 'Chưa ghép bill' };
  }

  const candidates = [
    { text: invoice.fullName, label: 'Họ và tên' },
    { text: invoice.billName, label: 'Nội dung' }
  ].filter(item => normalizeMatchText(item.text));

  if (!candidates.length) {
    return { matched: true, label: 'Bill không có tên để kiểm tra' };
  }

  for (const candidate of candidates) {
    if (hasStrongTokenMatch(img?.ocrText, candidate.text)) {
      return { matched: true, label: `${candidate.label} khớp` };
    }
  }

  return { matched: false, label: 'Tên trong ảnh không khớp' };
}

function getAutoMatchDecision(img, usedInvoiceIds = new Set()) {
  if (img.extractedAmount == null) {
    return { invoice: null, reason: 'missing-amount' };
  }

  const amountMatches = invoices.filter(inv => inv.price === img.extractedAmount && !usedInvoiceIds.has(inv.id));
  if (!amountMatches.length) {
    return { invoice: null, reason: 'amount-mismatch' };
  }

  const nameMatches = amountMatches.filter(inv => getInvoiceNameMatchInfo(img, inv).matched);
  if (nameMatches.length === 1) {
    return { invoice: nameMatches[0], reason: 'matched' };
  }

  if (nameMatches.length > 1) {
    return { invoice: null, reason: 'duplicate-name-match' };
  }

  return { invoice: null, reason: 'name-mismatch' };
}

function getImageReviewStatus(img, matchedInvoice) {
  if (img.status === 'pending') {
    return { statusText: 'Đang đợi phân tích', nameCheckText: 'Đang kiểm tra', detailText: 'OCR đang chạy' };
  }

  if (!matchedInvoice) {
    if (img.matchReason === 'name-mismatch') {
      return { statusText: 'Không khớp tên', nameCheckText: 'Không khớp', detailText: 'Trùng số tiền nhưng khác tên' };
    }
    if (img.matchReason === 'duplicate-name-match') {
      return { statusText: 'Nhiều bill giống nhau', nameCheckText: 'Chưa rõ', detailText: 'Có nhiều bill cùng tên và số tiền' };
    }
    if (img.status === 'unmatched') {
      return { statusText: 'Không khớp', nameCheckText: 'Không kiểm tra được', detailText: 'Không tìm thấy bill phù hợp' };
    }
    return { statusText: '—', nameCheckText: '—', detailText: '' };
  }

  const nameMatch = getInvoiceNameMatchInfo(img, matchedInvoice);
  return {
    statusText: nameMatch.matched ? 'Khớp' : 'Không khớp tên',
    nameCheckText: nameMatch.matched ? 'Khớp' : 'Không khớp',
    detailText: nameMatch.label
  };
}
function normalizeExcelDate(value) {
  if (value == null || value === '') return '';

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const d = new Date(value);

    // cộng thêm 1 ngày
    d.setDate(d.getDate() + 1);

    const vnDate =
      d.getDate().toString().padStart(2, '0') + '/' +
      (d.getMonth() + 1).toString().padStart(2, '0') + '/' +
      d.getFullYear();

    return vnDate;
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    // Excel serial date
    const parsed = XLSX.SSF.parse_date_code(value + 1);

    if (parsed) {
      const y = String(parsed.y).padStart(4, '0');
      const m = String(parsed.m).padStart(2, '0');
      const d = String(parsed.d).padStart(2, '0');

      return `${y}-${m}-${d}`;
    }
  }

  return String(value).trim();
}

function createInvoiceRecord(record, id) {
  return {
    id,
    billName: String(record.billName || '').trim(),
    priceRaw: record.priceRaw,
    price: normalizePriceToInt(record.priceRaw),
    date: normalizeExcelDate(record.date),
    fullName: String(record.fullName || '').trim(),
    fundType: String(record.fundType || '').trim()
  };
}

function findSimpleColumn(headerKeys, variants) {
  const normalized = headerKeys.map(k => normalizeColName(k));
  for (const variant of variants) {
    const idx = normalized.indexOf(normalizeColName(variant));
    if (idx >= 0) return headerKeys[idx];
  }
  return null;
}

function parseSimpleInvoiceSheet(sheet) {
  const json = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: true });
  if (!json.length) return [];
  const headerKeys = Object.keys(json[0]);
  const colBill = findSimpleColumn(headerKeys, ['tên bill', 'ten bill', 'name', 'nội dung', 'noi dung']);
  const colPrice = findSimpleColumn(headerKeys, ['số tiền', 'so tien', 'đơn giá', 'don gia', 'price', 'gia']);
  const colDate = findSimpleColumn(headerKeys, ['ngày xuất hoá đơn', 'ngay xuat hoa don', 'ngày tháng', 'ngay thang', 'date']);
  if (!colBill || !colPrice || !colDate) return [];
  const colFullName = findSimpleColumn(headerKeys, ['họ và tên', 'ho va ten', 'người nộp', 'nguoi nop', 'người nhận', 'nguoi nhan']);

  return json
    .map((row, index) => createInvoiceRecord({
      billName: row[colBill],
      priceRaw: row[colPrice],
      date: row[colDate],
      fullName: colFullName ? row[colFullName] : '',
      fundType: row.loại || row['Loại quỹ'] || ''
    }, index))
    .filter(inv => inv.price != null && (inv.billName || inv.fullName || inv.date));
}

function detectCashbookBlocks(rows) {
  const headerRowIndex = rows.findIndex(row => {
    const normalized = row.map(cell => normalizeColName(cell));
    return normalized.filter(cell => cell === 'so tien').length >= 2
      && normalized.filter(cell => cell === 'ngay thang').length >= 2
      && normalized.filter(cell => cell === 'ho va ten').length >= 2;
  });
  if (headerRowIndex <= 0) return null;

  const groupRow = rows[headerRowIndex - 1] || [];
  const normalizedGroups = groupRow.map(cell => normalizeColName(cell));
  const blockStarts = [];
  normalizedGroups.forEach((label, index) => {
    if (label) blockStarts.push({ label, index });
  });

  const blocks = blockStarts
    .map((block, index) => ({
      label: block.label,
      start: block.index,
      end: (blockStarts[index + 1] ? blockStarts[index + 1].index : rows[headerRowIndex].length) - 1
    }))
    .filter(block => block.label === 'thu' || block.label === 'chi')
    .map(block => {
      const headers = rows[headerRowIndex].slice(block.start, block.end + 1).map(cell => normalizeColName(cell));
      const findOffset = (variants) => {
        for (const variant of variants) {
          const idx = headers.indexOf(normalizeColName(variant));
          if (idx >= 0) return block.start + idx;
        }
        return -1;
      };
      return {
        fundType: block.label === 'thu' ? 'Thu' : 'Chi',
        dateIndex: findOffset(['ngày tháng', 'ngay thang', 'date']),
        fullNameIndex: findOffset(['họ và tên', 'ho va ten']),
        billNameIndex: findOffset(['nội dung', 'noi dung', 'tên bill', 'ten bill']),
        amountIndex: findOffset(['số tiền', 'so tien', 'đơn giá', 'don gia'])
      };
    })
    .filter(block => block.dateIndex >= 0 && block.billNameIndex >= 0 && block.amountIndex >= 0);

  if (!blocks.length) return null;
  return { headerRowIndex, blocks };
}

function parseCashbookSheet(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true });
  const detected = detectCashbookBlocks(rows);
  if (!detected) return [];

  const records = [];
  for (let rowIndex = detected.headerRowIndex + 1; rowIndex < rows.length; rowIndex++) {
    const row = rows[rowIndex] || [];
    detected.blocks.forEach(block => {
      const rawAmount = row[block.amountIndex];
      const amount = normalizePriceToInt(rawAmount);
      const billName = row[block.billNameIndex];
      const fullName = block.fullNameIndex >= 0 ? row[block.fullNameIndex] : '';
      const date = row[block.dateIndex];
      if (amount == null) return;
      if (!String(billName || '').trim() && !String(fullName || '').trim() && (date == null || date === '')) return;
      records.push(createInvoiceRecord({
        billName,
        priceRaw: rawAmount,
        date,
        fullName,
        fundType: block.fundType
      }, records.length));
    });
  }
  return records;
}

function parseInvoicesFromSheet(sheet) {
  const cashbookInvoices = parseCashbookSheet(sheet);
  if (cashbookInvoices.length) {
    return cashbookInvoices.map((record, index) => ({ ...record, id: index }));
  }
  const simpleInvoices = parseSimpleInvoiceSheet(sheet);
  return simpleInvoices.map((record, index) => ({ ...record, id: index }));
}

// Excel handling
excelInput.addEventListener('change', (e) => {
  const f = e.target.files[0];
  if (!f) return;
  const reader = new FileReader();
  reader.onload = function(evt) {
    try {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: 'array', raw: true, cellDates: true });
      const firstSheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[firstSheetName];
      const parsedInvoices = parseInvoicesFromSheet(sheet);
      if (!parsedInvoices.length) {
        excelError.textContent = 'Excel trống hoặc không đọc được.';
        return;
      }
      excelError.textContent = '';
      invoices = parsedInvoices;
      renderInvoicePreview();
      renderAllMappingUI();
      setStepProgress(2);
    } catch (err) {
      console.error(err);
      excelError.textContent = 'Lỗi đọc file Excel: ' + err.message;
    }
  };
  reader.readAsArrayBuffer(f);
});

function renderInvoicePreview() {
  if (!invoices.length) {
    invoicePreview.innerHTML = '<div class="text-sm text-gray-500">Chưa có hóa đơn nào.</div>';
    return;
  }
  let html = '<div class="overflow-x-auto"><table class="min-w-full text-sm"><thead><tr><th class="text-left p-2">#</th><th class="text-left p-2">Loại quỹ</th><th class="text-left p-2">Họ và tên</th><th class="text-left p-2">Nội dung</th><th class="text-left p-2">Số tiền</th><th class="text-left p-2">Ngày</th></tr></thead><tbody>';
  invoices.forEach(inv => {
    html += `<tr><td class="p-2">${inv.id+1}</td><td class="p-2">${inv.fundType || '—'}</td><td class="p-2">${inv.fullName || '—'}</td><td class="p-2">${inv.billName || '—'}</td><td class="p-2">${formatAmount(inv.priceRaw)}</td><td class="p-2">${inv.date || '—'}</td></tr>`;
  });
  html += '</tbody></table></div>';
  invoicePreview.innerHTML = html;
}

// Images handling & OCR
imagesInput.addEventListener('change', (e) => {
  const files = Array.from(e.target.files || []);
  if (!files.length) {
    if (imagesProgressTitle) imagesProgressTitle.textContent = '';
    return;
  }
  images = files.map((f, idx) => {
    return { id: idx, file: f, dataUrl: '', ocrText: '', extractedAmount: null, status: 'pending', matchedInvoiceId: null, ocrProgress: 0, matchReason: '' };
  });
  imagesList.innerHTML = '';
  if (imagesProgressTitle) imagesProgressTitle.textContent = `Đang phân tích 0/${images.length} ảnh`;
  let done = 0;
  images.forEach((img, i) => {
    const reader = new FileReader();
    reader.onload = function(evt) {
      img.dataUrl = evt.target.result;
      renderImagesList();
      runOCRForImage(img).then(() => {
        done++;
        if (imagesProgressTitle) imagesProgressTitle.textContent = `Đang phân tích ${done}/${images.length} ảnh`;
        // Khi xong hết thì hiện "Đã phân tích xong N ảnh"
        if (done === images.length && images.length > 0) {
          setTimeout(() => {
            if (imagesProgressTitle) imagesProgressTitle.textContent = `Đã phân tích xong ${images.length} ảnh`;
          }, 800);
        }
      });
      renderAllMappingUI();
    };
    reader.readAsDataURL(img.file);
  });
  renderAllMappingUI();
  setStepProgress(3);
});

function renderImagesList() {
  imagesList.innerHTML = images.map(img => {
    return `<div class="p-3 bg-white rounded shadow">
      <img src="${img.dataUrl}" class="img-preview img-popup mb-2" style="cursor:pointer;" />
      <div class="text-sm">Số tiền trong ảnh: ${img.extractedAmount != null ? formatAmount(img.extractedAmount) : '—'}</div>
      <div class="mt-2 progress-small"><div style="width:${img.ocrProgress}%"></div></div>
    </div>`;
  }).join('');
}

async function runOCRForImage(img) {
  try {
    img.ocrProgress = 5; renderImagesList();
    const { data: { text } } = await Tesseract.recognize(
      img.dataUrl,
      'vie+eng',
      {
        logger: m => {
          if (m.status === 'recognizing text' && m.progress) {
            img.ocrProgress = Math.round(m.progress * 100);
            renderImagesList();
          }
        }
      }
    );
    console.log('OCR text for image', img.id, ':', text);
    img.ocrText = text;
    img.ocrProgress = 100; renderImagesList();
    const extracted = extractAmountFromText(text);
    img.extractedAmount = extracted;
    img.matchReason = extracted != null ? 'ocr-ready' : 'missing-amount';
    img.status = extracted != null ? 'ocr' : 'unmatched';
    renderImagesList();
    renderAllMappingUI();
  } catch (err) {
    console.error('OCR error', err);
    img.status = 'unmatched';
    img.matchReason = 'ocr-error';
    img.ocrProgress = 100; renderImagesList();
    renderAllMappingUI();
  }
}

function extractAmountFromText(text) {
  if (!text) return null;
  const patterns = [
    /(\d{1,3}(?:[.,]\d{3})+)\s*(?:đ|VND|vnd|VNĐ|vnđ)/gi,
    /(?:số tiền|amount|tổng tiền|thành tiền|total)[^\d]*(\d{1,3}(?:[.,]\d{3})+)/gi,
    /(\d{1,3}(?:[.,]\d{3})+)/g,
    /(\d{4,})/g
  ];
  let matches = [];
  for (const p of patterns) {
    let m;
    while ((m = p.exec(text)) !== null) {
      // Lấy group cuối cùng khác null
      let val = null;
      for (let i = m.length - 1; i >= 1; --i) {
        if (m[i]) { val = m[i]; break; }
      }
      if (val) matches.push(val);
    }
    if (matches.length) break;
  }
  if (!matches.length) return null;
  // pick largest numeric value after normalization
  const nums = matches.map(s => parseInt(String(s).replace(/[^\d]/g, ''), 10)).filter(n => !Number.isNaN(n));
  if (!nums.length) return null;
  return Math.max(...nums);
}

// Auto mapping
runMappingBtn.addEventListener('click', () => {
  if (!invoices.length) { alert('Vui lòng upload Excel trước.'); return; }
  // Reset all invoice usage
  const usedInvoiceIds = new Set();
  images.forEach(img => {
    if (img.status === 'ocr') {
      const decision = getAutoMatchDecision(img, usedInvoiceIds);
      img.matchedInvoiceId = decision.invoice ? decision.invoice.id : null;
      img.matchReason = decision.reason;
      img.status = decision.invoice ? 'matched' : 'unmatched';
      if (decision.invoice) usedInvoiceIds.add(decision.invoice.id);
    }
  });
  renderImagesList();
  renderReviewTable();
  renderAllMappingUI();
  setStepProgress(4);
});

function getReviewRowData(img, idx) {
  const matched = invoices.find(i => i.id === img.matchedInvoiceId);
  const reviewStatus = getImageReviewStatus(img, matched);
  return {
    index: idx + 1,
    img,
    matched,
    imageFileName: img.file?.name || `anh-giao-dich-${idx + 1}`,
    extractedAmount: img.extractedAmount != null ? formatAmount(img.extractedAmount) : '—',
    fundType: matched ? matched.fundType || '—' : '—',
    fullName: matched ? matched.fullName || '—' : '—',
    billName: matched ? matched.billName || '—' : '—',
    date: matched ? matched.date || '—' : '—',
    nameCheckText: reviewStatus.nameCheckText || '—',
    statusText: reviewStatus.statusText || '—',
    detailText: reviewStatus.detailText || '',
    badgeClass: getStatusBadgeClass(reviewStatus.statusText)
  };
}

function getStatusBadgeClass(statusText) {
  if (statusText === 'Khớp') return 'badge-matched';
  if (statusText === 'Không khớp tên') return 'badge-name-mismatch';
  if (statusText === 'Nhiều bill giống nhau') return 'badge-duplicate';
  if (statusText === 'Đang đợi phân tích') return 'badge-pending';
  return 'badge-unmatched';
}

function renderReviewTable() {
  if (!images.length) {
    reviewTable.innerHTML = '<div class="text-sm text-gray-500">Chưa có ảnh nào.</div>';
    renderAllMappingUI();
    return;
  }

  let html = `<div class="overflow-x-auto"><table class="min-w-full text-sm review-table-with-sticky"><thead><tr><th class="p-2">#</th><th class="p-2">Ảnh</th><th class="p-2">Số tiền trong ảnh</th><th class="p-2">Loại quỹ</th><th class="p-2">Họ và tên</th><th class="p-2">Nội dung</th><th class="p-2">Kiểm tra tên</th><th class="p-2">Trạng thái <span title="\n<badge class='badge-matched'>Khớp</badge>: Ảnh đã ghép thành công với bill có cùng tên và số tiền.\n<badge class='badge-name-mismatch'>Không khớp tên</badge>: Có bill cùng số tiền nhưng tên trên ảnh không khớp.\n<badge class='badge-unmatched'>Không khớp</badge>: Không tìm thấy bill phù hợp hoặc OCR lỗi.\n<badge class='badge-duplicate'>Nhiều bill giống nhau</badge>: Có nhiều bill cùng tên và số tiền.\n<badge class='badge-pending'>Đang đợi phân tích</badge>: OCR đang chạy.\n" style="cursor: help; color: #2563eb;">&#9432;</span></th><th class="p-2">Word</th></tr></thead><tbody>`;

  images.forEach((img, idx) => {
    const row = getReviewRowData(img, idx);
    html += `<tr>
      <td class="p-2">${row.index}</td>
      <td class="p-2"><img src="${img.dataUrl}" class="img-popup" style="width:80px;height:50px;object-fit:cover;border-radius:6px;cursor:pointer;" /></td>
      <td class="p-2">${row.extractedAmount}</td>
      <td class="p-2">${row.fundType}</td>
      <td class="p-2">${row.fullName}</td>
      <td class="p-2">${row.billName}</td>
      <td class="p-2">${row.nameCheckText}</td>
      <td class="p-2"><span class="badge ${row.badgeClass}">${row.statusText}</span><div class="text-xs text-gray-500">${row.detailText}</div></td>
      <td class="p-2">
        <button type="button" data-export-review-row="${idx}" class="bg-green-600 hover:bg-green-700 text-white px-3 py-1 rounded text-xs whitespace-nowrap disabled:opacity-50" ${img.dataUrl ? '' : 'disabled'}>
          Tải Word
        </button>
      </td>
    </tr>`;
  });
  html += '</tbody></table></div>';
  reviewTable.innerHTML = html;
  renderAllMappingUI();
}

function getMatchedAndUnmatched() {
  // Images: unmatched if status !== 'matched' or matchedInvoiceId == null
  const matchedImages = images.filter(img => img.status === 'matched' && img.matchedInvoiceId != null);
  const unmatchedImages = images.filter(img => img.status !== 'matched' || img.matchedInvoiceId == null);

  // Bills: unmatched if not used by any image
  const matchedBillIds = new Set(matchedImages.map(img => img.matchedInvoiceId));
  const unmatchedBills = invoices.filter(inv => !matchedBillIds.has(inv.id));

  return { matchedImages, unmatchedImages, unmatchedBills };
}

function renderMappingSummary() {
  const { matchedImages, unmatchedImages, unmatchedBills } = getMatchedAndUnmatched();
  mappingSummary.innerHTML = `
    <div class="flex flex-wrap gap-4 items-center">
      <div class="bg-green-100 text-green-800 px-4 py-2 rounded shadow transition-all">
        Đã ghép: <span class="font-bold">${matchedImages.length}</span>
      </div>
      <div class="bg-yellow-100 text-yellow-800 px-4 py-2 rounded shadow transition-all">
        Ảnh chưa ghép: <span class="font-bold">${unmatchedImages.length}</span>
      </div>
      <div class="bg-red-100 text-red-800 px-4 py-2 rounded shadow transition-all">
        Bill chưa ghép: <span class="font-bold">${unmatchedBills.length}</span>
      </div>
    </div>
  `;
}

function renderUnmatchedImages() {
  const { unmatchedImages } = getMatchedAndUnmatched();
  // Get suspicious images (those matched in suspicious bill section)
  const suspicious = findSuspiciousBillMappings();
  const suspiciousImageIds = new Set();
  suspicious.forEach(item => item.images.forEach(img => suspiciousImageIds.add(img.id)));
  if (!unmatchedImages.length) {
    unmatchedImagesSection.innerHTML = '';
    return;
  }
  unmatchedImagesSection.innerHTML = `
    <div class="card p-4 bg-white rounded shadow transition-all duration-300">
      <h2 class="font-medium mb-3">Ảnh chưa được map</h2>
      <div class="scroll-container">
        <div class="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4">
          ${unmatchedImages.map(img => `
            <div class="p-3 bg-gray-50 rounded shadow flex flex-col items-center">
              <img src="${img.dataUrl}" class="mb-2 rounded max-h-32 object-contain border img-popup" style="cursor:pointer; border: ${suspiciousImageIds.has(img.id) ? '2px solid #fbbf24' : '1px solid #e5e7eb'};" />
              <div class="text-xs text-gray-700 truncate w-full">${img.file?.name || ''}</div>
              <div class="text-sm mt-1">Số tiền trong ảnh: <span class="font-semibold">${img.extractedAmount != null ? formatAmount(img.extractedAmount) : '—'}</span></div>
              <div class="text-xs mt-1 ${img.status === 'unmatched' ? 'text-red-600' : 'text-gray-500'}">
                ${img.matchReason === 'name-mismatch' ? 'Trùng số tiền nhưng tên không khớp' : (img.status === 'unmatched' ? (img.ocrText ? 'Không khớp bill nào' : 'OCR lỗi hoặc không đọc được số tiền') : (img.status === 'pending' ? 'Đang nhận diện...' : 'Đã nhận diện'))}
                ${suspiciousImageIds.has(img.id) ? '<span class="badge badge-suspicious ml-2">Nghi vấn ghép</span>' : ''}
              </div>
            </div>
          `).join('')}
        </div>
      </div>
    </div>
  `;
}

function renderUnmatchedBills() {
  const { unmatchedBills } = getMatchedAndUnmatched();
  // Get suspicious bills (those matched in suspicious bill section)
  const suspicious = findSuspiciousBillMappings();
  const suspiciousBillIds = new Set(suspicious.map(item => item.bill.id));
  if (!unmatchedBills.length) {
    unmatchedBillsSection.innerHTML = '';
    return;
  }
  unmatchedBillsSection.innerHTML = `
    <div class="card p-4 bg-white rounded shadow transition-all duration-300">
      <h2 class="font-medium mb-3">Bill chưa được map</h2>
      <div class="scroll-container overflow-x-auto">
        <table class="min-w-full text-sm">
          <thead>
            <tr>
              <th class="p-2 text-left">Loại quỹ</th>
              <th class="p-2 text-left">Họ và tên</th>
              <th class="p-2 text-left">Nội dung</th>
              <th class="p-2 text-left">Số tiền</th>
              <th class="p-2 text-left">Ngày</th>
              <th class="p-2 text-left">Ghi chú</th>
            </tr>
          </thead>
          <tbody>
            ${unmatchedBills.map(inv => `
              <tr>
                <td class="p-2">${inv.fundType || '—'}</td>
                <td class="p-2">${inv.fullName || '—'}</td>
                <td class="p-2">${inv.billName || '—'}</td>
                <td class="p-2">${formatAmount(inv.priceRaw)}</td>
                <td class="p-2">${inv.date || '—'}</td>
                <td class="p-2">${suspiciousBillIds.has(inv.id) ? '<span class=\"badge badge-suspicious\">Nghi vấn ghép</span>' : ''}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    </div>
  `;
}

function renderAllMappingUI() {
  renderMappingSummary();
  renderSuspiciousBillTable(); // Move suspicious bill section above unmatched images
  renderUnmatchedImages();
  renderUnmatchedBills();
}

// --- Suspicious combined bill detection ---

// Extract sub-amounts from bill text (handles Vietnamese conventions: 70k, 36k, 1.578.982, 462300, etc.)
function extractSubAmountsFromBillText(text) {
  if (!text) return [];
  const results = [];
  const seen = new Set();
  const addAmount = (n) => {
    if (n > 0 && !seen.has(n)) { seen.add(n); results.push(n); }
  };

  // Numbers followed by 'k' (case insensitive) — e.g. 70k, 36k, 150K
  const kPattern = /(\d{1,6})\s*k\b/gi;
  let m;
  while ((m = kPattern.exec(text)) !== null) {
    addAmount(parseInt(m[1], 10) * 1000);
  }

  // Formatted numbers with dot/comma separators — e.g. 1.578.982 or 1,578,982
  const fmtPattern = /(\d{1,3}(?:[.,]\d{3})+)/g;
  while ((m = fmtPattern.exec(text)) !== null) {
    addAmount(parseInt(m[1].replace(/[.,]/g, ''), 10));
  }

  // Plain numbers >= 4 digits — e.g. 462300
  const plainPattern = /(?<!\d)(\d{4,})(?!\d)/g;
  while ((m = plainPattern.exec(text)) !== null) {
    addAmount(parseInt(m[1], 10));
  }

  return results;
}

function findSuspiciousBillMappings() {
  const { unmatchedImages, unmatchedBills } = getMatchedAndUnmatched();
  if (unmatchedImages.length < 2 || unmatchedBills.length < 1) return [];

  const result = [];
  unmatchedBills.forEach(bill => {
    // Combine billName + fullName for analysis
    const combinedText = [bill.billName, bill.fullName].filter(Boolean).join(' ');
    const subAmounts = extractSubAmountsFromBillText(combinedText);

    // Only consider sub-amounts that are smaller than the bill price
    const validSubs = subAmounts.filter(x => bill.price != null && x < bill.price && x >= 1000);
    if (validSubs.length < 2) return;

    // Find unmatched images whose extractedAmount matches any sub-amount (with ±1000 tolerance)
    const matchedImgs = unmatchedImages.filter(img => {
      if (img.extractedAmount == null) return false;
      return validSubs.some(sub => Math.abs(img.extractedAmount - sub) <= 1000);
    });

    if (matchedImgs.length >= 2) {
      // Check if sum of matched image amounts is close to bill price
      const imgSum = matchedImgs.reduce((s, img) => s + (img.extractedAmount || 0), 0);
      const sumClose = bill.price != null && Math.abs(imgSum - bill.price) <= 2000;
      result.push({
        bill,
        images: matchedImgs,
        subAmounts: validSubs,
        imgSum,
        sumClose
      });
    }
  });
  return result;
}

function renderSuspiciousBillTable() {
  if (!suspiciousBillSection) return;
  const suspicious = findSuspiciousBillMappings();
  if (!suspicious.length) {
    suspiciousBillSection.innerHTML = '';
    return;
  }
  suspiciousBillSection.innerHTML = `
    <div class="card p-4 bg-white rounded shadow transition-all duration-300">
      <h2 class="font-medium mb-3 text-yellow-700">⚠️ Nghi vấn ghép bill <span id="suspicious-tooltip-icon" style="cursor:help; color:#2563eb; font-size:16px;">&#9432;</span></h2>
      <p class="text-xs text-gray-500 mb-3">Các bill dưới đây có thể là tổng của nhiều ảnh giao dịch. Vui lòng kiểm tra lại.</p>
      <div class="scroll-container overflow-x-auto">
        <table class="min-w-full text-sm">
          <thead>
            <tr>
              <th class="p-2 text-left">Nội dung bill</th>
              <th class="p-2 text-left">Số tiền bill</th>
              <th class="p-2 text-left">Số tiền con tìm được</th>
              <th class="p-2 text-left">Tổng ảnh</th>
              <th class="p-2 text-left">Ảnh nghi vấn</th>
              <th class="p-2 text-left">Word</th>
            </tr>
          </thead>
          <tbody>
            ${suspicious.map((item, suspiciousIndex) => `
              <tr class="border-t">
                <td class="p-2 align-top">${item.bill.billName || '—'}</td>
                <td class="p-2 align-top font-semibold">${formatAmount(item.bill.price)}</td>
                <td class="p-2 align-top">${item.subAmounts.map(a => formatAmount(a)).join(', ')}</td>
                <td class="p-2 align-top">
                  <span class="${item.sumClose ? 'text-green-600 font-bold' : 'text-orange-500'}">${formatAmount(item.imgSum)}</span>
                  ${item.sumClose ? ' ✅' : ' ⚠️'}
                </td>
                <td class="p-2 align-top" style="max-width:350px;">
                  <div style="display:flex; gap:8px; overflow-x:auto; padding:4px 0;">
                    ${item.images.map(img => `
                      <div style="flex-shrink:0; text-align:center;">
                        <img src="${img.dataUrl}" class="img-popup" style="height:70px; border-radius:6px; cursor:pointer; border:2px solid #fbbf24;" title="Click để xem lớn" />
                        <div class="text-xs mt-1">${formatAmount(img.extractedAmount)}</div>
                      </div>
                    `).join('')}
                  </div>
                </td>
                <td class="p-2 align-top">
                  <button type="button" data-export-suspicious-row="${suspiciousIndex}" class="bg-yellow-600 hover:bg-yellow-700 text-white px-3 py-1 rounded text-xs whitespace-nowrap disabled:opacity-50" ${item.images.length ? '' : 'disabled'}>
                    Tải Word
                  </button>
                </td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    </div>
  `;

  // Init Tippy tooltip on the info icon
  const tooltipEl = document.getElementById('suspicious-tooltip-icon');
  if (tooltipEl && typeof tippy === 'function') {
    tippy(tooltipEl, {
      content: `<div style="text-align:left; line-height:1.6;">
        <b>Logic phát hiện bill ghép:</b><br/>
        1. Phân tích nội dung bill để tìm các <b>số tiền con</b> bên trong (ví dụ: "điện 1.578.982 + 462.300" hoặc "rửa xe 70k và grap 36k").<br/>
        2. So khớp các số tiền con với <b>số tiền OCR</b> từ các ảnh chưa được map (sai số ±1.000₫).<br/>
        3. Nếu có <b>≥ 2 ảnh</b> khớp với các số tiền con trong cùng 1 bill → đánh dấu nghi vấn ghép.<br/>
        4. Kiểm tra tổng số tiền ảnh có gần bằng tổng bill không (✅ nếu khớp).
      </div>`,
      allowHTML: true,
      theme: 'light-border',
      placement: 'right',
      maxWidth: 380,
      interactive: true,
    });
  }
}

// Call renderAllMappingUI after mapping or manual changes

// Export Word files from Step 4 review table
if (exportAllBtn) {
  exportAllBtn.addEventListener('click', exportAllReviewRows);
}

document.addEventListener('click', async (e) => {
  const reviewBtn = e.target.closest('[data-export-review-row]');
  if (reviewBtn) {
    const rowIndex = Number(reviewBtn.dataset.exportReviewRow);
    await withDownloadButtonLoading(reviewBtn, () => exportReviewRowDoc(rowIndex));
    return;
  }

  const suspiciousBtn = e.target.closest('[data-export-suspicious-row]');
  if (suspiciousBtn) {
    const suspiciousIndex = Number(suspiciousBtn.dataset.exportSuspiciousRow);
    await withDownloadButtonLoading(suspiciousBtn, () => exportSuspiciousBillDoc(suspiciousIndex));
  }
});

async function withDownloadButtonLoading(btn, action) {
  btn.disabled = true;
  const originalText = btn.textContent;
  btn.textContent = 'Đang tạo...';

  try {
    await action();
  } finally {
    btn.disabled = false;
    btn.textContent = originalText;
  }
}

async function exportReviewRowDoc(rowIndex, options = {}) {
  if (!window.docx) {
    alert('Thiếu thư viện docx.js. Hãy kiểm tra thẻ script CDN trong index.html.');
    return;
  }

  const img = images[rowIndex];
  if (!img) {
    alert('Không tìm thấy row cần xuất Word.');
    return;
  }

  const row = getReviewRowData(img, rowIndex);
  await exportWordDocFromReviewRow(row);

  if (exportLog && !options.silent) {
    exportLog.innerHTML += `<div class="text-green-600">✔ Đã tải Word cho row ${row.index}: ${escapeHtml(row.billName !== '—' ? row.billName : row.imageFileName)}</div>`;
  }
}

async function exportAllReviewRows() {
  if (!images.length) {
    alert('Chưa có dữ liệu ở bảng Bước 4 để xuất Word.');
    return;
  }
  if (!window.docx) {
    alert('Thiếu thư viện docx.js. Hãy kiểm tra thẻ script CDN trong index.html.');
    return;
  }

  exportAllBtn.disabled = true;
  const originalText = exportAllBtn.textContent;
  exportAllBtn.textContent = 'Đang tạo Word...';
  if (exportLog) exportLog.innerHTML = '<div>Đang tạo file Word từ từng row trong bảng Bước 4...</div>';

  let success = 0;
  for (let i = 0; i < images.length; i++) {
    try {
      await exportReviewRowDoc(i, { silent: true });
      success++;
      if (exportLog) exportLog.innerHTML += `<div class="text-green-600">✔ Row ${i + 1}</div>`;
      // Small delay helps browsers handle multiple downloads from one click.
      await sleep(150);
    } catch (err) {
      console.error(err);
      if (exportLog) exportLog.innerHTML += `<div class="text-red-600">✖ Row ${i + 1}: ${escapeHtml(err.message)}</div>`;
    }
  }

  exportAllBtn.disabled = false;
  exportAllBtn.textContent = originalText;
  if (exportLog) exportLog.innerHTML += `<div class="font-semibold mt-2">Hoàn tất: ${success}/${images.length} file Word.</div>`;
  alert(`Đã tạo ${success}/${images.length} file Word.`);
  setStepProgress(5);
}

async function exportSuspiciousBillDoc(suspiciousIndex) {
  if (!window.docx) {
    alert('Thiếu thư viện docx.js. Hãy kiểm tra thẻ script CDN trong index.html.');
    return;
  }

  const suspicious = findSuspiciousBillMappings();
  const item = suspicious[suspiciousIndex];
  if (!item) {
    alert('Không tìm thấy dòng nghi vấn cần xuất Word. Hãy chạy lại quy trình rồi thử lại.');
    return;
  }

  await exportWordDocFromSuspiciousBill(item, suspiciousIndex);

  if (exportLog) {
    exportLog.innerHTML += `<div class="text-green-600">✔ Đã tải Word nghi vấn ghép bill: ${escapeHtml(item.bill.billName || `Dòng ${suspiciousIndex + 1}`)}</div>`;
  }
  setStepProgress(5);
}

async function exportWordDocFromSuspiciousBill(item, suspiciousIndex) {
  const docx = window.docx;
  const titleText = item.bill.billName || `Nghi vấn ghép bill ${suspiciousIndex + 1}`;

  const children = [
    new docx.Paragraph({
      children: [new docx.TextRun({ text: `Nghi vấn ghép bill - ${String(titleText)}`, bold: true, size: 32 })],
      spacing: { after: 240 }
    }),
    new docx.Table({
      width: { size: 100, type: docx.WidthType.PERCENTAGE },
      rows: [
        createDocxTableRow(docx, 'Nội dung bill', item.bill.billName || '—'),
        createDocxTableRow(docx, 'Họ và tên', item.bill.fullName || '—'),
        createDocxTableRow(docx, 'Loại quỹ', item.bill.fundType || '—'),
        createDocxTableRow(docx, 'Ngày', item.bill.date || '—'),
        createDocxTableRow(docx, 'Số tiền bill', `${formatAmount(item.bill.price)} VND`),
        createDocxTableRow(docx, 'Số tiền con tìm được', item.subAmounts.map(a => formatAmount(a)).join(', ') || '—'),
        createDocxTableRow(docx, 'Tổng tiền ảnh nghi vấn', `${formatAmount(item.imgSum)} VND`),
        createDocxTableRow(docx, 'Đánh giá tổng', item.sumClose ? 'Tổng ảnh gần khớp bill' : 'Tổng ảnh chưa khớp hoàn toàn'),
        createDocxTableRow(docx, 'Số ảnh nghi vấn', item.images.length)
      ]
    }),
    new docx.Paragraph({ text: '', spacing: { after: 180 } }),
    new docx.Paragraph({
      children: [new docx.TextRun({ text: 'Danh sách ảnh nghi vấn', bold: true, size: 24 })],
      spacing: { before: 180, after: 120 }
    })
  ];

  for (let i = 0; i < item.images.length; i++) {
    const img = item.images[i];
    children.push(
      new docx.Paragraph({
        children: [new docx.TextRun({ text: `Ảnh ${i + 1}: ${img.file?.name || `anh-nghi-van-${i + 1}`} - ${formatAmount(img.extractedAmount)} VND`, bold: true })],
        spacing: { before: 180, after: 100 }
      })
    );

    if (!img.dataUrl) {
      children.push(new docx.Paragraph({ children: [new docx.TextRun({ text: '[Ảnh này chưa có dữ liệu]' })] }));
      continue;
    }

    try {
      let dataUrl = img.dataUrl;
      if (dataUrl.startsWith('data:image/webp')) {
        dataUrl = await webpToPng(dataUrl);
      }

      const imgBuffer = await fetch(dataUrl).then(r => r.arrayBuffer());
      const transformation = await getImageTransformation(dataUrl, 520, 520);
      children.push(
        new docx.Paragraph({
          children: [
            new docx.ImageRun({
              data: imgBuffer,
              transformation,
              type: getDocxImageType(dataUrl)
            })
          ]
        })
      );
    } catch (err) {
      console.error('Lỗi khi thêm ảnh nghi vấn vào Word:', err);
      children.push(new docx.Paragraph({ children: [new docx.TextRun({ text: '[Không thể nhúng ảnh này]' })] }));
    }
  }

  const doc = new docx.Document({
    sections: [{ children }]
  });

  const blob = await docx.Packer.toBlob(doc);
  downloadBlob(blob, getSuspiciousDocFileName(item, suspiciousIndex));
}


async function exportWordDocFromReviewRow(row) {
  const docx = window.docx;
  const titleText = row.billName !== '—' ? row.billName : `Row ${row.index} - ${row.imageFileName}`;

  const children = [
    new docx.Paragraph({
      children: [new docx.TextRun({ text: String(titleText), bold: true, size: 32 })],
      spacing: { after: 240 }
    }),
    new docx.Table({
      width: { size: 100, type: docx.WidthType.PERCENTAGE },
      rows: [
        createDocxTableRow(docx, 'Row', row.index),
        createDocxTableRow(docx, 'Tên file ảnh', row.imageFileName),
        createDocxTableRow(docx, 'Số tiền trong ảnh', row.extractedAmount),
        createDocxTableRow(docx, 'Loại quỹ', row.fundType),
        createDocxTableRow(docx, 'Họ và tên', row.fullName),
        createDocxTableRow(docx, 'Nội dung', row.billName),
        createDocxTableRow(docx, 'Ngày', row.date),
        createDocxTableRow(docx, 'Kiểm tra tên', row.nameCheckText),
        createDocxTableRow(docx, 'Trạng thái', row.statusText),
        createDocxTableRow(docx, 'Chi tiết', row.detailText || '—')
      ]
    }),
    new docx.Paragraph({ text: '', spacing: { after: 180 } }),
    new docx.Paragraph({
      children: [new docx.TextRun({ text: 'Ảnh giao dịch', bold: true, size: 24 })],
      spacing: { before: 180, after: 120 }
    })
  ];

  if (row.img.dataUrl) {
    try {
      let dataUrl = row.img.dataUrl;
      if (dataUrl.startsWith('data:image/webp')) {
        dataUrl = await webpToPng(dataUrl);
      }

      const imgBuffer = await fetch(dataUrl).then(r => r.arrayBuffer());
      const transformation = await getImageTransformation(dataUrl, 520, 650);
      children.push(
        new docx.Paragraph({
          children: [
            new docx.ImageRun({
              data: imgBuffer,
              transformation,
              type: getDocxImageType(dataUrl)
            })
          ]
        })
      );
    } catch (err) {
      console.error('Lỗi khi thêm ảnh vào Word:', err);
      children.push(new docx.Paragraph({ children: [new docx.TextRun({ text: '[Không thể nhúng ảnh]' })] }));
    }
  } else {
    children.push(new docx.Paragraph({ children: [new docx.TextRun({ text: '[Row này chưa có ảnh]' })] }));
  }

  const doc = new docx.Document({
    sections: [{ children }]
  });

  const blob = await docx.Packer.toBlob(doc);
  downloadBlob(blob, getReviewDocFileName(row));
}

function createDocxTableRow(docx, label, value) {
  return new docx.TableRow({
    children: [
      new docx.TableCell({
        width: { size: 30, type: docx.WidthType.PERCENTAGE },
        children: [new docx.Paragraph({ children: [new docx.TextRun({ text: String(label), bold: true })] })]
      }),
      new docx.TableCell({
        width: { size: 70, type: docx.WidthType.PERCENTAGE },
        children: [new docx.Paragraph({ children: [new docx.TextRun({ text: String(value ?? '—') })] })]
      })
    ]
  });
}

function getDocxImageType(dataUrl) {
  if (dataUrl.startsWith('data:image/png')) return 'png';
  if (dataUrl.startsWith('data:image/gif')) return 'gif';
  if (dataUrl.startsWith('data:image/bmp')) return 'bmp';
  return 'jpg';
}

async function getImageTransformation(dataUrl, maxWidth, maxHeight) {
  return new Promise((resolve) => {
    const image = new window.Image();
    image.onload = () => {
      const ratio = Math.min(maxWidth / image.width, maxHeight / image.height, 1);
      resolve({
        width: Math.round(image.width * ratio),
        height: Math.round(image.height * ratio)
      });
    };
    image.onerror = () => resolve({ width: 400, height: 250 });
    image.src = dataUrl;
  });
}

function getReviewDocFileName(row) {
  const base = row.billName !== '—' ? row.billName : row.imageFileName || `review-row-${row.index}`;
  return `${String(row.index).padStart(2, '0')}_${sanitizeFileName(base)}.docx`;
}

function getSuspiciousDocFileName(item, suspiciousIndex) {
  const base = item.bill.billName || `nghi-van-ghep-bill-${suspiciousIndex + 1}`;
  return `nghi-van-${String(suspiciousIndex + 1).padStart(2, '0')}_${sanitizeFileName(base)}.docx`;
}

function sanitizeFileName(value) {
  return String(value || 'file')
    .normalize('NFC')
    .replace(/[\/:*?"<>|]/g, '-')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 120) || 'file';
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function uploadToDrive(inv, img) {
  try {
    exportLog.innerHTML += `
      <div>Đang tạo file: ${inv.billName}</div>
    `;

    const payload = {
      billName: inv.billName,
      price: inv.price,
      date: inv.date,
      imageBase64: img.dataUrl
    };

    const res = await fetch(
      'https://script.google.com/macros/s/AKfycbzcXQHVfGAV1plYvrrpxnnPaJjGSfoHXPhUvraciMKo313tV7S-zrDAQv4zR9kJFgc/exec',
      {
        method: 'POST',
        body: JSON.stringify(payload)
      }
    );

    const data = await res.json();

    if (!data.success) {
      throw new Error(data.error);
    }

    exportLog.innerHTML += `
      <div class="text-green-600">
        ✔
        <a
          href="${data.url}"
          target="_blank"
          class="underline"
        >
          ${inv.billName}
        </a>
      </div>
    `;

  } catch (err) {
    console.error(err);

    exportLog.innerHTML += `
      <div class="text-red-600">
        ✖ ${inv.billName}: ${err.message}
      </div>
    `;
  }
}

// Hàm chuyển webp sang png (canvas)
async function webpToPng(dataUrl) {
  return new Promise((resolve) => {
    const img = new window.Image();
    img.onload = function() {
      const canvas = document.createElement('canvas');
      canvas.width = img.width;
      canvas.height = img.height;
      const ctx = canvas.getContext('2d');
      ctx.drawImage(img, 0, 0);
      resolve(canvas.toDataURL('image/png'));
    };
    img.src = dataUrl;
  });
}

// --- CSS for sticky headers and badges ---
(function injectStickyTableHeaderCSS() {
  const style = document.createElement('style');
  style.innerHTML = `
    .review-table-with-sticky thead th, .min-w-full thead th {
      position: sticky;
      top: 0;
      background: #f9fafb;
      z-index: 2;
    }
    .badge {
      display: inline-block;
      padding: 2px 8px;
      border-radius: 12px;
      font-size: 12px;
      font-weight: 600;
      margin-right: 4px;
      vertical-align: middle;
    }
    .badge-matched { background: #bbf7d0; color: #166534; }
    .badge-name-mismatch { background: #fef08a; color: #a16207; }
    .badge-unmatched { background: #fecaca; color: #b91c1c; }
    .badge-duplicate { background: #fca5a5; color: #7f1d1d; }
    .badge-pending { background: #e0e7ff; color: #3730a3; }
  `;
  document.head.appendChild(style);
})();

// Add badge style for suspicious
(function injectSuspiciousBadgeCSS() {
  const style = document.createElement('style');
  style.innerHTML = `
    .badge-suspicious { background: #fbbf24; color: #92400e; padding: 2px 8px; border-radius: 12px; font-size: 12px; font-weight: 600; }
  `;
  document.head.appendChild(style);
})();
