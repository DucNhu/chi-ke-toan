// Globals
let invoices = []; // {id, billName, price, date}
let images = [];   // {id, file, dataUrl, ocrText, extractedAmount, status, matchedInvoiceId, ocrProgress}

const excelInput = document.getElementById('excel-input');
const invoicePreview = document.getElementById('invoice-preview');
const excelError = document.getElementById('excel-error');
const imagesInput = document.getElementById('images-input');
const imagesList = document.getElementById('images-list');
const runMappingBtn = document.getElementById('run-mapping');
const mappingTable = document.getElementById('mapping-table');
const reviewTable = document.getElementById('review-table');
// const exportAllBtn = document.getElementById('export-all');
const exportLog = document.getElementById('export-log');
const progressBar = document.getElementById('progress-bar');
const exportTemplateBtn = document.getElementById('export-template');
const mappingSummary = document.getElementById('mapping-summary');
const unmatchedImagesSection = document.getElementById('unmatched-images-section');
const unmatchedBillsSection = document.getElementById('unmatched-bills-section');
const imagesProgressTitle = document.getElementById('images-progress-title');

function setStepProgress(step) {
  const pct = Math.min(100, Math.max(0, (step-1) * 25));
  progressBar.style.width = pct + '%';
}
setStepProgress(1);

// Helpers
function normalizeColName(s) {
  return String(s || '').trim().toLowerCase();
}
function normalizePriceToInt(s) {
  if (s == null) return null;
  const str = String(s).replace(/\s+/g, '').replace(/[.,]/g, '');
  const n = parseInt(str, 10);
  return Number.isNaN(n) ? null : n;
}

// Excel handling
excelInput.addEventListener('change', (e) => {
  const f = e.target.files[0];
  if (!f) return;
  const reader = new FileReader();
  reader.onload = function(evt) {
    try {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, {type: 'array'});
      const firstSheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[firstSheetName];
      const json = XLSX.utils.sheet_to_json(sheet, {defval: ''});
      if (!json || !json.length) {
        excelError.textContent = 'Excel trống hoặc không đọc được.';
        return;
      }
      // Find required columns: tên bill, đơn giá, ngày xuất hoá đơn
      const headerKeys = Object.keys(json[0]);
      const normalized = headerKeys.map(k => normalizeColName(k));
      const findCol = (variants) => {
        for (let v of variants) {
          const idx = normalized.indexOf(normalizeColName(v));
          if (idx >= 0) return headerKeys[idx];
        }
        return null;
      };
      const colBill = findCol(['tên bill', 'ten bill', 'name']);
      const colPrice = findCol(['đơn giá', 'don gia', 'price', 'gia']);
      const colDate = findCol(['ngày xuất hoá đơn', 'ngay xuat hoa don', 'date']);
      if (!colBill || !colPrice || !colDate) {
        excelError.textContent = 'Không tìm thấy các cột cần thiết. Vui lòng kiểm tra tên cột (tên bill, đơn giá, ngày xuất hoá đơn).';
        return;
      }
      excelError.textContent = '';
      invoices = json.map((row, i) => {
        return {
          id: i,
          billName: row[colBill],
          priceRaw: row[colPrice],
          price: normalizePriceToInt(row[colPrice]),
          date: row[colDate]
        };
      });
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
  let html = '<div class="overflow-x-auto"><table class="min-w-full text-sm"><thead><tr><th class="text-left p-2">#</th><th class="text-left p-2">Tên bill</th><th class="text-left p-2">Đơn giá</th><th class="text-left p-2">Ngày</th></tr></thead><tbody>';
  invoices.forEach(inv => {
    html += `<tr><td class="p-2">${inv.id+1}</td><td class="p-2">${inv.billName}</td><td class="p-2">${inv.priceRaw}</td><td class="p-2">${inv.date}</td></tr>`;
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
    return { id: idx, file: f, dataUrl: '', ocrText: '', extractedAmount: null, status: 'pending', matchedInvoiceId: null, ocrProgress: 0 };
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
      <img src="${img.dataUrl}" class="img-preview mb-2" />
      <div class="text-sm">Giá trong ảnh: ${img.extractedAmount != null ? img.extractedAmount : '—'}</div>
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
    img.status = extracted != null ? 'ocr' : 'unmatched';
    renderImagesList();
    renderAllMappingUI();
  } catch (err) {
    console.error('OCR error', err);
    img.status = 'unmatched';
    img.ocrProgress = 100; renderImagesList();
    renderAllMappingUI();
  }
}

function extractAmountFromText(text) {
  if (!text) return null;
  const patterns = [
    /(\d{1,3}(?:[.,]\d{3})+)\s*(?:đ|VND|vnđ|d\b)/gi,
    /(?:số tiền|amount|tổng tiền|thành tiền|total)[^\d]*(\d{1,3}(?:[.,]\d{3})+)/gi,
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
      // Only allow each invoice to be matched once (unless manually reassigned)
      const matched = invoices.find(inv => inv.price === img.extractedAmount && !usedInvoiceIds.has(inv.id));
      img.matchedInvoiceId = matched ? matched.id : null;
      img.status = matched ? 'matched' : 'unmatched';
      if (matched) usedInvoiceIds.add(matched.id);
    }
  });
  renderImagesList();
  renderReviewTable();
  renderAllMappingUI();
  setStepProgress(4);
});

function renderReviewTable() {
  if (!images.length) {
    reviewTable.innerHTML = '<div class="text-sm text-gray-500">Chưa có ảnh nào.</div>';
    renderAllMappingUI();
    return;
  }
  let html = '<div class="overflow-x-auto"><table class="min-w-full text-sm"><thead><tr><th class="p-2">#</th><th class="p-2">Ảnh</th><th class="p-2">Số tiền thanh toán của ảnh</th><th class="p-2">Tên bill</th><th class="p-2">Trạng thái <span title="Chưa ghép: Ảnh chưa được ghép với bill nào.\nKhông khớp: Ảnh đã nhận diện nhưng không tìm thấy bill phù hợp.\nKhớp: Ảnh đã ghép thành công với bill." style="cursor: help; color: #2563eb;">&#9432;</span></th></tr></thead><tbody>';
  images.forEach((img, idx) => {
    const matched = invoices.find(i => i.id === img.matchedInvoiceId);
    let statusText = '';
    if (img.status === 'matched') statusText = 'Khớp';
    else if (img.status === 'unmatched') statusText = 'Không khớp';
    else if (img.status === 'pending') statusText = 'Đang đợi phân tích';
    else statusText = '—';
    html += `<tr><td class="p-2">${idx+1}</td><td class="p-2"><img src="${img.dataUrl}" style="width:80px;height:50px;object-fit:cover;border-radius:6px" /></td><td class="p-2">${img.extractedAmount || '—'}</td><td class="p-2">${matched ? matched.billName : '—'}</td><td class="p-2">${statusText}</td></tr>`;
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
  if (!unmatchedImages.length) {
    unmatchedImagesSection.innerHTML = '';
    return;
  }
  unmatchedImagesSection.innerHTML = `
    <div class="card p-4 bg-white rounded shadow transition-all duration-300">
      <h2 class="font-medium mb-3">Ảnh chưa được map</h2>
      <div class="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4">
        ${unmatchedImages.map(img => `
          <div class="p-3 bg-gray-50 rounded shadow flex flex-col items-center">
            <img src="${img.dataUrl}" class="mb-2 rounded max-h-32 object-contain border" />
            <div class="text-xs text-gray-700 truncate w-full">${img.file?.name || ''}</div>
            <div class="text-sm mt-1">Giá trong ảnh: <span class="font-semibold">${img.extractedAmount != null ? img.extractedAmount : '—'}</span></div>
            <div class="text-xs mt-1 ${img.status === 'unmatched' ? 'text-red-600' : 'text-gray-500'}">
              ${img.status === 'unmatched' ? (img.ocrText ? 'Không khớp bill nào' : 'OCR lỗi hoặc không đọc được số tiền') : (img.status === 'pending' ? 'Đang nhận diện...' : 'Đã nhận diện')}
            </div>
          </div>
        `).join('')}
      </div>
    </div>
  `;
}

function renderUnmatchedBills() {
  const { unmatchedBills } = getMatchedAndUnmatched();
  if (!unmatchedBills.length) {
    unmatchedBillsSection.innerHTML = '';
    return;
  }
  unmatchedBillsSection.innerHTML = `
    <div class="card p-4 bg-white rounded shadow transition-all duration-300">
      <h2 class="font-medium mb-3">Bill chưa được map</h2>
      <div class="overflow-x-auto">
        <table class="min-w-full text-sm">
          <thead>
            <tr>
              <th class="p-2 text-left">Tên bill</th>
              <th class="p-2 text-left">Đơn giá</th>
              <th class="p-2 text-left">Ngày xuất hoá đơn</th>
            </tr>
          </thead>
          <tbody>
            ${unmatchedBills.map(inv => `
              <tr>
                <td class="p-2">${inv.billName}</td>
                <td class="p-2">${inv.priceRaw}</td>
                <td class="p-2">${inv.date}</td>
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
  renderUnmatchedImages();
  renderUnmatchedBills();
}

// Call renderAllMappingUI after mapping or manual changes

// Export
// exportAllBtn.addEventListener('click', async () => {
//   if (!invoices.length) { alert('Chưa có dữ liệu để xuất.'); return; }
//   const matchedImages = images.filter(img => img.status === 'matched' && img.matchedInvoiceId != null);
//   if (!matchedImages.length) { alert('Không có cặp nào đã ghép để xuất.'); return; }
//   for (const img of matchedImages) {
//     const inv = invoices.find(i => i.id === img.matchedInvoiceId);
//     await uploadToDrive(inv, img);
//   }
//   alert('Đã xuất xong tất cả file Word!');
// });

async function exportWordDoc(inv, img) {
  if (!window.docx) {
    alert('docx.js chưa sẵn sàng!'); return;
  }
  const docx = window.docx;

  const children = [
    new docx.Paragraph({ children: [ new docx.TextRun({ text: String(inv.billName), bold: true, size: 28 }) ] }),
    new docx.Paragraph({ children: [ new docx.TextRun({ text: `Đơn giá: ${inv.price} VND` }) ] }),
    new docx.Paragraph({ children: [ new docx.TextRun({ text: `Ngày xuất hoá đơn: ${inv.date}` }) ] }),
    new docx.Paragraph({ text: '' })
  ];

  if (img.dataUrl) {
    try {
      let dataUrl = img.dataUrl;
      if (dataUrl.startsWith('data:image/webp')) {
        dataUrl = await webpToPng(dataUrl);
      }
      // docx v7: ImageRun thay cho Media.addImage
      const imgBuffer = await fetch(dataUrl).then(r => r.arrayBuffer());
      const ext = dataUrl.startsWith('data:image/png') ? 'png' : 'jpg';
      children.push(
        new docx.Paragraph({
          children: [
            new docx.ImageRun({
              data: imgBuffer,
              transformation: { width: 400, height: 250 },
              type: ext
            })
          ]
        })
      );
    } catch (err) {
      console.error('Lỗi khi thêm ảnh vào Word:', err);
      children.push(new docx.Paragraph({ children: [ new docx.TextRun({ text: '[Không thể nhúng ảnh]' }) ] }));
    }
  }

  // docx v7: sections truyền vào constructor
  const doc = new docx.Document({
    sections: [{ children }]
  });

  const blob = await docx.Packer.toBlob(doc);
  const filename = `${inv.billName || 'invoice'}.docx`;
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
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

// Export template
exportTemplateBtn.addEventListener('click', () => {
  const template = [
    ['Tên bill', 'Đơn giá', 'Ngày xuất hóa đơn'],
    ['Ví dụ: Điện nước', '100000', '2023-10-01']
  ];
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(template);
  XLSX.utils.book_append_sheet(wb, ws, 'Template');
  const fileName = 'template_' + new Date().toISOString().slice(0, 10) + '.xlsx';
  XLSX.writeFile(wb, fileName);
});