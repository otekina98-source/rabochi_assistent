const state = {
    orders: [],
    excludedIds: [],
    manualExcluded: new Set(),
    priceMap: {},
    mapping: new Map(),
    tableData: [],
    pdfFile: null,
    previewScale: 1,
    pageRotation: 0,
    visualW: 0,
    visualH: 0,
    fontBytes: null,
    fontSource: null,
    fontkitSource: null,
    textConfig: { fontSize: 5, color: '#000000' }
};

let previewFontFamily = 'Arial, sans-serif';
let lastOutBytes = null;

// Защита: не даём случайно закрыть/обновить страницу во время обработки
let processingActive = false;
window.addEventListener('beforeunload', (e) => {
    if (processingActive) {
        e.preventDefault();
        e.returnValue = '';
    }
});

const COMPRESS_PRESETS = {
    quality: { scale: 2,   quality: 0.9 },
    balance: { scale: 1.6, quality: 0.75 },
    min:     { scale: 1.3, quality: 0.6 }
};
const TARGET_BYTES = 70 * 1024 * 1024;
const OCR_SCALE = 1.5;
const OCR_LANG = 'eng';

// ===== Утилиты =====
function normalizeHeader(str) {
    if (str === null || str === undefined) return '';
    return String(str)
        .replace(/[\u00A0\u2007\u202F\u2009\u200A\u205F]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
}

function normalizeSku(value) {
    if (value === null || value === undefined) return '';
    let s = String(value);
    if (s.endsWith('.0')) s = s.slice(0, -2);
    return s.replace(/[\s\u00A0\u2007\u202F\u2009\u200A\u205F\u3000]/g, '').toLowerCase();
}

function normalizeAngle(a) {
    const r = ((Number(a) || 0) % 360 + 360) % 360;
    return [0, 90, 180, 270].includes(r) ? r : 0;
}

function loadScript(src) {
    return new Promise((resolve, reject) => {
        const s = document.createElement('script');
        s.src = src;
        s.onload = resolve;
        s.onerror = reject;
        document.head.appendChild(s);
    });
}

// ===== Поиск ТОЛЬКО номеров из таблицы (даже внутри слипшихся чисел) =====
function findTableOrderId(text, tableIds) {
    if (!text || !tableIds.length) return null;
    const compact = text.replace(/\s+/g, '');
    for (const id of tableIds) {
        if (id && id.length >= 5 && compact.includes(id)) return id;
    }
    return null;
}

// ===== Позиционирование текста — ТОЧНО как в рабочем коде (add_name_and_rename.py) =====
function pythonPlacement(rot, mw, mh) {
    if (rot === 90)  return { x: mw - 2,  y: 0 };
    if (rot === 270) return { x: 5,      y: mh - 1 };
    if (rot === 180) return { x: mw - 10, y: mh - 1 };
    return { x: 5, y: 1 };
}

function pageToVisual(rot, mw, mh, px, py) {
    switch (rot) {
        case 90:  return { x: py, y: px };
        case 180: return { x: mw - px, y: py };
        case 270: return { x: mh - py, y: px };
        default:  return { x: px, y: mh - py };
    }
}

function pageDeltaToVisual(rot, dx, dy) {
    switch (rot) {
        case 90:  return { x: dy, y: dx };
        case 180: return { x: -dx, y: dy };
        case 270: return { x: -dy, y: dx };
        default:  return { x: dx, y: -dy };
    }
}

// ===== Навигация и зоны загрузки =====
function goToStep(stepNum) {
    document.querySelectorAll('.step-container').forEach(el => el.classList.remove('active'));
    document.getElementById(`step${stepNum}`).classList.add('active');
    document.getElementById('globalProgress').style.width = `${(stepNum / 5) * 100}%`;
    if (stepNum === 3 && state.orders.length) rebuildTable();
    if (stepNum === 4 && state.pdfFile) initPdfPreview();
}

function setupDropZone(zoneId, inputId, handler) {
    const zone = document.getElementById(zoneId);
    zone.addEventListener('click', () => {
        const inp = document.getElementById(inputId);
        if (inp) inp.click();
    });
    zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('dragover'); });
    zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
    zone.addEventListener('drop', (e) => {
        e.preventDefault();
        zone.classList.remove('dragover');
        if (e.dataTransfer.files.length) handler(e.dataTransfer.files[0]);
    });
    zone.addEventListener('change', (e) => {
        if (e.target.files && e.target.files.length) handler(e.target.files[0]);
    });
}

const ordersZoneHTML = document.getElementById('ordersDropZone').innerHTML;
const priceZoneHTML = document.getElementById('priceDropZone').innerHTML;
const pdfZoneHTML = document.getElementById('pdfDropZone').innerHTML;

function resetOrdersZone() {
    const zone = document.getElementById('ordersDropZone');
    zone.innerHTML = ordersZoneHTML;
    zone.classList.remove('d-none');
    document.getElementById('ordersStats').classList.add('d-none');
}

function resetPriceZone() {
    const zone = document.getElementById('priceDropZone');
    zone.innerHTML = priceZoneHTML;
    zone.classList.remove('d-none');
    document.getElementById('priceStats').classList.add('d-none');
    document.getElementById('btnMatch').disabled = true;
}

function resetPdfZone() {
    const zone = document.getElementById('pdfDropZone');
    zone.innerHTML = pdfZoneHTML;
    zone.classList.remove('d-none');
    document.getElementById('pdfPreviewArea').classList.add('d-none');
    document.getElementById('btnStartProcess').disabled = true;
    state.pdfFile = null;
}

// ===== STEP 1: заказы =====
setupDropZone('ordersDropZone', 'ordersInput', async (file) => {
    try {
        document.getElementById('ordersDropZone').innerHTML = '<div class="spinner-border"></div>';
        const wb = XLSX.read(new Uint8Array(await file.arrayBuffer()), { type: 'array' });
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' });
        if (!rows.length) throw new Error('Файл пустой');

        const headers = Object.keys(rows[0] || {});
        const findCol = (kw) => headers.find(h => kw.some(k => normalizeHeader(h).includes(k)));
        const orderCol = findCol(['номер заказа', 'order id', 'order', 'заказ', 'номер']);
        const skuCol = findCol(['sku', 'артикул', 'код товара', 'код']);
        const cargoCol = findCol(['грузоместа', 'грузоместо', 'места', 'boxes', 'cargo']);
        if (!orderCol || !skuCol) throw new Error('Не найдены колонки «Номер заказа» и «SKU»');

        state.orders = [];
        state.excludedIds = [];
        for (const row of rows) {
            let cargoVal = 1;
            if (cargoCol && row[cargoCol] !== '' && row[cargoCol] !== null) {
                cargoVal = parseFloat(String(row[cargoCol]).replace(',', '.'));
                if (isNaN(cargoVal)) cargoVal = 1;
            }
            state.orders.push({
                orderId: normalizeSku(row[orderCol]),
                sku: normalizeSku(row[skuCol]),
                rawOrderId: String(row[orderCol] || '').trim(),
                cargo: cargoVal
            });
        }

        document.getElementById('statTotal').textContent = rows.length;
        document.getElementById('statExcluded').textContent = 0;
        document.getElementById('statRemaining').textContent = state.orders.length;
        document.getElementById('ordersStats').classList.remove('d-none');
        document.getElementById('ordersDropZone').classList.add('d-none');
    } catch (err) {
        alert('Ошибка обработки файла заказов: ' + err.message);
        location.reload();
    }
});

// ===== STEP 2: прайс =====
setupDropZone('priceDropZone', 'priceInput', async (file) => {
    try {
        document.getElementById('priceDropZone').innerHTML = '<div class="spinner-border"></div>';
        const wb = XLSX.read(new Uint8Array(await file.arrayBuffer()), { type: 'array' });
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '', blankrows: false });

        state.priceMap = {};
        for (const row of rows) {
            if (!row || !row.length) continue;
            let rawSku = null;
            for (let c = 0; c < Math.min(row.length, 5); c++) {
                if (row[c] !== '' && row[c] !== null && row[c] !== undefined) { rawSku = row[c]; break; }
            }
            if (rawSku === null) continue;
            let name = '';
            for (let c = 0; c < row.length; c++) {
                const cell = String(row[c] || '').trim();
                if (cell && cell !== String(rawSku)) { name = cell; break; }
            }
            if (!name) continue;
            const sku = normalizeSku(rawSku);
            if (sku) state.priceMap[sku] = name;
        }

        document.getElementById('priceCount').textContent = Object.keys(state.priceMap).length;
        document.getElementById('priceStats').classList.remove('d-none');
        document.getElementById('btnMatch').disabled = false;
        document.getElementById('priceDropZone').classList.add('d-none');
    } catch (err) {
        alert('Ошибка обработки прайс-листа: ' + err.message);
        location.reload();
    }
});

// ===== STEP 3: сопоставление (ручные исключения применяются ЗДЕСЬ, один раз) =====
function parseManualExclusions() {
    const el = document.getElementById('manualExcludeInput');
    const set = new Set();
    if (!el) return set;
    const nums = el.value.match(/\d{8,15}/g) || [];
    nums.forEach(n => set.add(normalizeSku(n)));
    return set;
}

function sortByName(a, b) {
    if (a.productName && !b.productName) return -1;
    if (!a.productName && b.productName) return 1;
    if (a.productName && b.productName) {
        return a.productName.localeCompare(b.productName, 'ru', { sensitivity: 'base' });
    }
    return 0;
}

function rebuildTable() {
    state.manualExcluded = parseManualExclusions();
    const activeOrders = state.orders.filter(o => !state.manualExcluded.has(o.orderId));
    state.mapping = new Map();
    state.tableData = activeOrders.map(order => {
        const productName = state.priceMap[order.sku] || null;
        if (productName && order.orderId) state.mapping.set(order.orderId, productName);
        return { ...order, productName, status: productName ? 'FOUND' : 'NOT_FOUND' };
    });
    state.tableData.sort(sortByName);
    renderTable(state.tableData);
}

function matchAndShowTable() {
    rebuildTable();
    if (state.manualExcluded.size) {
        document.getElementById('tableInfo').textContent += ` · Вручную исключено: ${state.manualExcluded.size}`;
    }
    goToStep(3);
}

function renderTable(data) {
    document.querySelector('#resultTable tbody').innerHTML = data.map(row => `
        <tr class="${row.status === 'NOT_FOUND' ? 'table-danger' : ''}">
            <td>${row.rawOrderId}</td>
            <td><code>${row.sku}</code></td>
            <td>${row.productName || '<em class="text-muted">-</em>'}</td>
            <td>${row.status === 'FOUND' ? '✅' : '❌'}</td>
        </tr>
    `).join('');
    document.getElementById('tableInfo').textContent = `Показано ${data.length} записей`;
}

function filterTable(query) {
    const q = query.toLowerCase();
    const filtered = state.tableData.filter(r =>
        r.rawOrderId.toLowerCase().includes(q) ||
        r.sku.includes(q) ||
        (r.productName && r.productName.toLowerCase().includes(q))
    );
    filtered.sort(sortByName);
    renderTable(filtered);
}

function exportTable() {
    const rows = state.tableData.map(r => ({
        'Номер заказа': r.rawOrderId,
        'SKU': r.sku,
        'Название товара': r.productName || '',
        'Грузоместа': r.cargo || 1
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Сопоставление');
    XLSX.writeFile(wb, 'matching_result.xlsx');
}

// ===== STEP 4: превью =====
let clientPdfDoc = null;
let textCanvas = null;

setupDropZone('pdfDropZone', 'pdfInput', (file) => {
    state.pdfFile = file;
    document.getElementById('pdfDropZone').innerHTML = `<p class="mt-2 mb-0">📄 ${file.name}</p>`;
    initPdfPreview();
});

async function initPdfPreview() {
    if (!state.pdfFile) return;

    const container = document.getElementById('pdfPreviewArea');
    container.classList.remove('d-none');
    document.getElementById('pdfDropZone').classList.add('d-none');

    const arrayBuffer = await state.pdfFile.arrayBuffer();
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

    clientPdfDoc = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    const page = await clientPdfDoc.getPage(1);
    state.pageRotation = page.rotate || 0;

    const base = page.getViewport({ scale: 1 });
    state.visualW = base.width;
    state.visualH = base.height;

    let availW = container.clientWidth - 24;
    let availH = container.clientHeight - 24;
    if (availW <= 0) availW = 800;
    if (availH <= 0) availH = 600;

    const scale = Math.min(availW / base.width, availH / base.height);
    state.previewScale = scale;
    const viewport = page.getViewport({ scale });

    const canvas = document.getElementById('pdfCanvas');
    canvas.width = Math.floor(viewport.width);
    canvas.height = Math.floor(viewport.height);

    const wrapper = document.getElementById('previewWrapper');
    wrapper.style.width = canvas.width + 'px';
    wrapper.style.height = canvas.height + 'px';

    await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;

    if (!textCanvas) {
        textCanvas = document.createElement('canvas');
        textCanvas.id = 'textPreviewCanvas';
        textCanvas.style.cssText = 'position:absolute;left:0;top:0;pointer-events:none;z-index:5;';
        wrapper.appendChild(textCanvas);
    }
    textCanvas.width = canvas.width;
    textCanvas.height = canvas.height;

    document.getElementById('btnStartProcess').disabled = false;

    ensureCompressControls();
    redrawTextPreview();
}

function ensureCompressControls() {
    if (document.getElementById('compressRow')) return;
    const btn = document.getElementById('btnStartProcess');
    const wrap = document.createElement('div');
    wrap.id = 'compressRow';
    wrap.className = 'alert alert-info small py-2 mt-2';
    wrap.innerHTML = `
        <div class="form-check form-switch">
            <input class="form-check-input" type="checkbox" id="cfgCompress" checked>
            <label class="form-check-label" for="cfgCompress">Сжать файл</label>
        </div>
        <select class="form-select form-select-sm mt-1" id="cfgCompressLevel">
            <option value="quality" selected>Без потери качества</option>
            <option value="balance">Баланс (размер / качество)</option>
            <option value="min">Минимальный размер</option>
        </select>
        <div class="text-muted" style="margin-top:4px;">Выключить — файл соберётся в исходном виде (большой размер).</div>
    `;
    btn.parentNode.insertBefore(wrap, btn);
}

function redrawTextPreview() {
    if (!textCanvas) return;
    const ctx = textCanvas.getContext('2d');
    ctx.clearRect(0, 0, textCanvas.width, textCanvas.height);
    const cw = textCanvas.width, ch = textCanvas.height;
    if (!cw || !ch) return;

    const rot = normalizeAngle(state.pageRotation);
    const VW = state.visualW, VH = state.visualH;
    if (!VW || !VH) return;
    const mw = (rot === 90 || rot === 270) ? VH : VW;
    const mh = (rot === 90 || rot === 270) ? VW : VH;

    const p = pythonPlacement(rot, mw, mh);
    const v0 = pageToVisual(rot, mw, mh, p.x, p.y);
    const dv = pageDeltaToVisual(rot, 0, 1);
    const angle = Math.atan2(dv.y, dv.x);
    const s = state.previewScale;

    const first = state.tableData.find(r => r.productName);
    const sample = first ? first.productName : 'тут будет название наклейки';

    ctx.font = `${(Number(state.textConfig.fontSize) || 5) * s}px ${previewFontFamily}`;
    ctx.fillStyle = state.textConfig.color || '#000000';
    ctx.textBaseline = 'alphabetic';
    ctx.save();
    ctx.translate(v0.x * s, v0.y * s);
    ctx.rotate(angle);
    ctx.fillText(sample, 0, 0);
    ctx.restore();
}

const fsEl = document.getElementById('cfgFontSize');
if (fsEl) fsEl.addEventListener('change', (e) => { state.textConfig.fontSize = parseInt(e.target.value) || 5; redrawTextPreview(); });
const colEl = document.getElementById('cfgColor');
if (colEl) colEl.addEventListener('change', (e) => { state.textConfig.color = e.target.value; redrawTextPreview(); });

// ===== Шрифт и fontkit =====
function b64ToBytes(b64) {
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
}

function looksLikeTtf(bytes) {
    return bytes && bytes.length > 50000 &&
        bytes[0] === 0x00 && bytes[1] === 0x01 && bytes[2] === 0x00 && bytes[3] === 0x00;
}

async function loadFontBytes() {
    if (window.DEJAVU_FONT_B64) {
        try {
            const b = b64ToBytes(window.DEJAVU_FONT_B64);
            if (looksLikeTtf(b)) { state.fontSource = 'font-data.js'; return b; }
        } catch (e) { /* идём дальше */ }
    }
    const urls = [
        'DejaVuSans.ttf',
        'fonts/DejaVuSans.ttf',
        'https://cdn.jsdelivr.net/gh/dejavu-fonts/dejavu@master/ttf/DejaVuSans.ttf',
        'https://raw.githubusercontent.com/dejavu-fonts/dejavu/master/ttf/DejaVuSans.ttf'
    ];
    for (const url of urls) {
        try {
            const res = await fetch(url);
            if (res.ok) {
                const bytes = new Uint8Array(await res.arrayBuffer());
                if (looksLikeTtf(bytes)) { state.fontSource = url; return bytes; }
            }
        } catch (e) { /* пробуем следующий */ }
    }
    try {
        const res = await fetch('https://cdn.jsdelivr.net/npm/pdfmake@0.2.10/build/vfs_fonts.js');
        if (res.ok) {
            const text = await res.text();
            const m = text.match(/Roboto-Regular\.ttf['"]?\s*:\s*['"]([A-Za-z0-9+/=]+)['"]/);
            if (m) {
                const b = b64ToBytes(m[1]);
                if (looksLikeTtf(b)) { state.fontSource = 'pdfmake CDN'; return b; }
            }
        }
    } catch (e) { /* ignore */ }
    return null;
}

async function setupPreviewFont(bytes) {
    try {
        const face = new FontFace('LabelFont', bytes.slice().buffer);
        await face.load();
        document.fonts.add(face);
        previewFontFamily = 'LabelFont, Arial, sans-serif';
        redrawTextPreview();
    } catch (e) { /* превью останется на Arial */ }
}

async function getFontkit() {
    if (window.fontkit) { state.fontkitSource = 'тег в index.html'; return window.fontkit; }
    await new Promise((resolve) => {
        const s = document.createElement('script');
        s.src = 'fontkit.local.js';
        s.onload = () => resolve();
        s.onerror = () => resolve();
        document.head.appendChild(s);
    });
    if (window.fontkit) { state.fontkitSource = 'fontkit.local.js'; return window.fontkit; }
    try {
        const mod = await import('https://cdn.jsdelivr.net/npm/@pdf-lib/fontkit@1.1.2/+esm');
        state.fontkitSource = 'CDN import';
        return mod.default || mod;
    } catch (e) {
        return null;
    }
}

(async function initFont() {
    state.fontBytes = await loadFontBytes();
    if (state.fontBytes) await setupPreviewFont(state.fontBytes);
    await getFontkit();

    if (!state.fontBytes || !window.fontkit) {
        alert(
            '🔍 Диагностика загрузки:\n\n' +
            '• font-data.js (base64 шрифта): ' + (window.DEJAVU_FONT_B64 ? '✅ есть' : '❌ НЕТ') + '\n' +
            '• Шрифт: ' + (state.fontBytes ? '✅ загружен (' + state.fontSource + ')' : '❌ НЕ загружен') + '\n' +
            '• fontkit: ' + (window.fontkit ? '✅ есть (' + state.fontkitSource + ')' : '❌ НЕТ (fontkit.local.js отсутствует или не сработал)') + '\n\n' +
            'Шрифт нужен для вставки текста с кириллицей.'
        );
    }
})();

// ===== Сортировка этикеток по названию товара (только включённые в PDF) =====
function sortPagesByProductName(results) {
    const named = [];
    const rest = [];
    results.forEach((r, idx) => {
        if (!r || r.status !== 'OK') return;
        if (r.productName) named.push({ idx, name: String(r.productName) });
        else rest.push(idx);
    });
    named.sort((a, b) => {
        const cmp = a.name.localeCompare(b.name, 'ru', { sensitivity: 'base' });
        return cmp !== 0 ? cmp : a.idx - b.idx;
    });
    return named.map(x => x.idx).concat(rest);
}

// ===== STEP 5: обработка (ориентир — таблица после сопоставления) =====
async function startProcessing() {
    goToStep(5);
    processingActive = true;

    const compBtn = document.getElementById('downloadCompressedBtn');
    if (compBtn) compBtn.classList.add('d-none');
    const compStatus = document.getElementById('compressStatus');
    if (compStatus) compStatus.classList.add('d-none');

    try {
        const pdfBytes = new Uint8Array(await state.pdfFile.arrayBuffer());
        const compress = document.getElementById('cfgCompress').checked;
        const preset = COMPRESS_PRESETS[document.getElementById('cfgCompressLevel').value] || COMPRESS_PRESETS.quality;

        document.getElementById('processStatus').textContent = 'Загрузка PDF и шрифта...';
        const pdfDoc = await PDFLib.PDFDocument.load(pdfBytes);

        const fontBytes = state.fontBytes || await loadFontBytes();
        const fontkitLib = await getFontkit();
        if (!fontBytes || !fontkitLib) {
            alert('Не удалось загрузить шрифт/fontkit. Убедитесь, что font-data.js и fontkit.local.js находятся в папке программы.');
            processingActive = false;
            goToStep(4);
            return;
        }
        pdfDoc.registerFontkit(fontkitLib);
        const font = await pdfDoc.embedFont(fontBytes);

        const pdfjs = await pdfjsLib.getDocument({ data: pdfBytes.slice() }).promise;
        const totalPages = pdfjs.numPages;
        const rotations = new Array(totalPages).fill(0);
        const results = new Array(totalPages).fill(null);
        const pageOut = new Array(totalPages).fill(undefined);

        document.getElementById('processStatus').textContent = 'Загрузка языковых пакетов OCR...';
        const CPU = navigator.hardwareConcurrency || 4;
        const POOL = Math.max(2, Math.min(CPU - 2 > 0 ? CPU - 2 : 2, 8));
        const workers = [];
        for (let i = 0; i < POOL; i++) {
            const w = await Tesseract.createWorker(OCR_LANG);
            try { await w.setParameters({ tessedit_char_whitelist: '0123456789' }); } catch (e) { /* необязательно */ }
            workers.push(w);
        }

        // ЕДИНСТВЕННЫЙ ОРИЕНТИР — номера из таблицы после сопоставления
        const tableIds = state.tableData.map(o => o.orderId).filter(Boolean);
        const cfg = state.textConfig;
        const renderScale = compress ? Math.max(OCR_SCALE, preset.scale) : OCR_SCALE;
        let next = 0, done = 0;
        const startTime = Date.now();

        async function runWorker(worker) {
            const cv = document.createElement('canvas');
            const cropCv = document.createElement('canvas');
            while (next < totalPages) {
                const idx = next++;
                const page = await pdfjs.getPage(idx + 1);
                rotations[idx] = page.rotate || 0;
                const viewport = page.getViewport({ scale: renderScale });
                cv.width = viewport.width;
                cv.height = viewport.height;
                await page.render({ canvasContext: cv.getContext('2d'), viewport }).promise;

                // Распознаём верхнюю треть этикетки — номер заказа там
                let text = '';
                const cropPx = Math.round(cv.height * 0.35);
                cropCv.width = cv.width;
                cropCv.height = cropPx;
                cropCv.getContext('2d').drawImage(cv, 0, 0, cv.width, cropPx, 0, 0, cv.width, cropPx);
                const fast = await worker.recognize(cropCv);
                text = fast.data.text || '';
                let orderId = findTableOrderId(text, tableIds);

                if (!orderId) {
                    const full = await worker.recognize(cv);
                    text = full.data.text || '';
                    orderId = findTableOrderId(text, tableIds);
                }

                if (!orderId) {
                    // Номера из таблицы на наклейке нет → удаляем автоматом
                    results[idx] = { status: 'NOT_IN_TABLE', ocrText: text.substring(0, 200) };
                    pageOut[idx] = { skip: true };
                } else {
                    const productName = state.mapping.get(orderId) || null;
                    let orderCargo = 1;
                    const foundOrder = state.orders.find(o => o.orderId === orderId);
                    if (foundOrder && foundOrder.cargo !== undefined) orderCargo = foundOrder.cargo;

                    if (orderCargo > 1) {
                        results[idx] = { status: 'CARGO', orderId, productName, cargo: orderCargo };
                        pageOut[idx] = { skip: true };
                    } else {
                        if (productName) {
                            drawLabel(pdfDoc, idx, font, rotations[idx], productName, cfg);
                        }
                        results[idx] = { status: 'OK', orderId, productName, cargo: orderCargo };
                        pageOut[idx] = { ok: true };
                    }
                }

                done++;
                const elapsed = (Date.now() - startTime) / 1000;
                const speed = done / elapsed || 0.01;
                const percent = Math.round((done / totalPages) * 100);
                document.getElementById('processProgressBar').style.width = `${percent}%`;
                document.getElementById('processProgressBar').textContent = `${percent}%`;
                document.getElementById('processStatus').textContent =
                    `Страница ${done} из ${totalPages}. Осталось ~${Math.round((totalPages - done) / speed)} сек.`;
            }
        }

        await Promise.all(workers.map(w => runWorker(w)));
        workers.forEach(w => w.terminate());

        document.getElementById('processStatus').textContent = 'Сортировка по названию и сборка итогового PDF...';

        const pageOrder = sortPagesByProductName(results);

        let skippedCargo = 0;
        let removedNotInTable = 0;
        results.forEach(r => {
            if (!r) return;
            if (r.status === 'CARGO') skippedCargo++;
            if (r.status === 'NOT_IN_TABLE') removedNotInTable++;
        });

        const outDoc = await PDFLib.PDFDocument.create();
        for (const i of pageOrder) {
            const item = pageOut[i];
            if (!item || !item.ok) continue;
            const [copied] = await outDoc.copyPages(pdfDoc, [i]);
            outDoc.addPage(copied);
        }
        const outBytes = await outDoc.save({
            useObjectStreams: true,
            addDefaultPage: false,
            objectsPerTick: 50
        });

        lastOutBytes = outBytes;

        const sizeMB = (outBytes.length / 1024 / 1024).toFixed(1);
        const pdfBlob = new Blob([outBytes], { type: 'application/pdf' });
        document.getElementById('downloadPdfBtn').href = URL.createObjectURL(pdfBlob);
        document.getElementById('downloadPdfBtn').download = 'processed_labels.pdf';

        let ok = 0;
        results.forEach(r => { if (r && r.status === 'OK') ok++; });

        // Журнал
        const wb = XLSX.utils.book_new();
        const wsData = [['Страница', 'Статус', 'Заказ', 'Товар', 'Грузоместа', 'Распознанный текст']];
        results.forEach((r, idx) => {
            if (!r) { wsData.push([idx + 1, '⚠️ Не обработано', '-', '-', '-', '-']); return; }
            let status;
            if (r.status === 'CARGO') status = '⏭️ Пропущена (Грузоместа > 1)';
            else if (r.status === 'NOT_IN_TABLE') status = '🚫 Удалена (номера нет в таблице)';
            else status = r.productName ? '✅ Обработано' : '✅ Обработано (без названия)';
            wsData.push([idx + 1, status, r.orderId || '-', r.productName || '-', r.cargo || 1, r.ocrText || '-']);
        });
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(wsData), 'Журнал');

        // Отдельный лист: удалённые страницы (номера нет в таблице)
        const nfRows = [['Страница в исходном PDF', 'Распознанный текст']];
        results.forEach((r, idx) => {
            if (r && r.status === 'NOT_IN_TABLE') {
                nfRows.push([idx + 1, (r.ocrText || '').replace(/\r?\n/g, ' ')]);
            }
        });
        if (nfRows.length > 1) {
            XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(nfRows), 'Не найдены');
        }

        const logOut = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const logBlob = new Blob([logOut], { type: 'application/octet-stream' });
        document.getElementById('downloadLogBtn').href = URL.createObjectURL(logBlob);
        document.getElementById('downloadLogBtn').download = 'processing_log.xlsx';

        document.getElementById('processingView').classList.add('d-none');
        const resultView = document.getElementById('resultView');
        resultView.classList.remove('d-none');
        let statsEl = document.getElementById('resultStats');
        if (!statsEl) {
            statsEl = document.createElement('p');
            statsEl.id = 'resultStats';
            statsEl.className = 'text-muted';
            resultView.insertBefore(statsEl, resultView.querySelector('.d-grid'));
        }
        const removedPages = results.map((r, i) => (r && r.status === 'NOT_IN_TABLE') ? i + 1 : null).filter(v => v !== null);
        statsEl.innerHTML =
            `📦 Размер файла: <strong>${sizeMB} МБ</strong> · В PDF этикеток: <strong>${ok}</strong> · Отсортированы по названию` +
            (skippedCargo > 0 ? `<br>⏭️ Пропущено (грузомест > 1): <strong>${skippedCargo}</strong>` : '') +
            (removedNotInTable > 0 ? `<br>🚫 Удалено (номера нет в таблице): <strong>${removedNotInTable}</strong> — страницы ${removedPages.join(', ')}             
        processingActive = false;
    } catch (err) {
        console.error(err);
        processingActive = false;
        alert('Ошибка обработки: ' + err.message);
        goToStep(4);
    }
}

// ===== Сжатие готового файла =====
async function compressResult() {
    if (!lastOutBytes) { alert('Сначала обработайте файл.'); return; }
    const btn = document.getElementById('btnCompress');
    const status = document.getElementById('compressStatus');
    btn.disabled = true;
    status.classList.remove('d-none');
    status.textContent = '🗜️ Сжатие: подготовка...';
    const wasMB = (lastOutBytes.length / 1024 / 1024).toFixed(1);
    try {
        const bytes = await compressPdfBytes(lastOutBytes, (d, t) => {
            status.textContent = `🗜️ Сжатие: страница ${d} из ${t}...`;
        });
        const sizeMB = (bytes.length / 1024 / 1024).toFixed(1);
        const blob = new Blob([bytes], { type: 'application/pdf' });
        const a = document.getElementById('downloadCompressedBtn');
        a.href = URL.createObjectURL(blob);
        a.download = 'processed_labels_compressed.pdf';
        a.classList.remove('d-none');
        status.textContent = `✅ Сжато: было ${wasMB} МБ → стало ${sizeMB} МБ. Текст остался доступен для поиска (Ctrl+F).`;
    } catch (err) {
        console.error(err);
        status.textContent = '❌ Ошибка сжатия: ' + err.message;
    }
    btn.disabled = false;
}

async function compressPdfBytes(bytes, onProgress) {
    const pdfjs = await pdfjsLib.getDocument({ data: bytes.slice() }).promise;
    const totalPages = pdfjs.numPages;
    const outDoc = await PDFLib.PDFDocument.create();
    const fontBytes = state.fontBytes || await loadFontBytes();
    const fontkitLib = await getFontkit();
    if (!fontBytes || !fontkitLib) throw new Error('Не удалось загрузить шрифт для сжатия.');
    outDoc.registerFontkit(fontkitLib);
    const font = await outDoc.embedFont(fontBytes);

    const live = { scale: 2, quality: 0.85 };
    let encodedBytes = 0, encodedCount = 0;
    const pageOut = new Array(totalPages).fill(undefined);
    let assembleIdx = 0;

    async function flush() {
        while (assembleIdx < totalPages && pageOut[assembleIdx] !== undefined) {
            const it = pageOut[assembleIdx];
            pageOut[assembleIdx] = null;
            assembleIdx++;
            if (!it) continue;
            const p = outDoc.addPage([it.w, it.h]);
            const img = await outDoc.embedJpg(it.jpeg);
            p.drawImage(img, { x: 0, y: 0, width: it.w, height: it.h });
            for (const t of it.texts) {
                try {
                    p.drawText(t.str, { x: t.x, y: t.y, size: t.size, font: font, opacity: 0, rotate: PDFLib.degrees(t.angle) });
                } catch (e) { /* пропускаем проблемные глифы */ }
            }
        }
    }

    const CPU = navigator.hardwareConcurrency || 4;
    const POOL = Math.max(2, Math.min(CPU - 2 > 0 ? CPU - 2 : 2, 6));
    let next = 0, done = 0;

    async function worker() {
        const cv = document.createElement('canvas');
        while (next < totalPages) {
            const idx = next++;
            const page = await pdfjs.getPage(idx + 1);
            const vp0 = page.getViewport({ scale: 1 });
            const visW = vp0.width, visH = vp0.height;
            const vp = page.getViewport({ scale: live.scale });
            cv.width = vp.width;
            cv.height = vp.height;
            await page.render({ canvasContext: cv.getContext('2d'), viewport: vp }).promise;
            const blob = await new Promise(res => cv.toBlob(res, 'image/jpeg', live.quality));
            const jpeg = new Uint8Array(await blob.arrayBuffer());

            encodedBytes += jpeg.length;
            encodedCount++;
            const projected = (encodedBytes / encodedCount) * totalPages;
            if (projected > TARGET_BYTES) {
                live.scale = Math.max(1.2, live.scale - 0.2);
                live.quality = Math.max(0.5, live.quality - 0.05);
            }

            const texts = [];
            try {
                const tc = await page.getTextContent();
                const tr = vp.transform;
                for (const item of tc.items) {
                    const str = item.str;
                    if (!str || !str.trim()) continue;
                    const it = item.transform;
                    const m0 = tr[0] * it[0] + tr[2] * it[1];
                    const m1 = tr[1] * it[0] + tr[3] * it[1];
                    const m2 = tr[0] * it[2] + tr[2] * it[3];
                    const m3 = tr[1] * it[2] + tr[3] * it[3];
                    const m4 = tr[0] * it[4] + tr[2] * it[5] + tr[4];
                    const m5 = tr[1] * it[4] + tr[3] * it[5] + tr[5];
                    const size = Math.min(Math.max(Math.hypot(m2, m3) || 8, 4), 14);
                    const angle = Math.round(-Math.atan2(m1, m0) * 180 / Math.PI);
                    texts.push({ str: str, x: m4, y: visH - m5, size: size, angle: angle });
                }
            } catch (e) { /* страница без текста */ }

            pageOut[idx] = { jpeg: jpeg, w: visW, h: visH, texts: texts };
            await flush();
            done++;
            if (onProgress) onProgress(done, totalPages);
        }
    }

    await Promise.all(Array.from({ length: POOL }, () => worker()));
    await flush();
    return outDoc.save({ useObjectStreams: true });
}

// ===== Вставка текста — ТОЧНО как в рабочем коде (add_name_and_rename.py) =====
function drawLabel(pdfDoc, pageIdx, font, rotation, productName, cfg) {
    const page = pdfDoc.getPage(pageIdx);
    const rot = normalizeAngle(rotation);
    const mb = page.getMediaBox();
    const mw = mb.width, mh = mb.height;
    const { x, y } = pythonPlacement(rot, mw, mh);
    page.drawText(productName, {
        x: x,
        y: y,
        size: Number(cfg.fontSize) || 5,
        font: font,
        color: parseColorLib(cfg.color),
        rotate: PDFLib.degrees(90)
    });
}

function parseColorLib(hex) {
    const h = String(hex || '#000000').replace('#', '');
    return PDFLib.rgb(
        parseInt(h.substring(0, 2), 16) / 255,
        parseInt(h.substring(2, 4), 16) / 255,
        parseInt(h.substring(4, 6), 16) / 255
    );
}