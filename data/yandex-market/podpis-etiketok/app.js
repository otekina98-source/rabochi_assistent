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
    fontBytes: null,
    fontSource: null,
    fontkitSource: null,
    textConfig: { x: 3.6, y: 91.7, width: 90, height: 7.1, fontSize: 5, lineSpacing: 1.2, align: 'left', color: '#000000', xShift: 0, yShift: 0 }
};

let previewFontFamily = 'Arial, sans-serif';

// Защита: не даём случайно закрыть/обновить страницу во время обработки
let processingActive = false;
window.addEventListener('beforeunload', (e) => {
    if (processingActive) {
        e.preventDefault();
        e.returnValue = '';
    }
});

// Уровни сжатия (scale = разрешение рендера, quality = качество JPEG)
const COMPRESS_PRESETS = {
    quality: { scale: 2,   quality: 0.9 },
    balance: { scale: 1.6, quality: 0.75 },
    min:     { scale: 1.3, quality: 0.6 }
};

// Масштаб рендера для OCR: 1.5 достаточно для крупных номеров
const OCR_SCALE = 1.5;
// Распознавание: только eng (быстрее), только цифры
const OCR_LANG = 'eng';

// ===== Нормализация и поиск =====
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

function extractOrderId(text, knownIds) {
    if (!text) return null;
    const compact = text.replace(/\s+/g, '');
    for (const id of knownIds) {
        if (id && id.length >= 6 && compact.includes(id)) return id;
    }
    const m1 = text.match(/(?:заказ|order)\s*[:\-№#]?\s*(\d{8,15})/i);
    if (m1) return m1[1];
    const longNums = text.match(/\d{10,15}/g);
    if (longNums) return longNums.sort((a, b) => b.length - a.length)[0];
    const shortNums = text.match(/\d{6,9}/g);
    if (shortNums) return shortNums.sort((a, b) => b.length - a.length)[0];
    return null;
}

function normalizeAngle(a) {
    const r = ((Number(a) || 0) % 360 + 360) % 360;
    return [0, 90, 180, 270].includes(r) ? r : 0;
}

function mapPoint(rot, mw, mh, vx, vy) {
    switch (rot) {
        case 90:  return { px: vy, py: vx };
        case 180: return { px: mw - vx, py: vy };
        case 270: return { px: vy, py: mh - vx };
        default:  return { px: vx, py: mh - vy };
    }
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

// ===== Навигация и зоны загрузки =====
function goToStep(stepNum) {
    document.querySelectorAll('.step-container').forEach(el => el.classList.remove('active'));
    document.getElementById(`step${stepNum}`).classList.add('active');
    document.getElementById('globalProgress').style.width = `${(stepNum / 5) * 100}%`;
    if (stepNum === 4 && state.pdfFile) initPdfPreview();
}

function setupDropZone(zoneId, inputId, handler) {
    const zone = document.getElementById(zoneId);
    const input = document.getElementById(inputId);
    zone.addEventListener('click', () => input.click());
    zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('dragover'); });
    zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
    zone.addEventListener('drop', (e) => {
        e.preventDefault();
        zone.classList.remove('dragover');
        if (e.dataTransfer.files.length) handler(e.dataTransfer.files[0]);
    });
    input.addEventListener('change', () => { if (input.files.length) handler(input.files[0]); });
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
            if (cargoCol && cargoVal > 1) {
                state.excludedIds.push(normalizeSku(row[orderCol]));
                continue;
            }
            state.orders.push({
                orderId: normalizeSku(row[orderCol]),
                sku: normalizeSku(row[skuCol]),
                rawOrderId: String(row[orderCol] || '').trim()
            });
        }

        document.getElementById('statTotal').textContent = rows.length;
        document.getElementById('statExcluded').textContent = state.excludedIds.length;
        document.getElementById('statRemaining').textContent = state.orders.length;
        document.getElementById('ordersStats').classList.remove('d-none');
        document.getElementById('ordersDropZone').classList.add('d-none');
    } catch (err) {
        alert('Ошибка обработки файла заказов: ' + err.message);
        location.reload();
    }
});

// ===== STEP 2: прайс (загрузка вручную) =====
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

// ===== STEP 3: сопоставление + ручные исключения =====
function parseManualExclusions() {
    const el = document.getElementById('manualExcludeInput');
    const set = new Set();
    if (!el) return set;
    el.value.split(/\r?\n/).forEach(line => {
        const id = normalizeSku(line.trim());
        if (id) set.add(id);
    });
    return set;
}

function matchAndShowTable() {
    state.manualExcluded = parseManualExclusions();
    const activeOrders = state.orders.filter(o => !state.manualExcluded.has(o.orderId));

    state.mapping = new Map();
    state.tableData = activeOrders.map(order => {
        const productName = state.priceMap[order.sku] || null;
        if (productName && order.orderId) state.mapping.set(order.orderId, productName);
        return { ...order, productName, status: productName ? 'FOUND' : 'NOT_FOUND' };
    });
    renderTable(state.tableData);
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
    renderTable(state.tableData.filter(r =>
        r.rawOrderId.toLowerCase().includes(q) ||
        r.sku.includes(q) ||
        (r.productName && r.productName.toLowerCase().includes(q))
    ));
}

function exportTable() {
    const rows = state.tableData.map(r => ({
        'Номер заказа': r.rawOrderId,
        'SKU': r.sku,
        'Название товара': r.productName || ''
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Сопоставление');
    XLSX.writeFile(wb, 'matching_result.xlsx');
}

// ===== STEP 4: превью и настройка зоны =====
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
        wrapper.insertBefore(textCanvas, document.getElementById('overlayBox'));
    }
    textCanvas.width = canvas.width;
    textCanvas.height = canvas.height;

    const overlay = document.getElementById('overlayBox');
    overlay.style.left = state.textConfig.x + '%';
    overlay.style.top = state.textConfig.y + '%';
    overlay.style.width = state.textConfig.width + '%';
    overlay.style.height = state.textConfig.height + '%';
    overlay.style.zIndex = '10';
    overlay.style.background = 'transparent';
    overlay.style.border = '1px dashed rgba(201,169,201,0.9)';

    document.getElementById('btnStartProcess').disabled = false;

    ensureFineTuneControls();
    ensureCompressControls();
    syncOverlayInputs();
    redrawTextPreview();
}

function ensureFineTuneControls() {
    if (document.getElementById('fineTuneRow')) return;
    const card = document.querySelector('#step4 .col-lg-4 .card-body');
    if (!card) return;
    const row = document.createElement('div');
    row.id = 'fineTuneRow';
    row.className = 'mt-2';
    row.innerHTML = `
        <label class="form-label small mb-1">Точная доводка (если текст уехал)</label>
        <div class="row g-2 mb-1">
            <div class="col">
                <small class="text-muted">← → горизонталь</small>
                <input type="range" class="form-range" id="cfgXShift" min="-10" max="10" step="0.2" value="0">
            </div>
            <div class="col">
                <small class="text-muted">↑ ↓ вертикаль</small>
                <input type="range" class="form-range" id="cfgYShift" min="-10" max="10" step="0.2" value="0">
            </div>
        </div>
    `;
    card.appendChild(row);
    document.getElementById('cfgXShift').addEventListener('input', (e) => { state.textConfig.xShift = parseFloat(e.target.value); redrawTextPreview(); });
    document.getElementById('cfgYShift').addEventListener('input', (e) => { state.textConfig.yShift = parseFloat(e.target.value); redrawTextPreview(); });
}

// Переключатель сжатия (создаётся автоматически, index.html не трогаем)
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
        <div class="text-muted" style="margin-top:4px;">Выключить — файл соберётся в исходном векторном виде (большой размер).</div>
    `;
    btn.parentNode.insertBefore(wrap, btn);
}

// ===== Рисование текста на канвасе (общее для превью и сжатия) =====
function paintLabel(ctx, cw, ch, scale, text, cfg) {
    const vx0 = (Number(cfg.x) / 100) * cw;
    const vy0 = (Number(cfg.y) / 100) * ch;
    const vw = (Number(cfg.width) / 100) * cw;
    const vh = (Number(cfg.height) / 100) * ch;
    const fontSizePx = (Number(cfg.fontSize) || 5) * scale;
    const lineHeightPx = fontSizePx * (Number(cfg.lineSpacing) || 1.2);
    const yShift = Number(cfg.yShift) || 0;
    const xShift = Number(cfg.xShift) || 0;

    ctx.font = `${fontSizePx}px ${previewFontFamily}`;
    ctx.fillStyle = cfg.color || '#000000';
    ctx.textBaseline = 'alphabetic';

    const lines = wrapCanvasText(ctx, text, Math.max(20, vw)).slice(0, 30);
    const N = lines.length;
    const bottomVis = Math.min(vy0 + vh, ch - 4 * scale) + (yShift / 100) * ch;

    lines.forEach((line, idx) => {
        const lineWidth = ctx.measureText(line).width;
        let dx = 0;
        if (cfg.align === 'center') dx = (vw - lineWidth) / 2;
        if (cfg.align === 'right') dx = vw - lineWidth;
        const baseY = bottomVis - fontSizePx * 0.2 - (N - 1 - idx) * lineHeightPx;
        ctx.fillText(line, Math.max(0, Math.min(cw, vx0 + dx + (xShift / 100) * cw)), baseY);
    });
}

function redrawTextPreview() {
    if (!textCanvas) return;
    const ctx = textCanvas.getContext('2d');
    ctx.clearRect(0, 0, textCanvas.width, textCanvas.height);
    const cw = textCanvas.width, ch = textCanvas.height;
    if (!cw || !ch) return;

    const first = state.tableData.find(r => r.productName);
    const sample = first ? first.productName : 'тут будет название наклейки';
    paintLabel(ctx, cw, ch, state.previewScale, sample, state.textConfig);
}

function wrapCanvasText(ctx, text, maxWidth) {
    const words = String(text).split(' ');
    const lines = [];
    let current = '';
    for (const word of words) {
        const test = current ? `${current} ${word}` : word;
        if (ctx.measureText(test).width > maxWidth && current) { lines.push(current); current = word; }
        else current = test;
    }
    if (current) lines.push(current);
    return lines;
}

const overlay = document.getElementById('overlayBox');
const wrapper = document.getElementById('previewWrapper');
let isDragging = false, startX, startY, startLeft, startTop;

overlay.addEventListener('mousedown', (e) => {
    isDragging = true;
    startX = e.clientX; startY = e.clientY;
    startLeft = overlay.offsetLeft; startTop = overlay.offsetTop;
    e.preventDefault();
});
window.addEventListener('mousemove', (e) => {
    if (!isDragging) return;
    let newLeft = startLeft + (e.clientX - startX);
    let newTop = startTop + (e.clientY - startY);
    newLeft = Math.max(0, Math.min(newLeft, wrapper.clientWidth - overlay.offsetWidth));
    newTop = Math.max(0, Math.min(newTop, wrapper.clientHeight - overlay.offsetHeight));
    overlay.style.left = `${newLeft}px`;
    overlay.style.top = `${newTop}px`;
    syncOverlayInputs();
    redrawTextPreview();
});
window.addEventListener('mouseup', () => isDragging = false);
new ResizeObserver(() => { syncOverlayInputs(); redrawTextPreview(); }).observe(overlay);

function syncOverlayInputs() {
    const cw = wrapper.clientWidth, ch = wrapper.clientHeight;
    if (!cw || !ch || cw < 10 || ch < 10) return;
    state.textConfig.x = (overlay.offsetLeft / cw) * 100;
    state.textConfig.y = (overlay.offsetTop / ch) * 100;
    state.textConfig.width = (overlay.offsetWidth / cw) * 100;
    state.textConfig.height = (overlay.offsetHeight / ch) * 100;
    document.getElementById('cfgX').value = state.textConfig.x.toFixed(1);
    document.getElementById('cfgY').value = state.textConfig.y.toFixed(1);
    document.getElementById('cfgW').value = state.textConfig.width.toFixed(1);
    document.getElementById('cfgH').value = state.textConfig.height.toFixed(1);
}

['cfgX', 'cfgY', 'cfgW', 'cfgH'].forEach(id => {
    document.getElementById(id).addEventListener('change', (e) => {
        const val = parseFloat(e.target.value) || 0;
        const prop = id.replace('cfg', '').toLowerCase();
        const key = prop === 'w' ? 'width' : prop === 'h' ? 'height' : prop;
        state.textConfig[key] = val;
        if (prop === 'x') overlay.style.left = `${val}%`;
        if (prop === 'y') overlay.style.top = `${val}%`;
        if (prop === 'w') overlay.style.width = `${val}%`;
        if (prop === 'h') overlay.style.height = `${val}%`;
        redrawTextPreview();
    });
});
document.getElementById('cfgFontSize').addEventListener('change', (e) => { state.textConfig.fontSize = parseInt(e.target.value) || 5; redrawTextPreview(); });
document.getElementById('cfgLineSpacing').addEventListener('change', (e) => { state.textConfig.lineSpacing = parseFloat(e.target.value) || 1.2; redrawTextPreview(); });
document.getElementById('cfgAlign').addEventListener('change', (e) => { state.textConfig.align = e.target.value; redrawTextPreview(); });
document.getElementById('cfgColor').addEventListener('change', (e) => { state.textConfig.color = e.target.value; redrawTextPreview(); });

// ===== Шрифт и fontkit =====
function b64ToBytes(b64) {
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
}

async function loadFontBytes() {
    if (window.DEJAVU_FONT_B64) {
        try {
            state.fontSource = 'font-data.js';
            return b64ToBytes(window.DEJAVU_FONT_B64);
        } catch (e) { /* идём дальше */ }
    }
    const urls = [
        'DejaVuSans.ttf',
        'fonts/DejaVuSans.ttf',
        'https://cdn.jsdelivr.net/gh/googlefonts/roboto@main/src/hinted/Roboto-Regular.ttf',
        'https://raw.githubusercontent.com/googlefonts/roboto/main/src/hinted/Roboto-Regular.ttf'
    ];
    for (const url of urls) {
        try {
            const res = await fetch(url);
            if (res.ok) {
                const bytes = new Uint8Array(await res.arrayBuffer());
                if (bytes.length > 10000) { state.fontSource = url; return bytes; }
            }
        } catch (e) { /* пробуем следующий */ }
    }
    try {
        const res = await fetch('https://cdn.jsdelivr.net/npm/pdfmake@0.2.10/build/vfs_fonts.js');
        if (res.ok) {
            const text = await res.text();
            const m = text.match(/Roboto-Regular\.ttf['"]?\s*:\s*['"]([A-Za-z0-9+/=]+)['"]/);
            if (m) { state.fontSource = 'pdfmake CDN'; return b64ToBytes(m[1]); }
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
            'Примечание: при включённом сжатии шрифт/fontkit не требуются.'
        );
    }
})();

// ===== Исключения =====
function isExcluded(orderId) {
    if (!orderId) return false;
    if (state.manualExcluded && state.manualExcluded.has(orderId)) return true;
    return state.excludedIds.some(e => e === orderId || (e && orderId.includes(e)));
}

// ===== STEP 5: обработка в браузере =====
async function startProcessing() {
    if (!(state.textConfig.width > 0) || !(state.textConfig.height > 0)) {
        Object.assign(state.textConfig, { x: 3.6, y: 91.7, width: 90, height: 7.1 });
    }
    goToStep(5);
    processingActive = true;

    try {
        const pdfBytes = new Uint8Array(await state.pdfFile.arrayBuffer());

        const compress = document.getElementById('cfgCompress').checked;
        const preset = COMPRESS_PRESETS[document.getElementById('cfgCompressLevel').value] || COMPRESS_PRESETS.quality;

        document.getElementById('processStatus').textContent = 'Загрузка PDF...';
        const pdfDoc = await PDFLib.PDFDocument.load(pdfBytes);

        // Шрифт и fontkit нужны только для векторного (несжатого) режима
        let font = null;
        if (!compress) {
            const fontBytes = state.fontBytes || await loadFontBytes();
            const fontkitLib = await getFontkit();
            if (!fontBytes || !fontkitLib) {
                alert('Не удалось загрузить шрифт/fontkit для векторного режима. Включите «Сжать файл» или запустите «node make-font.js».');
                processingActive = false;
                goToStep(4);
                return;
            }
            pdfDoc.registerFontkit(fontkitLib);
            font = await pdfDoc.embedFont(fontBytes);
        }

        const pdfjs = await pdfjsLib.getDocument({ data: pdfBytes.slice() }).promise;
        const totalPages = pdfjs.numPages;
        const rotations = new Array(totalPages).fill(0);
        const results = new Array(totalPages).fill(null);

        // Для режима сжатия: собираем PDF из JPEG по мере готовности
        let outDocCompress = null;
        let pageOut = null;
        let assembleIdx = 0;
        let skippedExcluded = 0;
        if (compress) {
            outDocCompress = await PDFLib.PDFDocument.create();
            pageOut = new Array(totalPages).fill(undefined);
        }

        async function flushPages() {
            while (assembleIdx < totalPages && pageOut[assembleIdx] !== undefined) {
                const item = pageOut[assembleIdx];
                if (item.skip) {
                    skippedExcluded++;
                } else {
                    const img = await outDocCompress.embedJpg(item.jpeg);
                    const p = outDocCompress.addPage([item.w, item.h]);
                    p.drawImage(img, { x: 0, y: 0, width: item.w, height: item.h });
                }
                pageOut[assembleIdx] = null;
                assembleIdx++;
            }
        }

        document.getElementById('processStatus').textContent = 'Загрузка языковых пакетов OCR...';
        const CPU = navigator.hardwareConcurrency || 4;
        const POOL = Math.max(2, Math.min(CPU - 2 > 0 ? CPU - 2 : 2, 8));
        const workers = [];
        for (let i = 0; i < POOL; i++) {
            const w = await Tesseract.createWorker(OCR_LANG);
            try { await w.setParameters({ tessedit_char_whitelist: '0123456789' }); } catch (e) { /* необязательно */ }
            workers.push(w);
        }

        const knownIds = Array.from(state.mapping.keys())
            .concat(state.excludedIds)
            .concat(Array.from(state.manualExcluded || []));
        const cfg = state.textConfig;
        // В режиме сжатия рендерим не мельче, чем нужно для выходного JPEG
        const renderScale = compress ? Math.max(OCR_SCALE, preset.scale) : OCR_SCALE;
        let next = 0, done = 0;
        const startTime = Date.now();

        async function runWorker(worker) {
            const cv = document.createElement('canvas');
            const cropCv = document.createElement('canvas');
            const outCv = compress ? document.createElement('canvas') : null;
            while (next < totalPages) {
                const idx = next++;
                const page = await pdfjs.getPage(idx + 1);
                rotations[idx] = page.rotate || 0;
                const vp1 = page.getViewport({ scale: 1 });
                const viewport = page.getViewport({ scale: renderScale });
                cv.width = viewport.width;
                cv.height = viewport.height;
                await page.render({ canvasContext: cv.getContext('2d'), viewport }).promise;

                // Распознаём верхнюю треть этикетки — номер заказа там
                let text = '';
                const cropH = Math.round(cv.height * (0.35 * (renderScale / OCR_SCALE) ) / (renderScale / OCR_SCALE)); // верхняя треть относительно высоты
                const cropPx = Math.round(cv.height * 0.35);
                cropCv.width = cv.width;
                cropCv.height = cropPx;
                cropCv.getContext('2d').drawImage(cv, 0, 0, cv.width, cropPx, 0, 0, cv.width, cropPx);
                const fast = await worker.recognize(cropCv);
                text = fast.data.text || '';
                let orderId = extractOrderId(text, knownIds);

                // Не нашлось сверху — распознаём страницу целиком
                if (!orderId) {
                    const full = await worker.recognize(cv);
                    text = full.data.text || '';
                    orderId = extractOrderId(text, knownIds);
                }

                const productName = orderId ? state.mapping.get(orderId) : null;
                const excluded = !!orderId && isExcluded(orderId);

                if (excluded) {
                    results[idx] = { status: 'EXCLUDED', orderId };
                    if (compress) { pageOut[idx] = { skip: true }; await flushPages(); }
                } else if (compress) {
                    const s = preset.scale;
                    const ow = Math.max(2, Math.round(vp1.width * s));
                    const oh = Math.max(2, Math.round(vp1.height * s));
                    outCv.width = ow;
                    outCv.height = oh;
                    const octx = outCv.getContext('2d');
                    octx.imageSmoothingEnabled = true;
                    octx.imageSmoothingQuality = 'high';
                    octx.drawImage(cv, 0, 0, cv.width, cv.height, 0, 0, ow, oh);
                    if (productName) paintLabel(octx, ow, oh, s, productName, cfg);
                    const blob = await new Promise(res => outCv.toBlob(res, 'image/jpeg', preset.quality));
                    const bytes = new Uint8Array(await blob.arrayBuffer());
                    pageOut[idx] = { jpeg: bytes, w: vp1.width, h: vp1.height };
                    results[idx] = { status: productName ? 'OK' : 'NOT_FOUND', orderId, productName };
                    await flushPages();
                } else {
                    if (productName) {
                        drawLabel(pdfDoc, idx, font, rotations[idx], productName, cfg);
                        results[idx] = { status: 'OK', orderId, productName };
                    } else {
                        results[idx] = { status: 'NOT_FOUND', orderId, ocrText: text.substring(0, 200) };
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

        document.getElementById('processStatus').textContent = 'Сборка итогового PDF...';

        let outBytes;
        if (compress) {
            await flushPages();
            outBytes = await outDocCompress.save();
        } else {
            const outDoc = await PDFLib.PDFDocument.create();
            for (let i = 0; i < totalPages; i++) {
                const r = results[i];
                if (r && r.orderId && isExcluded(r.orderId)) { skippedExcluded++; continue; }
                const [copied] = await outDoc.copyPages(pdfDoc, [i]);
                outDoc.addPage(copied);
            }
            outBytes = await outDoc.save();
        }

        const sizeMB = (outBytes.length / 1024 / 1024).toFixed(1);
        const pdfBlob = new Blob([outBytes], { type: 'application/pdf' });
        document.getElementById('downloadPdfBtn').href = URL.createObjectURL(pdfBlob);
        document.getElementById('downloadPdfBtn').download = 'processed_labels.pdf';

        let ok = 0, notFound = 0;
        results.forEach(r => {
            if (!r) notFound++;
            else if (r.status === 'OK') ok++;
            else if (r.status === 'NOT_FOUND') notFound++;
        });

        const wb = XLSX.utils.book_new();
        const wsData = [['Страница', 'Статус', 'Заказ', 'Товар', 'Распознанный текст']];
        results.forEach((r, idx) => {
            if (!r) { wsData.push([idx + 1, '⚠️ Не обработано', '-', '-', '-']); return; }
            let status;
            if (r.status === 'EXCLUDED') status = '🚫 Удалена (Грузоместа > 1 / вручную)';
            else if (r.status === 'OK') status = '✅ Обработано';
            else status = '❌ Не найден';
            wsData.push([idx + 1, status, r.orderId || '-', r.productName || '-', r.ocrText || '-']);
        });
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(wsData), 'Журнал');
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
        statsEl.innerHTML =
            `📦 Размер файла: <strong>${sizeMB} МБ</strong>` +
            (compress && Number(sizeMB) > 70 ? '<br>Больше 70 МБ — выберите «Минимальный размер» и повторите.' : '');

        processingActive = false;
    } catch (err) {
        console.error(err);
        processingActive = false;
        alert('Ошибка обработки: ' + err.message);
        goToStep(4);
    }
}

function drawLabel(pdfDoc, pageIdx, font, rotation, productName, cfg) {
    const page = pdfDoc.getPage(pageIdx);
    const rot = normalizeAngle(rotation);
    const mb = page.getMediaBox();
    const mw = mb.width, mh = mb.height;
    const VW = (rot === 90 || rot === 270) ? mh : mw;
    const VH = (rot === 90 || rot === 270) ? mw : mh;

    const fontSize = Number(cfg.fontSize) || 5;
    const lineHeight = fontSize * (Number(cfg.lineSpacing) || 1.2);
    const vx0 = (Number(cfg.x) / 100) * VW;
    const vy0 = (Number(cfg.y) / 100) * VH;
    const vw = (Number(cfg.width) / 100) * VW;
    const vh = (Number(cfg.height) / 100) * VH;
    const yShift = Number(cfg.yShift) || 0;
    const xShift = Number(cfg.xShift) || 0;

    const lines = wrapTextLib(font, productName, fontSize, vw).slice(0, 30);
    const N = lines.length;
    const color = parseColorLib(cfg.color);
    const bottomVis = Math.min(vy0 + vh, VH - 4) + (yShift / 100) * VH;

    lines.forEach((line, idx) => {
        const lineWidth = font.widthOfTextAtSize(line, fontSize);
        let dx = 0;
        if (cfg.align === 'center') dx = (vw - lineWidth) / 2;
        if (cfg.align === 'right') dx = vw - lineWidth;
        const pvx = Math.max(0, Math.min(VW, vx0 + dx + (xShift / 100) * VW));
        const pvy = bottomVis - fontSize * 0.2 - (N - 1 - idx) * lineHeight;
        const { px, py } = mapPoint(rot, mw, mh, pvx, pvy);
        page.drawText(line, {
            x: px, y: py, size: fontSize, font, color,
            rotate: PDFLib.degrees(rot)
        });
    });
}

function wrapTextLib(font, text, fontSize, maxWidth) {
    const words = String(text).split(' ');
    const lines = [];
    let current = '';
    for (const word of words) {
        const test = current ? `${current} ${word}` : word;
        if (font.widthOfTextAtSize(test, fontSize) > maxWidth && current) { lines.push(current); current = word; }
        else current = test;
    }
    if (current) lines.push(current);
    return lines;
}

function parseColorLib(hex) {
    const h = String(hex || '#000000').replace('#', '');
    return PDFLib.rgb(
        parseInt(h.substring(0, 2), 16) / 255,
        parseInt(h.substring(2, 4), 16) / 255,
        parseInt(h.substring(4, 6), 16) / 255
    );
}