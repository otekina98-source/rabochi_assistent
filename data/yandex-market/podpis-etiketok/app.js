/* ===== Подпись этикеток — поток: реестр 1С + этикетки PDF + прайс + список заказов ЛК =====
   Из реестра 1С берутся только заказы, где «Количество» = 1 (столбцы «Идентификатор МП» и «Количество»).
   В PDF остаются только этикетки этих заказов, остальные удаляются (список с причинами — внизу страницы).
   Артикулы (SKU) берутся из списка заказов ЛК, названия — из прайса.
   Этикетки подписываются и сортируются по алфавиту по наименованию. */

const PDFJS_WORKER_URL = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
const TESSERACT_URL = 'https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js';
const OCR_LANG = 'eng';

const state = {
    reestr: null,   // { name, orders: Map(orderId -> qty), qty1: Set(orderId) }
    labels: null,   // { name, file }
    price: null,    // { name, map: { sku -> name } }
    lk: null,       // { name, map: Map(orderId -> sku) }
    exportRows: [],
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
let clientPdfDoc = null;
let textCanvas = null;

// Защита: не даём случайно закрыть/обновить страницу во время обработки
let processingActive = false;
window.addEventListener('beforeunload', (e) => {
    if (processingActive) {
        e.preventDefault();
        e.returnValue = '';
    }
});

// ===== Утилиты =====
function normalizeHeader(str) {
    if (str === null || str === undefined) return '';
    return String(str)
        .replace(/[\u00A0\u2007\u202F\u2009\u200A\u205F]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
}

function normalizeOrderId(value) {
    if (value === null || value === undefined) return '';
    let s = String(value).trim();
    if (/^\d+\.0+$/.test(s)) s = s.replace(/\.0+$/, '');
    return s.replace(/[\s\u00A0\u2007\u202F\u2009\u200A\u205F\u3000]/g, '');
}

function normalizeSku(value) {
    if (value === null || value === undefined) return '';
    let s = String(value);
    if (s.endsWith('.0')) s = s.slice(0, -2);
    return s.replace(/[\s\u00A0\u2007\u202F\u2009\u200A\u205F\u3000]/g, '').toLowerCase();
}

function parseNum(value) {
    if (value === null || value === undefined || value === '') return null;
    const n = parseFloat(String(value).replace(/\s/g, '').replace(',', '.'));
    return isNaN(n) ? null : n;
}

function fmtQty(q) {
    return Number.isInteger(q) ? String(q) : String(Math.round(q * 100) / 100);
}

/* ===== Сжатие PDF без потери качества: удаление дубликатов объектов =====
   Если исходный файл содержит по копии одного и того же тяжёлого объекта
   на каждую этикетку (эмблема маркетплейса и т.п.), оставляем первый
   экземпляр и переносим все ссылки на него. Рендеринг не меняется ни на
   бит. Проверки типов — «утиные»: не зависят от экспорта классов pdf-lib. */

function pdfIsRef(v) {
    return !!(PDFLib.PDFRef && v instanceof PDFLib.PDFRef);
}
function pdfIsStream(o) {
    return !!(o && o.dict && o.contents instanceof Uint8Array);
}
function pdfIsDict(o) {
    return !!(o && typeof o.entries === 'function' && typeof o.set === 'function' && !(o.contents instanceof Uint8Array));
}
function pdfIsArray(o) {
    return !!(o && typeof o.size === 'function' && typeof o.get === 'function' &&
              typeof o.set === 'function' && typeof o.entries !== 'function');
}

function pdfHashBytes(u8) {
    // двойная свёртка: быстрая и практически без коллизий,
    // при совпадении хэша содержимое сверяется целиком
    let h1 = 0x811c9dc5, h2 = 0x01000193;
    for (let i = 0; i < u8.length; i++) {
        h1 = Math.imul((h1 ^ u8[i]) >>> 0, 16777619) >>> 0;
        h2 = (h2 + Math.imul(u8[i], i % 7 + 1)) >>> 0;
    }
    return u8.length + ':' + h1.toString(36) + h2.toString(36);
}

function pdfSameBytes(a, b) {
    if (!a || !b || a.length !== b.length) return false;
    for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
    return true;
}

/* Хэш объекта ПО СОДЕРЖИМОМУ (номера объектов не участвуют): два
   одинаковых объекта с разными номерами получают одинаковый хэш. */
function pdfContentHash(ctx, obj, memo, inProgress) {
    if (pdfIsRef(obj)) {
        const num = obj.number();
        if (memo.has(num)) return memo.get(num);
        if (inProgress.has(num)) return '~c';
        inProgress.add(num);
        const h = pdfContentHash(ctx, ctx.lookup(obj), memo, inProgress);
        inProgress.delete(num);
        memo.set(num, h);
        return h;
    }
    if (pdfIsStream(obj)) {
        const dictH = pdfContentHash(ctx, obj.dict, memo, inProgress);
        return 'S[' + pdfHashBytes(obj.contents) + '|' + dictH + ']';
    }
    if (pdfIsDict(obj)) {
        const parts = [];
        for (const [k, v] of obj.entries()) {
            parts.push(k.toString() + '=' + pdfContentHash(ctx, v, memo, inProgress));
        }
        parts.sort();
        return 'D{' + parts.join(';') + '}';
    }
    if (pdfIsArray(obj)) {
        const parts = [];
        for (let i = 0; i < obj.size(); i++) {
            parts.push(pdfContentHash(ctx, obj.get(i), memo, inProgress));
        }
        return 'A[' + parts.join(',') + ']';
    }
    try { return 'P' + obj.toString(); } catch (e) { return 'P?'; }
}

/* Возвращает число убранных дубликатов-потоков. */
function dedupePdfObjects(pdfDoc) {
    try {
        const ctx = pdfDoc.context;
        const objects = ctx.enumerateIndirectObjects();
        const byHash = new Map();    // хэш содержимого -> первая ссылка
        const dupToOrig = new Map(); // ссылка дубликата -> ссылка оригинала

        for (const [ref, obj] of objects) {
            if (!pdfIsStream(obj)) continue;            // тяжёлое — только потоки
            const bytes = obj.contents;
            if (!bytes || bytes.length < 1024) continue; // мелочь не ищем
            const h = pdfContentHash(ctx, obj, new Map(), new Set());
            const first = byHash.get(h);
            if (first === undefined) { byHash.set(h, ref); continue; }
            // страховка от коллизий: сверяем содержимое целиком
            const a = ctx.lookup(first);
            if (pdfIsStream(a) && pdfSameBytes(a.contents, bytes)) {
                dupToOrig.set(ref, first);
            }
        }
        if (!dupToOrig.size) return 0;

        // переносим все ссылки на дубликаты к оригиналам
        const visit = (obj, seen) => {
            if (!obj || typeof obj !== 'object' || seen.has(obj)) return;
            seen.add(obj);
            if (pdfIsDict(obj)) {
                for (const [k, v] of obj.entries()) {
                    if (pdfIsRef(v) && dupToOrig.has(v)) obj.set(k, dupToOrig.get(v));
                    else if (v && typeof v === 'object') visit(v, seen);
                }
            } else if (pdfIsArray(obj)) {
                for (let i = 0; i < obj.size(); i++) {
                    const v = obj.get(i);
                    if (pdfIsRef(v) && dupToOrig.has(v)) obj.set(i, dupToOrig.get(v));
                    else if (v && typeof v === 'object') visit(v, seen);
                }
            }
        };
        for (const [, obj] of objects) visit(obj, new Set());
        for (const ref of dupToOrig.keys()) ctx.delete(ref);
        return dupToOrig.size;
    } catch (e) {
        console.warn('Сжатие PDF не удалось (файл сохранён без сжатия):', e);
        return 0;
    }
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

function escapeHtml(s) {
    return String(s == null ? '' : s)
        .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

// ===== Позиционирование текста — как в рабочем коде (add_name_and_rename.py) =====
function pythonPlacement(rot, mw, mh) {
    if (rot === 90)  return { x: mw - 2, y: 0 };
    if (rot === 270) return { x: 5,     y: mh - 1 };
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

/* Текст всегда рисуется с поворотом 90° (базовая линия идёт в +y страницы,
   выносные вверх — в −x, вниз — в +x). Подбираем максимальный размер,
   при котором название целиком помещается в строку на этикетке. */
function fitTextLayout(measureWidth, text, mw, mh, x, y, maxSize) {
    const TOP_M = 2;           // не упираться в дальний край строки
    const ASC_RATIO = 0.95;    // высота букв над базовой линией (DejaVu Bold, с запасом)
    const DESC_RATIO = 0.27;   // выносные элементы вниз (у, р, д, щ…)
    const MIN_SIZE = 3;

    const run = Math.max(mh - y - TOP_M, 20);       // длина строки от точки старта
    const ascRoom = Math.max(x - 0.3, 2);           // место вверх до края
    const descRoom = Math.max(mw - x - 0.5, 1);     // место вниз до края

    let size = maxSize;
    const w = measureWidth(text, maxSize);
    if (w > run) size = maxSize * run / w;
    size = Math.min(size, ascRoom / ASC_RATIO);

    // если выносным вниз не хватает места — приподнимаем базовую линию (сдвиг в −x),
    // но не выше, чем позволяет запас сверху
    let shift = Math.max(0, DESC_RATIO * size - descRoom);
    shift = Math.min(shift, Math.max(0, x - 0.3 - ASC_RATIO * size));
    return { size: Math.max(Math.min(MIN_SIZE, maxSize), Math.floor(size * 10) / 10), shift };
}

// ===== Чтение Excel =====
/* Файлы из ЛК Яндекса содержат неверную запись <dimension> (например A1:N2 при 97 строках данных),
   из-за чего SheetJS видит только заголовки. Пересчитываем диапазон по фактическим ячейкам. */
function fixSheetRange(wb) {
    wb.SheetNames.forEach(name => {
        const ws = wb.Sheets[name];
        if (!ws) return;
        let minR = Infinity, minC = Infinity, maxR = -1, maxC = -1;
        Object.keys(ws).forEach(k => {
            if (k.charAt(0) === '!' || !Object.prototype.hasOwnProperty.call(ws, k)) return;
            const cell = ws[k];
            if (!cell || (cell.v === undefined && cell.t === undefined)) return;
            const { r, c } = XLSX.utils.decode_cell(k);
            if (r < minR) minR = r;
            if (c < minC) minC = c;
            if (r > maxR) maxR = r;
            if (c > maxC) maxC = c;
        });
        if (maxR >= 0) {
            ws['!ref'] = XLSX.utils.encode_range({ s: { r: minR, c: minC }, e: { r: maxR, c: maxC } });
        }
    });
}

async function readWorkbook(file) {
    const buf = new Uint8Array(await file.arrayBuffer());
    const wb = XLSX.read(buf, { type: 'array' });
    fixSheetRange(wb);
    return wb;
}

function firstSheetRows(wb) {
    return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '', blankrows: false });
}

/* Реестр 1С: любая структура файла, ищем строку заголовков со столбцами
   «Идентификатор МП» и «Количество», ниже — данные. */
function parseReestr(rows) {
    let hdr = -1, idCol = -1, qtyCol = -1;
    const limit = Math.min(rows.length, 40);
    for (let i = 0; i < limit; i++) {
        const cells = (rows[i] || []).map(normalizeHeader);
        let ic = cells.findIndex(c => c.includes('идентификатор') && c.replace(/[^а-яa-z ]/g, '').includes('мп'));
        const qc = cells.findIndex(c => c.replace(/[^а-яa-z]/g, '').startsWith('количество'));
        if (ic >= 0 && qc >= 0) { hdr = i; idCol = ic; qtyCol = qc; break; }
    }
    if (hdr < 0) {
        throw new Error('Не найдены столбцы «Идентификатор МП» и «Количество». Проверьте, что это реестр 1С с заказами.');
    }

    const orders = new Map();
    const rowCount = new Map(); // сколько строк с этим заказом (для поиска дублей)
    for (let i = hdr + 1; i < rows.length; i++) {
        const r = rows[i] || [];
        const id = normalizeOrderId(r[idCol]);
        if (!id || !/\d/.test(id)) continue;
        let qty = parseNum(r[qtyCol]);
        if (qty === null) qty = 1;
        rowCount.set(id, (rowCount.get(id) || 0) + 1);
        if (!orders.has(id)) orders.set(id, qty);
    }
    if (!orders.size) throw new Error('В реестре не найдено ни одного заказа (столбец «Идентификатор МП» пустой).');

    const dups = [];
    rowCount.forEach((count, id) => { if (count > 1) dups.push(id); });

    const qty1 = new Set();
    orders.forEach((q, id) => { if (q === 1) qty1.add(id); });
    return { orders, qty1, dups };
}

/* Список заказов из ЛК: «Номер заказа» + «SKU». В одном заказе может быть
   несколько товаров — сохраняем ВСЕ артикулы заказа (без повторов). */
function parseLk(rows) {
    let hdr = -1, orderCol = -1, skuCol = -1;
    const limit = Math.min(rows.length, 40);
    for (let i = 0; i < limit; i++) {
        const cells = (rows[i] || []).map(normalizeHeader);
        const oc = cells.findIndex(c => c === 'номер заказа');
        const sc = cells.findIndex(c => c === 'ваш sku' || c.includes('sku') || c.includes('артикул'));
        if (oc >= 0 && sc >= 0) { hdr = i; orderCol = oc; skuCol = sc; break; }
    }
    if (hdr < 0) {
        // запасной вариант: «Ваш номер заказа»
        for (let i = 0; i < limit; i++) {
            const cells = (rows[i] || []).map(normalizeHeader);
            const oc = cells.findIndex(c => c.includes('номер заказа'));
            const sc = cells.findIndex(c => c.includes('sku') || c.includes('артикул'));
            if (oc >= 0 && sc >= 0) { hdr = i; orderCol = oc; skuCol = sc; break; }
        }
    }
    if (hdr < 0) throw new Error('Не найдены столбцы «Номер заказа» и «SKU». Проверьте, что это список заказов из ЛК.');

    const map = new Map();
    let itemCount = 0;
    for (let i = hdr + 1; i < rows.length; i++) {
        const r = rows[i] || [];
        const id = normalizeOrderId(r[orderCol]);
        if (!id || !/\d/.test(id)) continue;
        const sku = normalizeSku(r[skuCol]);
        if (!sku) continue;
        let list = map.get(id);
        if (!list) { list = []; map.set(id, list); }
        if (!list.includes(sku)) { list.push(sku); itemCount++; }
    }
    if (!map.size) throw new Error('В списке заказов из ЛК не найдено ни одного заказа.');
    map.itemCount = itemCount;
    return map;
}

/* Прайс: первый непустой столбец — артикул, следующий непустой — название. */
function parsePrice(rows) {
    const map = {};
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
        if (sku) map[sku] = name;
    }
    return map;
}

// ===== Зоны загрузки =====
const ZONES = [
    { key: 'reestr', zone: 'reestrDropZone', input: 'reestrInput', stats: 'reestrStats', reset: 'reestrReset', handler: handleReestrFile },
    { key: 'labels', zone: 'labelsDropZone', input: 'labelsInput', stats: 'labelsStats', reset: 'labelsReset', handler: handleLabelsFile },
    { key: 'price',  zone: 'priceDropZone',  input: 'priceInput',  stats: 'priceStats',  reset: 'priceReset',  handler: handlePriceFile },
    { key: 'lk',     zone: 'lkDropZone',     input: 'lkInput',     stats: 'lkStats',     reset: 'lkReset',     handler: handleLkFile }
];
const zoneHTML = {};

function setupZones() {
    ZONES.forEach(z => {
        const zone = document.getElementById(z.zone);
        const inp = document.getElementById(z.input);
        zoneHTML[z.zone] = zone.innerHTML;
        zone.addEventListener('click', () => inp.click());
        zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('dragover'); });
        zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
        zone.addEventListener('drop', (e) => {
            e.preventDefault();
            zone.classList.remove('dragover');
            if (e.dataTransfer.files.length) z.handler(e.dataTransfer.files[0]);
        });
        inp.addEventListener('change', (e) => {
            if (e.target.files && e.target.files.length) z.handler(e.target.files[0]);
            inp.value = '';
        });
        document.getElementById(z.reset).addEventListener('click', () => resetZone(z.key));
    });
}

function setZoneBusy(zoneId) {
    document.getElementById(zoneId).innerHTML = '<div class="dz-spinner"></div><p class="dz-sub">Читаем файл...</p>';
}

function setZoneLoaded(zoneId, fileName) {
    document.getElementById(zoneId).innerHTML =
        `<div class="zone-loaded"><span class="ok-ico">✅</span><span>${escapeHtml(fileName)}</span></div>`;
}

function showStats(statsId, chipsHtml) {
    const el = document.getElementById(statsId);
    el.innerHTML = chipsHtml;
    el.classList.remove('hidden');
}

function resetZone(key) {
    const z = ZONES.find(x => x.key === key);
    state[key] = null;
    if (key === 'labels') {
        state.pdfFile = null;
        document.getElementById('previewBox').classList.add('hidden');
        clientPdfDoc = null;
        textCanvas = null;
    }
    const zone = document.getElementById(z.zone);
    zone.innerHTML = zoneHTML[z.zone];
    zone.classList.remove('hidden');
    document.getElementById(z.stats).classList.add('hidden');
    document.getElementById(z.reset).classList.add('hidden');
    refreshUi();
}

// ===== Обработчики файлов =====
async function handleReestrFile(file) {
    const z = ZONES.find(x => x.key === 'reestr');
    try {
        setZoneBusy(z.zone);
        const wb = await readWorkbook(file);
        const { orders, qty1, dups } = parseReestr(firstSheetRows(wb));
        state.reestr = { name: file.name, orders, qty1, dups };
        const over = orders.size - qty1.size;
        const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
        setZoneLoaded(z.zone, file.name);
        showStats(z.stats,
            `<span class="chip">Всего заказов: <b>${orders.size}</b></span>` +
            `<span class="chip chip-ok">С кол-вом 1: <b>${qty1.size}</b></span>` +
            (over ? `<span class="chip chip-warn">С кол-вом >1: <b>${over}</b></span>` : '') +
            (dups.length ? `<span class="chip chip-warn">⚠️ Есть дубли (${dups.length}): <b>${dups.slice(0, 5).map(esc).join(', ')}${dups.length > 5 ? '…' : ''}</b></span>` : '')
        );
        document.getElementById(z.reset).classList.remove('hidden');
        refreshUi();
    } catch (err) {
        zoneError(z, err);
    }
}

async function handleLabelsFile(file) {
    const z = ZONES.find(x => x.key === 'labels');
    try {
        setZoneBusy(z.zone);
        // быстрая проверка, что это читаемый PDF
        const buf = new Uint8Array(await file.arrayBuffer());
        pdfjsLib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER_URL;
        clientPdfDoc = await pdfjsLib.getDocument({ data: buf }).promise;
        state.labels = { name: file.name, file };
        state.pdfFile = file;
        setZoneLoaded(z.zone, file.name);
        showStats(z.stats, `<span class="chip">Этикеток (страниц PDF): <b>${clientPdfDoc.numPages}</b></span>`);
        document.getElementById(z.reset).classList.remove('hidden');
        refreshUi();
        initPdfPreview();
    } catch (err) {
        zoneError(z, err);
    }
}

async function handlePriceFile(file) {
    const z = ZONES.find(x => x.key === 'price');
    try {
        setZoneBusy(z.zone);
        const wb = await readWorkbook(file);
        const map = parsePrice(firstSheetRows(wb));
        if (!Object.keys(map).length) throw new Error('В прайсе не найдено ни одной строки «артикул + название».');
        state.price = { name: file.name, map };
        setZoneLoaded(z.zone, file.name);
        showStats(z.stats, `<span class="chip">Товаров в прайсе: <b>${Object.keys(map).length}</b></span>`);
        document.getElementById(z.reset).classList.remove('hidden');
        refreshUi();
    } catch (err) {
        zoneError(z, err);
    }
}

async function handleLkFile(file) {
    const z = ZONES.find(x => x.key === 'lk');
    try {
        setZoneBusy(z.zone);
        const wb = await readWorkbook(file);
        const map = parseLk(firstSheetRows(wb));
        state.lk = { name: file.name, map };
        setZoneLoaded(z.zone, file.name);
        showStats(z.stats,
            `<span class="chip">Заказов в списке: <b>${map.size}</b></span>` +
            `<span class="chip">Товаров: <b>${map.itemCount || map.size}</b></span>`
        );
        document.getElementById(z.reset).classList.remove('hidden');
        refreshUi();
    } catch (err) {
        zoneError(z, err);
    }
}

function zoneError(z, err) {
    const zone = document.getElementById(z.zone);
    zone.innerHTML = zoneHTML[z.zone];
    alert('❌ ' + err.message);
}

// ===== Общее состояние интерфейса =====
function allLoaded() {
    return !!(state.reestr && state.labels && state.price && state.lk);
}

function refreshUi() {
    const missing = [];
    if (!state.reestr) missing.push('реестр 1С');
    if (!state.labels) missing.push('этикетки');
    if (!state.price) missing.push('прайс');
    if (!state.lk) missing.push('список заказов из ЛК');

    // кнопка выгрузки доступна, когда есть данные заказов (PDF не нужен)
    const dataReady = !!(state.reestr && state.price && state.lk);
    state.exportRows = dataReady ? buildExportRows() : [];
    const exportBtn = document.getElementById('exportListBtn');
    const exportHint = document.getElementById('exportHint');
    exportBtn.disabled = !dataReady;
    exportHint.textContent = dataReady
        ? `Заказов: ${state.reestr.orders.size}, товаров: ${state.exportRows.length} — весь товар из заказов, по алфавиту. Жёлтым подсвечены заказы с количеством больше 1.`
        : 'Нужны: ' + missing.filter(m => m !== 'этикетки').join(', ');

    const btn = document.getElementById('btnStartProcess');
    const hint = document.getElementById('startHint');
    if (allLoaded()) {
        btn.disabled = false;
        hint.textContent = 'Все файлы загружены — можно запускать обработку.';
    } else {
        btn.disabled = true;
        hint.textContent = 'Ждём: ' + missing.join(', ') + '.';
    }
    redrawTextPreview();
}

// ===== Выгрузка «Заказы и наименования» =====
/* Все заказы из реестра 1С и весь их товар из списка ЛК: по одной строке
   на каждый товар (заказ с несколькими SKU даёт несколько строк).
   Название берём по артикулу из прайса, сортировка по алфавиту.
   Заказы с количеством больше 1 подсвечиваются жёлтым. */
function buildExportRows() {
    if (!state.reestr) return [];
    const rows = [];
    state.reestr.orders.forEach((qty, orderId) => {
        const skus = state.lk ? (state.lk.map.get(orderId) || []) : [];
        if (skus.length) {
            skus.forEach(sku => {
                const name = state.price ? (state.price.map[sku] || '') : '';
                rows.push({ orderId, sku, name, qty, sortKey: (name || '\uffff') + orderId + sku });
            });
        } else {
            // заказа нет в списке ЛК — строка с пустым артикулом, чтобы
            // заказ не потерялся из файла
            rows.push({ orderId, sku: '', name: '', qty, sortKey: '\uffff' + orderId });
        }
    });
    rows.sort((a, b) => a.sortKey.localeCompare(b.sortKey, 'ru'));
    return rows;
}

function exportOrdersFile() {
    const rows = state.exportRows && state.exportRows.length ? state.exportRows : buildExportRows();
    if (!rows.length) { alert('Сначала загрузите реестр 1С, прайс и список заказов из ЛК.'); return; }

    const ws = XLSX.utils.aoa_to_sheet([
        ['Заказ', 'SKU', 'Название товара'],
        ...rows.map(r => [r.orderId, r.sku, r.name])
    ]);
    ws['!cols'] = [{ wch: 16 }, { wch: 12 }, { wch: 95 }];

    // Подсветка номеров заказов с количеством больше 1 (нужна библиотека со стилями)
    const lib = window.XLSX_STYLED || XLSX;
    if (window.XLSX_STYLED) {
        ['A1', 'B1', 'C1'].forEach(a => {
            if (ws[a]) ws[a].s = { font: { bold: true }, fill: { fgColor: { rgb: 'E6E6FA' } } };
        });
        rows.forEach((r, i) => {
            if (r.qty > 1) {
                ws['A' + (i + 2)].s = {
                    fill: { fgColor: { rgb: 'FFEB9C' } },
                    font: { bold: true, color: { rgb: '9C6500' } }
                };
            }
        });
    }

    const wb = lib.utils.book_new();
    lib.utils.book_append_sheet(wb, ws, 'Заказы и наименования');
    lib.writeFile(wb, 'заказы_и_наименования.xlsx');
}

// ===== Превью первой этикетки =====
async function initPdfPreview() {
    if (!state.pdfFile) return;
    const box = document.getElementById('previewBox');
    box.classList.remove('hidden');

    const arrayBuffer = await state.pdfFile.arrayBuffer();
    pdfjsLib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER_URL;
    clientPdfDoc = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    const page = await clientPdfDoc.getPage(1);
    state.pageRotation = page.rotate || 0;

    const base = page.getViewport({ scale: 1 });
    state.visualW = base.width;
    state.visualH = base.height;

    const availW = Math.max(box.clientWidth - 30, 140);
    const availH = 330;
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
    redrawTextPreview();
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

    const first = (state.exportRows || []).find(r => r.name);
    const sample = first ? first.name : 'тут будет название';

    // тот же автоподбор размера, что и при финальной отрисовке (через метрики канваса)
    const maxSize = Number(state.textConfig.fontSize) || 5;
    ctx.font = `${maxSize * s}px ${previewFontFamily}`;
    const fit = fitTextLayout(
        (text, size) => { ctx.font = `${size * s}px ${previewFontFamily}`; return ctx.measureText(text).width / s; },
        sample, mw, mh, p.x, p.y, maxSize
    );
    const dsh = pageDeltaToVisual(rot, -fit.shift, 0);

    ctx.font = `${fit.size * s}px ${previewFontFamily}`;
    ctx.fillStyle = state.textConfig.color || '#000000';
    ctx.textBaseline = 'alphabetic';
    ctx.save();
    ctx.translate((v0.x + dsh.x) * s, (v0.y + dsh.y) * s);
    ctx.rotate(angle);
    ctx.fillText(sample, 0, 0);
    ctx.restore();
}

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
    // жирное начертание (лучше видно при мелком размере) — сначала встроенное, затем файлы/CDN
    if (window.DEJAVU_BOLD_FONT_B64) {
        try {
            const b = b64ToBytes(window.DEJAVU_BOLD_FONT_B64);
            if (looksLikeTtf(b)) { state.fontSource = 'font-data-bold.js'; return b; }
        } catch (e) { /* идём дальше */ }
    }
    if (window.DEJAVU_FONT_B64) {
        try {
            const b = b64ToBytes(window.DEJAVU_FONT_B64);
            if (looksLikeTtf(b)) { state.fontSource = 'font-data.js'; return b; }
        } catch (e) { /* идём дальше */ }
    }
    /* Жирное начертание — при мелком размере (5) читается заметно лучше;
       если жирного нет, откатываемся к обычному DejaVuSans. */
    const urls = [
        'DejaVuSans-Bold.ttf',
        'fonts/DejaVuSans-Bold.ttf',
        'https://cdn.jsdelivr.net/npm/dejavu-fonts-ttf@2.37.3/ttf/DejaVuSans-Bold.ttf',
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
    if (window.fontkit) return window.fontkit;
    await new Promise((resolve) => {
        const s = document.createElement('script');
        s.src = 'fontkit.local.js';
        s.onload = () => resolve();
        s.onerror = () => resolve();
        document.head.appendChild(s);
    });
    if (window.fontkit) return window.fontkit;
    try {
        const mod = await import('https://cdn.jsdelivr.net/npm/@pdf-lib/fontkit@1.1.2/+esm');
        return mod.default || mod;
    } catch (e) {
        return null;
    }
}

// ===== Сопоставление номера заказа на этикетке с реестром =====
const flexRegexCache = new Map();

function flexRegex(id) {
    let re = flexRegexCache.get(id);
    if (!re) {
        const escaped = id.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        re = new RegExp('(^|[^0-9])' + escaped.split('').join('\\s*') + '($|[^0-9])');
        flexRegexCache.set(id, re);
    }
    return re;
}

/* Поиск номеров реестра в тексте с допуском пробелов между цифрами (для OCR). */
function findReestrIdsFlexible(text) {
    const found = [];
    state.reestr.orders.forEach((_, id) => {
        if (flexRegex(id).test(text)) found.push(id);
    });
    return found;
}

function pickDisplayNumber(runs) {
    const preferred = [...runs].filter(d => /^\d{10,13}$/.test(d));
    if (preferred.length) return preferred.sort((a, b) => b.length - a.length)[0];
    return null;
}

function decideByReestrIds(ids) {
    const qty1 = ids.filter(id => state.reestr.qty1.has(id));
    if (qty1.length === 1) return { action: 'keep', orderId: qty1[0] };
    if (qty1.length > 1) {
        return { action: 'remove', reason: 'неоднозначно — несколько заказов с кол-вом 1: ' + qty1.join(', '), display: qty1[0] };
    }
    const qty = Math.max(...ids.map(id => state.reestr.orders.get(id)));
    return { action: 'remove', reason: `кол-во ${fmtQty(qty)} (больше 1)`, display: ids[0] };
}

/* Решение по тексту страницы этикетки.
   Возвращает { action: 'keep' | 'remove', ... } или { needOcr: true }. */
function decideByText(text, isOcr) {
    const txt = String(text || '');
    const runs = new Set((txt.match(/\d{8,}/g) || []).map(normalizeOrderId));
    const inReestr = [...runs].filter(id => state.reestr.orders.has(id));

    if (inReestr.length) return decideByReestrIds(inReestr);

    // цифры могли «разорваться» пробелами — пробуем гибкий поиск по номерам реестра
    const flex = findReestrIdsFlexible(txt);
    if (flex.length) return decideByReestrIds(flex);

    if (isOcr) {
        const display = pickDisplayNumber(runs);
        if (!display) return { action: 'remove', reason: 'номер не распознан', display: null };
        return { action: 'remove', reason: 'нет в реестре', display };
    }

    if (runs.size === 0) return { needOcr: true }; // текстовый слой пуст — вероятно, скан

    return { action: 'remove', reason: 'нет в реестре', display: pickDisplayNumber(runs) };
}

// ===== OCR (только для страниц без текстового слоя) =====
async function createOcrPool(size) {
    if (!window.Tesseract) await loadScript(TESSERACT_URL);
    const workers = [];
    for (let i = 0; i < size; i++) {
        const w = await Tesseract.createWorker(OCR_LANG);
        try { await w.setParameters({ tessedit_char_whitelist: '0123456789' }); } catch (e) { /* необязательно */ }
        workers.push(w);
    }
    return workers;
}

async function ocrAndDecide(worker, page) {
    const cv = document.createElement('canvas');
    const cropCv = document.createElement('canvas');
    const viewport = page.getViewport({ scale: 1.5 });
    cv.width = viewport.width;
    cv.height = viewport.height;
    await page.render({ canvasContext: cv.getContext('2d'), viewport }).promise;

    // номер заказа — в верхней части этикетки
    const cropPx = Math.round(cv.height * 0.35);
    cropCv.width = cv.width;
    cropCv.height = cropPx;
    cropCv.getContext('2d').drawImage(cv, 0, 0, cv.width, cropPx, 0, 0, cv.width, cropPx);

    let text = ((await worker.recognize(cropCv)).data.text) || '';
    let dec = decideByText(text, true);
    if (dec.action === 'remove' && (dec.reason === 'номер не распознан')) {
        text = ((await worker.recognize(cv)).data.text) || '';
        dec = decideByText(text, true);
    }
    return dec;
}

// ===== Сортировка этикеток по названию товара =====
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

// ===== Прогресс =====
function setProgress(done, total, note) {
    const percent = total ? Math.round((done / total) * 100) : 0;
    document.getElementById('progressFill').style.width = percent + '%';
    document.getElementById('progressPercent').textContent = percent + '%';
    document.getElementById('progressText').textContent = `Страница ${done} из ${total}`;
    document.getElementById('progressDetails').textContent = note || '';
}

// ===== Главная обработка =====
async function startProcessing() {
    if (!allLoaded()) return;
    processingActive = true;

    const progressSection = document.getElementById('progressSection');
    progressSection.classList.remove('hidden');
    progressSection.classList.add('processing');
    document.getElementById('resultSection').classList.add('hidden');
    document.getElementById('removedPlaque').classList.add('hidden');
    setProgress(0, 1, 'Читаем PDF и шрифт...');
    progressSection.scrollIntoView({ behavior: 'smooth', block: 'center' });

    try {
        const pdfBytes = new Uint8Array(await state.pdfFile.arrayBuffer());
        const pdfDoc = await PDFLib.PDFDocument.load(pdfBytes);

        const fontBytes = state.fontBytes || await loadFontBytes();
        const fontkitLib = await getFontkit();
        if (!fontBytes || !fontkitLib) {
            alert('Не удалось загрузить шрифт/fontkit. Убедитесь, что font-data.js и fontkit.local.js находятся в папке программы.');
            processingActive = false;
            progressSection.classList.add('hidden');
            return;
        }
        pdfDoc.registerFontkit(fontkitLib);
        const font = await pdfDoc.embedFont(fontBytes);

        pdfjsLib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER_URL;
        const pdfjs = await pdfjsLib.getDocument({ data: pdfBytes.slice() }).promise;
        const totalPages = pdfjs.numPages;
        const rotations = new Array(totalPages).fill(0);
        const texts = new Array(totalPages).fill('');
        const decisions = new Array(totalPages).fill(null);

        // 1) собираем текстовый слой всех страниц (быстро)
        for (let i = 0; i < totalPages; i++) {
            const page = await pdfjs.getPage(i + 1);
            rotations[i] = page.rotate || 0;
            try {
                const tc = await page.getTextContent();
                texts[i] = tc.items.map(it => it.str || '').join('\n');
            } catch (e) { texts[i] = ''; }
            if ((i % 10) === 0 || i === totalPages - 1) setProgress(i + 1, totalPages, 'Читаем текст этикеток...');
        }

        // 2) решения по текстовому слою; страницы без текста идут в OCR
        const needOcr = [];
        for (let i = 0; i < totalPages; i++) {
            const d = decideByText(texts[i], false);
            if (d.needOcr) { needOcr.push(i); decisions[i] = null; }
            else decisions[i] = d;
        }

        if (needOcr.length) {
            const CPU = navigator.hardwareConcurrency || 4;
            const POOL = Math.max(1, Math.min(CPU - 1 > 0 ? CPU - 1 : 1, 4));
            setProgress(0, needOcr.length, `Найдены страницы без текстового слоя: ${needOcr.length}. Запускаем распознавание (OCR)...`);
            const workers = await createOcrPool(POOL);
            let wIdx = 0, doneOcr = 0;
            const runOne = async (i) => {
                const worker = workers[wIdx++ % workers.length];
                const page = await pdfjs.getPage(i + 1);
                decisions[i] = await ocrAndDecide(worker, page);
                doneOcr++;
                setProgress(doneOcr, needOcr.length, 'Распознаём номера заказов (OCR) на сканированных страницах...');
            };
            // простой пул параллельной обработки
            const queue = needOcr.slice();
            await Promise.all(Array.from({ length: Math.min(POOL, queue.length) }, async () => {
                while (queue.length) {
                    const i = queue.shift();
                    await runOne(i);
                }
            }));
            workers.forEach(w => w.terminate());
        }

        // 3) применяем решения: подписываем оставшиеся этикетки
        const results = new Array(totalPages).fill(null);
        for (let i = 0; i < totalPages; i++) {
            const dec = decisions[i] || { action: 'remove', reason: 'номер не распознан', display: null };
            if (dec.action === 'keep') {
                const skus = (state.lk.map.get(dec.orderId) || []);
                const sku = skus[0] || '';
                const productName = sku ? (state.price.map[sku] || null) : null;
                if (productName) drawLabel(pdfDoc, i, font, rotations[i], productName, state.textConfig);
                results[i] = { status: 'OK', orderId: dec.orderId, sku, productName };
            } else {
                results[i] = { status: 'REMOVED', reason: dec.reason, orderId: dec.display || null, page: i + 1 };
            }
            if ((i % 10) === 0 || i === totalPages - 1) setProgress(i + 1, totalPages, 'Подписываем этикетки...');
        }

        // 4) сортировка по названию и сборка итогового PDF
        setProgress(totalPages, totalPages, 'Сортировка по названию и сборка PDF...');
        const pageOrder = sortPagesByProductName(results);
        const keep = pageOrder.filter(i => results[i] && results[i].status === 'OK');
        const outDoc = await PDFLib.PDFDocument.create();
        // ВАЖНО: все страницы копируем ОДНИМ вызовом — внутри одного
        // copyPages pdf-lib переиспользует общие объекты (эмблему,
        // шрифты), и файл не раздувается копией на каждую этикетку
        const copiedPages = await outDoc.copyPages(pdfDoc, keep);
        copiedPages.forEach(p => outDoc.addPage(p));

        // 4а) сжатие без потери качества: один и тот же тяжёлый объект
        // (например, векторная эмблема маркетплейса) копируется в каждую
        // этикетку и раздувает файл в десятки раз. Оставляем одну копию.
        setProgress(totalPages, totalPages, 'Сжимаем PDF: убираем дубликаты...');
        const dupesRemoved = dedupePdfObjects(outDoc);

        const outBytes = await outDoc.save({ useObjectStreams: true, addDefaultPage: false, objectsPerTick: 50 });

        // 5) скачивание PDF
        const sizeMB = (outBytes.length / 1024 / 1024).toFixed(1);
        const pdfBlob = new Blob([outBytes], { type: 'application/pdf' });
        document.getElementById('downloadPdfBtn').href = URL.createObjectURL(pdfBlob);
        document.getElementById('downloadPdfBtn').download = 'podpisannye_etiketki.pdf';

        // 6) журнал Excel
        const wb = XLSX.utils.book_new();
        const wsData = [['Страница', 'Статус', 'Заказ', 'SKU', 'Товар', 'Причина удаления']];
        results.forEach((r, idx) => {
            if (!r) { wsData.push([idx + 1, 'Не обработано', '', '', '', '']); return; }
            if (r.status === 'OK') {
                wsData.push([idx + 1, r.productName ? '✅ Оставлена, подписана' : '✅ Оставлена (без названия)', r.orderId, r.sku || '', r.productName || '', '']);
            } else {
                wsData.push([idx + 1, '🚫 Удалена', r.orderId || '', '', '', r.reason]);
            }
        });
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(wsData), 'Журнал');
        const logOut = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        document.getElementById('downloadLogBtn').href = URL.createObjectURL(new Blob([logOut], { type: 'application/octet-stream' }));
        document.getElementById('downloadLogBtn').download = 'podpis_zhurnal.xlsx';

        // 7) статистика
        let ok = 0, removedQty = 0, removedNotReg = 0, removedOther = 0, noName = 0;
        results.forEach(r => {
            if (!r) return;
            if (r.status === 'OK') {
                ok++;
                if (!r.productName) noName++;
            } else if (r.reason.startsWith('кол-во')) removedQty++;
            else if (r.reason === 'нет в реестре') removedNotReg++;
            else removedOther++;
        });

        document.getElementById('resultStats').innerHTML =
            `📦 Размер файла: <strong>${sizeMB} МБ</strong> · этикеток оставлено: <strong>${ok}</strong> из ${totalPages}, отсортированы по названию` +
            (dupesRemoved ? ` · 🗜 сжато без потери качества (убрано повторов: ${dupesRemoved})` : '') +
            (noName ? `<br>⚠️ Без названия (нет в списке ЛК или в прайсе): <strong>${noName}</strong> — подробности в таблице проверки` : '');
        document.getElementById('resultSection').classList.remove('hidden');

        // 8) плашка удалённых этикеток внизу страницы
        renderRemovedPlaque(results, { removedQty, removedNotReg, removedOther });

        setProgress(totalPages, totalPages, 'Готово!');
        progressSection.classList.remove('processing');
        processingActive = false;
        document.getElementById('removedPlaque').scrollIntoView({ behavior: 'smooth', block: 'center' });

    } catch (err) {
        console.error(err);
        processingActive = false;
        progressSection.classList.remove('processing');
        progressSection.classList.add('hidden');
        alert('Ошибка обработки: ' + err.message);
    }
}

function renderRemovedPlaque(results, counts) {
    // группы по причине: «кол-во больше 1», «нет в реестре», прочее
    const groups = new Map();
    results.forEach((r) => {
        if (!r || r.status !== 'REMOVED') return;
        const reason = r.reason.startsWith('кол-во') ? 'кол-во больше 1' : r.reason;
        if (!groups.has(reason)) groups.set(reason, { count: 0, numbers: new Map() });
        const g = groups.get(reason);
        g.count++;
        if (r.orderId) g.numbers.set(r.orderId, (g.numbers.get(r.orderId) || 0) + 1);
    });

    const order = ['кол-во больше 1', 'нет в реестре', 'номер не распознан'];
    const keys = [...groups.keys()].sort((a, b) => {
        const ia = order.indexOf(a), ib = order.indexOf(b);
        return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib);
    });

    const total = counts.removedQty + counts.removedNotReg + counts.removedOther;
    document.getElementById('removedCount').textContent = total + ' шт.';
    const box = document.getElementById('removedGroups');

    if (!total) {
        box.innerHTML = '<div class="plaque-empty">Ничего не удалено — все этикетки соответствуют заказам с количеством 1 ✨</div>';
    } else {
        box.innerHTML = keys.map(reason => {
            const g = groups.get(reason);
            const nums = [...g.numbers.entries()]
                .map(([n, c]) => `<span class="rg-num">${escapeHtml(n)}</span>${c > 1 ? ` <span class="rg-mult">×${c}</span>` : ''}`)
                .join(',  ');
            return `<div class="removed-group">` +
                `<span class="rg-label">${escapeHtml(reason)}: ${g.count}</span>` +
                (nums ? `<span class="rg-nums">${nums}</span>` : '') +
                `</div>`;
        }).join('');
    }
    document.getElementById('removedPlaque').classList.remove('hidden');
}

// ===== Вставка текста на этикетку =====
function drawLabel(pdfDoc, pageIdx, font, rotation, productName, cfg) {
    const page = pdfDoc.getPage(pageIdx);
    const rot = normalizeAngle(rotation);
    const mb = page.getMediaBox();
    const mw = mb.width, mh = mb.height;
    const { x, y } = pythonPlacement(rot, mw, mh);
    const maxSize = Number(cfg.fontSize) || 12;
    const fit = fitTextLayout(
        (text, size) => font.widthOfTextAtSize(text, size),
        productName, mw, mh, x, y, maxSize
    );
    page.drawText(productName, {
        x: x - fit.shift,
        y: y,
        size: fit.size,
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

// ===== Инициализация =====
document.addEventListener('DOMContentLoaded', () => {
    setupZones();

    document.getElementById('btnStartProcess').addEventListener('click', startProcessing);
    document.getElementById('exportListBtn').addEventListener('click', exportOrdersFile);

    document.getElementById('cfgFontSize').addEventListener('change', (e) => {
        let v = parseInt(e.target.value, 10);
        if (isNaN(v)) v = 5;
        v = Math.max(3, Math.min(20, v));
        e.target.value = v;
        state.textConfig.fontSize = v;
        redrawTextPreview();
    });
    document.getElementById('cfgColor').addEventListener('change', (e) => {
        state.textConfig.color = e.target.value;
        redrawTextPreview();
    });
});

(async function initFont() {
    state.fontBytes = await loadFontBytes();
    if (state.fontBytes) await setupPreviewFont(state.fontBytes);
    await getFontkit();
    if (!state.fontBytes || !window.fontkit) {
        console.warn('Шрифт или fontkit не загрузились:', {
            fontData: !!window.DEJAVU_FONT_B64,
            font: !!state.fontBytes,
            fontkit: !!window.fontkit
        });
    }
})();
