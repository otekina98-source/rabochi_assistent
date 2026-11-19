/* Лист подбора и наклейки — ОЗОН
   Вход:  Прайс (Excel) + Лист подбора (PDF) + Наклейки (PDF)
   Выход: Лист подбора (Excel, по алфавиту) + Наклейки (PDF в том же порядке) */

pdfjsLib.GlobalWorkerOptions.workerSrc =
    'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

/* ============================ Состояние ============================ */

const state = {
    price: null,      // { map: Map(артикул -> наименование), name, count }
    lp: null,         // { rows: [{num, article, name}], name }
    ticket: null      // { buffer, pageOf: Map(номер -> [страницы]), name, pages }
};

let built = null;     // { rows, xlsxBlob, pdfBlob, missing, extra }
let showAllMode = false;

/* ============================ Служебные ============================ */

const NUM_RE = /^\d{5,10}-\d{3,5}-\d{1,3}$/;
const STICKER_GLUE_RE = /^(\d{5,10}-\d{3,5}-\d{1,3})\d{4}$/; // номер + 4-значный код этикетки (наклейки)
const LABEL_ID_RE = /^ii\d{8,16}$/i; // новый полный ID этикетки Ozon
const LABEL_ID_SEARCH_RE = /(^|[^a-zа-яё0-9])(ii\d{8,16})(?=$|[^a-zа-яё0-9])/ig;
const LP_GLUE_RE = /^(\d{1,4})(\d{8,10}-\d{3,5}-\d{1,3})$/;  // порядковый номер + номер отправления (лист подбора, после 999-й строки)

function normArt(v) {
    return String(v == null ? '' : v).replace(/[\s\u00A0\u202F]+/g, '').trim();
}

function escapeXml(s) {
    return String(s).replace(/[&<>"']/g, c => ({
        '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;'
    }[c]));
}

function setStatus(id, text, cls) {
    const el = document.getElementById(id);
    el.textContent = text;
    el.className = 'status' + (cls ? ' ' + cls : '');
}

function setProgress(pct, text, details) {
    document.getElementById('progressFill').style.width = pct + '%';
    document.getElementById('progressPercent').textContent = pct + '%';
    if (text != null) document.getElementById('progressText').textContent = text;
    if (details != null) document.getElementById('progressDetails').textContent = details;
}

function updateStartBtn() {
    document.getElementById('startBtn').disabled =
        !(state.price && state.lp && state.ticket);
}

/* ============================ Загрузка файлов ============================ */

function wireUpload(areaId, inputId, statusId, nameId, handler) {
    const area = document.getElementById(areaId);
    const input = document.getElementById(inputId);
    input.addEventListener('change', e => {
        if (e.target.files[0]) handler(e.target.files[0]);
        input.value = '';
    });
    area.addEventListener('dragover', e => { e.preventDefault(); area.classList.add('dragover'); });
    area.addEventListener('dragleave', () => area.classList.remove('dragover'));
    area.addEventListener('drop', e => {
        e.preventDefault();
        area.classList.remove('dragover');
        const file = e.dataTransfer.files[0];
        if (file) handler(file);
    });
}

/* ============================ Прайс (Excel) ============================ */

async function handlePrice(file) {
    document.getElementById('priceFileName').textContent = file.name;
    document.getElementById('priceUploadArea').classList.remove('loaded');
    setStatus('priceStatus', '⏳ Чтение файла...', '');
    try {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        if (!ws) throw new Error('в файле нет таблиц');
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
        const map = new Map();
        let empty = 0;
        for (const r of rows) {
            const art = normArt(r[0]);
            const name = String(r[1] == null ? '' : r[1]).trim();
            if (!art || !name) continue;
            if (/^(артикул|арт|код)$/i.test(art) && /наимен|товар|назв/i.test(name)) continue; // шапка
            if (!map.has(art)) map.set(art, name); else empty++;
        }
        if (map.size === 0) throw new Error('не найдены столбцы «артикул + наименование»');
        state.price = { map, name: file.name, count: map.size };
        document.getElementById('priceUploadArea').classList.add('loaded');
        setStatus('priceStatus',
            `✅ Загружено ${map.size} наименований` + (empty ? ` (пропущено повторов: ${empty})` : ''),
            'success');
    } catch (err) {
        state.price = null;
        setStatus('priceStatus', '❌ Ошибка: ' + err.message, 'error');
        console.error(err);
    }
    updateStartBtn();
}

/* ============================ Разбор PDF: общее ============================ */

// Элемент текста pdf.js может содержать сразу несколько слов — делим на слова,
// распределяя координату x пропорционально длине подстроки.
function itemsToWords(items) {
    const words = [];
    for (const it of items) {
        const str = it.str;
        if (!str || !str.trim()) continue;
        const x0 = it.transform[4], y = it.transform[5];
        const w = it.width || 0;
        const len = str.length;
        let i = 0;
        while (i < len) {
            while (i < len && /\s/.test(str[i])) i++;
            if (i >= len) break;
            let j = i;
            while (j < len && !/\s/.test(str[j])) j++;
            const piece = str.slice(i, j);
            words.push({
                text: piece,
                x: x0 + w * (i / len),
                w: w * ((j - i) / len),
                y: y,
                h: it.height || 0
            });
            i = j;
        }
    }
    return words;
}

// Идентификатор ii в PDF может быть разделён на фрагменты разными шрифтами.
// Объединяем ТОЛЬКО соседние фрагменты одной печатной строки, а не произвольные цифры страницы.
function normalizedLabelPart(value) {
    return String(value || '').replace(/[\u0456\u0406]/g, 'i') // кириллическая і
        .replace(/[\u0131]/g, 'i')
        .replace(/[\s\u00a0\u202f]/g, '').toLowerCase();
}

function labelIdsFromWords(words) {
    const result = new Set();
    const lines = [];
    for (const w of words) {
        const line = lines.find(l => Math.abs(l.y - w.y) <= Math.max(3, Math.min(7, Math.max(w.h || 0, l.h || 0) * 0.65)));
        if (line) { line.words.push(w); line.h = Math.max(line.h, w.h || 0); }
        else lines.push({ y: w.y, h: w.h || 0, words: [w] });
    }
    for (const line of lines) {
        const ordered = line.words.sort((a, b) => a.x - b.x);
        for (let i = 0; i < ordered.length; i++) {
            let combined = '';
            let right = null;
            for (let j = i; j < ordered.length && j < i + 20; j++) {
                const w = ordered[j];
                const part = normalizedLabelPart(w.text);
                if (right !== null && w.x - right > Math.max(24, (w.h || 0) * 2.5)) break;
                if (!/^[i0-9]+$/.test(part)) break;
                const next = combined + part;
                if (!/^i{0,2}\d{0,16}$/.test(next)) break;
                combined = next;
                right = Math.max(right === null ? -Infinity : right, w.x + w.w);
            }
            // Не считаем короткий префикс подходящим, если после него стоит продолжение.
            // Длина у ii-идентификаторов определяется конкретным PDF.
            if (LABEL_ID_RE.test(combined)) result.add(combined);
        }
    }
    return [...result];
}

/* ============================ Лист подбора (PDF) ============================ */

async function handleLp(file) {
    document.getElementById('lpFileName').textContent = file.name;
    document.getElementById('lpUploadArea').classList.remove('loaded');
    setStatus('lpStatus', '⏳ Чтение PDF...', '');
    try {
        const buf = await file.arrayBuffer();
        const doc = await pdfjsLib.getDocument({ data: buf, useSystemFonts: true }).promise;
        const rowsMap = new Map(); // номер -> {num, article, name}
        let dupCount = 0;

        for (let p = 1; p <= doc.numPages; p++) {
            const page = await doc.getPage(p);
            const tc = await page.getTextContent();
            const words = itemsToWords(tc.items).filter(w => w.text.trim());

            // Якоря — номера отправлений (колонка «Номер отправления»).
            // После 999-й строки порядковый номер может склеиться с номером
            // отправления в одно слово — учитываем это.
            const anchors = [];
            for (const wd of words) {
                if (wd.x < 25 || wd.x >= 190) continue;
                if (NUM_RE.test(wd.text)) {
                    anchors.push({ ...wd, num: wd.text });
                } else {
                    const m = LP_GLUE_RE.exec(wd.text);
                    if (m) {
                        const frac = m[1].length / wd.text.length;
                        anchors.push({ ...wd, x: wd.x + wd.w * frac, num: m[2] });
                    }
                }
            }
            if (anchors.length) {
                anchors.sort((a, b) => b.y - a.y); // сверху вниз

                // Границы строк — середины между соседними якорями.
                const bounds = [];
                for (let i = 0; i < anchors.length; i++) {
                    const top = i === 0
                        ? anchors[0].y + 30 // чтобы шапка страницы не попала в первую строку
                        : (anchors[i - 1].y + anchors[i].y) / 2;
                    const bottom = i === anchors.length - 1
                        ? -Infinity
                        : (anchors[i].y + anchors[i + 1].y) / 2;
                    bounds.push([top, bottom]);
                }

                // Строка шапки таблицы — исключаем её слова из названия.
                const isHeader = (w) => {
                    if (!/^(Товар|Фото|Артикул|Кол-во|Этикетка|№Номер|отправления|№)$/.test(w.text)) return false;
                    return words.some(v =>
                        Math.abs(v.y - w.y) < 3 &&
                        ((v.x >= 395 && v.x < 480) || v.x >= 525));
                };

                const nameCol = [], artCol = [];
                for (const wd of words) {
                    if (isHeader(wd)) continue;
                    if (wd.x >= 186 && wd.x < 395) nameCol.push(wd);
                    else if (wd.x >= 395 && wd.x < 480) artCol.push(wd);
                }

                for (let i = 0; i < anchors.length; i++) {
                    const a = anchors[i];
                    const [top, bottom] = bounds[i];
                    const inBand = (wd) => wd.y <= top && wd.y >= bottom;

                    // Название: слова колонки «Товар», построчно сверху вниз.
                    const lines = new Map();
                    for (const wd of nameCol) {
                        if (!inBand(wd)) continue;
                        const key = Math.round(wd.y * 2) / 2;
                        if (!lines.has(key)) lines.set(key, []);
                        lines.get(key).push(wd);
                    }
                    const name = [...lines.keys()].sort((a, b) => b - a)
                        .map(k => lines.get(k).sort((a, b) => a.x - b.x).map(w => w.text).join(' '))
                        .join(' ').replace(/\s+/g, ' ').trim();

                    // Артикул — ближайшее слово своей колонки к строке якоря.
                    let best = null, bd = Infinity;
                    for (const wd of artCol) {
                        if (!inBand(wd)) continue;
                        const d = Math.abs(wd.y - a.y);
                        if (d < bd) { bd = d; best = wd; }
                    }

                    // Серый ID расположен под номером отправления в той же
                    // колонке. Берём его только в пределах соответствующей строки.
                    // Серый ШК часто расположен правее номера отправления и может
                    // попадать ниже 24 pt. Вся полоса строки безопаснее узкой зоны.
                    const labelCandidates = labelIdsFromWords(words.filter(w =>
                        w.x >= Math.max(20, a.x - 12) && w.x < 270 &&
                        w.y <= a.y + 5 && w.y >= bottom));
                    const labelId = labelCandidates.length === 1 ? labelCandidates[0] : null;
                    const existing = rowsMap.get(a.num);
                    if (existing) {
                        dupCount++;
                        if (!existing.name && name) existing.name = name;
                        if (!existing.article && best) existing.article = best.text;
                        if (!existing.labelId && labelId) existing.labelId = labelId;
                    } else {
                        rowsMap.set(a.num, { num: a.num, article: best ? best.text : '', name, labelId });
                    }
                }
            }
            setStatus('lpStatus', `⏳ Страница ${p} из ${doc.numPages}`, '');
        }

        const rows = [...rowsMap.values()];
        if (rows.length === 0) throw new Error('не найдено ни одного номера отправления — возможно, это не Лист подбора');
        state.lp = { rows, name: file.name };
        document.getElementById('lpUploadArea').classList.add('loaded');
        setStatus('lpStatus',
            `✅ Найдено ${rows.length} отправлений на ${doc.numPages} стр.` +
            (dupCount ? ` (повторов: ${dupCount})` : ''),
            'success');
    } catch (err) {
        state.lp = null;
        setStatus('lpStatus', '❌ Ошибка: ' + err.message, 'error');
        console.error(err);
    }
    updateStartBtn();
}

/* ============================ Наклейки (PDF) ============================ */

function stickerNumberFromPage(tc) {
    // В новых этикетках ii-префикс и четыре последние цифры могут быть
    // разными текстовыми элементами/шрифтами одной строки.
    const ids = labelIdsFromWords(itemsToWords(tc.items));
    if (ids.length > 1) return null; // неоднозначный ШК
    if (ids.length === 1) return ids[0];

    // Дополнительное чтение ii по текстовым элементам PDF, когда координаты
    // шрифтов немного различаются. Принимаем только ПОЛНЫЙ идентификатор.
    const chunks = tc.items.map(it => String(it.str || '').trim()).filter(Boolean);
    const joined = chunks.join(' ');
    const newIds = [...new Set([...joined.matchAll(/(?:^|[^\w])(ii\s*(?:\d\s*){8,16})(?=$|[^\w])/gi)]
        .map(m => normalizedLabelPart(m[1])).filter(v => LABEL_ID_RE.test(v)))];
    if (newIds.length > 1) return null;
    if (newIds.length === 1) return newIds[0];

    // Старые Ozon: «4414 0656» и «-0784-1» могут находиться в нескольких
    // pdf.js items. Разрешаем пробелы ВНУТРИ цифровой части, но не меняем
    // ни одной цифры и не используем последние четыре цифры как ключ.
    const oldIds = new Set();
    for (const m of joined.matchAll(/(?:^|[^0-9])((?:\d{8,10}|\d{4,7}[\s\u00a0\u202f]+\d{1,6})[\s\u00a0\u202f]*-\s*\d{3,5}-\s*\d{1,3})(?!\d)/g)) {
        const candidate = m[1].replace(/[\s\u00a0\u202f]+/g, '');
        if (NUM_RE.test(candidate)) oldIds.add(candidate);
    }
    // В старом варианте номер иногда хранится слитно с четырёхзначным кодом.
    for (const value of chunks) {
        const m = STICKER_GLUE_RE.exec(value.replace(/\s+/g, ''));
        if (m) oldIds.add(m[1]);
    }
    return oldIds.size === 1 ? [...oldIds][0] : null;
}

// Соответствие строится строго в порядке УЖЕ отсортированного листа.
// Каждая страница наклейки может использоваться только один раз.
function matchStickerPages(rows, pageOf, pageCount) {
    const pageIndices = [], missingStickers = [], usedPages = new Set();
    const labelOwners = new Map();
    for (const r of rows) {
        if (!r.labelId) continue;
        if (!labelOwners.has(r.labelId)) labelOwners.set(r.labelId, new Set());
        labelOwners.get(r.labelId).add(r.num);
    }
    for (const r of rows) {
        const key = r.labelId || r.num;
        const unique = !r.labelId || labelOwners.get(r.labelId).size === 1;
        const available = unique ? (pageOf.get(key) || []) : [];
        const pg = available.find(p => Number.isInteger(p) && p >= 1 && p <= pageCount && !usedPages.has(p)) || null;
        pageIndices.push(pg);
        if (pg === null) missingStickers.push(r.num);
        else usedPages.add(pg);
    }
    const extraStickers = [];
    for (let p = 1; p <= pageCount; p++) if (!usedPages.has(p)) extraStickers.push(p);
    if (pageIndices.filter(p => p !== null).length !== usedPages.size) {
        throw new Error('Ошибка проверки: одна страница наклейки назначена повторно');
    }
    return { pageIndices, missingStickers, extraStickers };
}

async function handleTicket(file) {
    document.getElementById('ticketFileName').textContent = file.name;
    document.getElementById('ticketUploadArea').classList.remove('loaded');
    setStatus('ticketStatus', '⏳ Чтение PDF...', '');
    try {
        const buf = await file.arrayBuffer();
        const doc = await pdfjsLib.getDocument({ data: buf.slice(0) }).promise;
        const pageOf = new Map();
        const pageToNum = new Map(); // страница -> номер этикетки (для сообщений о несовпадении)
        let noNum = 0;
        for (let p = 1; p <= doc.numPages; p++) {
            const page = await doc.getPage(p);
            const tc = await page.getTextContent();
            const num = stickerNumberFromPage(tc);
            if (num) {
                if (!pageOf.has(num)) pageOf.set(num, []);
                pageOf.get(num).push(p);
                pageToNum.set(p, num);
            } else {
                noNum++;
            }
            setStatus('ticketStatus', `⏳ Страница ${p} из ${doc.numPages}`, '');
        }
        state.ticket = { buffer: buf, pageOf, pageToNum, name: file.name, pages: doc.numPages };
        document.getElementById('ticketUploadArea').classList.add('loaded');
        setStatus('ticketStatus',
            `✅ Найдено ${pageOf.size} этикеток на ${doc.numPages} стр.` +
            (noNum ? ` (без номера: ${noNum})` : ''),
            'success');
    } catch (err) {
        state.ticket = null;
        setStatus('ticketStatus', '❌ Ошибка: ' + err.message, 'error');
        console.error(err);
    }
    updateStartBtn();
}

/* ============================ Генерация Excel ============================ */

// Минимальный корректный xlsx с оформлением: жирная шапка с заливкой и рамками,
// закреплённая строка, автофильтр, ширины столбцов, подсветка строк.
function buildXlsx(rows) {
    const zip = new JSZip();

    zip.file('[Content_Types].xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        '</Types>');

    zip.file('_rels/.rels',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        '</Relationships>');

    zip.file('xl/workbook.xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="Лист подбора" sheetId="1" r:id="rId1"/></sheets>' +
        '</workbook>');

    zip.file('xl/_rels/workbook.xml.rels',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
        '</Relationships>');

    zip.file('xl/styles.xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        '<fonts count="4">' +
        '<font><sz val="11"/><name val="Calibri"/></font>' + // 0 — обычный
        '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>' + // 1 — шапка
        '<font><sz val="11"/><color rgb="FF7A5B00"/><name val="Calibri"/></font>' + // 2 — подсвеченная строка
        '<font><b/><sz val="14"/><color rgb="FF005BFF"/><name val="Calibri"/></font>' + // 3 — выделенные четыре цифры ШК
        '</fonts>' +
        '<fills count="5">' +
        '<fill><patternFill patternType="none"/></fill>' +
        '<fill><patternFill patternType="gray125"/></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FF005BFF"/><bgColor indexed="64"/></patternFill></fill>' + // 2 — синяя шапка
        '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF3CD"/><bgColor indexed="64"/></patternFill></fill>' + // 3 — жёлтая подсветка
        '<fill><patternFill patternType="solid"><fgColor rgb="FFD9D9D9"/><bgColor indexed="64"/></patternFill></fill>' + // 4 — серая полоса блока товара
        '</fills>' +
        '<borders count="2">' +
        '<border><left/><right/><top/><bottom/><diagonal/></border>' +
        '<border><left style="thin"><color rgb="FF9DC3F7"/></left><right style="thin"><color rgb="FF9DC3F7"/></right><top style="thin"><color rgb="FF9DC3F7"/></top><bottom style="thin"><color rgb="FF9DC3F7"/></bottom><diagonal/></border>' +
        '</borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="5">' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' + // 0 — обычная ячейка
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">' +
        '<alignment horizontal="center" vertical="center"/>' + // 1 — шапка
        '</xf>' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' + // 2 — данные
        '<xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/>' + // 3 — данные с серой полосой
        '<xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' + // 4 — название не из прайса
        '</cellXfs>' +
        '<cellStyles count="1"><cellStyle name="Обычный" xfId="0" builtinId="0"/></cellStyles>' +
        '</styleSheet>');

    const n = rows.length;
    let sheet =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        `<dimension ref="A1:D${n + 1}"/>` +
        '<sheetViews><sheetView workbookViewId="0">' +
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' +
        '</sheetView></sheetViews>' +
        '<sheetFormatPr defaultRowHeight="15"/>' +
        '<cols>' +
        '<col min="1" max="1" width="20" customWidth="1"/>' +
        '<col min="2" max="2" width="12" customWidth="1"/>' +
        '<col min="3" max="3" width="84" customWidth="1"/>' +
        '<col min="4" max="4" width="22" customWidth="1"/>' +
        '</cols>' +
        '<sheetData>';

    const cell = (ref, value, style) =>
        `<c r="${ref}"${style ? ` s="${style}"` : ''} t="inlineStr"><is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`;

    // Номер ШК всегда полный: новый ii... или обычный номер отправления.
    // В XLSX выделяем исключительно нужную 4-значную группу, не изменяя текст номера.
    const labelCell = (ref, r, style) => {
        const value = r.labelId && LABEL_ID_RE.test(r.labelId) ? r.labelId : r.num;
        const newId = /^(ii\d+)(\d{4})$/i.exec(value);
        const oldId = /^(\d{1,6})(\d{4})(-\d{3,5}-\d{1,3})$/.exec(value);
        const parts = newId ? [newId[1], newId[2], '']
                    : oldId ? [oldId[1], oldId[2], oldId[3]] : null;
        if (!parts) return cell(ref, value, style);
        const run = (text, fontId) => text ?
            `<r><rPr><rFont val="Calibri"/><sz val="${fontId === 3 ? 14 : 11}"/>` +
            (fontId === 3 ? '<b/><color rgb="FF005BFF"/>' : '') +
            `</rPr><t xml:space="preserve">${escapeXml(text)}</t></r>` : '';
        return `<c r="${ref}"${style ? ` s="${style}"` : ''} t="inlineStr"><is>` +
            run(parts[0], 0) + run(parts[1], 3) + run(parts[2], 0) + '</is></c>';
    };

    sheet += `<row r="1">${cell('A1', 'Номер отправления', 1)}${cell('B1', 'Артикул', 1)}${cell('C1', 'Товар', 1)}${cell('D1', 'Номер ШК', 1)}</row>`;

    // Чередование серых полос по блокам одинаковых товаров (видно при ч/б печати)
    let band = false;
    let prevName = null;
    rows.forEach((r, i) => {
        if (r.name !== prevName) { band = !band; prevName = r.name; }
        const rn = i + 2;
        const s = !r.fromPrice ? 4 : (band ? 3 : 2);
        sheet += `<row r="${rn}">${cell('A' + rn, r.num, s)}${cell('B' + rn, r.article, s)}${cell('C' + rn, r.name, s)}${labelCell('D' + rn, r, s)}</row>`;
    });

    sheet += `</sheetData><autoFilter ref="A1:D${n + 1}"/></worksheet>`;
    zip.file('xl/worksheets/sheet1.xml', sheet);

    return zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } });
}

/* ============================ Сборка ============================ */

async function startPack() {
    if (!(state.price && state.lp && state.ticket)) return;
    const btn = document.getElementById('startBtn');
    btn.disabled = true;
    document.getElementById('resultsSection').classList.add('hidden');
    document.getElementById('warningsSection').classList.add('hidden');
    document.getElementById('progressSection').classList.remove('hidden');
    setProgress(0, 'Обработка...', '');

    try {
        // 1. Обогащаем названия из прайса
        setProgress(5, 'Подстановка наименований из прайса...', '');
        await new Promise(r => setTimeout(r, 30));
        const rows = state.lp.rows.map((r, idx) => {
            const art = normArt(r.article);
            const priceName = art ? state.price.map.get(art) : null;
            return {
                num: r.num,
                labelId: r.labelId || null,
                article: r.article,
                name: priceName != null ? priceName : r.name,
                fromPrice: priceName != null,
                origIdx: idx
            };
        });
        const notInPrice = rows.filter(r => !r.fromPrice).length;

        // 2. Сортировка по наименованию (алфавит, русский);
        //    при одинаковых названиях сохраняется исходный порядок листа
        setProgress(12, 'Сортировка по алфавиту...', '');
        await new Promise(r => setTimeout(r, 30));
        rows.sort((a, b) =>
            a.name.localeCompare(b.name, 'ru') || a.origIdx - b.origIdx);

        // 3. Excel
        setProgress(20, 'Формирование Листа подбора (Excel)...', '');
        const xlsxBlob = await buildXlsx(rows);
        setProgress(55, 'Лист подбора готов', '');

        // 4. Порядок наклеек = порядок строк
        setProgress(60, 'Сопоставление наклеек с листом...', '');
        await new Promise(r => setTimeout(r, 30));
        const pageOf = state.ticket.pageOf;
        const { pageIndices, missingStickers, extraStickers } =
            matchStickerPages(rows, pageOf, state.ticket.pages);
        const extraNums = [];
        for (const p of extraStickers) {
            extraNums.push(state.ticket.pageToNum.get(p) || ('стр. ' + p));
        }

        // 5. Пересборка PDF (в pdf-lib индексы страниц с нуля)
        setProgress(65, 'Пересборка PDF с наклейками...', '');
        const src = await PDFLib.PDFDocument.load(state.ticket.buffer, { ignoreEncryption: true });
        const out = await PDFLib.PDFDocument.create();
        const ordered = pageIndices.filter(p => p != null).map(p => p - 1);
        const extraIdx = extraStickers.map(p => p - 1);
        const copied = await out.copyPages(src, [...ordered, ...extraIdx]);
        copied.forEach(pg => out.addPage(pg));
        const pdfBytes = await out.save();
        const pdfBlob = new Blob([pdfBytes], { type: 'application/pdf' });
        setProgress(100, 'Готово!', `Позиций: ${rows.length} | Наклеек: ${ordered.length}`);

        built = { rows, xlsxBlob, pdfBlob, missingStickers, extraStickers, extraNums };
        showResults();
    } catch (err) {
        console.error(err);
        setProgress(0, '❌ Ошибка: ' + err.message, '');
    }
    btn.disabled = false;
    updateStartBtn();
}

/* ============================ Результаты ============================ */

function showResults() {
    const { rows, missingStickers, extraStickers, extraNums } = built;
    const fromPrice = rows.filter(r => r.fromPrice).length;

    document.getElementById('summary').innerHTML = `
        <div class="summary-card ok">
            <span class="num">${rows.length}</span>
            <div class="lbl">Позиций в листе</div>
        </div>
        <div class="summary-card ok">
            <span class="num">${fromPrice}</span>
            <div class="lbl">Названий из прайса</div>
        </div>
        <div class="summary-card ${notInPriceCount() ? 'warn' : 'ok'}">
            <span class="num">${notInPriceCount()}</span>
            <div class="lbl">Нет в прайсе (взято из листа)</div>
        </div>
        <div class="summary-card ${(missingStickers.length || extraStickers.length) ? 'miss' : 'ok'}">
            <span class="num">${rows.length - missingStickers.length}</span>
            <div class="lbl">Наклеек в порядке листа</div>
        </div>
    `;

    let warnHtml = '';
    if (missingStickers.length) {
        warnHtml += `<p><strong>⚠️ Нет наклейки для ${missingStickers.length} отправл.:</strong> ${missingStickers.slice(0, 20).join(', ')}${missingStickers.length > 20 ? ' …' : ''}</p>`;
    }
    if (extraStickers.length) {
        warnHtml += `<p><strong>⚠️ ${extraNums.length} наклеек нет в Листе подбора:</strong> ${extraNums.slice(0, 20).join(', ')}${extraNums.length > 20 ? ' …' : ''} — они добавлены в конец PDF.</p>`;
    }
    if (notInPriceCount()) {
        warnHtml += `<p><strong>💡 ${notInPriceCount()} наимен. нет в прайсе</strong> — взяты из Листа подбора (жёлтая подсветка в Excel и таблице).</p>`;
    }
    if (warnHtml) {
        document.getElementById('warningsBox').innerHTML = warnHtml;
        document.getElementById('warningsSection').classList.remove('hidden');
    } else {
        document.getElementById('warningsSection').classList.add('hidden');
    }

    showAllMode = false;
    renderPreview();
    const showAllBtn = document.getElementById('showAllBtn');
    if (rows.length > 100) {
        showAllBtn.classList.remove('hidden');
        showAllBtn.textContent = `Показать все строки (${rows.length})`;
    } else {
        showAllBtn.classList.add('hidden');
    }

    // Нижний блок: соответствие этикеток Листу подбора с конкретными номерами
    const diffEl = document.getElementById('labelsDiff');
    let diffHtml = '<h4>🏷️ Проверка этикеток</h4>';
    if (missingStickers.length) {
        diffHtml += `<p><strong>Не хватает наклеек — отправления из Листа подбора, для которых нет этикетки (${missingStickers.length} шт.):</strong><br><span class="nums">${missingStickers.join(', ')}</span></p>`;
    }
    if (extraNums.length) {
        diffHtml += `<p><strong>Наклеек больше, чем позиций в Листе подбора — эти номера есть среди этикеток, но отсутствуют в листе (${extraNums.length} шт., добавлены в конец PDF):</strong><br><span class="nums">${extraNums.join(', ')}</span></p>`;
    }
    if (!missingStickers.length && !extraNums.length) {
        diffHtml += '<p><strong>✅ Этикетки совпадают с Листом подбора</strong> — каждой позиции соответствует своя наклейка.</p>';
    }
    diffEl.innerHTML = diffHtml;
    diffEl.classList.toggle('warn', !!(missingStickers.length || extraNums.length));
    diffEl.classList.remove('hidden');

    document.getElementById('resultsSection').classList.remove('hidden');
    document.getElementById('resultsSection').scrollIntoView({ behavior: 'smooth' });
}

function notInPriceCount() {
    return built ? built.rows.filter(r => !r.fromPrice).length : 0;
}

function renderPreview() {
    const { rows } = built;
    const limit = showAllMode ? rows.length : Math.min(rows.length, 100);
    let band = false;
    let prevName = null;
    document.getElementById('resultsBody').innerHTML = rows.slice(0, limit).map((r, i) => {
        if (r.name !== prevName) { band = !band; prevName = r.name; }
        const cls = !r.fromPrice ? 'from-lp' : (band ? 'band' : '');
        return `
        <tr class="${cls}">
            <td>${i + 1}</td>
            <td>${r.num}</td>
            <td>${r.article || '—'}</td>
            <td class="name-cell">${r.name || '—'}</td>
            <td><span class="badge ${r.fromPrice ? 'ok' : 'warn'}">${r.fromPrice ? 'Из прайса' : 'Из листа'}</span></td>
        </tr>`;
    }).join('') +
    (limit < rows.length
        ? `<tr><td colspan="5" style="text-align:center; color:#86868b;">… и ещё ${rows.length - limit} строк</td></tr>`
        : '');
}

function showAllRows() {
    showAllMode = true;
    renderPreview();
    document.getElementById('showAllBtn').classList.add('hidden');
}

function download(blob, filename) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 30000);
}

/* ============================ Инициализация ============================ */

document.getElementById('downloadXlsxBtn').addEventListener('click', () => {
    if (built) download(built.xlsxBlob, 'Лист_подбора.xlsx');
});
document.getElementById('downloadPdfBtn').addEventListener('click', () => {
    if (built) download(built.pdfBlob, 'Наклейки_по_порядку.pdf');
});

wireUpload('priceUploadArea', 'priceFileInput', 'priceStatus', 'priceFileName', handlePrice);
wireUpload('lpUploadArea', 'lpFileInput', 'lpStatus', 'lpFileName', handleLp);
wireUpload('ticketUploadArea', 'ticketFileInput', 'ticketStatus', 'ticketFileName', handleTicket);
