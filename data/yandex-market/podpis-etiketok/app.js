'use strict';

/* =========================================================
   СОСТОЯНИЕ
========================================================= */

const state = {
    oneC: [],
    lk: [],
    priceMap: new Map(),

    mapping: new Map(),

    pdfFile: null,

    lastOutputPdf: null,
    lastOrdersXlsx: null,

    processedPages: [],

    manualExcluded: new Set(),

    fontBytes: null,
    fontSource: null,
    fontkitSource: null,

    textConfig: {
        fontSize: 5,
        color: '#000000'
    }
};

let processingActive = false;


/* =========================================================
   OCR
========================================================= */

const OCR_SCALE = 1.0;
const OCR_LANG = 'eng';
const OCR_WORKERS = 3;


/* =========================================================
   ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
========================================================= */

function log(message) {
    const logEl =
        document.getElementById('log');

    if (!logEl) return;

    const time =
        new Date().toLocaleTimeString();

    logEl.textContent +=
        `[${time}] ${message}\n`;

    logEl.scrollTop =
        logEl.scrollHeight;
}


function setStatus(
    title,
    progress = null
) {
    const box =
        document.getElementById('statusBox');

    const titleEl =
        document.getElementById('statusTitle');

    const bar =
        document.getElementById('progressBar');

    if (box) {
        box.classList.add('show');
    }

    if (titleEl) {
        titleEl.textContent = title;
    }

    if (
        bar &&
        progress !== null
    ) {
        const value =
            Math.max(
                0,
                Math.min(
                    100,
                    Number(progress)
                )
            );

        bar.style.width =
            value + '%';
    }
}


function clearStatus() {
    const box =
        document.getElementById('statusBox');

    const logEl =
        document.getElementById('log');

    const bar =
        document.getElementById('progressBar');

    const unknownBox =
        document.getElementById('unknownOrdersBox');

    if (unknownBox) {
        unknownBox.remove();
    }

    if (box) {
        box.classList.remove('show');
    }

    if (logEl) {
        logEl.textContent = '';
    }

    if (bar) {
        bar.style.width = '0%';
    }
}


function normalizeHeader(value) {
    return String(value ?? '')
        .replace(/\u00A0/g, ' ')
        .trim()
        .replace(/\s+/g, ' ')
        .toLowerCase();
}


/* =========================================================
   НАУЧНАЯ ЗАПИСЬ
========================================================= */

function expandScientific(value) {
    let s =
        String(value)
            .trim()
            .replace(',', '.');

    const match =
        s.match(
            /^([+-]?)(\d+)(?:\.(\d+))?[eE]([+-]?\d+)$/
        );

    if (!match) {
        return s;
    }

    const sign =
        match[1] || '';

    const integerPart =
        match[2] || '';

    const fractionPart =
        match[3] || '';

    const exponent =
        parseInt(
            match[4],
            10
        );

    const digits =
        integerPart +
        fractionPart;

    const decimalPosition =
        integerPart.length +
        exponent;

    let result;

    if (
        decimalPosition <= 0
    ) {
        result =
            '0.' +
            '0'.repeat(
                -decimalPosition
            ) +
            digits;

    } else if (
        decimalPosition >=
        digits.length
    ) {
        result =
            digits +
            '0'.repeat(
                decimalPosition -
                digits.length
            );

    } else {
        result =
            digits.slice(
                0,
                decimalPosition
            ) +
            '.' +
            digits.slice(
                decimalPosition
            );
    }

    if (
        result.includes('.')
    ) {
        result =
            result
                .replace(
                    /0+$/,
                    ''
                )
                .replace(
                    /\.$/,
                    ''
                );
    }

    return sign + result;
}


/* =========================================================
   SKU
========================================================= */

function normalizeSku(value) {
    if (
        value === null ||
        value === undefined
    ) {
        return '';
    }

    let s =
        String(value)
            .replace(
                /\u00A0/g,
                ' '
            )
            .trim();

    if (!s) {
        return '';
    }

    s =
        s.replace(
            /\s+/g,
            ''
        );

    if (
        /^[+-]?\d+(?:[.,]\d+)?e[+-]?\d+$/i
            .test(s)
    ) {
        s =
            expandScientific(s);
    }

    s =
        s.replace(
            /\.0+$/,
            ''
        );

    return s
        .replace(
            /[^\dA-Za-zА-Яа-я_-]/g,
            ''
        )
        .toLowerCase();
}


/* =========================================================
   НОМЕР ЗАКАЗА
========================================================= */

function normalizeOrderId(value) {
    if (
        value === null ||
        value === undefined
    ) {
        return '';
    }

    let s =
        String(value)
            .replace(
                /\u00A0/g,
                ' '
            )
            .trim();

    if (!s) {
        return '';
    }

    s =
        s.replace(
            /^['"]|['"]$/g,
            ''
        );

    if (
        /^[+-]?\d+(?:[.,]\d+)?e[+-]?\d+$/i
            .test(s)
    ) {
        s =
            expandScientific(s);
    }

    if (
        /^\d+\.0+$/.test(s)
    ) {
        s =
            s.replace(
                /\.0+$/,
                ''
            );
    }

    s =
        s.replace(
            /[^\d]/g,
            ''
        );

    s =
        s.replace(
            /^0+/,
            ''
        );

    return s;
}


/* =========================================================
   EXCEL
========================================================= */

async function readExcelFile(file) {
    const buffer =
        await file.arrayBuffer();

    const workbook =
        XLSX.read(
            buffer,
            {
                type: 'array',
                cellDates: false,
                raw: false,
                cellNF: false,
                cellText: true
            }
        );

    return workbook;
}


function getAllSheetsRows(workbook) {
    const result = [];

    if (
        !workbook ||
        !workbook.SheetNames ||
        !workbook.SheetNames.length
    ) {
        throw new Error(
            'В Excel-файле нет листов.'
        );
    }

    for (
        const sheetName
        of workbook.SheetNames
    ) {
        const sheet =
            workbook.Sheets[
                sheetName
            ];

        if (!sheet) {
            continue;
        }

        const rows =
            XLSX.utils.sheet_to_json(
                sheet,
                {
                    header: 1,
                    raw: false,
                    defval: ''
                }
            );

        result.push({
            sheetName,
            rows
        });
    }

    return result;
}


function getFirstSheetRows(workbook) {
    const sheets =
        getAllSheetsRows(
            workbook
        );

    if (!sheets.length) {
        throw new Error(
            'Не удалось прочитать Excel-файл.'
        );
    }

    return sheets[0].rows;
}


/* =========================================================
   ПОИСК СТРОКИ ЗАГОЛОВКОВ
========================================================= */

function findHeaderRow(
    rows,
    requiredHeaders
) {
    const wanted =
        requiredHeaders.map(
            normalizeHeader
        );

    const limit =
        Math.min(
            rows.length,
            100
        );

    for (
        let r = 0;
        r < limit;
        r++
    ) {
        const row =
            rows[r] || [];

        const normalized =
            row.map(
                normalizeHeader
            );

        const found =
            wanted.every(
                header =>
                    normalized.includes(
                        header
                    )
            );

        if (found) {
            return {
                index: r,
                headers: normalized
            };
        }
    }

    return null;
}


/* =========================================================
   1С
========================================================= */

function parseOneC(rows) {
    const headerInfo =
        findHeaderRow(
            rows,
            [
                'Идентификатор МП',
                'Количество'
            ]
        );

    if (!headerInfo) {
        throw new Error(
            'В файле 1С не найдена строка с заголовками "Идентификатор МП" и "Количество".'
        );
    }

    const headers =
        headerInfo.headers;

    const mpIndex =
        headers.indexOf(
            normalizeHeader(
                'Идентификатор МП'
            )
        );

    const quantityIndex =
        headers.indexOf(
            normalizeHeader(
                'Количество'
            )
        );

    const result = [];

    for (
        let r =
            headerInfo.index + 1;
        r < rows.length;
        r++
    ) {
        const row =
            rows[r] || [];

        const orderId =
            normalizeOrderId(
                row[mpIndex]
            );

        if (!orderId) {
            continue;
        }

        const quantity =
            parseQuantity(
                row[quantityIndex]
            );

        result.push({
            orderId,
            quantity,
            sourceRow:
                r + 1
        });
    }

    return result;
}


function parseQuantity(value) {
    if (
        value === null ||
        value === undefined ||
        String(value).trim() === ''
    ) {
        return 1;
    }

    const s =
        String(value)
            .trim()
            .replace(',', '.');

    const number =
        Number(s);

    if (
        !Number.isFinite(number)
    ) {
        return 1;
    }

    return number;
}


/* =========================================================
   ЛК
========================================================= */

function extractOrderIdFromCell(value) {
    if (
        value === null ||
        value === undefined
    ) {
        return '';
    }

    const raw =
        String(value)
            .replace(
                /\u00A0/g,
                ' '
            )
            .trim();

    if (!raw) {
        return '';
    }

    const direct =
        normalizeOrderId(raw);

    if (
        direct.length >= 6
    ) {
        return direct;
    }

    const scientificMatches =
        raw.match(
            /[+-]?\d+(?:[.,]\d+)?e[+-]?\d+/gi
        );

    if (scientificMatches) {
        for (
            const candidate
            of scientificMatches
        ) {
            const id =
                normalizeOrderId(
                    candidate
                );

            if (
                id.length >= 6
            ) {
                return id;
            }
        }
    }

    const digitGroups =
        raw.match(
            /\d{6,}/g
        );

    if (
        digitGroups?.length
    ) {
        for (
            const group
            of digitGroups
        ) {
            const id =
                normalizeOrderId(
                    group
                );

            if (
                id.length >= 6
            ) {
                return id;
            }
        }
    }

    return '';
}


function findMatchingOneCOrder(
    cellValue,
    oneCIds
) {
    const raw =
        String(
            cellValue ?? ''
        )
            .replace(
                /\u00A0/g,
                ' '
            )
            .trim();

    if (!raw) {
        return null;
    }

    const normalized =
        normalizeOrderId(raw);

    if (
        normalized &&
        oneCIds.has(normalized)
    ) {
        return normalized;
    }

    const scientific =
        raw.match(
            /[+-]?\d+(?:[.,]\d+)?e[+-]?\d+/gi
        );

    if (scientific) {
        for (
            const value
            of scientific
        ) {
            const id =
                normalizeOrderId(
                    value
                );

            if (
                id &&
                oneCIds.has(id)
            ) {
                return id;
            }
        }
    }

    const groups =
        raw.match(
            /\d{6,}/g
        );

    if (groups) {
        for (
            const group
            of groups
        ) {
            const id =
                normalizeOrderId(
                    group
                );

            if (
                id &&
                oneCIds.has(id)
            ) {
                return id;
            }
        }
    }

    return null;
}


/*
 * Читаем реальные ячейки Excel.
 */
function scanLKForOrders(
    workbook,
    oneC
) {
    const oneCIds =
        new Set(
            oneC.map(
                item =>
                    normalizeOrderId(
                        item.orderId
                    )
            )
        );

    const found =
        new Map();

    for (
        const sheetName
        of workbook.SheetNames
    ) {
        const sheet =
            workbook.Sheets[
                sheetName
            ];

        if (!sheet) {
            continue;
        }

        const cells =
            Object.entries(sheet)
                .filter(
                    ([key, cell]) =>
                        /^[A-Z]+[0-9]+$/.test(key) &&
                        cell
                );

        const rows =
            new Map();

        for (
            const [address, cell]
            of cells
        ) {
            const pos =
                XLSX.utils.decode_cell(
                    address
                );

            if (
                !rows.has(pos.r)
            ) {
                rows.set(
                    pos.r,
                    []
                );
            }

            rows.get(pos.r)[pos.c] =
                cell.v ?? '';
        }

        for (
            const [address, cell]
            of cells
        ) {
            const candidates = [
                cell.v,
                cell.w
            ];

            let orderId = null;

            for (
                const candidate
                of candidates
            ) {
                orderId =
                    findMatchingOneCOrder(
                        candidate,
                        oneCIds
                    );

                if (orderId) {
                    break;
                }
            }

            if (!orderId) {
                continue;
            }

            const pos =
                XLSX.utils.decode_cell(
                    address
                );

            if (
                !found.has(orderId)
            ) {
                found.set(
                    orderId,
                    []
                );
            }

            found.get(orderId).push({
                sheetName,
                rowIndex: pos.r,
                columnIndex: pos.c,
                row:
                    rows.get(pos.r) ||
                    []
            });
        }
    }

    return found;
}


/* =========================================================
   ПОИСК SKU В СТРОКЕ ЛК
========================================================= */

function findSkusInRow(
    row,
    priceMap
) {
    const result = [];

    for (
        let c = 0;
        c < row.length;
        c++
    ) {
        const value =
            row[c];

        if (
            value === null ||
            value === undefined ||
            String(value).trim() === ''
        ) {
            continue;
        }

        const sku =
            normalizeSku(value);

        if (!sku) {
            continue;
        }

        if (
            priceMap.has(sku) &&
            !result.includes(sku)
        ) {
            result.push(sku);
        }
    }

    return result;
}


/* =========================================================
   ПРАЙС
========================================================= */

function parsePrice(rows) {
    const map =
        new Map();

    for (
        let r = 0;
        r < rows.length;
        r++
    ) {
        const row =
            rows[r] || [];

        if (
            row.length < 2
        ) {
            continue;
        }

        const sku =
            normalizeSku(
                row[0]
            );

        const name =
            String(
                row[1] ?? ''
            )
                .replace(
                    /\u00A0/g,
                    ' '
                )
                .trim();

        if (
            !sku ||
            !name
        ) {
            continue;
        }

        map.set(
            sku,
            name
        );
    }

    return map;
}


/* =========================================================
   ПОСТРОЕНИЕ СВЯЗКИ
========================================================= */

function buildMapping(
    oneC,
    lkWorkbook,
    priceMap
) {
    const mapping =
        new Map();

    const foundOrders =
        scanLKForOrders(
            lkWorkbook,
            oneC
        );

    log(
        `В ЛК просканирован весь файл: найдено совпадений по заказам — ${foundOrders.size}.`
    );

    for (
        const oneCRow
        of oneC
    ) {
        const orderId =
            oneCRow.orderId;

        const matches =
            foundOrders.get(
                orderId
            );

        if (
            !matches ||
            !matches.length
        ) {
            mapping.set(
                orderId,
                {
                    orderId,
                    quantity:
                        oneCRow.quantity,
                    status:
                        'LK_NOT_FOUND',
                    skus: [],
                    productNames: [],
                    sku: '',
                    productName: '',
                    reason:
                        'Номер заказа не найден во всём файле ЛК'
                }
            );

            continue;
        }

        const skuSet =
            new Set();

        for (
            const match
            of matches
        ) {
            const skus =
                findSkusInRow(
                    match.row,
                    priceMap
                );

            for (
                const sku
                of skus
            ) {
                skuSet.add(sku);
            }
        }

        const skus =
            [...skuSet];

        if (!skus.length) {
            mapping.set(
                orderId,
                {
                    orderId,
                    quantity:
                        oneCRow.quantity,
                    status:
                        'SKU_NOT_FOUND',
                    skus: [],
                    productNames: [],
                    sku: '',
                    productName: '',
                    reason:
                        'Для найденного заказа не найден SKU, присутствующий в прайсе'
                }
            );

            continue;
        }

        const productNames =
            skus.map(
                sku =>
                    priceMap.get(
                        sku
                    ) || ''
            );

        const combinedProductName =
            productNames
                .filter(Boolean)
                .join(' / ');

        mapping.set(
            orderId,
            {
                orderId,
                quantity:
                    oneCRow.quantity,

                status:
                    'FOUND',

                skus,

                productNames,

                sku:
                    skus.join(', '),

                productName:
                    combinedProductName,

                reason:
                    Number(
                        oneCRow.quantity
                    ) > 1
                        ? 'Количество больше 1 — в PDF не включается'
                        : ''
            }
        );
    }

    return mapping;
}


/* =========================================================
   СТАТИСТИКА
========================================================= */

function getStats(mapping) {
    const stats = {
        total: 0,
        found: 0,
        excluded: 0,
        lkNotFound: 0,
        skuNotFound: 0,
        priceNotFound: 0,
        conflicts: 0
    };

    for (
        const item
        of mapping.values()
    ) {
        stats.total++;

        if (
            Number(item.quantity) > 1
        ) {
            stats.excluded++;
        }

        if (
            item.status === 'FOUND' &&
            Number(item.quantity) === 1
        ) {
            stats.found++;
        }

        switch (
            item.status
        ) {
            case 'LK_NOT_FOUND':
                stats.lkNotFound++;
                break;

            case 'SKU_NOT_FOUND':
                stats.skuNotFound++;
                break;

            case 'PRICE_NOT_FOUND':
                stats.priceNotFound++;
                break;

            case 'DUPLICATE_CONFLICT':
                stats.conflicts++;
                break;
        }
    }

    return stats;
}


/* =========================================================
   ЛОГ НОМЕРОВ
========================================================= */

function logOrderNumbers(
    title,
    items
) {
    if (
        !items ||
        !items.length
    ) {
        log(
            `${title}: нет`
        );
        return;
    }

    const ids =
        items.map(
            item =>
                typeof item === 'string'
                    ? item
                    : item.orderId
        );

    log(
        `${title} (${ids.length}):`
    );

    const chunkSize = 15;

    for (
        let i = 0;
        i < ids.length;
        i += chunkSize
    ) {
        log(
            '  ' +
            ids
                .slice(
                    i,
                    i + chunkSize
                )
                .join(', ')
        );
    }
}


/* =========================================================
   КОПИРОВАНИЕ
========================================================= */

async function copyTextToClipboard(text) {
    const value =
        String(text ?? '');

    if (!value) {
        return false;
    }

    try {
        if (
            navigator.clipboard &&
            typeof navigator.clipboard.writeText === 'function'
        ) {
            await navigator.clipboard.writeText(
                value
            );

            return true;
        }
    } catch (error) {
        console.warn(
            'Clipboard API недоступен, использую запасной способ.',
            error
        );
    }

    try {
        const textarea =
            document.createElement(
                'textarea'
            );

        textarea.value =
            value;

        textarea.setAttribute(
            'readonly',
            ''
        );

        textarea.style.position =
            'fixed';

        textarea.style.left =
            '-9999px';

        textarea.style.top =
            '0';

        textarea.style.opacity =
            '0';

        document.body.appendChild(
            textarea
        );

        textarea.focus();
        textarea.select();

        const successful =
            document.execCommand(
                'copy'
            );

        textarea.remove();

        return successful;

    } catch (error) {
        console.error(
            'Не удалось скопировать номер заказа.',
            error
        );

        return false;
    }
}


/* =========================================================
   БЛОК НЕ ОБРАБОТАННЫХ ЗАКАЗОВ
========================================================= */

function showUnknownOrdersBox(
    unknownActions
) {
    const oldBox =
        document.getElementById(
            'unknownOrdersBox'
        );

    if (oldBox) {
        oldBox.remove();
    }

    if (
        !unknownActions ||
        !unknownActions.length
    ) {
        return;
    }

    /*
     * ВАЖНО:
     *
     * Здесь дополнительно пропускаем
     * только номера, которые реально
     * существуют в реестре 1С.
     */
    const oneCIds =
        new Set(
            getAllOneCOrderIds()
        );

    const orders = [];

    for (
        const action
        of unknownActions
    ) {
        const orderId =
            normalizeOrderId(
                action.orderId
            );

        if (
            orderId &&
            oneCIds.has(orderId) &&
            !orders.includes(orderId)
        ) {
            orders.push(orderId);
        }
    }

    const noOrderIdCount =
        unknownActions.filter(
            action =>
                !action.orderId
        ).length;

    const box =
        document.createElement(
            'div'
        );

    box.id =
        'unknownOrdersBox';

    box.style.cssText = `
        margin-top: 18px;
        padding: 16px 18px;
        border: 2px solid #f0b429;
        border-radius: 14px;
        background: #fffaf0;
        box-shadow: 0 4px 14px rgba(180, 120, 20, 0.08);
        font-family: inherit;
        color: #4a3a12;
    `;

    const title =
        document.createElement(
            'div'
        );

    title.textContent =
        '⚠️ Требуют проверки';

    title.style.cssText = `
        font-size: 16px;
        font-weight: 800;
        margin-bottom: 8px;
    `;

    box.appendChild(
        title
    );

    const description =
        document.createElement(
            'div'
        );

    description.textContent =
        'Эти страницы не попали в готовый PDF. Нажмите на номер заказа, чтобы скопировать его. В списке отображаются только номера из реестра 1С.';

    description.style.cssText = `
        font-size: 13px;
        margin-bottom: 12px;
        line-height: 1.45;
    `;

    box.appendChild(
        description
    );

    const list =
        document.createElement(
            'div'
        );

    list.style.cssText = `
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
    `;

    if (orders.length) {
        for (
            const orderId
            of orders
        ) {
            const item =
                document.createElement(
                    'button'
                );

            item.type =
                'button';

            item.textContent =
                orderId;

            item.title =
                'Нажмите, чтобы скопировать номер заказа';

            item.style.cssText = `
                display: inline-flex;
                align-items: center;
                justify-content: center;
                padding: 7px 11px;
                border-radius: 8px;
                background: #ffffff;
                border: 1px solid #e6c66a;
                color: #4a3a12;
                font-size: 13px;
                font-weight: 700;
                font-family: monospace;
                cursor: pointer;
                transition:
                    background .18s ease,
                    border-color .18s ease,
                    color .18s ease,
                    transform .08s ease;
            `;

            item.addEventListener(
                'mouseenter',
                function () {
                    if (
                        item.dataset.copied !== '1'
                    ) {
                        item.style.background =
                            '#fff4c7';

                        item.style.borderColor =
                            '#d7ad32';
                    }
                }
            );

            item.addEventListener(
                'mouseleave',
                function () {
                    if (
                        item.dataset.copied !== '1'
                    ) {
                        item.style.background =
                            '#ffffff';

                        item.style.borderColor =
                            '#e6c66a';
                    }
                }
            );

            item.addEventListener(
                'mousedown',
                function () {
                    item.style.transform =
                        'scale(0.97)';
                }
            );

            item.addEventListener(
                'mouseup',
                function () {
                    item.style.transform =
                        'scale(1)';
                }
            );

            item.addEventListener(
                'click',
                async function () {
                    const copied =
                        await copyTextToClipboard(
                            orderId
                        );

                    if (!copied) {
                        return;
                    }

                    item.dataset.copied =
                        '1';

                    item.textContent =
                        '✓ ' + orderId;

                    item.style.background =
                        '#dcfce7';

                    item.style.borderColor =
                        '#86efac';

                    item.style.color =
                        '#166534';

                    item.style.cursor =
                        'default';

                    item.title =
                        'Номер уже скопирован';

                    item.style.transform =
                        'scale(1)';
                }
            );

            list.appendChild(
                item
            );
        }
    }

    if (
        noOrderIdCount > 0
    ) {
        const item =
            document.createElement(
                'span'
            );

        item.textContent =
            `Номер не распознан: ${noOrderIdCount} стр.`;

        item.style.cssText = `
            display: inline-flex;
            align-items: center;
            padding: 7px 11px;
            border-radius: 8px;
            background: #ffffff;
            border: 1px solid #e6c66a;
            color: #7a5a00;
            font-size: 13px;
            font-weight: 700;
        `;

        list.appendChild(
            item
        );
    }

    if (
        !orders.length &&
        !noOrderIdCount
    ) {
        const item =
            document.createElement(
                'span'
            );

        item.textContent =
            'Нет распознанных номеров из реестра 1С для проверки.';

        item.style.cssText = `
            font-size: 13px;
            color: #7a5a00;
        `;

        list.appendChild(
            item
        );
    }

    box.appendChild(
        list
    );

    const statusBox =
        document.getElementById(
            'statusBox'
        );

    if (
        statusBox &&
        statusBox.parentNode
    ) {
        statusBox.parentNode.insertBefore(
            box,
            statusBox.nextSibling
        );
    } else {
        document.body.appendChild(
            box
        );
    }
}


/* =========================================================
   ЗАГРУЗКА 1С
========================================================= */

async function handleOneCFile(file) {
    setStatus(
        'Читаю реестр 1С...',
        5
    );

    const workbook =
        await readExcelFile(
            file
        );

    const rows =
        getFirstSheetRows(
            workbook
        );

    state.oneC =
        parseOneC(rows);

    state.mapping =
        new Map();

    state.lastOutputPdf =
        null;

    log(
        `Реестр 1С: найдено ${state.oneC.length} заказов.`
    );

    const excluded =
        state.oneC.filter(
            x =>
                Number(
                    x.quantity
                ) > 1
        );

    log(
        `Количество > 1: ${excluded.length}.`
    );

    logOrderNumbers(
        'Заказы с количеством > 1',
        excluded
    );

    rebuildMappingIfPossible();
}


/* =========================================================
   ЗАГРУЗКА ЛК
========================================================= */

async function handleLKFile(file) {
    setStatus(
        'Читаю весь файл ЛК...',
        15
    );

    const workbook =
        await readExcelFile(
            file
        );

    state.lk =
        workbook;

    state.mapping =
        new Map();

    state.lastOutputPdf =
        null;

    let totalRows = 0;

    for (
        const sheetName
        of workbook.SheetNames
    ) {
        const sheet =
            workbook.Sheets[
                sheetName
            ];

        if (!sheet) {
            continue;
        }

        const rows =
            XLSX.utils.sheet_to_json(
                sheet,
                {
                    header: 1,
                    raw: false,
                    defval: ''
                }
            );

        totalRows +=
            rows.length;
    }

    log(
        `ЛК: прочитан весь файл. Листов: ${workbook.SheetNames.length}, строк: ${totalRows}.`
    );

    log(
        'Номера заказов будут искаться во всех ячейках файла, но только среди номеров из реестра 1С.'
    );

    rebuildMappingIfPossible();
}


/* =========================================================
   ЗАГРУЗКА ПРАЙСА
========================================================= */

async function handlePriceFile(file) {
    setStatus(
        'Читаю прайс...',
        25
    );

    const workbook =
        await readExcelFile(
            file
        );

    const rows =
        getFirstSheetRows(
            workbook
        );

    state.priceMap =
        parsePrice(rows);

    state.mapping =
        new Map();

    state.lastOutputPdf =
        null;

    log(
        `Прайс: найдено ${state.priceMap.size} SKU.`
    );

    rebuildMappingIfPossible();
}


/* =========================================================
   ЗАГРУЗКА PDF
========================================================= */

async function handlePdfFile(file) {
    state.pdfFile =
        file;

    state.lastOutputPdf =
        null;

    log(
        `Этикетки: ${file.name}`
    );

    log(
        `Размер PDF: ${(file.size / 1024 / 1024).toFixed(2)} МБ`
    );

    if (
        state.mapping.size
    ) {
        log(
            'Связка построена. Запускаю обработку этикеток...'
        );

        await processPdf();
    }
}


/* =========================================================
   ПЕРЕСТРОЕНИЕ СВЯЗКИ
========================================================= */

function rebuildMappingIfPossible() {
    if (
        !state.oneC.length ||
        !state.lk ||
        !state.lk.SheetNames ||
        !state.lk.SheetNames.length ||
        !state.priceMap.size
    ) {
        updateDownloadButtons();
        return;
    }

    state.mapping =
        buildMapping(
            state.oneC,
            state.lk,
            state.priceMap
        );

    const stats =
        getStats(
            state.mapping
        );

    const foundOrders = [];
    const quantityOrders = [];
    const lkNotFoundOrders = [];
    const skuNotFoundOrders = [];
    const conflictOrders = [];

    for (
        const item
        of state.mapping.values()
    ) {
        if (
            item.status === 'FOUND' &&
            Number(item.quantity) === 1
        ) {
            foundOrders.push(item);
        }

        if (
            Number(item.quantity) > 1
        ) {
            quantityOrders.push(item);
        }

        if (
            item.status === 'LK_NOT_FOUND'
        ) {
            lkNotFoundOrders.push(item);
        }

        if (
            item.status === 'SKU_NOT_FOUND'
        ) {
            skuNotFoundOrders.push(item);
        }

        if (
            item.status === 'DUPLICATE_CONFLICT'
        ) {
            conflictOrders.push(item);
        }
    }

    log('');

    log(
        '========================================'
    );

    log(
        'СОПОСТАВЛЕНИЕ'
    );

    log(
        '========================================'
    );

    log(
        `Всего заказов из 1С: ${stats.total}`
    );

    log(
        `Найдено и готово: ${stats.found}`
    );

    logOrderNumbers(
        'Заказы quantity = 1',
        foundOrders
    );

    log('');

    log(
        `Количество > 1: ${stats.excluded}`
    );

    logOrderNumbers(
        'Заказы quantity > 1',
        quantityOrders
    );

    log('');

    log(
        `Не найдено в ЛК: ${stats.lkNotFound}`
    );

    logOrderNumbers(
        'Заказы не найдены в ЛК',
        lkNotFoundOrders
    );

    log('');

    log(
        `SKU не найден: ${stats.skuNotFound}`
    );

    logOrderNumbers(
        'Заказы без SKU',
        skuNotFoundOrders
    );

    log('');

    log(
        `SKU не найден в прайсе: ${stats.priceNotFound}`
    );

    log(
        `Конфликты SKU: ${stats.conflicts}`
    );

    if (
        conflictOrders.length
    ) {
        logOrderNumbers(
            'Заказы с конфликтом SKU',
            conflictOrders
        );
    }

    log(
        '========================================'
    );

    log('');

    log(
        'НАЙДЕННЫЕ SKU ПО ЗАКАЗАМ:'
    );

    const foundWithSku =
        [...state.mapping.values()]
            .filter(
                item =>
                    item.status ===
                    'FOUND'
            );

    for (
        const item
        of foundWithSku
    ) {
        log(
            `${item.orderId} → ${item.skus.join(', ')}`
        );
    }

    log('');

    setStatus(
        'Данные сопоставлены',
        30
    );

    updateDownloadButtons();
}


/* =========================================================
   OCR — ТОЛЬКО НОМЕРА ИЗ 1С
========================================================= */

function getEligibleOrderIds() {
    const result = [];

    for (
        const item
        of state.mapping.values()
    ) {
        if (
            item.status === 'FOUND' &&
            Number(item.quantity) === 1 &&
            !state.manualExcluded.has(
                item.orderId
            )
        ) {
            result.push(
                item.orderId
            );
        }
    }

    return result;
}


function getExcludedQuantityOrderIds() {
    const result = [];

    for (
        const item
        of state.mapping.values()
    ) {
        if (
            item.status === 'FOUND' &&
            Number(item.quantity) > 1
        ) {
            result.push(
                item.orderId
            );
        }
    }

    return result;
}


/*
 * Все номера ИСКЛЮЧИТЕЛЬНО из реестра 1С.
 */
function getAllOneCOrderIds() {
    return [
        ...new Set(
            state.oneC
                .map(
                    item =>
                        normalizeOrderId(
                            item.orderId
                        )
                )
                .filter(Boolean)
        )
    ];
}


/*
 * =========================================================
 * ПОИСК НОМЕРА В OCR
 *
 * КРИТИЧЕСКОЕ ПРАВИЛО:
 *
 * Функция НИКОГДА не возвращает номер,
 * которого нет в orderIds.
 *
 * orderIds = номера только из реестра 1С.
 *
 * Также запрещено простое:
 *
 *   digitsOnly.includes(orderId)
 *
 * потому что оно могло найти, например,
 * номер 123456 внутри числа 991234567890.
 *
 * Теперь номер должен совпасть целиком.
 * Между цифрами допускаются пробелы/переносы.
 * =========================================================
 */
function findOrderIdInOcrText(
    text,
    orderIds
) {
    if (
        !text ||
        !orderIds ||
        !orderIds.length
    ) {
        return null;
    }

    const sourceText =
        String(text);

    const sortedIds =
        [
            ...new Set(
                orderIds
                    .map(
                        normalizeOrderId
                    )
                    .filter(Boolean)
            )
        ].sort(
            (a, b) =>
                String(b).length -
                String(a).length
        );

    /*
     * -----------------------------------------------------
     * Вариант 1.
     *
     * Ищем последовательность цифр,
     * разрешая пробелы и переносы между цифрами.
     *
     * Границы:
     * слева и справа НЕ должно быть цифры.
     *
     * Поэтому:
     *
     * 123456
     *
     * найдётся,
     *
     * но:
     *
     * 991234567890
     *
     * не даст ложного совпадения.
     * -----------------------------------------------------
     */

    for (
        const orderId
        of sortedIds
    ) {
        const pattern =
            orderId
                .split('')
                .join(
                    '[\\s\\r\\n]*'
                );

        const regex =
            new RegExp(
                `(?<!\\d)${pattern}(?!\\d)`
            );

        if (
            regex.test(
                sourceText
            )
        ) {
            return orderId;
        }
    }

    /*
     * -----------------------------------------------------
     * Вариант 2.
     *
     * OCR иногда вставляет символы вроде:
     *
     * 123 456 789
     *
     * или:
     *
     * 123-456-789
     *
     * Но здесь всё равно проверяем только
     * номера из реестра 1С.
     *
     * Разрешаем между цифрами пробелы,
     * дефисы и точки.
     * -----------------------------------------------------
     */

    for (
        const orderId
        of sortedIds
    ) {
        const pattern =
            orderId
                .split('')
                .join(
                    '[\\s\\-._]*'
                );

        const regex =
            new RegExp(
                `(?<!\\d)${pattern}(?!\\d)`
            );

        if (
            regex.test(
                sourceText
            )
        ) {
            return orderId;
        }
    }

    return null;
}


/* =========================================================
   PDF.JS
========================================================= */

async function renderPdfPage(
    pdf,
    pageNumber
) {
    const page =
        await pdf.getPage(
            pageNumber
        );

    const viewport =
        page.getViewport({
            scale: OCR_SCALE
        });

    const canvas =
        document.createElement(
            'canvas'
        );

    const context =
        canvas.getContext(
            '2d',
            {
                willReadFrequently:
                    true
            }
        );

    canvas.width =
        Math.ceil(
            viewport.width
        );

    canvas.height =
        Math.ceil(
            viewport.height
        );

    await page.render({
        canvasContext:
            context,
        viewport
    }).promise;

    /*
     * Основной быстрый OCR:
     * верхние 35%.
     */
    const ocrCanvas =
        document.createElement(
            'canvas'
        );

    const ocrContext =
        ocrCanvas.getContext(
            '2d',
            {
                willReadFrequently:
                    true
            }
        );

    ocrCanvas.width =
        canvas.width;

    ocrCanvas.height =
        Math.max(
            1,
            Math.floor(
                canvas.height * 0.35
            )
        );

    ocrContext.drawImage(
        canvas,
        0,
        0,
        canvas.width,
        ocrCanvas.height,
        0,
        0,
        ocrCanvas.width,
        ocrCanvas.height
    );

    return {
        canvas,
        ocrCanvas
    };
}


/* =========================================================
   ШРИФТ
========================================================= */

async function getFontBytes() {
    if (state.fontBytes) {
        return state.fontBytes;
    }

    if (
        window.DEJAVU_FONT_B64
    ) {
        try {
            const base64 =
                window.DEJAVU_FONT_B64
                    .replace(
                        /^data:.*?base64,/,
                        ''
                    )
                    .replace(
                        /\s/g,
                        ''
                    );

            const binary =
                atob(base64);

            const bytes =
                new Uint8Array(
                    binary.length
                );

            for (
                let i = 0;
                i < binary.length;
                i++
            ) {
                bytes[i] =
                    binary.charCodeAt(i);
            }

            state.fontBytes =
                bytes;

            state.fontSource =
                'font-data.js';

            return bytes;

        } catch (error) {
            console.warn(
                'Не удалось загрузить шрифт из font-data.js',
                error
            );
        }
    }

    try {
        const response =
            await fetch(
                'fonts/DejaVuSans.ttf'
            );

        if (response.ok) {
            state.fontBytes =
                new Uint8Array(
                    await response.arrayBuffer()
                );

            state.fontSource =
                'fonts/DejaVuSans.ttf';

            return state.fontBytes;
        }

    } catch (error) {
        console.warn(
            'Локальный шрифт недоступен',
            error
        );
    }

    throw new Error(
        'Не найден DejaVuSans.ttf. Проверьте font-data.js или папку fonts.'
    );
}


async function getFontkit() {
    if (window.fontkit) {
        return window.fontkit;
    }

    await loadScript(
        'https://cdn.jsdelivr.net/npm/@pdf-lib/fontkit@1.1.1/dist/fontkit.umd.min.js'
    );

    if (window.fontkit) {
        return window.fontkit;
    }

    throw new Error(
        'Не удалось загрузить fontkit.'
    );
}


function loadScript(src) {
    return new Promise(
        (
            resolve,
            reject
        ) => {
            const script =
                document.createElement(
                    'script'
                );

            script.src =
                src;

            script.onload =
                () => resolve();

            script.onerror =
                () =>
                    reject(
                        new Error(
                            `Не удалось загрузить ${src}`
                        )
                    );

            document.head.appendChild(
                script
            );
        }
    );
}


/* =========================================================
   ЦВЕТ
========================================================= */

function parseColorLib(value) {
    const color =
        String(
            value ||
            '#000000'
        )
            .trim()
            .replace(
                '#',
                ''
            );

    if (
        /^[0-9a-fA-F]{6}$/.test(
            color
        )
    ) {
        const r =
            parseInt(
                color.slice(0, 2),
                16
            ) / 255;

        const g =
            parseInt(
                color.slice(2, 4),
                16
            ) / 255;

        const b =
            parseInt(
                color.slice(4, 6),
                16
            ) / 255;

        return PDFLib.rgb(
            r,
            g,
            b
        );
    }

    return PDFLib.rgb(
        0,
        0,
        0
    );
}


/* =========================================================
   РАЗМЕЩЕНИЕ
   НЕ МЕНЯЕМ
========================================================= */

function normalizeAngle(angle) {
    let a =
        Number(angle) || 0;

    a =
        (
            (a % 360) +
            360
        ) % 360;

    if (a === 360) {
        a = 0;
    }

    return a;
}


function pythonPlacement(
    rot,
    mw,
    mh
) {
    switch (
        normalizeAngle(rot)
    ) {
        case 90:
            return {
                x: mw - 2,
                y: 0
            };

        case 270:
            return {
                x: 5,
                y: mh - 1
            };

        case 180:
            return {
                x: mw - 10,
                y: mh - 1
            };

        case 0:
        default:
            return {
                x: 5,
                y: 1
            };
    }
}


function drawLabel(
    page,
    font,
    productName,
    cfg
) {
    const rot =
        normalizeAngle(
            page.getRotation().angle
        );

    const mb =
        page.getMediaBox();

    const mw =
        mb.width;

    const mh =
        mb.height;

    const {
        x,
        y
    } =
        pythonPlacement(
            rot,
            mw,
            mh
        );

    page.drawText(
        String(
            productName || ''
        ),
        {
            x,
            y,
            size:
                Number(
                    cfg.fontSize
                ) || 5,
            font,
            color:
                parseColorLib(
                    cfg.color
                ),
            rotate:
                PDFLib.degrees(90)
        }
    );
}


/* =========================================================
   СОРТИРОВКА PDF
========================================================= */

function compareProductNames(
    a,
    b
) {
    const nameA =
        String(
            a?.item?.productName ||
            ''
        )
            .trim();

    const nameB =
        String(
            b?.item?.productName ||
            ''
        )
            .trim();

    const emptyA =
        nameA === '';

    const emptyB =
        nameB === '';

    if (
        emptyA &&
        !emptyB
    ) {
        return 1;
    }

    if (
        !emptyA &&
        emptyB
    ) {
        return -1;
    }

    const nameCompare =
        nameA.localeCompare(
            nameB,
            'ru',
            {
                sensitivity:
                    'base',
                numeric: true
            }
        );

    if (
        nameCompare !== 0
    ) {
        return nameCompare;
    }

    return (
        a.pageIndex -
        b.pageIndex
    );
}


/* =========================================================
   OCR WORKERS
========================================================= */

async function createOcrWorkers() {
    const count =
        Math.min(
            OCR_WORKERS,
            navigator.hardwareConcurrency
                ? Math.max(
                    1,
                    navigator.hardwareConcurrency - 1
                )
                : OCR_WORKERS
        );

    log(
        `Запускаю OCR-worker: ${count}`
    );

    const workers = [];

    for (
        let i = 0;
        i < count;
        i++
    ) {
        const worker =
            await Tesseract.createWorker(
                OCR_LANG
            );

        await worker.setParameters({
            tessedit_pageseg_mode:
                Tesseract.PSM.SPARSE_TEXT
        });

        workers.push(worker);
    }

    return workers;
}


/* =========================================================
   OCR СТРАНИЦЫ
========================================================= */

async function recognizePdfPage(
    pdf,
    pageIndex,
    orderIds,
    worker
) {
    const {
        canvas,
        ocrCanvas
    } =
        await renderPdfPage(
            pdf,
            pageIndex + 1
        );

    /*
     * -------------------------------------------------------
     * ЭТАП 1.
     *
     * Быстрый OCR верхних 35%.
     * -------------------------------------------------------
     */

    let ocrResult;

    try {
        ocrResult =
            await worker.recognize(
                ocrCanvas
            );
    } catch (error) {
        /*
         * Если быстрый OCR упал,
         * всё равно пробуем полную страницу.
         */
        ocrResult = null;

        console.warn(
            `Быстрый OCR не сработал для страницы ${pageIndex + 1}`,
            error
        );
    }

    let text =
        ocrResult?.data?.text ||
        '';

    let orderId =
        findOrderIdInOcrText(
            text,
            orderIds
        );

    /*
     * -------------------------------------------------------
     * ЭТАП 2.
     *
     * Если сверху номер не найден,
     * OCRируем ВСЮ страницу.
     *
     * При этом findOrderIdInOcrText()
     * всё равно разрешает вернуть
     * ТОЛЬКО номер из реестра 1С.
     * -------------------------------------------------------
     */

    if (!orderId) {
        try {
            ocrResult =
                await worker.recognize(
                    canvas
                );

            text =
                ocrResult?.data?.text ||
                '';

            orderId =
                findOrderIdInOcrText(
                    text,
                    orderIds
                );

        } catch (error) {
            console.warn(
                `Полный OCR не сработал для страницы ${pageIndex + 1}`,
                error
            );
        }
    }

    return {
        pageIndex,
        orderId
    };
}


/* =========================================================
   ОБРАБОТКА PDF
========================================================= */

async function processPdf() {
    if (processingActive) {
        return;
    }

    if (!state.pdfFile) {
        return;
    }

    if (!state.oneC.length) {
        alert(
            'Сначала загрузите реестр 1С.'
        );
        return;
    }

    if (
        !state.lk ||
        !state.lk.SheetNames ||
        !state.lk.SheetNames.length
    ) {
        alert(
            'Сначала загрузите список заказов из ЛК.'
        );
        return;
    }

    if (!state.priceMap.size) {
        alert(
            'Сначала загрузите прайс.'
        );
        return;
    }

    if (!state.mapping.size) {
        rebuildMappingIfPossible();
    }

    const eligibleIds =
        getEligibleOrderIds();

    const excludedQuantityIds =
        getExcludedQuantityOrderIds();

    if (
        !eligibleIds.length &&
        !excludedQuantityIds.length
    ) {
        alert(
            'Нет заказов, готовых к обработке.'
        );
        return;
    }

    processingActive =
        true;

    try {
        clearStatus();


        /* =====================================================
           ЭТАП 1. ЗАГРУЖАЕМ PDF
        ===================================================== */

        setStatus(
            'Загружаю PDF...',
            35
        );

        const pdfBytes =
            new Uint8Array(
                await state.pdfFile.arrayBuffer()
            );

        log(
            `PDF: размер байтов = ${pdfBytes.length}`
        );

        log(
            `PDF: первые байты = ${String.fromCharCode(
                ...pdfBytes.slice(0, 5)
            )}`
        );

        const pdf =
            await pdfjsLib
                .getDocument({
                    data:
                        new Uint8Array(
                            pdfBytes
                        )
                })
                .promise;

        const pageCount =
            pdf.numPages;

        log(
            `Страниц в исходном PDF: ${pageCount}`
        );

        log(
            `Заказов quantity = 1: ${eligibleIds.length}`
        );

        log(
            `Заказов quantity > 1: ${excludedQuantityIds.length}`
        );

        /*
         * ВАЖНО:
         *
         * sourcePdf загружается из исходных байтов.
         * Ничего из PDF предварительно не вырезаем.
         */
        const sourcePdf =
            await PDFLib.PDFDocument.load(
                pdfBytes
            );


        /* =====================================================
           ЭТАП 2. OCR
        ===================================================== */

        setStatus(
            'Запускаю быстрый OCR...',
            36
        );

        const pageActions = [];

        let pagesFound = 0;
        let pagesRemoved = 0;
        let pagesNotFound = 0;

        /*
         * КРИТИЧЕСКИ ВАЖНО:
         *
         * allOrderIds содержит ТОЛЬКО номера
         * из реестра 1С.
         *
         * Поэтому OCR физически не может
         * вернуть номер, которого нет в 1С.
         */
        const allOrderIds =
            getAllOneCOrderIds();

        log(
            `OCR будет искать только номера из реестра 1С: ${allOrderIds.length}`
        );

        const workers =
            await createOcrWorkers();

        try {
            let nextPageIndex = 0;
            let completedPages = 0;

            async function workerLoop(
                worker
            ) {
                while (true) {
                    const pageIndex =
                        nextPageIndex++;

                    if (
                        pageIndex >=
                        pageCount
                    ) {
                        return;
                    }

                    const result =
                        await recognizePdfPage(
                            pdf,
                            pageIndex,
                            allOrderIds,
                            worker
                        );

                    completedPages++;

                    const percent =
                        36 +
                        (
                            completedPages /
                            pageCount
                        ) * 50;

                    setStatus(
                        `Быстрый OCR: ${completedPages} из ${pageCount}...`,
                        percent
                    );

                    pageActions.push(
                        result
                    );
                }
            }

            await Promise.all(
                workers.map(
                    worker =>
                        workerLoop(
                            worker
                        )
                )
            );

        } finally {
            await Promise.all(
                workers.map(
                    worker =>
                        worker.terminate()
                )
            );
        }


        /* =====================================================
           ЭТАП 3. ВОССТАНАВЛИВАЕМ ПОРЯДОК
        ===================================================== */

        pageActions.sort(
            (a, b) =>
                a.pageIndex -
                b.pageIndex
        );

        const classifiedActions =
            [];

        for (
            const result
            of pageActions
        ) {
            const {
                pageIndex,
                orderId
            } =
                result;

            /*
             * Если OCR не нашёл НИ ОДНОГО
             * номера из реестра 1С.
             */
            if (
                !orderId
            ) {
                pagesNotFound++;

                classifiedActions.push({
                    pageIndex,
                    orderId: null,
                    action:
                        'UNKNOWN'
                });

                continue;
            }

            /*
             * Дополнительная страховка:
             *
             * даже если что-то пошло не так,
             * номер обязан существовать
             * в реестре 1С.
             */
            const oneCIdsSet =
                new Set(
                    allOrderIds
                );

            if (
                !oneCIdsSet.has(
                    orderId
                )
            ) {
                pagesNotFound++;

                classifiedActions.push({
                    pageIndex,
                    orderId: null,
                    action:
                        'UNKNOWN'
                });

                continue;
            }

            const item =
                state.mapping.get(
                    orderId
                );

            if (!item) {
                pagesNotFound++;

                classifiedActions.push({
                    pageIndex,
                    orderId,
                    action:
                        'UNKNOWN'
                });

                continue;
            }

            pagesFound++;

            /*
             * quantity > 1:
             * физически не копируем страницу.
             */
            if (
                Number(
                    item.quantity
                ) > 1
            ) {
                classifiedActions.push({
                    pageIndex,
                    orderId,
                    action:
                        'REMOVE',
                    item
                });

                pagesRemoved++;

                continue;
            }

            /*
             * В итоговый PDF попадают только:
             *
             * 1. номер есть в 1С;
             * 2. есть в mapping;
             * 3. FOUND;
             * 4. quantity === 1;
             * 5. не исключён вручную.
             */
            if (
                allOrderIds.includes(
                    item.orderId
                ) &&
                item.status === 'FOUND' &&
                Number(
                    item.quantity
                ) === 1 &&
                !state.manualExcluded.has(
                    item.orderId
                )
            ) {
                classifiedActions.push({
                    pageIndex,
                    orderId,
                    action:
                        'DRAW',
                    item
                });

                continue;
            }

            classifiedActions.push({
                pageIndex,
                orderId,
                action:
                    'UNKNOWN',
                item
            });
        }


        /* =====================================================
           ЭТАП 4. РАЗБИРАЕМ РЕЗУЛЬТАТ
        ===================================================== */

        const drawActions =
            classifiedActions.filter(
                action =>
                    action.action ===
                    'DRAW'
            );

        const removeActions =
            classifiedActions.filter(
                action =>
                    action.action ===
                    'REMOVE'
            );

        const unknownActions =
            classifiedActions.filter(
                action =>
                    action.action ===
                    'UNKNOWN'
            );

        const sortedActions =
            [...drawActions].sort(
                compareProductNames
            );

        const uniqueDrawOrderIds =
            new Set(
                sortedActions.map(
                    action =>
                        action.orderId
                )
            );

        const uniqueRemovedOrderIds =
            new Set(
                removeActions.map(
                    action =>
                        action.orderId
                )
            );

        log('');

        log(
            '========================================'
        );

        log(
            'РЕЗУЛЬТАТ OCR'
        );

        log(
            '========================================'
        );

        log(
            `Страниц распознано: ${pagesFound}`
        );

        log(
            `Уникальных заказов quantity = 1: ${uniqueDrawOrderIds.size}`
        );

        log(
            `Страниц quantity = 1: ${sortedActions.length}`
        );

        log(
            `Уникальных заказов quantity > 1: ${uniqueRemovedOrderIds.size}`
        );

        log(
            `Страниц quantity > 1: ${removeActions.length}`
        );

        log(
            `Страниц, которые не удалось определить: ${unknownActions.length}`
        );

        log(
            '========================================'
        );

        /*
         * Блок проверки.
         *
         * showUnknownOrdersBox() дополнительно
         * фильтрует номера через реестр 1С.
         */
        showUnknownOrdersBox(
            unknownActions
        );


        /* =====================================================
           ЭТАП 5. НОВЫЙ PDF
        ===================================================== */

        setStatus(
            'Собираю PDF в алфавитном порядке...',
            90
        );

        const outputPdf =
            await PDFLib.PDFDocument.create();

        const copiedPages =
            await outputPdf.copyPages(
                sourcePdf,
                sortedActions.map(
                    action =>
                        action.pageIndex
                )
            );

        for (
            const copiedPage
            of copiedPages
        ) {
            outputPdf.addPage(
                copiedPage
            );
        }


        /* =====================================================
           ЭТАП 6. НАНОСИМ НАИМЕНОВАНИЯ
        ===================================================== */

        setStatus(
            'Наношу наименования...',
            93
        );

        let font = null;

        if (
            sortedActions.length
        ) {
            const fontkit =
                await getFontkit();

            outputPdf.registerFontkit(
                fontkit
            );

            const fontBytes =
                await getFontBytes();

            font =
                await outputPdf.embedFont(
                    fontBytes,
                    {
                        subset: true
                    }
                );
        }

        for (
            let i = 0;
            i < sortedActions.length;
            i++
        ) {
            const action =
                sortedActions[i];

            const page =
                outputPdf.getPage(
                    i
                );

            drawLabel(
                page,
                font,
                action.item.productName,
                state.textConfig
            );
        }


        /* =====================================================
           ЭТАП 7. СОХРАНЕНИЕ
        ===================================================== */

        setStatus(
            'Сохраняю готовый PDF...',
            98
        );

        const outputBytes =
            await outputPdf.save({
                useObjectStreams:
                    true
            });

        state.lastOutputPdf =
            new Blob(
                [outputBytes],
                {
                    type:
                        'application/pdf'
                }
            );

        state.processedPages =
            sortedActions.map(
                action => ({
                    pageIndex:
                        action.pageIndex,
                    orderId:
                        action.orderId,
                    productName:
                        action.item.productName,
                    sku:
                        action.item.sku,
                    quantity:
                        action.item.quantity
                })
            );


        /* =====================================================
           ИТОГ
        ===================================================== */

        log('');

        log(
            '========================================'
        );

        log(
            'ГОТОВО'
        );

        log(
            '========================================'
        );

        log(
            `Уникальных заказов в готовом PDF: ${uniqueDrawOrderIds.size}`
        );

        log(
            `Этикеток в готовом PDF: ${sortedActions.length}`
        );

        log(
            `Удалено уникальных заказов quantity > 1: ${uniqueRemovedOrderIds.size}`
        );

        log(
            `Удалено страниц quantity > 1: ${removeActions.length}`
        );

        log(
            `Не попало в готовый PDF из-за нераспознанных/неподходящих страниц: ${unknownActions.length}`
        );

        log(
            `Размер готового PDF: ${(outputBytes.length / 1024 / 1024).toFixed(2)} МБ`
        );

        log(
            'Порядок PDF: по наименованию товара.'
        );

        log(
            'Названия нанесены ПОСЛЕ сортировки страниц.'
        );

        log(
            'В PDF допускаются только номера, существующие в реестре 1С.'
        );

        setStatus(
            'Готово',
            100
        );

        updateDownloadButtons();

    } catch (error) {
        console.error(error);

        log('');

        log(
            'ОШИБКА: ' +
            (
                error?.message ||
                error
            )
        );

        setStatus(
            'Ошибка обработки',
            0
        );

        alert(
            'Ошибка обработки PDF:\n\n' +
            (
                error?.message ||
                error
            )
        );

    } finally {
        processingActive =
            false;
    }
}


/* =========================================================
   EXCEL: ЗАКАЗЫ + SKU + НАИМЕНОВАНИЕ
========================================================= */

function createOrdersAndNamesWorkbook() {
    const dataRows = [];

    for (
        const oneCRow
        of state.oneC
    ) {
        const item =
            state.mapping.get(
                oneCRow.orderId
            );

        if (!item) {
            dataRows.push({
                orderId:
                    oneCRow.orderId,
                quantity:
                    oneCRow.quantity,
                sku: '',
                productName: ''
            });

            continue;
        }

        if (
            item.status === 'FOUND' &&
            Array.isArray(item.skus) &&
            item.skus.length
        ) {
            for (
                let i = 0;
                i < item.skus.length;
                i++
            ) {
                const sku =
                    item.skus[i];

                const productName =
                    item.productNames?.[i] ||
                    state.priceMap.get(
                        sku
                    ) ||
                    '';

                dataRows.push({
                    orderId:
                        item.orderId,
                    quantity:
                        item.quantity,
                    sku,
                    productName
                });
            }

            continue;
        }

        dataRows.push({
            orderId:
                item.orderId,
            quantity:
                item.quantity,
            sku: '',
            productName: ''
        });
    }


    /* =====================================================
       СОРТИРОВКА
    ===================================================== */

    dataRows.sort(
        (a, b) => {
            const nameA =
                String(
                    a.productName ||
                    ''
                ).trim();

            const nameB =
                String(
                    b.productName ||
                    ''
                ).trim();

            const emptyA =
                nameA === '';

            const emptyB =
                nameB === '';

            if (
                emptyA &&
                !emptyB
            ) {
                return 1;
            }

            if (
                !emptyA &&
                emptyB
            ) {
                return -1;
            }

            const result =
                nameA.localeCompare(
                    nameB,
                    'ru',
                    {
                        sensitivity:
                            'base',
                        numeric:
                            true
                    }
                );

            if (
                result !== 0
            ) {
                return result;
            }

            const orderCompare =
                String(
                    a.orderId
                ).localeCompare(
                    String(
                        b.orderId
                    ),
                    'ru',
                    {
                        numeric:
                            true
                    }
                );

            if (
                orderCompare !== 0
            ) {
                return orderCompare;
            }

            return String(
                a.sku
            ).localeCompare(
                String(
                    b.sku
                ),
                'ru',
                {
                    numeric:
                        true
                }
            );
        }
    );


    /* =====================================================
       РОВНО 3 КОЛОНКИ
    ===================================================== */

    const rows = [
        [
            'Номер заказа',
            'SKU',
            'Наименование'
        ]
    ];

    for (
        const item
        of dataRows
    ) {
        rows.push([
            item.orderId,
            item.sku,
            item.productName
        ]);
    }


    const worksheet =
        XLSX.utils.aoa_to_sheet(
            rows
        );


    /* =====================================================
       ЗАГОЛОВКИ
    ===================================================== */

    for (
        let c = 0;
        c < rows[0].length;
        c++
    ) {
        const cell =
            worksheet[
                XLSX.utils.encode_cell({
                    r: 0,
                    c
                })
            ];

        if (cell) {
            cell.s = {
                font: {
                    bold: true
                }
            };
        }
    }


    /* =====================================================
       quantity > 1 — ЖЁЛТЫЕ СТРОКИ
    ===================================================== */

    for (
        let r = 1;
        r < rows.length;
        r++
    ) {
        const orderId =
            normalizeOrderId(
                rows[r][0]
            );

        const item =
            state.mapping.get(
                orderId
            );

        if (
            item &&
            Number(
                item.quantity
            ) > 1
        ) {
            for (
                let c = 0;
                c < rows[r].length;
                c++
            ) {
                const cell =
                    worksheet[
                        XLSX.utils.encode_cell({
                            r,
                            c
                        })
                    ];

                if (cell) {
                    cell.s = {
                        fill: {
                            fgColor: {
                                rgb:
                                    'FFF2CC'
                            }
                        },
                        font: {
                            color: {
                                rgb:
                                    '7F6000'
                            }
                        }
                    };
                }
            }
        }
    }


    worksheet['!cols'] = [
        {
            wch: 22
        },
        {
            wch: 18
        },
        {
            wch: 70
        }
    ];


    const workbook =
        XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(
        workbook,
        worksheet,
        'Заказы и наименования'
    );

    const output =
        XLSX.write(
            workbook,
            {
                bookType:
                    'xlsx',
                type:
                    'array'
            }
        );

    return new Blob(
        [output],
        {
            type:
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }
    );
}


/* =========================================================
   СКАЧИВАНИЕ
========================================================= */

function downloadWithPrefix(
    blob,
    filename
) {
    let finalName =
        String(filename);

    const prefix =
        window.DOWNLOAD_PREFIX ||
        'шк-';

    if (
        !finalName.startsWith(
            prefix
        )
    ) {
        finalName =
            prefix +
            finalName;
    }

    const url =
        URL.createObjectURL(
            blob
        );

    const a =
        document.createElement(
            'a'
        );

    a.href =
        url;

    a.download =
        finalName;

    document.body.appendChild(
        a
    );

    a.click();

    a.remove();

    setTimeout(
        () =>
            URL.revokeObjectURL(
                url
            ),
        1000
    );
}


function downloadOrdersAndNames() {
    if (
        !state.oneC.length ||
        !state.mapping.size
    ) {
        alert(
            'Сначала загрузите реестр 1С, список ЛК и прайс.'
        );

        return;
    }

    const blob =
        createOrdersAndNamesWorkbook();

    state.lastOrdersXlsx =
        blob;

    downloadWithPrefix(
        blob,
        'заказы и наименования.xlsx'
    );
}


function downloadLabels() {
    if (
        !state.lastOutputPdf
    ) {
        alert(
            'Готовый PDF ещё не сформирован.'
        );

        return;
    }

    downloadWithPrefix(
        state.lastOutputPdf,
        'готовые-этикетки.pdf'
    );
}


/* =========================================================
   КНОПКИ
========================================================= */

function updateDownloadButtons() {
    const ordersBtn =
        document.getElementById(
            'downloadOrdersNamesBtn'
        );

    const labelsBtn =
        document.getElementById(
            'downloadLabelsBtn'
        );

    if (ordersBtn) {
        ordersBtn.disabled =
            !(
                state.oneC.length &&
                state.lk &&
                state.lk.SheetNames &&
                state.lk.SheetNames.length &&
                state.priceMap.size &&
                state.mapping.size
            );
    }

    if (labelsBtn) {
        labelsBtn.disabled =
            !state.lastOutputPdf;
    }
}


/* =========================================================
   ГАЛОЧКА ПОСЛЕ ЗАГРУЗКИ
========================================================= */

function markFileLoaded(
    input,
    file
) {
    if (!input) return;

    let container =
        input.closest('label');

    if (
        !container &&
        input.id
    ) {
        try {
            container =
                document.querySelector(
                    `label[for="${CSS.escape(input.id)}"]`
                );
        } catch (error) {
            container = null;
        }
    }

    if (!container) {
        container =
            input.parentElement;
    }

    if (!container) return;

    const oldMark =
        container.querySelector(
            '.file-loaded-mark'
        );

    if (oldMark) {
        oldMark.remove();
    }

    const mark =
        document.createElement(
            'span'
        );

    mark.className =
        'file-loaded-mark';

    mark.textContent =
        ' ✓';

    mark.title =
        `Загружен: ${file.name}`;

    mark.style.cssText = `
        color: #16a34a;
        font-weight: 800;
        font-size: 20px;
        margin-left: 8px;
        display: inline-block;
        vertical-align: middle;
        line-height: 1;
    `;

    container.appendChild(
        mark
    );
}


/* =========================================================
   ЛИЛОВЫЕ ПЛАШКИ
   HTML НЕ ТРОГАЕМ
========================================================= */

function styleFileCards() {
    const inputIds = [
        'orders1cFile',
        'lkFile',
        'priceFile',
        'pdfFile'
    ];

    for (
        const inputId
        of inputIds
    ) {
        const input =
            document.getElementById(
                inputId
            );

        if (!input) {
            continue;
        }

        let container =
            input.closest('label');

        if (
            !container &&
            input.id
        ) {
            try {
                container =
                    document.querySelector(
                        `label[for="${CSS.escape(input.id)}"]`
                    );
            } catch (error) {
                container = null;
            }
        }

        if (!container) {
            container =
                input.parentElement;
        }

        if (!container) {
            continue;
        }

        container.style.border =
            '2px solid #b58cff';

        container.style.borderRadius =
            '14px';

        container.style.boxShadow =
            '0 0 0 1px rgba(181, 140, 255, 0.12), 0 4px 14px rgba(120, 80, 180, 0.08)';

        container.style.transition =
            'border-color .2s ease, box-shadow .2s ease';

        container.addEventListener(
            'mouseenter',
            function () {
                container.style.borderColor =
                    '#9b6cff';

                container.style.boxShadow =
                    '0 0 0 3px rgba(181, 140, 255, 0.15), 0 6px 18px rgba(120, 80, 180, 0.12)';
            }
        );

        container.addEventListener(
            'mouseleave',
            function () {
                container.style.borderColor =
                    '#b58cff';

                container.style.boxShadow =
                    '0 0 0 1px rgba(181, 140, 255, 0.12), 0 4px 14px rgba(120, 80, 180, 0.08)';
            }
        );
    }
}


/* =========================================================
   СОБЫТИЯ
========================================================= */

function bindInput(
    inputId,
    handler
) {
    const input =
        document.getElementById(
            inputId
        );

    if (!input) {
        console.warn(
            `Не найден input #${inputId}`
        );

        return;
    }

    input.addEventListener(
        'change',
        async function () {
            const file =
                this.files?.[0];

            if (!file) {
                return;
            }

            try {
                await handler(file);

                markFileLoaded(
                    this,
                    file
                );

                updateDownloadButtons();

            } catch (error) {
                console.error(
                    error
                );

                log(
                    'ОШИБКА: ' +
                    (
                        error?.message ||
                        error
                    )
                );

                alert(
                    error?.message ||
                    error
                );
            }
        }
    );
}


/* =========================================================
   ЗАПУСК
========================================================= */

document.addEventListener(
    'DOMContentLoaded',
    function () {
        bindInput(
            'orders1cFile',
            handleOneCFile
        );

        bindInput(
            'lkFile',
            handleLKFile
        );

        bindInput(
            'priceFile',
            handlePriceFile
        );

        bindInput(
            'pdfFile',
            handlePdfFile
        );

        const ordersBtn =
            document.getElementById(
                'downloadOrdersNamesBtn'
            );

        const labelsBtn =
            document.getElementById(
                'downloadLabelsBtn'
            );

        if (ordersBtn) {
            ordersBtn.addEventListener(
                'click',
                downloadOrdersAndNames
            );
        }

        if (labelsBtn) {
            labelsBtn.addEventListener(
                'click',
                downloadLabels
            );
        }

        styleFileCards();

        updateDownloadButtons();
    }
);


/* =========================================================
   ПЕРЕД ЗАКРЫТИЕМ
========================================================= */

window.addEventListener(
    'beforeunload',
    function () {
        state.lastOutputPdf = null;
        state.lastOrdersXlsx = null;
    }
);