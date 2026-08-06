console.log('🚀 make-price.js запущен...');
const fs = require('fs');
const path = require('path');

// Ищем библиотеку xlsx в нескольких местах
let XLSX = null;
const candidates = [
    __dirname,
    path.join(__dirname, '..', '..', '..', '..', 'marketplace-label-processor'),
    'C:\\Users\\Admin\\node_modules'
];
for (const p of candidates) {
    try {
        XLSX = require(require.resolve('xlsx', { paths: [p] }));
        console.log('📦 xlsx найден в: ' + p);
        break;
    } catch (e) { /* пробуем следующее место */ }
}
if (!XLSX) {
    console.log('❌ Не найдена библиотека xlsx ни в одном из мест. Пришлите скриншот этого окна.');
    process.exit(1);
}

const file = process.argv[2];
if (!file) {
    console.log('Как использовать: node make-price.js "Прайс.xlsx"');
    process.exit(1);
}

try {
    function normalizeSku(value) {
        if (value === null || value === undefined) return '';
        let s = String(value);
        if (s.endsWith('.0')) s = s.slice(0, -2);
        return s.replace(/[\s\u00A0\u2007\u202F\u2009\u200A\u205F\u3000]/g, '').toLowerCase();
    }

    console.log('📄 Читаю файл: ' + file);
    const wb = XLSX.readFile(path.resolve(file));
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '', blankrows: false });
    console.log('📄 Строк в файле: ' + rows.length);

    const map = {};
    let count = 0;
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
        if (sku) { map[sku] = name; count++; }
    }

    fs.writeFileSync(path.join(__dirname, 'price-data.js'), 'window.DEFAULT_PRICE_MAP = ' + JSON.stringify(map) + ';\n');
    console.log('✅ Создан price-data.js: ' + count + ' товаров');
} catch (e) {
    console.log('❌ Ошибка: ' + e.message);
}