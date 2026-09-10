let listNumbers = [];
let reestrNumbers = [];
let results = [];

// Drag & drop
const uploadArea = document.getElementById('pdfUploadArea');
const pdfInput = document.getElementById('pdfFileInput');

uploadArea.addEventListener('dragover', e => {
    e.preventDefault();
    uploadArea.classList.add('dragover');
});
uploadArea.addEventListener('dragleave', () => uploadArea.classList.remove('dragover'));
uploadArea.addEventListener('drop', e => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    if (file && file.type === 'application/pdf') handlePdf(file);
});
pdfInput.addEventListener('change', e => {
    if (e.target.files[0]) handlePdf(e.target.files[0]);
});

// Счётчик номеров в textarea
const reestrText = document.getElementById('reestrText');
const reestrCounter = document.getElementById('reestrCounter');

function updateCounter() {
    const nums = reestrText.value
        .split(/[\r\n]+/)
        .map(n => n.trim().replace(/\s+/g, '').replace(/[–—]/g, '-'))
        .filter(n => /^\d+-\d+-\d+(-\d+)?$/.test(n));
    reestrCounter.textContent = nums.length + ' ' + pluralize(nums.length, 'номер', 'номера', 'номеров');
}
reestrText.addEventListener('input', updateCounter);

function pluralize(n, one, two, five) {
    let abs = Math.abs(n) % 100;
    if (abs >= 11 && abs <= 19) return five;
    abs = abs % 10;
    if (abs === 1) return one;
    if (abs >= 2 && abs <= 4) return two;
    return five;
}

// 🚀 УМНАЯ СБОРКА ТЕКСТА С УЧЁТОМ ПОЗИЦИЙ ЭЛЕМЕНТОВ PDF
async function handlePdf(file) {
    document.getElementById('pdfFileName').textContent = file.name;
    const status = document.getElementById('pdfStatus');
    status.textContent = ' Чтение PDF...';
    status.className = 'status';

    try {
        const buffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
        let text = '';

        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const content = await page.getTextContent();
            
            // Умная сборка текста с учётом позиций элементов
            const items = content.items;
            let pageText = '';
            
            for (let j = 0; j < items.length; j++) {
                const item = items[j];
                const str = item.str;
                
                if (j > 0) {
                    const prevItem = items[j-1];
                    // Вычисляем расстояние между элементами по X-координате
                    const prevEndX = prevItem.transform[4] + (prevItem.width || 0);
                    const currStartX = item.transform[4];
                    const gap = currStartX - prevEndX;
                    
                    // Если элементы близко (< 3 единицы), склеиваем без пробела
                    // Если чуть дальше — добавляем пробел
                    if (gap < 3) {
                        pageText += str;
                    } else {
                        pageText += ' ' + str;
                    }
                } else {
                    pageText += str;
                }
            }
            
            text += pageText + '\n';
            status.textContent = `⏳ Страница ${i} из ${pdf.numPages}`;
            if (i % 5 === 0) await new Promise(r => setTimeout(r, 0));
        }

        // 🚀 АГРЕССИВНАЯ НОРМАЛИЗАЦИЯ: убираем все пробелы вокруг дефисов
        text = text.replace(/(\d)\s+([-–—])\s*/g, '$1$2');
        text = text.replace(/\s*([-–—])\s*(\d)/g, '$1$2');
        text = text.replace(/([-–—])\s+([-–—])/g, '$1');
        // Убираем множественные пробелы
        text = text.replace(/\s{2,}/g, ' ');

        // Отладка: выводим первые 2000 символов в консоль
        console.log('=== ПЕРВЫЕ 2000 СИМВОЛОВ ТЕКСТА ИЗ PDF ===');
        console.log(text.substring(0, 2000));
        console.log('===========================================');

        // Ищем номера: 3 или 4 группы цифр через дефис
        const pattern = /\d+[-–—]\d+[-–—]\d+(?:[-–—]\d+)?/g;
        const matches = text.match(pattern) || [];

        listNumbers = [...new Set(matches.map(n => n.replace(/[–—]/g, '-')))];

        status.textContent = `✅ Найдено ${listNumbers.length} номеров на ${pdf.numPages} стр.`;
        status.className = 'status success';
        console.log('Найдено номеров:', listNumbers.length);
        console.log('Примеры:', listNumbers.slice(0, 10));
        
        // Если не найдено номеров, показываем предупреждение
        if (listNumbers.length === 0) {
            alert('⚠️ Номера не найдены! Проверьте консоль браузера (F12) — там показан пример текста из PDF.');
        }
    } catch (err) {
        status.textContent = '❌ Ошибка: ' + err.message;
        status.className = 'status error';
        console.error(err);
    }
}

async function startSverka() {
    const text = document.getElementById('reestrText').value.trim();
    if (!text) { alert('Введите номера из 1С'); return; }
    if (listNumbers.length === 0) { alert('Загрузите Лист отгрузки'); return; }

    reestrNumbers = [...new Set(
        text.split(/[\r\n]+/)
            .map(n => n.trim().replace(/\s+/g, '').replace(/[–—]/g, '-'))
            .filter(n => /^\d+-\d+-\d+(-\d+)?$/.test(n))
    )];

    if (reestrNumbers.length === 0) {
        alert('Не найдено номеров в формате «цифры-цифры-цифры» или «цифры-цифры-цифры-цифры»');
        return;
    }

    document.getElementById('progressSection').classList.remove('hidden');
    document.getElementById('resultsSection').classList.add('hidden');
    document.getElementById('startBtn').disabled = true;

    results = [];
    const matchedFrom1C = new Set();
    const matchedPdfIdx = new Set();

    const total = listNumbers.length;
    let done = 0;

    for (let i = 0; i < listNumbers.length; i++) {
        const pdfNum = listNumbers[i];
        let matched = false;

        for (const reestrNum of reestrNumbers) {
            if (pdfNum === reestrNum ||
                pdfNum.includes(reestrNum) ||
                reestrNum.includes(pdfNum)) {
                matchedFrom1C.add(reestrNum);
                matchedPdfIdx.add(i);
                matched = true;
                break;
            }
        }

        if (!matched) {
            results.push({
                number: pdfNum,
                inList: true,
                inReestr: false,
                status: 'Есть в листе отгрузки, нет в 1С'
            });
        }

        done++;
        if (done % 50 === 0 || done === total) {
            const pct = Math.round((done / total) * 100);
            document.getElementById('progressFill').style.width = pct + '%';
            document.getElementById('progressPercent').textContent = pct + '%';
            document.getElementById('progressDetails').textContent =
                `Обработано ${done}/${total}`;
            await new Promise(r => setTimeout(r, 0));
        }
    }

    for (const reestrNum of reestrNumbers) {
        if (!matchedFrom1C.has(reestrNum)) {
            results.push({
                number: reestrNum,
                inList: false,
                inReestr: true,
                status: 'Есть в 1С, нет в листе отгрузки'
            });
        }
    }

    showResults();
    document.getElementById('startBtn').disabled = false;
}

function showResults() {
    const onlyList = results.filter(r => r.inList && !r.inReestr).length;
    const onlyReestr = results.filter(r => !r.inList && r.inReestr).length;
    const totalDiff = onlyList + onlyReestr;
    const totalChecked = reestrNumbers.length;
    const matched = totalChecked - onlyReestr;

    document.getElementById('summary').innerHTML = `
        <div class="summary-card ok">
            <span class="num">${matched}</span>
            <div class="lbl">${pluralize(matched, 'номер совпадает', 'номера совпадают', 'номеров совпадают')}</div>
        </div>
        <div class="summary-card miss">
            <span class="num">${onlyReestr}</span>
            <div class="lbl">Есть в 1С, нет в листе</div>
        </div>
        <div class="summary-card miss">
            <span class="num">${onlyList}</span>
            <div class="lbl">Есть в листе, нет в 1С</div>
        </div>
    `;

    const body = document.getElementById('resultsBody');
    if (results.length === 0) {
        body.innerHTML = `
            <tr><td colspan="4" style="text-align:center; padding:30px; color:#28a745; font-family:-apple-system, sans-serif;">
                ✅ Расхождений не найдено.
            </td></tr>
        `;
    } else {
        const sorted = [...results].sort((a, b) => {
            if (a.inReestr && !b.inReestr) return -1;
            if (!a.inReestr && b.inReestr) return 1;
            return 0;
        });

        body.innerHTML = sorted.map(r => {
            const badgeClass = r.inReestr ? 'miss-reestr' : 'miss-list';
            return `<tr>
                <td><span class="num-copy" title="Нажмите, чтобы скопировать" onclick="copyNumber('${r.number}', this)">${r.number}</span></td>
                <td>${r.inList ? '✅' : '❌'}</td>
                <td>${r.inReestr ? '✅' : '❌'}</td>
                <td><span class="badge ${badgeClass}">${r.status}</span></td>
            </tr>`;
        }).join('');
    }

    document.getElementById('resultsSection').classList.remove('hidden');
}

function copyNumber(num, el) {
    navigator.clipboard.writeText(num).then(() => {
        el.classList.add('copied');
        showToast('Скопировано: ' + num);
        setTimeout(() => el.classList.remove('copied'), 1200);
    });
}

function copyAllNumbers() {
    if (results.length === 0) { showToast('Нет номеров для копирования'); return; }
    const text = results.map(r => r.number).join('\n');
    navigator.clipboard.writeText(text).then(() => {
        showToast('Скопировано ' + results.length + ' ' + pluralize(results.length, 'номер', 'номера', 'номеров'));
    });
}

function showToast(msg) {
    let t = document.getElementById('toast');
    if (!t) {
        t = document.createElement('div');
        t.id = 'toast';
        t.className = 'toast';
        document.body.appendChild(t);
    }
    t.textContent = msg;
    t.classList.add('show');
    clearTimeout(t._timer);
    t._timer = setTimeout(() => t.classList.remove('show'), 1800);
}

function exportCSV() {
    if (results.length === 0) { showToast('Нет данных для экспорта'); return; }
    const csv = [
        'Номер;В 1С;В листе отгрузки;Статус',
        ...results.map(r =>
            `${r.number};${r.inReestr ? 'Да' : 'Нет'};${r.inList ? 'Да' : 'Нет'};${r.status}`
        )
    ].join('\n');
    const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'raschozhdeniya_ozon.csv';
    a.click();
}