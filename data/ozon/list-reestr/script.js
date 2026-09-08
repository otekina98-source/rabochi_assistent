let listNumbers = [];
let reestrNumbers = [];
let results = [];

// Drag & drop
const uploadArea = document.getElementById('pdfUploadArea');
const pdfInput = document.getElementById('pdfFileInput');

uploadArea.addEventListener('dragover', e => { e.preventDefault(); uploadArea.classList.add('dragover'); });
uploadArea.addEventListener('dragleave', () => uploadArea.classList.remove('dragover'));
uploadArea.addEventListener('drop', e => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    if (file && file.type === 'application/pdf') handlePdf(file);
});

pdfInput.addEventListener('change', e => { if (e.target.files[0]) handlePdf(e.target.files[0]); });

async function handlePdf(file) {
    document.getElementById('pdfFileName').textContent = file.name;
    const status = document.getElementById('pdfStatus');
    status.textContent = '⏳ Чтение PDF...';
    status.className = 'status';
    
    try {
        const buffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
        let text = '';
        
        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const content = await page.getTextContent();
            text += content.items.map(it => it.str).join(' ') + '\n';
            
            status.textContent = `⏳ Страница ${i} из ${pdf.numPages}`;
            // Отдаем управление браузеру каждые 5 страниц, чтобы не было зависания
            if (i % 5 === 0) await new Promise(r => setTimeout(r, 0));
        }
        
        // 🚀 КРИТИЧЕСКИ ВАЖНО: убираем пробелы вокруг дефисов между цифрами
        // Превращает "84486774 - 0082 - 1" в "84486774-0082-1"
        text = text.replace(/(\d)\s+[-–—]\s+(\d)/g, '$1-$2');
        text = text.replace(/(\d)\s+[-–—]\s+(\d)/g, '$1-$2'); // Повторяем для надежности
        
        // 🚀 ГИБКИЙ ПАТТЕРН: ищем ЛЮБЫЕ номера с двумя дефисами (без жестких ограничений длины)
        const pattern = /\b\d+[-–—]\d+[-–—]\d+\b/g;
        const matches = text.match(pattern) || [];
        
        // Нормализуем: приводим все типы дефисов к обычному "-" и убираем дубликаты
        listNumbers = [...new Set(matches.map(n => n.replace(/[–—]/g, '-')))];
        
        status.textContent = `✅ Найдено ${listNumbers.length} номеров на ${pdf.numPages} стр.`;
        status.className = 'status success';
        console.log('Найдено номеров:', listNumbers.length);
        console.log('Примеры:', listNumbers.slice(0, 10));
        
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
    
    // 🚀 Парсим и нормализуем номера из 1С (убираем пробелы, приводим дефисы к одному виду)
    reestrNumbers = [...new Set(
        text.split(/[\r\n]+/)
            .map(n => n.trim().replace(/\s+/g, '').replace(/[–—]/g, '-'))
            .filter(n => n.length > 0)
    )];
    
    document.getElementById('progressSection').classList.remove('hidden');
    document.getElementById('resultsSection').classList.add('hidden');
    document.getElementById('startBtn').disabled = true;
    
    // 🚀 ОПТИМИЗАЦИЯ: используем Set для мгновенного поиска O(1) вместо медленного .includes() O(N)
    const listSet = new Set(listNumbers);
    const reestrSet = new Set(reestrNumbers);
    const all = new Set([...listSet, ...reestrSet]);
    const total = all.size;
    let done = 0;
    results = [];
    
    for (const num of all) {
        const inList = listSet.has(num);
        const inReestr = reestrSet.has(num);
        results.push({ number: num, inList, inReestr });
        done++;
        
        // Обновляем прогресс-бар каждые 50 шагов, чтобы не тормозить браузер
        if (done % 50 === 0 || done === total) {
            const pct = Math.round((done / total) * 100);
            document.getElementById('progressFill').style.width = pct + '%';
            document.getElementById('progressPercent').textContent = pct + '%';
            document.getElementById('progressDetails').textContent = 
                `Обработано ${done}/${total} | Лист: ${listNumbers.length} | 1С: ${reestrNumbers.length}`;
            await new Promise(r => setTimeout(r, 0));
        }
    }
    
    showResults();
    document.getElementById('startBtn').disabled = false;
}

function showResults() {
    const match = results.filter(r => r.inList && r.inReestr).length;
    const onlyList = results.filter(r => r.inList && !r.inReestr).length;
    const onlyReestr = results.filter(r => !r.inList && r.inReestr).length;
    const totalDiff = onlyList + onlyReestr;
    
    document.getElementById('summary').innerHTML = `
        <div class="summary-card miss">
            <span class="num">${totalDiff}</span>
            <div class="lbl">Всего расхождений</div>
        </div>
        <div class="summary-card miss">
            <span class="num">${onlyList}</span>
            <div class="lbl">Только в листе</div>
        </div>
        <div class="summary-card miss">
            <span class="num">${onlyReestr}</span>
            <div class="lbl">Только в 1С</div>
        </div>
    `;
    
    const diff = results.filter(r => !(r.inList && r.inReestr));
    
    if (diff.length === 0) {
        document.getElementById('resultsBody').innerHTML = `
            <tr><td colspan="4" style="text-align:center; padding:30px; color:#28a745; font-family:-apple-system, sans-serif;">
                ✅ Все номера совпадают! Расхождений нет.
            </td></tr>
        `;
    } else {
        const sorted = [...diff].sort((a, b) => {
            if (a.inList && !b.inList) return -1;
            if (!a.inList && b.inList) return 1;
            return 0;
        });
        
        document.getElementById('resultsBody').innerHTML = sorted.map(r => {
            let badge;
            if (r.inList && !r.inReestr) badge = '⚠️ Только в листе';
            else badge = '⚠️ Только в 1С';
            return `<tr>
                <td><span class="num-copy" title="Нажмите, чтобы скопировать" onclick="copyNumber('${r.number}', this)">${r.number}</span></td>
                <td>${r.inList ? '✅' : '❌'}</td>
                <td>${r.inReestr ? '✅' : '❌'}</td>
                <td><span class="badge miss">${badge}</span></td>
            </tr>`;
        }).join('');
    }
    
    document.getElementById('resultsSection').classList.remove('hidden');
}

// --- Функции копирования и экспорта ---
function copyNumber(num, el) {
    navigator.clipboard.writeText(num).then(() => {
        el.classList.add('copied');
        showToast('Скопировано: ' + num);
        setTimeout(() => el.classList.remove('copied'), 1200);
    });
}

function copyAllNumbers() {
    const diff = results.filter(r => !(r.inList && r.inReestr));
    if (diff.length === 0) { showToast('Нет номеров для копирования'); return; }
    const text = diff.map(r => r.number).join('\n');
    navigator.clipboard.writeText(text).then(() => {
        showToast('Скопировано ' + diff.length + ' номеров');
    });
}

function showToast(msg) {
    // Создаем toast, если его нет в HTML
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
    const diff = results.filter(r => !(r.inList && r.inReestr));
    const csv = [
        'Номер;В листе;В 1С;Статус',
        ...diff.map(r => `${r.number};${r.inList?'Да':'Нет'};${r.inReestr?'Да':'Нет'};${r.inList?'Только в листе':'Только в 1С'}`)
    ].join('\n');
    
    const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'raschozhdeniya_ozon_list.csv';
    a.click();
}