let aktNumbers = [];
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
            status.textContent = ` Страница ${i} из ${pdf.numPages}`;
        }
        // Ищем номера заказов (10-13 цифр)
        const matches = text.match(/\b\d{10,13}\b/g) || [];
        aktNumbers = [...new Set(matches)];
        status.textContent = `✅ Найдено ${aktNumbers.length} номеров на ${pdf.numPages} стр.`;
        status.className = 'status success';
    } catch (err) {
        status.textContent = '❌ Ошибка: ' + err.message;
        status.className = 'status error';
    }
}

function toggleSource() {
    const src = document.querySelector('input[name="src"]:checked').value;
    document.getElementById('reestrFileBtn').style.display = src === 'file' ? 'inline-block' : 'none';
}

function loadReestrFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    document.getElementById('reestrFileName').textContent = file.name;
    const reader = new FileReader();
    reader.onload = ev => document.getElementById('reestrText').value = ev.target.result;
    reader.readAsText(file);
}

async function startSverka() {
    const text = document.getElementById('reestrText').value.trim();
    if (!text) { alert('Введите реестр'); return; }
    if (aktNumbers.length === 0) { alert('Загрузите PDF с актом'); return; }
    reestrNumbers = [...new Set(text.split(/[\n,;\s\t]+/).filter(n => n.trim()))];
    
    document.getElementById('progressSection').classList.remove('hidden');
    document.getElementById('progressSection').classList.add('processing'); // ← анимация
    document.getElementById('resultsSection').classList.add('hidden');
    document.getElementById('startBtn').disabled = true;
    
    const all = new Set([...aktNumbers, ...reestrNumbers]);
    const total = all.size;
    let done = 0;
    results = [];
    
    for (const num of all) {
        const inAkt = aktNumbers.includes(num);
        const inReestr = reestrNumbers.includes(num);
        results.push({ number: num, inAkt, inReestr });
        done++;
        const pct = Math.round((done / total) * 100);
        document.getElementById('progressFill').style.width = pct + '%';
        document.getElementById('progressPercent').textContent = pct + '%';
        document.getElementById('progressDetails').textContent = 
            `Обработано ${done}/${total} | Акт: ${aktNumbers.length} | Реестр: ${reestrNumbers.length}`;
        if (done % 10 === 0) await new Promise(r => setTimeout(r, 10));
    }
    
    document.getElementById('progressSection').classList.remove('processing'); // ← убираем анимацию
    showResults();
    document.getElementById('startBtn').disabled = false;
}

function showResults() {
    const match = results.filter(r => r.inAkt && r.inReestr).length;
    const onlyAkt = results.filter(r => r.inAkt && !r.inReestr).length;
    const onlyReestr = results.filter(r => !r.inAkt && r.inReestr).length;
    const totalDiff = onlyAkt + onlyReestr;
    
    document.getElementById('summary').innerHTML = `
        <div class="summary-card miss">
            <span class="num">${totalDiff}</span>
            <div class="lbl">Всего расхождений</div>
        </div>
        <div class="summary-card" style="border-color: #9B59B6; background: #FAF0FF;">
            <span class="num" style="color: #9B59B6;">${onlyAkt}</span>
            <div class="lbl">Только в акте</div>
        </div>
        <div class="summary-card" style="border-color: #dc3545; background: #fff5f5;">
            <span class="num" style="color: #dc3545;">${onlyReestr}</span>
            <div class="lbl">Только в реестре</div>
        </div>
    `;
    
    // Показываем ТОЛЬКО расхождения
    const diff = results.filter(r => !(r.inAkt && r.inReestr));
    if (diff.length === 0) {
        document.getElementById('resultsBody').innerHTML = `
            <tr><td colspan="4" style="text-align:center; padding:30px; color:#28a745; font-family:-apple-system, sans-serif;">
                ✅ Все номера совпадают! Расхождений нет.
            </td></tr>
        `;
    } else {
        const sorted = [...diff].sort((a, b) => {
            if (a.inAkt && !b.inAkt) return -1;
            if (!a.inAkt && b.inAkt) return 1;
            return 0;
        });
        document.getElementById('resultsBody').innerHTML = sorted.map((r, idx) => {
            let badgeClass, badgeText, rowClass;
            
            if (r.inAkt && !r.inReestr) {
                badgeClass = 'only_akt';
                badgeText = '🟣 Только в акте';
                rowClass = 'row-only_akt';
            } else {
                badgeClass = 'only_reestr';
                badgeText = '🔴 Только в реестре';
                rowClass = 'row-only_reestr';
            }
            
            return `<tr class="${rowClass}" style="animation-delay: ${idx * 0.03}s">
                <td>${r.number}</td>
                <td style="text-align:center; font-size:1.2em;">${r.inAkt ? '✅' : '❌'}</td>
                <td style="text-align:center; font-size:1.2em;">${r.inReestr ? '✅' : '❌'}</td>
                <td><span class="badge ${badgeClass}">${badgeText}</span></td>
            </tr>`;
        }).join('');
    }
    document.getElementById('resultsSection').classList.remove('hidden');
}

function exportCSV() {
    const diff = results.filter(r => !(r.inAkt && r.inReestr));
    const csv = [
        'Номер;В акте;В реестре;Статус',
        ...diff.map(r => `${r.number};${r.inAkt?'Да':'Нет'};${r.inReestr?'Да':'Нет'};${r.inAkt?'Только акт':'Только реестр'}`)
    ].join('\n');
    const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'raschozhdeniya_akt_yandex.csv';
    a.click();
}

function copyDiffNumbers() {
    const diff = results.filter(r => !(r.inAkt && r.inReestr));
    if (diff.length === 0) {
        alert('Нет расхождений для копирования');
        return;
    }
    const numbers = diff.map(r => r.number).join('\n');
    navigator.clipboard.writeText(numbers).then(() => {
        const btn = event.target;
        const originalText = btn.textContent;
        btn.textContent = '✅ Скопировано!';
        setTimeout(() => { btn.textContent = originalText; }, 2000);
    }).catch(err => {
        const textarea = document.createElement('textarea');
        textarea.value = numbers;
        textarea.style.position = 'fixed';
        textarea.style.opacity = '0';
        document.body.appendChild(textarea);
        textarea.select();
        try {
            document.execCommand('copy');
            alert('✅ Скопировано ' + diff.length + ' номеров!');
        } catch (e) {
            alert('Не удалось скопировать');
        }
        document.body.removeChild(textarea);
    });
}