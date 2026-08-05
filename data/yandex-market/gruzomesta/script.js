let lkData = {};
let reestrData = {};
let results = [];

function parseData(text) {
    const lines = text.split('\n');
    const data = {};
    for (const line of lines) {
        const trimmed = line.trim();
        if (!trimmed) continue;
        const parts = trimmed.split(/[\t;]+|\s{2,}/);
        if (parts.length >= 2) {
            const number = parts[0].trim();
            const gruzomesta = parseFloat(parts[1].replace(',', '.'));
            if (number && !isNaN(gruzomesta)) {
                data[number] = gruzomesta;
            }
        }
    }
    return data;
}

async function startSverka() {
    const lkText = document.getElementById('lkText').value.trim();
    const reestrText = document.getElementById('reestrText').value.trim();
    if (!lkText) {
        alert('Вставьте список из ЛК');
        return;
    }
    if (!reestrText) {
        alert('Вставьте реестр из 1С');
        return;
    }
    lkData = parseData(lkText);
    reestrData = parseData(reestrText);
    if (Object.keys(lkData).length === 0) {
        alert('Не найдено данных в списке из ЛК. Проверьте формат: номер + табуляция + грузоместа');
        return;
    }
    if (Object.keys(reestrData).length === 0) {
        alert('Не найдено данных в реестре из 1С. Проверьте формат: номер + табуляция + грузоместа');
        return;
    }
    document.getElementById('progressSection').classList.remove('hidden');
    document.getElementById('progressSection').classList.add('processing'); // ← анимация
    document.getElementById('resultsSection').classList.add('hidden');
    document.getElementById('startBtn').disabled = true;
    
    const allNumbers = new Set([...Object.keys(lkData), ...Object.keys(reestrData)]);
    const total = allNumbers.size;
    let done = 0;
    results = [];
    
    for (const number of allNumbers) {
        const lkGruz = lkData[number];
        const reestrGruz = reestrData[number];
        let status;
        if (lkGruz === undefined) status = 'only_reestr';
        else if (reestrGruz === undefined) status = 'only_lk';
        else if (lkGruz !== reestrGruz) status = 'mismatch';
        else status = 'match';
        results.push({
            number,
            lkGruz: lkGruz !== undefined ? lkGruz : 0,
            reestrGruz: reestrGruz !== undefined ? reestrGruz : 0,
            diff: Math.abs((lkGruz || 0) - (reestrGruz || 0)),
            status
        });
        done++;
        const pct = Math.round((done / total) * 100);
        document.getElementById('progressFill').style.width = pct + '%';
        document.getElementById('progressPercent').textContent = pct + '%';
        document.getElementById('progressDetails').textContent = 
            `Обработано ${done}/${total} | ЛК: ${Object.keys(lkData).length} | 1С: ${Object.keys(reestrData).length}`;
        if (done % 10 === 0) await new Promise(r => setTimeout(r, 10));
    }
    
    document.getElementById('progressSection').classList.remove('processing'); // ← убираем анимацию
    showResults();
    document.getElementById('startBtn').disabled = false;
}

function showResults() {
    const match = results.filter(r => r.status === 'match').length;
    const mismatch = results.filter(r => r.status === 'mismatch').length;
    const onlyLK = results.filter(r => r.status === 'only_lk').length;
    const onlyReestr = results.filter(r => r.status === 'only_reestr').length;
    const totalDiff = mismatch + onlyLK + onlyReestr;
    
    document.getElementById('summary').innerHTML = `
        <div class="summary-card miss">
            <span class="num">${totalDiff}</span>
            <div class="lbl">Всего расхождений</div>
        </div>
        <div class="summary-card miss">
            <span class="num">${mismatch}</span>
            <div class="lbl">Разное кол-во грузомест</div>
        </div>
        <div class="summary-card miss">
            <span class="num">${onlyLK + onlyReestr}</span>
            <div class="lbl">Нет в одном из источников</div>
        </div>
    `;
    
    const diff = results.filter(r => r.status !== 'match');
    if (diff.length === 0) {
        document.getElementById('resultsBody').innerHTML = `
            <tr><td colspan="5" style="text-align:center; padding:30px; color:#28a745; font-family:-apple-system, sans-serif;">
                ✅ Все грузоместа совпадают! Расхождений нет.
            </td></tr>
        `;
    } else {
        const sorted = [...diff].sort((a, b) => {
            const order = { only_lk: 0, only_reestr: 1, mismatch: 2 };
            return order[a.status] - order[b.status];
        });
        document.getElementById('resultsBody').innerHTML = sorted.map((r, idx) => {
            let badgeClass, badgeText, rowClass;
            
            if (r.status === 'only_lk') {
                badgeClass = 'only_lk';
                badgeText = '🟣 Только в ЛК';
                rowClass = 'row-only_lk';
            } else if (r.status === 'only_reestr') {
                badgeClass = 'only_reestr';
                badgeText = ' Только в 1С';
                rowClass = 'row-only_reestr';
            } else {
                badgeClass = 'mismatch';
                badgeText = '🔴 Разное количество';
                rowClass = 'row-mismatch';
            }
            
            return `<tr class="${rowClass}" style="animation-delay: ${idx * 0.03}s">
                <td>${r.number}</td>
                <td>${r.lkGruz}</td>
                <td>${r.reestrGruz}</td>
                <td>${r.diff > 0 ? r.diff : '-'}</td>
                <td><span class="badge ${badgeClass}">${badgeText}</span></td>
            </tr>`;
        }).join('');
    }
    document.getElementById('resultsSection').classList.remove('hidden');
}

function exportCSV() {
    const diff = results.filter(r => r.status !== 'match');
    const csv = [
        'Номер;Грузомест в ЛК;Грузомест в 1С;Разница;Статус',
        ...diff.map(r => {
            let statusText;
            if (r.status === 'only_lk') statusText = 'Только в ЛК';
            else if (r.status === 'only_reestr') statusText = 'Только в 1С';
            else statusText = 'Разное количество';
            return `${r.number};${r.lkGruz};${r.reestrGruz};${r.diff};${statusText}`;
        })
    ].join('\n');
    const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'raschozhdeniya_gruzomesta_yandex.csv';
    a.click();
}

function copyDiffNumbers() {
    const diff = results.filter(r => r.status !== 'match');
    if (diff.length === 0) {
        alert('Нет расхождений для копирования');
        return;
    }
    const numbers = diff.map(r => r.number).join('\n');
    navigator.clipboard.writeText(numbers).then(() => {
        const btn = event.target;
        const originalText = btn.textContent;
        btn.textContent = '✅ Скопировано!';
        setTimeout(() => {
            btn.textContent = originalText;
        }, 2000);
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