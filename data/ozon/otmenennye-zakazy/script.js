let reestrNumbers = [];
let cancelledNumbers = [];
let matches = [];

/* ---------- Утилиты ---------- */

function normText(v) {
    if (v === null || v === undefined) return '';
    return String(v)
        .replace(/[\s\u00A0\u2028\u2029]+/g, '')
        .replace(/^['"`«»]+|['"`«»]+$/g, '');
}

function normNumber(v) {
    return normText(v).toLowerCase();
}

function lenientKey(num) {
    return num.split('-').map(g => g.replace(/^0+/, '') || '0').join('-');
}

function plural(n, one, few, many) {
    const m10 = n % 10, m100 = n % 100;

    if (m10 === 1 && m100 !== 11) return one;

    if (
        m10 >= 2 &&
        m10 <= 4 &&
        (m100 < 10 || m100 >= 20)
    ) {
        return few;
    }

    return many;
}

/* ---------- Чтение текста ---------- */

// Извлекает номера вида 67893508-0017-1 из вставленного текста.
// Игнорирует вертикальные черты, пробелы, строки-разделители и прочие символы.

function extractOrderNumbers(text) {
    const found = text.match(/\b\d{6,12}-\d{3,5}-\d{1,2}\b/g) || [];

    const seen = new Set();
    const values = [];
    const dups = new Set();

    for (const raw of found) {
        const num = raw.trim();

        if (seen.has(num)) {
            dups.add(num);
        } else {
            seen.add(num);
            values.push(num);
        }
    }

    return {
        values,
        dups: [...dups]
    };
}

/* ---------- Чтение реестра 1С ---------- */

function readReestrText() {
    const text = document.getElementById('reestrText').value;
    const status = document.getElementById('reestrStatus');

    const res = extractOrderNumbers(text);

    reestrNumbers = res.values;

    if (reestrNumbers.length === 0) {
        status.textContent = text.trim()
            ? '❌ Не нашли номера. Проверьте формат.'
            : 'Вставьте список номеров из 1С.';

        status.className = 'status';
        return;
    }

    let msg = `✅ Найдено ${reestrNumbers.length} ${plural(
        reestrNumbers.length,
        'номер',
        'номера',
        'номеров'
    )}`;

    if (res.dups.length) {
        msg += `\n⚠️ Дубли: ${res.dups.slice(0, 10).join(', ')}`;

        if (res.dups.length > 10) {
            msg += '…';
        }
    }

    status.textContent = msg;

    status.className =
        'status' + (res.dups.length ? ' error' : ' success');
}

/* ---------- Чтение отменённых из ЛК ---------- */

function readPostingsText() {
    const text = document.getElementById('postingsText').value;
    const status = document.getElementById('postingsStatus');

    const res = extractOrderNumbers(text);

    cancelledNumbers = res.values;

    if (cancelledNumbers.length === 0) {
        status.textContent = text.trim()
            ? '❌ Не нашли номера отправлений. Проверьте формат.'
            : 'Вставьте список отменённых заказов из ЛК.';

        status.className = 'status';
        return;
    }

    let msg = `✅ Найдено ${cancelledNumbers.length} ${plural(
        cancelledNumbers.length,
        'номер',
        'номера',
        'номеров'
    )}`;

    if (res.dups.length) {
        msg += `\n⚠️ Дубли: ${res.dups.slice(0, 10).join(', ')}`;

        if (res.dups.length > 10) {
            msg += '…';
        }
    }

    status.textContent = msg;

    status.className =
        'status' + (res.dups.length ? ' error' : ' success');
}

/* ---------- Сверка ---------- */

async function startSverka() {
    if (reestrNumbers.length === 0) {
        alert('Загрузите реестр 1С');
        return;
    }

    if (cancelledNumbers.length === 0) {
        alert('Загрузите отменённые заказы из ЛК');
        return;
    }

    document.getElementById('progressSection').classList.remove('hidden');
    document.getElementById('resultsSection').classList.add('hidden');
    document.getElementById('startBtn').disabled = true;

    const strictPost = new Set(cancelledNumbers);
    const lenientPost = new Set(
        cancelledNumbers.map(lenientKey)
    );

    const total = reestrNumbers.length;

    let done = 0;

    matches = [];

    for (const num of reestrNumbers) {
        if (
            strictPost.has(num) ||
            lenientPost.has(lenientKey(num))
        ) {
            matches.push(num);
        }

        done++;

        const pct = Math.round((done / total) * 100);

        document.getElementById('progressFill').style.width =
            pct + '%';

        document.getElementById('progressPercent').textContent =
            pct + '%';

        document.getElementById('progressDetails').textContent =
            `Обработано ${done}/${total} | Реестр 1С: ${total} | Отменённых в ЛК: ${cancelledNumbers.length}`;

        if (done % 200 === 0) {
            await new Promise(r => setTimeout(r, 0));
        }
    }

    showResults();

    document.getElementById('startBtn').disabled = false;
}

/* ---------- Результаты ---------- */

function showResults() {
    document.getElementById('summary').innerHTML = `
        <div class="summary-card ok">
            <span class="num">${reestrNumbers.length}</span>
            <div class="lbl">Номеров в реестре 1С</div>
        </div>

        <div class="summary-card ok">
            <span class="num">${cancelledNumbers.length}</span>
            <div class="lbl">Отменённых в ЛК</div>
        </div>

        <div class="summary-card miss">
            <span class="num">${matches.length}</span>
            <div class="lbl">Отменённых из реестра</div>
        </div>
    `;

    copiedNumbers.clear();

    if (matches.length === 0) {
        document.getElementById('resultsBody').innerHTML = `
            <tr>
                <td
                    colspan="2"
                    style="
                        text-align:center;
                        padding:30px;
                        color:#28a745;
                        font-family:-apple-system, sans-serif;
                    "
                >
                    ✅ Совпадений нет — ни один заказ из реестра не отменён.
                </td>
            </tr>
        `;
    } else {
        document.getElementById('resultsBody').innerHTML =
            matches.map((num, i) => `
                <tr>
                    <td class="col-num">${i + 1}</td>

                    <td>
                        <button
                            class="copy-number ${
                                copiedNumbers.has(num) ? 'copied' : ''
                            }"
                            data-number-index="${i}"
                            type="button"
                            title="Нажмите, чтобы скопировать номер"
                        >${num}</button>
                    </td>
                </tr>
            `).join('');

        document.querySelectorAll('.copy-number').forEach(btn => {
            btn.addEventListener('click', () => {
                const index = Number(
                    btn.dataset.numberIndex
                );

                copySingleNumber(
                    matches[index],
                    btn
                );
            });
        });
    }

    document
        .getElementById('resultsSection')
        .classList.remove('hidden');
}

/* ---------- Копирование отдельных номеров ---------- */

// Номера, которые уже скопировали

const copiedNumbers = new Set();

function copySingleNumber(num, btn) {
    const done = () => {
        copiedNumbers.add(num);

        btn.classList.add('copied');

        btn.title = 'Скопировано';
    };

    if (
        navigator.clipboard &&
        navigator.clipboard.writeText
    ) {
        navigator.clipboard.writeText(num)
            .then(done)
            .catch(() => {
                fallbackCopy(num, done);
            });
    } else {
        fallbackCopy(num, done);
    }
}

function fallbackCopy(text, done) {
    const ta = document.createElement('textarea');

    ta.value = text;

    ta.style.position = 'fixed';
    ta.style.opacity = '0';

    document.body.appendChild(ta);

    ta.select();

    try {
        document.execCommand('copy');
    } catch (e) {}

    document.body.removeChild(ta);

    done();
}

/* ---------- Копирование всего списка ---------- */

function copyList(btn) {
    const text = matches.join('\n');

    if (
        navigator.clipboard &&
        navigator.clipboard.writeText
    ) {
        navigator.clipboard.writeText(text)
            .then(() => {
                btn.textContent = '✅ Скопировано';

                setTimeout(() => {
                    btn.textContent = '📋 Копировать список';
                }, 1500);
            })
            .catch(() => {
                fallbackCopy(text, () => {
                    btn.textContent = '✅ Скопировано';

                    setTimeout(() => {
                        btn.textContent =
                            '📋 Копировать список';
                    }, 1500);
                });
            });
    } else {
        fallbackCopy(text, () => {
            btn.textContent = '✅ Скопировано';

            setTimeout(() => {
                btn.textContent =
                    '📋 Копировать список';
            }, 1500);
        });
    }
}

/* ---------- Экспорт CSV ---------- */

function exportCSV() {
    const csv = [
        'Номер',
        ...matches
    ].join('\n');

    const blob = new Blob(
        ['\ufeff' + csv],
        {
            type: 'text/csv;charset=utf-8;'
        }
    );

    const a = document.createElement('a');

    a.href = URL.createObjectURL(blob);

    a.download = 'otmenennye_iz_reestra.csv';

    a.click();
}

/* ---------- Автоматическое чтение при вводе/вставке ---------- */

document
    .getElementById('reestrText')
    .addEventListener('input', readReestrText);

document
    .getElementById('postingsText')
    .addEventListener('input', readPostingsText);