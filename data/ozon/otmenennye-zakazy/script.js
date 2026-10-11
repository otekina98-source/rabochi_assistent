let reestrNumbers = [];    // номера из столбика Телефон1 (реестр 1С)
let cancelledNumbers = []; // номера из столбика Номер отправления (отменённые из ЛК)
let matches = [];


/* ---------- Утилиты ---------- */

function normText(v) {
    // Приводим ячейку к строке и убираем всё лишнее: пробелы, неразрывные
    // пробелы, кавычки и апострофы, которые Excel/1С иногда добавляют.
    if (v === null || v === undefined) return '';

    return String(v)
        .replace(/[\s\u00A0\u2028\u2029]+/g, '')
        .replace(/^['"`«»]+|['"`«»]+$/g, '');
}


function normNumber(v) {
    return normText(v).toLowerCase();
}


// Мягкий ключ: убираем ведущие нули в каждой группе, чтобы
// «0197595052-0004-1» и «197595052-4-1» считались одним номером.
function lenientKey(num) {
    return num
        .split('-')
        .map(g => g.replace(/^0+/, '') || '0')
        .join('-');
}


function plural(n, one, few, many) {

    const m10 = n % 10;
    const m100 = n % 100;

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


function isExcelLike(file) {
    return /\.(xlsx|xls|csv)$/i.test(file.name);
}


function libReady() {
    return typeof XLSX !== 'undefined';
}


// Разворачиваем файл в массив листов,
// каждый — массив строк (через SheetJS).
async function readSheets(file) {

    if (/\.csv$/i.test(file.name)) {

        const text = await file.text();

        const wb = XLSX.read(
            text,
            {
                type: 'string'
            }
        );

        return wb.SheetNames.map(
            name =>
                XLSX.utils.sheet_to_json(
                    wb.Sheets[name],
                    {
                        header: 1,
                        raw: false,
                        defval: null
                    }
                )
        );
    }


    const buffer =
        await file.arrayBuffer();


    const wb =
        XLSX.read(
            buffer,
            {
                type: 'array'
            }
        );


    return wb.SheetNames.map(
        name =>
            XLSX.utils.sheet_to_json(
                wb.Sheets[name],
                {
                    header: 1,
                    raw: false,
                    defval: null
                }
            )
    );
}


// Ищем в листе строку заголовков
// и столбик с нужным названием.
//
// Заголовки могут быть не в первой строке
// (в реестре 1С — в третьей),
// поэтому сканируем первые 100 строк каждого листа.
function findColInSheet(rows, targetHeader) {

    const limit =
        Math.min(rows.length, 100);


    for (let r = 0; r < limit; r++) {

        const row =
            rows[r] || [];


        for (let c = 0; c < row.length; c++) {

            if (
                normText(row[c]).toLowerCase() === targetHeader
            ) {

                return {
                    col: c,
                    headerRow: r
                };

            }

        }

    }


    return null;
}


// Собираем значения столбика со всех листов
// уникальные, в порядке следования.
// Попутно считаем дубли.
//
// Возвращает { values, dups }
// или null, если столбик не найден ни в одном листе.
function collectColumn(sheets, targetHeader) {

    const seen =
        new Map();


    let foundIn = 0;


    for (const rows of sheets) {

        const found =
            findColInSheet(
                rows,
                targetHeader
            );


        if (!found) continue;


        foundIn++;


        for (
            let r = found.headerRow + 1;
            r < rows.length;
            r++
        ) {

            const key =
                normNumber(
                    (rows[r] || [])[found.col]
                );


            if (!key) continue;


            seen.set(
                key,
                (seen.get(key) || 0) + 1
            );

        }

    }


    if (foundIn === 0) {
        return null;
    }


    const values = [];
    const dups = [];


    seen.forEach(
        (count, key) => {

            values.push(key);

            if (count > 1) {
                dups.push(key);
            }

        }
    );


    return {
        values,
        dups
    };
}


/* ---------- Загрузка файлов ---------- */

function setupUpload(
    areaId,
    inputId,
    handler
) {

    const area =
        document.getElementById(areaId);


    const input =
        document.getElementById(inputId);


    area.addEventListener(
        'dragover',
        e => {

            e.preventDefault();

            area.classList.add(
                'dragover'
            );

        }
    );


    area.addEventListener(
        'dragleave',
        () =>
            area.classList.remove(
                'dragover'
            )
    );


    area.addEventListener(
        'drop',
        e => {

            e.preventDefault();

            area.classList.remove(
                'dragover'
            );


            const file =
                e.dataTransfer.files[0];


            if (
                file &&
                isExcelLike(file)
            ) {

                handler(file);

            }

        }
    );


    input.addEventListener(
        'change',
        e => {

            if (e.target.files[0]) {

                handler(
                    e.target.files[0]
                );

            }

        }
    );
}


/* ---------- Загрузка реестра 1С ---------- */

async function handleReestr(file) {

    document
        .getElementById('reestrFileName')
        .textContent =
        file.name;


    const status =
        document.getElementById(
            'reestrStatus'
        );


    status.textContent =
        '⏳ Чтение файла...';


    status.className =
        'status';


    reestrNumbers = [];


    if (!libReady()) {

        status.textContent =
            '❌ Не загрузилась библиотека для чтения Excel. Проверьте интернет и обновите страницу.';

        status.className =
            'status error';

        return;

    }


    try {

        const sheets =
            await readSheets(file);


        const res =
            collectColumn(
                sheets,
                'телефон1'
            );


        if (!res) {

            status.textContent =
                '❌ Не нашли столбик «Телефон1» — проверьте, тот ли файл загружен.';

            status.className =
                'status error';

            return;

        }


        reestrNumbers =
            res.values;


        let msg =
            `✅ Найдено ${res.values.length} ${plural(
                res.values.length,
                'номер',
                'номера',
                'номеров'
            )} (столбик «Телефон1»)`;


        if (res.dups.length) {

            msg +=
                `\n⚠️ Внимание, есть дубли: ${res.dups
                    .slice(0, 10)
                    .join(', ')}${res.dups.length > 10 ? '…' : ''}`;

        }


        status.textContent =
            msg;


        status.className =
            'status' +
            (
                res.dups.length
                    ? ' error'
                    : ' success'
            );


    } catch (err) {

        status.textContent =
            '❌ Ошибка: ' +
            err.message;


        status.className =
            'status error';


        console.error(err);

    }

}


/* ---------- Загрузка отменённых из ЛК ---------- */

async function handlePostings(file) {

    document
        .getElementById('postingsFileName')
        .textContent =
        file.name;


    const status =
        document.getElementById(
            'postingsStatus'
        );


    status.textContent =
        '⏳ Чтение файла...';


    status.className =
        'status';


    cancelledNumbers = [];


    if (!libReady()) {

        status.textContent =
            '❌ Не загрузилась библиотека для чтения Excel. Проверьте интернет и обновите страницу.';

        status.className =
            'status error';

        return;

    }


    try {

        const sheets =
            await readSheets(file);


        // Собираем все номера отправления;
        // если в листе есть столбик «Статус» —
        // параллельно только те,
        // у которых статус отмены.

        const all = [];
        const cancelled = [];

        const seenAll =
            new Set();

        const seenCancelled =
            new Set();


        let hasStatusCol = false;

        let foundAny = false;


        for (const rows of sheets) {

            const nf =
                findColInSheet(
                    rows,
                    'номеротправления'
                );


            if (!nf) continue;


            foundAny = true;


            const sf =
                findColInSheet(
                    rows,
                    'статус'
                );


            if (sf) {
                hasStatusCol = true;
            }


            for (
                let r = nf.headerRow + 1;
                r < rows.length;
                r++
            ) {

                const row =
                    rows[r] || [];


                const key =
                    normNumber(
                        row[nf.col]
                    );


                if (!key) continue;


                if (!seenAll.has(key)) {

                    seenAll.add(key);

                    all.push(key);

                }


                if (
                    sf &&
                    /отмен|cancel/i.test(
                        normText(
                            row[sf.col]
                        )
                    ) &&
                    !seenCancelled.has(key)
                ) {

                    seenCancelled.add(key);

                    cancelled.push(key);

                }

            }

        }


        if (!foundAny) {

            status.textContent =
                '❌ Не нашли столбик «Номер отправления» — проверьте, тот ли файл загружен.';

            status.className =
                'status error';

            return;

        }


        let numbers = all;

        let note = '';


        if (
            hasStatusCol &&
            cancelled.length > 0
        ) {

            numbers =
                cancelled;

        } else if (hasStatusCol) {

            note =
                ' (взяты все строки файла)';

        }


        cancelledNumbers =
            numbers;


        status.textContent =
            `✅ Найдено ${numbers.length} ${plural(
                numbers.length,
                'отменённое отправление',
                'отменённых отправления',
                'отменённых отправлений'
            )} (столбик «Номер отправления»)${note}`;


        status.className =
            'status success';


    } catch (err) {

        status.textContent =
            '❌ Ошибка: ' +
            err.message;


        status.className =
            'status error';


        console.error(err);

    }

}


/* ---------- Сверка ---------- */

async function startSverka() {

    if (
        reestrNumbers.length === 0
    ) {

        alert(
            'Загрузите реестр 1С'
        );

        return;

    }


    if (
        cancelledNumbers.length === 0
    ) {

        alert(
            'Загрузите отменённые заказы из ЛК'
        );

        return;

    }


    document
        .getElementById(
            'progressSection'
        )
        .classList
        .remove('hidden');


    document
        .getElementById(
            'resultsSection'
        )
        .classList
        .add('hidden');


    document
        .getElementById(
            'startBtn'
        )
        .disabled = true;


    const strictPost =
        new Set(
            cancelledNumbers
        );


    const lenientPost =
        new Set(
            cancelledNumbers.map(
                lenientKey
            )
        );


    const total =
        reestrNumbers.length;


    let done = 0;


    matches = [];


    for (
        const num of reestrNumbers
    ) {

        if (
            strictPost.has(num) ||
            lenientPost.has(
                lenientKey(num)
            )
        ) {

            matches.push(num);

        }


        done++;


        const pct =
            Math.round(
                (done / total) * 100
            );


        document
            .getElementById(
                'progressFill'
            )
            .style.width =
            pct + '%';


        document
            .getElementById(
                'progressPercent'
            )
            .textContent =
            pct + '%';


        document
            .getElementById(
                'progressDetails'
            )
            .textContent =
            `Обработано ${done}/${total} | Реестр 1С: ${total} | Отменённых в ЛК: ${cancelledNumbers.length}`;


        if (
            done % 200 === 0
        ) {

            await new Promise(
                r =>
                    setTimeout(
                        r,
                        0
                    )
            );

        }

    }


    showResults();


    document
        .getElementById(
            'startBtn'
        )
        .disabled = false;

}


/* ---------- Результаты ---------- */

function showResults() {

    document
        .getElementById(
            'summary'
        )
        .innerHTML = `

        <div class="summary-card ok">

            <span class="num">
                ${reestrNumbers.length}
            </span>

            <div class="lbl">
                Номеров в реестре 1С
            </div>

        </div>


        <div class="summary-card ok">

            <span class="num">
                ${cancelledNumbers.length}
            </span>

            <div class="lbl">
                Отменённых в ЛК
            </div>

        </div>


        <div class="summary-card miss">

            <span class="num">
                ${matches.length}
            </span>

            <div class="lbl">
                Отменённых из реестра
            </div>

        </div>

    `;


    if (
        matches.length === 0
    ) {

        document
            .getElementById(
                'resultsBody'
            )
            .innerHTML = `

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

                    ✅ Совпадений нет —
                    ни один заказ из реестра не отменён.

                </td>

            </tr>

        `;

    } else {

        document
            .getElementById(
                'resultsBody'
            )
            .innerHTML =

            matches
                .map(
                    (num, i) => `

                <tr>

                    <td class="col-num">

                        ${i + 1}

                    </td>


                    <td>

                        <span
                            class="order-copy"
                            title="Скопировать номер"
                            onclick="copySingleNumber('${num}', this.closest('tr'))"
                        >

                            ${num}

                        </span>

                    </td>

                </tr>

            `
                )
                .join('');

    }


    document
        .getElementById(
            'resultsSection'
        )
        .classList
        .remove('hidden');

}


/* ---------- Копирование ОДНОГО номера ---------- */

function copySingleNumber(
    number,
    row
) {

    const done = () => {

        if (row) {

            row.classList.add(
                'copied-row'
            );

        }

    };


    if (
        navigator.clipboard &&
        navigator.clipboard.writeText
    ) {

        navigator.clipboard
            .writeText(number)
            .then(done)
            .catch(
                () =>
                    fallbackCopy(
                        number,
                        done
                    )
            );

    } else {

        fallbackCopy(
            number,
            done
        );

    }

}


/* ---------- Копировать весь список ---------- */

function copyList(btn) {

    if (
        matches.length === 0
    ) {
        return;
    }


    const text =
        matches.join('\n');


    const done = () => {

        if (btn) {

            const old =
                btn.textContent;


            btn.textContent =
                '✅ Скопировано!';


            setTimeout(
                () =>
                    btn.textContent =
                    old,
                1500
            );

        }

    };


    if (
        navigator.clipboard &&
        navigator.clipboard.writeText
    ) {

        navigator.clipboard
            .writeText(text)
            .then(done)
            .catch(
                () =>
                    fallbackCopy(
                        text,
                        done
                    )
            );

    } else {

        fallbackCopy(
            text,
            done
        );

    }

}


function fallbackCopy(
    text,
    done
) {

    const ta =
        document.createElement(
            'textarea'
        );


    ta.value =
        text;


    ta.style.position =
        'fixed';


    ta.style.opacity =
        '0';


    document.body
        .appendChild(ta);


    ta.select();


    try {

        document.execCommand(
            'copy'
        );

    } catch (e) {}


    document.body
        .removeChild(ta);


    done();

}


/* ---------- CSV ---------- */

function exportCSV() {

    const csv =
        [
            'Номер',
            ...matches
        ]
        .join('\n');


    const blob =
        new Blob(
            [
                '\ufeff' +
                csv
            ],
            {
                type:
                    'text/csv;charset=utf-8;'
            }
        );


    const a =
        document.createElement(
            'a'
        );


    a.href =
        URL.createObjectURL(
            blob
        );


    a.download =
        'otmenennye_iz_reestra.csv';


    a.click();

}


/* ---------- Инициализация ---------- */

setupUpload(
    'reestrUploadArea',
    'reestrFileInput',
    handleReestr
);


setupUpload(
    'postingsUploadArea',
    'postingsFileInput',
    handlePostings
);