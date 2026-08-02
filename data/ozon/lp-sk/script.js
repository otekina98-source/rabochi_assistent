// ============================================================
// ГЛОБАЛЬНЫЕ ПЕРЕМЕННЫЕ
// ============================================================
let lpData = [];
let skPdfBytes = null;
let skPages = {};
let priceData = {};
let sortedData = [];

// ============================================================
// ОБРАБОТЧИКИ ЗАГРУЗКИ ФАЙЛОВ
// ============================================================
document.getElementById('lpFileInput').addEventListener('change', e => {
    if (e.target.files[0]) handleLP(e.target.files[0]);
});

document.getElementById('skFileInput').addEventListener('change', e => {
    if (e.target.files[0]) handleSK(e.target.files[0]);
});

document.getElementById('priceFileInput').addEventListener('change', e => {
    if (e.target.files[0]) handlePrice(e.target.files[0]);
});

// ============================================================
// ЗАГРУЗКА ЛИСТА ПОДБОРА (PDF)
// ============================================================
async function handleLP(file) {
    document.getElementById('lpFileName').textContent = file.name;
    updateStatus('lpStatus', '', 'Чтение...');

    try {
        const buffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
        let text = '';

        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const content = await page.getTextContent();
            text += content.items.map(it => it.str).join(' ') + '\n';
        }

        lpData = [];
        const pattern = /(\d{4,15}-\d{3,5}-\d{1,2})\s+(.+?)\s+(\d{6})\b/g;
        let match;

        while ((match = pattern.exec(text)) !== null) {
            lpData.push({
                number: match[1],
                name: match[2].trim(),
                article: match[3]
            });
        }

        console.log('✅ ЛП загружено:', lpData.length, 'товаров');
        if (lpData.length > 0) console.log('Пример:', lpData[0]);

        updateStatus('lpStatus', '✅', `Загружено: ${lpData.length} товаров`);
        checkReady();
    } catch (err) {
        console.error('❌ Ошибка ЛП:', err);
        updateStatus('lpStatus', '❌', 'Ошибка: ' + err.message);
    }
}

// ============================================================
// ЗАГРУЗКА ШК (PDF)
// ============================================================
async function handleSK(file) {
    document.getElementById('skFileName').textContent = file.name;
    updateStatus('skStatus', '', 'Чтение...');

    try {
        const buffer = await file.arrayBuffer();

        const header = new Uint8Array(buffer.slice(0, 5));
        const headerStr = String.fromCharCode(...header);
        if (headerStr !== '%PDF-') {
            throw new Error('Файл не является PDF (заголовок: ' + headerStr + ')');
        }

        skPdfBytes = new Uint8Array(buffer.byteLength);
        skPdfBytes.set(new Uint8Array(buffer));

        console.log('✅ ШК PDF сохранён (копия), размер:', skPdfBytes.length, 'байт');

        try {
            const base64Pdf = btoa(String.fromCharCode(...skPdfBytes));
            sessionStorage.setItem('skPdfBase64', base64Pdf);
            console.log('✅ ШК PDF сохранён в sessionStorage');
        } catch (storageErr) {
            console.warn('⚠️ Не удалось сохранить в sessionStorage:', storageErr.message);
        }

        const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
        skPages = {};
        console.log('Всего страниц в ШК:', pdf.numPages);

        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const content = await page.getTextContent();
            const pageText = content.items.map(it => it.str).join(' ');
            const pattern = /\b\d{4,15}-\d{3,5}-\d{1,2}\b/g;
            const matches = pageText.match(pattern);

            if (matches && matches.length > 0) {
                const number = matches[0];
                skPages[number] = i - 1;
            }
        }

        console.log('✅ Всего найдено наклеек:', Object.keys(skPages).length);
        console.log('Примеры:', Object.keys(skPages).slice(0, 5));

        updateStatus('skStatus', '✅', `Загружено: ${Object.keys(skPages).length} наклеек`);
        checkReady();
    } catch (err) {
        console.error('❌ Ошибка ШК:', err);
        updateStatus('skStatus', '', 'Ошибка: ' + err.message);
    }
}

// ============================================================
// ЗАГРУЗКА ПРАЙСА (EXCEL)
// ============================================================
function handlePrice(file) {
    document.getElementById('priceFileName').textContent = file.name;
    updateStatus('priceStatus', '', 'Чтение...');

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });

            priceData = {};
            let processedCount = 0;
            let skippedCount = 0;

            for (let i = 0; i < jsonData.length; i++) {
                const row = jsonData[i];
                if (row[0] && row[1]) {
                    const article = String(row[0]).replace(/\s/g, '').trim();
                    if (/^\d{6}$/.test(article)) {
                        priceData[article] = String(row[1]).trim();
                        processedCount++;
                    } else {
                        skippedCount++;
                        if (skippedCount <= 5) {
                            console.log(`⚠️ Пропущен артикул: "${row[0]}" → "${article}"`);
                        }
                    }
                }
            }

            console.log('✅ Прайс загружен:', processedCount, 'товаров');
            console.log('⚠️ Пропущено:', skippedCount, 'строк');

            updateStatus('priceStatus', '✅', `Загружено: ${processedCount} товаров`);
            checkReady();
        } catch (err) {
            console.error('❌ Ошибка Прайса:', err);
            updateStatus('priceStatus', '❌', 'Ошибка: ' + err.message);
        }
    };
    reader.readAsArrayBuffer(file);
}

// ============================================================
// ПРОВЕРКА ГОТОВНОСТИ
// ============================================================
function checkReady() {
    const ready = lpData.length > 0 && skPdfBytes !== null && Object.keys(priceData).length > 0;
    document.getElementById('generateBtn').disabled = !ready;

    console.log('📊 Статус загрузки:');
    console.log('  - ЛП:', lpData.length > 0 ? '✅' : '', `(${lpData.length} товаров)`);
    console.log('  - ШК:', skPdfBytes !== null ? '✅' : '❌', `(${skPdfBytes ? skPdfBytes.length : 0} байт)`);
    console.log('  - Прайс:', Object.keys(priceData).length > 0 ? '✅' : '❌', `(${Object.keys(priceData).length} товаров)`);
    console.log('  - Готовность:', ready ? '✅ МОЖНО ФОРМИРОВАТЬ' : '❌');
}

// ============================================================
// ОБНОВЛЕНИЕ СТАТУСА
// ============================================================
function updateStatus(elementId, icon, text) {
    const el = document.getElementById(elementId);
    el.innerHTML = `<span class="status-icon">${icon}</span><span>${text}</span>`;
}

// ============================================================
// ОСНОВНАЯ ГЕНЕРАЦИЯ
// ============================================================
async function generateFiles() {
    try {
        console.log('\n=== 🚀 НАЧАЛО ГЕНЕРАЦИИ ===');
        console.log('skPdfBytes:', skPdfBytes ? skPdfBytes.length : 'null', 'байт');
        console.log('skPages:', Object.keys(skPages).length, 'наклеек');

        if (!skPdfBytes || skPdfBytes.length === 0) {
            console.warn('⚠️ skPdfBytes пустой, пробуем восстановить из sessionStorage...');
            try {
                const base64Pdf = sessionStorage.getItem('skPdfBase64');
                if (base64Pdf) {
                    const binaryString = atob(base64Pdf);
                    skPdfBytes = new Uint8Array(binaryString.length);
                    for (let i = 0; i < binaryString.length; i++) {
                        skPdfBytes[i] = binaryString.charCodeAt(i);
                    }
                    console.log('✅ Восстановлено из sessionStorage:', skPdfBytes.length, 'байт');
                } else {
                    console.error(' sessionStorage тоже пустой!');
                }
            } catch (err) {
                console.error(' Ошибка восстановления:', err);
            }
        }

        document.getElementById('progressSection').classList.remove('hidden');
        document.getElementById('generateBtn').disabled = true;
        updateProgress(10, 'Сопоставление данных...');

        sortedData = lpData.map(item => ({
            number: item.number,
            name: priceData[item.article] || item.name,
            article: item.article
        }));

        console.log('\n📋 ПЕРЕД СОРТИРОВКОЙ (первые 3):');
        sortedData.slice(0, 3).forEach((item, i) => {
            console.log(`${i+1}. "${item.name.substring(0, 50)}..."`);
        });

        updateProgress(30, 'Сортировка по алфавиту...');

        sortedData.sort((a, b) => {
            const nameA = a.name.trim().toLowerCase();
            const nameB = b.name.trim().toLowerCase();
            return nameA.localeCompare(nameB, 'ru');
        });

        console.log('\n ПОСЛЕ СОРТИРОВКИ (первые 3 и последние 3):');
        sortedData.slice(0, 3).forEach((item, i) => {
            console.log(`${i+1}. "${item.name.substring(0, 60)}..."`);
        });
        console.log('...');
        sortedData.slice(-3).forEach((item, i) => {
            console.log(`${sortedData.length-2+i}. "${item.name.substring(0, 60)}..."`);
        });

        updateProgress(50, 'Подготовка предпросмотра...');
        showPreview();

        updateProgress(100, 'Готово к скачиванию!');
        setTimeout(() => {
            document.getElementById('progressSection').classList.add('hidden');
            document.getElementById('generateBtn').disabled = false;
        }, 500);

    } catch (error) {
        console.error('❌ Ошибка генерации:', error);
        alert('Ошибка: ' + error.message);
        document.getElementById('progressSection').classList.add('hidden');
        document.getElementById('generateBtn').disabled = false;
    }
}

// ============================================================
// ОТОБРАЖЕНИЕ ПРЕДПРОСМОТРА
// ============================================================
function showPreview() {
    document.getElementById('resultSection').classList.remove('hidden');
    document.getElementById('lpCount').textContent = sortedData.length;
    document.getElementById('skCount').textContent = sortedData.length;

    const lpBody = document.getElementById('lpTableBody');
    lpBody.innerHTML = sortedData.map((item, idx) => `
        <tr>
            <td>${idx + 1}</td>
            <td>${item.number}</td>
            <td>${item.name}</td>
            <td>${item.article}</td>
        </tr>
    `).join('');

    const stickersList = document.getElementById('stickersList');
    stickersList.innerHTML = sortedData.map((item, idx) => `
        <div class="sticker-item">
            <span class="sticker-idx">${idx + 1}</span>
            <span class="sticker-num">${item.number}</span>
        </div>
    `).join('');

    document.getElementById('resultSection').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ============================================================
// ПЕРЕКЛЮЧЕНИЕ ВКЛАДОК
// ============================================================
function switchTab(tab) {
    document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));

    event.target.classList.add('active');
    document.getElementById(tab === 'lp' ? 'tabLp' : 'tabSk').classList.add('active');
}

// ============================================================
// СКАЧИВАНИЕ EXCEL (оптимизировано для 10 000+ строк)
// ============================================================
function downloadExcel() {
    try {
        console.log('📊 Начинаем генерацию Excel (', sortedData.length, ' строк)');

        // 1. Создаём рабочую книгу
        const wb = XLSX.utils.book_new();

        // 2. Генерируем заголовки
        const header = ['№', 'Номер отправления', 'Название товара', 'Артикул'];

        // 3. Создаём массив данных (без заголовка в json_to_sheet)
        const dataRows = sortedData.map((item, idx) => [
            idx + 1,
            item.number,
            item.name,
            item.article
        ]);

        // 4. Объединяем заголовок + данные
        const allRows = [header, ...dataRows];

        // 5. Преобразуем в лист (через aoa_to_sheet — оптимизировано для больших массивов)
        const ws = XLSX.utils.aoa_to_sheet(allRows);

        // 6. Добавляем лист в книгу
        XLSX.utils.book_append_sheet(wb, ws, 'Лист подбора');

        // 7. Оптимизация ширины колонок (опционально)
        const colWidths = [
            { wch: 6 },   // №
            { wch: 20 },  // Номер отправления
            { wch: 50 },  // Название товара
            { wch: 12 }   // Артикул
        ];
        ws['!cols'] = colWidths;

        // 8. Сохраняем через write (синхронно, но для 10k+ строк работает быстро)
        console.time('Excel generation');
        const excelBuffer = XLSX.write(wb, {
            bookType: 'xlsx',
            type: 'array',
            compression: true // Включаем сжатие для больших файлов
        });
        console.timeEnd('Excel generation');

        // 9. Создаём Blob и скачиваем
        const excelBlob = new Blob([excelBuffer], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
        saveAs(excelBlob, 'noviy_list_podbora.xlsx');

        console.log('✅ Excel успешно скачан (', sortedData.length, ' строк)');

    } catch (error) {
        console.error('❌ Ошибка генерации Excel:', error);
        alert('Ошибка при создании Excel: ' + error.message + '\n\nПопробуйте использовать меньший файл или обновить страницу.');
    }
}

// ============================================================
// СКАЧИВАНИЕ PDF (оптимизировано для 1000+ наклеек)
// ============================================================
async function downloadPDF() {
    try {
        console.log('\n=== 📥 СКАЧИВАНИЕ PDF ===');
        console.log('skPdfBytes:', skPdfBytes ? skPdfBytes.length : 'null');

        if (!skPdfBytes || skPdfBytes.length === 0) {
            try {
                const base64Pdf = sessionStorage.getItem('skPdfBase64');
                if (base64Pdf) {
                    const binaryString = atob(base64Pdf);
                    skPdfBytes = new Uint8Array(binaryString.length);
                    for (let i = 0; i < binaryString.length; i++) {
                        skPdfBytes[i] = binaryString.charCodeAt(i);
                    }
                    console.log('✅ Восстановлено из sessionStorage:', skPdfBytes.length, 'байт');
                }
            } catch (err) {
                console.error('❌ Ошибка восстановления:', err);
            }
        }

        if (!skPdfBytes || skPdfBytes.length === 0) {
            throw new Error('PDF с наклейками не загружен. Пожалуйста, загрузите файл ШК заново.');
        }

        if (Object.keys(skPages).length === 0) {
            throw new Error('Страницы с наклейками не найдены. Проверьте файл ШК.');
        }

        const { PDFDocument } = PDFLib;

        console.log('Загрузка оригинального PDF...');
        const originalPdf = await PDFDocument.load(skPdfBytes, {
            updateMetadata: false,
            ignoreEncryption: true
        });

        console.log('Страниц в оригинале:', originalPdf.getPageCount());
        console.log('Нужно скопировать:', sortedData.length, 'наклеек');

        const newPdf = await PDFDocument.create();
        let copiedCount = 0;
        let notFoundCount = 0;
        let skippedNumbers = [];

        // ✅ ПАКЕТНОЕ КОПИРОВАНИЕ (по 100 страниц за раз) для избежания переполнения памяти
        const BATCH_SIZE = 100;
        const pagesToCopy = [];

        for (const item of sortedData) {
            const pageIndex = skPages[item.number];
            if (pageIndex !== undefined && pageIndex < originalPdf.getPageCount()) {
                pagesToCopy.push(pageIndex);
            } else {
                console.warn(`⚠️ Наклейка ${item.number} не найдена в ШК`);
                skippedNumbers.push(item.number);
                notFoundCount++;
            }
        }

        console.log(`✅ Найдено страниц для копирования: ${pagesToCopy.length}`);

        // Копируем пакетами
        for (let i = 0; i < pagesToCopy.length; i += BATCH_SIZE) {
            const batch = pagesToCopy.slice(i, i + BATCH_SIZE);
            const copiedPages = await newPdf.copyPages(originalPdf, batch);
            for (const page of copiedPages) {
                newPdf.addPage(page);
            }
            copiedCount += batch.length;
            console.log(`🔄 Скопировано ${copiedCount} / ${pagesToCopy.length} страниц`);

            // Даём браузеру время на сборку мусора
            await new Promise(resolve => setTimeout(resolve, 50));
        }

        console.log('✅ Скопировано:', copiedCount);
        console.log('⚠️ Не найдено:', notFoundCount);
        if (skippedNumbers.length > 0) {
            console.log('Пропущенные номера (первые 10):', skippedNumbers.slice(0, 10));
        }

        if (copiedCount === 0) {
            throw new Error('Не удалось скопировать ни одной наклейки.\n\nПроверьте:\n1. Что номера в ЛП совпадают с номерами в ШК\n2. Что файл ШК содержит наклейки');
        }

        // ✅ Сохраняем с оптимизацией (без сжатия для скорости)
        const pdfBytes = await newPdf.save({ useObjectStreams: false });

        const pdfBlob = new Blob([pdfBytes], { type: 'application/pdf' });
        saveAs(pdfBlob, 'nakleyki_po_poryadku.pdf');

        console.log('✅ PDF создан и скачан');

    } catch (error) {
        console.error('❌ Ошибка создания PDF:', error);
        alert('Ошибка при создании PDF: ' + error.message);
    }
}

// ============================================================
// ОБНОВЛЕНИЕ ПРОГРЕССА
// ============================================================
function updateProgress(percent, text) {
    document.getElementById('progressFill').style.width = percent + '%';
    document.getElementById('progressPercent').textContent = percent + '%';
    document.getElementById('progressText').textContent = text;
}

// ============================================================
// УНИВЕРСАЛЬНОЕ СОХРАНЕНИЕ ФАЙЛА
// ============================================================
function saveAs(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
}