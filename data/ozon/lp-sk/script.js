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
                    }
                }
            }

            console.log('✅ Прайс загружен:', processedCount, 'товаров');
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
}

// ============================================================
// ОБНОВЛЕНИЕ СТАТУСА
// ============================================================
function updateStatus(elementId, icon, text) {
    const el = document.getElementById(elementId);
    el.innerHTML = `<span class="status-icon">${icon}</span><span>${text}</span>`;
}

// ============================================================
// ОСНОВНАЯ ГЕНЕРАЦИЯ (С ДОБАВЛЕННЫМ ПРЕДУПРЕЖДЕНИЕМ)
// ============================================================
async function generateFiles() {
    try {
        console.log('\n=== 🚀 НАЧАЛО ГЕНЕРАЦИИ ===');

        // Восстановление PDF, если упала память
        if (!skPdfBytes || skPdfBytes.length === 0) {
            try {
                const base64Pdf = sessionStorage.getItem('skPdfBase64');
                if (base64Pdf) {
                    const binaryString = atob(base64Pdf);
                    skPdfBytes = new Uint8Array(binaryString.length);
                    for (let i = 0; i < binaryString.length; i++) {
                        skPdfBytes[i] = binaryString.charCodeAt(i);
                    }
                }
            } catch (err) {
                console.error('Ошибка восстановления:', err);
            }
        }

        document.getElementById('progressSection').classList.remove('hidden');
        document.getElementById('generateBtn').disabled = true;
        updateProgress(5, 'Сопоставление данных...');

        sortedData = lpData.map(item => ({
            number: item.number,
            name: priceData[item.article] || item.name,
            article: item.article
        }));

        updateProgress(20, 'Сортировка по алфавиту...');
        sortedData.sort((a, b) => {
            const nameA = a.name.trim().toLowerCase();
            const nameB = b.name.trim().toLowerCase();
            return nameA.localeCompare(nameB, 'ru');
        });

        updateProgress(40, 'Подготовка предпросмотра...');
        showPreview();

        // ФИНАЛЬНОЕ ПРЕДУПРЕЖДЕНИЕ
        const totalItems = sortedData.length;
        if (totalItems > 500) {
            updateProgress(60, `⚠️ ВАЖНО: Идет формирование ${totalItems} наклеек. Пожалуйста, подождите! Это займет от 30 секунд до 3 минут в зависимости от объема данных.`);
        } else {
            updateProgress(60, 'Подготовка к скачиванию...');
        }

        // Даем браузеру время отрисовать сообщение
        await new Promise(resolve => setTimeout(resolve, 100));

        updateProgress(80, 'Файлы сформированы. Начинается загрузка...');

        updateProgress(100, 'Готово!');
        setTimeout(() => {
            document.getElementById('progressSection').classList.add('hidden');
            document.getElementById('generateBtn').disabled = false;
        }, 1500);

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
// СКАЧИВАНИЕ EXCEL (Оптимизировано для 10 000+ строк)
// ============================================================
function downloadExcel() {
    try {
        console.log('📊 Генерация Excel (', sortedData.length, ' строк)');
        
        const wb = XLSX.utils.book_new();
        const header = ['№', 'Номер отправления', 'Название товара', 'Артикул'];
        const dataRows = sortedData.map((item, idx) => [
            idx + 1,
            item.number,
            item.name,
            item.article
        ]);
        const allRows = [header, ...dataRows];
        const ws = XLSX.utils.aoa_to_sheet(allRows);
        
        ws['!cols'] = [
            { wch: 6 }, { wch: 20 }, { wch: 50 }, { wch: 12 }
        ];
        
        XLSX.utils.book_append_sheet(wb, ws, 'Лист подбора');
        const excelBuffer = XLSX.write(wb, {
            bookType: 'xlsx',
            type: 'array',
            compression: true
        });

        const excelBlob = new Blob([excelBuffer], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
        saveAs(excelBlob, 'noviy_list_podbora.xlsx');
        console.log('✅ Excel скачан');

    } catch (error) {
        console.error('❌ Ошибка генерации Excel:', error);
        alert('Ошибка при создании Excel: ' + error.message);
    }
}

// ============================================================
// СКАЧИВАНИЕ PDF (ОПТИМИЗИРОВАНО ДЛЯ 3000+ НАКЛЕЕК)
// BATCH_SIZE = 25 (максимальная стабильность)
// ============================================================
async function downloadPDF() {
    try {
        console.log('\n=== 📥 СКАЧИВАНИЕ PDF ===');
        
        // Восстановление данных
        if (!skPdfBytes || skPdfBytes.length === 0) {
            try {
                const base64Pdf = sessionStorage.getItem('skPdfBase64');
                if (base64Pdf) {
                    const binaryString = atob(base64Pdf);
                    skPdfBytes = new Uint8Array(binaryString.length);
                    for (let i = 0; i < binaryString.length; i++) {
                        skPdfBytes[i] = binaryString.charCodeAt(i);
                    }
                }
            } catch (err) {
                console.error('❌ Ошибка восстановления:', err);
            }
        }

        if (!skPdfBytes || skPdfBytes.length === 0) {
            throw new Error('PDF с наклейками не загружен.');
        }

        if (Object.keys(skPages).length === 0) {
            throw new Error('Страницы с наклейками не найдены.');
        }

        // Подготовка списка страниц для копирования
        const pagesToCopy = [];
        let notFoundCount = 0;
        let skippedNumbers = [];

        for (const item of sortedData) {
            const pageIndex = skPages[item.number];
            if (pageIndex !== undefined) {
                pagesToCopy.push(pageIndex);
            } else {
                console.warn(`⚠️ Наклейка ${item.number} не найдена в ШК`);
                skippedNumbers.push(item.number);
                notFoundCount++;
            }
        }

        if (pagesToCopy.length === 0) {
            throw new Error('Не удалось найти ни одной наклейки.');
        }

        const { PDFDocument } = PDFLib;
        const newPdf = await PDFDocument.create();
        
        let copiedCount = 0;
        const BATCH_SIZE = 25; // Максимальная стабильность для 3000+ страниц

        // Основной цикл с пакетной обработкой и сбросом памяти
        for (let i = 0; i < pagesToCopy.length; i += BATCH_SIZE) {
            const batch = pagesToCopy.slice(i, i + BATCH_SIZE);
            
            // Загружаем оригинал заново для каждого пакета (очистка памяти браузера)
            const originalPdf = await PDFDocument.load(skPdfBytes, {
                updateMetadata: false,
                ignoreEncryption: true
            });

            const copiedPages = await newPdf.copyPages(originalPdf, batch);
            for (const page of copiedPages) {
                newPdf.addPage(page);
            }
            
            copiedCount += batch.length;
            
            // Плавное обновление прогресс-бара с понятным текстом
            const percent = Math.round((copiedCount / pagesToCopy.length) * 90) + 10; // От 10% до 100%
            document.getElementById('progressFill').style.width = percent + '%';
            document.getElementById('progressPercent').textContent = percent + '%';
            document.getElementById('progressText').textContent = `Формирование PDF: ${copiedCount} / ${pagesToCopy.length} наклеек (обработано ${Math.round((copiedCount / pagesToCopy.length) * 100)}%)`;

            // Даем браузеру время на сборку мусора перед следующим пакетом
            await new Promise(resolve => setTimeout(resolve, 50));
        }

        document.getElementById('progressText').textContent = 'Сохранение готового файла...';
        document.getElementById('progressPercent').textContent = '99%';

        const pdfBytes = await newPdf.save({ useObjectStreams: false });
        const pdfBlob = new Blob([pdfBytes], { type: 'application/pdf' });
        saveAs(pdfBlob, 'nakleyki_po_poryadku.pdf');

        document.getElementById('progressText').textContent = 'Файл скачан!';
        document.getElementById('progressPercent').textContent = '100%';
        document.getElementById('progressFill').style.width = '100%';

        console.log('✅ PDF создан и скачан (', copiedCount, ' наклеек)');

    } catch (error) {
        console.error('❌ Ошибка создания PDF:', error);
        alert('Ошибка при создании PDF: ' + error.message + '\n\nПопробуйте перезагрузить страницу и загрузить файлы заново.');
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
// СОХРАНЕНИЕ ФАЙЛА
// ============================================================
function saveAs(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
}