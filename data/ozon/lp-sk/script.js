let lpData = [];
let skPdfBytes = null;
let skPages = {};
let priceData = {};
let sortedData = [];

// Обработчики загрузки файлов
document.getElementById('lpFileInput').addEventListener('change', e => {
    if (e.target.files[0]) handleLP(e.target.files[0]);
});

document.getElementById('skFileInput').addEventListener('change', e => {
    if (e.target.files[0]) handleSK(e.target.files[0]);
});

document.getElementById('priceFileInput').addEventListener('change', e => {
    if (e.target.files[0]) handlePrice(e.target.files[0]);
});

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
        if (lpData.length > 0) {
            console.log('Пример:', lpData[0]);
        }
        
        updateStatus('lpStatus', '✅', `Загружено: ${lpData.length} товаров`);
        checkReady();
    } catch (err) {
        console.error(' Ошибка ЛП:', err);
        updateStatus('lpStatus', '❌', 'Ошибка: ' + err.message);
    }
}

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
        updateStatus('skStatus', '❌', 'Ошибка: ' + err.message);
    }
}

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
                    // Убираем ВСЕ пробелы из артикула
                    const article = String(row[0]).replace(/\s/g, '').trim();
                    
                    // Проверяем что артикул ровно 6 цифр
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

function checkReady() {
    const ready = lpData.length > 0 && skPdfBytes !== null && Object.keys(priceData).length > 0;
    document.getElementById('generateBtn').disabled = !ready;
    
    console.log('📊 Статус загрузки:');
    console.log('  - ЛП:', lpData.length > 0 ? '✅' : '❌', `(${lpData.length} товаров)`);
    console.log('  - ШК:', skPdfBytes !== null ? '✅' : '', `(${skPdfBytes ? skPdfBytes.length : 0} байт)`);
    console.log('  - Прайс:', Object.keys(priceData).length > 0 ? '✅' : '❌', `(${Object.keys(priceData).length} товаров)`);
    console.log('  - Готовность:', ready ? '✅ МОЖНО ФОРМИРОВАТЬ' : '❌');
}

function updateStatus(elementId, icon, text) {
    const el = document.getElementById(elementId);
    el.innerHTML = `<span class="status-icon">${icon}</span><span>${text}</span>`;
}

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
                    console.error('❌ sessionStorage тоже пустой!');
                }
            } catch (err) {
                console.error('❌ Ошибка восстановления:', err);
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
        
        // ПРОСТАЯ СОРТИРОВКА по кодам символов Unicode
        sortedData.sort((a, b) => {
            const nameA = a.name.trim().toLowerCase();
            const nameB = b.name.trim().toLowerCase();
            
            if (nameA < nameB) return -1;
            if (nameA > nameB) return 1;
            return 0;
        });
        
        console.log('\n📋 ПОСЛЕ СОРТИРОВКИ (первые 3 и последние 3):');
        sortedData.slice(0, 3).forEach((item, i) => {
            console.log(`${i+1}. "${item.name.substring(0, 60)}..."`);
        });
        console.log('...');
        sortedData.slice(-3).forEach((item, i) => {
            console.log(`${sortedData.length-2+i}. "${item.name.substring(0, 60)}..."`);
        });
        
        // Проверяем конкретные товары
        const stiralnaya = sortedData.find(item => item.name.includes('Стиральная машина с инвертором и паром Weissgauff WM 47147'));
        const holodilnik1 = sortedData.find(item => item.name.includes('Холодильник Weissgauff WSBS 600 BeG'));
        const holodilnik2 = sortedData.find(item => item.name.includes('Холодильник Weissgauff WSBS 600 W NoFrost'));
        
        console.log('\n🔍 ПРОВЕРКА ПРОБЛЕМНЫХ ТОВАРОВ:');
        if (stiralnaya) {
            console.log('Стиральная:', stiralnaya.name);
            console.log('  Первый символ:', stiralnaya.name.charAt(0), '(', stiralnaya.name.charCodeAt(0), ')');
            console.log('  Позиция в списке:', sortedData.indexOf(stiralnaya) + 1);
        }
        if (holodilnik1) {
            console.log('Холодильник 1:', holodilnik1.name);
            console.log('  Первый символ:', holodilnik1.name.charAt(0), '(', holodilnik1.name.charCodeAt(0), ')');
            console.log('  Позиция в списке:', sortedData.indexOf(holodilnik1) + 1);
        }
        if (holodilnik2) {
            console.log('Холодильник 2:', holodilnik2.name);
            console.log('  Первый символ:', holodilnik2.name.charAt(0), '(', holodilnik2.name.charCodeAt(0), ')');
            console.log('  Позиция в списке:', sortedData.indexOf(holodilnik2) + 1);
        }
        
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

function switchTab(tab) {
    document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
    
    event.target.classList.add('active');
    document.getElementById(tab === 'lp' ? 'tabLp' : 'tabSk').classList.add('active');
}

function downloadExcel() {
    const excelData = sortedData.map((item, idx) => ({
        '№': idx + 1,
        'Номер отправления': item.number,
        'Название товара': item.name,
        'Артикул': item.article
    }));
    
    const ws = XLSX.utils.json_to_sheet(excelData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Лист подбора');
    const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const excelBlob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(excelBlob, 'noviy_list_podbora.xlsx');
}

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
                console.error(' Ошибка восстановления:', err);
            }
        }
        
        if (!skPdfBytes || skPdfBytes.length === 0) {
            throw new Error('PDF с наклейками не загружен. Пожалуйста, загрузите файл ШК заново.\n\nЕсли проблема повторяется, попробуйте:\n1. Обновить страницу\n2. Загрузить все файлы заново\n3. Открыть через локальный сервер (не file://)');
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
        
        const totalPages = sortedData.length;
        console.log('Страниц в оригинале:', originalPdf.getPageCount());
        console.log('Нужно скопировать:', totalPages, 'наклеек');
        
        // Показываем предупреждение для большого объема
        if (totalPages > 100) {
            alert(`⏱️ Внимание! Формируется большой файл (${totalPages} наклеек).\n\nЭто может занять несколько секунд. Пожалуйста, подождите...`);
        }
        
        const newPdf = await PDFDocument.create();
        let copiedCount = 0;
        let notFoundCount = 0;
        let skippedNumbers = [];
        
        // Показываем прогресс
        const progressDiv = document.createElement('div');
        progressDiv.className = 'pdf-progress';
        progressDiv.innerHTML = `
            <div style="position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%); background: white; padding: 30px; border-radius: 16px; box-shadow: 0 10px 40px rgba(0,0,0,0.3); z-index: 10000; min-width: 400px;">
                <h3 style="margin-bottom: 20px; color: #1d1d1f;">📥 Создание PDF с наклейками...</h3>
                <div style="background: #E6E6FA; height: 20px; border-radius: 10px; overflow: hidden; margin-bottom: 15px;">
                    <div id="pdfProgressFill" style="background: linear-gradient(90deg, #005BFF, #0046CC); height: 100%; width: 0%; transition: width 0.3s;"></div>
                </div>
                <p id="pdfProgressText" style="margin: 0; color: #86868b; font-size: 0.95em;">Обработка 0 из ${totalPages}...</p>
                <p style="margin-top: 15px; font-size: 0.85em; color: #86868b;">⏱️ Пожалуйста, подождите. Не закрывайте страницу.</p>
            </div>
        `;
        document.body.appendChild(progressDiv);
        
        // Копируем страницы порциями для лучшей производительности
        const batchSize = 100; // Обрабатываем по 100 страниц за раз
        
        for (let batchStart = 0; batchStart < sortedData.length; batchStart += batchSize) {
            const batchEnd = Math.min(batchStart + batchSize, sortedData.length);
            const batch = sortedData.slice(batchStart, batchEnd);
            
            for (const item of batch) {
                const pageIndex = skPages[item.number];
                
                if (pageIndex !== undefined) {
                    if (pageIndex < originalPdf.getPageCount()) {
                        const [copiedPage] = await newPdf.copyPages(originalPdf, [pageIndex]);
                        newPdf.addPage(copiedPage);
                        copiedCount++;
                    } else {
                        console.warn(`⚠️ Страница ${pageIndex} не найдена (всего: ${originalPdf.getPageCount()})`);
                    }
                } else {
                    console.warn(`️ Наклейка ${item.number} не найдена в ШК`);
                    skippedNumbers.push(item.number);
                    notFoundCount++;
                }
            }
            
            // Обновляем прогресс
            const progress = Math.round((batchEnd / totalPages) * 100);
            const progressFill = document.getElementById('pdfProgressFill');
            const progressText = document.getElementById('pdfProgressText');
            if (progressFill) progressFill.style.width = progress + '%';
            if (progressText) progressText.textContent = `Обработка ${batchEnd} из ${totalPages}... (${progress}%)`;
            
            // Даём браузеру передохнуть
            await new Promise(resolve => setTimeout(resolve, 10));
        }
        
        console.log('✅ Скопировано:', copiedCount);
        console.log('️ Не найдено:', notFoundCount);
        if (skippedNumbers.length > 0) {
            console.log('Пропущенные номера:', skippedNumbers.slice(0, 10));
        }
        
        if (copiedCount === 0) {
            progressDiv.remove();
            throw new Error('Не удалось скопировать ни одной наклейки.\n\nПроверьте:\n1. Что номера в ЛП совпадают с номерами в ШК\n2. Что файл ШК содержит наклейки\n\nПервые 5 номеров из ЛП:\n' + sortedData.slice(0, 5).map(i => i.number).join('\n'));
        }
        
        // Сохраняем PDF
        const pdfProgressText = document.getElementById('pdfProgressText');
        if (pdfProgressText) pdfProgressText.textContent = 'Сохранение PDF...';
        
        const pdfBytes = await newPdf.save();
        const pdfBlob = new Blob([pdfBytes], { type: 'application/pdf' });
        
        if (pdfProgressText) pdfProgressText.textContent = 'Готово! Скачивание...';
        await new Promise(resolve => setTimeout(resolve, 500));
        
        saveAs(pdfBlob, 'nakleyki_po_poryadku.pdf');
        
        // Удаляем прогресс
        setTimeout(() => {
            progressDiv.remove();
        }, 1000);
        
        console.log('✅ PDF создан и скачан');
        
    } catch (error) {
        console.error('❌ Ошибка создания PDF:', error);
        alert('Ошибка при создании PDF: ' + error.message);
        // Удаляем прогресс если есть
        const progressDiv = document.querySelector('.pdf-progress');
        if (progressDiv) progressDiv.remove();
    }
}

function updateProgress(percent, text) {
    document.getElementById('progressFill').style.width = percent + '%';
    document.getElementById('progressPercent').textContent = percent + '%';
    document.getElementById('progressText').textContent = text;
}

function saveAs(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
}