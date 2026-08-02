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
        updateStatus('lpStatus', '✅', `Загружено: ${lpData.length} товаров`);
        checkReady();
    } catch (err) {
        console.error('❌ Ошибка ЛП:', err);
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
        
        console.log('✅ ШК PDF сохранён, размер:', skPdfBytes.length, 'байт');
        
        const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
        skPages = {};
        
        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const content = await page.getTextContent();
            const pageText = content.items.map(it => it.str).join(' ');
            
            const pattern = /\b\d{4,15}-\d{3,5}-\d{1,2}\b/g;
            const matches = pageText.match(pattern);
            
            if (matches && matches.length > 0) {
                skPages[matches[0]] = i - 1;
            }
        }
        
        console.log('✅ Найдено наклеек:', Object.keys(skPages).length);
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
            
            for (let i = 0; i < jsonData.length; i++) {
                const row = jsonData[i];
                if (row[0] && row[1]) {
                    const article = String(row[0]).replace(/\s/g, '').trim();
                    if (/^\d{6}$/.test(article)) {
                        priceData[article] = String(row[1]).trim();
                        processedCount++;
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

function checkReady() {
    const ready = lpData.length > 0 && skPdfBytes !== null && Object.keys(priceData).length > 0;
    document.getElementById('generateBtn').disabled = !ready;
}

function updateStatus(elementId, icon, text) {
    const el = document.getElementById(elementId);
    el.innerHTML = `<span class="status-icon">${icon}</span><span>${text}</span>`;
}

async function generateFiles() {
    try {
        document.getElementById('progressSection').classList.remove('hidden');
        document.getElementById('generateBtn').disabled = true;
        updateProgress(10, 'Сопоставление данных...');
        
        sortedData = lpData.map(item => ({
            number: item.number,
            name: priceData[item.article] || item.name,
            article: item.article
        }));
        
        updateProgress(30, 'Сортировка по алфавиту...');
        sortedData.sort((a, b) => {
            const nameA = a.name.trim().toLowerCase();
            const nameB = b.name.trim().toLowerCase();
            if (nameA < nameB) return -1;
            if (nameA > nameB) return 1;
            return 0;
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
        console.log('\n' + '='.repeat(60));
        console.log('🚀 НАЧАЛО СКАЧИВАНИЯ PDF');
        console.log('='.repeat(60));
        
        // 1. Проверяем данные
        console.log('\n📊 ШАГ 1: Проверка данных');
        console.log('  skPdfBytes размер:', skPdfBytes ? skPdfBytes.length : 'NULL');
        console.log('  sortedData длина:', sortedData.length);
        console.log('  skPages количество:', Object.keys(skPages).length);
        
        if (!skPdfBytes || skPdfBytes.length === 0) {
            throw new Error('PDF с наклейками не загружен');
        }
        
        // 2. Загружаем оригинальный PDF
        console.log('\n📊 ШАГ 2: Загрузка оригинального PDF');
        const { PDFDocument } = PDFLib;
        const originalPdf = await PDFDocument.load(skPdfBytes, {
            updateMetadata: false,
            ignoreEncryption: true
        });
        console.log('  Страниц в оригинале:', originalPdf.getPageCount());
        
        // 3. Создаём новый PDF
        console.log('\n📊 ШАГ 3: Создание нового PDF');
        const newPdf = await PDFDocument.create();
        let copiedCount = 0;
        let notFoundCount = 0;
        let errorCount = 0;
        let lastSuccessIndex = -1;
        let lastErrorIndex = -1;
        let lastErrorNumber = '';
        
        // Показываем прогресс
        const progressDiv = document.createElement('div');
        progressDiv.style.cssText = 'position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%); background: white; padding: 30px; border-radius: 16px; box-shadow: 0 10px 40px rgba(0,0,0,0.3); z-index: 10000; min-width: 400px;';
        progressDiv.innerHTML = `
            <h3 style="margin-bottom: 20px; color: #1d1d1f;">📥 Создание PDF...</h3>
            <div style="background: #E6E6FA; height: 20px; border-radius: 10px; overflow: hidden; margin-bottom: 15px;">
                <div id="pdfProgressFill" style="background: linear-gradient(90deg, #005BFF, #0046CC); height: 100%; width: 0%; transition: width 0.3s;"></div>
            </div>
            <p id="pdfProgressText" style="margin: 0; color: #86868b; font-size: 0.95em;">Обработка 0 из ${sortedData.length}...</p>
            <p style="margin-top: 15px; font-size: 0.85em; color: #86868b;">⏱️ Пожалуйста, подождите. Не закрывайте страницу.</p>
        `;
        document.body.appendChild(progressDiv);
        
        // 4. Копируем страницы
        console.log('\n ШАГ 4: Копирование страниц');
        const batchSize = 50;
        
        for (let batchStart = 0; batchStart < sortedData.length; batchStart += batchSize) {
            const batchEnd = Math.min(batchStart + batchSize, sortedData.length);
            
            for (let i = batchStart; i < batchEnd; i++) {
                const item = sortedData[i];
                const pageIndex = skPages[item.number];
                
                try {
                    if (pageIndex !== undefined && pageIndex < originalPdf.getPageCount()) {
                        const [copiedPage] = await newPdf.copyPages(originalPdf, [pageIndex]);
                        newPdf.addPage(copiedPage);
                        copiedCount++;
                        lastSuccessIndex = i;
                    } else {
                        notFoundCount++;
                        if (notFoundCount <= 3) {
                            console.log(`  ⚠️ [${i+1}] Не найдена: ${item.number} (pageIndex: ${pageIndex})`);
                        }
                    }
                } catch (err) {
                    errorCount++;
                    lastErrorIndex = i;
                    lastErrorNumber = item.number;
                    if (errorCount <= 3) {
                        console.log(`  ❌ [${i+1}] Ошибка: ${item.number} - ${err.message}`);
                    }
                }
            }
            
            // Обновляем прогресс
            const progress = Math.round((batchEnd / sortedData.length) * 100);
            const progressFill = document.getElementById('pdfProgressFill');
            const progressText = document.getElementById('pdfProgressText');
            if (progressFill) progressFill.style.width = progress + '%';
            if (progressText) progressText.textContent = `Обработка ${batchEnd} из ${sortedData.length}... (${progress}%)`;
            
            // Логируем каждые 100
            if (batchEnd % 100 === 0 || batchEnd === sortedData.length) {
                console.log(`  📍 Прогресс: ${batchEnd}/${sortedData.length} (${progress}%) | Скопировано: ${copiedCount} | Ошибок: ${errorCount}`);
            }
            
            await new Promise(resolve => setTimeout(resolve, 5));
        }
        
        // 5. Итоги копирования
        console.log('\n ШАГ 5: Итоги копирования');
        console.log('  ✅ Скопировано:', copiedCount);
        console.log('  ⚠️ Не найдено:', notFoundCount);
        console.log('  ❌ Ошибок:', errorCount);
        console.log('  📍 Последний успешный индекс:', lastSuccessIndex + 1, 'из', sortedData.length);
        if (lastErrorIndex >= 0) {
            console.log('  ❌ Последняя ошибка на индексе:', lastErrorIndex + 1, 'номер:', lastErrorNumber);
        }
        console.log('  📄 Страниц в новом PDF:', newPdf.getPageCount());
        
        if (copiedCount === 0) {
            progressDiv.remove();
            throw new Error('Не удалось скопировать ни одной наклейки');
        }
        
        // 6. Сохраняем PDF
        console.log('\n📊 ШАГ 6: Сохранение PDF');
        const pdfProgressText = document.getElementById('pdfProgressText');
        if (pdfProgressText) pdfProgressText.textContent = 'Сохранение PDF...';
        
        const pdfBytes = await newPdf.save();
        console.log('  Размер PDF:', (pdfBytes.length / 1024 / 1024).toFixed(2), 'MB');
        console.log('  Длина массива байт:', pdfBytes.length);
        
        if (pdfProgressText) pdfProgressText.textContent = 'Готово! Скачивание...';
        
        // 7. Скачиваем
        console.log('\n📊 ШАГ 7: Скачивание файла');
        const pdfBlob = new Blob([pdfBytes], { type: 'application/pdf' });
        console.log('  Размер Blob:', (pdfBlob.size / 1024 / 1024).toFixed(2), 'MB');
        
        saveAs(pdfBlob, 'nakleyki_po_poryadku.pdf');
        
        setTimeout(() => {
            progressDiv.remove();
        }, 1000);
        
        console.log('\n' + '='.repeat(60));
        console.log('✅ PDF СОЗДАН И СКАЧАН');
        console.log('='.repeat(60));
        console.log('\n📋 ИТОГОВАЯ СВОДКА:');
        console.log('  Всего должно быть:', sortedData.length);
        console.log('  Реально скопировано:', copiedCount);
        console.log('  Разница:', sortedData.length - copiedCount);
        console.log('  Размер файла:', (pdfBytes.length / 1024 / 1024).toFixed(2), 'MB');
        
        if (copiedCount !== sortedData.length) {
            console.warn('\n⚠️ ВНИМАНИЕ! Количество не совпадает!');
            console.warn('  Ожидалось:', sortedData.length);
            console.warn('  Получено:', copiedCount);
            console.warn('  Потеряно:', sortedData.length - copiedCount, 'наклеек');
        }
        
    } catch (error) {
        console.error('\n❌ КРИТИЧЕСКАЯ ОШИБКА:', error);
        alert('Ошибка при создании PDF: ' + error.message);
        const progressDiv = document.querySelector('div[style*="position: fixed"]');
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
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}