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
        
        console.log('📄 ЛП PDF страниц:', pdf.numPages);
        
        lpData = [];
        let skippedRows = 0;
        let skippedExamples = [];
        
        // Читаем все страницы
        for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
            const page = await pdf.getPage(pageNum);
            const content = await page.getTextContent();
            
            // Собираем текст с координатами
            const items = content.items.map(item => ({
                text: item.str,
                x: Math.round(item.transform[4]),
                y: Math.round(item.transform[5])
            }));
            
            // Группируем элементы по строкам (одинаковый Y)
            const rows = {};
            items.forEach(item => {
                const yKey = item.y;
                if (!rows[yKey]) rows[yKey] = [];
                rows[yKey].push(item);
            });
            
            // Сортируем строки сверху вниз (Y убывает)
            const sortedY = Object.keys(rows).sort((a, b) => b - a);
            
            for (const yKey of sortedY) {
                const rowItems = rows[yKey].sort((a, b) => a.x - b.x);
                const rowText = rowItems.map(item => item.text).join(' ');
                
                // Ищем номер отправления (8-10 цифр - 4 цифры - 1-2 цифры)
                const numberMatch = rowText.match(/(\d{8,10}-\d{4}-\d{1,2})/);
                
                if (!numberMatch) {
                    continue;
                }
                
                const number = numberMatch[1];
                
                // Ищем артикул (6 цифр) в той же строке
                const articleMatch = rowText.match(/\b(\d{6})\b/);
                
                if (!articleMatch) {
                    skippedRows++;
                    if (skippedExamples.length < 3) {
                        skippedExamples.push(`Стр.${pageNum}: нет артикула в "${rowText.substring(0, 100)}"`);
                    }
                    continue;
                }
                
                const article = articleMatch[1];
                
                // Ищем название товара (текст после артикула)
                const articleIndex = rowText.indexOf(article);
                const name = rowText.substring(articleIndex + 6).trim();
                
                lpData.push({
                    number: number,
                    article: article,
                    name: name
                });
                
                // Логируем первые 3 успешные строки
                if (lpData.length <= 3) {
                    console.log(`✅ Стр.${pageNum}: номер="${number}", артикул="${article}", название="${name.substring(0, 50)}..."`);
                }
            }
        }
        
        console.log('\n📊 ИТОГИ ЗАГРУЗКИ ЛП:');
        console.log('  ✅ Загружено:', lpData.length, 'товаров');
        console.log('  ️ Пропущено:', skippedRows, 'строк');
        if (skippedExamples.length > 0) {
            console.log('  Примеры пропущенных:', skippedExamples);
        }
        console.log('  📋 Первые 3:', lpData.slice(0, 3));
        console.log('  📋 Последние 3:', lpData.slice(-3));
        
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
            
            // Ищем номер отправления на странице
            const pattern = /\d{8,10}-\d{4}-\d{1,2}/g;
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
                    // Убираем ВСЕ пробелы из артикула
                    const article = String(row[0]).replace(/\s/g, '').trim();
                    
                    // Проверяем что артикул ровно 6 цифр
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
        
        if (!skPdfBytes || skPdfBytes.length === 0) {
            throw new Error('PDF с наклейками не загружен');
        }
        
        console.log('\n📊 ДАННЫЕ:');
        console.log('  sortedData длина:', sortedData.length);
        console.log('  skPages количество:', Object.keys(skPages).length);
        
        const { PDFDocument } = PDFLib;
        const originalPdf = await PDFDocument.load(skPdfBytes, {
            updateMetadata: false,
            ignoreEncryption: true
        });
        
        console.log('  Страниц в оригинале:', originalPdf.getPageCount());
        
        let foundCount = 0;
        let notInSkPages = 0;
        
        for (const item of sortedData) {
            if (skPages[item.number] !== undefined) {
                foundCount++;
            } else {
                notInSkPages++;
            }
        }
        
        console.log('\n🔍 ПРОВЕРКА:');
        console.log('  Найдено в skPages:', foundCount);
        console.log('  Не найдено в skPages:', notInSkPages);
        
        if (foundCount === 0) {
            throw new Error('Ни одна наклейка не найдена в ШК! Проверите файл ШК.');
        }
        
        // Разбиваем на файлы по 999 наклеек
        const MAX_PER_FILE = 999;
        const totalFiles = Math.ceil(sortedData.length / MAX_PER_FILE);
        
        console.log('\n📦 ПЛАН:');
        console.log('  Всего наклеек:', sortedData.length);
        console.log('  Максимум в файле:', MAX_PER_FILE);
        console.log('  Будет создано файлов:', totalFiles);
        
        // Показываем прогресс
        const progressDiv = document.createElement('div');
        progressDiv.style.cssText = 'position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%); background: white; padding: 30px; border-radius: 16px; box-shadow: 0 10px 40px rgba(0,0,0,0.3); z-index: 10000; min-width: 400px;';
        progressDiv.innerHTML = `
            <h3 style="margin-bottom: 20px; color: #1d1d1f;">📥 Создание PDF...</h3>
            <div style="background: #E6E6FA; height: 20px; border-radius: 10px; overflow: hidden; margin-bottom: 15px;">
                <div id="pdfProgressFill" style="background: linear-gradient(90deg, #005BFF, #0046CC); height: 100%; width: 0%; transition: width 0.3s;"></div>
            </div>
            <p id="pdfProgressText" style="margin: 0; color: #86868b; font-size: 0.95em;">Обработка 0 из ${sortedData.length}...</p>
            <p style="margin-top: 15px; font-size: 0.85em; color: #86868b;">⏱️ Пожалуйста, подождите.</p>
        `;
        document.body.appendChild(progressDiv);
        
        let totalCopied = 0;
        
        // Создаём файлы по частям
        for (let fileIndex = 0; fileIndex < totalFiles; fileIndex++) {
            const startIdx = fileIndex * MAX_PER_FILE;
            const endIdx = Math.min(startIdx + MAX_PER_FILE, sortedData.length);
            const batch = sortedData.slice(startIdx, endIdx);
            
            console.log(`\n📄 ФАЙЛ ${fileIndex + 1} из ${totalFiles}:`);
            console.log('  Наклейки:', startIdx + 1, '-', endIdx);
            
            const newPdf = await PDFDocument.create();
            let copiedCount = 0;
            let notFoundInBatch = 0;
            
            for (let i = 0; i < batch.length; i++) {
                const item = batch[i];
                const pageIndex = skPages[item.number];
                
                if (pageIndex !== undefined && pageIndex < originalPdf.getPageCount()) {
                    try {
                        const [copiedPage] = await newPdf.copyPages(originalPdf, [pageIndex]);
                        newPdf.addPage(copiedPage);
                        copiedCount++;
                    } catch (err) {
                        console.error(`  ❌ Ошибка копирования ${item.number}:`, err.message);
                    }
                } else {
                    notFoundInBatch++;
                }
            }
            
            console.log('  ✅ Скопировано:', copiedCount);
            if (notFoundInBatch > 0) {
                console.log('  ⚠️ Не найдено:', notFoundInBatch);
            }
            
            // Сохраняем файл
            const pdfBytes = await newPdf.save();
            const pdfBlob = new Blob([pdfBytes], { type: 'application/pdf' });
            
            const filename = totalFiles === 1 
                ? 'nakleyki_po_poryadku.pdf'
                : `nakleyki_part_${fileIndex + 1}_of_${totalFiles}.pdf`;
            
            console.log('  💾 Сохраняем:', filename, `(${(pdfBytes.length / 1024 / 1024).toFixed(2)} MB)`);
            saveAs(pdfBlob, filename);
            
            totalCopied += copiedCount;
            
            // Обновляем прогресс
            const progress = Math.round((endIdx / sortedData.length) * 100);
            const progressFill = document.getElementById('pdfProgressFill');
            const progressText = document.getElementById('pdfProgressText');
            if (progressFill) progressFill.style.width = progress + '%';
            if (progressText) progressText.textContent = `Файл ${fileIndex + 1} из ${totalFiles}... (${progress}%)`;
            
            // Даём браузеру передохнуть между файлами
            await new Promise(resolve => setTimeout(resolve, 1000));
        }
        
        console.log('\n' + '='.repeat(60));
        console.log('✅ ВСЕ PDF СОЗДАНЫ И СКАЧАНЫ');
        console.log('='.repeat(60));
        console.log('\n ИТОГОВАЯ СВОДКА:');
        console.log('  Всего должно быть:', sortedData.length);
        console.log('  Реально скопировано:', totalCopied);
        console.log('  Создано файлов:', totalFiles);
        
        setTimeout(() => progressDiv.remove(), 1000);
        
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