let excelData = []; // все записи из Excel
let filteredData = []; // отфильтрованные данные
let groupedData = {}; // сгруппированные данные

// Drag & drop
const uploadArea = document.getElementById('excelUploadArea');
const excelInput = document.getElementById('excelFileInput');

uploadArea.addEventListener('dragover', e => { e.preventDefault(); uploadArea.classList.add('dragover'); });
uploadArea.addEventListener('dragleave', () => uploadArea.classList.remove('dragover'));
uploadArea.addEventListener('drop', e => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    if (file) handleExcel(file);
});

excelInput.addEventListener('change', e => { if (e.target.files[0]) handleExcel(e.target.files[0]); });

function handleExcel(file) {
    document.getElementById('excelFileName').textContent = file.name;
    const status = document.getElementById('excelStatus');
    status.textContent = '⏳ Чтение Excel...';
    status.className = 'status';
    
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array', cellDates: true });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet, { defval: '' });
            
            console.log('Сырые данные:', jsonData.length, 'записей');
            console.log('Первые 3 записи:', jsonData.slice(0, 3));
            
            // Нормализуем данные
            excelData = jsonData.map(row => {
                const normalized = {};
                for (const key in row) {
                    const lowerKey = key.toLowerCase().trim();
                    if (lowerKey.includes('стикер')) normalized.sticker = String(row[key]).trim();
                    else if (lowerKey.includes('наименование')) normalized.name = String(row[key]).trim();
                    else if (lowerKey.includes('склад')) {
                        const warehouseRaw = String(row[key]).trim();
                        // Определяем короткий код склада
                        if (warehouseRaw.includes('МБТ') || warehouseRaw.includes('30825')) {
                            normalized.warehouse = 'МБТ';
                            normalized.warehouseFull = warehouseRaw;
                        } else if (warehouseRaw.includes('СГТ')) {
                            normalized.warehouse = 'СГТ';
                            normalized.warehouseFull = warehouseRaw;
                        } else if (warehouseRaw.includes('СПБ')) {
                            normalized.warehouse = 'СПБ';
                            normalized.warehouseFull = warehouseRaw;
                        } else {
                            normalized.warehouse = warehouseRaw;
                            normalized.warehouseFull = warehouseRaw;
                        }
                    }
                    else if (lowerKey.includes('дата') && lowerKey.includes('доставк')) {
                        normalized.deliveryDate = formatDate(row[key]);
                    }
                    else if (lowerKey.includes('статус')) normalized.status = String(row[key]).trim();
                    else if (lowerKey.includes('фио') && lowerKey.includes('курьер')) normalized.courierName = String(row[key]).trim();
                    else if (lowerKey.includes('телефон') && lowerKey.includes('курьер')) normalized.courierPhone = String(row[key]).trim();
                    else if (lowerKey.includes('номер') && lowerKey.includes('авто')) normalized.courierCar = String(row[key]).trim();
                }
                return normalized;
            }).filter(row => row.sticker); // убираем пустые
            
            console.log('После нормализации:', excelData.length, 'записей');
            console.log('Пример:', excelData.slice(0, 2));
            
            status.textContent = `✅ Загружено ${excelData.length} записей`;
            status.className = 'status success';
            
            document.getElementById('recordCount').textContent = excelData.length;
            document.getElementById('dataPreview').classList.remove('hidden');
            document.getElementById('generateBtn').disabled = false;
            
            // Заполняем фильтры дат
            populateDateFilters();
            
        } catch (err) {
            console.error('Ошибка загрузки:', err);
            status.textContent = ' Ошибка: ' + err.message;
            status.className = 'status error';
        }
    };
    reader.readAsArrayBuffer(file);
}

function formatDate(date) {
    if (!date) return '';
    if (date instanceof Date) {
        const d = date.getDate().toString().padStart(2, '0');
        const m = (date.getMonth() + 1).toString().padStart(2, '0');
        const y = date.getFullYear();
        return `${d}.${m}.${y}`;
    }
    return String(date).trim();
}

function populateDateFilters() {
    const dates = [...new Set(excelData.map(r => r.deliveryDate).filter(d => d))];
    dates.sort();
    
    const container = document.getElementById('dateFilters');
    if (dates.length === 0) {
        container.innerHTML = '<p class="empty-text">Даты не найдены</p>';
        return;
    }
    
    container.innerHTML = dates.map(date => `
        <label class="checkbox-label">
            <input type="checkbox" name="date" value="${date}" checked>
            <span>${date}</span>
        </label>
    `).join('');
}

async function generateFiles() {
    try {
        // Получаем выбранные фильтры
        const selectedDates = [...document.querySelectorAll('input[name="date"]:checked')].map(cb => cb.value);
        const statusFilter = document.getElementById('statusFilter').value.trim().toLowerCase();
        const selectedWarehouses = [...document.querySelectorAll('input[name="warehouse"]:checked')].map(cb => cb.value);
        
        if (selectedDates.length === 0) { alert('Выберите хотя бы одну дату'); return; }
        if (selectedWarehouses.length === 0) { alert('Выберите хотя бы один склад'); return; }
        
        console.log('========================================');
        console.log('Всего загружено:', excelData.length, 'записей');
        console.log('Выбранные склады:', selectedWarehouses);
        console.log('Выбранные даты:', selectedDates);
        console.log('Фильтр статуса:', statusFilter);
        console.log('========================================');
        
        // Проверяем каждый заказ
        const lostOrders = [];
        filteredData = excelData.filter((row, index) => {
            const dateMatch = selectedDates.includes(row.deliveryDate);
            const statusMatch = !statusFilter || row.status.toLowerCase().includes(statusFilter);
            const warehouseMatch = selectedWarehouses.includes(row.warehouse);
            
            const isMatch = dateMatch && statusMatch && warehouseMatch;
            
            if (!isMatch) {
                lostOrders.push({
                    index: index + 1,
                    sticker: row.sticker,
                    warehouse: row.warehouse,
                    deliveryDate: row.deliveryDate,
                    status: row.status,
                    reason: !warehouseMatch ? 'СКЛАД' : (!dateMatch ? 'ДАТА' : 'СТАТУС')
                });
            }
            
            return isMatch;
        });
        
        console.log('========================================');
        console.log('Прошло фильтрацию:', filteredData.length, 'заказов');
        console.log('Не прошло:', lostOrders.length, 'заказов');
        console.log('Потерянные заказы:');
        lostOrders.forEach(order => {
            console.log(`#${order.index} ${order.sticker} - причина: ${order.reason}`);
            console.log(`  Склад: "${order.warehouse}" (выбрано: ${selectedWarehouses})`);
            console.log(`  Дата: "${order.deliveryDate}" (выбрано: ${selectedDates})`);
            console.log(`  Статус: "${order.status}" (фильтр: "${statusFilter}")`);
        });
        console.log('========================================');
        
        if (filteredData.length === 0) {
            alert('Нет данных по выбранным фильтрам');
            return;
        }
        
        // Группируем по складам
        groupedData = {};
        selectedWarehouses.forEach(w => groupedData[w] = []);
        
        filteredData.forEach(row => {
            if (groupedData[row.warehouse]) {
                groupedData[row.warehouse].push(row);
            }
        });
        
        console.log('Сгруппированные данные:');
        Object.entries(groupedData).forEach(([warehouse, orders]) => {
            console.log(`  ${warehouse}: ${orders.length} заказов`);
        });
        
        // Показываем предпросмотр
        showPreview();
        
    } catch (error) {
        console.error('Ошибка генерации:', error);
        alert('Ошибка при формировании: ' + error.message);
    }
}

function showPreview() {
    // Создаём модальное окно с предпросмотром
    const modal = document.createElement('div');
    modal.className = 'preview-modal';
    modal.innerHTML = `
        <div class="preview-content">
            <div class="preview-header">
                <h2> Предпросмотр</h2>
                <button class="close-btn" onclick="this.closest('.preview-modal').remove()">✕</button>
            </div>
            <div class="preview-info">
                <p><strong>Всего записей:</strong> ${filteredData.length}</p>
                <p><strong>По складам:</strong></p>
                <ul>
                    ${Object.entries(groupedData).map(([warehouse, orders]) => 
                        `<li>${warehouse}: ${orders.length} заказов</li>`
                    ).join('')}
                </ul>
            </div>
            <div class="preview-actions">
                <button class="preview-btn" onclick="showTextPreview()">👁️ Показать текстом</button>
                <button class="download-btn" onclick="downloadFiles()">📥 Скачать файлы</button>
            </div>
        </div>
    `;
    document.body.appendChild(modal);
}

function showTextPreview() {
    let previewText = '';
    
    for (const warehouse of Object.keys(groupedData)) {
        const orders = groupedData[warehouse];
        if (orders.length === 0) continue;
        
        previewText += `\n\n=== СКЛАД: ${warehouse} ===\n\n`;
        
        // Без курьера
        const withoutCourier = orders.filter(o => !o.courierName);
        if (withoutCourier.length > 0) {
            previewText += 'НЕТ ДАННЫХ! ПРИШЛЮ ПОЗЖЕ!\n';
            withoutCourier.forEach((order, idx) => {
                previewText += `${idx + 1}. ${order.sticker}\n`;
            });
            previewText += '\n';
        }
        
        // С курьерами
        const withCourier = orders.filter(o => o.courierName);
        const byCourier = {};
        withCourier.forEach(order => {
            const key = order.courierName;
            if (!byCourier[key]) byCourier[key] = { info: order, orders: [] };
            byCourier[key].orders.push(order);
        });
        
        for (const courierName in byCourier) {
            const courier = byCourier[courierName];
            previewText += `ФИО: ${courier.info.courierName}\n`;
            if (courier.info.courierPhone) previewText += `Телефон: ${courier.info.courierPhone}\n`;
            if (courier.info.courierCar) previewText += `Номер авто: ${courier.info.courierCar}\n`;
            previewText += `Склад: ${warehouse}\n\n`;
            
            courier.orders.forEach((order, idx) => {
                previewText += `${idx + 1}. ${order.sticker}\n`;
            });
            previewText += '\n';
        }
    }
    
    // Показываем в новом окне
    const previewWindow = window.open('', '_blank');
    previewWindow.document.write(`
        <html>
        <head><title>Предпросмотр</title></head>
        <body style="font-family: monospace; white-space: pre-wrap;">
            <h2>Предпросмотр файлов</h2>
            <pre>${previewText}</pre>
        </body>
        </html>
    `);
    previewWindow.document.close();
}

async function downloadFiles() {
    try {
        document.querySelector('.preview-modal').remove();
        
        const progressModal = document.createElement('div');
        progressModal.className = 'progress-modal';
        progressModal.innerHTML = `
            <div class="progress-content">
                <h3>Формирование файлов...</h3>
                <div class="progress-bar"><div class="progress-fill" style="width: 0%"></div></div>
                <p class="progress-text">0%</p>
            </div>
        `;
        document.body.appendChild(progressModal);
        
        // Файл с наименованиями
        updateProgress(progressModal, 20, 'Создание файла с наименованиями...');
        const docWithName = createDocument(groupedData, true);
        const blobWithName = await docx.Packer.toBlob(docWithName);
        
        updateProgress(progressModal, 60, 'Сохранение...');
        saveAs(blobWithName, `kurery_s_naimenovaniyami.docx`);
        
        // Файл без наименований
        updateProgress(progressModal, 80, 'Создание файла без наименований...');
        const docWithoutName = createDocument(groupedData, false);
        const blobWithoutName = await docx.Packer.toBlob(docWithoutName);
        
        updateProgress(progressModal, 100, 'Готово!');
        saveAs(blobWithoutName, `kurery_bez_naimenovaniy.docx`);
        
        setTimeout(() => {
            progressModal.remove();
        }, 1500);
        
    } catch (error) {
        console.error('Ошибка скачивания:', error);
        alert('Ошибка при скачивании: ' + error.message);
    }
}

function updateProgress(modal, percent, text) {
    modal.querySelector('.progress-fill').style.width = percent + '%';
    modal.querySelector('.progress-text').textContent = text;
}

function createDocument(byWarehouse, withNames) {
    const children = [];
    
    for (const warehouse of Object.keys(byWarehouse)) {
        const orders = byWarehouse[warehouse];
        if (orders.length === 0) continue;
        
        // Заголовок склада
        children.push(new docx.Paragraph({
            children: [new docx.TextRun({ text: `Склад: ${warehouse}`, bold: true, size: 28 })],
            spacing: { before: 300, after: 200 }
        }));
        
        const withoutCourier = orders.filter(o => !o.courierName);
        const withCourier = orders.filter(o => o.courierName);
        
        // Без курьера
        if (withoutCourier.length > 0) {
            children.push(new docx.Paragraph({
                children: [new docx.TextRun({ text: 'НЕТ ДАННЫХ! ПРИШЛЮ ПОЗЖЕ!', bold: true, color: 'FF0000', size: 24 })],
                spacing: { before: 200, after: 100 }
            }));
            
            withoutCourier.forEach((order, idx) => {
                const text = withNames 
                    ? `${idx + 1}. ${order.sticker} - ${order.name}`
                    : `${idx + 1}. ${order.sticker}`;
                children.push(new docx.Paragraph({
                    children: [new docx.TextRun({ text, size: 22 })],
                    spacing: { after: 60 }
                }));
            });
        }
        
        // С курьерами
        const byCourier = {};
        withCourier.forEach(order => {
            const key = order.courierName;
            if (!byCourier[key]) byCourier[key] = { info: order, orders: [] };
            byCourier[key].orders.push(order);
        });
        
        for (const courierName in byCourier) {
            const courier = byCourier[courierName];
            
            children.push(new docx.Paragraph({
                children: [new docx.TextRun({ text: `ФИО: ${courier.info.courierName}`, bold: true, size: 24 })],
                spacing: { before: 300, after: 60 }
            }));
            
            if (courier.info.courierPhone) {
                children.push(new docx.Paragraph({
                    children: [new docx.TextRun({ text: `Телефон: ${courier.info.courierPhone}`, size: 22 })],
                    spacing: { after: 60 }
                }));
            }
            
            if (courier.info.courierCar) {
                children.push(new docx.Paragraph({
                    children: [new docx.TextRun({ text: `Номер авто: ${courier.info.courierCar}`, size: 22 })],
                    spacing: { after: 60 }
                }));
            }
            
            children.push(new docx.Paragraph({
                children: [new docx.TextRun({ text: `Склад: ${warehouse}`, size: 22 })],
                spacing: { after: 150 }
            }));
            
            courier.orders.forEach((order, idx) => {
                const text = withNames 
                    ? `${idx + 1}. ${order.sticker} - ${order.name}`
                    : `${idx + 1}. ${order.sticker}`;
                children.push(new docx.Paragraph({
                    children: [new docx.TextRun({ text, size: 22 })],
                    spacing: { after: 60 }
                }));
            });
        }
        
        children.push(new docx.Paragraph({
            children: [new docx.TextRun({ text: '' })],
            spacing: { before: 200 }
        }));
    }
    
    return new docx.Document({
        sections: [{
            properties: {},
            children
        }]
    });
}