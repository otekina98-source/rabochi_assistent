let actOrderNumbers = new Set();
let ordersData = [];
let matchedOrders = [];
let priceData = {};
let skPdfBytes = null;
let skPages = {};
let finalData = [];

let config = { x: 10, y: 500, fontSize: 5, color: '#000000', fontWeight: 'normal', maxWidth: 250 };
let sampleImageData = null;
let isDragging = false;
let dragOffsetX = 0, dragOffsetY = 0;
let currentSampleOrder = null;
let skFileBuffer = null; 

document.getElementById('actFileInput').addEventListener('change', e => { if (e.target.files[0]) handleAct(e.target.files[0]); });
document.getElementById('ordersFileInput').addEventListener('change', e => { if (e.target.files[0]) handleOrders(e.target.files[0]); });
document.getElementById('priceFileInput').addEventListener('change', e => { if (e.target.files[0]) handlePrice(e.target.files[0]); });
document.getElementById('skFileInput').addEventListener('change', e => { if (e.target.files[0]) handleSK(e.target.files[0]); });

// --- ЧТЕНИЕ ФАЙЛОВ ---

async function handleAct(file) {
    document.getElementById('actFileName').textContent = file.name;
    updateStatus('actStatus', '', 'Чтение...');
    try {
        const buffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
        actOrderNumbers.clear();
        let text = '';
        for (let p = 1; p <= pdf.numPages; p++) {
            const page = await pdf.getPage(p);
            const content = await page.getTextContent();
            text += content.items.map(i => i.str).join(' ') + '\n';
        }
        const matches = text.match(/\b(\d{11})\b/g);
        if (matches) matches.forEach(n => actOrderNumbers.add(n));
        updateStatus('actStatus', '✅', `АКТ: ${actOrderNumbers.size} номеров`);
    } catch (err) { updateStatus('actStatus', '❌', 'Ошибка: ' + err.message); }
}

async function handleOrders(file) {
    document.getElementById('ordersFileName').textContent = file.name;
    updateStatus('ordersStatus', '', 'Чтение...');
    try {
        const data = new Uint8Array(await file.arrayBuffer());
        const wb = XLSX.read(data, { type: 'array' });
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });
        if (rows.length < 2) throw new Error('Мало строк');
        
        const headers = rows[0].map(h => String(h).toLowerCase().trim());
        let oIdx = headers.findIndex(h => h.includes('ваш номер заказа'));
        let sIdx = headers.findIndex(h => h.includes('ваш sku'));
        let cIdx = headers.findIndex(h => h.includes('грузомест'));
        if (oIdx === -1 || sIdx === -1 || cIdx === -1) throw new Error('Не найдены колонки');

        ordersData = [];
        for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            let num = String(row[oIdx] || '').trim().replace(/[\s.]/g, '').replace('+', '');
            const sku = String(row[sIdx] || '').trim().replace(/\s/g, '');
            const cargo = parseInt(String(row[cIdx] || '0').trim()) || 0;
            if (num.includes('E')) num = String(parseFloat(num)).replace('.', '');
            if (num && num.length >= 8 && sku) ordersData.push({ orderNumber: num, sku: sku, cargo: cargo });
        }
        updateStatus('ordersStatus', '✅', `Список заказов: ${ordersData.length} шт`);
    } catch (err) { updateStatus('ordersStatus', '❌', 'Ошибка: ' + err.message); }
}

function handlePrice(file) {
    document.getElementById('priceFileName').textContent = file.name;
    updateStatus('priceStatus', '', 'Чтение...');
    const r = new FileReader();
    r.onload = (e) => {
        try {
            const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
            const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });
            priceData = {};
            rows.forEach(row => {
                const sku = String(row[0]).replace(/\s/g, '').trim();
                const name = String(row[1]).trim();
                if (sku && name && sku.length > 5) priceData[sku] = name;
            });
            updateStatus('priceStatus', '✅', `ПРАЙС: ${Object.keys(priceData).length} тов.`);
        } catch (err) { updateStatus('priceStatus', '❌', 'Ошибка: ' + err.message); }
    };
    r.readAsArrayBuffer(file);
}

async function handleSK(file) {
    document.getElementById('skFileName').textContent = file.name;
    updateStatus('skStatus', '', 'Чтение...');
    try {
        skFileBuffer = await file.arrayBuffer();
        const workBuffer = skFileBuffer.slice(0);
        skPdfBytes = new Uint8Array(workBuffer);
        
        const pdf = await pdfjsLib.getDocument({ data: workBuffer }).promise;
        skPages = {};
        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const txt = (await page.getTextContent()).items.map(it => it.str).join(' ');
            const m = txt.match(/\b(\d{11})\b/);
            if (m) skPages[m[1]] = i - 1;
        }
        updateStatus('skStatus', '✅', `ШК (общий): ${Object.keys(skPages).length} наклеек`);
    } catch (err) { 
        updateStatus('skStatus', '❌', 'Ошибка: ' + err.message); 
    }
}

// --- ЛОГИКА ОБРАБОТКИ ---

function checkReady() {
    const ready = matchedOrders.length > 0 && Object.keys(priceData).length > 0 && skPdfBytes !== null;
    document.getElementById('configBtn').disabled = !ready;
    document.getElementById('pdfBtn').disabled = !ready || (config.x === 0 && config.y === 0);
}

function processOrders() {
    try {
        if (actOrderNumbers.size === 0) return alert('Загрузите АКТ!');
        if (ordersData.length === 0) return alert('Загрузите Список заказов!');
        
        showProgress('Поиск совпадений...', 10);
        matchedOrders = ordersData.filter(o => actOrderNumbers.has(o.orderNumber) && o.cargo === 1);
        
        finalData = [];
        for (const order of matchedOrders) {
            const name = priceData[order.sku] || '';
            if (name) finalData.push({ orderNumber: order.orderNumber, sku: order.sku, name: name, cargo: order.cargo });
        }
        
        updateProgress(100, 'Готово!');
        setTimeout(() => {
            document.getElementById('progressSection').classList.add('hidden');
            document.getElementById('processBtn').disabled = false;
        }, 500);
        alert(`✅ Найдено ${matchedOrders.length} заказов`);
        showPreview();
        checkReady();
    } catch (e) { alert('Ошибка: ' + e.message); hideProgress(); }
}

// --- НАСТРОЙКА НА ХОЛСТЕ ---

async function openConfig() {
    if (!skFileBuffer) return alert('Сначала загрузите ШК (общий файл)!');
    if (finalData.length === 0) return alert('Сначала нажмите "ОБРАБОТАТЬ"!');

    document.getElementById('configOverlay').style.display = 'flex';
    const canvas = document.getElementById('configCanvas');
    const ctx = canvas.getContext('2d');

    try {
        const renderBuffer = skFileBuffer.slice(0);
        const pdf = await pdfjsLib.getDocument({ data: renderBuffer }).promise;
        const page = await pdf.getPage(1);
        const viewport = page.getViewport({ scale: 2 });
        canvas.width = viewport.width;
        canvas.height = viewport.height;

        await page.render({ canvasContext: ctx, viewport: viewport }).promise;
        sampleImageData = ctx.getImageData(0, 0, canvas.width, canvas.height);

        currentSampleOrder = finalData[0];
        document.getElementById('configPreviewName').textContent = '"' + currentSampleOrder.name + '"';
        document.getElementById('fontSizeInput').value = config.fontSize;
        document.getElementById('colorInput').value = config.color;
        document.getElementById('fontWeightInput').value = config.fontWeight;
        document.getElementById('maxWidthInput').value = config.maxWidth;

        drawLabel(ctx);
        setupCanvasEvents();

        document.getElementById('fontSizeInput').oninput = (e) => { config.fontSize = parseFloat(e.target.value); drawLabel(ctx); };
        document.getElementById('colorInput').oninput = (e) => { config.color = e.target.value; drawLabel(ctx); };
        document.getElementById('fontWeightInput').onchange = (e) => { config.fontWeight = e.target.value; drawLabel(ctx); };
        document.getElementById('maxWidthInput').oninput = (e) => { config.maxWidth = parseInt(e.target.value); drawLabel(ctx); };

    } catch (error) {
        alert('Ошибка отрисовки: ' + error.message);
        document.getElementById('configOverlay').style.display = 'none';
    }
}

function drawLabel(ctx) {
    const canvas = document.getElementById('configCanvas');
    ctx.putImageData(sampleImageData, 0, 0);
    if (!currentSampleOrder) return;

    ctx.font = `${config.fontWeight} ${config.fontSize}pt Helvetica`;
    ctx.fillStyle = config.color;
    ctx.textBaseline = 'bottom';

    const words = currentSampleOrder.name.split(' ');
    let lines = [], currentLine = words[0];
    for (let i = 1; i < words.length; i++) {
        const testLine = currentLine + ' ' + words[i];
        if (ctx.measureText(testLine).width > config.maxWidth) { lines.push(currentLine); currentLine = words[i]; } 
        else { currentLine = testLine; }
    }
    lines.push(currentLine);

    let currentY = config.y;
    for (let i = lines.length - 1; i >= 0; i--) {
        ctx.fillText(lines[i], config.x, currentY);
        currentY -= config.fontSize * 1.2;
    }
}

function setupCanvasEvents() {
    const canvas = document.getElementById('configCanvas');
    const newCanvas = canvas.cloneNode(true);
    canvas.parentNode.replaceChild(newCanvas, canvas);
    
    newCanvas.addEventListener('mousedown', (e) => {
        const rect = newCanvas.getBoundingClientRect();
        const scaleX = newCanvas.width / rect.width;
        const scaleY = newCanvas.height / rect.height;
        const clickX = (e.clientX - rect.left) * scaleX;
        const clickY = (e.clientY - rect.top) * scaleY;

        const ctx = newCanvas.getContext('2d');
        ctx.font = `${config.fontWeight} ${config.fontSize}pt Helvetica`;
        const tw = ctx.measureText(currentSampleOrder.name).width;
        const th = config.fontSize * 1.2;

        if (clickX >= config.x && clickX <= config.x + Math.min(tw, config.maxWidth) + 20 &&
            clickY >= config.y - th - 10 && clickY <= config.y + 10) {
            isDragging = true;
            dragOffsetX = clickX - config.x;
            dragOffsetY = clickY - config.y;
        } else {
            config.x = Math.round(clickX);
            config.y = Math.round(clickY);
            drawLabel(ctx);
        }
    });

    newCanvas.addEventListener('mousemove', (e) => {
        if (!isDragging) return;
        const rect = newCanvas.getBoundingClientRect();
        const scaleX = newCanvas.width / rect.width;
        const scaleY = newCanvas.height / rect.height;
        config.x = Math.round((e.clientX - rect.left) * scaleX - dragOffsetX);
        config.y = Math.round((e.clientY - rect.top) * scaleY - dragOffsetY);
        drawLabel(newCanvas.getContext('2d'));
    });

    window.addEventListener('mouseup', () => { isDragging = false; });
}

function closeConfig() { document.getElementById('configOverlay').style.display = 'none'; }

function saveConfig() {
    if (config.x === 0 && config.y === 0) return alert('Кликните по месту на наклейке!');
    document.getElementById('configOverlay').style.display = 'none';
    document.getElementById('pdfBtn').disabled = false;
    alert('✅ Настройки сохранены! Скачайте PDF.');
}

// --- ПРЕВЬЮ И СКАЧИВАНИЕ ---

function showPreview() {
    document.getElementById('resultSection').classList.remove('hidden');
    document.getElementById('tableCount').textContent = final