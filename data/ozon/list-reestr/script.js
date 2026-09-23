let listNumbers = [];
let reestrNumbers = [];
let results = [];
let copiedNumbers = new Set(); // Храним скопированные номера
// Drag & drop
const uploadArea = document.getElementById('pdfUploadArea');
const pdfInput = document.getElementById('pdfFileInput');
uploadArea.addEventListener('click', (e) => {
if (e.target.closest('.upload-btn')) return;
pdfInput.click();
});
uploadArea.addEventListener('dragover', e => {
e.preventDefault();
e.stopPropagation();
uploadArea.classList.add('dragover');
});
uploadArea.addEventListener('dragleave', e => {
e.preventDefault();
e.stopPropagation();
uploadArea.classList.remove('dragover');
});
uploadArea.addEventListener('drop', e => {
e.preventDefault();
e.stopPropagation();
uploadArea.classList.remove('dragover');
const files = e.dataTransfer.files;
console.log('Drop files:', files);
if (files.length > 0) {
const file = files[0];
if (isPdfFile(file)) {
handlePdf(file);
} else {
alert('️ Пожалуйста, выберите PDF файл. Получен файл: ' + file.name);
}
}
});
pdfInput.addEventListener('change', e => {
console.log('Input change:', e.target.files);
if (e.target.files[0]) {
const file = e.target.files[0];
if (isPdfFile(file)) {
handlePdf(file);
} else {
alert('️ Пожалуйста, выберите PDF файл. Получен файл: ' + file.name);
}
}
});
function isPdfFile(file) {
if (file.type === 'application/pdf') return true;
if (file.name && file.name.toLowerCase().endsWith('.pdf')) return true;
return false;
}
// Счётчик номеров в textarea
const reestrText = document.getElementById('reestrText');
const reestrCounter = document.getElementById('reestrCounter');
function updateCounter() {
// Считаем все непустые строки, чтобы не терять некорректно записанные значения
const nums = reestrText.value
.split(/[\r\n]+/)
.map(n => n.trim())
.filter(n => n.length > 0);
reestrCounter.textContent = nums.length + ' ' + pluralize(nums.length, 'номер', 'номера', 'номеров');
}
reestrText.addEventListener('input', updateCounter);
function pluralize(n, one, two, five) {
let abs = Math.abs(n) % 100;
if (abs >= 11 && abs <= 19) return five;
abs = abs % 10;
if (abs === 1) return one;
if (abs >= 2 && abs <= 4) return two;
return five;
}
/**
Нормализация номера заказа
Теперь принимает любой формат "числа-числа-числа" без строгих ограничений на количество цифр
*/
function normalizeOrderNumber(num) {
if (!num) return null;
// Убираем пробелы, заменяем тире
num = num.trim().replace(/\s+/g, '').replace(/[–—]/g, '-');
// Разбиваем по дефису
const parts = num.split('-');
if (parts.length < 3) return null;
// Проверяем, что каждая из первых трёх частей содержит хотя бы одну цифру
if (!parts[0] || !parts[1] || !parts[2]) return null;
// Возвращаем первые три части
return parts.slice(0, 3).join('-');
}
async function handlePdf(file) {
console.log('Начинаем обработку файла:', file);
console.log('Имя:', file.name, 'Размер:', file.size, 'Тип:', file.type);
document.getElementById('pdfFileName').textContent = file.name;
const status = document.getElementById('pdfStatus');
status.textContent = '⏳ Чтение PDF...';
status.className = 'status';
try {
const buffer = await file.arrayBuffer();
console.log('📦 Буфер получен, размер:', buffer.byteLength);
const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
console.log('📑 PDF загружен, страниц:', pdf.numPages);
let text = '';
for (let i = 1; i <= pdf.numPages; i++) {
const page = await pdf.getPage(i);
const content = await page.getTextContent();
text += content.items.map(it => it.str).join(' ') + '\n';
status.textContent = `📄 Страница ${i} из ${pdf.numPages}`;
if (i % 5 === 0) await new Promise(r => setTimeout(r, 0));
}
console.log('Текст извлечён, длина:', text.length);
    // Убираем пробелы вокруг дефисов между цифрами
     text = text.replace(/(\d)\s+[-–—]\s+(\d)/g, '$1-$2');
     text = text.replace(/(\d)\s+[-–—]\s+(\d)/g, '$1-$2');
     // Ищем все последовательности цифр-дефис-цифры-дефис-цифры
     const pattern = /\d+[-–—]\d+[-–—]\d+(?:[-–—]\d+)?/g;
     const rawMatches = text.match(pattern) || [];
     // Нормализуем каждый найденный номер
     const normalizedSet = new Set();
     for (const raw of rawMatches) {
         const normalized = normalizeOrderNumber(raw);
         if (normalized) {
             normalizedSet.add(normalized);
         }
     }
     listNumbers = [...normalizedSet];
     status.textContent = `✅ Найдено ${listNumbers.length} номеров на ${pdf.numPages} стр.`;
     status.className = 'status success';
     console.log('✅ Найдено номеров:', listNumbers.length);
     console.log('Примеры:', listNumbers.slice(0, 10));
     if (listNumbers.length === 0) {
         alert('⚠️ Номера не найдены! Проверьте консоль браузера (F12).');
     }
 } catch (err) {
     console.error('❌ Ошибка обработки PDF:', err);
     status.textContent = '❌ Ошибка: ' + err.message;
     status.className = 'status error';
     alert('Ошибка при чтении PDF: ' + err.message + '\n\nПроверьте консоль браузера (F12) для деталей.');
 }
}
async function startSverka() {
const text = document.getElementById('reestrText').value.trim();
if (!text) { alert('Введите номера из 1С'); return; }
if (listNumbers.length === 0) { alert('Загрузите Лист отгрузки'); return; }
// Парсим реестр 1С с мягкой нормализацией.
// Сохраняем ВСЕ введённые значения: если номер нормализуется — кладём нормализованный,
// иначе — исходное значение. Это гарантирует, что некорректно записанные номера не исчезнут.
const rawLines = text.split(/[\r\n]+/);
const reestrSet = new Set();
for (const line of rawLines) {
    const trimmed = line.trim();
    if (!trimmed) continue;
    const normalized = normalizeOrderNumber(trimmed);
    reestrSet.add(normalized || trimmed);
}
reestrNumbers = [...reestrSet];
if (reestrNumbers.length === 0) {
    alert('Не найдено номеров в реестре 1С');
    return;
}
document.getElementById('progressSection').classList.remove('hidden');
document.getElementById('resultsSection').classList.add('hidden');
document.getElementById('startBtn').disabled = true;
results = [];
copiedNumbers.clear(); // Сбрасываем скопированные номера
// Создаём Set для быстрого поиска O(1)
const reestrSetLookup = new Set(reestrNumbers);
const listSetLookup = new Set(listNumbers);
const total = listNumbers.length + reestrNumbers.length;
let done = 0;
// Находим номера, которые есть в Листе отгрузки но нет в 1С
for (const listNum of listNumbers) {
    if (!reestrSetLookup.has(listNum)) {
        results.push({
            number: listNum,
            inList: true,
            inReestr: false,
            status: 'Есть в Листе отгрузки, нет в 1С'
        });
    }
    done++;
    if (done % 100 === 0 || done === total) {
        const pct = Math.round((done / total) * 100);
        document.getElementById('progressFill').style.width = pct + '%';
        document.getElementById('progressPercent').textContent = pct + '%';
        document.getElementById('progressDetails').textContent = `Обработано ${done}/${total}`;
        await new Promise(r => setTimeout(r, 0));
    }
}
// Находим номера, которые есть в 1С но нет в Листе отгрузки
for (const reestrNum of reestrNumbers) {
    if (!listSetLookup.has(reestrNum)) {
        results.push({
            number: reestrNum,
            inList: false,
            inReestr: true,
            status: 'Есть в 1С, нет в Листе отгрузки'
        });
    }
    done++;
    if (done % 100 === 0 || done === total) {
        const pct = Math.round((done / total) * 100);
        document.getElementById('progressFill').style.width = pct + '%';
        document.getElementById('progressPercent').textContent = pct + '%';
        document.getElementById('progressDetails').textContent = `Обработано ${done}/${total}`;
        await new Promise(r => setTimeout(r, 0));
    }
}
showResults();
document.getElementById('startBtn').disabled = false;
}
function showResults() {
const onlyList = results.filter(r => r.inList && !r.inReestr).length;
const onlyReestr = results.filter(r => !r.inList && r.inReestr).length;
const totalChecked = reestrNumbers.length;
const matched = totalChecked - onlyReestr;
document.getElementById('summary').innerHTML = `
     <div class="summary-card ok">
         <span class="num">${matched}</span>
         <div class="lbl">${pluralize(matched, 'номер совпадает', 'номера совпадают', 'номеров совпадают')}</div>
     </div>
     <div class="summary-card miss">
         <span class="num">${onlyReestr}</span>
         <div class="lbl">Есть в 1С, нет в Листе отгрузки</div>
     </div>
     <div class="summary-card miss">
         <span class="num">${onlyList}</span>
         <div class="lbl">Есть в Листе отгрузки, нет в 1С</div>
     </div>
 `;
 const body = document.getElementById('resultsBody');
 if (results.length === 0) {
     body.innerHTML = `
         <tr><td colspan="4" style="text-align:center; padding:30px; color:#28a745; font-family:-apple-system, sans-serif;">
             ✅ Расхождений не найдено.
         </td></tr>
     `;
 } else {
     const sorted = [...results].sort((a, b) => {
         if (a.inReestr && !b.inReestr) return -1;
         if (!a.inReestr && b.inReestr) return 1;
         return a.number.localeCompare(b.number);
     });
     body.innerHTML = sorted.map(r => {
         const badgeClass = r.inReestr ? 'miss-reestr' : 'miss-list';
         const isCopied = copiedNumbers.has(r.number) ? 'copied-persistent' : '';
         return `<tr>
             <td><span class="num-copy ${isCopied}" title="Нажмите, чтобы скопировать" onclick="copyNumber('${r.number}', this)">${r.number}</span></td>
             <td>${r.inList ? '✅' : '❌'}</td>
             <td>${r.inReestr ? '✅' : '❌'}</td>
             <td><span class="badge ${badgeClass}">${r.status}</span></td>
         </tr>`;
     }).join('');
 }
 document.getElementById('resultsSection').classList.remove('hidden');
}
// Функция копирования одного номера
async function copyNumber(num, el) {
try {
if (navigator.clipboard && navigator.clipboard.writeText) {
await navigator.clipboard.writeText(num);
} else {
fallbackCopyTextToClipboard(num);
}
    // Добавляем в множество скопированных
     copiedNumbers.add(num);
     if (el) {
         el.classList.add('copied-persistent');
     }
     showToast('📋 Номер скопирован: ' + num);
 } catch (err) {
     console.error('Ошибка копирования:', err);
     try {
         fallbackCopyTextToClipboard(num);
         copiedNumbers.add(num);
         if (el) {
             el.classList.add('copied-persistent');
         }
         showToast(' Номер скопирован: ' + num);
     } catch (fallbackErr) {
         showToast(' Не удалось скопировать');
         console.error('Fallback также не сработал:', fallbackErr);
     }
 }
}
function fallbackCopyTextToClipboard(text) {
const textArea = document.createElement('textarea');
textArea.value = text;
textArea.style.position = 'fixed';
textArea.style.top = '0';
textArea.style.left = '0';
textArea.style.width = '2em';
textArea.style.height = '2em';
textArea.style.padding = '0';
textArea.style.border = 'none';
textArea.style.outline = 'none';
textArea.style.boxShadow = 'none';
textArea.style.background = 'transparent';
textArea.style.opacity = '0';
document.body.appendChild(textArea);
textArea.focus();
textArea.select();
try {
const successful = document.execCommand('copy');
if (!successful) throw new Error('execCommand copy failed');
} catch (err) {
console.error('Fallback copy failed:', err);
throw err;
} finally {
document.body.removeChild(textArea);
}
}
// Функция копирования всех номеров
function copyAllNumbers() {
if (results.length === 0) { showToast('Нет номеров для копирования'); return; }
const text = results.map(r => r.number).join('\n');
if (navigator.clipboard && navigator.clipboard.writeText) {
navigator.clipboard.writeText(text).then(() => {
// Помечаем все как скопированные
results.forEach(r => copiedNumbers.add(r.number));
// Обновляем отображение
document.querySelectorAll('.num-copy').forEach(el => {
el.classList.add('copied-persistent');
});
showToast('📋 Скопировано ' + results.length + ' ' + pluralize(results.length, 'номер', 'номера', 'номеров'));
}).catch(err => {
console.error('Ошибка копирования:', err);
fallbackCopyTextToClipboard(text);
results.forEach(r => copiedNumbers.add(r.number));
document.querySelectorAll('.num-copy').forEach(el => {
el.classList.add('copied-persistent');
});
showToast('📋 Скопировано ' + results.length + ' ' + pluralize(results.length, 'номер', 'номера', 'номеров'));
});
} else {
fallbackCopyTextToClipboard(text);
results.forEach(r => copiedNumbers.add(r.number));
document.querySelectorAll('.num-copy').forEach(el => {
el.classList.add('copied-persistent');
});
showToast('📋 Скопировано ' + results.length + ' ' + pluralize(results.length, 'номер', 'номера', 'номеров'));
}
}
function showToast(msg) {
let t = document.getElementById('toast');
if (!t) {
t = document.createElement('div');
t.id = 'toast';
t.className = 'toast';
document.body.appendChild(t);
}
t.textContent = msg;
t.classList.add('show');
clearTimeout(t._timer);
t._timer = setTimeout(() => t.classList.remove('show'), 1800);
}
function exportCSV() {
if (results.length === 0) { showToast('Нет данных для экспорта'); return; }
const csv = [
'Номер;В 1С;В Листе отгрузки;Статус',
...results.map(r =>
`${r.number};${r.inReestr ? 'Да' : 'Нет'};${r.inList ? 'Да' : 'Нет'};${r.status}`
)
].join('\n');
const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
const a = document.createElement('a');
a.href = URL.createObjectURL(blob);
a.download = 'raschozhdeniya_ozon.csv';
a.click();
}