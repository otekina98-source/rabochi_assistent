const stores = [
    { id: 'general', title: 'Инструкции', class: 'instructions', items: allInstructions },
    { id: 'yandex', title: 'ЯндексМаркет', class: 'yandex', items: ymChecks },
    { id: 'wildberries', title: 'ВБ', class: 'wildberries', items: wbChecks },
    { id: 'ozon', title: 'ОЗОН', class: 'ozon', items: ozonChecks }
];

document.addEventListener('DOMContentLoaded', () => {
    renderStores();
    setupSearch();
});

function renderStores() {
    const grid = document.getElementById('storesGrid');
    grid.innerHTML = stores.map(store => `
        <div class="store-card ${store.class}">
            <h3>${store.title}</h3>
            <div class="store-items">
                <ul>
                    ${store.items.map(item => `
                        <li onclick="handleItemClick('${store.id}', '${item.id}')">${item.title}</li>
                    `).join('')}
                </ul>
            </div>
        </div>
    `).join('');
}

function handleItemClick(storeId, itemId) {
    const store = stores.find(s => s.id === storeId);
    const item = store.items.find(i => i.id === itemId);
    
    if (!store || !item) return;
    
    // Если есть путь — переходим на отдельную страницу (сверка-приложение)
    if (item.path) {
        window.location.href = item.path;
        return;
    }
    
    // Иначе — открываем инструкцию
    openInstruction(store, item);
}

function openInstruction(store, item) {
    const content = document.getElementById('instructionContent');
    
    // Если это главная инструкция по сайту — показываем HTML напрямую
    if (item.isMainInstruction) {
        content.innerHTML = `
            <div class="site-instruction">
                <h2> Инструкция по работе с сайтом "Рабочий Ассистент"</h2>
                
                <div class="instruction-section">
                    <h3> Что это за сайт?</h3>
                    <p><strong>Рабочий Ассистент</strong> — это единая платформа для сбора всех сверок и инструкций по работе с маркетплейсами (ЯндексМаркет, ОЗОН, Wildberries).</p>
                    <p><strong>Цель:</strong> собрать все необходимые инструменты в одном месте для удобной и быстрой работы.</p>
                </div>

                <div class="instruction-section warning">
                    <h3>⚠️ ВАЖНО!</h3>
                    <p><strong>Если в инструкции к сверке НЕ написано, что нужно переводить PDF файл в Excel или какой-то другой формат, то мы этого НЕ ДЕЛАЕМ!</strong></p>
                    <p>Загружайте файлы в том формате, в котором они указаны в инструкции (PDF остаётся PDF, Excel остаётся Excel).</p>
                </div>

                <div class="instruction-section">
                    <h3>📚 Разделы сайта</h3>
                    <h4>Инструкция</h4>
                    <p>Здесь описано, как пользоваться сайтом. Выберите инструмент ниже — название ведёт на его страницу.</p>
                    <h4>🚚 ЯндексМаркет</h4>
                    <ul>
                        <li><a href="data/yandex-market/shk-reestr/index.html" class="sverka-link">Сверка ШК и реестр</a> — Загрузите PDF с ШК и вставьте реестр номеров из 1С, затем сравните списки.</li>
                        <li><a href="data/yandex-market/akt-reestr/index.html" class="sverka-link">Сверка АКТ и реестр</a> — Загрузите акт в PDF и вставьте реестр номеров из 1С, чтобы найти расхождения.</li>
                        <li><a href="data/yandex-market/gruzomesta/index.html" class="sverka-link">Сверка по грузоместам</a> — Загрузите Excel из ЛК и добавьте реестр 1С, затем сравните количество грузомест.</li>
                        <li><a href="data/yandex-market/podpis-etiketok/index.html" class="sverka-link">Подпись этикеток</a> — Загрузите реестр 1С, этикетки PDF, прайс и список заказов из ЛК; подпишите этикетки товаров. Останутся заказы с количеством 1 из реестра.</li>
                        <li><a href="data/yandex-market/vycherk-zakazov/index.html" class="sverka-link">Вычерк заказов</a> — Загрузите PDF-акт и добавьте номера ранее отгруженных заказов, чтобы выделить их в акте.</li>
                        <li><a href="data/yandex-market/otmeny/index.html" class="sverka-link">Отмены Яндекс</a> — Вставьте номера отменённых заказов из ЛК и номера из реестра 1С; найдите совпадения — заказы, которые нужно отменить в 1С.</li>
                    </ul>
                    <h4>🏪 ОЗОН</h4>
                    <ul>
                        <li><a href="data/ozon/list-reestr/index.html" class="sverka-link">Сверка Лист подбора(отгрузки) и реестр</a> — Загрузите лист подбора или отгрузки в PDF и добавьте реестр 1С, затем сравните номера отправлений.</li>
                        <li><a href="data/ozon/otmenennye-zakazy/index.html" class="sverka-link">Сверка реестра 1С и отменённых заказов</a> — Загрузите реестр 1С и файл отменённых заказов из ЛК Ozon; получите список заказов, которые нужно отменить в 1С.</li>
                        <li><a href="data/ozon/podbor-stickers/index.html" class="sverka-link">Сортировка ЛП и ШК по порядку БЕЗ картинок</a> — Загрузите прайс Excel, лист подбора PDF и наклейки PDF; скачайте лист подбора без картинок и наклейки в одном алфавитном порядке.</li>
                        <li><a href="data/ozon/sortirovka-LP-kartinki/index.html" class="sverka-link">Сортировка ЛП и ШК с картинками</a> — Загрузите лист сборки PDF, этикетки PDF, прайс и заказы из 1С; скачайте Excel с фото и PDF с ШК в порядке строк листа.</li>
                    </ul>
                    <h4>🏪 Wildberries</h4>
                    <ul>
                        <li><a href="data/wildberries/couriers/index.html" class="sverka-link">Курьеры экспресс</a> — Загрузите Excel с данными курьеров, выберите даты, статус и склады; скачайте два Word-файла — с наименованиями и без.</li>
                        <li><a href="data/wildberries/wb-lp-pics/index.html" class="sverka-link">Сортировка ЛП и ШК по порядку С картинками</a> — Загрузите лист подбора PDF, этикетки ШК PDF, прайс и заказы из 1С; скачайте Excel с фото и этикетками, а также PDF с ШК в том же порядке.</li>
                        <li><a href="data/wildberries/wb-lp-no-pics/index.html" class="sverka-link">Сортировка ЛП и ШК по порядку БЕЗ картинок</a> — Загрузите те же четыре файла, что и в варианте с картинками; скачайте Excel только с колонками «№ документа», «№ задания», «Наименование 1С», «Артикул», «Стикер» и отдельный PDF с ШК в том же порядке.</li>
                    </ul>
                </div>

                <div class="instruction-section">
                    <h3>🔍 Поиск</h3>
                    <p>В верхней части сайта есть строка поиска. Можно искать по названию сверки, по шагам инструкции, по шаблонам сообщений.</p>
                </div>

                <div class="instruction-section">
                    <h3>💡 Советы по работе</h3>
                    <ol>
                        <li><strong>Всегда проверяйте формат файлов</strong> — загружайте только те форматы, которые указаны в инструкции</li>
                        <li><strong>Используйте поиск</strong> — если не можете найти нужную сверку</li>
                        <li><strong>Следите за обновлениями</strong> — инструкция может дополняться новыми сверками</li>
                    </ol>
                </div>

                <div class="instruction-section">
                    <h3>❓ Вопросы</h3>
                    <p>Если возникли вопросы или проблемы:</p>
                    <ol>
                        <li>Проверьте, правильно ли загружены файлы</li>
                        <li>Убедитесь, что формат файлов соответствует инструкции</li>
                        <li>Попробуйте обновить страницу</li>
                        <li><strong>Если выходит ошибка — сделайте скриншот экрана</strong>, чтобы можно было быстро её устранить</li>
                        <li><strong>Обратитесь к разработчику</strong> и отправьте скриншот ошибки</li>
                    </ol>
                </div>

                <div class="instruction-section footer">
                    <p><strong>Удачной работы! </strong></p>
                </div>
            </div>
        `;
    } else {
        // Обычная инструкция
        content.innerHTML = `
            <h2>${item.title}</h2>
            <p>${store.title}</p>
            <div class="instruction-steps">
                <ol>${item.steps.map(s => `<li>${s}</li>`).join('')}</ol>
            </div>
            ${item.messageTemplate ? `
                <div class="message-template">
                    <h4>📝 Шаблон сообщения</h4>
                    <p id="templateText">${item.messageTemplate}</p>
                    <button class="copy-btn" onclick="copyTemplate()">📋 Копировать</button>
                </div>
            ` : ''}
        `;
    }
    
    document.getElementById('instructionScreen').classList.remove('hidden');
    document.getElementById('storesGrid').classList.add('hidden');
    document.querySelector('.search-box').classList.add('hidden');
}

function closeInstruction() {
    document.getElementById('instructionScreen').classList.add('hidden');
    document.getElementById('storesGrid').classList.remove('hidden');
    document.querySelector('.search-box').classList.remove('hidden');
}

function copyTemplate() {
    const text = document.getElementById('templateText').textContent;
    navigator.clipboard.writeText(text).then(() => {
        const btn = document.querySelector('.copy-btn');
        btn.textContent = '✅ Скопировано!';
        setTimeout(() => btn.textContent = ' Копировать', 1500);
    });
}

function setupSearch() {
    const input = document.getElementById('searchInput');
    input.addEventListener('input', (e) => {
        const query = e.target.value.toLowerCase().trim();
        if (query.length < 2) { renderStores(); return; }
        
        const results = [];
        stores.forEach(store => {
            store.items.forEach(item => {
                const inTitle = item.title.toLowerCase().includes(query);
                const inSteps = item.steps && item.steps.some(s => s.toLowerCase().includes(query));
                if (inTitle || inSteps) {
                    results.push({ ...item, storeTitle: store.title, storeId: store.id });
                }
            });
        });
        
        const grid = document.getElementById('storesGrid');
        if (results.length === 0) {
            grid.innerHTML = '<p style="text-align:center; color:#86868b; padding:40px; grid-column:1/-1;">Ничего не найдено</p>';
        } else {
            grid.innerHTML = results.map(item => `
                <div class="store-card" onclick="handleItemClick('${item.storeId}', '${item.id}')">
                    <h3>${item.title}</h3>
                    <p>${item.storeTitle}</p>
                </div>
            `).join('');
        }
    });
}