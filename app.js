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
    
    if (item.path) {
        window.location.href = item.path;
        return;
    }
    
    openInstruction(store, item);
}

function openInstruction(store, item) {
    const content = document.getElementById('instructionContent');
    
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
        setTimeout(() => btn.textContent = '📋 Копировать', 1500);
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