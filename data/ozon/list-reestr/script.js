// Улучшенная функция копирования с обработкой ошибок и fallback
async function copyNumber(num, el) {
    try {
        // Проверяем поддержку современного API
        if (navigator.clipboard && navigator.clipboard.writeText) {
            await navigator.clipboard.writeText(num);
        } else {
            // Fallback для старых браузеров
            fallbackCopyTextToClipboard(num);
        }
        
        // Визуальная обратная связь
        if (el) {
            const originalText = el.textContent;
            el.textContent = '✓ Скопировано';
            el.classList.add('copied');
            
            setTimeout(() => {
                el.textContent = originalText;
                el.classList.remove('copied');
            }, 2000);
        }
        
        showToast('📋 Номер скопирован: ' + num);
        
    } catch (err) {
        console.error('Ошибка копирования:', err);
        // Пробуем fallback при ошибке
        try {
            fallbackCopyTextToClipboard(num);
            if (el) {
                el.textContent = '✓ Скопировано';
                el.classList.add('copied');
                setTimeout(() => {
                    el.textContent = num;
                    el.classList.remove('copied');
                }, 2000);
            }
            showToast('📋 Номер скопирован: ' + num);
        } catch (fallbackErr) {
            showToast('❌ Не удалось скопировать');
            console.error('Fallback также не сработал:', fallbackErr);
        }
    }
}

// Fallback функция для старых браузеров и HTTP
function fallbackCopyTextToClipboard(text) {
    const textArea = document.createElement('textarea');
    textArea.value = text;
    
    // Стили чтобы не мелькало
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
        if (!successful) {
            throw new Error('execCommand copy failed');
        }
    } catch (err) {
        console.error('Fallback copy failed:', err);
        throw err;
    } finally {
        document.body.removeChild(textArea);
    }
}

// Функция копирования всех номеров
function copyAllNumbers() {
    const diff = results.filter(r => !(r.inList && r.inReestr));
    if (diff.length === 0) { 
        showToast('Нет номеров для копирования'); 
        return; 
    }
    
    const text = diff.map(r => r.number).join('\n');
    
    // Используем нашу улучшенную функцию
    if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text).then(() => {
            showToast(`📋 Скопировано ${diff.length} номеров`);
        }).catch(err => {
            console.error('Ошибка копирования:', err);
            fallbackCopyTextToClipboard(text);
            showToast(`📋 Скопировано ${diff.length} номеров`);
        });
    } else {
        fallbackCopyTextToClipboard(text);
        showToast(` Скопировано ${diff.length} номеров`);
    }
}