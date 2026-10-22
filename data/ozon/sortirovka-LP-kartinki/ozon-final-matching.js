/* Один окончательный модуль сопоставления Ozon. Подключать последним.
   Перед сборкой перечитывает PDF и строит единый полный индекс страниц.
   Excel сохраняет номера ШК из листа сборки, PDF использует точно такие же ключи.
   Если нет полного соответствия, выдача PDF запрещена. */
(function () {
    'use strict';
    const clean = v => String(v == null ? '' : v)
       .replace(/[\s\u00a0\u2007\u202f\u0456\u0406]/g, c => /[\u0456\u0406]/.test(c) ? 'i' : '')
       .toLowerCase();
    const II = /^ii\d{8,16}$/;
    const POST = /^\d{5,10}-\d{3,5}-\d{1,3}$/;

    function extractPageId(tc) {
        const chunks = tc.items.map(it => String(it.str || '').trim()).filter(Boolean);
        // В реальном PDF всегда две части: "ii5005890" и " 3451".
        // Нельзя пытаться сопоставлять номера, ориентируясь лишь на суффикс.
        const candidates = new Set();
        for (let i = 0; i < chunks.length; i++) {
            const a = clean(chunks[i]);
            if (II.test(a)) candidates.add(a);
            if (/^ii\d{4,12}$/.test(a) && i + 1 < chunks.length) {
                const b = clean(chunks[i + 1]);
                if (/^\d{4}$/.test(b) && II.test(a + b)) candidates.add(a + b);
            }
        }
        // Строго полное значение. Числа адреса не используем.
        const text = chunks.join(' ');
        for (const m of text.matchAll(/(?:^|[^a-z0-9])(ii\s*\d{4,12}\s+\d{4})(?!\d)/gi)) {
            const id = clean(m[1]);
            if (II.test(id)) candidates.add(id);
        }
        if (candidates.size === 1) return [...candidates][0];
        if (candidates.size > 1) throw new Error('Несколько возможных ii-номеров на одной наклейке: ' + [...candidates].join(', '));
        // Старые этикетки обрабатывает существующая проверенная функция.
        const old = clean(postingFromPage(tc));
        return POST.test(old) ? old : null;
    }

    async function rebuildIndex(shk) {
        const doc = await pdfjsLib.getDocument({
            data: new Uint8Array(shk.buffer).slice(), disableFontFace: true
        }).promise;
        const pageOf = new Map(), pageToNum = new Map(), unreadable = [];
        for (let p = 1; p <= doc.numPages; p++) {
            const tc = await (await doc.getPage(p)).getTextContent();
            const id = extractPageId(tc);
            if (!id) { unreadable.push(p); continue; }
            if (!pageOf.has(id)) pageOf.set(id, []);
            pageOf.get(id).push(p);
            pageToNum.set(p, id);
        }
        if (unreadable.length) {
            throw new Error('Полный номер не распознан на страницах PDF: ' + unreadable.join(', '));
        }
        shk.pageOf = pageOf;
        shk.pageToNum = pageToNum;
        shk.pages = doc.numPages;
        console.info('Ozon ШК: прочитаны страницы', [...pageToNum.entries()]);
    }

    // PDF собирается по реальным идентификаторам страниц, а не по псевдонимам
    // из предыдущих вариантов кода. Отдельный поиск по 4 цифрам запрещён.
    buildShkPdf = async function (rows, shk) {
        const used = new Set(), order = [], missing = [], duplicated = [];
        const owners = new Map((listData?.rows || []).map(r => [clean(r.posting), clean(r.shkId)]));
        for (const row of rows) {
            const posting = clean(row.posting);
            const id = clean(row.shkId || owners.get(posting));
            const key = id || posting;
            if ((id && !II.test(id)) || (!id && !POST.test(posting))) {
                missing.push(posting + ' → неверный идентификатор ' + key);
                continue;
            }
            const found = [];
            for (const [pg, pageId] of shk.pageToNum.entries()) {
                if (clean(pageId) === key) found.push(pg);
            }
            if (found.length !== 1) {
                (found.length ? duplicated : missing).push(posting + ' → ' + key);
                continue;
            }
            if (used.has(found[0])) { duplicated.push(posting + ' → страница ' + found[0]); continue; }
            used.add(found[0]);
            order.push(found[0] - 1);
        }
        const extras = [];
        for (let p=1; p<=shk.pages; p++) if (!used.has(p)) extras.push(p);
        if (missing.length || duplicated.length || extras.length || order.length !== rows.length) {
            throw new Error('Наклейки не полностью соответствуют листу подбора. PDF не создан. ' +
               (missing.length ? 'Нет совпадений: ' + missing.join('; ') + '. ' : '') +
               (duplicated.length ? 'Повторы: ' + duplicated.join('; ') + '. ' : '') +
               (extras.length ? 'Неиспользованные страницы: ' + extras.join(', ') : ''));
        }
        const source = await PDFLib.PDFDocument.load(shk.buffer, {ignoreEncryption:true});
        const dest = await PDFLib.PDFDocument.create();
        const pages = await dest.copyPages(source, order);
        pages.forEach(p => dest.addPage(p));
        console.info('Ozon: порядок страниц PDF по Excel', order.map(p => p+1));
        return dest.save();
    };

    const previousBuild = startBuild;
    startBuild = async function () {
        if (shkData?.buffer) {
            try { await rebuildIndex(shkData); }
            catch(e) { console.error(e); alert('Ошибка чтения ШК: ' + e.message); return; }
        }
        return previousBuild();
    };
})();
