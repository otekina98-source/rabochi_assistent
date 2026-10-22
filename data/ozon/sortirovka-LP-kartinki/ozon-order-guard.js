/* Единое сопоставление Excel и PDF. Подключается ПОСЛЕ двух других патчей.
   Наклейка для строки = её полный ШК (ii...) либо номер отправления.
   Никогда не подставляет страницы без совпадения и не выдаёт PDF с неверным порядком. */
(function () {
    'use strict';
    const clean = value => String(value == null ? '' : value)
        .replace(/[\s\u00a0\u202f\u2007]+/g, '').toLowerCase();
    const validII = /^ii\d{8,16}$/;
    const validPosting = /^\d{5,10}-\d{3,5}-\d{1,3}$/;

    function createPlan(rows, shk) {
        if (!shk || !shk.pageOf || !shk.buffer) throw new Error('Не загружены этикетки PDF');
        const known = new Map();
        for (const row of (listData && listData.rows || [])) {
            known.set(clean(row.posting), clean(row.shkId));
        }
        const used = new Set();
        const order = [];
        const missing = [];
        const ambiguous = [];
        for (const row of rows) {
            const posting = clean(row.posting);
            const id = clean(row.shkId || known.get(posting));
            const key = id || posting;
            if (id && !validII.test(id)) {
                ambiguous.push(`${posting}: неверный ШК ${id}`);
                continue;
            }
            if (!id && !validPosting.test(posting)) {
                ambiguous.push(`${posting}: неверный номер отправления`);
                continue;
            }
            // Для новых заказов выбираем только ii-ключ, для старых — полный номер отправления.
            // Алиас по номеру заказа, созданный прежним патчем, допустим исключительно если
            // страница в действительности имеет полный ii-идентификатор этой строки.
            let pages = shk.pageOf.get(key) || [];
            if (id && !pages.length) pages = shk.pageOf.get(posting) || [];
            pages = [...new Set(pages)].filter(p => Number.isInteger(p) && p >= 1 && p <= shk.pages);
            const exact = pages.filter(p => clean(shk.pageToNum && shk.pageToNum.get(p)) === key);
            if (exact.length !== 1) {
                if (!exact.length) missing.push(`${posting} → ${key}`);
                else ambiguous.push(`${posting}: ${exact.length} этикеток с одинаковым ШК ${key}`);
                continue;
            }
            if (used.has(exact[0])) {
                ambiguous.push(`${posting}: страница ${exact[0]} уже назначена`);
                continue;
            }
            used.add(exact[0]);
            order.push(exact[0] - 1);
        }
        const extra = [];
        for (let p = 1; p <= shk.pages; p++) if (!used.has(p)) extra.push(p);
        if (missing.length || ambiguous.length || extra.length || order.length !== rows.length) {
            const details = [
                missing.length && `Не найдены: ${missing.slice(0, 5).join('; ')}`,
                ambiguous.length && `Неоднозначные: ${ambiguous.slice(0, 5).join('; ')}`,
                extra.length && `Неиспользованные страницы PDF: ${extra.slice(0, 15).join(', ')}`
            ].filter(Boolean).join('\n');
            throw new Error('Этикетки не полностью соответствуют листу подбора. PDF не будет создан в неправильном порядке.\n' + details);
        }
        return order;
    }

    buildShkPdf = async function (rows, shk) {
        const order = createPlan(rows, shk);
        const source = await PDFLib.PDFDocument.load(shk.buffer, { ignoreEncryption: true });
        if (source.getPageCount() !== shk.pages) throw new Error('Число страниц исходного PDF изменилось');
        const out = await PDFLib.PDFDocument.create();
        const pages = await out.copyPages(source, order);
        pages.forEach(p => out.addPage(p));
        console.info('Ozon PDF: порядок исходных страниц строго по Excel:', order.map(p=>p+1));
        return out.save();
    };

    // Предотвращаем ложное сообщение об успешном PDF, если часть этикеток
    // не сопоставилась уже на этапе формирования строк (включая изображения).
    const baseBuild = startBuild;
    startBuild = async function () {
        if (listData && shkData) {
            // Окончательное соответствие проверится после заполнения ШК старым патчем.
            console.info('Ozon: включена проверка порядка наклеек и Excel');
        }
        return baseBuild();
    };
})();
