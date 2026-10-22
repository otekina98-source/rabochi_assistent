/* Расширение для «ЛП Озон с картинками и 1С» v4.
   Работает поверх исходного script.js. Не меняет таблицу, изображения,
   получение данных 1С, оформление Excel и алфавитную сортировку. */
(function () {
    'use strict';
    const captured = Object.create(null);
    // Сохраняем ссылку на файл для обеих поддерживаемых загрузок:
    // «Выбрать PDF» И перетаскивание в область загрузки.
    document.addEventListener('change', e => {
        if (e.target?.id === 'docxFileInput' && e.target.files?.[0]) {
            captured.docxFileInput = e.target.files[0];
        }
    }, true);
    document.addEventListener('drop', e => {
        if (e.target?.closest?.('#docxUploadArea') && e.dataTransfer?.files?.[0]) {
            captured.docxFileInput = e.dataTransfer.files[0];
        }
    }, true);
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://misha-spec.github.io/sistema_bystrogo_scheta_po_trakhtenbergu/assets/vendor/pdf.worker.min.js';

    const norm = s => String(s || '').replace(/[\s\u00a0\u2007\u202f]/g, '').toLowerCase();
    const validPost = /^\d{5,10}-\d{3,5}-\d{1,3}$/;
    const validII = /^ii\d{8,16}$/;

    function linesFromText(tc) {
        const words = itemsToWords(tc.items).filter(w => w.text && w.text.trim());
        const lines = [];
        for (const w of words.sort((a,b) => b.y-a.y || a.x-b.x)) {
            let line = lines.find(l => Math.abs(l.y-w.y) < 2.5);
            if (!line) { line = {y:w.y, words:[]}; lines.push(line); }
            line.words.push(w);
        }
        return lines.sort((a,b) => b.y-a.y).map(l => ({
            y:l.y, words:l.words.sort((a,b) => a.x-b.x),
            get text() { return this.words.map(w=>w.text).join(' '); },
            get x() { return Math.min(...this.words.map(w=>w.x)); }
        }));
    }
    function readII(lines) {
        for (const line of lines) {
            const m=/(?:^|[^a-z0-9])(ii\s*(?:\d\s*){8,16})(?!\d)/i.exec(line.text);
            if (m && validII.test(norm(m[1]))) return norm(m[1]);
        }
        return '';
    }
    function fromStickerPage(tc) {
        // В реальном ticket.pdf две части одного ШК имеют РАЗНЫЕ базовые линии:
        // "ii5005890" + " 3451". Не требуем совпадения y для двух соседних
        // текстовых элементов, но используем целый ID, а не 4 последние цифры.
        const newChunks=tc.items.map(it => String(it.str || '').trim()).filter(Boolean);
        const iiMatches=new Set();
        for(let i=0;i<newChunks.length;i++) {
            const first=norm(newChunks[i]);
            if(!/^ii\d{4,12}$/.test(first) && !validII.test(first))continue;
            if(validII.test(first))iiMatches.add(first);
            // Суффикс находится непосредственно после префикса в порядке
            // объектов PDF. Цифры адреса забирать нельзя: только отдельный
            // соседний фрагмент и итоговая проверка полного номера.
            if(i+1<newChunks.length && /^\d{4}$/.test(norm(newChunks[i+1]))) {
                const full=first+norm(newChunks[i+1]);
                if(validII.test(full))iiMatches.add(full);
            }
        }
        // Ещё один вариант: pdf.js отдаёт номер одной печатной строкой.
        const direct=(tc.items.map(it=>String(it.str||'')).join(' ')).match(/(?:^|[^a-z0-9])(ii\s*(?:\d\s*){8,16})(?!\d)/ig)||[];
        for(const raw of direct) {
            const id=norm(raw).replace(/^[^i]+/,'');
            if(validII.test(id))iiMatches.add(id);
        }
        if(iiMatches.size===1)return [...iiMatches][0];
        if(iiMatches.size>1)return null;
        const lines=linesFromText(tc);
        const ii=readII(lines);
        if(ii)return ii;
        // Часть PDF хранит два цифровых блока номера в разных координатах
        // и даже в неожиданном порядке объектов; не используем цифры адреса.
        const chunks=tc.items.map(it=>String(it.str||'').trim());
        const suffixItem=chunks.findIndex(t=>/^\s*-\d{3,5}\s*-\d{1,3}\s*$/.test(t));
        if(suffixItem>=0){
            const tail=/-\s*(\d{3,5})\s*-\s*(\d{1,3})/.exec(chunks[suffixItem]);
            const pieces=chunks.slice(0,suffixItem).filter(t=>/^\d{3,10}$/.test(t));
            let head='';
            if(pieces.length>=2){
                const combined=pieces.slice(-2).join('');
                if(combined.length>=8 && combined.length<=10)head=combined;
            }
            if(!head && pieces.length && pieces[pieces.length-1].length>=8)head=pieces[pieces.length-1];
            if(head)return head+'-'+tail[1]+'-'+tail[2];
        }
        // Для старых наклеек номер состоит из двух визуальных строк.
        // Изучаем непосредственно цифровую строку над суффиксом, не адрес.
        for(let i=0;i<lines.length;i++) {
            const lower=lines[i];
            const mt=/-\s*(\d{3,5})\s*-\s*(\d{1,3})(?!\d)/.exec(lower.text);
            if(!mt)continue;
            const prefix=lower.text.slice(0,mt.index).replace(/\s+/g,'');
            if(/^\d{5,10}$/.test(prefix))return prefix+'-'+mt[1]+'-'+mt[2];
            for(let j=i-1;j>=Math.max(0,i-2);j--) {
                if(Math.abs(lines[j].y-lower.y)>35)continue;
                const head=lines[j].text.replace(/\s+/g,'');
                if(/^\d{5,10}$/.test(head))return head+'-'+mt[1]+'-'+mt[2];
            }
        }
        // Страница с готовым полным номером.
        for(const item of tc.items){
            const m=/(?:^|\D)(\d{5,10}-\d{3,5}-\d{1,3})(?!\d)/.exec(norm(item.str));
            if(m&&validPost.test(m[1]))return m[1];
        }
        return null;
    }
    // Исходный обработчик этикеток использует именно postingFromPage.
    postingFromPage = fromStickerPage;

    // Привязка ШК строго к строке PDF: номер в левой колонке,
    // ii расположен под ним, до начала следующего отправления.
    async function collectIIFromAssembly(file) {
        const result = new Map();
        const doc = await pdfjsLib.getDocument({
            data: new Uint8Array(await file.arrayBuffer()), useSystemFonts: true
        }).promise;
        for (let p=1;p<=doc.numPages;p++) {
            const tc=await (await doc.getPage(p)).getTextContent();
            const words=itemsToWords(tc.items).filter(w => w.text?.trim());
            const anchors=[];
            for(const w of words) {
                if(w.x<25 || w.x>=190)continue;
                if(validPost.test(w.text))anchors.push({...w, posting:w.text});
                else {
                    const m=/^(\d{1,4})(\d{8,10}-\d{3,5}-\d{1,3})$/.exec(w.text);
                    if(m)anchors.push({...w,posting:m[2]});
                }
            }
            anchors.sort((a,b)=>b.y-a.y);
            for(let i=0;i<anchors.length;i++) {
                const a=anchors[i];
                const below= i+1<anchors.length ? (a.y+anchors[i+1].y)/2 : -Infinity;
                const inRow=words.filter(w=>w.x>=25 && w.x<190 && w.y<=a.y+3 && w.y>below);
                const ids=[];
                // PDF сохраняет серый номер цельным либо отдельными частями.
                for(const line of linesFromText({items:tc.items.filter(it=>
                    it.transform[4]>=25 && it.transform[4]<190 &&
                    it.transform[5]<=a.y+3 && it.transform[5]>below
                )})) {
                    const m=/(?:^|[^a-z0-9])(ii\s*(?:\d\s*){8,16})(?!\d)/i.exec(line.text);
                    if(m && validII.test(norm(m[1]))) ids.push(norm(m[1]));
                }
                // Иногда pdf.js ставит разные базовые линии внутри одного ШК.
                if(!ids.length) {
                    const sorted=inRow.sort((x,y)=>y.y-x.y||x.x-y.x);
                    for(const w of sorted) {
                        const id=norm(w.text);
                        if(validII.test(id))ids.push(id);
                    }
                }
                const unique=[...new Set(ids)];
                if(unique.length>1)throw new Error('Несколько номеров ii для '+a.posting);
                if(unique.length===1)result.set(a.posting,unique[0]);
            }
        }
        return result;
    }

    const originalBuild=startBuild;
    startBuild=async function () {
        if(!(listData && shkData))return originalBuild();
        // Загруженный файл запоминается при выборе и при перетаскивании.
        // Не заменяем его по имени другим PDF и не используем пустой input
        // (оригинальный интерфейс очищает input.value после загрузки).
        if (!captured.docxFileInput) {
            alert('Исходный PDF не сохранился в памяти страницы. Перезагрузите страницу и загрузите лист сборки снова.');
            return;
        }
        try {
            const iiByPost=await collectIIFromAssembly(captured.docxFileInput);
            const owners=new Map();
            const fullMatch=new Map();
            for(const row of listData.rows) {
                row.shkId=iiByPost.get(row.posting)||'';
                const key=norm(row.shkId||row.posting);
                const pages=shkData.pageOf.get(key)||[];
                if(pages.length>1)throw new Error('Несколько наклеек с одинаковым номером '+key);
                if(pages.length===1) {
                    if(owners.has(pages[0]))throw new Error('Одна наклейка соответствует двум заказам');
                    owners.set(pages[0],row.posting);
                    fullMatch.set(row.posting,pages.slice());
                }
            }
            // Удаляем псевдонимы от предыдущей сборки, чтобы повторное нажатие
            // кнопки не оставляло устаревших совпадений.
            for(const row of listData.rows) {
                if(row.shkId)shkData.pageOf.delete(row.posting);
            }
            for(const [posting,pages] of fullMatch)shkData.pageOf.set(posting,pages);
            const newRows=listData.rows.filter(r=>r.shkId);
            console.info('Новые ШК ii:',newRows.map(r=>r.posting+' → '+r.shkId));
            const missing=listData.rows.filter(r=>!fullMatch.has(r.posting));
            if(missing.length)console.warn('Наклейки не найдены:',missing.map(r=>r.shkId||r.posting));
        } catch(err) {
            console.error(err);
            alert('Ошибка сопоставления ШК: '+err.message+'\nСборка остановлена.');
            return;
        }
        return originalBuild();
    };
})();
