/* Расширение для «ЛП Озон с картинками и 1С» v4.
   Работает поверх исходного script.js. Не меняет таблицу, изображения,
   получение данных 1С, оформление Excel и алфавитную сортировку. */
(function () {
    'use strict';
    const captured = Object.create(null);
    document.addEventListener('change', e => {
        if (e.target && e.target.type === 'file' && e.target.files?.length) {
            captured[e.target.id] = e.target.files[0];
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

    async function collectIIFromAssembly(file) {
        const result=new Map();
        const doc=await pdfjsLib.getDocument({data:new Uint8Array(await file.arrayBuffer()),useSystemFonts:true}).promise;
        for(let p=1;p<=doc.numPages;p++) {
            const tc=await (await doc.getPage(p)).getTextContent();
            const words=itemsToWords(tc.items);
            const lines=linesFromText(tc);
            const anchors=words.filter(w=>w.x>=25&&w.x<190&&validPost.test(w.text));
            // На каждой строке Ozon ii напечатан под номером заказа слева.
            for(const a of anchors) {
                const relevant=lines.filter(l=>l.y<=a.y+2&&l.y>=a.y-25&&l.words.some(w=>w.x>=20&&w.x<190))
                  .map(l=>({...l,words:l.words.filter(w=>w.x>=20&&w.x<190)}));
                const id=readII(relevant);
                if(id)result.set(a.text,id);
            }
        }
        return result;
    }

    const originalBuild=startBuild;
    startBuild=async function () {
        // Перестраиваем ключи ДО начала штатной сборки.
        // Этого достаточно и для картинок в Excel, и для PDF: обе ветки
        // исходного приложения используют одну shkData.pageOf.
        if (listData && shkData && captured.docxFileInput) {
            try {
                const iiByPost=await collectIIFromAssembly(captured.docxFileInput);
                const map=new Map();
                for(const [key,pages] of shkData.pageOf)map.set(norm(key),pages.slice());
                let missing=0, duplicate=0;
                const owners=new Map();
                for(const row of listData.rows) {
                    const identifier=iiByPost.get(row.posting)||row.posting;
                    const pages=map.get(norm(identifier));
                    if(!pages?.length){missing++;continue;}
                    for(const pg of pages) {
                        const existing=owners.get(pg);
                        if(existing&&existing!==row.posting)duplicate++;
                        owners.set(pg,row.posting);
                    }
                    // Штатный код сопоставляет по posting — добавляем точный
                    // алиас заказа на страницу с новым ii, не меняя сам номер.
                    shkData.pageOf.set(row.posting,pages.slice());
                }
                if(duplicate)throw new Error('Одна и та же наклейка сопоставлена нескольким заказам: '+duplicate);
                if(missing)console.warn('Ozon: заказов без найденной наклейки:',missing);
            } catch(err) {
                console.error('Ошибка нового сопоставления:',err);
                alert('Не удалось проверить соответствия наклеек: '+err.message+'\nСборка остановлена, чтобы не получить PDF с неправильным порядком.');
                return;
            }
        }
        return originalBuild();
    };
})();
