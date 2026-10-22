/* Номер ШК вместо картинки этикетки в готовом Excel.
   Остальные функции и исходная сортировка остаются без изменений.
   Для старой этикетки: первые 8 цифр, из них выделяем последние 4 перед ПЕРВЫМ дефисом.
   Для новой: ii... и последние четыре цифры выделяем.
*/
(function () {
    'use strict';
    const original = buildLpXlsx;
    const xmlEscape = s => String(s == null ? '' : s).replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[ch]));
    const columnLetter = index => {
        let s = '';
        for (let n = index + 1; n; n = Math.floor((n-1)/26)) s = String.fromCharCode(65+(n-1)%26)+s;
        return s;
    };
    const numberToDisplay = (r, iiMap) => String(iiMap.get(r.posting) || r.posting || '').replace(/\s+/g, '');
    function parts(num) {
        if (/^ii\d{8,16}$/i.test(num)) return [num.slice(0,-4),num.slice(-4),''];
        const m = /^(\d{4,6})(\d{4})(-\d{3,5}-\d{1,3})$/.exec(num);
        if (m) return [m[1],m[2],m[3]];
        return [num,'',''];
    }
    const basic = s => '<r><rPr><sz val="11"/><rFont val="Calibri"/></rPr><t xml:space="preserve">'+xmlEscape(s)+'</t></r>';
    const highlight = s => '<r><rPr><b/><sz val="14"/><color rgb="FF005BFF"/><rFont val="Calibri"/></rPr><t xml:space="preserve">'+xmlEscape(s)+'</t></r>';
    function richCell(ref, value, style) {
        const [a,b,c] = parts(value);
        const runs = b ? basic(a)+highlight(b)+(c?basic(c):'') : basic(a);
        return '<c r="'+ref+'" s="'+style+'" t="inlineStr"><is>'+runs+'</is></c>';
    }

    // Для Excel картинки этикеток больше не нужны: читаем только идентификаторы.
    // При этом сохраняем страницы PDF для старой логики сопоставления и экспорта.
    async function readLabelsWithoutImages(file) {
        document.getElementById('shkFileName').textContent=file.name;
        setStatus('shkStatus','⏳ Чтение номеров этикеток...','');
        try {
            const buf=new Uint8Array(await file.arrayBuffer());
            const keep=new Uint8Array(buf);
            const doc=await pdfjsLib.getDocument({data:buf,disableFontFace:true}).promise;
            const pageOf=new Map(),pageToNum=new Map(),ordered=[];
            let noNum=0;
            for(let p=1;p<=doc.numPages;p++) {
                const tc=await (await doc.getPage(p)).getTextContent();
                const num=postingFromPage(tc);
                // Лёгкая метка: означает, что страница существует. PNG не создаём.
                ordered.push({bytes:new Uint8Array(0),w:1,h:1});
                if(num){
                    if(!pageOf.has(num))pageOf.set(num,[]);
                    pageOf.get(num).push(p);
                    pageToNum.set(p,num);
                } else noNum++;
                if(p%10===0)await new Promise(resolve=>setTimeout(resolve,0));
            }
            shkData={ordered,pageOf,pageToNum,pages:doc.numPages,buffer:keep,name:file.name};
            setStatus('shkStatus','✅ Этикеток: '+doc.numPages+' (номеров: '+(doc.numPages-noNum)+(noNum?', без номера: '+noNum:'')+')','success');
        } catch(err){
            shkData=null;
            setStatus('shkStatus','❌ Ошибка: '+err.message,'error');
            console.error(err);
        }
        updateStartBtn();
    }
    // Заменяем только обработчик загрузки наклеек. Другие загрузчики не трогаем.
    const oldArea=document.getElementById('shkUploadArea');
    const newArea=oldArea.cloneNode(true);
    oldArea.parentNode.replaceChild(newArea,oldArea);
    const newInput=document.getElementById('shkFileInput');
    newInput.addEventListener('change',e=>{
        if(e.target.files[0])readLabelsWithoutImages(e.target.files[0]);
        newInput.value='';
    });
    newArea.addEventListener('dragover',e=>{e.preventDefault();newArea.classList.add('dragover');});
    newArea.addEventListener('dragleave',()=>newArea.classList.remove('dragover'));
    newArea.addEventListener('drop',e=>{
        e.preventDefault();newArea.classList.remove('dragover');
        if(e.dataTransfer.files[0])readLabelsWithoutImages(e.dataTransfer.files[0]);
    });

    buildLpXlsx = async function(rows) {
        // Оригинальная генерация сохраняет фото товара и все данные.
        // Лишь картинки ЭТИКЕТОК временно исключаем из xlsx.
        const labels = rows.map(r => r.label);
        const iiMap = new Map((listData && listData.rows || []).map(r => [r.posting,r.shkId || '']));
        let result;
        try {
            rows.forEach(r => { r.label = null; });
            result = await original(rows);
        } finally {
            rows.forEach((r,i) => { r.label = labels[i]; });
        }
        const zip = await JSZip.loadAsync(result);
        let sheet = await zip.file('xl/worksheets/sheet1.xml').async('string');
        // Находим последние две колонки по исходной шапке (Кол-во может отсутствовать).
        const headers = [...sheet.matchAll(/<c\b[^>]*\br="([A-Z]+)1"[^>]*>[\s\S]*?<\/c>/g)];
        const labelHead = headers.find(m => m[0].includes('Этикетка'));
        const stickerHead = headers.find(m => m[0].includes('Стикер'));
        if (!labelHead || !stickerHead) throw new Error('Не удалось обнаружить колонки «Стикер» и «Этикетка»');
        const labelCol=labelHead[1], stickerCol=stickerHead[1];
        const lastIndex = str => [...str].reduce((n,ch)=>n*26+ch.charCodeAt(0)-64,0);
        const labelIndex=lastIndex(labelCol);
        const stickerIndex=lastIndex(stickerCol);
        if(labelIndex !== stickerIndex+1) throw new Error('Неожиданный порядок колонок: изменение Excel остановлено');
        // Удаляем пустую колонку «Этикетка» (включая заголовок и ширину).
        const escaped = labelCol.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
        sheet=sheet.replace(new RegExp('<c\\b[^>]*\\br="'+escaped+'\\d+"[^>]*?(?:\\/>|>[\\s\\S]*?<\\/c>)','g'),'');
        sheet=sheet.replace(new RegExp('<col\\b(?=[^>]*\\bmin="'+labelIndex+'")(?=[^>]*\\bmax="'+labelIndex+'")[^>]*/>','g'),'');
        sheet=sheet.replace(/(<dimension ref="[A-Z]+1:)[A-Z]+(\d+"\/>)/,'$1'+stickerCol+'$2');
        sheet=sheet.replace(/(<autoFilter ref="[A-Z]+1:)[A-Z]+(\d+"\/>)/,'$1'+stickerCol+'$2');
        sheet=sheet.replace('>Стикер</t>', '>Номер ШК</t>');
        rows.forEach((row,i) => {
            const ref=stickerCol+(i+2);
            const escapedRef=ref.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
            const re=new RegExp('<c\\b[^>]*\\br="'+escapedRef+'"[^>]*?(?:\\/>|>[\\s\\S]*?<\\/c>)');
            const matched=re.exec(sheet);
            if(!matched)throw new Error('Не найдена ячейка '+ref);
            const style=/<c\b[^>]*\bs="(\d+)"/.exec(matched[0])?.[1] || '3';
            sheet=sheet.replace(re,richCell(ref,numberToDisplay(row,iiMap),style));
        });
        // Колонка теперь содержит полный номер ШК, а не 4 цифры.
        sheet=sheet.replace(new RegExp('<col\\b(?=[^>]*\\bmin="'+stickerIndex+'")(?=[^>]*\\bmax="'+stickerIndex+'")[^>]*/>'),
          m => m.replace(/width="[^"]+"/,'width="25"'));
        zip.file('xl/worksheets/sheet1.xml',sheet);
        return zip.generateAsync({type:'blob',compression:'DEFLATE',compressionOptions:{level:6}});
    };
})();
