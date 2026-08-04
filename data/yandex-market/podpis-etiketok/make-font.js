const fs = require('fs');
const path = require('path');
const https = require('https');

// ---------- вспомогательные ----------
function httpsGet(url) {
    return new Promise((resolve, reject) => {
        https.get(url, (res) => {
            if (res.statusCode !== 200) {
                res.resume();
                return reject(new Error('HTTP ' + res.statusCode));
            }
            let data = '';
            res.on('data', (c) => (data += c));
            res.on('end', () => resolve(data));
        }).on('error', reject);
    });
}

function wrapCjs(src) {
    return '(function(){var module={exports:{}};var exports=module.exports;\n' + src + '\n;window.fontkit=module.exports;})();\n';
}

function prepareFontkitSource(src) {
    // UMD-сборка работает в браузере как есть; CJS нужно обернуть
    if (/typeof\s+(self|window|globalThis)/.test(src) || /\bdefine\(/.test(src)) return src;
    if (/module\.exports/.test(src)) return wrapCjs(src);
    return src;
}

// ---------- 1) шрифт ----------
try {
    const ttfPath = path.join(__dirname, 'DejaVuSans.ttf');
    const outFont = path.join(__dirname, 'font-data.js');
    const b64 = fs.readFileSync(ttfPath).toString('base64');
    fs.writeFileSync(outFont, 'window.DEJAVU_FONT_B64 = "' + b64 + '";\n');
    console.log('✅ font-data.js создан (' + (fs.statSync(outFont).size / 1024).toFixed(0) + ' КБ)');
} catch (e) {
    console.log('⚠️ Ошибка упаковки шрифта: ' + e.message);
}

// ---------- 2) fontkit ----------
(async () => {
    const outFk = path.join(__dirname, 'fontkit.local.js');
    let src = null;

    // Попытка 1: взять из node_modules старого серверного проекта
    try {
        const oldProject = path.join(__dirname, '..', '..', '..', '..', 'marketplace-label-processor');
        const resolved = require.resolve('@pdf-lib/fontkit', { paths: [oldProject] });
        src = fs.readFileSync(resolved, 'utf8');
        console.log('✅ fontkit найден локально: ' + resolved);
    } catch (e) { /* идём дальше */ }

    // Попытка 2: узнать реальный список файлов пакета и скачать нужный
    if (!src) {
        try {
            const meta = JSON.parse(await httpsGet('https://data.jsdelivr.com/v1/package/npm/@pdf-lib/fontkit@1.1.2'));
            const files = [];
            (function walk(nodes) {
                for (const n of nodes || []) {
                    if (n.type === 'file') files.push(n.name);
                    if (n.files) walk(n.files);
                }
            })(meta.files);
            const pick =
                files.find((f) => /umd(\.min)?\.js$/.test(f)) ||
                files.find((f) => /^\/dist\/.+\.js$/.test(f)) ||
                files.find((f) => /\.js$/.test(f));
            if (pick) {
                console.log('Скачиваю fontkit: ' + pick);
                src = await httpsGet('https://cdn.jsdelivr.net/npm/@pdf-lib/fontkit@1.1.2' + pick);
            }
        } catch (e) {
            console.log('⚠️ jsdelivr API: ' + e.message);
        }
    }

    // Попытка 3: прямые адреса
    if (!src) {
        const urls = [
            'https://cdn.jsdelivr.net/npm/@pdf-lib/fontkit@1.1.2/dist/fontkit.cjs.js',
            'https://unpkg.com/@pdf-lib/fontkit@1.1.2/dist/fontkit.cjs.js'
        ];
        for (const u of urls) {
            try {
                src = await httpsGet(u);
                console.log('Скачиваю fontkit: ' + u);
                break;
            } catch (e) { /* следующий */ }
        }
    }

    if (!src) {
        console.log('⚠️ Не удалось получить fontkit. Для этого шага нужен интернет (или старый проект с node_modules).');
        return;
    }

    fs.writeFileSync(outFk, prepareFontkitSource(src));
    console.log('✅ fontkit.local.js создан (' + (fs.statSync(outFk).size / 1024).toFixed(0) + ' КБ)');
})();