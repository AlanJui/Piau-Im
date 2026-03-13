
/**
 * 漢字標音切換器 (Phonetic Switcher) - 佈局與邏輯精修版
 * 修正項目：
 * 1. 調整 rtc 間距，解決與漢字重疊的問題。
 * 2. 十五音零聲母自動補上「英」。
 * 3. 嚴格遵循使用者提供的樣式規範。
 */
document.addEventListener('DOMContentLoaded', function() {
    let phoneticMapping = null;

    fetch('./assets/javascripts/phonetic_mapping.json')
        .then(response => response.json())
        .then(data => {
            phoneticMapping = data;
            initSwitcherUI();
        })
        .catch(err => console.error('無法載入標音對照表:', err));

    const schemes = {
        "original": { label: "重設", up: null, right: null },
        "方音符號": { label: "方音符號", up: "", right: "方音符號" },
        "台語音標": { label: "台語音標", up: "台語音標", right: "" },
        "閩拼方案": { label: "閩拼方案", up: "閩拼方案", right: "" },
        "十五音": { label: "十五音", up: "十五音", right: "" },
        "十五音+方音符號": { label: "十五音+方音符號", up: "十五音", right: "方音符號" },
        "台語音標+方音符號": { label: "台語音標+方音符號", up: "台語音標", right: "方音符號" },
        "閩拼方案+方音符號": { label: "閩拼方案+方音符號", up: "閩拼方案", right: "方音符號" },
        "台語音標+十五音": { label: "台語音標+十五音", up: "台語音標", right: "十五音" }
    };

    function initSwitcherUI() {
        const toolbar = document.createElement('div');
        toolbar.className = 'phonetic-switcher-toolbar';
        toolbar.style.cssText = "position:sticky; top:0; z-index:1000; background:#f8f9fa; padding:10px; border-bottom:1px solid #ddd; display:flex; flex-wrap:wrap; gap:10px; align-items:center; justify-content:center;";

        const homeBtn = document.createElement('button');
        homeBtn.innerHTML = "🏠 回首頁";
        homeBtn.style.cssText = "padding:5px 15px; cursor:pointer; font-weight:bold;";
        homeBtn.onclick = () => window.location.href = 'index.html';
        toolbar.appendChild(homeBtn);

        const btnGroup = document.createElement('div');
        btnGroup.style.cssText = "display:flex; flex-wrap:wrap; gap:5px; border-left:1px solid #ccc; padding-left:10px;";
        Object.keys(schemes).forEach(key => {
            const btn = document.createElement('button');
            btn.textContent = schemes[key].label;
            btn.style.padding = "5px 10px"; btn.style.cursor = "pointer";
            btn.onclick = () => applyScheme(key);
            btnGroup.appendChild(btn);
        });
        toolbar.appendChild(btnGroup);

        const customDiv = document.createElement('div');
        customDiv.style.cssText = "border-left:1px solid #ccc; padding-left:10px; display:flex; align-items:center; gap:5px;";
        customDiv.innerHTML = `
            <span style="font-weight:bold">自訂</span>
            上: <select id="select-up"></select>
            右: <select id="select-right"></select>
        `;
        toolbar.appendChild(customDiv);

        const systems = ["無", "十五音", "方音符號", "台語音標", "台羅拼音", "白話字", "閩拼方案", "閩拼調號", "注音二式"];
        const selUp = customDiv.querySelector('#select-up');
        const selRight = customDiv.querySelector('#select-right');
        systems.forEach(s => {
            const val = (s === "無" ? "" : s);
            selUp.add(new Option(s, val));
            selRight.add(new Option(s, val));
        });

        const applyCustom = () => applyPhonetics(selUp.value, selRight.value);
        selUp.onchange = applyCustom; selRight.onchange = applyCustom;

        const main = document.querySelector('main.page');
        if (main) {
            main.insertBefore(toolbar, main.firstChild);
            const footer = toolbar.cloneNode(true);
            footer.style.position = "static"; footer.style.marginTop = "30px";
            footer.querySelector('button').onclick = () => window.location.href = 'index.html';
            footer.querySelectorAll('div button').forEach((btn, idx) => {
                btn.onclick = () => applyScheme(Object.keys(schemes)[idx]);
            });
            const fSelUp = footer.querySelector('#select-up'); const fSelRight = footer.querySelector('#select-right');
            fSelUp.onchange = () => { selUp.value = fSelUp.value; applyCustom(); };
            fSelRight.onchange = () => { selRight.value = fSelRight.value; applyCustom(); };
            main.appendChild(footer);
        }
    }

    function applyScheme(key) {
        if (key === 'original') { location.reload(); return; }
        const s = schemes[key];
        document.getElementById('select-up').value = s.up || "";
        document.getElementById('select-right').value = s.right || "";
        applyPhonetics(s.up, s.right);
    }

    function splitTLPA(tlpa) {
        if (!tlpa) return null;
        let p = tlpa.toLowerCase();
        let tiau = p.slice(-1);
        if (!isNaN(tiau)) { p = p.slice(0, -1); }
        else {
            if (['p', 't', 'k', 'h'].includes(tiau)) tiau = '4';
            else tiau = '1';
        }
        let siann = "", un = "";
        const initials = ["tsh", "kh", "ph", "th", "ng", "bb", "gg", "ch", "zh", "ji", "ts", "z", "c", "p", "k", "t", "l", "s", "h", "b", "g", "m", "n", "j"];
        initials.sort((a, b) => b.length - a.length);
        for (let i of initials) { if (p.startsWith(i)) { siann = i; un = p.slice(i.length); break; } }
        if (!siann) un = p;
        return { siann: siann || "ø", un: un, tiau: tiau };
    }

    function decodeUnicode(str) {
        if (!str) return "";
        return str.replace(/U\+([0-9A-Fa-f]{4})/g, (match, hex) => String.fromCharCode(parseInt(hex, 16)));
    }

    function applyBPToneMark(pingIm, toneMark) {
        if (!toneMark) return pingIm;
        let cleanIm = pingIm.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
        let targetIdx = -1;
        if (cleanIm.includes('iu')) targetIdx = cleanIm.indexOf('u');
        else if (cleanIm.includes('ui')) targetIdx = cleanIm.indexOf('i');
        else {
            const priority = ['a', 'o', 'e', 'i', 'u', 'v'];
            for (let v of priority) { if (cleanIm.includes(v)) { targetIdx = cleanIm.indexOf(v); break; } }
        }
        if (targetIdx === -1) return cleanIm + toneMark;
        return cleanIm.slice(0, targetIdx + 1) + toneMark + cleanIm.slice(targetIdx + 1);
    }

    function convertOne(tlpa, targetSystem) {
        if (!targetSystem) return "";
        const parts = splitTLPA(tlpa); if (!parts) return "";

        let initialMatch = phoneticMapping.initials.find(i => 
            i.台語音標 === parts.siann || (parts.siann === "ø" && (i.台語音標 === "" || i.台語音標 === "Ø"))
        );
        let finalMatch = phoneticMapping.finals.find(f => f.台語音標 === parts.un);
        let toneMatch = phoneticMapping.tones.find(t => String(t.台羅調號) === parts.tiau || String(t.台語音標).endsWith(parts.tiau));

        if (targetSystem === '十五音') {
            let iName = initialMatch ? initialMatch['十五音'] : "";
            // 修正：零聲母補上「英」
            if (!iName && parts.siann === "ø") iName = "英";
            const fName = finalMatch ? finalMatch['十五音'] : "";
            const toneCN = ["", "一", "二", "三", "四", "五", "六", "七", "八"][parseInt(parts.tiau)] || parts.tiau;
            return fName + toneCN + iName;
        }

        if (targetSystem === '方音符號') {
            let iTPS = initialMatch ? initialMatch['方音符號'] : "";
            let fTPS = finalMatch ? finalMatch['方音符號'] : "";
            let tTPS = decodeUnicode((toneMatch && toneMatch['注音調符編碼'])) || "";
            if (parts.un.startsWith('i')) {
                if (parts.siann === 'z' || parts.siann === 'ts') iTPS = 'ㄐ';
                else if (parts.siann === 'c' || parts.siann === 'tsh') iTPS = 'ㄑ';
                else if (parts.siann === 's') iTPS = 'ㄒ';
                else if (parts.siann === 'j' || parts.siann === 'ji') iTPS = 'ㆢ';
            }
            return (iTPS === "Ø" ? "" : iTPS) + fTPS + tTPS;
        }

        let sysKey = (targetSystem === "閩拼方案" || targetSystem === "閩拼調號") ? "閩拼方案" : targetSystem;
        let iPin = initialMatch ? (initialMatch[sysKey] || initialMatch['台羅拼音'] || parts.siann) : parts.siann;
        let fPin = finalMatch ? (finalMatch[sysKey] || finalMatch['台羅拼音'] || parts.un) : parts.un;
        if (iPin === "Ø" || iPin === "ø") iPin = "";
        
        if (targetSystem === "閩拼方案") {
            let mark = (toneMatch && toneMatch['拼音調符編碼']) || "";
            if (!mark && toneMatch && toneMatch['羅馬拼音調符']) { mark = toneMatch['羅馬拼音調符'].replace(/[a-zA-Z◌]/g, '').trim(); }
            return applyBPToneMark(iPin + fPin, decodeUnicode(mark));
        }
        return iPin + fPin + parts.tiau;
    }

    function applyPhonetics(upSystem, rightSystem) {
        document.querySelectorAll('article.article_content > div').forEach(div => {
            div.style.cssText = "display: block; line-height: 2.2; font-family: 'Noto Serif TC', serif; font-size: 2.8rem; text-align: justify;";
            div.classList.remove('Siang_Pai', 'pin_yin', 'Piau_Im', 'fifteen_yin');
        });

        document.querySelectorAll('ruby[data-tlpa]').forEach(ruby => {
            const tlpa = ruby.getAttribute('data-tlpa');
            let hanJi = "";
            for (let node of ruby.childNodes) {
                if (node.nodeType === 3) { hanJi = node.textContent.trim(); break; }
            }
            if (!hanJi && ruby.innerText) hanJi = ruby.innerText.split('\n')[0].trim();

            ruby.innerHTML = hanJi;
            // 修正：增加 ruby 間距，解決注音符號與下一個字太擠的問題
            ruby.style.cssText = "display: inline-ruby !important; ruby-position: over; margin-right: 0.8em; margin-bottom: 0.5em;";

            if (upSystem) {
                const rt = document.createElement('rt');
                rt.textContent = convertOne(tlpa, upSystem);
                rt.style.cssText = `
                    font-family: Arial, "Noto Sans TC", sans-serif;
                    color: #00c4ab; font-size: 50%; margin-left: -0.4em;
                    font-variant-east-asian: salt; -webkit-font-feature-settings: "salt"; font-feature-settings: "salt";
                    white-space: nowrap;
                `;
                ruby.appendChild(rt);
            }

            if (rightSystem) {
                const rtc = document.createElement('rtc');
                rtc.textContent = convertOne(tlpa, rightSystem);
                rtc.style.cssText = `
                    font-family: BopomofoRuby; color: #B94B6A; font-size: 40%; font-weight: bold; line-height: 1;
                    ruby-position: inter-character; -webkit-writing-mode: vertical-lr; writing-mode: vertical-lr;
                    vertical-align: middle;
                    margin-left: -0.1em; /* 修正：進一步拉開與漢字的距離，解決重疊問題 */
                    margin-right: 0.2em;
                `;
                ruby.appendChild(rtc);
            }
        });
    }
});
