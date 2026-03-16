
/**
 * 漢字標音切換器 (Phonetic Switcher) - 聲韻學強修版
 * 1. 新增【國際音標 (IPA)】轉換功能。
 * 2. 羅馬拼音調符全效支援 (TL, POJ, BP)。
 * 3. 啟動即修復拉丁調符偏移問題。
 */
document.addEventListener('DOMContentLoaded', function() {
    let phoneticMapping = null;

    // --- [1. 樣式修復] ---
    const injectLatinFix = () => {
        if (!document.getElementById('latin-phonetic-fix')) {
            const style = document.createElement('style');
            style.id = 'latin-phonetic-fix';
            style.textContent = `
                rt.latin-phonetic, rtc.latin-phonetic { 
                    font-family: Arial, Helvetica, sans-serif !important; 
                    font-feature-settings: "normal" !important; 
                    font-variant-east-asian: normal !important; 
                    text-align: center !important;
                    letter-spacing: -0.01em !important;
                }
            `;
            document.head.appendChild(style);
        }
    };

    const autoFix = () => {
        const latinRegex = /[a-zA-Z\u0300-\u036f]/;
        document.querySelectorAll('rt, rtc').forEach(el => {
            if (latinRegex.test(el.textContent)) el.classList.add('latin-phonetic');
        });
    };

    injectLatinFix();
    autoFix();

    // --- [2. 資料載入] ---
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
        "台羅拼音": { label: "台羅拼音", up: "台羅拼音", right: "" },
        "白話字": { label: "白話字", up: "白話字", right: "" },
        "閩拼方案": { label: "閩拼方案", up: "閩拼方案", right: "" },
        "十五音+方音符號": { label: "十五音+方音符號", up: "十五音", right: "方音符號" },
        "台語音標+十五音": { label: "台語音標+十五音", up: "台語音標", right: "十五音" }
    };

    function initSwitcherUI() {
        const toolbar = document.createElement('div');
        toolbar.className = 'phonetic-switcher-toolbar';
        toolbar.style.cssText = "position:sticky; top:0; z-index:1000; background:#f8f9fa; padding:8px 15px; border-bottom:1px solid #ddd; display:flex; flex-wrap:nowrap; gap:8px; align-items:center; overflow-x:auto; white-space:nowrap;";

        const homeBtn = document.createElement('button');
        homeBtn.innerHTML = "🏠 回首頁";
        homeBtn.style.cssText = "padding:4px 10px; cursor:pointer; font-weight:bold; flex-shrink:0;";
        homeBtn.onclick = () => window.location.href = 'index.html';
        toolbar.appendChild(homeBtn);

        const btnGroup = document.createElement('div');
        btnGroup.style.cssText = "display:flex; gap:4px; border-left:1px solid #ccc; padding-left:8px; flex-shrink:0;";
        Object.keys(schemes).forEach(key => {
            const btn = document.createElement('button');
            btn.textContent = schemes[key].label;
            btn.style.cssText = "padding:4px 8px; cursor:pointer; font-size:13px; flex-shrink:0;";
            btn.onclick = () => applyScheme(key);
            btnGroup.appendChild(btn);
        });
        toolbar.appendChild(btnGroup);

        const customDiv = document.createElement('div');
        customDiv.style.cssText = "border-left:1px solid #ccc; padding-left:8px; display:flex; align-items:center; gap:5px; flex-shrink:0; font-size:13px;";
        customDiv.innerHTML = `
            <span style="font-weight:bold">自訂</span>
            上:<select id="select-up" style="font-size:12px"></select>
            右:<select id="select-right" style="font-size:12px"></select>
        `;
        toolbar.appendChild(customDiv);

        // --- 修正：在選單中加入【國際音標】 ---
        const systems = ["無", "十五音", "方音符號", "國際音標", "台語音標", "台羅拼音", "白話字", "閩拼方案", "閩拼調號", "注音二式"];
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
            footer.querySelectorAll('div button').forEach((btn, idx) => { btn.onclick = () => applyScheme(Object.keys(schemes)[idx]); });
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
            if (['p', 't', 'k', 'h'].includes(tiau)) tiau = '4'; else tiau = '1';
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

    function applyToneMark(base, tiau, system) {
        const marks = {
            "TL":  { "2":"\u0301", "3":"\u0300", "5":"\u0302", "6":"\u030c", "7":"\u0304", "8":"\u030d" },
            "POJ": { "2":"\u0301", "3":"\u0300", "5":"\u0302", "7":"\u0304", "8":"\u030d" },
            "BP":  { "1":"\u0304", "5":"\u0301", "2":"\u030c", "6":"\u030c", "3":"\u0300", "7":"\u0302", "4":"\u0304", "8":"\u0301" }
        };
        let sys = system === "白話字" ? "POJ" : (system === "台羅拼音" ? "TL" : "BP");
        let mark = marks[sys][tiau] || "";
        if (!mark) return base;

        let s = base.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
        let pos = -1;
        if (s.includes("ere")) pos = s.indexOf("ere") + 2;
        else if (s.includes("iu")) pos = s.indexOf("u");
        else if (s.includes("ui")) pos = s.indexOf("i");
        else if (s.includes("oo")) pos = s.indexOf("o");
        else if (s.includes("ng") && !/[aeiou]/.test(s)) pos = s.indexOf("n");
        else if (s.includes("m") && !/[aeiou]/.test(s)) pos = s.indexOf("m");
        else {
            const priority = ['a', 'o', 'e', 'i', 'u', 'v'];
            for (let v of priority) { if (s.includes(v)) { pos = s.indexOf(v); break; } }
        }
        if (pos === -1) return s + mark;
        let res = s.slice(0, pos + 1) + mark + s.slice(pos + 1);
        if (system === "白話字" && base.includes("o\u0358")) res = res.replace("o", "o\u0358");
        return res.normalize("NFC");
    }

    function convertOne(tlpa, targetSystem) {
        if (!targetSystem) return "";
        const parts = splitTLPA(tlpa); if (!parts) return "";

        let initialMatch = phoneticMapping.initials.find(i => i.台語音標 === parts.siann || (parts.siann === "ø" && (i.台語音標 === "" || i.台語音標 === "Ø" || i.台語音標 === "ø")));
        let finalMatch = phoneticMapping.finals.find(f => f.台語音標 === parts.un);

        if (targetSystem === '十五音') {
            let iName = initialMatch ? initialMatch['十五音'] : (parts.siann === "ø" ? "英" : "");
            const fName = finalMatch ? finalMatch['十五音'] : "";
            const toneCN = ["", "一", "二", "三", "四", "五", "六", "七", "八"][parseInt(parts.tiau)] || parts.tiau;
            return fName + toneCN + iName;
        }

        if (targetSystem === '方音符號') {
            let iTPS = initialMatch ? initialMatch['方音符號'] : "";
            let fTPS = finalMatch ? finalMatch['方音符號'] : "";
            let toneMatch = phoneticMapping.tones.find(t => String(t.台羅調號) === parts.tiau);
            let tTPS = decodeUnicode((toneMatch && toneMatch['注音調符編碼'])) || "";
            if (tTPS.startsWith('U+')) tTPS = String.fromCharCode(parseInt(tTPS.substring(2), 16));
            if (parts.un.startsWith('i') && ['z','ts','c','tsh','s','j','ji'].includes(parts.siann)) {
                if (parts.siann === 's') iTPS = 'ㄒ'; else if (parts.siann === 'j' || parts.siann === 'ji') iTPS = 'ㆢ';
                else if (parts.siann === 'z' || parts.siann === 'ts') iTPS = 'ㄐ'; else iTPS = 'ㄑ';
            }
            return (iTPS === "Ø" ? "" : iTPS) + fTPS + tTPS;
        }

        // --- 修正：處理國際音標 (IPA) ---
        if (targetSystem === '國際音標') {
            let iIPA = initialMatch ? initialMatch['國際音標'] : (parts.siann === "ø" ? "" : "");
            let fIPA = finalMatch ? finalMatch['國際音標'] : parts.un;
            // 處理 Ø 顯示
            if (iIPA === "Ø" || iIPA === "ø") iIPA = "";
            return iIPA + fIPA + parts.tiau; // 直接帶數字調號
        }

        let bSiann = parts.siann === "ø" ? "" : parts.siann;
        let bUn = parts.un;

        if (targetSystem === "白話字") {
            if (bSiann === "ts" || bSiann === "z") bSiann = "ch"; else if (bSiann === "tsh" || bSiann === "c") bSiann = "chh";
            bUn = bUn.replace(/ue/g, "oe").replace(/ua/g, "oa").replace(/ik/g, "ek").replace(/ing/g, "eng").replace(/oo/g, "o\u0358").replace(/nn/g, "\u207F");
            return applyToneMark(bSiann + bUn, parts.tiau, "白話字");
        }

        if (targetSystem === "台羅拼音") {
            if (bSiann === "z") bSiann = "ts"; if (bSiann === "c") bSiann = "tsh";
            return applyToneMark(bSiann + bUn, parts.tiau, "台羅拼音");
        }

        if (targetSystem === "閩拼方案" || targetSystem === "閩拼調號") {
            if (bSiann === "") {
                if (bUn.startsWith('i')) { if (bUn === 'i' || /^(in|im|ing|it|ip|ik|ih)/.test(bUn)) bSiann = "y"; else { bSiann = "y"; bUn = bUn.substring(1); } }
                else if (bUn.startsWith('u')) { if (bUn === 'u' || /^(un|ut|uh)/.test(bUn)) bSiann = "w"; else { bSiann = "w"; bUn = bUn.substring(1); } }
            }
            if (targetSystem === "閩拼調號") {
                const numMap = {"1":"1", "5":"2", "2":"3", "6":"4", "3":"5", "7":"6", "4":"7", "8":"8"};
                return bSiann + bUn + (numMap[parts.tiau] || parts.tiau);
            }
            return applyToneMark(bSiann + bUn, parts.tiau, "BP");
        }

        return bSiann + bUn + parts.tiau;
    }

    function applyPhonetics(upSystem, rightSystem) {
        injectLatinFix();
        document.querySelectorAll('article.article_content > div').forEach(div => { div.className = 'Siang_Pai'; div.style.cssText = ""; });
        const latinSystems = ["台語音標", "台羅拼音", "白話字", "閩拼方案", "閩拼調號", "注音二式", "國際音標"];

        document.querySelectorAll('ruby[data-tlpa]').forEach(ruby => {
            const tlpa = ruby.getAttribute('data-tlpa');
            let hanJi = "";
            for (let node of ruby.childNodes) { if (node.nodeType === 3) { hanJi = node.textContent.trim(); break; } }
            if (!hanJi && ruby.innerText) hanJi = ruby.innerText.split('\n')[0].trim();
            ruby.innerHTML = hanJi;
            if (upSystem) {
                const rt = document.createElement('rt');
                rt.textContent = convertOne(tlpa, upSystem);
                if (latinSystems.includes(upSystem)) rt.classList.add('latin-phonetic');
                ruby.appendChild(rt);
            }
            if (rightSystem) {
                const rtc = document.createElement('rtc');
                rtc.textContent = convertOne(tlpa, rightSystem);
                if (latinSystems.includes(rightSystem)) rtc.classList.add('latin-phonetic');
                ruby.appendChild(rtc);
            }
        });
    }
});
