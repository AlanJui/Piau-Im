
/**
 * 漢字標音切換器 (Phonetic Switcher) - 渲染修復版
 * 修正：
 * 1. 使用 NFC 規範化解決調符偏移問題。
 * 2. 修正閩拼鼻化音與零聲母規則。
 * 3. 確保標調位置與聲調映射完全符合規範。
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
        "閩拼調號": { label: "閩拼調號", up: "閩拼調號", right: "" },
        "十五音": { label: "十五音", up: "十五音", right: "" },
        "十五音+方音符號": { label: "十五音+方音符號", up: "十五音", right: "方音符號" },
        "台語音標+方音符號": { label: "台語音標+方音符號", up: "台語音標", right: "方音符號" },
        "閩拼方案+方音符號": { label: "閩拼方案+方音符號", up: "閩拼方案", right: "方音符號" },
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

    /**
     * 閩拼方案 (BP) 專家級核心轉換
     */
    function convertToBP(siann, un, tiauTLPA, isNumeric) {
        let bSiann = siann === "ø" ? "" : siann;
        let bUn = un;

        // 1. 韻母基礎對應與鼻化前置
        bUn = bUn.replace(/iauh/g, "iaoh").replace(/auh/g, "aoh")
                 .replace(/iau/g, "iao").replace(/au/g, "ao");
        if (bUn.endsWith("nn")) { bUn = "n" + bUn.slice(0, -2); }

        // 2. 零聲母智慧 y/w 規則
        if (bSiann === "") {
            if (bUn.startsWith('n')) {
                if (un.startsWith('i')) bSiann = "y";
                else if (un.startsWith('u')) bSiann = "w";
            } else if (bUn.startsWith('i')) {
                if (bUn === 'i' || /^(in|im|ing|it|ip|ik|ih)/.test(bUn)) bSiann = "y";
                else { bSiann = "y"; bUn = bUn.substring(1); }
            } else if (bUn.startsWith('u')) {
                if (bUn === 'u' || /^(un|ut|uh)/.test(bUn)) bSiann = "w";
                else { bSiann = "w"; bUn = bUn.substring(1); }
            }
        }

        let base = bSiann + bUn;
        
        // 3. 聲調映射 (TLPA -> BP)
        const numMap = {"1":"1", "5":"2", "2":"3", "6":"4", "3":"5", "7":"6", "4":"7", "8":"8"};
        if (isNumeric) return base + (numMap[tiauTLPA] || tiauTLPA);

        const markMap = {
            "1": "\u0304", "5": "\u0301", "2": "\u030c", "6": "\u030c",
            "3": "\u0300", "7": "\u0302", "4": "\u0304", "8": "\u0301"
        };
        let mark = markMap[tiauTLPA] || "";
        
        // 4. 響度優先標調演算法
        let s = base.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
        let pos = -1;
        if (s.includes("ere")) pos = s.indexOf("ere") + 2;
        else if (s.includes("iu")) pos = s.indexOf("u");
        else if (s.includes("ui")) pos = s.indexOf("i");
        else if (s.includes("oo")) pos = s.indexOf("oo");
        else if (s.includes("ng") && !/[aeiou]/.test(s)) pos = s.indexOf("n");
        else if (s.includes("m") && !/[aeiou]/.test(s)) pos = s.indexOf("m");
        else {
            const priority = ['a', 'o', 'e', 'i', 'u', 'v'];
            for (let v of priority) { if (s.includes(v)) { pos = s.indexOf(v); break; } }
        }

        // 5. 輸出並進行 NFC 規範化 (解決調符偏移的關鍵)
        let result = (pos === -1 || !mark) ? s : (s.slice(0, pos + 1) + mark + s.slice(pos + 1));
        return result.normalize("NFC");
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
            if (!iName && parts.siann === "ø") iName = "英";
            const fName = finalMatch ? finalMatch['十五音'] : "";
            const toneCN = ["", "一", "二", "三", "四", "五", "六", "七", "八"][parseInt(parts.tiau)] || parts.tiau;
            return fName + toneCN + iName;
        }

        if (targetSystem === '方音符號') {
            let iTPS = initialMatch ? initialMatch['方音符號'] : "";
            let fTPS = finalMatch ? finalMatch['方音符號'] : "";
            let tTPS = (toneMatch && toneMatch['注音調符編碼']) || "";
            if (tTPS.startsWith('U+')) tTPS = String.fromCharCode(parseInt(tTPS.substring(2), 16));
            if (parts.un.startsWith('i') && ['z','ts','c','tsh','s','j','ji'].includes(parts.siann)) {
                if (parts.siann === 's') iTPS = 'ㄒ';
                else if (parts.siann === 'j' || parts.siann === 'ji') iTPS = 'ㆢ';
                else if (parts.siann === 'z' || parts.siann === 'ts') iTPS = 'ㄐ';
                else iTPS = 'ㄑ';
            }
            return (iTPS === "Ø" ? "" : iTPS) + fTPS + tTPS;
        }

        if (targetSystem === "閩拼方案" || targetSystem === "閩拼調號") {
            return convertToBP(parts.siann, parts.un, parts.tiau, targetSystem === "閩拼調號");
        }

        let iPin = initialMatch ? (initialMatch[targetSystem] || initialMatch['台羅拼音'] || parts.siann) : parts.siann;
        let fPin = finalMatch ? (finalMatch[targetSystem] || finalMatch['台羅拼音'] || parts.un) : parts.un;
        if (iPin === "Ø" || iPin === "ø") iPin = "";
        return iPin + fPin + parts.tiau;
    }

    function applyPhonetics(upSystem, rightSystem) {
        // 注入專屬的拉丁拼音修復樣式 (只注入一次)
        if (!document.getElementById('latin-phonetic-fix')) {
            const style = document.createElement('style');
            style.id = 'latin-phonetic-fix';
            // 針對拉丁字母拼音：關閉 salt 特性，並優先使用純英文字型，徹底解決調符偏移與全形字距問題
            style.textContent = `
                rt.latin-phonetic, rtc.latin-phonetic {
                    font-family: Arial, Helvetica, sans-serif !important;
                    font-feature-settings: "normal" !important;
                    font-variant-east-asian: normal !important;
                }
            `;
            document.head.appendChild(style);
        }

        document.querySelectorAll('article.article_content > div').forEach(div => {
            div.className = 'Siang_Pai';
            div.style.cssText = ""; 
        });

        const latinSystems = ["台語音標", "台羅拼音", "白話字", "閩拼方案", "閩拼調號", "注音二式"];

        document.querySelectorAll('ruby[data-tlpa]').forEach(ruby => {
            const tlpa = ruby.getAttribute('data-tlpa');
            let hanJi = "";
            for (let node of ruby.childNodes) { if (node.nodeType === 3) { hanJi = node.textContent.trim(); break; } }
            if (!hanJi && ruby.innerText) hanJi = ruby.innerText.split('\n')[0].trim();

            ruby.innerHTML = hanJi;
            
            if (upSystem) {
                const rt = document.createElement('rt');
                rt.textContent = convertOne(tlpa, upSystem);
                if (latinSystems.includes(upSystem)) rt.className = "latin-phonetic";
                ruby.appendChild(rt);
            }
            if (rightSystem) {
                const rtc = document.createElement('rtc');
                rtc.textContent = convertOne(tlpa, rightSystem);
                if (latinSystems.includes(rightSystem)) rtc.className = "latin-phonetic";
                ruby.appendChild(rtc);
            }
        });
    }
});
