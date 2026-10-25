/**
 * 漢字標音切換器 (Phonetic Switcher) - 聲韻學強修版
 * 1. 新增【國際音標 (IPA)】轉換功能。
 * 2. 羅馬拼音調符全效支援 (TL, POJ, BP)。
 * 3. 啟動即修復拉丁調符偏移問題。
 * 4. 【雅俗通】為韻＋調＋聲（公三時）；【十五音】為聲＋韻＋調（時公三）。
 * 5. 支援【使用變調標音格式】：依 data-pian-tiau 產出本調／變調 <rt>。
 */
document.addEventListener('DOMContentLoaded', function() {
    let phoneticMapping = null;

    const SUP_DIGITS = "⁰¹²³⁴⁵⁶⁷⁸⁹";
    const CN_TONES = { 1: "一", 2: "二", 3: "三", 4: "四", 5: "五", 7: "七", 8: "八" };
    const TL_TO_BP = { 0: 0, 1: 1, 2: 3, 3: 5, 4: 7, 5: 2, 7: 6, 8: 8 };
    const SANDHI_SHU = { 1: 7, 2: 1, 3: 2, 5: 7, 7: 3 };
    const BP_TONE_SYSTEMS = new Set(["閩拼調號", "閩拼方案", "閩拼調符"]);
    const LATIN_SYSTEMS = ["台語音標", "台羅拼音", "白話字", "閩拼方案", "閩拼調號", "台語注音二式", "國際音標"];

    function getMeta(name) {
        const el = document.querySelector(`meta[name="${name}"]`);
        return el ? String(el.getAttribute("content") || "").trim() : "";
    }

    function usePianTiauFormat() {
        return getMeta("使用變調標音格式") === "是";
    }

    function toSuperscript(n) {
        return String(n).split("").map(ch => {
            const d = parseInt(ch, 10);
            return Number.isNaN(d) ? ch : SUP_DIGITS[d];
        }).join("");
    }

    function normalizeTlTone(tiau) {
        let t = parseInt(tiau, 10);
        if (Number.isNaN(t)) t = 1;
        return t === 6 ? 7 : t;
    }

    function getSandhiToneTl(tiau, un) {
        const t = normalizeTlTone(tiau);
        if (SANDHI_SHU[t] !== undefined) return SANDHI_SHU[t];
        if (t === 4 || t === 8) {
            if (/[ptk]$/.test(un || "")) return t === 4 ? 8 : 4;
            return t === 4 ? 2 : 3; // -h
        }
        return t;
    }

    function toDisplayTone(tlTone, system) {
        if (BP_TONE_SYSTEMS.has(system)) return TL_TO_BP[tlTone] ?? tlTone;
        return tlTone;
    }

    // --- [1. 樣式修復] ---
    const injectLatinFix = () => {
        if (!document.getElementById('latin-phonetic-fix')) {
            const style = document.createElement('style');
            style.id = 'latin-phonetic-fix';
            style.textContent = `
                @font-face {
                    font-family: 'Charis SIL';
                    src: url('./assets/fonts/Charis-Regular.ttf') format('truetype');
                    font-weight: normal;
                    font-style: normal;
                }
                @font-face {
                    font-family: 'Charis SIL';
                    src: url('./assets/fonts/Charis-Bold.ttf') format('truetype');
                    font-weight: bold;
                    font-style: normal;
                }
                rt.latin-phonetic, rtc.latin-phonetic {
                    font-family: 'Charis SIL', 'Source Han Sans TC', sans-serif !important;
                    font-variant-ligatures: common-ligatures !important;
                    font-feature-settings: "kern" 1, "liga" 1, "mark" 1, "mkmk" 1, "ccmp" 1 !important;
                    font-variant-east-asian: normal !important;
                    text-align: center !important;
                    letter-spacing: -0.01em !important;
                    line-height: 1 !important;
                }
                /* 變調格式較長時，避免 <rt> 視覺上壓到相鄰標點 */
                ruby {
                    ruby-overhang: none;
                    -webkit-ruby-overhang: none;
                }
                .Siang_Pai > p > span,
                span.title_mark {
                    ruby-position: unset;
                }
            `;
            document.head.appendChild(style);
        }
    };

    const autoFix = () => {
        const latinRegex = /[a-zA-Z\u0300-\u036f]/;
        document.querySelectorAll('rt, rtc').forEach(el => {
            let text = el.textContent.trim();
            if (latinRegex.test(text)) {
                el.classList.add('latin-phonetic');
                // 修復初始載入的位移問題：加入零寬空格並正規化
                // 注意：勿破壞上標調號（¹⁻⁷），僅對調符式拉丁音標做 NFD
                if (!text.startsWith('\u200B')) {
                    const hasSupTone = /[⁰¹²³⁴⁵⁶⁷⁸⁹]/.test(text);
                    el.textContent = '\u200B' + (hasSupTone ? text : text.normalize("NFD"));
                }
            }
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
        "閩拼調號": { label: "閩拼調號", up: "閩拼調號", right: "" },
        "雅俗通+方音符號": { label: "雅俗通+方音符號", up: "雅俗通", right: "方音符號" },
        "十五音+方音符號": { label: "十五音+方音符號", up: "十五音", right: "方音符號" },
        "台語音標+雅俗通": { label: "台語音標+雅俗通", up: "台語音標", right: "雅俗通" },
        "台語音標+十五音": { label: "台語音標+十五音", up: "台語音標", right: "十五音" },
        "台語音標+方音符號": { label: "台語音標+方音符號", up: "台語音標", right: "方音符號" },
        "閩拼調號+方音符號": { label: "閩拼調號+方音符號", up: "閩拼調號", right: "方音符號" }
    };

    function initSwitcherUI() {
        const toolbar = document.createElement('div');
        toolbar.className = 'phonetic-switcher-toolbar';

        const homeBtn = document.createElement('button');
        homeBtn.className = 'ps-home-btn';
        homeBtn.innerHTML = "🏠 回首頁";
        homeBtn.onclick = () => window.location.href = 'index.html';
        toolbar.appendChild(homeBtn);

        const btnGroup = document.createElement('div');
        btnGroup.className = 'ps-btn-group';
        Object.keys(schemes).forEach(key => {
            const btn = document.createElement('button');
            btn.className = 'ps-scheme-btn';
            btn.textContent = schemes[key].label;
            btn.onclick = () => applyScheme(key);
            btnGroup.appendChild(btn);
        });
        toolbar.appendChild(btnGroup);

        if (usePianTiauFormat()) {
            const badge = document.createElement('span');
            badge.className = 'ps-custom-label';
            badge.style.marginLeft = '8px';
            badge.textContent = '變調格式';
            badge.title = '本頁啟用【使用變調標音格式】：上邊依 data-pian-tiau 顯示本調／變調';
            toolbar.appendChild(badge);
        }

        const customDiv = document.createElement('div');
        customDiv.className = 'ps-custom-div';
        customDiv.innerHTML = `
            <span class="ps-custom-label">自訂</span>
            <span class="ps-custom-text">上:</span><select id="select-up"></select>
            <span class="ps-custom-text">右:</span><select id="select-right"></select>
        `;
        toolbar.appendChild(customDiv);

        const systems = ["無", "雅俗通", "十五音", "方音符號", "國際音標", "台語音標", "台羅拼音", "白話字", "閩拼方案", "閩拼調號", "台語注音二式"];
        const selUp = customDiv.querySelector('#select-up');
        const selRight = customDiv.querySelector('#select-right');
        systems.forEach(s => {
            const val = (s === "無" ? "" : s);
            selUp.add(new Option(s, val));
            selRight.add(new Option(s, val));
        });

        // 依頁面 meta 預設「上／右」選項
        const metaUp = getMeta("上邊標音");
        const metaRight = getMeta("右邊標音");
        if (systems.includes(metaUp)) selUp.value = metaUp;
        if (systems.includes(metaRight)) selRight.value = metaRight;

        const applyCustom = () => applyPhonetics(selUp.value, selRight.value);
        selUp.onchange = applyCustom; selRight.onchange = applyCustom;

        const main = document.querySelector('main.page');
        if (main) {
            main.insertBefore(toolbar, main.firstChild);
            const footer = toolbar.cloneNode(true);
            footer.classList.add('ps-footer');
            footer.querySelector('button').onclick = () => window.location.href = 'index.html';
            footer.querySelectorAll('.ps-btn-group button').forEach((btn, idx) => {
                btn.onclick = () => applyScheme(Object.keys(schemes)[idx]);
            });
            const fSelUp = footer.querySelector('#select-up');
            const fSelRight = footer.querySelector('#select-right');
            fSelUp.value = selUp.value;
            fSelRight.value = selRight.value;
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
        const fSelUp = document.querySelector('.ps-footer #select-up');
        const fSelRight = document.querySelector('.ps-footer #select-right');
        if (fSelUp) fSelUp.value = s.up || "";
        if (fSelRight) fSelRight.value = s.right || "";
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
        const initials = ["tsh", "kh", "ph", "th", "ng", "bb", "gg", "ts", "z", "c", "p", "k", "t", "l", "s", "h", "b", "g", "m", "n", "j"];
        initials.sort((a, b) => b.length - a.length);
        for (let i of initials) { if (p.startsWith(i)) { siann = i; un = p.slice(i.length); break; } }
        if (!siann) un = p;
        return { siann: siann || "ø", un: un, tiau: tiau };
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
        if (pos === -1) return (s + mark).normalize("NFD");
        let res = s.slice(0, pos + 1) + mark + s.slice(pos + 1);
        if (system === "白話字" && base.includes("o\u0358")) res = res.replace("o", "o\u0358");
        return res.normalize("NFD");
    }

    function composeSipNgooIm(parts, initialMatch, finalMatch) {
        let siann = initialMatch ? initialMatch['十五音'] : (parts.siann === "ø" ? "英" : "");
        const un = finalMatch ? finalMatch['十五音'] : "";
        let tiauNum = parseInt(parts.tiau, 10);
        if (tiauNum === 6) tiauNum = 7;
        const tiau = ["", "一", "二", "三", "四", "五", "六", "七", "八"][tiauNum] || parts.tiau;
        return { siann: siann, un: un, tiau: tiau };
    }

    function findInitialFinal(parts) {
        let initialMatch = phoneticMapping.initials.find(i =>
            i.台語音標 === parts.siann ||
            (parts.siann === "ø" && (i.台語音標 === "" || i.台語音標 === "Ø" || i.台語音標 === "ø"))
        );
        let finalMatch = phoneticMapping.finals.find(f => f.台語音標 === parts.un);
        return { initialMatch, finalMatch };
    }

    /** 取得不含調號的音節主體（供本調／變調格式） */
    function getSyllableBody(parts, targetSystem, initialMatch, finalMatch) {
        if (targetSystem === '雅俗通' || targetSystem === '十五音') {
            const sipNgoo = composeSipNgooIm(parts, initialMatch, finalMatch);
            // 本調／變調皆採十五音順序：聲＋韻
            return sipNgoo.siann + sipNgoo.un;
        }
        if (targetSystem === '方音符號') {
            let iTPS = initialMatch ? initialMatch['方音符號'] : "";
            let fTPS = finalMatch ? finalMatch['方音符號'] : "";
            if (parts.un.startsWith('i') && ['z','ts','c','tsh','s','j','ji'].includes(parts.siann)) {
                if (parts.siann === 's') iTPS = 'ㄒ';
                else if (parts.siann === 'j' || parts.siann === 'ji') iTPS = 'ㆢ';
                else if (parts.siann === 'z' || parts.siann === 'ts') iTPS = 'ㄐ';
                else iTPS = 'ㄑ';
            }
            if (iTPS === "Ø" || iTPS === "ø") iTPS = "";
            return iTPS + fTPS;
        }
        if (targetSystem === '台語注音二式') {
            let iMPS = initialMatch ? initialMatch['台語注音二式'] : "";
            let fMPS = finalMatch ? finalMatch['台語注音二式'] : parts.un;
            if (iMPS === "Ø" || iMPS === "ø") iMPS = "";
            if (fMPS.startsWith('i') && ['z', 'c', 's', 'zz'].includes(iMPS)) {
                if (iMPS === 'z') iMPS = 'j';
                else if (iMPS === 'c') iMPS = 'ch';
                else if (iMPS === 's') iMPS = 'sh';
                else if (iMPS === 'zz') iMPS = 'jj';
            }
            return iMPS + fMPS;
        }

        let bSiann = parts.siann === "ø" ? "" : parts.siann;
        let bUn = parts.un;

        if (targetSystem === "白話字") {
            if (bSiann === "ts" || bSiann === "z") bSiann = "ch";
            else if (bSiann === "tsh" || bSiann === "c") bSiann = "chh";
            bUn = bUn.replace(/nnh/g, "h\u207F").replace(/nn/g, "\u207F");
            bUn = bUn.replace(/ue/g, "oe").replace(/ua/g, "oa").replace(/ik/g, "ek").replace(/ing/g, "eng").replace(/oo/g, "o\u0358");
            return bSiann + bUn;
        }
        if (targetSystem === "台羅拼音") {
            if (bSiann === "z") bSiann = "ts";
            if (bSiann === "c") bSiann = "tsh";
            return bSiann + bUn;
        }
        if (targetSystem === "閩拼方案" || targetSystem === "閩拼調號" || targetSystem === "閩拼調符") {
            bUn = bUn.replace(/au/g, "ao");
            if (bUn.endsWith("nn")) {
                let core = bUn.slice(0, -2);
                if (core === "io") core = "ioo";
                else if (core === "o") core = "oo";
                bUn = "n" + core;
            }
            if (bSiann === "") {
                if (bUn.startsWith('i')) {
                    if (bUn === 'i' || /^(in|im|ing|it|ip|ik|ih)/.test(bUn)) bSiann = "y";
                    else { bSiann = "y"; bUn = bUn.substring(1); }
                } else if (bUn.startsWith('u')) {
                    if (bUn === 'u' || /^(un|ut|uh)/.test(bUn)) bSiann = "w";
                    else { bSiann = "w"; bUn = bUn.substring(1); }
                }
            } else {
                bSiann = initialMatch ? initialMatch['閩拼方案'] : bSiann;
            }
            return bSiann + bUn;
        }
        // 台語音標等
        return (bSiann === "ø" ? "" : bSiann) + bUn;
    }

    function formatBunOrBian(body, origDisp, sandhiDisp, useBian, targetSystem) {
        if (targetSystem === '十五音' || targetSystem === '雅俗通') {
            const o = CN_TONES[origDisp] || String(origDisp);
            if (!useBian) return body + o;
            const s = CN_TONES[sandhiDisp] || String(sandhiDisp);
            return body + o + "-" + s;
        }
        if (!useBian) return body + toSuperscript(origDisp);
        return body + toSuperscript(origDisp) + "\u207b" + toSuperscript(sandhiDisp);
    }

    /**
     * 標準標音轉換（非變調格式模式，或右邊標音）。
     */
    function convertOne(tlpa, targetSystem) {
        if (!targetSystem) return "";
        const parts = splitTLPA(tlpa); if (!parts) return "";
        let result = "";
        const { initialMatch, finalMatch } = findInitialFinal(parts);

        if (targetSystem === '雅俗通' || targetSystem === '十五音') {
            const sipNgoo = composeSipNgooIm(parts, initialMatch, finalMatch);
            result = targetSystem === '雅俗通'
                ? sipNgoo.un + sipNgoo.tiau + sipNgoo.siann
                : sipNgoo.siann + sipNgoo.un + sipNgoo.tiau;
        }
        else if (targetSystem === '方音符號') {
            let iTPS = initialMatch ? initialMatch['方音符號'] : "";
            let fTPS = finalMatch ? finalMatch['方音符號'] : "";
            if (parts.un.startsWith('i') && ['z','ts','c','tsh','s','j','ji'].includes(parts.siann)) {
                if (parts.siann === 's') iTPS = 'ㄒ';
                else if (parts.siann === 'j' || parts.siann === 'ji') iTPS = 'ㆢ';
                else if (parts.siann === 'z' || parts.siann === 'ts') iTPS = 'ㄐ';
                else iTPS = 'ㄑ';
            }
            const toneMap = { "1":"", "2":"\u02CB", "3":"\u02EA", "4":"", "5":"\u02CA", "6":"", "7":"\u02EB", "8":"\u02D9" };
            let tTPS = toneMap[parts.tiau] || "";
            result = (iTPS === "Ø" || iTPS === "ø" ? "" : iTPS) + fTPS + tTPS;
        }
        else if (targetSystem === '國際音標') {
            let iIPA = initialMatch ? initialMatch['國際音標'] : "";
            let fIPA = finalMatch ? finalMatch['國際音標'] : parts.un;
            if (iIPA === "Ø" || iIPA === "ø") iIPA = "";
            result = iIPA + fIPA + parts.tiau;
        }
        else if (targetSystem === '台語注音二式') {
            let iMPS = initialMatch ? initialMatch['台語注音二式'] : "";
            let fMPS = finalMatch ? finalMatch['台語注音二式'] : parts.un;
            if (iMPS === "Ø" || iMPS === "ø") iMPS = "";
            if (fMPS.startsWith('i') && ['z', 'c', 's', 'zz'].includes(iMPS)) {
                if (iMPS === 'z') iMPS = 'j';
                else if (iMPS === 'c') iMPS = 'ch';
                else if (iMPS === 's') iMPS = 'sh';
                else if (iMPS === 'zz') iMPS = 'jj';
            }
            result = iMPS + fMPS + parts.tiau;
        }
        else {
            let bSiann = parts.siann === "ø" ? "" : parts.siann;
            let bUn = parts.un;

            if (targetSystem === "白話字") {
                if (bSiann === "ts" || bSiann === "z") bSiann = "ch"; else if (bSiann === "tsh" || bSiann === "c") bSiann = "chh";
                bUn = bUn.replace(/nnh/g, "h\u207F").replace(/nn/g, "\u207F");
                bUn = bUn.replace(/ue/g, "oe").replace(/ua/g, "oa").replace(/ik/g, "ek").replace(/ing/g, "eng").replace(/oo/g, "o\u0358");
                result = applyToneMark(bSiann + bUn, parts.tiau, "白話字");
            }
            else if (targetSystem === "台羅拼音") {
                if (bSiann === "z") bSiann = "ts"; if (bSiann === "c") bSiann = "tsh";
                result = applyToneMark(bSiann + bUn, parts.tiau, "台羅拼音");
            }
            else if (targetSystem === "閩拼方案" || targetSystem === "閩拼調號") {
                bUn = bUn.replace(/au/g, "ao");
                if (bUn.endsWith("nn")) {
                    let core = bUn.slice(0, -2);
                    if (core === "io") core = "ioo";
                    else if (core === "o") core = "oo";
                    bUn = "n" + core;
                }
                if (bSiann === "") {
                    if (bUn.startsWith('i')) {
                        if (bUn === 'i' || /^(in|im|ing|it|ip|ik|ih)/.test(bUn)) bSiann = "y";
                        else { bSiann = "y"; bUn = bUn.substring(1); }
                    }
                    else if (bUn.startsWith('u')) {
                        if (bUn === 'u' || /^(un|ut|uh)/.test(bUn)) bSiann = "w";
                        else { bSiann = "w"; bUn = bUn.substring(1); }
                    }
                } else {
                    bSiann = initialMatch ? initialMatch['閩拼方案'] : bSiann;
                }

                if (targetSystem === "閩拼調號") {
                    const numMap = {"1":"1", "5":"2", "2":"3", "6":"4", "3":"5", "7":"6", "4":"7", "8":"8"};
                    result = bSiann + bUn + (numMap[parts.tiau] || parts.tiau);
                } else {
                    result = applyToneMark(bSiann + bUn, parts.tiau, "BP");
                }
            } else {
                result = bSiann + bUn + parts.tiau;
            }
        }
        return "\u200B" + result;
    }

    /**
     * 上邊標音：若啟用變調格式，依 data-pian-tiau 產出本調／變調。
     */
    function convertUp(tlpa, targetSystem, pianTiau) {
        if (!targetSystem) return "";
        if (!usePianTiauFormat()) {
            return convertOne(tlpa, targetSystem);
        }
        const parts = splitTLPA(tlpa);
        if (!parts) return "";
        const { initialMatch, finalMatch } = findInitialFinal(parts);
        const body = getSyllableBody(parts, targetSystem, initialMatch, finalMatch);
        const tlOrig = normalizeTlTone(parts.tiau);
        const tlSandhi = getSandhiToneTl(parts.tiau, parts.un);
        const dOrig = toDisplayTone(tlOrig, targetSystem);
        const dSandhi = toDisplayTone(tlSandhi, targetSystem);
        const useBian = pianTiau === true || pianTiau === "true";
        return "\u200B" + formatBunOrBian(body, dOrig, dSandhi, useBian, targetSystem);
    }

    function applyPhonetics(upSystem, rightSystem) {
        injectLatinFix();
        document.querySelectorAll('article.article_content > div').forEach(div => {
            div.className = 'Siang_Pai';
            div.style.cssText = "";
        });

        document.querySelectorAll('ruby[data-tlpa]').forEach(ruby => {
            const tlpa = ruby.getAttribute('data-tlpa');
            const pianTiau = ruby.getAttribute('data-pian-tiau') || "false";

            // 只取漢字本文（略過 rt/rtc）；標點不會有 data-tlpa，故不會進入此迴圈
            let hanJi = "";
            for (let node of ruby.childNodes) {
                if (node.nodeType === 3) {
                    const t = node.textContent.trim();
                    if (t) { hanJi = t; break; }
                }
            }
            if (!hanJi) {
                // 後備：取非 rt/rtc 的文字
                const clone = ruby.cloneNode(true);
                clone.querySelectorAll('rt, rtc').forEach(n => n.remove());
                hanJi = clone.textContent.trim();
            }
            if (!hanJi) return;

            // 重建時保留 data-tlpa、data-pian-tiau
            ruby.innerHTML = "";
            ruby.appendChild(document.createTextNode(hanJi));

            if (upSystem) {
                const rt = document.createElement('rt');
                rt.textContent = convertUp(tlpa, upSystem, pianTiau);
                if (LATIN_SYSTEMS.includes(upSystem) || /[a-zA-Z]/.test(rt.textContent)) {
                    rt.classList.add('latin-phonetic');
                }
                ruby.appendChild(rt);
            }
            if (rightSystem) {
                const rtc = document.createElement('rtc');
                // 右邊標音維持標準格式（不變調標示）
                rtc.textContent = convertOne(tlpa, rightSystem);
                if (LATIN_SYSTEMS.includes(rightSystem)) rtc.classList.add('latin-phonetic');
                ruby.appendChild(rtc);
            }
        });
        autoFix();
    }
});
