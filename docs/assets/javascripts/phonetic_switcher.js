/**
 * 漢字標音切換器 (Phonetic Switcher)
 * 依 PRG-web01：預設標音樣式／自訂上端・右端／變調標音開關。
 * - 變調標音預設「關閉」；啟用後僅 <rt> 依 data-pian-tiau 產出本調／變調；<rtc> 維持標準格式。
 * - data-pian-tiau 缺省等同 "false"（本調格式）。
 * - 聲母 n + 韻母 ng（nng7）：調符標在第 2 個 n（nn̄g），不標在第 1 個 n。
 */
document.addEventListener("DOMContentLoaded", function () {
    let phoneticMapping = null;

    const SUP_DIGITS = "⁰¹²³⁴⁵⁶⁷⁸⁹";
    const CN_TONES = { 1: "一", 2: "二", 3: "三", 4: "四", 5: "五", 7: "七", 8: "八" };
    const TL_TO_BP = { 0: 0, 1: 1, 2: 3, 3: 5, 4: 7, 5: 2, 7: 6, 8: 8 };
    const SANDHI_SHU = { 1: 7, 2: 1, 3: 2, 5: 7, 7: 3 };
    const BP_TONE_SYSTEMS = new Set(["閩拼調號", "閩拼方案", "閩拼調符"]);
    const LATIN_SYSTEMS = ["台語音標", "台羅拼音", "白話字", "閩拼方案", "閩拼調號", "閩拼調符", "台語注音二式", "國際音標"];

    const PREDEFINE_SCHEMES = {
        "十五音+方音符號": { up: "十五音", right: "方音符號" },
        "台語音標+雅俗通": { up: "台語音標", right: "雅俗通" },
        "台語音標+十五音": { up: "台語音標", right: "十五音" },
        "台語音標+方音符號": { up: "台語音標", right: "方音符號" },
        "閩拼調號+方音符號": { up: "閩拼調號", right: "方音符號" },
        "方音符號": { up: "無", right: "方音符號" },
        "台語音標": { up: "台語音標", right: "無" },
        "台羅拼音": { up: "台羅拼音", right: "無" },
        "白話字": { up: "白話字", right: "無" },
        "閩拼方案": { up: "閩拼方案", right: "無" },
    };

    const TOP_SYSTEMS = [
        "無", "雅俗通", "十五音", "方音符號", "國際音標",
        "台語音標", "台羅拼音", "白話字", "閩拼方案", "閩拼調號", "台語注音二式",
    ];
    const RIGHT_SYSTEMS = ["無", "雅俗通", "十五音", "方音符號"];

    function getMeta(name) {
        const el = document.querySelector(`meta[name="${name}"]`);
        return el ? String(el.getAttribute("content") || "").trim() : "";
    }

    function normSystem(v) {
        if (!v || v === "無" || v === "None") return "";
        return v;
    }

    function isPianTiauOn() {
        const sel = document.querySelector(
            '.phonetic-switcher-toolbar:not(.ps-footer) [name="pian-tiau-switcher"]'
        );
        return sel ? sel.value === "on" : false;
    }

    function toSuperscript(n) {
        return String(n).split("").map((ch) => {
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
            return t === 4 ? 2 : 3;
        }
        return t;
    }

    function toDisplayTone(tlTone, system) {
        if (BP_TONE_SYSTEMS.has(system)) return TL_TO_BP[tlTone] ?? tlTone;
        return tlTone;
    }

    // --- [1. 樣式修復] ---
    const injectLatinFix = () => {
        if (document.getElementById("latin-phonetic-fix")) return;
        const style = document.createElement("style");
        style.id = "latin-phonetic-fix";
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
    };

    const autoFix = () => {
        const latinRegex = /[a-zA-Z\u0300-\u036f]/;
        document.querySelectorAll("rt, rtc").forEach((el) => {
            let text = el.textContent.trim();
            if (latinRegex.test(text)) {
                el.classList.add("latin-phonetic");
                if (!text.startsWith("\u200B")) {
                    const hasSupTone = /[⁰¹²³⁴⁵⁶⁷⁸⁹]/.test(text);
                    el.textContent = "\u200B" + (hasSupTone ? text : text.normalize("NFD"));
                }
            }
        });
    };

    injectLatinFix();
    autoFix();

    // --- [2. 資料載入] ---
    fetch("./assets/javascripts/phonetic_mapping.json")
        .then((response) => response.json())
        .then((data) => {
            phoneticMapping = data;
            initSwitcherUI();
        })
        .catch((err) => console.error("無法載入標音對照表:", err));

    function makeDivider() {
        const span = document.createElement("span");
        span.className = "ps-divider";
        span.setAttribute("aria-hidden", "true");
        return span;
    }

    function fillSelect(sel, options, selected) {
        sel.innerHTML = "";
        options.forEach((opt) => {
            sel.add(new Option(opt, opt));
        });
        if (selected && options.includes(selected)) {
            sel.value = selected;
        }
    }

    function findMatchingPredefine(up, right) {
        const u = up || "無";
        const r = right || "無";
        for (const [key, s] of Object.entries(PREDEFINE_SCHEMES)) {
            if (s.up === u && s.right === r) return key;
        }
        return "";
    }

    function getToolbarControls(root) {
        return {
            predefine: root.querySelector('[name="phonetic-styles-predefine"]'),
            top: root.querySelector('[name="phonetic-styles-custom-top"]'),
            right: root.querySelector('[name="phonetic-styles-custom-right"]'),
            pian: root.querySelector('[name="pian-tiau-switcher"]'),
        };
    }

    function syncAllToolbars(source) {
        document.querySelectorAll(".phonetic-switcher-toolbar").forEach((tb) => {
            if (tb === source) return;
            const from = getToolbarControls(source);
            const to = getToolbarControls(tb);
            if (from.predefine && to.predefine) to.predefine.value = from.predefine.value;
            if (from.top && to.top) to.top.value = from.top.value;
            if (from.right && to.right) to.right.value = from.right.value;
            if (from.pian && to.pian) to.pian.value = from.pian.value;
        });
    }

    function currentSystems() {
        const top = document.querySelector(
            '.phonetic-switcher-toolbar:not(.ps-footer) [name="phonetic-styles-custom-top"]'
        );
        const right = document.querySelector(
            '.phonetic-switcher-toolbar:not(.ps-footer) [name="phonetic-styles-custom-right"]'
        );
        return {
            up: normSystem(top ? top.value : ""),
            right: normSystem(right ? right.value : ""),
        };
    }

    function applyCurrent() {
        const { up, right } = currentSystems();
        applyPhonetics(up, right);
    }

    function buildToolbar() {
        const toolbar = document.createElement("div");
        toolbar.className = "phonetic-switcher-toolbar";

        const homeBtn = document.createElement("button");
        homeBtn.type = "button";
        homeBtn.className = "ps-home-btn";
        homeBtn.textContent = "🏠 回首頁";
        homeBtn.onclick = () => { window.location.href = "index.html"; };
        toolbar.appendChild(homeBtn);
        toolbar.appendChild(makeDivider());

        // 預設標音樣式
        const preGroup = document.createElement("div");
        preGroup.className = "ps-group";
        const preLabel = document.createElement("label");
        preLabel.htmlFor = "phonetic-styles-predefine";
        preLabel.textContent = "預設標音樣式：";
        const preSelect = document.createElement("select");
        preSelect.id = "phonetic-styles-predefine";
        preSelect.name = "phonetic-styles-predefine";
        Object.keys(PREDEFINE_SCHEMES).forEach((key) => {
            preSelect.add(new Option(key, key));
        });
        preGroup.appendChild(preLabel);
        preGroup.appendChild(preSelect);
        toolbar.appendChild(preGroup);
        toolbar.appendChild(makeDivider());

        // 自訂標音樣式
        const customGroup = document.createElement("div");
        customGroup.className = "ps-group ps-custom-div";
        const customTitle = document.createElement("label");
        customTitle.textContent = "自訂標音樣式";
        customTitle.className = "ps-custom-label";

        const topLabel = document.createElement("label");
        topLabel.htmlFor = "phonetic-styles-custom-top";
        topLabel.textContent = "上端：";
        const topSelect = document.createElement("select");
        topSelect.id = "phonetic-styles-custom-top";
        topSelect.name = "phonetic-styles-custom-top";

        const rightLabel = document.createElement("label");
        rightLabel.htmlFor = "phonetic-styles-custom-right";
        rightLabel.textContent = "右端：";
        const rightSelect = document.createElement("select");
        rightSelect.id = "phonetic-styles-custom-right";
        rightSelect.name = "phonetic-styles-custom-right";

        customGroup.appendChild(customTitle);
        customGroup.appendChild(topLabel);
        customGroup.appendChild(topSelect);
        customGroup.appendChild(rightLabel);
        customGroup.appendChild(rightSelect);
        toolbar.appendChild(customGroup);
        toolbar.appendChild(makeDivider());

        // 變調標音
        const pianGroup = document.createElement("div");
        pianGroup.className = "ps-group";
        const pianLabel = document.createElement("label");
        pianLabel.htmlFor = "pian-tiau-switcher";
        pianLabel.textContent = "變調標音：";
        const pianSelect = document.createElement("select");
        pianSelect.id = "pian-tiau-switcher";
        pianSelect.name = "pian-tiau-switcher";
        pianSelect.add(new Option("關閉", "off"));
        pianSelect.add(new Option("啟用", "on"));
        pianSelect.value = "off";
        pianGroup.appendChild(pianLabel);
        pianGroup.appendChild(pianSelect);
        toolbar.appendChild(pianGroup);

        return toolbar;
    }

    function wireToolbar(toolbar) {
        const { predefine, top, right, pian } = getToolbarControls(toolbar);

        predefine.onchange = () => {
            const scheme = PREDEFINE_SCHEMES[predefine.value];
            if (scheme) {
                top.value = scheme.up;
                right.value = scheme.right;
            }
            syncAllToolbars(toolbar);
            applyCurrent();
        };

        const onCustomChange = () => {
            const match = findMatchingPredefine(top.value, right.value);
            if (match) predefine.value = match;
            syncAllToolbars(toolbar);
            applyCurrent();
        };
        top.onchange = onCustomChange;
        right.onchange = onCustomChange;

        pian.onchange = () => {
            syncAllToolbars(toolbar);
            applyCurrent();
        };
    }

    function initSwitcherUI() {
        const metaUp = getMeta("上邊標音") || "無";
        const metaRight = getMeta("右邊標音") || "無";
        const initUp = TOP_SYSTEMS.includes(metaUp) ? metaUp : "無";
        const initRight = RIGHT_SYSTEMS.includes(metaRight) ? metaRight : "無";
        // 變調標音：規範預設為關閉；若頁面 meta 為「是」則初始啟用（與 a400 產出一致）
        const initPian = getMeta("使用變調標音格式") === "是" ? "on" : "off";

        const mainEl = document.querySelector("main.page");
        if (!mainEl) return;

        function prepareToolbar(isFooter) {
            const toolbar = buildToolbar();
            if (isFooter) {
                toolbar.classList.add("ps-footer");
                toolbar.querySelectorAll("[id]").forEach((el) => {
                    el.id = "footer-" + el.id;
                });
                toolbar.querySelectorAll("label[for]").forEach((lab) => {
                    if (!lab.htmlFor.startsWith("footer-")) {
                        lab.htmlFor = "footer-" + lab.htmlFor;
                    }
                });
            }
            const ctrls = getToolbarControls(toolbar);
            fillSelect(ctrls.top, TOP_SYSTEMS, initUp);
            fillSelect(ctrls.right, RIGHT_SYSTEMS, initRight);
            ctrls.pian.value = initPian;
            const match = findMatchingPredefine(initUp, initRight);
            if (match) ctrls.predefine.value = match;
            wireToolbar(toolbar);
            return toolbar;
        }

        mainEl.insertBefore(prepareToolbar(false), mainEl.firstChild);
        mainEl.appendChild(prepareToolbar(true));

        if (initPian === "on") {
            applyCurrent();
        }
    }

    function splitTLPA(tlpa) {
        if (!tlpa) return null;
        let p = tlpa.toLowerCase();
        let tiau = p.slice(-1);
        if (!isNaN(tiau)) {
            p = p.slice(0, -1);
        } else {
            if (["p", "t", "k", "h"].includes(tiau)) tiau = "4";
            else tiau = "1";
        }
        let siann = "";
        let un = "";
        const initials = ["tsh", "kh", "ph", "th", "ng", "bb", "gg", "ts", "z", "c", "p", "k", "t", "l", "s", "h", "b", "g", "m", "n", "j"];
        initials.sort((a, b) => b.length - a.length);
        for (const i of initials) {
            if (p.startsWith(i)) {
                siann = i;
                un = p.slice(i.length);
                break;
            }
        }
        if (!siann) un = p;
        return { siann: siann || "ø", un: un, tiau: tiau };
    }

    function applyToneMark(base, tiau, system) {
        const marks = {
            TL: { 2: "\u0301", 3: "\u0300", 5: "\u0302", 6: "\u030c", 7: "\u0304", 8: "\u030d" },
            POJ: { 2: "\u0301", 3: "\u0300", 5: "\u0302", 7: "\u0304", 8: "\u030d" },
            BP: { 1: "\u0304", 5: "\u0301", 2: "\u030c", 6: "\u030c", 3: "\u0300", 7: "\u0302", 4: "\u0304", 8: "\u0301" },
        };
        const sys = system === "白話字" ? "POJ" : system === "台羅拼音" ? "TL" : "BP";
        const mark = marks[sys][tiau] || "";
        if (!mark) return base;

        let s = base.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
        let pos = -1;
        if (s.includes("ere")) pos = s.indexOf("ere") + 2;
        else if (s.includes("iu")) pos = s.indexOf("u");
        else if (s.includes("ui")) pos = s.indexOf("i");
        else if (s.includes("oo")) pos = s.indexOf("o");
        else if (s.includes("ng") && !/[aeiou]/.test(s)) {
            // 韻母 ng 的調符標在 ng 的 n。nng（聲母 n + 韻母 ng）標在第 2 個 n。
            const ngN = s.search(/n(?=g)/i);
            pos = ngN === -1 ? s.indexOf("n") : ngN;
        } else if (s.includes("m") && !/[aeiou]/.test(s)) pos = s.indexOf("m");
        else {
            for (const v of ["a", "o", "e", "i", "u", "v"]) {
                if (s.includes(v)) {
                    pos = s.indexOf(v);
                    break;
                }
            }
        }
        if (pos === -1) return (s + mark).normalize("NFD");
        let res = s.slice(0, pos + 1) + mark + s.slice(pos + 1);
        if (system === "白話字" && base.includes("o\u0358")) res = res.replace("o", "o\u0358");
        return res.normalize("NFD");
    }

    function composeSipNgooIm(parts, initialMatch, finalMatch) {
        const siann = initialMatch ? initialMatch["十五音"] : parts.siann === "ø" ? "英" : "";
        const un = finalMatch ? finalMatch["十五音"] : "";
        let tiauNum = parseInt(parts.tiau, 10);
        if (tiauNum === 6) tiauNum = 7;
        const tiau = ["", "一", "二", "三", "四", "五", "六", "七", "八"][tiauNum] || parts.tiau;
        return { siann, un, tiau };
    }

    function findInitialFinal(parts) {
        const initialMatch = phoneticMapping.initials.find(
            (i) =>
                i.台語音標 === parts.siann ||
                (parts.siann === "ø" && (i.台語音標 === "" || i.台語音標 === "Ø" || i.台語音標 === "ø"))
        );
        const finalMatch = phoneticMapping.finals.find((f) => f.台語音標 === parts.un);
        return { initialMatch, finalMatch };
    }

    function getSyllableBody(parts, targetSystem, initialMatch, finalMatch) {
        if (targetSystem === "雅俗通" || targetSystem === "十五音") {
            const sipNgoo = composeSipNgooIm(parts, initialMatch, finalMatch);
            // 本調／變調：雅俗通與十五音皆採「聲＋韻」（地觀）
            return sipNgoo.siann + sipNgoo.un;
        }
        if (targetSystem === "方音符號") {
            let iTPS = initialMatch ? initialMatch["方音符號"] : "";
            let fTPS = finalMatch ? finalMatch["方音符號"] : "";
            if (parts.un.startsWith("i") && ["z", "ts", "c", "tsh", "s", "j", "ji"].includes(parts.siann)) {
                if (parts.siann === "s") iTPS = "ㄒ";
                else if (parts.siann === "j" || parts.siann === "ji") iTPS = "ㆢ";
                else if (parts.siann === "z" || parts.siann === "ts") iTPS = "ㄐ";
                else iTPS = "ㄑ";
            }
            if (iTPS === "Ø" || iTPS === "ø") iTPS = "";
            return iTPS + fTPS;
        }
        if (targetSystem === "台語注音二式") {
            let iMPS = initialMatch ? initialMatch["台語注音二式"] : "";
            let fMPS = finalMatch ? finalMatch["台語注音二式"] : parts.un;
            if (iMPS === "Ø" || iMPS === "ø") iMPS = "";
            if (fMPS.startsWith("i") && ["z", "c", "s", "zz"].includes(iMPS)) {
                if (iMPS === "z") iMPS = "j";
                else if (iMPS === "c") iMPS = "ch";
                else if (iMPS === "s") iMPS = "sh";
                else if (iMPS === "zz") iMPS = "jj";
            }
            return iMPS + fMPS;
        }

        let bSiann = parts.siann === "ø" ? "" : parts.siann;
        let bUn = parts.un;

        if (targetSystem === "白話字") {
            if (bSiann === "ts" || bSiann === "z") bSiann = "ch";
            else if (bSiann === "tsh" || bSiann === "c") bSiann = "chh";
            bUn = bUn.replace(/nnh/g, "h\u207F").replace(/nn/g, "\u207F");
            bUn = bUn
                .replace(/ue/g, "oe")
                .replace(/ua/g, "oa")
                .replace(/ik/g, "ek")
                .replace(/ing/g, "eng")
                .replace(/oo/g, "o\u0358");
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
                if (bUn.startsWith("i")) {
                    if (bUn === "i" || /^(in|im|ing|it|ip|ik|ih)/.test(bUn)) bSiann = "y";
                    else {
                        bSiann = "y";
                        bUn = bUn.substring(1);
                    }
                } else if (bUn.startsWith("u")) {
                    if (bUn === "u" || /^(un|ut|uh)/.test(bUn)) bSiann = "w";
                    else {
                        bSiann = "w";
                        bUn = bUn.substring(1);
                    }
                }
            } else {
                bSiann = initialMatch ? initialMatch["閩拼方案"] : bSiann;
            }
            return bSiann + bUn;
        }
        return (bSiann === "ø" ? "" : bSiann) + bUn;
    }

    function formatBunOrBian(body, origDisp, sandhiDisp, useBian, targetSystem) {
        if (targetSystem === "十五音" || targetSystem === "雅俗通") {
            const o = CN_TONES[origDisp] || String(origDisp);
            if (!useBian) return body + o;
            const s = CN_TONES[sandhiDisp] || String(sandhiDisp);
            return body + o + "-" + s;
        }
        if (!useBian) return body + toSuperscript(origDisp);
        return body + toSuperscript(origDisp) + "\u207b" + toSuperscript(sandhiDisp);
    }

    /** 標準標音（變調關閉時） */
    function convertOne(tlpa, targetSystem) {
        if (!targetSystem) return "";
        const parts = splitTLPA(tlpa);
        if (!parts) return "";
        let result = "";
        const { initialMatch, finalMatch } = findInitialFinal(parts);

        if (targetSystem === "雅俗通" || targetSystem === "十五音") {
            const sipNgoo = composeSipNgooIm(parts, initialMatch, finalMatch);
            result =
                targetSystem === "雅俗通"
                    ? sipNgoo.un + sipNgoo.tiau + sipNgoo.siann
                    : sipNgoo.siann + sipNgoo.un + sipNgoo.tiau;
        } else if (targetSystem === "方音符號") {
            let iTPS = initialMatch ? initialMatch["方音符號"] : "";
            let fTPS = finalMatch ? finalMatch["方音符號"] : "";
            if (parts.un.startsWith("i") && ["z", "ts", "c", "tsh", "s", "j", "ji"].includes(parts.siann)) {
                if (parts.siann === "s") iTPS = "ㄒ";
                else if (parts.siann === "j" || parts.siann === "ji") iTPS = "ㆢ";
                else if (parts.siann === "z" || parts.siann === "ts") iTPS = "ㄐ";
                else iTPS = "ㄑ";
            }
            const toneMap = {
                1: "",
                2: "\u02CB",
                3: "\u02EA",
                4: "",
                5: "\u02CA",
                6: "",
                7: "\u02EB",
                8: "\u02D9",
            };
            const tTPS = toneMap[parts.tiau] || "";
            result = (iTPS === "Ø" || iTPS === "ø" ? "" : iTPS) + fTPS + tTPS;
        } else if (targetSystem === "國際音標") {
            let iIPA = initialMatch ? initialMatch["國際音標"] : "";
            const fIPA = finalMatch ? finalMatch["國際音標"] : parts.un;
            if (iIPA === "Ø" || iIPA === "ø") iIPA = "";
            result = iIPA + fIPA + parts.tiau;
        } else if (targetSystem === "台語注音二式") {
            let iMPS = initialMatch ? initialMatch["台語注音二式"] : "";
            let fMPS = finalMatch ? finalMatch["台語注音二式"] : parts.un;
            if (iMPS === "Ø" || iMPS === "ø") iMPS = "";
            if (fMPS.startsWith("i") && ["z", "c", "s", "zz"].includes(iMPS)) {
                if (iMPS === "z") iMPS = "j";
                else if (iMPS === "c") iMPS = "ch";
                else if (iMPS === "s") iMPS = "sh";
                else if (iMPS === "zz") iMPS = "jj";
            }
            result = iMPS + fMPS + parts.tiau;
        } else {
            let bSiann = parts.siann === "ø" ? "" : parts.siann;
            let bUn = parts.un;

            if (targetSystem === "白話字") {
                if (bSiann === "ts" || bSiann === "z") bSiann = "ch";
                else if (bSiann === "tsh" || bSiann === "c") bSiann = "chh";
                bUn = bUn.replace(/nnh/g, "h\u207F").replace(/nn/g, "\u207F");
                bUn = bUn
                    .replace(/ue/g, "oe")
                    .replace(/ua/g, "oa")
                    .replace(/ik/g, "ek")
                    .replace(/ing/g, "eng")
                    .replace(/oo/g, "o\u0358");
                result = applyToneMark(bSiann + bUn, parts.tiau, "白話字");
            } else if (targetSystem === "台羅拼音") {
                if (bSiann === "z") bSiann = "ts";
                if (bSiann === "c") bSiann = "tsh";
                result = applyToneMark(bSiann + bUn, parts.tiau, "台羅拼音");
            } else if (targetSystem === "閩拼方案" || targetSystem === "閩拼調號" || targetSystem === "閩拼調符") {
                bUn = bUn.replace(/au/g, "ao");
                if (bUn.endsWith("nn")) {
                    let core = bUn.slice(0, -2);
                    if (core === "io") core = "ioo";
                    else if (core === "o") core = "oo";
                    bUn = "n" + core;
                }
                if (bSiann === "") {
                    if (bUn.startsWith("i")) {
                        if (bUn === "i" || /^(in|im|ing|it|ip|ik|ih)/.test(bUn)) bSiann = "y";
                        else {
                            bSiann = "y";
                            bUn = bUn.substring(1);
                        }
                    } else if (bUn.startsWith("u")) {
                        if (bUn === "u" || /^(un|ut|uh)/.test(bUn)) bSiann = "w";
                        else {
                            bSiann = "w";
                            bUn = bUn.substring(1);
                        }
                    }
                } else {
                    bSiann = initialMatch ? initialMatch["閩拼方案"] : bSiann;
                }

                if (targetSystem === "閩拼調號") {
                    const numMap = { 1: "1", 5: "2", 2: "3", 6: "4", 3: "5", 7: "6", 4: "7", 8: "8" };
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

    /** 本調／變調格式（僅套用於 <rt>） */
    function convertBunBian(tlpa, targetSystem, pianTiau) {
        if (!targetSystem) return "";
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

    /** <rt>：變調啟用時採本調／變調；否則標準格式 */
    function convertRt(tlpa, targetSystem, pianTiauAttr) {
        if (!targetSystem) return "";
        if (isPianTiauOn()) {
            return convertBunBian(tlpa, targetSystem, pianTiauAttr);
        }
        return convertOne(tlpa, targetSystem);
    }

    function applyPhonetics(upSystem, rightSystem) {
        injectLatinFix();
        document.querySelectorAll("article.article_content > div").forEach((div) => {
            div.className = "Siang_Pai";
            div.style.cssText = "";
        });

        document.querySelectorAll("ruby[data-tlpa]").forEach((ruby) => {
            const tlpa = ruby.getAttribute("data-tlpa");
            const pianTiau = ruby.getAttribute("data-pian-tiau") || "false";

            let hanJi = "";
            for (const node of ruby.childNodes) {
                if (node.nodeType === 3) {
                    const t = node.textContent.trim();
                    if (t) {
                        hanJi = t;
                        break;
                    }
                }
            }
            if (!hanJi) {
                const clone = ruby.cloneNode(true);
                clone.querySelectorAll("rt, rtc").forEach((n) => n.remove());
                hanJi = clone.textContent.trim();
            }
            if (!hanJi) return;

            ruby.innerHTML = "";
            ruby.appendChild(document.createTextNode(hanJi));

            if (upSystem) {
                const rt = document.createElement("rt");
                rt.textContent = convertRt(tlpa, upSystem, pianTiau);
                if (LATIN_SYSTEMS.includes(upSystem) || /[a-zA-Z]/.test(rt.textContent)) {
                    rt.classList.add("latin-phonetic");
                }
                ruby.appendChild(rt);
            }
            if (rightSystem) {
                const rtc = document.createElement("rtc");
                // <rtc> 一律標準標音格式
                rtc.textContent = convertOne(tlpa, rightSystem);
                if (LATIN_SYSTEMS.includes(rightSystem)) rtc.classList.add("latin-phonetic");
                ruby.appendChild(rtc);
            }
        });
        autoFix();
    }
});
