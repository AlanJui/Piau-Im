
document.addEventListener('DOMContentLoaded', function() {
    let phoneticMapping = null;

    fetch('assets/javascripts/phonetic_mapping.json')
        .then(response => response.json())
        .then(data => {
            phoneticMapping = data;
            initSwitcher();
        })
        .catch(err => console.error('無法載入標音對照表:', err));

    function initSwitcher() {
        const toolbar = document.createElement('div');
        toolbar.className = 'phonetic-switcher-toolbar';
        
        const label = document.createElement('label');
        label.textContent = '切換標音方式：';
        toolbar.appendChild(label);

        const select = document.createElement('select');
        const options = [
            { value: 'original', text: '預設 (檔案設定)' },
            { value: '十五音', text: '十五音' },
            { value: '方音符號', text: '方音符號' },
            { value: '台語音標', text: '台語音標 (TLPA+)' },
            { value: '台羅拼音', text: '台羅拼音' },
            { value: '白話字', text: '白話字 (POJ)' },
            { value: '閩拼方案', text: '閩拼方案 (BP)' },
            { value: '注音二式', text: '注音二式' }
        ];

        options.forEach(opt => {
            const el = document.createElement('option');
            el.value = opt.value;
            el.textContent = opt.text;
            select.appendChild(el);
        });

        select.addEventListener('change', function() {
            switchPhonetics(this.value);
        });

        toolbar.appendChild(select);
        const main = document.querySelector('main.page');
        if (main) main.insertBefore(toolbar, main.firstChild);
    }

    function splitTLPA(tlpa) {
        if (!tlpa) return null;
        let p = tlpa.toLowerCase();
        
        let tiau = p.slice(-1);
        if (!isNaN(tiau)) {
            p = p.slice(0, -1);
        } else {
            if (['p', 't', 'k', 'h'].includes(tiau)) tiau = '4';
            else tiau = '1';
        }

        let siann = "";
        let un = "";
        
        // 聲母比對：長字優先，並包含 z, c
        const initials = ["tsh", "kh", "ph", "th", "ng", "bb", "gg", "ch", "zh", "ji", "ts", "z", "c", "p", "k", "t", "l", "s", "h", "b", "g", "m", "n", "j"];
        initials.sort((a, b) => b.length - a.length);

        for (let i of initials) {
            if (p.startsWith(i)) {
                siann = i;
                un = p.slice(i.length);
                break;
            }
        }
        
        if (!siann) un = p;
        return { siann: siann || "ø", un: un, tiau: tiau };
    }

    function convertTo(tlpa, targetSystem) {
        const parts = splitTLPA(tlpa);
        if (!parts) return "";

        let initialMatch = phoneticMapping.initials.find(i => 
            i.台語音標 === parts.siann || (parts.siann === "ø" && i.台語音標 === "")
        );
        
        let finalMatch = phoneticMapping.finals.find(f => f.台語音標 === parts.un);
        
        let toneMatch = phoneticMapping.tones.find(t => 
            String(t.台羅調號) === parts.tiau || String(t.台語音標).endsWith(parts.tiau)
        );

        if (targetSystem === '十五音') {
            const iName = initialMatch ? initialMatch['十五音'] : "";
            const fName = finalMatch ? finalMatch['十五音'] : "";
            const tNum = toneMatch ? toneMatch['識別號'] : parts.tiau;
            return iName + fName + (tNum || "");
        }

        if (targetSystem === '方音符號') {
            let iTPS = initialMatch ? initialMatch['方音符號'] : "";
            let fTPS = finalMatch ? finalMatch['方音符號'] : "";
            const tTPS = (toneMatch && toneMatch['注音調符編碼']) || "";

            // 修正顎化音 (j/z/c/s + i)
            if (parts.un.startsWith('i')) {
                let changed = false;
                if (parts.siann === 'z') { iTPS = 'ㄐ'; changed = true; }
                else if (parts.siann === 'c') { iTPS = 'ㄑ'; changed = true; }
                else if (parts.siann === 's') { iTPS = 'ㄒ'; changed = true; }
                else if (parts.siann === 'j') { iTPS = 'ㆢ'; changed = true; }

                if (changed && fTPS.startsWith('ㄧ')) {
                    fTPS = fTPS.substring(1); // 併入 ㄧ
                }
            }

            return (iTPS === "Ø" ? "" : iTPS) + fTPS + tTPS;
        }

        // 其餘拼音方式
        let iPin = initialMatch ? (initialMatch[targetSystem] || initialMatch['台羅拼音'] || parts.siann) : parts.siann;
        let fPin = finalMatch ? (finalMatch[targetSystem] || finalMatch['台羅拼音'] || parts.un) : parts.un;
        
        // 處理 Ø 顯示
        iPin = (iPin === "Ø" ? "" : iPin);
        
        return iPin + fPin + parts.tiau;
    }

    function switchPhonetics(system) {
        const rubies = document.querySelectorAll('ruby[data-tlpa]');
        
        rubies.forEach(ruby => {
            const tlpa = ruby.getAttribute('data-tlpa');
            if (!tlpa) return;

            if (system === 'original') {
                location.reload();
                return;
            }

            const converted = convertTo(tlpa, system);
            
            if (system === '方音符號') {
                // 使用 <rtc> 顯示方音符號
                let rtc = ruby.querySelector('rtc');
                if (!rtc) {
                    rtc = document.createElement('rtc');
                    ruby.appendChild(rtc);
                }
                rtc.textContent = converted;
                rtc.style.display = 'inline-flex';
                
                // 隱藏 <rt>
                const rt = ruby.querySelector('rt');
                if (rt) rt.style.display = 'none';
            } else {
                // 其餘標音一律使用 <rt>
                let rt = ruby.querySelector('rt');
                if (!rt) {
                    rt = document.createElement('rt');
                    const rtc = ruby.querySelector('rtc');
                    if (rtc) ruby.insertBefore(rt, rtc);
                    else ruby.appendChild(rt);
                }
                rt.textContent = converted;
                rt.style.display = 'block';
                
                // 隱藏 <rtc>
                const rtc = ruby.querySelector('rtc');
                if (rtc) rtc.style.display = 'none';
            }
        });
    }
});
