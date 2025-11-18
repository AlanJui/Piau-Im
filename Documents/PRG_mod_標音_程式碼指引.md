# 程式碼指引 mod_標音

## a200_查找及填入漢字標音


循序圖

```plantuml
@startuml
title a200.py L217–L223：查字並取得漢字標音

actor "a200_查找及填入漢字標音.py" as A200
participant "ca_ji_kiat_ko_tng_piau_im()" as CA
participant "PiauIm.han_ji_piau_im_tng_huan()" as HJT
participant "PiauIm.BP_piau_im()" as BP
participant "PiauIm._get_BP_syllable()" as GET

note over A200
  第 217–223 行：
  tai_gi_im_piau, han_ji_piau_im =
  ca_ji_kiat_ko_tng_piau_im(...)
end note

A200 -> CA : ca_ji_kiat_ko_tng_piau_im(result,…)
activate CA
CA -> HJT : han_ji_piau_im_tng_huan(piau_im_huat, siann_bu, un_bu, tiau_ho)
activate HJT
HJT -> BP : BP_piau_im(siann_bu, un_bu, tiau_ho)
activate BP
BP -> GET : _get_BP_syllable(..., with_tone_number=True)
activate GET
GET --> BP : 閩拼音節字串或 None
deactivate GET
BP --> HJT : 閩拼音節字串 (或空字串)
deactivate BP
HJT --> CA : han_ji_piau_im (若失敗為空字串)
deactivate HJT
CA --> A200 : (tai_gi_im_piau, han_ji_piau_im)
deactivate CA
@enduml

```

## 【mod_標音】模組

### def ca_ji_kiat_ko_tng_piau_im()

查字結果轉標音：利用【查漢字庫】所得之【台語音標】，於指明【漢字標音法】後，生成所欲引用之
【漢字標音】。

流程概述

- 依 han_ji_khoo 判定為「河洛話」或「文讀音」，分別取得聲母、韻母、調號；河洛話需進行韻母轉換並將調號 6→7。
- 將 siann_bu + un_bu + tiau_ho 合併成 tai_gi_im_piau。
- 若標音法為「十五音/雅俗通」且聲母為空，補為 ø。
- 呼叫 piau_im.han_ji_piau_im_tng_huan() 嘗試轉換，成功則 ok=True，否則記錄警告或例外並保持 ok=False。
- 最後依 ok 真假回傳 (tai_gi_im_piau, "") 或 (tai_gi_im_piau, han_ji_piau_im)。

```plantuml
@startuml
start
:輸入 result, han_ji_khoo, piau_im, piau_im_huat;

if (han_ji_khoo == "河洛話"?) then (是)
  :siann_bu = result[0]["聲母"];
  :un_bu = result[0]["韻母"];
  :un_bu = tai_gi_im_piau_tng_un_bu(un_bu);
  :tiau_ho = result[0]["聲調"];
  if (tiau_ho == "6"?) then (是)
    :tiau_ho = "7";
  endif
else
  :siann_bu, un_bu, tiau_ho = split_tai_gi_im_piau(result[0]["標音"]);
  if (siann_bu 為空?) then (是)
    :siann_bu = "ø";
  endif
endif

:tai_gi_im_piau = siann_bu + un_bu + tiau_ho;

if ((piau_im_huat in {"十五音","雅俗通"}) and siann_bu 為空?) then (是)
  :siann_bu = "ø";
endif

:ok = False;
:han_ji_piau_im = "";

:呼叫 piau_im.han_ji_piau_im_tng_huan(...);

if (呼叫過程發生例外?) then (是)
  :logging_exception(...);
  :han_ji_piau_im = "";
  :ok = False;
else (否)
  if (han_ji_piau_im 非空字串?) then (是)
    :ok = True;
  else
    :logging_warning(轉換失敗);
    :ok = False;
  endif
endif

if (ok == False?) then (是)
  :回傳 (tai_gi_im_piau, "");
else
  :回傳 (tai_gi_im_piau, han_ji_piau_im);
endif
stop
@enduml
```

---

### Piau_Im.han_ji_piau_im_tng_huan()

轉換【漢字標音】

流程概念

- 讀入 piau_im_huat、siann_bu、un_bu、tiau_ho。
- 依照 piau_im_huat 逐一判斷：十五音 / 方音符號 / 注音二式 / 雅俗通 / 白話字 / 台羅 / 閩拼調號 / 閩拼調符 / 台語音標。
- 命中對應分支就呼叫對應轉換函式並立刻回傳。
- 若是「台語音標」，先從字典取聲母韻母後組合並回傳。
- 若都不符合，最後回傳空字串。

```plantuml
@startuml
start
:輸入 piau_im_huat, siann_bu, un_bu, tiau_ho;

if (piau_im_huat == "十五音"?) then (是)
  :return SNI_piau_im(...);
  stop
endif

if (piau_im_huat == "方音符號"?) then (是)
  :return TPS_piau_im(...);
  stop
endif

if (piau_im_huat == "注音二式"?) then (是)
  :return MPS2_piau_im(...);
  stop
endif

if (piau_im_huat == "雅俗通"?) then (是)
  :return NST_piau_im(...);
  stop
endif

if (piau_im_huat == "白話字"?) then (是)
  :return POJ_piau_im(...);
  stop
endif

if (piau_im_huat == "台羅拼音"?) then (是)
  :return TL_piau_im(...);
  stop
endif

if (piau_im_huat == "閩拼調號"?) then (是)
  :return BP_piau_im(...);
  stop
endif

if (piau_im_huat == "閩拼調符"?) then (是)
  :return BP_piau_im_with_tiau_hu(...);
  stop
endif

if (piau_im_huat == "台語音標"?) then (是)
  :siann = Siann_Bu_Dict[siann_bu]["台語音標"] 或 "";
  if (siann in {"", None, "Ø", "ø"}?) then (是)
    :siann = "";
  endif
  :un = Un_Bu_Dict[un_bu]["台語音標"];
  :return siann + un + tiau_ho;
  stop
endif

:return "";
stop
@enduml
```

### Piau_Im.BP_piau_im()

呼叫內部轉換 → 檢查 None → 回傳字串

若 _get_BP_syllable() 在轉換成【漢字標音】過程，發生【執行時期錯誤】，
該函數會進行【意外處理(Exception)】，抑止程式因而中斷；但函數會返回 None 值。

外部函數可據此 None 值，判斷返回值，當看成（聲、韻、調）Tumple？還是轉換不成功
的結果。若遇轉換不成功，則回傳【空字串】以表無法依據傳入函數之參數：聲母、韻母、調號
及拚音系統名，轉換成【漢字標音】。

```plantuml
@startuml
start
:輸入 siann_bu, un_bu, tiau_ho;

:result = _get_BP_syllable(
  siann_bu, un_bu, tiau_ho,
  with_tone_number=True);

if (result 為 None?) then (是)
  :回傳 "";
  stop
endif

:回傳 result;
stop
@enduml
```

### Piau_Im._get_BP_syllable()

流程重點

- 先將上標調號轉成一般數字並嘗試轉成 int；若失敗則記錄警告並回傳 None。
- 依 Tiau_Ho_Remap 對映調號；若無法對映同樣記錄警告並回傳 None。
- 聲母為空 (""/None/Ø/ø) 則設為空字串；否則從 Siann_Bu_Dict 查表，若結果為空則回傳 ("", "", "")。
- 韻母從 Un_Bu_Dict 查表，查不到也回傳 ("", "", "")。
- 若為零聲母且韻母以 i 或 u 起頭，依規則改為 y/w 開頭。
- with_tone_number=True 時回傳單一字串（其他兩欄留空）；否則回傳 (siann, un, tiau)。

```plantuml
@startuml
start
:接收 siann_bu, un_bu, tiau_ho, with_tone_number;
:replace_superscript_digits(tiau_ho);

if (tiau_ho 可轉成整數?) then (否)
  :logging_warning(調號不可解析);
  stop
endif
if (tiau_ho == 6?) then (是)
  :tiau_ho_int = 7;
else
  :tiau_ho_int = int(tiau_ho);
endif

:tiau = Tiau_Ho_Remap[tiau_ho_int];
if (tiau 存在?) then (否)
  :logging_warning(調號無對映);
  stop
endif

if (siann_bu in {"", None, "Ø", "ø"}?) then (是)
  :siann = "";
else
  :siann = Siann_Bu_Dict[siann_bu]["閩拼方案"];
  if (siann 為空?) then (是)
    :logging_warning(聲母無對映);
    stop
  endif
endif

:un = Un_Bu_Dict[un_bu]["閩拼方案"];
if (un 為空?) then (是)
  :logging_warning(韻母無對映);
  stop
endif

if (siann == "") then (是)
  if (un 以 "i" 開頭?) then (是)
    if (len(un) == 1?) then (是)
      :un = "y" + un;
    else
      :un = "y" + un[1:];
    endif
  elseif (un 以 "u" 開頭?) then (是)
    if (len(un) == 1?) then (是)
      :un = "w" + un;
    else
      :un = "w" + un[1:];
    endif
  endif
endif

if (with_tone_number?) then (是)
  :syllable = siann + un + tiau;
  :回傳 (syllable, "", "");
else
  :回傳 (siann, un, tiau);
endif
stop
@enduml
```