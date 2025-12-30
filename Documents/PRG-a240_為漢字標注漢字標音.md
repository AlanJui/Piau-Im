# 程式說明文件 a240 為漢字標注漢字標音

## 主流程

```plantuml
@startuml
title a240_為漢字標注漢字標音.py 主流程

start
:載入 env / logging / 模組;
:main() 取得程式資訊與路徑;
:取得 Excel 作用中活頁簿 (xw.apps.active.books.active);
if (成功取得?) then (是)
  :status_code = process(wb);
  if (status_code == EXIT_CODE_SUCCESS?) then (是)
    :wb.sheets["漢字注音"].activate();
    :file_path = save_as_new_file(wb);
    if (file_path 存在?) then (是)
      :logging 成功儲存;
    else (否)
      :logging 儲存失敗;
      stop
    endif
  else (否)
    :logging_exc_error(程式異常終止);
    stop
  endif
else (否)
  :logging.error(無作用中活頁簿);
  stop
endif
:logging_process_step(程式終止);
stop
@enduml
```

## han_ji_piau_im()

### 主流程

- 初始化：讀取命名範圍、建立 PiauIm、鎖定處理範圍與旗標。
- 逐列逐欄：檢查結束標記/換行/非漢字，或拆解台語音標後寫回漢字標音並印出訊息。
- 結束處理：離開巢狀迴圈、回到表頭、寫入日誌，回傳 EXIT_CODE_SUCCESS。

```plantuml
@startuml
title han_ji_piau_im(wb, sheet_name="漢字注音")

start
:讀取 named range (漢字庫/語音類型/標音方法);
:建立 PiauIm 並確保工作表存在;
:計算 TOTAL_LINES / CHARS_PER_ROW / ROWS_PER_LINE;
:設定 start_row/start_col 與 EOF=False, line=1;

repeat
  :設定作用儲存格至列首;
  :col = start_col;

  repeat
    :讀取 tai_gi_cell / han_ji_cell / han_ji_piau_im_cell;
    if (han_ji_cell == 'φ'?) then (是)
      :EOF = True;
      :msg = 《文章終止》;
      break
    elseif (han_ji_cell == '\n'?) then (是)
      :msg = 《換行》;
      break
    elseif (not is_han_ji(han_ji_cell)) then (是)
      :msg = 直接顯示字元;
    else
      if (tai_gi_cell 有值?) then (是)
        :siann_bu, un_bu, tiau_ho = split_tai_gi_im_piau(tai_gi_cell);
        :han_ji_piau_im = piau_im.han_ji_piau_im_tng_huan(...);
        :han_ji_piau_im_cell = han_ji_piau_im;
        :msg = "漢字 [TLPA] / [漢字標音]";
      else
        :msg = "無台語音標";
      endif
    endif
    :print((row, col) = msg);
    :col += 1;
  repeat while (col < end_col and not break)

  :print((row, col) = msg);
  :row += ROWS_PER_LINE;
  :line += 1;
repeat while (not EOF and line <= TOTAL_LINES)

:han_ji_piau_im_sheet.activate();
:select A1;
:logging_process_step(作業完成與參數資訊);
:return EXIT_CODE_SUCCESS;
stop
@enduml
```

### 子流程：處理單一漢字

「處理單一漢字」子流程，專注在 han_ji_piau_im() 內針對每個 (row, col) 位置的判斷與動作。

- 讀出對應的台語音標儲存格、漢字儲存格、漢字標音儲存格。
- 若遇 φ、換行或非漢字，直接設定訊息並結束當前欄位。
- 若為漢字且有台語音標，則拆解聲母/韻母/調號，呼叫 piau_im.han_ji_piau_im_tng_huan() 取得漢字標音並寫回儲存格。
- 所有分支都會留下 msg，最後印出 (row, col) 與 msg。

```plantuml
@startuml
title han_ji_piau_im()：單一漢字子流程

start
:讀取 tai_gi_cell, han_ji_cell, han_ji_piau_im_cell;
if (han_ji_cell == 'φ'?) then (是)
  :EOF = True;
  :msg = 《文章終止》;
  stop
elseif (han_ji_cell == '\n'?) then (是)
  :msg = 《換行》;
  stop
elseif (not is_han_ji(han_ji_cell)) then (是)
  :msg = han_ji_cell 內容;
else
  if (tai_gi_cell 有值?) then (是)
    :siann_bu, un_bu, tiau_ho = split_tai_gi_im_piau(tai_gi_cell);
    :han_ji_piau_im = piau_im.han_ji_piau_im_tng_huan(...);
    :han_ji_piau_im_cell = han_ji_piau_im;
    :tlpa_im_piau = siann_bu + un_bu + tiau_ho;
    :msg = "漢字 [TLPA] / [漢字標音]";
  else (否)
    :msg = "無台語音標";
  endif
endif
:print((row, col) = msg);
stop
@enduml
```