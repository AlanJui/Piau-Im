# 程式說明文件 a320 人工標音更正漢字自動標音

## process()

流程總覽

- process()：初始化 → 呼叫 jin_kang_piau_imm() 生成台語/漢字標音 → 記錄完成並返回成功碼。
- jin_kang_piau_imm()：建置字庫 → 逐列逐欄判斷 φ/換行/標點/空白 → 對有人工標音的漢字轉換 TLPA 與漢字標音、同步到兩個字庫 → 寫回 Excel 並記錄進度。


```plantuml
@startuml
title a320.process()

start
:logging "<----------- 作業開始 ---------->";
:han_ji_khoo_name = named range "漢字庫";
:ue_im_lui_piat = named range "語音類型";
:logging 取得設定;

:status = jin_kang_piau_imm(
  sheet="漢字注音",
  ue_im_lui_piat,
  han_ji_khoo=han_ji_khoo_name);

if (status == EXIT_CODE_SUCCESS?) then (否)
  :logging_exc_error("查找漢字標音時發生錯誤");
  stop
endif

:logging "自動查找台語音標完成";
:logging 語音類型資訊;
:wb.sheets["漢字注音"].activate();
:logging "<----------- 作業結束 ---------->";
:return EXIT_CODE_SUCCESS;
stop
@enduml
```


## jin_kang_piau_im()



```plantuml
@startuml
title a320.jin_kang_piau_imm()

start
:讀取漢字庫與標音方法;
:piau_im = PiauIm(...);
:piau_im_ji_khoo = JiKhooDict("標音字庫");
:jin_kang_ji_khoo = JiKhooDict("人工標音字庫");
:設定 TOTAL_LINES / ROWS_PER_LINE / CHARS_PER_ROW;
:EOF=False; line=1;

repeat
  :Two_Empty_Cells=0;
  :sheet.range((row,1)).select();
  repeat
    :cell = sheet[row, col];
    :reset cell font/color;
    :msg="";
    :jin_kang = cell.offset(-2,0).value;

    if (cell == 'φ'?) then (是)
      :EOF=True; msg="【文字終結】";
      break
    elseif (cell == '\n'?) then (是)
      :msg="【換行】";
      break
    elseif (not is_han_ji(cell)) then (是)
      if (is_punctuation(cell)?) then (是)
        :msg="標點符號";
      elseif (cell 為整數字串?) then (是)
        :轉成整數字串; msg="英/數半形";
      else
        :Two_Empty_Cells 累計; 若達 2 設 EOF=True;
        :msg="【空白】";
      endif
    else
      if (jin_kang 存在且漢字非空?) then (是)
        :tai_gi, han_ji_piau = jin_kang_piau_im_cu_han_ji_piau_im(...);
        :sheet[row-1,col] = tai_gi;
        :sheet[row+1,col] = han_ji_piau;
        :msg = "漢字 [台語音標]/[標音]《人工標音》";
        :jin_kang_ji_khoo.add_entry(...);
        :existing_entries = piau_im_ji_khoo[han_ji];
        :從 existing_entries 移除 (row,col) 座標;
        :piau_im_ji_khoo.add_entry(..., kenn_ziann="N/A");
        :cell.font.color = red; cell.color = yellow;
      endif
    endif

    :print 進度;
  repeat while (col < end_col and not break)

  if (EOF or line > TOTAL_LINES?) then (是)
    break
  endif
  :line += 1;
repeat while (not EOF)

:piau_im_ji_khoo.write_to_excel_sheet(...);
:jin_kang_ji_khoo.write_to_excel_sheet(...);
:logging "已完成台語音標/漢字標音標注";
:return EXIT_CODE_SUCCESS;
stop
@enduml
```


## 非漢字處理

```plantuml
@startuml
title 子流程：非漢字處理

start
:cell_value = 當前儲存格;
if (is_punctuation(cell_value)?) then (是)
  :msg = cell_value + "【標點符號】";
else
  if (cell_value 為整數型 float?) then (是)
    :cell_value = str(int(cell_value));
    :msg = cell_value + "【英/數半形字元】";
  else
    if (cell_value 為空或僅空白?) then (是)
      if (Two_Empty_Cells == 0?) then (是)
        :Two_Empty_Cells += 1;
      else
        :Two_Empty_Cells += 1;
        :EOF = True;
      endif
      :msg = "【空白】";
    else
      :msg = f"{cell_value}【未定義非漢字】";
    endif
  endif
endif
:print 進度;
stop
@enduml
```


## 同步人工/標音字庫

```plantuml
@startuml
title 子流程：同步人工/標音字庫

start
:'[tai_gi, han_ji_piau] =
  jin_kang_piau_im_cu_han_ji_piau_im(...)';
:sheet[row-1,col] = tai_gi;
:sheet[row+1,col] = han_ji_piau;

:jin_kang_ji_khoo.add_entry(
  han_ji, tai_gi, jin_kang_piau_im, (row,col));

:existing_entries = piau_im_ji_khoo[han_ji];
if (existing_entries 含 (row,col)?) then (是)
  :existing_entry["coordinates"].remove((row,col));
endif

:piau_im_ji_khoo.add_entry(
  han_ji, tai_gi, "N/A", (row,col));

:cell.font.color = 紅色;
:cell.color = 黃色;
:msg += "【人工標音】";
:print 進度;
stop
@enduml
```