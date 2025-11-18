# 程式說明文件 a300 手動更正漢字標音

## process()

- 初始化：記錄開始、擷取 漢字庫 與 標音方法、建立 PiauIm、啟用 漢字注音 工作表；任一失敗直接回傳 EXIT_CODE_PROCESS_FAILURE。
- 缺字表補標音：呼叫 update_khuat_ji_piau_by_jin_kang_piau_im()，將缺字表資料同步到標音字庫；失敗即終止。
- 人工標音回填：呼叫 update_by_jin_kang_piau_im()，把「漢字注音」表的人工標音寫入字庫並反映在工作表。
- 校正音標覆寫：呼叫 update_by_piau_im_ji_khoo()，以標音字庫內的校正音標覆蓋「漢字注音」表的台語音標/漢字標音。
- 收尾：選取 A1、記錄結束訊息、回傳 EXIT_CODE_SUCCESS。

```plantuml
@startuml
title a300.process()

start
:logging "<----------- 作業開始 ---------->";
:取得漢字庫/標音方法;
:create PiauIm & activate 漢字注音;
if (初始化成功?) then (否)
  :logging_exc_error;
  stop
endif

:'缺字表'流程 = update_khuat_ji_piau_by_jin_kang_piau_im();
if (成功?) then (否)
  :logging_exc_error;
  stop
endif
:logging "缺字表已補入標音字庫";

:'人工標音'流程 = update_by_jin_kang_piau_im();
if (成功?) then (否)
  :logging_exc_error;
  stop
endif
:logging "人工標音已寫入字庫";

:'標音字庫回填'流程 = update_by_piau_im_ji_khoo();
if (成功?) then (否)
  :logging_exc_error;
  stop
endif
:logging "校正音標已更新漢字注音";

:漢字注音範圍跳回 A1;
:logging "<----------- 作業結束 ---------->";
:return EXIT_CODE_SUCCESS;
stop
@enduml
```


## update_by_jin_kang_piau_im()

- 初始化：鎖定「漢字注音」表、讀取 漢字庫／標音方法，建立 PiauIm。建立「人工標音字庫」「標音字庫」兩個 JiKhooDict。
- 設定表格範圍並進入雙層迴圈（行→列），逐一取出漢字、台語音標、人工標音儲存格，並先還原底色字色。
- 結束條件：遇 φ 設 EOF=True，遇 \n 視為換行跳出本列。對非漢字再判斷標點 / 英數 / 空白並計數空白行。
- 若為漢字：
- 無人工標音：若儲存格仍有舊值則清空並還原樣式。
- 有人工標音：標註底色，透過 jin_kang_piau_im_cu_han_ji_piau_im() 轉換成台語音標 + 漢字標音，寫回儲存格。
- 同步字庫：jinkang 字庫記錄 add_or_update_entry()；piau_im_ji_khoo.update_kau_ziang_im_piau() 更新校正音標。
- 內層每欄列印進度，外層每列換行，遇 EOF 或超過總行數即跳出。
- 收尾：把更新後的兩個字典寫回各自工作表，選取 A1，回傳成功。

```plantuml
@startuml
title update_by_jin_kang_piau_im()

start
:確保「漢字注音」表存在並讀取設定;
:建立 PiauIm 及讀取 標音方法;
:載入「人工標音字庫」「標音字庫」為 JiKhooDict;
:計算 start_row/start_col 等範圍;
:EOF=False, line=1;

repeat
  :activate 漢字注音並定位列首;
  :Empty_Cells_Total=0;
  repeat
    :抓取 han_ji/tai_gi/jin_kang cell;
    :reset cell 顏色 & 字型;
    if (han_ji == 'φ'?) then (是)
      :EOF=True; msg="《文章終止》";
      break
    elseif (han_ji == '\n'?) then (是)
      :msg="《換行》";
      break
    elseif (not is_han_ji(han_ji)) then (是)
      :判斷標點/半形/空白 → msg;
      if (連續空白達 2?) then (是)
        :EOF=True;
      endif
    else
      if (jin_kang_piau_im 為空?) then (是)
        :必要時清空舊人工標音並還原顏色;
        :msg+="無人工標音";
      else
        :標記底色=黃、字紅;
        :kenn_ziann, han_ji_piau = jin_kang_piau_im_cu_han_ji_piau_im(...);
        :tai_gi_cell = kenn_ziann;
        :han_ji_piau_im_cell = han_ji_piau;
        :人工字庫.add_or_update_entry(...);
        :標音字庫.update_kau_ziang_im_piau(...);
        :msg+="[TLPA]/[漢字標音]【人工標音】";
      endif
    endif
    :print 目前欄位 msg;
  repeat while (col < end_col and not break)

  :print 空行區隔;
  :line += 1;
repeat while (not EOF and line <= TOTAL_LINES)

:標音字庫/人工字庫寫回工作表;
:漢字注音 activate 並選 A1;
:return EXIT_CODE_SUCCESS;
stop
@enduml
```