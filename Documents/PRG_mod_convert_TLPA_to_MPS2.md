# mod_convert_TLPA_to_MPS2 程式說明文件

將文字檔中之【台語音標】轉換成【台語注音二式音標】。

## 流程圖

### 主要流程圖

```mermaid
flowchart TD
    A[啟動程式] --> B[讀取輸入檔案]
    B --> C[逐行處理檔案內容]
    C --> D{是否進入詞條區?}
    D -- 否 --> C
    D -- 是 --> E{是否為空行或註解?}
    E -- 是 --> C
    E -- 否 --> F[分割行為多欄]
    F --> G{是否有至少兩欄?}
    G -- 否 --> C
    G -- 是 --> H[呼叫 convert_TLPA_to_MPS2]
    H --> I[更新第二欄內容]
    I --> C
    C --> J[寫入輸出檔案]
    J --> K[結束程式]
```

### convert_TLPA_to_MPS2() 流程圖

```mermaid
flowchart TD
    A[接收 TLPA 音標字串] --> B{是否符合格式?}
    B -- 否 --> C[直接回傳原字串]
    B -- 是 --> D[分割為聲母與韻母+聲調]
    D --> E[比對聲母映射表]
    E --> F{是否找到對應聲母?}
    F -- 否 --> G[聲母保持不變]
    F -- 是 --> H[替換為對應聲母]
    G --> I[比對韻母映射表]
    H --> I
    I --> J{韻母是否在映射表中?}
    J -- 否 --> K{韻母是否以 o 結尾？}
    K -- 是 --> L[替換 o 為 or]
    K -- 否 --> M[保持韻母不變]
    J -- 是 --> N[替換為對應韻母]
    L --> O[合併聲母、韻母與聲調]
    M --> O
    N --> O
    O --> P[回傳轉換後的 MPS2 字串]
```

## 台語音標【調號】轉成【台語二式音標】

```plantuml
@startjson
title 台語音標聲調對照表
{
    "1": "陰平",
    "2": "上聲",
    "3": "陰去",
    "4": "陰入",
    "5": "陽平",
    "6": "上聲",
    "7": "陽去",
    "8": "陽入"
}
@endjson
```

```plantuml
@startjson
title 閩拼音標聲調對照表
{
    "1": "陰平",
    "2": "陽平",
    "3": "上聲",
    "4": "上聲",
    "5": "陰去",
    "6": "陽去",
    "7": "陰入",
    "8": "陽入"
}
@endjson
```

## PlantUML 活動圖

想研究 PlantUML 在【流程圖】繪製之便利性。

### 主流程

```plantuml
@startuml
start
:讀取輸入檔案;
:逐行處理檔案內容;
repeat
  :是否進入詞條區?;
  if (是) then (是)
    if (是否為空行或註解?) then (是)
      :跳過該行;
    else (否)
      :分割行為多欄;
      if (是否有至少兩欄?) then (是)
        :呼叫 convert_TLPA_to_MPS2;
        :更新第二欄內容;
      else (否)
        :跳過該行;
      endif
    endif
  endif
repeat while (還有未處理的行?)
:寫入輸出檔案;
stop
@enduml
```

### 模組：convert_TLPA_to_MPS2

```plantuml
@startuml
start
:接收 TLPA 音標字串;
if (是否符合格式?) then (否)
  #pink:直接回傳原字串;
  stop
else (是)
  :分割為聲母與韻母+聲調;
  :比對聲母映射表;
  if (是否找到對應聲母?) then (否)
    #pink:聲母保持不變;
  else (是)
    :替換為對應聲母;
  endif
  :比對韻母映射表;
  if (韻母是否在映射表中?) then (否)
    if (<color:red>韻母是否以 "o" 結尾?) then (是)
      :替換 "o" 為 "or";
    else (否)
      #pink:保持韻母不變;
    endif
  else (是)
    :替換為對應韻母;
  endif
  :合併聲母、韻母與聲調;
  :回傳轉換後的 MPS2 字串;
  stop
@enduml
```

