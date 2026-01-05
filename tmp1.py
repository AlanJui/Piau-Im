def process(wb):
    """
    更新【漢字注音】表中【台語音標】儲存格的內容，依據【標音字庫】中的【校正音標】欄位進行更新，並將【校正音標】覆蓋至原【台語音標】。
    """
    logging_process_step("<=========== 開始處理流程作業！==========>")
    try:
        # 取得工作表
        target_sheet_name = "漢字注音"
        ensure_sheet_exists(wb, target_sheet_name)
        han_ji_piau_im_sheet = wb.sheets["漢字注音"]
        han_ji_piau_im_sheet.activate()
    except Exception as e:
        logging_exc_error(msg=f"找不到【漢字注音】工作表 ！", error=e)
        return EXIT_CODE_PROCESS_FAILURE
    logging_process_step(f"已完成作業所需之初始化設定！")

    # -------------------------------------------------------------------------
    # 將【缺字表】工作表，已填入【台語音標】之資料，登錄至【標音字庫】工作表
    # 使用【缺字表】工作表中的【校正音標】，更正【漢字注音】工作表中之【台語音標】、【漢字標音】；
    # 並依【缺字表】工作表中的【台語音標】儲存格內容，更新【標音字庫】工作表中之【台語音標】及【校正音標】欄位
    # -------------------------------------------------------------------------
    try:
        sheet_name = "缺字表"
        logging_process_step(f"以【{sheet_name}】工作表之【校正音標】，更新【{target_sheet_name}】工作表之【台語音標】與【漢字標音】！")
        print("\n\n")
        print("=" * 100)
        print(f"使用【{sheet_name}】工作表的【校正音標】欄位，更新【{target_sheet_name}】工作表之【台語音標】、【漢字標音】：")
        print("=" * 100)
        # 將【缺字表】工作表中的【台語音標】儲存格內容，更新至【標音字庫】工作表中之【台語音標】及【校正音標】欄位
        # update_khuat_ji_piau(wb=wb)
        # 依據【缺字表】工作表紀錄，並參考【漢字注音】工作表在【人工標音】欄位的內容，更新【缺字表】工作表中的【校正音標】及【台語音標】欄位
        # 即使用者為【漢字】補入查找不到的【台語音標】時，若是在【缺字表】工作表中之【校正音標】直接填寫
        # 則應執行 a310*.py 程式；但使用者若是在【漢字注音】工作表中之【人工標音】欄位填寫，則應執行 a320*.py 程式
        # a300*.py 之本程式
        update_khuat_ji_piau_by_jin_kang_piau_im(wb=wb, sheet_name=sheet_name)
    except Exception as e:
        logging_exc_error(msg=f"處理【缺字表】作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE
    # -------------------------------------------------------------------------
    # 將【漢字注音】工作表，【漢字】填入【人工標音】內容，登錄至【人工標音字庫】及【標音字庫】工作表
    # -------------------------------------------------------------------------
    try:
        sheet_name = "人工標音字庫"
        logging_process_step(f"以【{sheet_name}】工作表之【校正音標】，更新【{target_sheet_name}】工作表之【台語音標】與【漢字標音】！")
        print("\n\n")
        print("=" * 100)
        print(f"使用【{sheet_name}】工作表的【校正音標】欄位，更新【{target_sheet_name}】工作表之【台語音標】、【漢字標音】：")
        print("=" * 100)
        update_by_jin_kang_piau_im(wb=wb, sheet_name=sheet_name)
    except Exception as e:
        logging_exc_error(msg=f"處理【漢字】之【人工標音】作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE
    # -------------------------------------------------------------------------
    # 根據【標音字庫】工作表，更新【漢字注音】工作表中的【台語音標】及【漢字標音】欄位
    # -------------------------------------------------------------------------
    try:
        sheet_name = "標音字庫"
        logging_process_step(f"以【{sheet_name}】工作表之【校正音標】，更新【{target_sheet_name}】工作表之【台語音標】與【漢字標音】！")
        print("\n\n")
        print("=" * 100)
        print(f"使用【{sheet_name}】工作表中的【校正音標】，更新【漢字注音】工作表中的【台語音標】：")
        print("=" * 100)
        update_by_piau_im_ji_khoo(wb, sheet_name=sheet_name)
    except Exception as e:
        logging_exc_error(msg=f"處理以【標音字庫】更新【漢字注音】工作表之作業，發生執行異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE
    # --------------------------------------------------------------------------
    # 結束作業
    # --------------------------------------------------------------------------
    han_ji_piau_im_sheet.activate()
    logging_process_step("<=========== 完成處理流程作業！==========>")

    return EXIT_CODE_SUCCESS
