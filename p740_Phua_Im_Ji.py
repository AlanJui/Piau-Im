# p740_Phua_Im_Ji.py   v0.0.0.1

import json
import pickle


class PhuaImJi:
    """
    注音管理器，用於管理漢字與注音對應關係。
    """

    def __init__(self, ji_tian_name="Phua_Im_Ji.json"):
        # 【人工注音字典】存放漢字與注音
        self.phua_im_ji_tian = {}
        self.Ji_Tian_Name = ji_tian_name


    def ka_phua_im_ji(self, han_ji, piau_im):
        """
        將【漢字】與【注音】加入【人工注音字典】。

        參數：
        - han_ji: str，單一漢字。
        - piau_im: str，注音符號。
        """
        if len(han_ji) != 1:
            raise ValueError("輸入的 char 必須是一個單一漢字。")
        self.phua_im_ji_tian[han_ji] = piau_im
        print(f"漢字：【{han_ji}】之注音【{piau_im}】已加入【人工注音字典】。")


    def ca_phua_im_ji(self, han_ji):
        """
        查找漢字是否存在注音，返回【注音】字串或 None。

        參數：
        - han_ji: str，單一漢字。

        返回：
        - str 或 None：若存在於【人工注音字典】則返回注音，否則返回 None。
        """
        return self.phua_im_ji_tian.get(han_ji, None)


    def save_to_file(self):
        """
        將【人工注音字典】以純文字 JSON 格式存入檔案。
        """
        file_path = self.Ji_Tian_Name
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(self.phua_im_ji_tian, f, ensure_ascii=False, indent=4)
        print(f"【人工注音字典】已儲存至 {file_path}")


    def load_from_file(self):
        """
        從 JSON 格式檔案讀取【人工注音字典】。
        """
        file_path = self.Ji_Tian_Name
        with open(file_path, 'r', encoding='utf-8') as f:
            self.phua_im_ji_tian = json.load(f)
        print(f"已從 {file_path} 載入【人工注音字典】")


    def dump_phua_im_ji_tian(self):
        """
        在螢幕上輸出【人工注音字典】的內容，以純文字格式顯示。
        """
        if not self.phua_im_ji_tian:
            print("【人工注音字典】為空。")
        else:
            print("【人工注音字典】內容如下：")
            print("{")
            for han_ji, piau_im in self.phua_im_ji_tian.items():
                print(f"  '{han_ji}': '{piau_im}',")
            print("}")


    def save_to_bin_file(self, file_path):
        """
        將【人工注音字典】存入檔案。
        """
        with open(file_path, 'wb') as f:
            pickle.dump(self.phua_im_ji_tian, f)
        print(f"【人工注音字典】已儲存至 {file_path}")


    def load_from_bin_file(self, file_path):
        """
        從檔案讀取【人工注音字典】。
        """
        with open(file_path, 'rb') as f:
            self.phua_im_ji_tian = pickle.load(f)
        print(f"已從 {file_path} 載入【人工注音字典】")


# 單元測試
if __name__ == "__main__":
    import pickle

    # phua_im_ji = PhuaImJi()
    phua_im_ji = PhuaImJi('tmp.json')

    # 手動加入破音
    phua_im_ji.ka_phua_im_ji("行", "ㄒㄧㄥˊ")
    phua_im_ji.ka_phua_im_ji("重", "ㄓㄨㄥˋ")

    #==========================================================================

    # 儲存成 JSON 純文字檔案
    phua_im_ji.save_to_file()

    # 清空內部資料結構，模擬重新載入
    phua_im_ji.phua_im_ji_tian.clear()

    # 從 JSON 檔案讀取
    phua_im_ji.load_from_file()

    # 在螢幕上輸出【人工注音字典】
    phua_im_ji.dump_phua_im_ji_tian()

    # 查詢注音
    print(phua_im_ji.ca_phua_im_ji("行"))  # 返回 "ㄒㄧㄥˊ"
    print(phua_im_ji.ca_phua_im_ji("重"))  # 返回 "ㄓㄨㄥˋ"
    print(phua_im_ji.ca_phua_im_ji("新"))  # 返回 None

    #==========================================================================
    # BIN 字典資料
    data = {"行": "ㄒㄧㄥˊ", "重": "ㄓㄨㄥˋ"}

    # 儲存至檔案
    with open("data.pkl", "wb") as f:
        pickle.dump(data, f)
    print("資料已儲存至 data.pkl")

    # 從檔案讀取
    with open("data.pkl", "rb") as f:
        loaded_data = pickle.load(f)
    print("已載入資料：", loaded_data)