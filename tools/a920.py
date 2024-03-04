import os

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from dotenv import load_dotenv

# Get the path to he directory this file is in
BASEDIR = os.path.abspath(os.path.dirname(__file__))
# Load environment variables
load_dotenv(os.path.join(BASEDIR, 'config.env'))
CHROMEDRIVER_PATH = os.getenv('CHROMEDRIVER_PATH')

# from config_dev_env import CHROMEDRIVER_PATH, WAIT_TIME, KONG_UN_DICT_URL
KONG_UN_DICT_URL = 'https://ctext.org/dictionary.pl?if=gb'
WAIT_TIME = 5  # seconds

service = Service(executable_path=CHROMEDRIVER_PATH)
driver = webdriver.Chrome(service=service)

def fetch_guangyun_info(character):
    driver.get(KONG_UN_DICT_URL)

    # 等待搜索框出现
    search_box = WebDriverWait(driver, WAIT_TIME).until(
        EC.visibility_of_element_located((By.NAME, "char"))
    )
    search_box.clear()
    search_box.send_keys(character)
    search_box.submit()

    # 等待搜索结果加载
    WebDriverWait(driver, WAIT_TIME).until(
        EC.visibility_of_element_located((By.ID, "content"))
    )

    # 解析查询结果
    result = []
    # 定位到包含切语信息的td元素
    info_td = driver.find_element(By.CSS_SELECTOR, "#content table.info tr:nth-last-child(2) td")
    links = info_td.find_elements(By.TAG_NAME, "a")
    # 每四个链接为一组，对应一个切语信息
    for i in range(0, len(links), 5):
        # 解析切语、调、韵系和切语下字
        tshiat_gu = links[i].text if i < len(links) else ''
        tiau = links[i+2].text if i+2 < len(links) else ''
        un_he = links[i+3].text if i+3 < len(links) else ''
        tshia_gu_ha_ji = links[i+4].text if i+4 < len(links) else ''
        result.append({
            "tshiat_gu": tshiat_gu,  # 切语
            "tiau": tiau,            # 调
            "un_he": un_he,          # 韵系
            "tshia_gu_ha_ji": tshia_gu_ha_ji,  # 切语下字
        })

    driver.quit()
    return result

# 示例：查询字符"无"的信息
guangyun_info = fetch_guangyun_info("無")
print(guangyun_info)

# guangyun_info = fetch_guangyun_info("不")
# print(guangyun_info)
