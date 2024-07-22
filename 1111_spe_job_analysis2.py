import os
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from bs4 import BeautifulSoup
import pandas as pd
import jieba
from wordcloud import WordCloud
import matplotlib.pyplot as plt
from PIL import Image
import numpy as np
import time
import openpyxl

# 設置 Edge WebDriver 路徑，根據自己的環境調整
edge_driver_path = r'C:\path\to\edgedriver.exe'  # 請替換成你的Edge WebDriver路徑
# 設置瀏覽器選項
edge_options = webdriver.EdgeOptions()
# 輸入職稱並生成網址
job_title = input('請輸入想查詢的職稱：')
url = f"https://www.1111.com.tw/search/job?ks={job_title}"
# 設置日誌配置
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# 建立重試機制的最大次數和初始等待時間
max_retry = 3
retry_delay = 5  # seconds
# 初始化重試計數器
retry_count = 0
success = False

while retry_count < max_retry and not success:
    try:
        # 檢查 stopWord.txt 檔案是否存在
        stopfile = "stopWord.txt"  # 停用詞檔案的路徑
        if not os.path.exists(stopfile):
            raise FileNotFoundError(f"{stopfile} 文件不存在或路徑錯誤")
        # 讀取停用詞檔案
        with open(stopfile, 'r', encoding='utf-8') as f:
            stopwords = [line.strip() for line in f.readlines() if line.strip()]
        # 建立 WebDriver 對象，使用 Edge
        driver = webdriver.Edge(options=edge_options)
        logger.info("Edge WebDriver 已啟動")
        # 打開網頁
        driver.get(url)
        time.sleep(10)  # 等待網頁加載完全
        # 使用鍵盤向下按鈕捲動瀏覽器
        actions = ActionChains(driver)
        elem = driver.find_element(By.CLASS_NAME, "btnApply")
        actions.move_to_element(elem).perform()
        time.sleep(2)
        for _ in range(500):
            elem.send_keys(Keys.ARROW_DOWN)
            time.sleep(0.1)
        # 取得網頁原始碼
        html_source = driver.page_source
        soup = BeautifulSoup(html_source, "html.parser")
        # 找到職缺和公司名稱
        jobs = soup.find_all("div", class_="item__job job-item card work")[:100]
        # 創建 DataFrame 存儲職缺和公司
        data = []
        for job in jobs:
            job_title = job.find("div", class_="title position0").find("a").text.strip()
            company_name = job.find("div", class_="company organ").find("a").text.strip()
            data.append({"職缺": job_title, "公司": company_name})
        df = pd.DataFrame(data)
        # 存成 Excel
        excel_file = "1111_spe.xlsx"
        df.to_excel(excel_file, sheet_name="1111_spe_職缺分析", index=False)
        logger.info(f"Excel 檔案已保存：{excel_file}")
        # 關閉瀏覽器
        driver.quit()
        logger.info("Edge WebDriver 已關閉")

        # 文字雲生成部分
        dictfile = "dict.txt.big.txt"  # 設定常用字典
        fontpath = "TaipeiSansTCBeta-Regular.ttf"  # 繁體中文字體檔案的路徑
        pngfile = "cloud.png"  # 自訂文字雲遮罩
        mdfile = excel_file  # 資料來源檔案的路徑
        # 使用 PIL 加載遮罩圖片並轉換為 numpy 數組
        bgmask = np.array(Image.open(pngfile))
        # 讀取 Excel 資料
        df = pd.read_excel(mdfile, sheet_name='1111_spe_職缺分析')
        text = ' '.join(df['公司'].astype(str))
        # 創建 WordCloud 對象，設置參數
        wc = WordCloud(font_path=fontpath,
                       background_color="white",  # 背景顏色
                       mask=bgmask,  # 遮罩圖像
                       stopwords=stopwords,  # 停用詞列表
                       contour_color='steelblue',  # 邊框顏色
                       contour_width=5,  # 邊框寬度
                       max_words=150,  # 最大詞數
                       max_font_size=100,  # 最大字體大小
                       collocations=False  # 是否包括雙詞組合
                       )
        # 使用 jieba 分詞，並統計詞頻
        jieba.set_dictionary(dictfile)
        word_list = jieba.lcut(text, cut_all=False)
        segmented_text = ' '.join(word_list)
        # 生成文字雲
        wc.generate(segmented_text)
        # 顯示文字雲
        plt.figure(figsize=(12, 12))
        plt.imshow(wc, interpolation='bilinear')
        plt.axis('off')
        plt.tight_layout()
        # 保存文字雲圖片
        outputfile = '1111_spe.png'
        plt.savefig(outputfile)
        # 顯示生成的文字雲圖片
        plt.show()
        logger.info(f"文字雲已保存至 {outputfile}")
        success = True  # 操作成功完成

    except FileNotFoundError as fnf_error:
        logger.error(f"第 {retry_count + 1} 次嘗試時出現錯誤：{fnf_error}")
        retry_count += 1
        time.sleep(retry_delay)
    except Exception as e:
        logger.error(f"第 {retry_count + 1} 次嘗試時出現錯誤：{e}")
        retry_count += 1
        time.sleep(retry_delay)
        
# 如果達到最大重試次數仍未成功，顯示錯誤訊息
if not success:
    logger.error(f"網頁加載失敗，已嘗試 {max_retry} 次仍未成功。")
