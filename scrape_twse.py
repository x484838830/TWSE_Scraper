import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime

def parse_years(year_string):
    """解析使用者輸入的年份字串，轉換為年份列表"""
    years = set()
    parts = year_string.replace('、', ',').split(',')
    for part in parts:
        part = part.strip()
        if '~' in part:
            start_year, end_year = part.split('~')
            start_year, end_year = int(start_year), int(end_year)
            for year in range(start_year, end_year + 1):
                years.add(year)
        elif part.isdigit():
            years.add(int(part))
    
    # 驗證年份並排序
    valid_years = sorted([y for y in years if y >= 100])
    return valid_years


def get_user_inputs():
    """引導使用者輸入年份和公司代號"""
    # 取得年份輸入
    print("--- 股東會資料查詢設定 ---")
    print("請輸入您想查詢的民國年份(需>=100年)，支援以下格式：")
    print("1. 單一年份可以這樣輸入: 110")
    print("2. 需要不同年份(不連續)可以這樣輸入: 109、111、113 (可用全形、或半形逗號分隔)")
    print("3. 連續年份可以這樣輸入: 109~111")
    
    while True:
        year_input_str = input("請輸入查詢年份: ")
        try:
            year_list = parse_years(year_input_str)
            if not year_list:
                print("錯誤：未輸入有效年份或年份小於100，請重新輸入。")
                continue
            print(f"將會查詢以下年份: {year_list}")
            break
        except ValueError:
            print("錯誤：輸入格式不正確，請依照提示格式輸入。")

    # 取得公司代號輸入
    print("\nNote: 您可以指定單一公司，或抓取所有公司。")
    company_code = input("請輸入公司代號 (如 2330)，或直接按 Enter 抓取全部: ").strip()

    return year_input_str, year_list, company_code

def main():
    """主執行程式"""
    year_input_str, year_list, company_code = get_user_inputs()
    
    driver = None
    all_data = []
    company_name = ""

    try:
        print("\n正在啟動瀏覽器...")
        options = webdriver.ChromeOptions()
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        options.add_argument('--log-level=3')
        
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        
        for year in year_list:
            print(f"\n--- 正在查詢 民國 {year} 年的資料 ---")
            driver.get("https://mopsov.twse.com.tw/mops/web/t108sb31_q1")
            
            wait = WebDriverWait(driver, 10)
            year_input_element = wait.until(EC.presence_of_element_located((By.ID, "YEAR")))
            
            year_input_element.clear()
            year_input_element.send_keys(str(year))
            
            if company_code:
                # 填寫「公司代號 起」
                co_id1_input = driver.find_element(By.ID, "co_id1")
                co_id1_input.clear()
                co_id1_input.send_keys(company_code)
                
                # 同時填寫「公司代號 迄」
                co_id2_input = driver.find_element(By.ID, "co_id2")
                co_id2_input.clear()
                co_id2_input.send_keys(company_code)
            
            search_button = driver.find_element(By.XPATH, "//input[@type='button' and @value=' 查詢 ']")
            search_button.click()

            try:
                wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='table01']//tr[@class='even' or @class='odd']")))
                rows = driver.find_elements(By.XPATH, "//div[@id='table01']//tr[@class='even' or @class='odd']")
                
                print(f"民國 {year} 年資料載入完成，找到 {len(rows)} 列資料，開始動態解析...")

                # ======================================================
                # === 核心修正：改用 while 迴圈動態判斷單列/雙列格式 ===
                # ======================================================
                i = 0
                while i < len(rows):
                    main_row = rows[i]
                    cols = main_row.find_elements(By.TAG_NAME, 'td')
                    
                    if not cols: # 如果是空行，跳過
                        i += 1
                        continue
                    
                    first_cell = cols[0]
                    # 判斷第一個儲存格是否有 rowspan="2" 屬性
                    if first_cell.get_attribute('rowspan') == '2':
                        # --- 雙列格式 (新版) ---
                        date_row = rows[i+1]
                        date_cols = date_row.find_elements(By.TAG_NAME, 'td')
                        
                        company_data = {
                            '查詢年份': year,
                            '公司代號': cols[0].text if len(cols) > 0 else 'N/A',
                            '公司名稱': cols[1].text if len(cols) > 1 else 'N/A',
                            '股東會類型': cols[3].text if len(cols) > 3 else 'N/A',
                            '股東會日期': cols[4].text if len(cols) > 4 else 'N/A',
                            '停止過戶起日': date_cols[0].text if len(date_cols) > 0 else 'N/A',
                            '停止過戶迄日': date_cols[1].text if len(date_cols) > 1 else 'N/A',
                            '召開方式': cols[7].text if len(cols) > 7 else 'N/A',
                            '開會地點': cols[8].text if len(cols) > 8 else 'N/A',
                            '是否改選董監': cols[10].text if len(cols) > 10 else 'N/A',
                            '電子投票平台': cols[15].text if len(cols) > 15 else '資料不存在',
                            '投票網址': cols[16].text if len(cols) > 16 else '資料不存在'
                        }
                        all_data.append(company_data)
                        i += 2 # 處理完一組(兩列)，索引跳2
                    else:
                        # --- 單列格式 (舊版) ---
                        company_data = {
                            '查詢年份': year,
                            '公司代號': cols[0].text if len(cols) > 0 else 'N/A',
                            '公司名稱': cols[1].text if len(cols) > 1 else 'N/A',
                            '股東會類型': cols[3].text if len(cols) > 3 else 'N/A',
                            '股東會日期': cols[4].text if len(cols) > 4 else 'N/A',
                            '停止過戶起日': '資料不存在', # 舊格式無此欄位
                            '停止過戶迄日': '資料不存在', # 舊格式無此欄位
                            '召開方式': '資料不存在', # 舊格式無此欄位
                            '開會地點': cols[5].text if len(cols) > 5 else 'N/A',
                            '是否改選董監': cols[6].text if len(cols) > 6 else 'N/A',
                            '電子投票平台': cols[11].text if len(cols) > 11 else '資料不存在',
                            '投票網址': cols[12].text if len(cols) > 12 else '資料不存在'
                        }
                        all_data.append(company_data)
                        i += 1 # 處理完一組(單列)，索引跳1
                        
                    # 抓取公司名稱用於檔案命名
                    if company_code and not company_name and len(cols) > 1:
                        company_name = cols[1].text.strip()
                # ======================================================

            except TimeoutException:
                print(f"民國 {year} 年查無任何資料。")
                continue

        if not all_data:
            print("\n查詢結束，未抓取到任何有效資料。")
            return

        df = pd.DataFrame(all_data)
        print("\n\n--- 所有資料抓取完畢 ---")
        print(df)
        print(f"\n成功抓取總計 {len(df)} 筆資料。")

        date_str = datetime.now().strftime("%Y%m%d")
        year_str_for_file = year_input_str.replace(',', '_').replace('~', '至')
        
        if company_code:
            company_str = f"{company_code}{company_name}資料" if company_name else f"{company_code}資料"
        else:
            company_str = "全部公司資料"
            
        file_name = f"{date_str}_{year_str_for_file}_{company_str}.csv"
        
        df.to_csv(file_name, index=False, encoding='utf-8-sig')
        print(f"\n資料已成功儲存至檔案: {file_name}")

    finally:
        if driver:
            print("\n正在關閉瀏覽器...")
            driver.quit()

if __name__ == "__main__":
    main()
