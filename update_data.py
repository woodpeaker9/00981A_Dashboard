import requests
import pandas as pd
import os
from datetime import datetime
import urllib3

# 禁用 SSL 警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

FUND_CODE = "49YTW"
DOWNLOAD_URL = f"https://www.ezmoney.com.tw/ETF/Fund/AssetExcelNPOI?fundCode={FUND_CODE}"
DATA_DIR = "etf_data"
os.makedirs(DATA_DIR, exist_ok=True)

def run_update():
    today_str = datetime.now().strftime("%Y%m%d")
    file_path = os.path.join(DATA_DIR, f"{FUND_CODE}_{today_str}.xlsx")
    
    # 1. 下載當日 Excel
    print(f"正在下載今日資料: {today_str}...")
    try:
        response = requests.get(DOWNLOAD_URL, verify=False, timeout=30)
        if response.status_code == 200:
            with open(file_path, 'wb') as f:
                f.write(response.content)
        else:
            print(f"下載失敗: {response.status_code}")
            return
    except Exception as e:
        print(f"連線錯誤: {e}")
        return

    # 2. 檢查是否有至少兩天的資料
    all_files = sorted([f for f in os.listdir(DATA_DIR) if f.endswith('.xlsx')], reverse=True)
    if len(all_files) < 2:
        print("目前資料不足兩天，無法產生對比報告。")
        return

    # 3. 解析函數 (張數換算)
    def parse_etf(path):
        raw = pd.read_excel(path)
        # 尋找表頭「股票代號」
        header_row = raw[raw.iloc[:, 0] == "股票代號"].index[0]
        df = pd.read_excel(path, skiprows=header_row + 1).dropna(subset=['股票代號', '股票名稱'])
        
        # 股數轉張數 (1000股 = 1張)
        if df['股數'].dtype == object:
            df['股數'] = df['股數'].str.replace(',', '', regex=True).astype(float)
        else:
            df['股數'] = df['股數'].astype(float)
        
        df['張數'] = df['股數'] / 1000
        return df[['股票代號', '股票名稱', '張數', '持股權重']]

    df_now = parse_etf(os.path.join(DATA_DIR, all_files[0]))
    df_prev = parse_etf(os.path.join(DATA_DIR, all_files[1]))

    # 4. 比對邏輯
    merged = pd.merge(df_now, df_prev, on='股票代號', how='outer', suffixes=('_今', '_昨'))
    merged['張數_今'] = merged['張數_今'].fillna(0)
    merged['張數_昨'] = merged['張數_昨'].fillna(0)
    merged['股票名稱'] = merged['股票名稱_今'].fillna(merged['股票名稱_昨'])
    merged['張數變動'] = (merged['張數_今'] - merged['張數_昨']).round(3)
    
    # 增減幅 %
    def calc_pct(row):
        if row['張數_昨'] == 0: return 100.0 if row['張數_今'] > 0 else 0.0
        return round((row['張數變動'] / row['張數_昨']) * 100, 2)
    merged['增減幅(%)'] = merged.apply(calc_pct, axis=1)

    # 狀態與排序
    def get_status(row):
        if row['張數_昨'] == 0: return "🆕新進", 0
        if row['張數_今'] == 0: return "❌剔除", 1
        if row['張數變動'] > 0: return "➕加碼", 2
        if row['張數變動'] < 0: return "➖減碼", 3
        return "＝持平", 4

    status_data = merged.apply(get_status, axis=1)
    merged['狀態'] = [x[0] for x in status_data]
    merged['sort'] = [x[1] for x in status_data]

    # 排序：有變動的在前
    report = merged.sort_values(['sort', '張數變動'], ascending=[True, False]).drop(columns=['sort'])

    # 5. 產出 README 內容
    summary = f"# 00981A 每日持股變動監測\n\n"
    summary += f"> **更新時間**：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} (台北時間)\n"
    summary += f"> **比對區間**：{all_files[1][6:14]} ➔ {all_files[0][6:14]}\n\n"
    summary += "### 📊 持股變動明細 (單位：張)\n\n"
    summary += report[['股票代號', '股票名稱', '張數_昨', '張數_今', '張數變動', '增減幅(%)', '狀態']].to_markdown(index=False)
    summary += "\n\n---\n*自動更新由 GitHub Actions 提供*"

    with open("README.md", "w", encoding="utf-8") as f:
        f.write(summary)

if __name__ == "__main__":
    run_update()
