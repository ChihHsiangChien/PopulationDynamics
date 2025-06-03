# -*- coding: utf-8 -*-
"""新竹市人口圖 - 中文版"""

import pandas as pd
import matplotlib.pyplot as plt
import requests
import os

def setup_chinese_font():
    """設定中文字型"""
    import matplotlib.font_manager as fm
    font_path = "C:/Windows/Fonts/msjh.ttc" if os.name == 'nt' else "/System/Library/Fonts/STHeiti Medium.ttc"
    my_font = fm.FontProperties(fname=font_path)
    plt.rcParams["font.family"] = my_font.get_name()
    return my_font

def download_and_load_data(url):
    """下載並載入數據"""
    response = requests.get(url)
    response.raise_for_status()
    
    with open("hsinchuCityPopulation.xlsx", 'wb') as f:
        f.write(response.content)
    
    return pd.read_excel("hsinchuCityPopulation.xlsx")

def transform_data(df):
    """數據轉換"""
    # 轉換年月格式
    df["西元年月"] = df["年月"].apply(lambda x: pd.to_datetime(f"{int(str(x)[:3]) + 1911}-{int(str(x)[3:]):02d}"))
    
    # 篩選男女合計數據
    df_filtered = df[df["性別"] == "男女合計"]
    
    # 按區域和年月分組
    return df_filtered.groupby(["區域別", "西元年月"]).agg({
        "人口數": "sum", "遷入人數": "sum", "遷出人數": "sum", 
        "出生人數": "sum", "死亡人數": "sum"
    })

def create_charts(grouped_df):
    """生成圖表"""
    os.makedirs("output_charts", exist_ok=True)
    
    colors = {'population': 'tab:blue', 'in': 'tab:red', 'out': 'tab:green', 
              'birth': 'tab:orange', 'death': 'tab:purple'}
    
    for region in grouped_df.index.get_level_values(0).unique():
        data = grouped_df.loc[region]
        
        fig, ax1 = plt.subplots(figsize=(12, 6))
        
        # 主軸：人口數
        ax1.plot(data.index, data['人口數'], color=colors['population'], label='總人口')
        ax1.set_xlabel('年月', fontsize=12)
        ax1.set_ylabel('人口數', color=colors['population'], fontsize=12)
        ax1.tick_params(axis='y', labelcolor=colors['population'])
        
        # 次軸：其他指標
        ax2 = ax1.twinx()
        ax2.plot(data.index, data['遷入人數'], color=colors['in'], label='遷入人數')
        ax2.plot(data.index, data['遷出人數'], color=colors['out'], label='遷出人數')
        ax2.plot(data.index, data['出生人數'], color=colors['birth'], label='出生人數')
        ax2.plot(data.index, data['死亡人數'], color=colors['death'], label='死亡人數')
        ax2.set_ylabel('遷徙與出生死亡人數', color='tab:red', fontsize=12)
        ax2.tick_params(axis='y', labelcolor='tab:red')
        
        plt.title(f"{region}人口與相關變化趨勢圖", fontsize=14)
        fig.legend(loc="upper left", bbox_to_anchor=(0.13, 0.88))
        plt.xticks(rotation=45)
        plt.tight_layout()
        
        plt.savefig(f"output_charts/{region}_趨勢圖.png", dpi=300)
        plt.close()
        print(f"已生成 {region} 趨勢圖")

def main():
    """主函數"""
    url = "https://odws.hccg.gov.tw/001/Upload/25/opendataback/9059/57/debabe0e-c828-4a63-8dde-30ce3a5a7f87.xlsx"
    
    try:
        # 設定字型
        setup_chinese_font()
        
        # 載入並處理數據
        df = download_and_load_data(url)
        print(f"數據載入完成: {df.shape}")
        
        grouped_df = transform_data(df)
        print(f"數據處理完成: {grouped_df.shape}")
        
        # 生成圖表
        create_charts(grouped_df)
        print("所有圖表生成完成！")
        
        return grouped_df
        
    except Exception as e:
        print(f"錯誤: {e}")
        return None

# 執行
if __name__ == "__main__":
    result = main()
    if result is not None:
        print("\n數據預覽:")
        print(result.head())