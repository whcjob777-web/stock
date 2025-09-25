from tools import *

import os

proxy = 'http://127.0.0.1:7890' # 代理设置，此处修改
os.environ['HTTP_PROXY'] = proxy
os.environ['HTTPS_PROXY'] = proxy

# 美股主要指数
MAJOR_INDICES = {
    '^DJI': '道琼斯',
    '^GSPC': '标普500',
    '^NDX': '纳斯达克100',
    "^RUT": "罗素2000"
}

# 七大科技巨头（Magnificent Seven）
MAGNIFICENT_SEVEN = {
    'AAPL': '苹果',
    'MSFT': '微软',
    'GOOGL': '谷歌',
    'AMZN': '亚马逊',
    'META': 'Meta',
    'NVDA': '英伟达',
    'TSLA': '特斯拉'
}

# 主要板块ETF
SECTOR_ETF = {
    'XLK': '科技',
    'XLC': '通信',
    'XLF': '金融',
    'XLY': '可选消费',
    'XLP': '必选消费',
    'XLE': '能源',
    'XLV': '医疗',
    'XLI': '工业',
    'XLB': '材料',
    'XLRE': '房地产',
    'XLU': '公用事业'
}

# 加密股
CRYPTO_STOCKS = {
    "MSTR":"MSTR",
    "BMNR":"BMNR",
    "SBET":"SBET",
}


def main():
    print("正在获取美股市场数据...")
    date_str = get_previous_weekday()
    print(f"数据日期: {date_str}")
    
    # 获取主要指数数据
    print("\n正在获取主要指数数据...")
    indices_data = fetch_stock_data(MAJOR_INDICES, period="5d")
    
    # # 获取七大科技巨头数据
    print("\n正在获取七大科技巨头数据...")
    magnificent_seven_data = fetch_stock_data(MAGNIFICENT_SEVEN, period="5d")
    
    # # 获取板块数据
    print("\n正在获取主要板块数据...")
    sector_data = fetch_stock_data(SECTOR_ETF, period="5d")

    # # 获取加密股数据
    print("\n正在获取加密股数据...")
    crypto_data = fetch_stock_data(CRYPTO_STOCKS, period="5d")

    all_data = {
        "主要指数":indices_data,
        "七巨头":magnificent_seven_data,
        "板块":sector_data,
        "加密股":crypto_data
    }
    
    # 生成PDF报告
    print("\n正在生成PDF报告...")
    filename = generate_pdf_report(all_data)
    print(f"PDF报告已生成: {filename}")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"程序执行出错: {str(e)}")
        sys.exit(1)