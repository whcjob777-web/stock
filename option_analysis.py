from tools import *

proxy = 'http://127.0.0.1:7890'  # 代理设置，此处修改
os.environ['HTTP_PROXY'] = proxy
os.environ['HTTPS_PROXY'] = proxy

OPTINONS = ["QQQ", "SPY", "TSLA"]

def main():
    """
    主函数
    """
    all_data = {}
    for option in OPTINONS:
        print(f"正在获取{option}期权数据...")
        opt_chain, stock_price = get_option_data(option)
        all_data[option] = {
            "option_chain": opt_chain,
            "stock_price": stock_price
        }

    print("正在生成PDF报告...")
    filename = generate_options_pdf(all_data)
    print(f"PDF报告已生成: {filename}")


if __name__ == "__main__":
    main()