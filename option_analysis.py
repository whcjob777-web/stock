import yfinance as yf
import matplotlib.pyplot as plt
import pandas as pd
import time
import os
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import sys
import tempfile

from tools import *

proxy = 'http://127.0.0.1:7890'  # 代理设置，此处修改
os.environ['HTTP_PROXY'] = proxy
os.environ['HTTPS_PROXY'] = proxy

def plot_open_interest_separate(calls, puts, stock_price, ticker_symbol):
    """
    分别绘制单个股票的期权未平仓合约数
    """
    # 设置中文字体
    plt.rcParams['font.sans-serif'] = ['Arial', 'SimHei']
    plt.rcParams['axes.unicode_minus'] = False

    # 创建图表
    fig, ax = plt.subplots(figsize=(12, 8))

    # 绘制看涨期权未平仓数
    if calls is not None and 'strike' in calls.columns and 'openInterest' in calls.columns:
        # 移除未平仓数为0的行
        calls_filtered = calls[calls['openInterest'] > 0]
        if not calls_filtered.empty:
            ax.plot(calls_filtered['strike'], calls_filtered['openInterest'],
                    label=f'{ticker_symbol} Calls', marker='o', linewidth=2, color='blue')
        else:
            print(f"{ticker_symbol}看涨期权没有有效的未平仓数据")
    else:
        print(f"{ticker_symbol}看涨期权数据不可用")

    # 绘制看跌期权未平仓数
    if puts is not None and 'strike' in puts.columns and 'openInterest' in puts.columns:
        # 移除未平仓数为0的行
        puts_filtered = puts[puts['openInterest'] > 0]
        if not puts_filtered.empty:
            ax.plot(puts_filtered['strike'], puts_filtered['openInterest'],
                    label=f'{ticker_symbol} Puts', marker='s', linewidth=2, color='red')
        else:
            print(f"{ticker_symbol}看跌期权没有有效的未平仓数据")
    else:
        print(f"{ticker_symbol}看跌期权数据不可用")

    # 添加股票当前价格的竖线
    if stock_price is not None:
        ax.axvline(x=stock_price, color='green', linestyle='--', alpha=0.7,
                   label=f'{ticker_symbol} Price: {stock_price:.2f}')

    # 设置图表标题和标签
    ax.set_xlabel('Strike Price')
    ax.set_ylabel('Open Interest')
    ax.set_title(f'{ticker_symbol} Options Open Interest')
    ax.legend()
    ax.grid(True, alpha=0.3)

    # 保存图表到临时文件
    temp_file = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
    plt.tight_layout()
    plt.savefig(temp_file.name, dpi=300, bbox_inches='tight')
    plt.close()
    
    return temp_file.name


def generate_options_pdf(all_data):
    """
    生成PDF报告
    """
    font_name = get_font()

    current_dir = os.path.dirname(os.path.abspath(__file__))
    # 创建PDF文档
    filename = f"{current_dir}/期权未平仓数据分析_{time.strftime('%Y%m%d')}.pdf"
    doc = SimpleDocTemplate(filename, pagesize=A4)
    styles = getSampleStyleSheet()

    # 更新样式以使用中文字体
    styles['Normal'].fontName = font_name
    styles['Heading1'].fontName = font_name
    styles['Heading2'].fontName = font_name

    story = []

    # 标题
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=18,
        spaceAfter=30,
        alignment=1,  # 居中
        fontName=font_name
    )

    title = Paragraph("期权未平仓数据分析报告", title_style)
    story.append(title)

    # 日期
    date_str = time.strftime('%Y-%m-%d')
    date_para = Paragraph(f"数据日期: {date_str}", styles['Normal'])
    story.append(date_para)
    story.append(Spacer(1, 0.2*inch))
    del_path = []
    for key,value in all_data.items():
        story.append(Paragraph(f"{key}期权未平仓数据", styles['Heading1']))
        calls = value["option_chain"].calls
        puts = value["option_chain"].puts
        price = value["stock_price"]
        chart_path = plot_open_interest_separate(calls, puts, price, str(key))
        chart = Image(chart_path, width=6*inch, height=4*inch)
        story.append(chart)
        story.append(Spacer(1, 0.3*inch))

        if price is not None:
            info = Paragraph(f"{key}当前价格: {price:.2f}", styles['Normal'])
            story.append(info)
            story.append(Spacer(1, 0.1*inch))
        
        del_path.append(chart_path)
    
    doc.build(story)

    for path in del_path:
        os.remove(path)

    return filename



OPTINONS = ["QQQ","SPY"]

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