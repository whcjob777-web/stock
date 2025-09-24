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

proxy = 'http://127.0.0.1:7890'  # 代理设置，此处修改
os.environ['HTTP_PROXY'] = proxy
os.environ['HTTPS_PROXY'] = proxy


def get_option_data(ticker_symbol, max_retries=3):
    """
    获取指定股票代码的期权数据，包含重试机制
    """
    for attempt in range(max_retries):
        try:
            print(f"正在获取 {ticker_symbol} 的期权数据... (尝试 {attempt + 1}/{max_retries})")
            ticker = yf.Ticker(ticker_symbol)
            expiration_dates = ticker.options

            # 获取股票当前价格
            stock_price = None
            try:
                stock_info = ticker.info
                stock_price = stock_info.get('regularMarketPrice', None)
                if stock_price is None:
                    # 尝试其他可能的价格字段
                    stock_price = stock_info.get('currentPrice', None)
                    if stock_price is None:
                        stock_price = stock_info.get('previousClose', None)
            except Exception as e:
                print(f"获取 {ticker_symbol} 当前价格时出错: {e}")

            # 获取最近的到期日数据
            if expiration_dates:
                # 使用最近的到期日
                nearest_expiration = expiration_dates[0]
                opt_chain = ticker.option_chain(nearest_expiration)
                return opt_chain.calls, opt_chain.puts, stock_price
            else:
                print(f"无法获取 {ticker_symbol} 的期权数据")
                raise

        except yf.exceptions.YFRateLimitError:
            if attempt < max_retries - 1:
                print(f"遇到请求限制，等待10秒后重试...")
                time.sleep(10)
            else:
                print(f"达到最大重试次数，无法获取 {ticker_symbol} 的数据")
                return None, None, None
        except Exception as e:
            print(f"获取 {ticker_symbol} 数据时出错: {e}")
            if attempt < max_retries - 1:
                print(f"等待5秒后重试...")
                time.sleep(5)
            else:
                print(f"达到最大重试次数，无法获取 {ticker_symbol} 的数据")
                return None, None, None

    return None, None, None


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


def generate_pdf_report(qqq_calls, qqq_puts, qqq_price, spy_calls, spy_puts, spy_price):
    """
    生成PDF报告
    """
    # 注册中文字体
    font_name = 'Helvetica'  # 默认字体

    # 尝试多种macOS系统自带的中文字体
    mac_chinese_fonts = [
        ('PingFang SC', '/System/Library/Fonts/PingFang.ttc'),
        ('Heiti SC', '/System/Library/Fonts/STHeiti Light.ttc'),
        ('Songti SC', '/System/Library/Fonts/Songti.ttc'),
        ('Hiragino Sans GB', '/System/Library/Fonts/Hiragino Sans GB.ttc'),
        ('Arial Unicode MS', '/System/Library/Fonts/Arial Unicode.ttf')
    ]

    for font_name_try, font_path in mac_chinese_fonts:
        try:
            pdfmetrics.registerFont(TTFont(font_name_try, font_path))
            font_name = font_name_try
            print(f"成功注册中文字体: {font_name}")
            break
        except Exception as e:
            print(f"注册字体 {font_name_try} 失败: {str(e)}")
            continue

    # 如果上述字体都不可用，尝试系统中可能存在的其他中文字体
    if font_name == 'Helvetica':
        chinese_fonts_fallback = [
            ('SimSun', 'SimSun.ttf'),
            ('STSong', 'STSong.ttf')
        ]

        for font_name_try, font_path in chinese_fonts_fallback:
            try:
                pdfmetrics.registerFont(TTFont(font_name_try, font_path))
                font_name = font_name_try
                print(f"成功注册备选中文字体: {font_name}")
                break
            except:
                continue

    if font_name == 'Helvetica':
        print("警告: 未找到可用的中文字体，将使用默认字体Helvetica")
    else:
        print(f"使用字体: {font_name}")

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
    
    # 绘制QQQ图表
    story.append(Paragraph("QQQ期权未平仓数据", styles['Heading1']))
    qqq_chart_path = plot_open_interest_separate(qqq_calls, qqq_puts, qqq_price, "QQQ")
    qqq_chart = Image(qqq_chart_path, width=6*inch, height=4*inch)
    story.append(qqq_chart)
    story.append(Spacer(1, 0.3*inch))
    
    # 添加QQQ相关信息
    if qqq_price is not None:
        qqq_info = Paragraph(f"QQQ当前价格: {qqq_price:.2f}", styles['Normal'])
        story.append(qqq_info)
        story.append(Spacer(1, 0.1*inch))
    
    # 绘制SPY图表
    story.append(PageBreak())
    story.append(Paragraph("SPY期权未平仓数据", styles['Heading1']))
    spy_chart_path = plot_open_interest_separate(spy_calls, spy_puts, spy_price, "SPY")
    spy_chart = Image(spy_chart_path, width=6*inch, height=4*inch)
    story.append(spy_chart)
    story.append(Spacer(1, 0.3*inch))
    
    # 添加SPY相关信息
    if spy_price is not None:
        spy_info = Paragraph(f"SPY当前价格: {spy_price:.2f}", styles['Normal'])
        story.append(spy_info)
        story.append(Spacer(1, 0.1*inch))

    # 构建PDF
    doc.build(story)
    
    # 清理临时文件
    os.unlink(qqq_chart_path)
    os.unlink(spy_chart_path)
    
    return filename


def main():
    """
    主函数
    """
    print("正在获取QQQ期权数据...")
    qqq_calls, qqq_puts, qqq_price = get_option_data("QQQ")
    
    # 等待一段时间避免触发速率限制
    time.sleep(1)
    
    print("正在获取SPY期权数据...")
    spy_calls, spy_puts, spy_price = get_option_data("SPY")
    
    print("正在生成PDF报告...")
    filename = generate_pdf_report(qqq_calls, qqq_puts, qqq_price, spy_calls, spy_puts, spy_price)
    print(f"PDF报告已生成: {filename}")


if __name__ == "__main__":
    main()