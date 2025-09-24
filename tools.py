import yfinance as yf
import pandas as pd
from datetime import datetime, timedelta
import sys
import time
from typing import Dict, Optional
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.graphics.shapes import Drawing
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.textlabels import Label



def get_previous_weekday():
    """
    获取最近的交易日（周一到周五）
    如果今天是周末，则返回最近的周五
    """
    today = datetime.now()
    # 如果是周日，减去2天到周五
    if today.weekday() == 6:
        return (today - timedelta(days=2)).strftime('%Y-%m-%d')
    # 如果是周六，减去1天到周五
    elif today.weekday() == 5:
        return (today - timedelta(days=1)).strftime('%Y-%m-%d')
    else:
        return today.strftime('%Y-%m-%d')

def fetch_single_ticker(symbol: str, period: str = "5d") -> Optional[pd.DataFrame]:
    """
    获取单个股票或指数数据，包含重试机制
    """
    max_retries = 3
    for attempt in range(max_retries):
        try:
            ticker = yf.Ticker(symbol)
            hist = ticker.history(period=period)
            return hist
        except Exception as e:
            if "Too Many Requests" in str(e) or "Rate limited" in str(e):
                wait_time = (attempt + 1) * 5  # 逐步增加等待时间
                print(f"请求频率限制，等待 {wait_time} 秒后重试 ({attempt + 1}/{max_retries})...")
                time.sleep(wait_time)
                continue
            else:
                print(f"获取 {symbol} 数据时出错: {str(e)}")
                return None
    return None

def fetch_stock_data(symbols_dict: Dict[str, str], period: str = "5d"):
    """
    获取股票或指数数据，添加延迟以避免API限制
    """
    data = {}
    symbol_count = len(symbols_dict)
    
    for i, (symbol, name) in enumerate(symbols_dict.items()):
        try:
            print(f"正在获取 {name} ({symbol}) 数据 ({i+1}/{symbol_count})...")
            hist = fetch_single_ticker(symbol, period)
            
            if hist is not None and not hist.empty:
                current_price = hist['Close'].iloc[-1]
                if len(hist) > 1:
                    previous_price = hist['Close'].iloc[-2]
                    change = current_price - previous_price
                    change_percent = (change / previous_price) * 100
                else:
                    change = 0
                    change_percent = 0
                
                data[name] = {
                    'symbol': symbol,
                    'price': round(current_price, 2),
                    'change': round(change, 2),
                    'change_percent': round(change_percent, 2)
                }
            else:
                data[name] = None
                
            # 在请求之间添加延迟以避免频率限制
            if i < symbol_count - 1:  # 不需要在最后一个请求后等待
                time.sleep(1)
                
        except Exception as e:
            print(f"获取 {name} ({symbol}) 数据时出错: {str(e)}")
            data[name] = None
            
    return data

def create_table_data(title, data):
    """
    创建用于PDF表格的数据
    """
    table_data = [[title, '', '', '', '']]
    table_data.append(['名称', '代码', '价格', '涨跌额', '涨跌幅(%)'])
    
    for name, info in data.items():
        if info:
            change_str = f"{info['change']:+.2f}"
            change_percent_str = f"{info['change_percent']:+.2f}"
            table_data.append([name, info['symbol'], str(info['price']), change_str, change_percent_str])
        else:
            table_data.append([name, 'N/A', 'N/A', 'N/A', 'N/A'])
            
    return table_data

def generate_pdf_report(indices_data, magnificent_seven_data=None, sector_data=None):
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
    
    # 创建PDF文档
    import os
    current_dir = os.path.dirname(os.path.abspath(__file__))
    filename = f"{current_dir}/美股市场每日数据_{get_previous_weekday().replace('-', '')}.pdf"
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
    
    title = Paragraph("美股市场每日数据报告", title_style)
    story.append(title)
    
    # 日期
    date_str = get_previous_weekday()
    date_para = Paragraph(f"数据日期: {date_str}", styles['Normal'])
    story.append(date_para)
    story.append(Spacer(1, 0.2*inch))
    
    # 创建柱状图
    
    # 主要指数图表
    story.append(Paragraph("主要指数涨跌幅", styles['Heading1']))
    indices_chart = create_bar_chart(indices_data, "主要指数", font_name)
    story.append(indices_chart)
    story.append(Spacer(1, 0.3*inch))
    
    # 七大科技巨头图表
    story.append(PageBreak())
    story.append(Paragraph("七大科技巨头涨跌幅", styles['Heading1']))
    magnificent_chart = create_bar_chart(magnificent_seven_data, "七大科技巨头", font_name)
    story.append(magnificent_chart)
    story.append(Spacer(1, 0.3*inch))
    
    # 主要板块图表
    story.append(PageBreak())
    story.append(Paragraph("主要板块涨跌幅", styles['Heading1']))
    sector_chart = create_bar_chart(sector_data, "主要板块", font_name)
    story.append(sector_chart)
    story.append(Spacer(1, 0.3*inch))
    
    # 市场分析
    story.append(PageBreak())
    analysis_title = Paragraph("市场分析", styles['Heading1'])
    story.append(analysis_title)
    
    # 计算上涨和下跌的指数数量
    indices_up = sum(1 for info in indices_data.values() if info and info['change'] > 0)
    indices_down = sum(1 for info in indices_data.values() if info and info['change'] < 0)
    
    # 计算上涨和下跌的科技巨头数量
    magnificant_up = sum(1 for info in magnificent_seven_data.values() if info and info['change'] > 0)
    magnificant_down = sum(1 for info in magnificent_seven_data.values() if info and info['change'] < 0)
    
    # 计算上涨和下跌的板块数量
    sector_up = sum(1 for info in sector_data.values() if info and info['change'] > 0)
    sector_down = sum(1 for info in sector_data.values() if info and info['change'] < 0)
    
    analysis_data = [
        ['市场统计', '上涨数量', '下跌数量'],
        ['主要指数', str(indices_up), str(indices_down)],
        ['科技七巨头', str(magnificant_up), str(magnificant_down)],
        ['主要板块', str(sector_up), str(sector_down)]
    ]
    
    analysis_table = Table(analysis_data, colWidths=[1.5*inch, 1*inch, 1*inch])
    analysis_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, -1), font_name),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    story.append(analysis_table)
    story.append(Spacer(1, 0.3*inch))
    
    # 找出表现最好的和最差的
    all_data = {
        **{f"指数-{name}": info for name, info in indices_data.items()},
        **{f"科技-{name}": info for name, info in magnificent_seven_data.items()},
        **{f"板块-{name}": info for name, info in sector_data.items()}
    }
    
    valid_data = {name: info for name, info in all_data.items() if info}
    
    if valid_data:
        best_performer = max(valid_data.items(), key=lambda x: x[1]['change_percent'] if x[1] else -float('inf'))
        worst_performer = min(valid_data.items(), key=lambda x: x[1]['change_percent'] if x[1] else float('inf'))
        
        performance_data = []
        performance_data.append(['表现类型', '名称', '涨跌幅(%)'])
        
        if best_performer[1]:
            performance_data.append(['最佳表现', best_performer[0], f"+{best_performer[1]['change_percent']:.2f}%"])
            
        if worst_performer[1]:
            performance_data.append(['最差表现', worst_performer[0], f"{worst_performer[1]['change_percent']:.2f}%"])
            
        performance_table = Table(performance_data, colWidths=[1.5*inch, 2*inch, 1*inch])
        performance_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, -1), font_name),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        story.append(performance_table)
    
    # 构建PDF
    doc.build(story)
    return filename

def create_bar_chart(data, title, font_name):
    """
    创建柱状图
    """
    # 过滤掉无效数据
    valid_data = {name: info for name, info in data.items() if info}
    
    if not valid_data:
        return Paragraph("无有效数据", getSampleStyleSheet()['Normal'])
    
    # 准备数据
    names = list(valid_data.keys())
    changes = [info['change_percent'] for info in valid_data.values()]
    
    # 创建绘图对象
    drawing = Drawing(400, 200)
    
    # 创建柱状图
    bc = VerticalBarChart()
    bc.x = 50
    bc.y = 50
    bc.height = 120
    bc.width = 300
    bc.data = [changes]
    
    # 设置柱状图属性
    bc.barWidth = 8
    bc.groupSpacing = 10
    bc.valueAxis.valueMin = min(changes) - 1 if min(changes) < 0 else -1
    bc.valueAxis.valueMax = max(changes) + 1 if max(changes) > 0 else 1
    bc.valueAxis.valueStep = (bc.valueAxis.valueMax - bc.valueAxis.valueMin) / 10
    
    # 设置柱状图颜色 - 绿涨红跌
    for i, change in enumerate(changes):
        if change >= 0:
            # 上涨用绿色
            bc.bars[(0, i)].fillColor = colors.green
        else:
            # 下跌用红色
            bc.bars[(0, i)].fillColor = colors.red
    
    # 设置X轴标签
    bc.categoryAxis.labels.boxAnchor = 'ne'
    bc.categoryAxis.labels.angle = 30
    bc.categoryAxis.labels.dx = 5
    bc.categoryAxis.labels.dy = -5
    bc.categoryAxis.categoryNames = names
    bc.categoryAxis.labels.fontName = font_name
    bc.categoryAxis.labels.fontSize = 8
    
    # 设置Y轴标签
    bc.valueAxis.labels.fontName = font_name
    bc.valueAxis.labels.fontSize = 8
    
    # 计算每个柱子的高度比例
    value_range = bc.valueAxis.valueMax - bc.valueAxis.valueMin
    bar_heights = [(change - bc.valueAxis.valueMin) / value_range * bc.height for change in changes]
    max_height = max(bar_heights)
    # 添加数据标签，并确保与柱体顶部对齐
    for i, (change, bar_height) in enumerate(zip(changes, bar_heights)):
        label = Label()
        # 计算标签的y坐标，使其位于柱体顶部上方5个像素处
        x_position = bc.x + (bc.barWidth) + ((i) * (bc.barWidth + bc.groupSpacing))
        # x_position = bc.x + #+  (i * (bc.barWidth + bc.groupSpacing))
        # 计算标签的y坐标，使其位于柱体正上方
        y_position = bc.y + max_height + 5
        label.setOrigin(x_position , y_position)
        label.setText(f"{change:.1f}%")
        label.fontSize = 8
        label.fontName = font_name
        label.boxAnchor = 's'
        drawing.add(label)

    drawing.add(bc)

    return drawing