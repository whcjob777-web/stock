#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
from PyPDF2 import PdfMerger
import glob
from datetime import datetime

def merge_pdfs():
    """
    合并两个PDF文件到一个文件中
    """
    # 获取最新的两个PDF文件
    current_dir = os.path.dirname(os.path.abspath(__file__))
    stock_report_pattern = f"{current_dir}/美股市场每日数据_*.pdf"
    option_report_pattern = f"{current_dir}/期权未平仓数据分析_*.pdf"
    
    # 查找最新的股票数据报告
    stock_reports = glob.glob(stock_report_pattern)
    if not stock_reports:
        print("未找到美股市场每日数据报告")
        return
    
    # 按修改时间排序，获取最新的文件
    stock_reports.sort(key=os.path.getmtime, reverse=True)
    latest_stock_report = stock_reports[0]
    print(f"找到最新的美股市场每日数据报告: {latest_stock_report}")
    
    # 查找最新的期权分析报告
    option_reports = glob.glob(option_report_pattern)
    if not option_reports:
        print("未找到期权未平仓数据分析报告")
        return
    
    # 按修改时间排序，获取最新的文件
    option_reports.sort(key=os.path.getmtime, reverse=True)
    latest_option_report = option_reports[0]
    print(f"找到最新的期权未平仓数据分析报告: {latest_option_report}")
    
    # 创建PDF合并对象
    merger = PdfMerger()
    
    # 添加文件到合并对象
    print("正在合并PDF文件...")
    merger.append(latest_stock_report)
    merger.append(latest_option_report)
    
    # 生成合并后的文件名
    today = datetime.now().strftime("%Y%m%d")
    output_filename = f"{current_dir}/output/综合市场分析报告_{today}.pdf"
    
    # 写入合并后的PDF文件
    with open(output_filename, "wb") as output_file:
        merger.write(output_file)
    
    # 关闭合并对象
    merger.close()
    os.unlink(latest_stock_report)
    os.unlink(latest_option_report)
    
    print(f"PDF文件已成功合并为: {output_filename}")

def merge_specific_pdfs(stock_pdf, option_pdf):
    """
    合并指定的两个PDF文件
    
    Args:
        stock_pdf (str): 股票数据报告PDF文件路径
        option_pdf (str): 期权分析报告PDF文件路径
    """
    # 检查文件是否存在
    if not os.path.exists(stock_pdf):
        print(f"文件不存在: {stock_pdf}")
        return
    
    if not os.path.exists(option_pdf):
        print(f"文件不存在: {option_pdf}")
        return
    
    # 创建PDF合并对象
    merger = PdfMerger()
    
    # 添加文件到合并对象
    print("正在合并PDF文件...")
    merger.append(stock_pdf)
    merger.append(option_pdf)
    
    # 生成合并后的文件名
    today = datetime.now().strftime("%Y%m%d")
    current_dir = os.path.dirname(os.path.abspath(__file__))
    output_filename = f"{current_dir}/综合市场分析报告_{today}.pdf"
    
    # 写入合并后的PDF文件
    with open(output_filename, "wb") as output_file:
        merger.write(output_file)
    
    # 关闭合并对象
    merger.close()
    os.unlink(stock_pdf)
    os.unlink(option_pdf)
    
    print(f"PDF文件已成功合并为: {output_filename}")

if __name__ == "__main__":
    import sys
    
    if len(sys.argv) == 3:
        # 如果提供了两个文件路径参数，则合并指定的文件
        stock_pdf_file = sys.argv[1]
        option_pdf_file = sys.argv[2]
        merge_specific_pdfs(stock_pdf_file, option_pdf_file)
    else:
        # 否则合并最新的两个PDF文件
        merge_pdfs()