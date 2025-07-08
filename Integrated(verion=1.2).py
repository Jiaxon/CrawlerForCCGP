# -*- coding=utf-8 -*-
import math
import time
import requests
from lxml import etree
from chardet import detect
from datetime import datetime, timedelta
import random
import xlsxwriter
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from selenium import webdriver
from selenium.webdriver.common.by import By
import csv
import json
import openpyxl
import signal
import sys

# ------------------------- 配置文件 -------------------------
# 邮件服务器配置
SMTP_SERVER = "smtp.example.com"  # SMTP服务器地址
SMTP_PORT = 465  # SMTP端口
SENDER_EMAIL = "sender@example.com"  # 发件邮箱
SENDER_PASSWORD = "your_password"  # 发件邮箱密码
RECEIVER_EMAIL = "receiver@example.com"  # 收件邮箱

# 全局变量用于保存当前抓取的数据
current_data = []

def signal_handler(signum, frame):
    """处理用户中断信号"""
    print("\n检测到用户中断，正在保存已抓取的数据...")
    if current_data:
        head = ['序号', '类型', '名称', '日期', '招标人', '代理机构', '区域', '详情', '项目概况']
        output_filename = "interrupted_data_" + datetime.now().strftime("%Y%m%d_%H%M%S")
        writer_excel(current_data, head, '中标公告', output_filename)
        print(f"已保存 {len(current_data)} 条数据到 {output_filename}.xlsx")
    else:
        print("没有数据需要保存。")
    sys.exit(0)

# 注册信号处理器
signal.signal(signal.SIGINT, signal_handler)

# ------------------------- 数据爬取模块 -------------------------
def get_request_headers(referer=None):
    """生成随机的HTTP请求头，用于绕过反爬机制"""
    user_agents = [
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.64 Safari/537.36',
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:100.0) Gecko/20100101 Firefox/100.0',
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.64 Safari/537.36'
    ]
    ua = random.choice(user_agents)  # 随机选择一个User-Agent

    headers = {
        "User-Agent": ua,
        "Host": "search.ccgp.gov.cn",
        "Referer": referer if referer else "http://search.ccgp.gov.cn/",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
        "Accept-Encoding": "gzip, deflate",
        "Accept-Language": "zh-CN,zh;q=0.9",
        "Connection": "keep-alive",
        "Upgrade-Insecure-Requests": "1"
    }
    return headers


def open_url(url, params, refer=None):
    """发送HTTP请求并返回响应"""
    headers = get_request_headers(refer)
    
    # 显示延迟提示
    delay_seconds = random.randint(2, 6)
    print(f"等待 {delay_seconds} 秒以避免频繁请求...")
    
    try:
        time.sleep(delay_seconds)  # 随机延迟，避免频繁请求
    except KeyboardInterrupt:
        print("用户中断了延迟等待")
        raise
    
    response = requests.get(url, headers=headers, params=params, allow_redirects=True)
    if response.status_code != 200:
        print(f"请求失败: {response.status_code}")
    return response


def crawler_ccgp(sheetdata=[], year='', buyerName=''):
    """爬取中国政府采购网的招标公告数据"""
    global current_data
    current_data = sheetdata
    
    url = 'http://search.ccgp.gov.cn/bxsearch?'
    curr_date = datetime.now()
    start_date = curr_date - timedelta(days=3)  # 默认抓取最近3天的数据
    start_time = start_date.strftime("%Y:%m:%d")
    end_time = curr_date.strftime("%Y:%m:%d")

    params = {
        'searchtype': 1,
        'page_index': 1, # 页码
        'bidSort': 0, # 公告类型
        'buyerName': buyerName,  # 采购人
        'projectId': '', # 项目编号
        'pinMu': 0, # 品目：0表示所有 1表示货物类 2表示工程类 3表示服务类
        'bidType': 0, # 公告类型 0表示所有类别 1表示中央公告 2表示地方公告
        'dbselect': 'bidx',
        'kw': '公告',  # 搜索关键词
        'start_time': start_time, #筛选开始时间
        'end_time': end_time, # 筛选结束时间
        'timeType': 1, # 时间类型 0表示今日 1表示近三天 2表示近一周 3表示近一个月 4表示近三个月 5表示近半年 6表示指定时间，可通过开始结束时间设置
        'displayZone': '',  # 区域筛选
        'zoneId': '45',  # 区域Id 区域筛选和区域Id必须一一对应，例如区域筛选是广西 那区域Id必须是45 这样才能实现筛选广西地区
        'pppStatus': 0, #ppp项目状态
        'agentName': '' # 代理机构名称
    }

    try:
        print("开始获取数据...")
        resp = open_url(url, params)
        resp.raise_for_status()  # 检查响应状态
        html = resp.content.decode('utf-8')
        tree = etree.HTML(html)

        try:
            total = int(tree.xpath('/html/body/div[5]/div[1]/div/p[1]/span[2]')[0].text.strip())  # 获取总数据量
        except IndexError:
            print("未找到数据总数")
            return sheetdata

        print(f"找到 {total} 条数据")
        
        if total > 0:
            pagesize = math.ceil(total / 20)  # 计算总页数
            print(f"总共 {pagesize} 页数据需要抓取")
            
            for curr_page in range(1, pagesize + 1):
                print(f"正在抓取第 {curr_page}/{pagesize} 页数据...")
                
                list = tree.xpath('/html/body/div[5]/div[2]/div/div/div[1]/ul/li')
                for i, li in enumerate(list):
                    try:
                        title = li[0]
                        summary = li[1]
                        span = li[2]
                        info = span.xpath('string()').replace(' ', '').replace('\r', '').replace('\n', '').replace('\t', '')

                        str1 = info[:info.index('公告')]
                        str2 = info[info.index('公告'):].replace('公告', '')
                        strs = str2.split('|')

                        if len(strs) > 1:
                            row = [len(sheetdata) + 1, '公告', title.text.strip()]
                            link_href = title.get('href')

                            # 安全地处理 str1 的分割
                            str1_parts = str1.split('|')
                            
                            # 使用安全的索引访问，提供默认值
                            date_part = str1_parts[0][:10] if len(str1_parts) > 0 else ''
                            buyer_part = str1_parts[1].replace('采购人：', '') if len(str1_parts) > 1 else ''
                            agent_part = str1_parts[2].replace('代理机构：', '') if len(str1_parts) > 2 else ''
                            
                            row.extend([
                                date_part,  # 日期
                                buyer_part,  # 招标人
                                agent_part,  # 代理机构
                                strs[0],  # 区域
                                link_href,  # 详情链接
                                summary.text.strip() if summary.text else ''  # 项目概况
                            ])
                            sheetdata.append(row)
                            current_data = sheetdata
                            print(f"  已获取第 {i+1} 条数据: {title.text.strip()[:30]}...")
                            
                    except (ValueError, IndexError) as e:
                        # 如果解析失败，记录错误并跳过此条记录
                        print(f"解析数据时出错，跳过此条记录: {e}")
                        continue
                    except KeyboardInterrupt:
                        print("用户中断了数据抓取")
                        raise

                # 获取下一页
                if curr_page < pagesize:
                    params['page_index'] = curr_page + 1
                    print(f"准备抓取下一页数据...")
                    resp = open_url(url, params, resp.url)
                    html = resp.content.decode('utf-8')
                    tree = etree.HTML(html)

    except KeyboardInterrupt:
        print("数据抓取被用户中断")
        raise
    except Exception as e:
        print(f"抓取数据时发生错误: {e}")
        return sheetdata

    return sheetdata


# ------------------------- 数据处理模块 -------------------------
def load_existing_data(file_path):
    """使用openpyxl加载已有的Excel数据"""
    try:
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
        data = []
        headers = []
        
        # 读取表头
        for cell in ws[1]:
            headers.append(cell.value)
        
        # 读取数据
        for row in ws.iter_rows(min_row=2, values_only=True):
            if any(cell is not None for cell in row):  # 跳过空行
                data.append(dict(zip(headers, row)))
        
        return data, headers
    except Exception as e:
        print(f"读取文件失败: {e}")
        return None, None


def get_existing_titles(data):
    """提取已有数据的标题"""
    if data is not None:
        return {item.get('名称', '') for item in data if item.get('名称')}
    return set()


def filter_duplicates(new_data, existing_titles):
    """过滤掉重复的数据"""
    filtered_data = [row for row in new_data if row[2] not in existing_titles]
    return filtered_data


# ------------------------- 邮件通知模块 -------------------------
def send_email(subject, body):
    """发送邮件通知"""
    msg = MIMEMultipart()
    msg["From"] = SENDER_EMAIL
    msg["To"] = RECEIVER_EMAIL
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "html"))

    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.sendmail(SENDER_EMAIL, RECEIVER_EMAIL, msg.as_string())
            print("邮件发送成功!")
    except Exception as e:
        print(f"邮件发送失败: {e}")


def generate_email_body(new_data):
    """生成HTML格式的邮件内容，不使用pandas"""
    if not new_data:
        return None

    columns = ['序号', '类型', '名称', '日期', '招标人', '代理机构', '区域', '详情', '项目概况']
    
    # 手动构建HTML表格
    html_rows = []
    
    # 表头
    header_row = '<tr>' + ''.join([f'<th>{col}</th>' for col in columns]) + '</tr>'
    html_rows.append(header_row)
    
    # 数据行
    for row in new_data:
        data_row = '<tr>' + ''.join([f'<td>{cell}</td>' for cell in row]) + '</tr>'
        html_rows.append(data_row)
    
    html_table = '<table class="data-table">' + ''.join(html_rows) + '</table>'

    style = """
    <style>
        .data-table { width: 100%; border-collapse: collapse; font-family: Arial, sans-serif; font-size: 14px; color: #333; }
        .data-table th, .data-table td { padding: 12px; text-align: left; border-bottom: 1px solid #ddd; }
        .data-table th { background-color: #f8f9fa; font-weight: bold; color: #333; }
        .data-table tr:hover { background-color: #f1f1f1; }
        .data-table a { color: #007bff; text-decoration: none; }
        .data-table a:hover { text-decoration: underline; }
    </style>
    """

    body = f"""
    <html>
        <head>{style}</head>
        <body>
            <h3>发现 {len(new_data)} 条新招标公告：</h3>
            {html_table}
            <p>请及时查看附件或访问网站获取详细信息。</p>
        </body>
    </html>
    """
    return body


def writer_excel(data, head=['A1','A2','A3','A4','A5','A6','A7','A8'] ,  sheetname='sheet1',filename='DataFile'):
    "用XlsxWriter库把数据写入Excel文件"
    workbook = xlsxwriter.Workbook(filename+'.xlsx')
    worksheet = workbook.add_worksheet(sheetname)

    row = 0
    col = 0

    # 插入表头
    cvi = 0
    for cv in head:
        worksheet.write(row, col + cvi, cv)
        cvi += 1
    row += 1
    # 插入表数据
    for rowdata in data:
        cvindex  = 0
        for cv in rowdata:
            worksheet.write(row, col + cvindex, cv)
            cvindex += 1
        row += 1
    workbook.close()


# ------------------------- 主程序 -------------------------
def main():
    try:
        print("开始执行数据爬取任务...")
        print("提示: 按 Ctrl+C 可以中断程序并保存已抓取的数据")
        
        # 抓取数据
        sheetdata = crawler_ccgp([], str(datetime.now().year), '')

        # 加载历史数据
        existing_data, headers = load_existing_data("existing_data.xlsx")
        existing_titles = get_existing_titles(existing_data)

        # 过滤重复数据
        filtered_data = filter_duplicates(sheetdata, existing_titles)

        # 保存结果
        head = ['序号', '类型', '名称', '日期', '招标人', '代理机构', '区域', '详情', '项目概况']
        if filtered_data:
            # 发送邮件通知
            email_body = generate_email_body(filtered_data)
            send_email("[招标公告更新提醒]发现新数据", email_body)

            # 保存新数据到Excel
            output_filename = "filtered_data_" + datetime.now().strftime("%Y%m%d_%H%M%S")
            writer_excel(filtered_data, head, '中标公告', output_filename)
            print(f"发现 {len(filtered_data)} 条新数据，已发送邮件通知并保存到Excel文件。")
        else:
            print("未发现新数据，未发送邮件。")

        print(f"原始数据条数: {len(sheetdata)}")
        print(f"过滤后数据条数: {len(filtered_data)}")
        print("任务完成!")
        
    except KeyboardInterrupt:
        print("程序被用户中断")
    except Exception as e:
        print(f"程序执行过程中发生错误: {e}")


if __name__ == "__main__":
    main()