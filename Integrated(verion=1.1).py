# -*- coding=utf-8 -*-
import math
import time
import requests
from lxml import etree
from chardet import detect
from datetime import datetime, timedelta
import random
import pandas as pd
import xlsxwriter
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from selenium import webdriver
from selenium.webdriver.common.by import By

# ------------------------- 配置文件 -------------------------
# 邮件服务器配置
SMTP_SERVER = "smtp.example.com"  # SMTP服务器地址
SMTP_PORT = 465  # SMTP端口
SENDER_EMAIL = "sender@example.com"  # 发件邮箱
SENDER_PASSWORD = "your_password"  # 发件邮箱密码
RECEIVER_EMAIL = "receiver@example.com"  # 收件邮箱

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
    time.sleep(random.randint(2, 6))  # 随机延迟，避免频繁请求
    response = requests.get(url, headers=headers, params=params, allow_redirects=True)
    if response.status_code != 200:
        print(f"请求失败: {response.status_code}")
    return response


def crawler_ccgp(sheetdata=[], year='', buyerName=''):
    """爬取中国政府采购网的招标公告数据"""
    url = 'http://search.ccgp.gov.cn/bxsearch?'
    curr_date = datetime.now()
    start_date = curr_date - timedelta(days=30)  # 默认抓取最近30天的数据
    start_time = start_date.strftime("%Y:%m:%d")
    end_time = curr_date.strftime("%Y:%m:%d")

    params = {
        'searchtype': 1,
        'page_index': 1,
        'bidSort': 0,
        'buyerName': buyerName,
        'projectId': '',
        'pinMu': 0,
        'bidType': 0,
        'dbselect': 'bidx',
        'kw': '等级保护',  # 搜索关键词
        'start_time': start_time,
        'end_time': end_time,
        'timeType': 6,
        'displayZone': '',  # 目标区域
        'zoneId': '',
        'pppStatus': 0,
        'agentName': ''
    }

    resp = open_url(url, params)
    resp.raise_for_status()  # 检查响应状态
    html = resp.content.decode(detect(resp.content).get('encoding', 'utf-8'))
    tree = etree.HTML(html)

    try:
        total = int(tree.xpath('/html/body/div[5]/div[1]/div/p[1]/span[2]')[0].text.strip())  # 获取总数据量
    except IndexError:
        print("未找到数据总数")
        return sheetdata

    if total > 0:
        pagesize = math.ceil(total / 20)  # 计算总页数
        for curr_page in range(1, pagesize + 1):
            list = tree.xpath('/html/body/div[5]/div[2]/div/div/div[1]/ul/li')
            for li in list:
                title = li[0]
                summary = li[1]
                span = li[2]
                info = span.xpath('string()').replace(' ', '').replace('\r', '').replace('\n', '').replace('\t', '')

                str1 = info[:info.index('公告')]
                str2 = info[info.index('公告'):].replace('公告', '')
                strs = str2.split('|')

                if len(strs) > 1:
                    row = [len(sheetdata) + 1, '公告', title.text.strip()]
                    # 使用Selenium获取详情链接
                    driver = webdriver.Chrome()
                    driver.get(resp.url)
                    link_element = driver.find_element(By.XPATH, f"//a[contains(text(), '{title.text.strip()}')]")
                    link_href = link_element.get_attribute("href")
                    driver.quit()

                    row.extend([
                        str1.split('|')[0][:10],  # 日期
                        str1.split('|')[1].replace('采购人：', ''),  # 招标人
                        str1.split('|')[2].replace('代理机构：', ''),  # 代理机构
                        strs[0],  # 区域
                        link_href,  # 详情链接
                        summary.text.strip()  # 项目概况
                    ])
                    sheetdata.append(row)

            if curr_page < pagesize:
                params['page_index'] = curr_page + 1
                resp = open_url(url, params, resp.url)
                html = resp.content.decode(detect(resp.content).get('encoding', 'utf-8'))
                tree = etree.HTML(html)

    return sheetdata


# ------------------------- 数据处理模块 -------------------------
def load_existing_data(file_path):
    """加载已有的Excel数据"""
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        print(f"读取文件失败: {e}")
        return None


def get_existing_titles(df):
    """提取已有数据的标题"""
    if df is not None and '名称' in df.columns:
        return set(df['名称'].tolist())
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
    """生成HTML格式的邮件内容"""
    if not new_data:
        return None

    columns = ['序号', '类型', '名称', '日期', '招标人', '代理机构', '区域', '详情', '项目概况']
    df = pd.DataFrame(new_data, columns=columns)
    html_table = df.to_html(index=False, border=0, classes="data-table")

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
    # 抓取数据
    sheetdata = crawler_ccgp([], str(datetime.now().year), '')

    # 加载历史数据
    existing_df = load_existing_data("existing_data.xlsx")
    existing_titles = get_existing_titles(existing_df)

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


if __name__ == "__main__":
    main()