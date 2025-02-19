# -*- coding=utf-8 -*-
import pandas as pd
from datetime import datetime
import xlsxwriter
import ccgp_get
# ------------------------- email -------------------------
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd
from datetime import datetime
from config import SMTP_SERVER, SMTP_PORT, SENDER_EMAIL, SENDER_PASSWORD, RECEIVER_EMAIL


# ------------------------- docprocess -------------------------
def log_document_name(filename, log_file="document_log.txt"):
    with open(log_file, "a") as f:
        f.write(filename + "\n")


# read document name
def read_document_names(log_file="document_log.txt"):
    try:
        with open(log_file, "r") as f:
            return [line.strip() for line in f.readlines()]
    except FileNotFoundError:
        return []


# load historical data
def load_historical_data(log_file="document_log.txt"):
    document_names = read_document_names(log_file)
    historical_data = []
    for doc in document_names:
        try:
            df = pd.read_excel(doc + ".xlsx", sheet_name=0)
            historical_data.extend(df.to_dict("records"))
        except Exception as e:
            print(f"load document {doc} failed: {e}")
    return historical_data


# Filter duplicate data
def filter_duplicates(new_data, historical_data):
    historical_titles = set(row["名称"] for row in historical_data)  # 假设 "名称"是唯一标识
    filter_data = [row for row in new_data if row[2] not in historical_titles]  # 假设名称在第 3 列
    return filter_data


def writer_excel(data, head, sheetname, filename):
    workbook = xlsxwriter.Workbook(filename + ".xlsx")
    worksheet = workbook.add_worksheet(sheetname)

    # insert sheet head
    for col, header in enumerate(head):
        worksheet.write(0, col, header)

    # insert sheet data
    for row, rowdata in enumerate(data, start=1):
        for col, value in enumerate(rowdata):
            worksheet.write(row, col, value)

    workbook.close()
    log_document_name(filename)  # recorde document name


# ------------------------- email -------------------------
def send_email(subject, body):
    """send email"""
    msg = MIMEMultipart()
    msg["From"] = SENDER_EMAIL
    msg["To"] = RECEIVER_EMAIL
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "html"))

    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.sendmail(SENDER_EMAIL, RECEIVER_EMAIL, msg.as_string())
            print("send email successed!")
    except Exception as e:
        print(f"send failed: {e}")


def generate_email_body(new_data):
    """Generate message body(styled HTML table)"""
    if not new_data:
        return None

    # transfer to DataFrame
    columns = ['序号', '类型', '名称', '日期', '招标人', '代理机构', '区域', '详情', '项目概况']
    df = pd.DataFrame(new_data, columns=columns)

    # generate HTML sheet
    html_table = df.to_html(index=False, border=0, classes="data-table")

    # Add css styling
    style = """
    <style>
        .data-table {
            width: 100%;
            border-collapse: collapse;
            font-family: Arial, sans-serif;
            font-size: 14px;
            color: #333;
        }
        .data-table th, .data-table td {
            padding: 12px;
            text-align: left;
            border-bottom: 1px solid #ddd;
        }
        .data-table th {
            background-color: #f8f9fa;
            font-weight: bold;
            color: #333;
        }
        .data-table tr:hover {
            background-color: #f1f1f1;
        }
        .data-table a {
        color: #007bff;
        text-decoration: none;
        }
        .data-table a:hover {
            text-decoration: underline;
        }
    </style>
    """

    # Message body template
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


# ------------------------- main -------------------------
def main():
    # catch data
    sheetdata = ccgp_get.crawler_ccgp([], str(datetime.now().year), '')

    # load historical data
    historical_data = load_historical_data()

    # filter duplicate data
    filtered_data = filter_duplicates(sheetdata, historical_data)

    # save result
    head = ['序号 ', '类型', '名称', '日期', '招标人', '代理机构', '区域', '详情', '项目概况']

    if filtered_data:
        # generate the content of the email &  send it
        email_body = generate_email_body(filtered_data)
        send_email("[招标公告更新提醒]发现新数据", email_body)

        # save new data
        output_filename = "filtered_data_" + datetime.now().strftime("%Y%m%d_%H%M%S")
        writer_excel(filtered_data, head, '中标公告', output_filename)
        print(f"{len(filtered_data)} new data has been discovered, and an email notification has been sent")
    else:
        print("No new data was found, no email was sent")
    print(f"原始数据条数: {len(sheetdata)}")
    print(f"过滤后数据条数: {len(filtered_data)}")


if __name__ == "__main__":
    main()
