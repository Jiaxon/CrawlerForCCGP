import pandas as pd
from datetime import datetime
import ccgp_get

# 读取另一份文档
def load_existing_data(file_path):
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        print(f"读取文件失败: {e}")
        return None

# 提取关键字段
def get_existing_titles(df):
    if df is not None and '名称' in df.columns:
        return set(df['名称'].tolist())  # 转换为集合，方便快速查找
    return set()

# 过滤重复数据
def filter_duplicates(sheetdata, existing_titles):
    filtered_data = []
    for row in sheetdata:
        title = row[2]  # 假设名称在第 3 列（索引为 2）
        if title not in existing_titles:
            filtered_data.append(row)
    return filtered_data

def main():
    # 抓取数据
    sheetdata = ccgp_get.crawler_ccgp([], str(datetime.now().year), '')

    # 读取另一份文档
    existing_df = load_existing_data("existing_data.xlsx")
    existing_titles = get_existing_titles(existing_df)

    # 过滤重复数据
    filtered_data = filter_duplicates(sheetdata, existing_titles)

    # 保存结果
    head = ['序号', '类型', '名称', '日期', '招标人', '代理机构', '区域', '详情', '项目概况']
    ccgp_get.writer_excel(filtered_data, head, '中标公告', 'filtered_data')

    print(f"原始数据条数: {len(sheetdata)}")
    print(f"过滤后数据条数: {len(filtered_data)}")

if __name__ == "__main__":
    main()