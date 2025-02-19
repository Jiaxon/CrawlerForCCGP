# 中国政府采购网(`http://www.ccgp.gov.cn` )爬虫 

## 项目概述
本项目用于自动抓取中国政府采购网招标公告数据, 实现数据清洗、去重存储和邮件通知功能。  
系统包含数据爬取、数据处理、邮件通知三大核心模块.

---

## 模块说明

### 1. 核心模块 (`Integrated(version=1.0).py`)  
- **功能**:
  - 数据爬取: 调用`ccgp_get.py`模块获取最新招标数据
  - 历史数据加载: 通过`document_log.txt`记录文件读取历史数据
  - 数据去重: 基于"名称"字段过滤重复数据
  - 数据存储: 生成带时间戳的Excel文件
  - 邮件通知: 通过HTML格式邮件发送新数据通知

- **关键方法**：
  ```python
  def main():
      # 可修改年份参数控制抓取范围
      sheetdata = ccgp_get.crawler_ccgp([], str(datetime.now().year), '')
  
--- 

### 2. 数据爬取模块 (`ccgp_get.py`)  
- **功能**:
  - 自动生成随机请求头绕过反爬机制
  - 支持多页爬取和时间范围过滤
  - 自动提取公告详情链接

- **关键参数**：
  ```python
  params = {
        'kw': '',        # 搜索关键词
        'displayZone': '',   # 目标区域
        'start_time': '2025:02:01',  # 开始日期
        'end_time': '2025:02:14'     # 结束日期
    }  

---

### 3. 数据处理模块 (`DataProcessing.py`)  
- **功能**:
  - 与现有数据文件比对去重
  - 生成过滤后的Excel文件

- **使用方法**：
  ```python
  existing_df = load_existing_data("existing_data.xlsx")  # 自定义历史数据文件路径  

--- 

### 4. 配置文件 (`config.py`) 
- **配置项**：
  ```python
  SMTP_SERVER = "smtp.example.com"    # SMTP服务器地址
  SENDER_EMAIL = "sender@example.com" # 发件邮箱
  RECEIVER_EMAIL = "receiver@example.com" # 收件邮箱
  
---

## 运行指南

### 环境要求
- Python 3.8+   
- 依赖库：pandas xlsxwriter requests selenium lxml

### 快速启动

### &nbsp;&nbsp;1.配置SMTP
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;编辑`config.py`填写邮箱服务信息

### &nbsp;&nbsp;2.首次运行
  ```bash
  python Integrated(version=1.0).py  
  ```

### &nbsp;&nbsp;3.自定义爬取参数
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;修改`ccgp_get.py`中的爬取参数：
  ```python
  # 修改搜索关键词
  params['kw'] = '网络安全'
  
  # 修改目标区域
  params['displayZone'] = '北京'
  params['zoneId'] = '11'  # 区域编码参考网站实际值
  ```

### &nbsp;&nbsp;4.调整时间范围
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;修改`crawler_ccgp`函数中的时间计算逻辑：
  ```python
  # 默认抓取最近30天数据
  start_date = curr_date - timedelta(days=30)  # 修改days值调整范围
  ```

---

## 输出文件
- 生成的Excel文件命名格式：`filtered_data_YYYYMMDD_HHMMSS.xlsx`
- 日志文件：`document_log.txt` 记录所有生成的文件名

---
## 注意事项
1. 遵守robots.txt协议控制爬取频率    
2. 每日运行次数建议不超过3次    
3. 若需长期运行，建议设置定时任务(如crontab)  
4. 使用前请确认代理机构名称字段是否需要特殊处理  
5. 本爬虫程序仅供个人学习、研究和技术实验使用，不得用于商业目的、恶意用途或违反任何相关法律法规的行为。
6. 使用本爬虫程序的用户应确保符合目标网站的使用条款、隐私政策及相关法律规定。用户应自行承担因爬取数据而产生的所有法律责任，程序开发者不对由此产生的任何法律纠纷或经济损失负责。
6. 使用本程序时，用户应遵循目标网站的robots.txt文件和相关的访问限制规定，不得爬取未经授权的数据。若目标网站明确禁止爬虫访问，用户应立即停止相关操作。
7. 本程序不会收集、存储或共享任何爬取的数据，所有数据仅限于用户个人使用。若有数据泄露或滥用的行为，程序开发者不承担任何责任。
8. 本程序是开源工具，使用过程中可能存在技术风险或安全隐患，用户应自担风险，确保程序使用符合所有相关的法律法规。
---
## 文件结构
```
project/
├── Integrated(version=1.0).py  # 主程序
├── ccgp_get.py                 # 数据爬取模块
├── DataProcessing.py           # 数据处理模块
├── config.py                   # 配置文件
├── document_log.txt            # 日志文件
├── filtered_data_*.xlsx        # 生成的Excel文件
└── README.md                   # 项目说明文档
```

---

## 示例运行
### 1. 抓取指定关键词数据
修改`ccgp_get.py`中的`params`参数：
```python
params = {
    'kw': '网络安全',  # 修改关键词
    'displayZone': '北京',  # 修改区域
    'start_time': '2025:02:01',  # 修改开始时间
    'end_time': '2025:02:14'     # 修改结束时间
}
```
### 2. 运行主程序
```bash
python Integrated(version=1.0).py
```
### 3. 查看结果
- 新数据保存在`filtered_data_YYYYMMDD_HHMMSS.xlsx`中
- 日志记录在`document_log.txt`中
- 新数据通知邮件将发送至配置的邮箱

---

### 依赖包说明
1. `pandas`：
- 用于数据处理和分析，支持将数据存储为 Excel 文件。
2. `xlsxwriter`：
- 用于生成 Excel 文件，支持格式化、图表等功能。
3. `requests`：
- 用于发送 HTTP 请求，抓取网页内容。
4. `lxml`：
- 用于解析 HTML 和 XML 文档，提取所需数据。
5. `selenium`：
- 用于处理动态网页内容，支持浏览器自动化操作。
6. `beautifulsoup4`：
- 用于解析 HTML 文档，提取结构化数据。
7. `smtplib` 和 `email`：
- 用于发送邮件通知，支持 HTML 格式的邮件内容。
8. `numpy`：
- 用于数值计算，支持高效的数据处理。
9. `openpyxl`：
- 用于读取和写入 Excel 文件，支持复杂的数据操作。
10. `python-dateutil`：
- 用于处理日期和时间，支持日期格式化和计算。
11. `chardet`：
- 用于检测网页编码，确保正确解析网页内容。

---
  
## 部分更新描述(version = 1.1)

- 将需要用到的方法进行整理并合并到Integrated(verion=1.1).py文件中
- 更新并补充了详细的注释，方便理解每个模块的作用和使用方法。

--- 

## 联系方式  
如有问题或建议，请联系:  
- 邮箱：commonboeotian@gmail.com
