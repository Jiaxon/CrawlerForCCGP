#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import json
import os
import math
import time
import random
from datetime import datetime, timedelta

# 检查依赖
missing_modules = []

try:
    import requests
except ImportError:
    missing_modules.append("requests")

try:
    from lxml import etree
except ImportError:
    missing_modules.append("lxml")

try:
    import xlsxwriter
except ImportError:
    missing_modules.append("xlsxwriter")

try:
    from PyQt6.QtWidgets import (
        QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QLineEdit, 
        QPushButton, QTextEdit, QGroupBox, QSpinBox, QComboBox, QTabWidget, QFormLayout,
        QDateEdit, QProgressBar, QStatusBar, QCheckBox, QFileDialog, QMessageBox
    )
    from PyQt6.QtCore import QThread, pyqtSignal, QObject, QDate
except ImportError:
    missing_modules.append("PyQt6")

if missing_modules:
    print("缺少以下必要模块:")
    for module in missing_modules:
        print(f"  - {module}")
    print("\n请运行以下命令安装:")
    print("pip install requests lxml xlsxwriter PyQt6")
    sys.exit(1)

# 如果所有模块都可用，继续执行
print("所有依赖模块检查通过!")

# ------------------------- Worker类 -------------------------
class Worker(QObject):
    """
    将爬虫逻辑放在一个单独的QObject中，以便可以移动到QThread中执行，防止UI阻塞。
    """
    finished = pyqtSignal()
    progress_update = pyqtSignal(str)
    progress_bar_update = pyqtSignal(int, int)
    error = pyqtSignal(str)
    data_saved = pyqtSignal(str)

    def __init__(self, config):
        super().__init__()
        self.config = config
        self.is_running = True
        self.current_crawled_data = []
        self.session = requests.Session()  # 使用会话管理网络连接
        
        # 禁用代理以避免连接问题
        self.session.trust_env = False
        
        # 根据配置设置代理
        if config.get('use_proxy', False):
            proxy_host = config.get('proxy_host', '127.0.0.1')
            proxy_port = config.get('proxy_port', 7890)
            proxy_url = f"http://{proxy_host}:{proxy_port}"
            self.session.proxies = {
                'http': proxy_url,
                'https': proxy_url
            }
        else:
            self.session.proxies = {}

    def stop(self):
        self.progress_update.emit("正在请求停止...")
        self.is_running = False
        # 关闭网络会话
        if hasattr(self, 'session'):
            try:
                self.session.close()
            except:
                pass

    def run(self):
        """执行爬虫任务"""
        try:
            self.progress_update.emit("开始执行数据爬取任务...")
            
            # 1. 抓取数据
            self.current_crawled_data = self._crawler_ccgp_threaded()
            
            if not self.is_running:
                self._save_interrupted_data()
                self.progress_update.emit("任务已手动停止。")
                self.finished.emit()
                return
            
            self.progress_update.emit("数据抓取完成，开始处理数据...")
            
            head = ['序号', '关键字', '名称', '日期', '采购人', '代理机构', '公告类型', '详情', '项目概况']
            
            # 根据是否有新数据，决定后续操作
            if self.current_crawled_data:
                self.progress_update.emit(f"共抓取到 {len(self.current_crawled_data)} 条数据。")
                
                # 自动保存
                if self.config.get('auto_save', True):
                    output_filename = self.config.get('output_prefix', 'filtered_data_') + datetime.now().strftime("%Y%m%d_%H%M%S")
                    self._writer_excel(self.current_crawled_data, head, output_filename)
                    self.data_saved.emit(f"数据已保存到 {output_filename}.xlsx")
            else:
                self.progress_update.emit("未抓取到任何数据。")
            
            self.progress_update.emit(f"本次共抓取数据条数: {len(self.current_crawled_data)}")
            self.progress_update.emit("任务完成!")
            
        except Exception as e:
            self.error.emit(f"程序执行过程中发生错误: {e}")
        finally:
            # 确保在结束时关闭网络会话
            if hasattr(self, 'session'):
                try:
                    self.session.close()
                except:
                    pass
            self.finished.emit()

    def _save_interrupted_data(self):
        if self.current_crawled_data:
            self.progress_update.emit("正在保存已抓取的数据...")
            head = ['序号', '关键字', '名称', '日期', '采购人', '代理机构', '公告类型', '详情', '项目概况']
            output_filename = "interrupted_data_" + datetime.now().strftime("%Y%m%d_%H%M%S")
            self._writer_excel(self.current_crawled_data, head, output_filename)
            self.data_saved.emit(f"已保存 {len(self.current_crawled_data)} 条数据到 {output_filename}.xlsx")
        else:
            self.progress_update.emit("没有数据需要保存。")

    def _get_bid_type_name(self, bid_type_code):
        """根据公告类型代码获取对应的名称"""
        bid_type_map = {
            "0": "所有", "1": "公开招标", "2": "询价公告", "3": "竞争性谈判",
            "4": "单一来源", "5": "资格预审", "6": "邀请公告", "7": "中标公告",
            "8": "更正公告", "9": "其他公告", "10": "竞争性磋商", "11": "成交公告",
            "12": "废标公告"
        }
        return bid_type_map.get(bid_type_code, "未知类型")

    def _get_request_headers(self, referer=None):
        user_agents = [
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36',
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/119.0'
        ]
        ua = random.choice(user_agents)
        headers = {
            "User-Agent": ua, 
            "Host": "search.ccgp.gov.cn",
            "Referer": referer if referer else "http://search.ccgp.gov.cn/",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
            "Accept-Encoding": "gzip, deflate", 
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
            "Connection": "keep-alive", 
            "Upgrade-Insecure-Requests": "1",
            "Cache-Control": "max-age=0"
        }
        return headers

    def _open_url(self, url, params, refer=None):
        headers = self._get_request_headers(refer)
        delay_seconds = 3  # 固定延迟
        self.progress_update.emit(f"等待 {delay_seconds} 秒...")
        time.sleep(delay_seconds)
        
        try:
            response = self.session.get(url, headers=headers, params=params, timeout=30)
            return response
        except Exception as e:
            self.progress_update.emit(f"网络错误: {str(e)[:30]}")
            raise

    def _crawler_ccgp_threaded(self):
        sheetdata = []
        url = 'http://search.ccgp.gov.cn/bxsearch?'
        
        # 使用GUI传入的日期
        start_date_str = self.config['start_date']
        end_date_str = self.config['end_date']
        
        # 添加调试信息
        self.progress_update.emit(f"使用时间范围: {start_date_str} 至 {end_date_str}")
        
        # 使用最简单的参数格式
        params = {
            'searchtype': 1, 
            'page_index': 1, 
            'bidSort': 0,
            'buyerName': self.config.get('buyer_name', ''), 
            'projectId': '', 
            'pinMu': 0, 
            'bidType': self.config.get('bid_type', '0'),
            'dbselect': 'bidx', 
            'kw': self.config.get('keyword', ''),  # 不设置默认值
            'start_time': start_date_str,
            'end_time': end_date_str,
            'timeType': self.config.get('time_type', 6),  # 使用配置中的timeType
            'displayZone': '',
            'zoneId': self.config.get('zone_id', ''),
            'pppStatus': 0, 
            'agentName': self.config.get('agent_name', '')
        }
        
        # 简单清理空参数，但保留重要参数
        cleaned_params = {}
        for k, v in params.items():
            if k == 'kw':  # 关键字参数处理
                # 如果用户没有输入关键字，使用空字符串而不是默认值
                user_keyword = self.config.get('keyword', '').strip()
                cleaned_params[k] = user_keyword if user_keyword else ''
            elif k == 'dbselect':  # dbselect必须是字符串
                cleaned_params[k] = 'bidx'
            elif k in ['searchtype', 'page_index', 'bidSort', 'pinMu', 'pppStatus', 'timeType']:
                # 这些参数转换为整数
                try:
                    cleaned_params[k] = int(v) if v != '' else 0
                except (ValueError, TypeError):
                    cleaned_params[k] = 0
            elif k == 'bidType':
                # bidType保持字符串格式
                cleaned_params[k] = str(v) if v is not None else '0'
            elif k in ['start_time', 'end_time']:  # 时间参数必须保留
                cleaned_params[k] = v
            elif v is not None and str(v).strip():
                cleaned_params[k] = v
        
        # 添加参数调试信息
        self.progress_update.emit(f"API参数: {cleaned_params}")
        
        # 构建并显示完整的API URL
        from urllib.parse import urlencode
        full_url = url + urlencode(cleaned_params)
        self.progress_update.emit(f"完整API URL: {full_url}")
        
        try:
            self.progress_update.emit("开始获取数据...")
            resp = self._open_url(url, cleaned_params)
            if not self.is_running: return sheetdata
            resp.raise_for_status()
            html = resp.content.decode('utf-8')
            tree = etree.HTML(html)
            
            total_text = tree.xpath('/html/body/div[5]/div[1]/div/p[1]/span[2]/text()')
            total = int(total_text[0].strip()) if total_text else 0
            self.progress_update.emit(f"找到 {total} 条数据")
            
            if total > 0:
                pagesize = math.ceil(total / 20)
                self.progress_update.emit(f"总共 {pagesize} 页数据需要抓取")
                
                for curr_page in range(1, pagesize + 1):
                    if not self.is_running: break
                    self.progress_update.emit(f"正在抓取第 {curr_page}/{pagesize} 页数据...")
                    self.progress_bar_update.emit(curr_page, pagesize)
                    
                    # 获取下一页内容
                    if curr_page > 1:
                        params['page_index'] = curr_page
                        resp = self._open_url(url, params, resp.url)
                        if not self.is_running: break
                        html = resp.content.decode('utf-8')
                        tree = etree.HTML(html)
                    
                    list_items = tree.xpath('/html/body/div[5]/div[2]/div/div/div[1]/ul/li')
                    for i, li in enumerate(list_items):
                        if not self.is_running: break
                        try:
                            # 添加空值检查
                            title_element = li.find('a')
                            summary_element = li.find('p')
                            span_element = li.find('span')
                            
                            if title_element is None or summary_element is None or span_element is None: 
                                self.progress_update.emit("  跳过无效数据项")
                                continue
                                
                            # 安全获取文本内容
                            title = title_element.text.strip() if title_element.text else ''
                            if not title:
                                self.progress_update.emit("  跳过空标题项")
                                continue
                                
                            link_href = title_element.get('href', '')
                            summary = summary_element.text.strip() if summary_element.text else ''
                            
                            # 安全获取span内容
                            span_text = span_element.xpath('string()')
                            if not span_text:
                                self.progress_update.emit("  跳过空span项")
                                continue
                                
                            info = span_text.replace(' ', '').replace('\r', '').replace('\n', '').replace('\t', '')
                            
                            if len(info) < 10:
                                self.progress_update.emit("  跳过信息不完整项")
                                continue
                            
                            date_part = info[:10]
                            remaining_info = info[10:]
                            
                            # 初始化
                            buyer_part = ''
                            agent_part = ''
                            region_part = ''
                            
                            # 调试信息 - 减少输出频率
                            if i % 5 == 0:  # 每5条记录输出一次
                                self.progress_update.emit(f"  解析数据: {remaining_info[:30]}...")
                            
                            # 使用更精确的方式解析
                            # 先找到所有标识位置
                            buyer_pos = remaining_info.find('采购人：')
                            agent_pos = remaining_info.find('代理机构：')
                            
                            # 处理采购人
                            if buyer_pos != -1:
                                buyer_start = buyer_pos + 4
                                # 找到下一个分隔符的位置
                                next_sep = remaining_info.find('|', buyer_start)
                                if next_sep != -1:
                                    buyer_part = remaining_info[buyer_start:next_sep].strip()
                                else:
                                    # 如果没有|，看是否有代理机构标识
                                    if agent_pos > buyer_pos:
                                        buyer_part = remaining_info[buyer_start:agent_pos].strip()
                                    else:
                                        buyer_part = remaining_info[buyer_start:].strip()
                            
                            # 处理代理机构
                            if agent_pos != -1:
                                agent_start = agent_pos + 5
                                # 找到下一个分隔符的位置
                                next_sep = remaining_info.find('|', agent_start)
                                if next_sep != -1:
                                    agent_part = remaining_info[agent_start:next_sep].strip()
                                else:
                                    agent_part = remaining_info[agent_start:].strip()
                            
                            # 处理区域信息 - 从最后一个|开始的部分
                            last_pipe = remaining_info.rfind('|')
                            if last_pipe != -1:
                                potential_region = remaining_info[last_pipe + 1:].strip()
                                # 确保这部分不包含采购人或代理机构标识
                                if '采购人：' not in potential_region and '代理机构：' not in potential_region:
                                    region_part = potential_region
                            
                            # 调试输出解析结果 - 减少输出频率
                            if i % 5 == 0:  # 每5条记录输出一次
                                self.progress_update.emit(f"    解析结果: 采购人={buyer_part[:20]}, 代理={agent_part[:20]}, 区域={region_part[:20]}")
                            
                            # 获取公告类型名称
                            bid_type_name = self._get_bid_type_name(self.config.get('bid_type', '0'))
                            
                            # 获取搜索关键字
                            search_keyword = self.config.get('keyword', '')
                            
                            # 更新数据行结构: 序号、关键字、名称、日期、采购人、代理机构、公告类型、详情、项目概况
                            row = [len(sheetdata) + 1, search_keyword, title, date_part, buyer_part, agent_part, bid_type_name, link_href, summary]
                            sheetdata.append(row)
                            
                            # 减少日志输出频率
                            if i % 10 == 0:  # 每10条记录输出一次
                                self.progress_update.emit(f"  已获取第 {i+1} 条数据: {title[:20]}...")
                        except (ValueError, IndexError, AttributeError) as e:
                            self.progress_update.emit(f"解析数据时出错，跳过此条记录: {e}")
                            continue
                        except Exception as e:
                            self.progress_update.emit(f"处理数据时出现未知错误，跳过此条记录: {e}")
                            continue
                            
        except Exception as e:
            self.error.emit(f"抓取数据时发生错误: {e}")
        return sheetdata

    def _writer_excel(self, data, head, filename):
        # 获取保存路径
        save_path = self.config.get('save_path', '')
        if save_path and save_path.strip():
            full_path = os.path.join(save_path, filename + '.xlsx')
        else:
            full_path = filename + '.xlsx'
        
        workbook = xlsxwriter.Workbook(full_path)
        worksheet = workbook.add_worksheet("中标公告")
        for cvi, cv in enumerate(head):
            worksheet.write(0, cvi, cv)
        for row_idx, rowdata in enumerate(data, start=1):
            for col_idx, cell_data in enumerate(rowdata):
                worksheet.write(row_idx, col_idx, cell_data)
        workbook.close()


# ------------------------- PyQt6 GUI 主窗口 -------------------------
class MainWindow(QMainWindow):
    CONFIG_FILE = "config.json"

    def __init__(self):
        super().__init__()
        self.worker = None
        self.thread = None
        self.crawled_data = []
        self.init_ui()
        self.load_config()

    def cleanup_resources(self):
        """清理资源"""
        try:
            if self.thread and self.thread.isRunning():
                if self.worker:
                    self.worker.stop()
                self.thread.quit()
                self.thread.wait(5000)  # 增加等待时间到5秒
                if self.thread.isRunning():
                    self.thread.terminate()
                    self.thread.wait(2000)  # 等待终止完成
            
            # 断开所有信号连接
            if self.worker:
                try:
                    self.worker.disconnect()
                except:
                    pass
                self.worker.deleteLater()
                self.worker = None
            if self.thread:
                self.thread.deleteLater()
                self.thread = None
        except Exception as e:
            print(f"清理资源时出错: {e}")
            pass

    def init_ui(self):
        self.setWindowTitle('中国政府采购网公告爬虫')
        self.setGeometry(100, 100, 900, 700)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Tab Widget
        self.tab_widget = QTabWidget()
        main_layout.addWidget(self.tab_widget)

        # Create Tabs - 去掉邮件标签页
        self.crawler_tab = self._create_crawler_tab()
        self.advanced_tab = self._create_advanced_tab()

        self.tab_widget.addTab(self.crawler_tab, "爬虫设置")
        self.tab_widget.addTab(self.advanced_tab, "高级设置")

        # Status Bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("就绪")

        # Progress Bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.status_bar.addPermanentWidget(self.progress_bar)

    def _create_crawler_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Search Parameters Group - 3列布局
        search_group = QGroupBox("搜索参数")
        search_layout = QGridLayout()

        # 第一列：关键词和采购人名称
        search_layout.addWidget(QLabel("关键词:"), 0, 0)
        self.keyword_input = QLineEdit()  # 移除默认值
        search_layout.addWidget(self.keyword_input, 1, 0)

        # 采购人名称标签和输入框在同一行
        buyer_container = QWidget()
        buyer_layout = QHBoxLayout(buyer_container)
        buyer_layout.setContentsMargins(0, 0, 0, 0)
        buyer_layout.addWidget(QLabel("采购人:"))
        self.buyer_name_input = QLineEdit()
        buyer_layout.addWidget(self.buyer_name_input)
        search_layout.addWidget(buyer_container, 2, 0)

        # 时间预设
        search_layout.addWidget(QLabel("时间预设:"), 3, 0)
        self.time_preset_combo = QComboBox()
        time_presets = [
            ("自定义", "custom"),
            ("今天", "today"), 
            ("三天内", "3days"),
            ("一周内", "1week"),
            ("一月内", "1month"),
            ("三月内", "3months"),
            ("半年内", "6months")
        ]
        for name, value in time_presets:
            self.time_preset_combo.addItem(name, value)
        self.time_preset_combo.setCurrentIndex(0)  # Default to "自定义"
        self.time_preset_combo.currentIndexChanged.connect(self._on_time_preset_changed)
        search_layout.addWidget(self.time_preset_combo, 4, 0)

        # 第二列：公告类型和代理机构名称
        search_layout.addWidget(QLabel("公告类型:"), 0, 1)
        self.bid_type_combo = QComboBox()
        bid_types = [
            ("所有", "0"), ("公开招标", "1"), ("询价公告", "2"), ("竞争性谈判", "3"),
            ("单一来源", "4"), ("资格预审", "5"), ("邀请公告", "6"), ("中标公告", "7"),
            ("更正公告", "8"), ("其他公告", "9"), ("竞争性磋商", "10"), ("成交公告", "11"),
            ("废标公告", "12")
        ]
        for name, code in bid_types:
            self.bid_type_combo.addItem(name, code)
        self.bid_type_combo.setCurrentText("所有")
        search_layout.addWidget(self.bid_type_combo, 1, 1)

        # 代理机构名称标签和输入框在同一行
        agent_container = QWidget()
        agent_layout = QHBoxLayout(agent_container)
        agent_layout.setContentsMargins(0, 0, 0, 0)
        agent_layout.addWidget(QLabel("代理机构:"))
        self.agent_name_input = QLineEdit("")
        agent_layout.addWidget(self.agent_name_input)
        search_layout.addWidget(agent_container, 2, 1)

        # 自定义时间
        search_layout.addWidget(QLabel("自定义时间:"), 3, 1)
        date_container = QWidget()
        date_v_layout = QVBoxLayout(date_container)
        date_v_layout.setContentsMargins(0, 0, 0, 0)
        
        self.start_date_input = QDateEdit()
        self.start_date_input.setDate(QDate.currentDate().addDays(-3))
        self.start_date_input.setCalendarPopup(True)
        date_v_layout.addWidget(self.start_date_input)
        
        date_v_layout.addWidget(QLabel("至"))
        
        self.end_date_input = QDateEdit()
        self.end_date_input.setDate(QDate.currentDate())
        self.end_date_input.setCalendarPopup(True)
        date_v_layout.addWidget(self.end_date_input)
        
        search_layout.addWidget(date_container, 4, 1)

        # 第三列：区域
        search_layout.addWidget(QLabel("区域:"), 0, 2)
        self.region_combo = QComboBox()
        regions = [
            ("全国", ""), ("北京", "11"), ("天津", "12"), ("河北", "13"), ("山西", "14"),
            ("内蒙古", "15"), ("辽宁", "21"), ("吉林", "22"), ("黑龙江", "23"), ("上海", "31"),
            ("江苏", "32"), ("浙江", "33"), ("安徽", "34"), ("福建", "35"), ("江西", "36"),
            ("山东", "37"), ("河南", "41"), ("湖北", "42"), ("湖南", "43"), ("广东", "44"),
            ("广西", "45"), ("海南", "46"), ("重庆", "50"), ("四川", "51"), ("贵州", "52"),
            ("云南", "53"), ("西藏", "54"), ("陕西", "61"), ("甘肃", "62"), ("青海", "63"),
            ("宁夏", "64"), ("新疆", "65")
        ]
        for name, code in regions:
            self.region_combo.addItem(name, code)
        self.region_combo.setCurrentText("广西")
        search_layout.addWidget(self.region_combo, 1, 2)

        search_group.setLayout(search_layout)
        layout.addWidget(search_group)

        # Output Settings Group - 简化布局
        output_group = QGroupBox("输出设置")
        output_layout = QGridLayout()

        # 保存路径
        output_layout.addWidget(QLabel("保存路径:"), 0, 0)
        save_path_container = QWidget()
        save_path_h_layout = QHBoxLayout(save_path_container)
        save_path_h_layout.setContentsMargins(0, 0, 0, 0)
        self.save_path_input = QLineEdit()
        self.save_path_input.setPlaceholderText("选择保存路径（留空则保存到当前目录）")
        self.browse_save_path_button = QPushButton("浏览...")
        self.browse_save_path_button.clicked.connect(self._browse_save_path)
        save_path_h_layout.addWidget(self.save_path_input)
        save_path_h_layout.addWidget(self.browse_save_path_button)
        output_layout.addWidget(save_path_container, 0, 1, 1, 2)

        # 输出文件前缀
        output_layout.addWidget(QLabel("输出文件前缀:"), 1, 0)
        self.output_prefix_input = QLineEdit("filtered_data_")
        output_layout.addWidget(self.output_prefix_input, 1, 1, 1, 2)

        output_group.setLayout(output_layout)
        layout.addWidget(output_group)

        # Control Buttons - 去掉邮件按钮
        button_group = QGroupBox("操作")
        button_layout = QHBoxLayout()
        self.start_button = QPushButton('🚀 开始抓取')
        self.start_button.setStyleSheet("QPushButton { font-weight: bold; padding: 8px 16px; }")
        self.stop_button = QPushButton('⏹️ 停止')
        self.stop_button.setEnabled(False)
        self.save_results_button = QPushButton('💾 保存结果')
        self.save_results_button.setEnabled(False)

        button_layout.addWidget(self.start_button)
        button_layout.addWidget(self.stop_button)
        button_layout.addWidget(self.save_results_button)
        button_layout.addStretch()
        button_group.setLayout(button_layout)
        layout.addWidget(button_group)

        # Log Output Area
        log_group = QGroupBox("运行日志")
        log_layout = QVBoxLayout()
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setMaximumHeight(200)
        log_layout.addWidget(self.log_output)
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)

        # Connect Signals
        self.start_button.clicked.connect(self._start_crawling)
        self.stop_button.clicked.connect(self._stop_crawling)
        self.save_results_button.clicked.connect(self._save_results)

        return tab

    def _create_advanced_tab(self):
        tab = QWidget()
        layout = QFormLayout(tab)

        delay_h_layout = QHBoxLayout()
        self.min_delay_input = QSpinBox()
        self.min_delay_input.setRange(1, 10)
        self.min_delay_input.setValue(2)
        self.max_delay_input = QSpinBox()
        self.max_delay_input.setRange(2, 20)
        self.max_delay_input.setValue(6)
        delay_h_layout.addWidget(self.min_delay_input)
        delay_h_layout.addWidget(QLabel("至"))
        delay_h_layout.addWidget(self.max_delay_input)
        delay_h_layout.addWidget(QLabel("秒"))
        layout.addRow("请求延迟:", delay_h_layout)

        self.auto_save_checkbox = QCheckBox("爬取完成后自动保存结果")
        self.auto_save_checkbox.setChecked(True)
        layout.addRow("", self.auto_save_checkbox)

        # 添加代理设置
        self.use_proxy_checkbox = QCheckBox("使用代理服务器")
        self.use_proxy_checkbox.setChecked(False)
        layout.addRow("网络设置:", self.use_proxy_checkbox)

        proxy_h_layout = QHBoxLayout()
        self.proxy_host_input = QLineEdit("127.0.0.1")
        self.proxy_port_input = QSpinBox()
        self.proxy_port_input.setRange(1, 65535)
        self.proxy_port_input.setValue(7890)
        proxy_h_layout.addWidget(QLabel("代理地址:"))
        proxy_h_layout.addWidget(self.proxy_host_input)
        proxy_h_layout.addWidget(QLabel("端口:"))
        proxy_h_layout.addWidget(self.proxy_port_input)
        layout.addRow("", proxy_h_layout)

        return tab

    def _log(self, message):
        self.log_output.append(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}")
        self.log_output.verticalScrollBar().setValue(self.log_output.verticalScrollBar().maximum())

    def _update_progress_bar(self, current, total):
        progress = int((current / total) * 100) if total > 0 else 0
        self.progress_bar.setValue(progress)

    def _browse_save_path(self):
        dir_path = QFileDialog.getExistingDirectory(
            self, "选择保存目录", ""
        )
        if dir_path:
            self.save_path_input.setText(dir_path)

    def _on_time_preset_changed(self):
        """时间预设选择改变时更新日期范围"""
        preset = self.time_preset_combo.currentData()
        if preset == "custom":
            return
        
        current_date = QDate.currentDate()
        
        if preset == "today":
            start_date = current_date
        elif preset == "3days":
            start_date = current_date.addDays(-3)
        elif preset == "1week":
            start_date = current_date.addDays(-7)
        elif preset == "1month":
            start_date = current_date.addMonths(-1)
        elif preset == "3months":
            start_date = current_date.addMonths(-3)
        elif preset == "6months":
            start_date = current_date.addMonths(-6)
        else:
            start_date = current_date.addDays(-3)
        
        self.start_date_input.setDate(start_date)
        self.end_date_input.setDate(current_date)

    def _get_time_type(self):
        """根据时间预设选择返回对应的timeType值"""
        preset = self.time_preset_combo.currentData()
        time_type_map = {
            "today": 0,      # 今日
            "3days": 1,      # 三日内
            "1week": 2,      # 一周内
            "1month": 3,     # 一月内
            "3months": 4,    # 三月内
            "6months": 5,    # 半年内
            "custom": 6      # 自定义时间
        }
        return time_type_map.get(preset, 6)  # 默认为自定义时间

    def _get_current_config(self):
        # 获取原始日期
        start_date = self.start_date_input.date()
        end_date = self.end_date_input.date()
        
        # 尝试不同的日期格式
        start_date_str = start_date.toString("yyyy-MM-dd")
        end_date_str = end_date.toString("yyyy-MM-dd")
        
        return {
            # Crawler Config
            "buyer_name": self.buyer_name_input.text(),
            "keyword": self.keyword_input.text(),
            "start_date": start_date_str,
            "end_date": end_date_str,
            "zone_id": self.region_combo.currentData(),
            "bid_type": self.bid_type_combo.currentData(),
            "save_path": self.save_path_input.text(),
            "output_prefix": self.output_prefix_input.text(),
            "agent_name": self.agent_name_input.text(),
            "time_type": self._get_time_type(),  # 添加timeType

            # Advanced Config
            "min_delay": self.min_delay_input.value(),
            "max_delay": self.max_delay_input.value(),
            "auto_save": self.auto_save_checkbox.isChecked(),
            
            # Proxy Config
            "use_proxy": self.use_proxy_checkbox.isChecked(),
            "proxy_host": self.proxy_host_input.text(),
            "proxy_port": self.proxy_port_input.value(),
        }

    def save_config(self):
        config = self._get_current_config()
        try:
            with open(self.CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
            self._log("配置已保存。")
        except Exception as e:
            self._log(f"保存配置失败: {e}")

    def load_config(self):
        if not os.path.exists(self.CONFIG_FILE):
            self._log("未找到配置文件，将使用默认设置。")
            self.start_date_input.setDate(QDate.currentDate().addDays(-3))
            self.end_date_input.setDate(QDate.currentDate())
            self.min_delay_input.setValue(2)
            self.max_delay_input.setValue(6)
            self.auto_save_checkbox.setChecked(True)
            return
        try:
            with open(self.CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)

            # Crawler Config
            self.buyer_name_input.setText(config.get("buyer_name", ""))
            self.keyword_input.setText(config.get("keyword", ""))
            self.start_date_input.setDate(QDate.fromString(config.get("start_date", QDate.currentDate().addDays(-3).toString("yyyy-MM-dd")), "yyyy-MM-dd"))
            self.end_date_input.setDate(QDate.fromString(config.get("end_date", QDate.currentDate().toString("yyyy-MM-dd")), "yyyy-MM-dd"))

            # Set ComboBoxes by data, not text
            zone_id = config.get("zone_id", "45")
            index = self.region_combo.findData(zone_id)
            if index != -1: self.region_combo.setCurrentIndex(index)

            bid_type = config.get("bid_type", "0")
            index = self.bid_type_combo.findData(bid_type)
            if index != -1: self.bid_type_combo.setCurrentIndex(index)

            self.save_path_input.setText(config.get("save_path", ""))
            self.output_prefix_input.setText(config.get("output_prefix", "filtered_data_"))
            self.agent_name_input.setText(config.get("agent_name", ""))

            # Advanced Config
            self.min_delay_input.setValue(config.get("min_delay", 2))
            self.max_delay_input.setValue(config.get("max_delay", 6))
            self.auto_save_checkbox.setChecked(config.get("auto_save", True))
            
            # Proxy Config
            self.use_proxy_checkbox.setChecked(config.get("use_proxy", False))
            self.proxy_host_input.setText(config.get("proxy_host", "127.0.0.1"))
            self.proxy_port_input.setValue(config.get("proxy_port", 7890))

            self._log("配置已加载。")
        except Exception as e:
            self._log(f"加载配置失败: {e}")

    def _start_crawling(self):
        """启动爬虫"""
        # 如果已有线程在运行，先清理
        if self.thread and self.thread.isRunning():
            self._log("停止当前运行的爬虫...")
            self._stop_crawling()
            self.thread.quit()
            self.thread.wait(2000)
        
        self.save_config()
        config = self._get_current_config()

        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.save_results_button.setEnabled(False)
        self.log_output.clear()
        self.progress_bar.setValue(0)
        self.status_bar.showMessage("爬虫运行中...")
        self._log("正在启动爬虫线程...")

        # 创建新的线程和worker
        self.thread = QThread()
        self.worker = Worker(config)
        self.worker.moveToThread(self.thread)

        # 连接信号
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self._crawler_finished)
        
        # 连接日志和进度信号
        self.worker.progress_update.connect(self._log)
        self.worker.progress_bar_update.connect(self._update_progress_bar)
        self.worker.error.connect(self._log)
        self.worker.data_saved.connect(self._log)

        # 启动线程
        self.thread.start()

    def _stop_crawling(self):
        """停止爬虫"""
        if self.worker:
            self.worker.stop()
        self.stop_button.setEnabled(False)
        self.status_bar.showMessage("正在停止爬虫...")

    def _crawler_finished(self):
        """爬虫完成后的处理"""
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.save_results_button.setEnabled(True)
        self.status_bar.showMessage("爬虫任务完成")
        
        # 安全地获取数据
        try:
            if self.worker and hasattr(self.worker, 'current_crawled_data'):
                self.crawled_data = self.worker.current_crawled_data.copy()  # 创建副本
            else:
                self.crawled_data = []
        except Exception as e:
            self._log(f"获取爬虫数据时出错: {e}")
            self.crawled_data = []
            
        self._log("爬虫线程已结束。")
        
        # 延迟清理线程引用，避免内存访问错误
        if self.thread:
            self.thread.deleteLater()
            self.thread = None
        if self.worker:
            self.worker.deleteLater()
            self.worker = None

    def _save_results(self):
        if not self.crawled_data:
            QMessageBox.warning(self, "警告", "没有数据可保存！")
            return

        try:
            output_filename = self.output_prefix_input.text() + datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # 获取保存路径
            save_path = self.save_path_input.text()
            if save_path and save_path.strip():
                full_path = os.path.join(save_path, output_filename + '.xlsx')
                display_path = full_path
            else:
                full_path = output_filename + '.xlsx'
                display_path = output_filename + '.xlsx'
            
            head = ['序号', '关键字', '名称', '日期', '采购人', '代理机构', '公告类型', '详情', '项目概况']
            workbook = xlsxwriter.Workbook(full_path)
            worksheet = workbook.add_worksheet("中标公告")
            for cvi, cv in enumerate(head):
                worksheet.write(0, cvi, cv)
            for row_idx, rowdata in enumerate(self.crawled_data, start=1):
                for col_idx, cell_data in enumerate(rowdata):
                    worksheet.write(row_idx, col_idx, cell_data)
            workbook.close()
            
            self._log(f"数据已手动保存到 {display_path}")
            QMessageBox.information(self, "成功", f"数据已成功保存到 {display_path}")
        except Exception as e:
            self._log(f"手动保存数据时出错: {e}")
            QMessageBox.critical(self, "错误", f"手动保存数据时出错: {e}")

    def closeEvent(self, event):
        """处理窗口关闭事件，确保线程正常退出"""
        try:
            self.save_config()
            
            # 如果线程正在运行，先停止
            if self.thread and self.thread.isRunning():
                self._log("正在停止爬虫线程...")
                
                # 停止工作线程
                if self.worker:
                    self.worker.stop()
                
                # 等待线程结束
                self.thread.quit()
                if not self.thread.wait(3000):  # 等待3秒
                    self._log("强制终止爬虫线程...")
                    self.thread.terminate()
                    self.thread.wait(1000)  # 等待1秒确保终止
                
                # 清理引用
                self.worker = None
                self.thread = None
                
            # 接受关闭事件
            event.accept()
            
        except Exception as e:
            self._log(f"关闭程序时出错: {e}")
            # 即使出错也要关闭
            event.accept()


if __name__ == '__main__':
    app = None
    try:
        # 设置Qt平台插件路径
        if hasattr(sys, 'frozen'):
            # 如果是打包后的应用
            plugin_path = os.path.join(os.path.dirname(sys.executable), 'platforms')
            os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = plugin_path
        else:
            # 如果是从源码运行
            try:
                import PyQt6
                os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = os.path.join(os.path.dirname(PyQt6.__file__), 'Qt6', 'plugins', 'platforms')
            except ImportError:
                pass

        app = QApplication(sys.argv)
        ex = MainWindow()
        ex.show()
        exit_code = app.exec()
        
        # 确保在退出前清理资源
        ex.cleanup_resources()
        sys.exit(exit_code)
        
    except Exception as e:
        print(f"应用启动失败，错误信息: {e}")
        print("\n可能的解决方案:")
        print("1. 确保已安装PyQt6: pip install PyQt6")
        print("2. 安装Microsoft Visual C++ Redistributable")
        print("3. 如果错误信息包含'platform plugin'，请确保Qt平台插件正确安装")
        print("4. 重新启动计算机后再试")
        sys.exit(1)
    finally:
        if app:
            app.quit()
            app.deleteLater()