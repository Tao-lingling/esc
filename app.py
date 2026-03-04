import sys
import os
import pandas as pd
import sqlite3
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QPushButton, QLabel, QFileDialog, QMessageBox, 
                             QTableWidget, QTableWidgetItem, QHBoxLayout, QTabWidget,
                             QLineEdit, QComboBox, QDateEdit, QGroupBox)
from PyQt5.QtGui import QFont, QColor, QBrush
from PyQt5.QtCore import Qt, QDate
from datetime import datetime

# ================= 全局配置 =================
# 1. 规格书配置
SPEC = {
    '漏电流(uA)': {'max': 5.0, 'unit': 'uA'},
    '接触电阻(Ohm)': {'min': 80, 'max': 120, 'unit': 'Ohm'},
    'HE漏(uA)': {'max': 2.0, 'unit': 'uA'},
    '表面温度(°C)': {'min': 15, 'max': 40, 'unit': '°C'}
}

# 2. 数据库初始化
DB_FILE = "esc_data.db"

def init_db():
    """初始化本地数据库"""
    if not os.path.exists(DB_FILE):
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute('''CREATE TABLE IF NOT EXISTS test_records
                     (id INTEGER PRIMARY KEY AUTOINCREMENT,
                      sn TEXT NOT NULL,
                      product_model TEXT NOT NULL,
                      batch TEXT,
                      current REAL,
                      resistance REAL,
                      he_leak REAL,
                      temp REAL,
                      tester TEXT,
                      test_time TEXT NOT NULL,
                      overall_status TEXT,
                      remark TEXT)''')
        conn.commit()
        conn.close()

# ================= 核心业务逻辑 =================
def parse_excel(file_path):
    """解析公司内网的原始Excel"""
    try:
        df = pd.read_excel(file_path)
        
        # 数据清洗：去除空行
        df = df.dropna(subset=['SN序列号'], how='all') 
        return df
    except Exception as e:
        raise Exception(f"解析Excel失败:{str(e)}")

def analyze_and_save(df, tester="产线操作员"):
    """
    分析数据并保存到数据库
    【关键升级】:从SN号自动识别产品型号
    """
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    
    for index, row in df.iterrows():
        try:
            # 1. 提取基础数据
            sn = str(row['SN序列号']).strip()
            
            # 🔧 核心升级：自动识别产品型号
            # 规则：如果SN包含887或SYM3，则自动归类；否则标记为其他
            if '887' in sn.upper():
                product_model = 'ESC-887'
            elif 'SYM3' in sn.upper():
                product_model = 'SYM3'
            else:
                product_model = '其他'

            batch = str(row.get('批次号', '未知批次')).strip()
            current = float(row['漏电流(uA)'])
            resistance = float(row['接触电阻(Ohm)'])
            he_leak = float(row['HE漏(uA)'])
            temp = float(row['表面温度(°C)'])

            # 2. 自动判定
            def check_param(value, spec):
                if 'max' in spec:
                    return value <= spec['max']
                if 'min' in spec and 'max' in spec:
                    return spec['min'] <= value <= spec['max']
                return True

            ok_current = check_param(current, SPEC['漏电流(uA)'])
            ok_resistance = check_param(resistance, SPEC['接触电阻(Ohm)'])
            ok_he = check_param(he_leak, SPEC['HE漏(uA)'])
            ok_temp = check_param(temp, SPEC['表面温度(°C)'])

            overall_ok = ok_current and ok_resistance and ok_he and ok_temp
            overall_status = "✅ 全项通过" if overall_ok else "⚠️ 存在异常"
            
            # 3. 生成备注
            remarks = []
            if not ok_current: remarks.append("漏电流超")
            if not ok_resistance: remarks.append("电阻偏")
            if not ok_he: remarks.append("HE漏高")
            if not ok_temp: remarks.append("温度异")
            remark_str = " | ".join(remarks)

            # 4. 插入数据库
            c.execute('''INSERT INTO test_records 
                      (sn, product_model, batch, current, resistance, he_leak, temp, tester, test_time, overall_status, remark) 
                      VALUES (?,?,?,?,?,?,?,?,?,?,?)''',
                      (sn, product_model, batch, current, resistance, he_leak, temp, tester,
                       datetime.now().strftime('%Y-%m-%d %H:%M:%S'), # 记录上传时间
                       overall_status, remark_str))

        except KeyError as e:
            raise Exception(f"第{index+2}行缺少关键字段：{e}")

    conn.commit()
    conn.close()
    return len(df)

# ================= 数据查询模块 =================
def query_data_by_filters(sn_filter="", model_filter="", start_date="", end_date=""):
    """
    按多条件筛选数据
    :param sn_filter: SN号搜索
    :param model_filter: 产品型号 (ESC-887/SYM3/其他)
    :param start_date: 开始日期 YYYY-MM-DD
    :param end_date: 结束日期 YYYY-MM-DD
    :return: DataFrame
    """
    conn = sqlite3.connect(DB_FILE)
    
    # 构建基础查询
    query = "SELECT * FROM test_records WHERE 1=1"
    params = []

    # 1. SN 筛选
    if sn_filter:
        query += " AND sn LIKE ?"
        params.append(f'%{sn_filter}%')

    # 2. 产品型号筛选
    if model_filter:
        query += " AND product_model = ?"
        params.append(model_filter)

    # 3. 日期范围筛选
    if start_date:
        query += " AND date(test_time) >= ?"
        params.append(start_date)
    if end_date:
        query += " AND date(test_time) <= ?"
        params.append(end_date)

    # 按时间倒序
    query += " ORDER BY test_time DESC"

    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df

# ================= UI 界面代码 =================
class ESCManagerWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        init_db()
        self.setWindowTitle('ESC 静电卡盘数据管理系统 v4.0 (智能查询版)')
        self.setGeometry(100, 100, 1300, 800)
        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # 1. 标题
        title_font = QFont("Microsoft YaHei", 20, QFont.Weight.Bold)
        title_label = QLabel('🚗 ESC 静电卡盘全链路数据管理中心')
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setFont(title_font)
        title_label.setStyleSheet("color: #2C3E50; margin: 20px 0;")
        main_layout.addWidget(title_label)

        # 2. 功能区：上传 + 筛选
        func_layout = QHBoxLayout()
        
        # 左侧：上传
        self.btn_upload = QPushButton('📂 上传内网原始报表 (Excel)')
        self.btn_upload.setStyleSheet("background-color: #3498db; color: white; padding: 15px; font-size: 14px;")
        self.btn_upload.clicked.connect(self.upload_file)
        func_layout.addWidget(self.btn_upload)

        # 右侧：筛选组 (新增功能)
        filter_group = QGroupBox("🔎 历史数据精准筛选")
        filter_layout = QHBoxLayout(filter_group)
        
        self.cmb_model = QComboBox()
        self.cmb_model.addItems(["全部", "ESC-887", "SYM3", "其他"])
        self.cmb_model.setPlaceholderText("选择产品型号")
        
        self.txt_sn = QLineEdit()
        self.txt_sn.setPlaceholderText("输入SN号查询 (如 ESC-887-001)")
        
        self.dt_start = QDateEdit(QDate.currentDate().addDays(-7))
        self.dt_start.setDisplayFormat("yyyy-MM-dd")
        self.dt_end = QDateEdit(QDate.currentDate())
        self.dt_end.setDisplayFormat("yyyy-MM-dd")
        
        self.btn_filter = QPushButton("🔄 筛选数据")
        self.btn_filter.clicked.connect(self.filter_data)
        
        filter_layout.addWidget(QLabel("型号:"))
        filter_layout.addWidget(self.cmb_model)
        filter_layout.addWidget(QLabel("SN:"))
        filter_layout.addWidget(self.txt_sn)
        filter_layout.addWidget(QLabel("日期从:"))
        filter_layout.addWidget(self.dt_start)
        filter_layout.addWidget(QLabel("到:"))
        filter_layout.addWidget(self.dt_end)
        filter_layout.addWidget(self.btn_filter)
        
        func_layout.addWidget(filter_group)
        main_layout.addLayout(func_layout)

        # 3. 标签页
        self.tabs = QTabWidget()
        self.tab_data = QWidget() 
        self.tab_stat = QWidget() 
        self.tabs.addTab(self.tab_data, "📊 详细数据档案")
        self.tabs.addTab(self.tab_stat, "📈 质量统计看板")
        main_layout.addWidget(self.tabs)

        # 初始化表格
        self.init_data_tab()
        self.init_stat_tab()

    def init_data_tab(self):
        layout = QVBoxLayout(self.tab_data)
        self.table = QTableWidget()
        self.table.setStyleSheet("font-size: 10pt;")
        layout.addWidget(self.table)
        self.load_data_to_table() # 初始加载全部数据

    def init_stat_tab(self):
        layout = QVBoxLayout(self.tab_stat)
        self.stat_label = QLabel("统计数据将显示在这里...")
        self.stat_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.stat_label)

    def upload_file(self):
        """选择文件并上传解析"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择公司内网的原始Excel报表", 
            "", "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            reply = QMessageBox.question(
                self, "确认上传", 
                f"确认要解析文件：\n{os.path.basename(file_path)}\n\n数据将自动区分887/SYM3并入库。",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                try:
                    QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
                    df = parse_excel(file_path)
                    count = analyze_and_save(df)
                    QMessageBox.information(self, "成功", f"成功解析 {count} 条数据！\n已按SN号自动分类入库。")
                    self.load_data_to_table() # 刷新
                except Exception as e:
                    QMessageBox.critical(self, "失败", f"处理失败：{str(e)}")
                finally:
                    QApplication.restoreOverrideCursor()

    def load_data_to_table(self, df=None):
        """加载数据到表格,若不传df则加载全部"""
        if df is None:
            df = query_data_by_filters()
        
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(df.columns))
        self.table.setHorizontalHeaderLabels(df.columns)

        for row in range(len(df)):
            for col in range(len(df.columns)):
                val = df.iloc[row, col]
                item = QTableWidgetItem(str(val))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                
                # 综合判定列上色
                if df.columns[col] == 'overall_status':
                    if '全项通过' in str(val):
                        item.setBackground(QBrush(QColor(46, 204, 113, 100)))
                    else:
                        item.setBackground(QBrush(QColor(231, 76, 60, 100)))
                        item.setForeground(QBrush(Qt.GlobalColor.white))
                # 型号列也可以上色区分
                elif df.columns[col] == 'product_model':
                    if '887' in str(val):
                        item.setBackground(QBrush(QColor(52, 152, 219, 50)))
                    elif 'SYM3' in str(val):
                        item.setBackground(QBrush(QColor(155, 89, 182, 50)))
                        
                self.table.setItem(row, col, item)
        
        self.table.resizeColumnsToContents()
        self.update_statistics(df)

    def filter_data(self):
        """
        执行筛选逻辑
        """
        # 1. 获取筛选条件
        sn = self.txt_sn.text().strip()
        model = self.cmb_model.currentText()
        if model == "全部": model = "" # 全部则不传参
        
        start_date = self.dt_start.date().toString("yyyy-MM-dd")
        end_date = self.dt_end.date().toString("yyyy-MM-dd")

        # 2. 查询
        df_filtered = query_data_by_filters(sn_filter=sn, model_filter=model, start_date=start_date, end_date=end_date)
        
        # 3. 展示
        if df_filtered.empty:
            QMessageBox.information(self, "提示", "未找到符合条件的数据记录！")
        else:
            QMessageBox.information(self, "筛选完成", f"共找到 {len(df_filtered)} 条记录。")
        
        self.load_data_to_table(df_filtered)

    def update_statistics(self, df):
        """更新统计看板"""
        if df.empty:
            self.stat_label.setText("暂无数据记录。")
            return
        
        total = len(df)
        passed = len(df[df['overall_status'] == '✅ 全项通过'])
        pass_rate = (passed / total) * 100 if total > 0 else 0
        abnormal = len(df[df['overall_status'].str.contains('异常', na=False)])

        stats_text = f"""
        <div style='font-size: 14px; line-height: 1.6; padding: 20px;'>
            <h2 style='color: #2C3E50; border-bottom: 1px solid #eee; padding-bottom: 10px;'>实时质量统计</h2>
            <p>📊 <b>当前筛选总数量：</b> {total}</p >
            <p>✅ <b>合格数量：</b> {passed} &nbsp;&nbsp; <span style='color: #27ae60; font-weight: bold;'>合格率：{pass_rate:.2f}%</span></p >
            <p>⚠️ <b>异常数量：</b> {abnormal} &nbsp;&nbsp; <span style='color: #e74c3c; font-weight: bold;'>异常率：{(abnormal/total)*100:.2f}%</span></p >
            <p>🌡️ <b>平均温度：</b> {df['temp'].mean():.2f}°C</p >
            <p>💎 <b>887型号占比:</b> {len(df[df['product_model']=='ESC-887'])/total*100:.1f}%</p >
        </div>
        """
        self.stat_label.setText(stats_text)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setFont(QFont("Microsoft YaHei"))
    window = ESCManagerWindow()
    window.show()
    sys.exit(app.exec())