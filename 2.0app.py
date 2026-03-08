import sys
import os
import pandas as pd
import numpy as np
import sqlite3
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QTabWidget, QGroupBox, QFormLayout,
    QLineEdit, QComboBox, QTextEdit, QDoubleSpinBox, QFileDialog,
    QMessageBox, QStackedWidget, QListWidget, QSplitter, QSizePolicy
)
from PyQt5.QtCore import Qt, QSize, QDate
from PyQt5.QtGui import QFont, QPixmap, QIcon
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import matplotlib.pyplot as plt
from datetime import datetime

# ===================== 全局配置 =====================
# 数据库初始化
DB_FILE = "esc_quality.db"
def init_db():
    """初始化品质管理数据库"""
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    # 基础产品信息表
    c.execute('''CREATE TABLE IF NOT EXISTS product_info
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  sn TEXT UNIQUE NOT NULL,  # 序列号
                  model TEXT NOT NULL,     # 型号
                  test_date TEXT NOT NULL, # 检测日期
                  operator TEXT NOT NULL,  # 操作员
                  role TEXT DEFAULT 'operator') # 最后操作角色
              ''')
    # 外观检测表
    c.execute('''CREATE TABLE IF NOT EXISTS appearance_test
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  sn TEXT NOT NULL,
                  he_hole_spec_max REAL,    # HE孔规格上限
                  he_hole_spec_min REAL,    # HE孔规格下限
                  he_hole_measure REAL,     # HE孔测量值
                  abnormal_type TEXT,       # 异常类型（堵塞/侧漏/无）
                  abnormal_pos_3d_x REAL,   # 3D异常位置X
                  abnormal_pos_3d_y REAL,   # 3D异常位置Y
                  abnormal_pos_3d_z REAL,   # 3D异常位置Z
                  abnormal_desc TEXT,       # 异常描述
                  photo_path TEXT,          # 异常照片路径
                  result TEXT DEFAULT 'NG') # 结果（OK/NG）
              ''')
    # 功能检测表（介电层/粗糙度/电性能等）
    c.execute('''CREATE TABLE IF NOT EXISTS function_test
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  sn TEXT NOT NULL,
                  # 介电层厚度（外圈8点+内圈8点+中间1点，共17点）
                  dielectric_outer1 REAL, dielectric_outer2 REAL, dielectric_outer3 REAL,
                  dielectric_outer4 REAL, dielectric_outer5 REAL, dielectric_outer6 REAL,
                  dielectric_outer7 REAL, dielectric_outer8 REAL,
                  dielectric_inner1 REAL, dielectric_inner2 REAL, dielectric_inner3 REAL,
                  dielectric_inner4 REAL, dielectric_inner5 REAL, dielectric_inner6 REAL,
                  dielectric_inner7 REAL, dielectric_inner8 REAL,
                  dielectric_center REAL,
                  # 粗糙度（最多8点）
                  roughness1 REAL, roughness2 REAL, roughness3 REAL, roughness4 REAL,
                  roughness5 REAL, roughness6 REAL, roughness7 REAL, roughness8 REAL,
                  roughness_spec_max REAL, roughness_spec_min REAL,
                  # 电性能
                  voltage REAL,             # 测试电压
                  current_spec_max REAL,    # 电流阈值上限
                  current_measure REAL,     # 实测电流
                  insulation_resistance REAL, # 绝缘电阻
                  heater_resistance_spec REAL, # 加热器阻值规格
                  heater_resistance_measure REAL, # 加热器阻值实测
                  he_leak_spec_max REAL,    # HE漏阈值上限
                  he_leak_measure REAL,     # HE漏实测值
                  # 测温/超声波
                  temp_max REAL, temp_min REAL, temp_diff REAL,
                  thermal_map_path TEXT,    # 热力图路径
                  ultrasonic_surface TEXT,  # 超声波表面图
                  ultrasonic_electrode TEXT,# 超声波电极图
                  ultrasonic_banding TEXT,  # 超声波banding图
                  # 结果
                  dielectric_result TEXT,
                  roughness_result TEXT,
                  electric_result TEXT,
                  heater_result TEXT,
                  he_leak_result TEXT)
              ''')
    conn.commit()
    conn.close()

# ===================== 3D可视化组件 =====================
class Plotly3DWidget(QWidget):
    """Plotly 3D可视化组件（支持拖拽标注）"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.layout = QVBoxLayout(self)
        self.current_sn = ""
        self.abnormal_pos = None  # 存储3D标注的异常位置
    
    def show_3d_appearance(self, sn):
        """显示外观检测3D图（带异常标注）"""
        self.current_sn = sn
        # 查询该SN的异常位置
        conn = sqlite3.connect(DB_FILE)
        df = pd.read_sql(f"SELECT * FROM appearance_test WHERE sn='{sn}'", conn)
        conn.close()
        
        # 生成基础3D模型（ESC卡盘简化模型）
        theta = np.linspace(0, 2*np.pi, 100)
        phi = np.linspace(0, np.pi, 100)
        x = 10 * np.outer(np.cos(theta), np.sin(phi))
        y = 10 * np.outer(np.sin(theta), np.sin(phi))
        z = 2 * np.outer(np.ones(np.size(theta)), np.cos(phi))
        
        # 创建3D图
        fig = go.Figure(data=[
            go.Surface(x=x, y=y, z=z, colorscale='Blues', opacity=0.7, name='ESC本体'),
        ])
        
        # 添加异常标注（如果有）
        if not df.empty:
            self.abnormal_pos = (
                df.iloc[0]['abnormal_pos_3d_x'],
                df.iloc[0]['abnormal_pos_3d_y'],
                df.iloc[0]['abnormal_pos_3d_z']
            )
            fig.add_trace(go.Scatter3d(
                x=[self.abnormal_pos[0]], y=[self.abnormal_pos[1]], z=[self.abnormal_pos[2]],
                mode='markers+text',
                marker=dict(size=10, color='red', symbol='x'),
                text=[f"异常：{df.iloc[0]['abnormal_type']}"],
                textposition='top center',
                name='异常位置'
            ))
        
        # 配置布局
        fig.update_layout(
            scene=dict(xaxis_title='X轴', yaxis_title='Y轴', zaxis_title='Z轴'),
            title=f"ESC外观3D检测图（SN：{sn}）",
            width=800, height=600
        )
        
        # 显示在Qt中（Plotly生成HTML，可嵌入WebEngineView，简化版先保存为HTML）
        html_path = f"3d_appearance_{sn}.html"
        fig.write_html(html_path)
        # 注：完整版本可使用PyQt5.QtWebEngineWidgets.QWebEngineView加载HTML
        QMessageBox.information(self, "提示", f"3D图已生成：{html_path}\n可双击打开查看（支持拖拽旋转）")
    
    def show_3d_dielectric(self, sn):
        """显示介电层厚度3D高低图"""
        self.current_sn = sn
        # 查询介电层数据
        conn = sqlite3.connect(DB_FILE)
        df = pd.read_sql(f"SELECT * FROM function_test WHERE sn='{sn}'", conn)
        conn.close()
        
        if df.empty:
            QMessageBox.warning(self, "提示", "暂无介电层数据！")
            return
        
        # 提取17个点的厚度数据
        dielectric_points = [
            # 外圈8点
            df.iloc[0]['dielectric_outer1'], df.iloc[0]['dielectric_outer2'],
            df.iloc[0]['dielectric_outer3'], df.iloc[0]['dielectric_outer4'],
            df.iloc[0]['dielectric_outer5'], df.iloc[0]['dielectric_outer6'],
            df.iloc[0]['dielectric_outer7'], df.iloc[0]['dielectric_outer8'],
            # 内圈8点
            df.iloc[0]['dielectric_inner1'], df.iloc[0]['dielectric_inner2'],
            df.iloc[0]['dielectric_inner3'], df.iloc[0]['dielectric_inner4'],
            df.iloc[0]['dielectric_inner5'], df.iloc[0]['dielectric_inner6'],
            df.iloc[0]['dielectric_inner7'], df.iloc[0]['dielectric_inner8'],
            # 中间1点
            df.iloc[0]['dielectric_center']
        ]
        
        # 生成17个点的3D坐标
        outer_theta = np.linspace(0, 2*np.pi, 8)
        inner_theta = np.linspace(0, 2*np.pi, 8)
        x = list(8 * np.cos(outer_theta)) + list(4 * np.cos(inner_theta)) + [0]
        y = list(8 * np.sin(outer_theta)) + list(4 * np.sin(inner_theta)) + [0]
        z = dielectric_points
        
        # 创建3D散点图（颜色表示厚度高低）
        fig = go.Figure(data=[
            go.Scatter3d(
                x=x, y=y, z=z,
                mode='markers+text',
                marker=dict(size=8, color=z, colorscale='Viridis', colorbar_title='厚度(μm)'),
                text=[f"点{i+1}: {z[i]:.2f}" for i in range(len(z))],
                textposition='top center'
            )
        ])
        
        fig.update_layout(
            scene=dict(xaxis_title='X轴', yaxis_title='Y轴', zaxis_title='厚度(μm)'),
            title=f"介电层厚度3D分布（SN：{sn}）",
            width=800, height=600
        )
        
        # 保存并提示
        html_path = f"3d_dielectric_{sn}.html"
        fig.write_html(html_path)
        QMessageBox.information(self, "提示", f"介电层3D图已生成：{html_path}")

# ===================== 粗糙度波浪线组件 =====================
class RoughnessPlotWidget(FigureCanvas):
    """Matplotlib粗糙度波浪线可视化"""
    def __init__(self, parent=None):
        self.fig, self.ax = plt.subplots(figsize=(10, 3))
        super().__init__(self.fig)
        self.setParent(parent)
        self.spec_max = 0
        self.spec_min = 0
    
    def plot_roughness(self, roughness_data):
        """绘制粗糙度波浪线（阈值内绿色平滑，阈值外红色褶皱）"""
        self.ax.clear()
        x = np.arange(1, len(roughness_data)+1)
        y = np.array(roughness_data)
        
        # 区分合格/异常点
        normal_mask = (y >= self.spec_min) & (y <= self.spec_max)
        abnormal_mask = ~normal_mask
        
        # 绘制波浪线（合格：绿色平滑，异常：红色褶皱）
        if np.any(normal_mask):
            self.ax.plot(x[normal_mask], y[normal_mask], 'g-', linewidth=2, label='合格', alpha=0.8)
        if np.any(abnormal_mask):
            # 褶皱效果：添加随机波动
            abnormal_y = y[abnormal_mask] + 0.5 * np.random.randn(len(y[abnormal_mask]))
            self.ax.plot(x[abnormal_mask], abnormal_y, 'r-', linewidth=2, label='异常', alpha=0.8)
        
        # 绘制阈值线
        self.ax.axhline(y=self.spec_max, color='orange', linestyle='--', label='规格上限')
        self.ax.axhline(y=self.spec_min, color='orange', linestyle='--', label='规格下限')
        
        # 标注数值
        for i, (xi, yi) in enumerate(zip(x, y)):
            self.ax.text(xi, yi+0.1, f'{yi:.2f}', ha='center')
        
        self.ax.set_xlabel('检测点')
        self.ax.set_ylabel('粗糙度值')
        self.ax.set_title('ESC粗糙度检测波形图')
        self.ax.legend()
        self.ax.grid(True, alpha=0.3)
        self.draw()

# ===================== 主窗口 =====================
class ESCQualityManager(QMainWindow):
    def __init__(self):
        super().__init__()
        init_db()
        self.setWindowTitle('ESC品质数据管理系统')
        self.setGeometry(100, 100, 1400, 900)
        self.current_role = 'operator'  # 默认操作员角色
        self.current_sn = ""            # 当前操作的SN
        self.init_ui()
    
    def init_ui(self):
        """初始化界面"""
        # 主布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # 1. 顶部栏（角色切换+SN输入）
        top_layout = QHBoxLayout()
        # 角色切换
        self.role_btn = QPushButton(f'当前角色：{self.current_role}', self)
        self.role_btn.clicked.connect(self.switch_role)
        self.role_btn.setStyleSheet("background-color: #3498db; color: white; padding: 10px;")
        top_layout.addWidget(self.role_btn)
        
        # SN输入
        self.sn_edit = QLineEdit()
        self.sn_edit.setPlaceholderText("输入产品序列号（SN）")
        self.sn_edit.setFixedWidth(200)
        top_layout.addWidget(QLabel("产品SN："))
        top_layout.addWidget(self.sn_edit)
        
        # 加载按钮
        self.load_btn = QPushButton("加载数据")
        self.load_btn.clicked.connect(self.load_sn_data)
        self.load_btn.setStyleSheet("background-color: #2ecc71; color: white; padding: 10px;")
        top_layout.addWidget(self.load_btn)
        
        # 导出按钮（仅决策者可见）
        self.export_btn = QPushButton("一键导出数据")
        self.export_btn.clicked.connect(self.export_data)
        self.export_btn.setStyleSheet("background-color: #e67e22; color: white; padding: 10px;")
        self.export_btn.setVisible(False)  # 默认隐藏
        top_layout.addWidget(self.export_btn)
        
        top_layout.addStretch()
        main_layout.addLayout(top_layout)
        
        # 2. 核心内容区（堆叠窗口：操作员/决策者）
        self.stacked_widget = QStackedWidget()
        # 2.1 操作员界面
        self.operator_widget = self.create_operator_widget()
        # 2.2 决策者界面
        self.decision_widget = self.create_decision_widget()
        
        self.stacked_widget.addWidget(self.operator_widget)
        self.stacked_widget.addWidget(self.decision_widget)
        self.stacked_widget.setCurrentWidget(self.operator_widget)
        main_layout.addWidget(self.stacked_widget)
    
    def switch_role(self):
        """切换操作员/决策者角色"""
        if self.current_role == 'operator':
            self.current_role = 'decision'
            self.role_btn.setText('当前角色：decision')
            self.stacked_widget.setCurrentWidget(self.decision_widget)
            self.export_btn.setVisible(True)
        else:
            self.current_role = 'operator'
            self.role_btn.setText('当前角色：operator')
            self.stacked_widget.setCurrentWidget(self.operator_widget)
            self.export_btn.setVisible(False)
        QMessageBox.information(self, "提示", f"已切换至{self.current_role}角色")
    
    def load_sn_data(self):
        """加载指定SN的数据"""
        self.current_sn = self.sn_edit.text().strip()
        if not self.current_sn:
            QMessageBox.warning(self, "提示", "请输入SN！")
            return
        
        # 检查是否存在该SN，不存在则创建基础信息
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("SELECT * FROM product_info WHERE sn=?", (self.current_sn,))
        if not c.fetchone():
            # 新增产品信息
            model, ok = QInputDialog.getText(self, "新增产品", "输入产品型号：")
            if not ok or not model:
                QMessageBox.warning(self, "提示", "型号不能为空！")
                return
            operator, ok = QInputDialog.getText(self, "新增产品", "输入操作员姓名：")
            if not ok or not operator:
                QMessageBox.warning(self, "提示", "操作员不能为空！")
                return
            c.execute(
                "INSERT INTO product_info (sn, model, test_date, operator) VALUES (?,?,?,?)",
                (self.current_sn, model, datetime.now().strftime('%Y-%m-%d'), operator)
            )
        # 更新最后操作角色
        c.execute("UPDATE product_info SET role=? WHERE sn=?", (self.current_role, self.current_sn))
        conn.commit()
        conn.close()
        
        QMessageBox.information(self, "提示", f"已加载SN：{self.current_sn}的数据")
        # 刷新当前界面数据
        if self.current_role == 'operator':
            self.refresh_operator_data()
        else:
            self.refresh_decision_data()
    
    def create_operator_widget(self):
        """创建操作员界面"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # 标签页：外观检测/尺寸检测/功能检测
        tab_widget = QTabWidget()
        
        # 1. 外观检测标签页
        appearance_tab = self.create_appearance_tab()
        tab_widget.addTab(appearance_tab, "外观检测")
        
        # 2. 尺寸检测标签页（预留）
        size_tab = QWidget()
        size_layout = QVBoxLayout(size_tab)
        size_layout.addWidget(QLabel("尺寸检测模块暂未实现"))
        tab_widget.addTab(size_tab, "尺寸检测")
        
        # 3. 功能检测标签页
        function_tab = self.create_function_tab()
        tab_widget.addTab(function_tab, "功能检测")
        
        layout.addWidget(tab_widget)
        return widget
    
    def create_appearance_tab(self):
        """创建外观检测标签页"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # 1. HE孔检测
        he_group = QGroupBox("HE孔检测")
        he_layout = QFormLayout(he_group)
        # 规格输入
        self.he_spec_min = QDoubleSpinBox()
        self.he_spec_min.setRange(0, 100)
        self.he_spec_min.setDecimals(2)
        he_layout.addRow("规格下限：", self.he_spec_min)
        
        self.he_spec_max = QDoubleSpinBox()
        self.he_spec_max.setRange(0, 100)
        self.he_spec_max.setDecimals(2)
        he_layout.addRow("规格上限：", self.he_spec_max)
        
        # 测量值
        self.he_measure = QDoubleSpinBox()
        self.he_measure.setRange(0, 100)
        self.he_measure.setDecimals(2)
        he_layout.addRow("测量值：", self.he_measure)
        
        # 结果显示
        self.he_result = QLabel("结果：未检测")
        self.he_result.setStyleSheet("font-size: 14px; font-weight: bold;")
        he_layout.addRow("", self.he_result)
        
        # 生成折线图按钮
        self.he_plot_btn = QPushButton("生成检测折线图")
        self.he_plot_btn.clicked.connect(self.plot_he_hole)
        he_layout.addRow("", self.he_plot_btn)
        
        # 2. 外观异常标注
        abnormal_group = QGroupBox("外观异常标注")
        abnormal_layout = QFormLayout(abnormal_group)
        # 异常类型
        self.abnormal_type = QComboBox()
        self.abnormal_type.addItems(["无", "堵塞", "侧漏"])
        abnormal_layout.addRow("异常类型：", self.abnormal_type)
        
        # 3D异常位置
        self.abnormal_x = QDoubleSpinBox()
        self.abnormal_x.setRange(-20, 20)
        self.abnormal_x.setDecimals(1)
        abnormal_layout.addRow("3D位置X：", self.abnormal_x)
        
        self.abnormal_y = QDoubleSpinBox()
        self.abnormal_y.setRange(-20, 20)
        self.abnormal_y.setDecimals(1)
        abnormal_layout.addRow("3D位置Y：", self.abnormal_y)
        
        self.abnormal_z = QDoubleSpinBox()
        self.abnormal_z.setRange(0, 10)
        self.abnormal_z.setDecimals(1)
        abnormal_layout.addRow("3D位置Z：", self.abnormal_z)
        
        # 异常描述
        self.abnormal_desc = QTextEdit()
        self.abnormal_desc.setFixedHeight(60)
        abnormal_layout.addRow("异常描述：", self.abnormal_desc)
        
        # 异常照片上传
        self.photo_path = QLabel("未上传照片")
        self.upload_photo_btn = QPushButton("上传异常照片")
        self.upload_photo_btn.clicked.connect(self.upload_photo)
        abnormal_layout.addRow("", self.upload_photo_btn)
        abnormal_layout.addRow("", self.photo_path)
        
        # 3D预览按钮
        self.3d_preview_btn = QPushButton("3D模型预览（标注异常位置）")
        self.3d_preview_btn.clicked.connect(self.show_3d_appearance)
        abnormal_layout.addRow("", self.3d_preview_btn)
        
        # 保存按钮
        self.save_appearance_btn = QPushButton("保存外观检测数据")
        self.save_appearance_btn.setStyleSheet("background-color: #27ae60; color: white; padding: 10px;")
        self.save_appearance_btn.clicked.connect(self.save_appearance_data)
        
        layout.addWidget(he_group)
        layout.addWidget(abnormal_group)
        layout.addWidget(self.save_appearance_btn, alignment=Qt.AlignCenter)
        layout.addStretch()
        return widget
    
    def create_function_tab(self):
        """创建功能检测标签页"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # 1. 介电层厚度
        dielectric_group = QGroupBox("介电层厚度检测（外圈8点+内圈8点+中间1点）")
        dielectric_layout = QFormLayout(dielectric_group)
        # 外圈8点
        self.dielectric_outer = [QDoubleSpinBox() for _ in range(8)]
        for i, spin in enumerate(self.dielectric_outer):
            spin.setRange(0, 100)
            spin.setDecimals(2)
            dielectric_layout.addRow(f"外圈{i+1}点：", spin)
        # 内圈8点
        self.dielectric_inner = [QDoubleSpinBox() for _ in range(8)]
        for i, spin in enumerate(self.dielectric_inner):
            spin.setRange(0, 100)
            spin.setDecimals(2)
            dielectric_layout.addRow(f"内圈{i+1}点：", spin)
        # 中间点
        self.dielectric_center = QDoubleSpinBox()
        self.dielectric_center.setRange(0, 100)
        self.dielectric_center.setDecimals(2)
        dielectric_layout.addRow("中间点：", self.dielectric_center)
        # 3D预览
        self.dielectric_3d_btn = QPushButton("生成3D厚度分布")
        self.dielectric_3d_btn.clicked.connect(self.show_3d_dielectric)
        dielectric_layout.addRow("", self.dielectric_3d_btn)
        
        # 2. 粗糙度检测
        roughness_group = QGroupBox("粗糙度检测")
        roughness_layout = QFormLayout(roughness_group)
        # 规格
        self.rough_spec_min = QDoubleSpinBox()
        self.rough_spec_min.setRange(0, 100)
        self.rough_spec_min.setDecimals(2)
        roughness_layout.addRow("规格下限：", self.rough_spec_min)
        
        self.rough_spec_max = QDoubleSpinBox()
        self.rough_spec_max.setRange(0, 100)
        self.rough_spec_max.setDecimals(2)
        roughness_layout.addRow("规格上限：", self.rough_spec_max)
        
        # 测量点（最多8点）
        self.roughness_points = [QDoubleSpinBox() for _ in range(8)]
        for i, spin in enumerate(self.roughness_points):
            spin.setRange(0, 100)
            spin.setDecimals(2)
            roughness_layout.addRow(f"检测点{i+1}：", spin)
        
        # 波形图预览
        self.rough_plot_btn = QPushButton("生成粗糙度波形图")
        self.rough_plot_btn.clicked.connect(self.plot_roughness)
        roughness_layout.addRow("", self.rough_plot_btn)
        
        # 波形图显示区域
        self.rough_plot_widget = RoughnessPlotWidget()
        roughness_layout.addRow("", self.rough_plot_widget)
        
        # 3. 电性能/HE漏等
        electric_group = QGroupBox("电性能/HE漏/加热器阻值检测")
        electric_layout = QFormLayout(electric_group)
        # 电压&电流
        self.voltage = QDoubleSpinBox()
        self.voltage.setRange(0, 1000)
        electric_layout.addRow("测试电压(V)：", self.voltage)
        
        self.current_spec_max = QDoubleSpinBox()
        self.current_spec_max.setRange(0, 100)
        electric_layout.addRow("电流阈值上限(mA)：", self.current_spec_max)
        
        self.current_measure = QDoubleSpinBox()
        self.current_measure.setRange(0, 100)
        electric_layout.addRow("实测电流(mA)：", self.current_measure)
        
        # HE漏
        self.he_leak_spec_max = QDoubleSpinBox()
        self.he_leak_spec_max.setRange(0, 10)
        self.he_leak_spec_max.setDecimals(2)
        electric_layout.addRow("HE漏阈值上限(uA)：", self.he_leak_spec_max)
        
        self.he_leak_measure = QDoubleSpinBox()
        self.he_leak_measure.setRange(0, 10)
        self.he_leak_measure.setDecimals(2)
        electric_layout.addRow("HE漏实测值(uA)：", self.he_leak_measure)
        
        # 保存按钮
        self.save_function_btn = QPushButton("保存功能检测数据")
        self.save_function_btn.setStyleSheet("background-color: #27ae60; color: white; padding: 10px;")
        self.save_function_btn.clicked.connect(self.save_function_data)
        
        layout.addWidget(dielectric_group)
        layout.addWidget(roughness_group)
        layout.addWidget(electric_group)
        layout.addWidget(self.save_function_btn, alignment=Qt.AlignCenter)
        layout.addStretch()
        return widget
    
    def create_decision_widget(self):
        """创建决策者界面"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # 数据概览
        overview_group = QGroupBox("数据概览")
        overview_layout = QVBoxLayout(overview_group)
        self.overview_label = QLabel("请加载SN查看数据概览")
        overview_layout.addWidget(self.overview_label)
        
        # 可视化标签页
        plot_tab = QTabWidget()
        # 外观3D
        appearance_3d_tab = QWidget()
        appearance_3d_layout = QVBoxLayout(appearance_3d_tab)
        self.appearance_3d_btn = QPushButton("查看外观3D异常标注")
        self.appearance_3d_btn.clicked.connect(self.show_3d_appearance)
        appearance_3d_layout.addWidget(self.appearance_3d_btn)
        
        # 介电层3D
        dielectric_3d_tab = QWidget()
        dielectric_3d_layout = QVBoxLayout(dielectric_3d_tab)
        self.dielectric_3d_decision_btn = QPushButton("查看介电层3D厚度分布")
        self.dielectric_3d_decision_btn.clicked.connect(self.show_3d_dielectric)
        dielectric_3d_layout.addWidget(self.dielectric_3d_decision_btn)
        
        # 粗糙度波形
        rough_plot_tab = QWidget()
        rough_plot_layout = QVBoxLayout(rough_plot_tab)
        self.rough_plot_decision_btn = QPushButton("查看粗糙度波形图")
        self.rough_plot_decision_btn.clicked.connect(self.plot_roughness)
        rough_plot_layout.addWidget(self.rough_plot_decision_btn)
        self.rough_plot_decision_widget = RoughnessPlotWidget()
        rough_plot_layout.addWidget(self.rough_plot_decision_widget)
        
        plot_tab.addTab(appearance_3d_tab, "外观3D")
        plot_tab.addTab(dielectric_3d_tab, "介电层3D")
        plot_tab.addTab(rough_plot_tab, "粗糙度波形")
        
        layout.addWidget(overview_group)
        layout.addWidget(plot_tab)
        return widget
    
    # ===================== 核心功能实现 =====================
    def upload_photo(self):
        """上传异常照片"""
        file_path, _ = QFileDialog.getOpenFileName(self, "选择异常照片", "", "图片文件 (*.jpg *.png *.jpeg)")
        if file_path:
            self.photo_path.setText(f"照片路径：{file_path}")
    
    def plot_he_hole(self):
        """绘制HE孔检测折线图（阈值红绿标注）"""
        spec_min = self.he_spec_min.value()
        spec_max = self.he_spec_max.value()
        measure = self.he_measure.value()
        
        # 创建图表
        fig, ax = plt.subplots(figsize=(8, 4))
        # 绘制阈值区间
        ax.axhspan(spec_min, spec_max, color='green', alpha=0.2, label='合格区间')
        # 绘制测量点
        color = 'green' if spec_min <= measure <= spec_max else 'red'
        ax.plot([1], [measure], 'o', color=color, markersize=10, label=f'测量值：{measure}')
        # 标注结果
        result = "OK" if color == 'green' else "NG"
        ax.text(1.1, measure, result, fontsize=12, fontweight='bold', color=color)
        
        ax.set_xlim(0.5, 1.5)
        ax.set_ylim(spec_min-1, spec_max+1)
        ax.set_xlabel("HE孔检测")
        ax.set_ylabel("测量值")
        ax.set_title("HE孔检测结果")
        ax.legend()
        ax.grid(True)
        plt.show()
        
        # 更新结果标签
        self.he_result.setText(f"结果：{result}")
        self.he_result.setStyleSheet(f"color: {color}; font-size: 14px; font-weight: bold;")
    
    def plot_roughness(self):
        """绘制粗糙度波浪线"""
        # 获取规格
        self.rough_plot_widget.spec_min = self.rough_spec_min.value()
        self.rough_plot_widget.spec_max = self.rough_spec_max.value()
        # 获取测量数据
        roughness_data = [spin.value() for spin in self.roughness_points if spin.value() > 0]
        if not roughness_data:
            QMessageBox.warning(self, "提示", "请输入至少一个粗糙度测量值！")
            return
        # 绘制
        self.rough_plot_widget.plot_roughness(roughness_data)
    
    def show_3d_appearance(self):
        """显示外观3D图"""
        if not self.current_sn:
            QMessageBox.warning(self, "提示", "请先加载SN！")
            return
        self.3d_widget = Plotly3DWidget()
        self.3d_widget.show_3d_appearance(self.current_sn)
    
    def show_3d_dielectric(self):
        """显示介电层3D图"""
        if not self.current_sn:
            QMessageBox.warning(self, "提示", "请先加载SN！")
            return
        self.3d_widget = Plotly3DWidget()
        self.3d_widget.show_3d_dielectric(self.current_sn)
    
    def save_appearance_data(self):
        """保存外观检测数据"""
        if not self.current_sn:
            QMessageBox.warning(self, "提示", "请先加载SN！")
            return
        
        # 判定结果
        he_measure = self.he_measure.value()
        he_spec_min = self.he_spec_min.value()
        he_spec_max = self.he_spec_max.value()
        abnormal_type = self.abnormal_type.currentText()
        result = "OK" if (he_spec_min <= he_measure <= he_spec_max) and (abnormal_type == "无") else "NG"
        
        # 保存到数据库
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        # 先删除旧数据
        c.execute("DELETE FROM appearance_test WHERE sn=?", (self.current_sn,))
        # 插入新数据
        c.execute('''INSERT INTO appearance_test
                     (sn, he_hole_spec_max, he_hole_spec_min, he_hole_measure,
                      abnormal_type, abnormal_pos_3d_x, abnormal_pos_3d_y, abnormal_pos_3d_z,
                      abnormal_desc, photo_path, result)
                     VALUES (?,?,?,?,?,?,?,?,?,?,?)''',
                  (self.current_sn, he_spec_max, he_spec_min, he_measure,
                   abnormal_type, self.abnormal_x.value(), self.abnormal_y.value(), self.abnormal_z.value(),
                   self.abnormal_desc.toPlainText(), self.photo_path.text().replace("照片路径：", ""), result))
        conn.commit()
        conn.close()
        
        QMessageBox.information(self, "提示", "外观检测数据保存成功！")
    
    def save_function_data(self):
        """保存功能检测数据"""
        if not self.current_sn:
            QMessageBox.warning(self, "提示", "请先加载SN！")
            return
        
        # 提取介电层数据
        outer_data = [spin.value() for spin in self.dielectric_outer]
        inner_data = [spin.value() for spin in self.dielectric_inner]
        center_data = self.dielectric_center.value()
        
        # 提取粗糙度数据
        rough_data = [spin.value() for spin in self.roughness_points]
        rough_spec_min = self.rough_spec_min.value()
        rough_spec_max = self.rough_spec_max.value()
        # 判定粗糙度结果
        rough_result = "OK" if all(rough_spec_min <= val <= rough_spec_max for val in rough_data if val > 0) else "NG"
        
        # 提取电性能数据
        voltage = self.voltage.value()
        current_spec_max = self.current_spec_max.value()
        current_measure = self.current_measure.value()
        electric_result = "OK" if current_measure <= current_spec_max else "NG"
        
        # HE漏结果
        he_leak_spec_max = self.he_leak_spec_max.value()
        he_leak_measure = self.he_leak_measure.value()
        he_leak_result = "OK" if he_leak_measure <= he_leak_spec_max else "NG"
        
        # 保存到数据库
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        # 先删除旧数据
        c.execute("DELETE FROM function_test WHERE sn=?", (self.current_sn,))
        # 插入新数据
        c.execute('''INSERT INTO function_test
                     (sn, dielectric_outer1, dielectric_outer2, dielectric_outer3, dielectric_outer4,
                      dielectric_outer5, dielectric_outer6, dielectric_outer7, dielectric_outer8,
                      dielectric_inner1, dielectric_inner2, dielectric_inner3, dielectric_inner4,
                      dielectric_inner5, dielectric_inner6, dielectric_inner7, dielectric_inner8,
                      dielectric_center, roughness1, roughness2, roughness3, roughness4,
                      roughness5, roughness6, roughness7, roughness8,
                      roughness_spec_max, roughness_spec_min, voltage, current_spec_max,
                      current_measure, he_leak_spec_max, he_leak_measure,
                      dielectric_result, roughness_result, electric_result, he_leak_result)
                     VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                  (self.current_sn, *outer_data, *inner_data, center_data, *rough_data,
                   rough_spec_max, rough_spec_min, voltage, current_spec_max,
                   current_measure, he_leak_spec_max, he_leak_measure,
                   "OK", rough_result, electric_result, he_leak_result))
        conn.commit()
        conn.close()
        
        QMessageBox.information(self, "提示", "功能检测数据保存成功！")
    
    def refresh_decision_data(self):
        """刷新决策者界面数据"""
        # 查询数据概览
        conn = sqlite3.connect(DB_FILE)
        # 产品基础信息
        product_df = pd.read_sql(f"SELECT * FROM product_info WHERE sn='{self.current_sn}'", conn)
        # 外观检测
        appearance_df = pd.read_sql(f"SELECT * FROM appearance_test WHERE sn='{self.current_sn}'", conn)
        # 功能检测
        function_df = pd.read_sql(f"SELECT * FROM function_test WHERE sn='{self.current_sn}'", conn)
        conn.close()
        
        # 生成概览文本
        overview_text = f"""
        <h3>产品基础信息</h3>
        <p>SN：{self.current_sn}</p>
        <p>型号：{product_df.iloc[0]['model'] if not product_df.empty else '未知'}</p>
        <p>检测日期：{product_df.iloc[0]['test_date'] if not product_df.empty else '未知'}</p>
        <p>操作员：{product_df.iloc[0]['operator'] if not product_df.empty else '未知'}</p>
        
        <h3>外观检测结果</h3>
        <p>HE孔结果：{appearance_df.iloc[0]['result'] if not appearance_df.empty else '未检测'}</p>
        <p>异常类型：{appearance_df.iloc[0]['abnormal_type'] if not appearance_df.empty else '无'}</p>
        
        <h3>功能检测结果</h3>
        <p>粗糙度结果：{function_df.iloc[0]['roughness_result'] if not function_df.empty else '未检测'}</p>
        <p>电性能结果：{function_df.iloc[0]['electric_result'] if not function_df.empty else '未检测'}</p>
        <p>HE漏结果：{function_df.iloc[0]['he_leak_result'] if not function_df.empty else '未检测'}</p>
        """
        self.overview_label.setText(overview_text)
    
    def export_data(self):
        """一键导出数据"""
        if not self.current_sn:
            QMessageBox.warning(self, "提示", "请先加载SN！")
            return
        
        # 选择导出路径
        export_path, _ = QFileDialog.getSaveFileName(self, "导出数据", f"ESC_{self.current_sn}.xlsx", "Excel文件 (*.xlsx)")
        if not export_path:
            return
        
        # 查询所有数据
        conn = sqlite3.connect(DB_FILE)
        product_df = pd.read_sql(f"SELECT * FROM product_info WHERE sn='{self.current_sn}'", conn)
        appearance_df = pd.read_sql(f"SELECT * FROM appearance_test WHERE sn='{self.current_sn}'", conn)
        function_df = pd.read_sql(f"SELECT * FROM function_test WHERE sn='{self.current_sn}'", conn)
        conn.close()
        
        # 导出到Excel（多sheet）
        with pd.ExcelWriter(export_path, engine='openpyxl') as writer:
            product_df.to_excel(writer, sheet_name='产品基础信息', index=False)
            appearance_df.to_excel(writer, sheet_name='外观检测', index=False)
            function_df.to_excel(writer, sheet_name='功能检测', index=False)
        
        QMessageBox.information(self, "提示", f"数据已成功导出到：{export_path}")

# ===================== 程序入口 =====================
if __name__ == '__main__':
    # 修复QInputDialog导入
    from PyQt5.QtWidgets import QInputDialog
    app = QApplication(sys.argv)
    window = ESCQualityManager()
    window.show()
    sys.exit(app.exec_())