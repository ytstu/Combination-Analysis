import sys
import os
import pandas as pd
from datetime import datetime
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import re

class SplashScreen(QSplashScreen):
    """启动加载画面"""
    def __init__(self):
        # 创建启动画面
        splash_pixmap = QPixmap(500, 300)
        splash_pixmap.fill(QColor("#2c3e50"))
        super().__init__(splash_pixmap)
        
        # 创建标签
        label = QLabel(self)
        label.setGeometry(0, 0, 500, 300)
        label.setAlignment(Qt.AlignCenter)
        label.setText("""
            <div style='text-align: center; color: white;'>
                <h2 style='font-size: 24px; margin-bottom: 20px;'>组合装数据分析工具</h2>
                <p style='font-size: 14px; margin-bottom: 30px;'>正在加载中，请稍候...</p>
                <p style='font-size: 12px; color: #bdc3c7;'>首次加载可能需要几秒钟</p>
            </div>
        """)
        label.setStyleSheet("background-color: transparent;")
        
        # 进度条
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setGeometry(150, 200, 200, 15)
        self.progress_bar.setRange(0, 0)  # 无限进度
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 2px solid #3498db;
                border-radius: 5px;
                background-color: white;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #3498db;
                border-radius: 3px;
            }
        """)

class LoadDatabasesThread(QThread):
    """后台加载数据库的线程"""
    finished = pyqtSignal(object, object)
    error = pyqtSignal(str)
    
    def __init__(self, product_path, combo_path):
        super().__init__()
        self.product_path = product_path
        self.combo_path = combo_path
        
    def run(self):
        try:
            product_df = None
            combo_df = None
            
            # 加载商品资料
            if os.path.exists(self.product_path):
                product_df = pd.read_excel(self.product_path)
                
            # 加载组合资料
            if os.path.exists(self.combo_path):
                combo_df = pd.read_excel(self.combo_path)
                
            self.finished.emit(product_df, combo_df)
        except Exception as e:
            self.error.emit(str(e))

class ExcelDataProcessor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.product_db_path = r'\\Desktop-inv4qoc\图片数据\Temu_半托项目组\倒表格\数据\组合分析\商品资料.xlsx'
        self.combo_db_path = r'\\Desktop-inv4qoc\图片数据\Temu_半托项目组\倒表格\数据\组合分析\组合资料.xlsx'
        self.product_df = None
        self.combo_df = None
        self.input_df = None
        self.processed_df = None
        
        # 显示启动画面
        self.splash = SplashScreen()
        self.splash.show()
        
        self.initUI()
        
        # 异步加载数据库
        self.load_databases_async()
        
    def initUI(self):
        self.setWindowTitle('组合装分析工具')
        self.setGeometry(100, 100, 360, 650)
        
        # 中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setSpacing(10)
        
        # 标题
        title_label = QLabel('使用方法：\n1.将 聚水潭商品编码缺失订单导出，\n2.拖拽或选择文件导入进行分析，\n3.点击【开始分析】处理数据，\n4.导出结果，\n5.将导出结果上传到组合装资料中')
        title_label.setStyleSheet("""
            QLabel {
                font-size: 14px;
                font-weight: bold;
                color: #2c3e50;
                padding: 2px;
            }
        """)
        #title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)
        
        # 数据库状态行（精简）
        db_layout = QHBoxLayout()
        self.product_db_label = QLabel("商品库: 加载中...")
        self.combo_db_label = QLabel("组合库: 加载中...")
        
        reload_btn = QPushButton("刷新")
        reload_btn.setFixedSize(60, 25)
        reload_btn.clicked.connect(self.load_databases_async)
        reload_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border-radius: 3px;
                font-size: 11px;
            }
            QPushButton:hover { background-color: #2980b9; }
        """)
        
        db_layout.addWidget(self.product_db_label)
        db_layout.addWidget(self.combo_db_label)
        db_layout.addStretch()
        db_layout.addWidget(reload_btn)
        layout.addLayout(db_layout)
        
        # 步骤指示器
        steps_widget = QWidget()
        steps_layout = QHBoxLayout(steps_widget)
        steps_layout.setContentsMargins(0, 5, 0, 10)
        
        self.step1_label = QLabel("① 导入数据")
        self.step1_label.setStyleSheet("color: #3498db; font-weight: bold;")
        
        self.step2_label = QLabel("② 开始分析")
        
        self.step3_label = QLabel("③ 导出结果")
        
        steps_layout.addWidget(self.step1_label)
        steps_layout.addWidget(QLabel(" → "))
        steps_layout.addWidget(self.step2_label)
        steps_layout.addWidget(QLabel(" → "))
        steps_layout.addWidget(self.step3_label)
        steps_layout.addStretch()
        layout.addWidget(steps_widget)
        
        # 文件导入区域（精简拖拽上传）
        self.import_frame = QFrame()
        self.import_frame.setStyleSheet("""
            QFrame {
                border: 2px dashed #3498db;
                border-radius: 6px;
                background-color: #f8f9fa;
                min-height: 80px;
            }
            QFrame:hover { background-color: #ebf5fb; }
        """)
        self.setAcceptDrops(True)
        
        import_inner_layout = QVBoxLayout(self.import_frame)
        
        self.import_label = QLabel("拖拽Excel文件到此处 或 点击下方按钮")
        self.import_label.setAlignment(Qt.AlignCenter)
        self.import_label.setStyleSheet("font-size: 14px; color: #34495e; padding: 10px;")
        import_inner_layout.addWidget(self.import_label)
        
        # 按钮行
        btn_row = QHBoxLayout()
        browse_btn = QPushButton("选择文件")
        browse_btn.setFixedSize(100, 28)
        browse_btn.clicked.connect(self.browse_file)
        browse_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover { background-color: #2980b9; }
        """)
        
        btn_row.addStretch()
        btn_row.addWidget(browse_btn)
        btn_row.addStretch()
        import_inner_layout.addLayout(btn_row)
        
        layout.addWidget(self.import_frame)
        
        # 文件名显示
        self.file_name_label = QLabel("")
        self.file_name_label.setAlignment(Qt.AlignCenter)
        self.file_name_label.setStyleSheet("color: #27ae60; font-size: 12px; padding: 5px;")
        layout.addWidget(self.file_name_label)
        
        # 操作按钮行
        action_layout = QHBoxLayout()
        
        self.process_btn = QPushButton("开始分析")
        self.process_btn.setEnabled(False)
        self.process_btn.clicked.connect(self.process_data)
        self.process_btn.setStyleSheet("""
            QPushButton {
                background-color: #9b59b6;
                color: white;
                padding: 10px;
                border-radius: 4px;
                font-weight: bold;
                font-size: 14px;
                min-width: 120px;
            }
            QPushButton:hover { background-color: #8e44ad; }
            QPushButton:disabled { background-color: #bdc3c7; }
        """)
        
        self.export_btn = QPushButton("导出结果")
        self.export_btn.setEnabled(False)
        self.export_btn.clicked.connect(self.export_data)
        self.export_btn.setStyleSheet("""
            QPushButton {
                background-color: #e67e22;
                color: white;
                padding: 10px;
                border-radius: 4px;
                font-weight: bold;
                font-size: 14px;
                min-width: 120px;
            }
            QPushButton:hover { background-color: #d35400; }
            QPushButton:disabled { background-color: #bdc3c7; }
        """)
        
        action_layout.addStretch()
        action_layout.addWidget(self.process_btn)
        action_layout.addWidget(self.export_btn)
        action_layout.addStretch()
        layout.addLayout(action_layout)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #bdc3c7;
                border-radius: 3px;
                text-align: center;
                height: 22px;
            }
            QProgressBar::chunk {
                background-color: #3498db;
                border-radius: 2px;
            }
        """)
        layout.addWidget(self.progress_bar)
        
        # 状态标签
        self.status_label = QLabel("就绪")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("font-size: 12px; color: #7f8c8d; padding: 5px;")
        layout.addWidget(self.status_label)
        
        # 数据预览表格（精简显示）
        preview_frame = QGroupBox("数据预览")
        preview_frame.setStyleSheet("QGroupBox { font-weight: bold; }")
        preview_layout = QVBoxLayout()
        
        self.table_widget = QTableWidget()
        self.table_widget.setAlternatingRowColors(True)
        self.table_widget.setStyleSheet("""
            QTableWidget {
                gridline-color: #bdc3c7;
                font-size: 11px;
            }
        """)
        self.table_widget.setColumnCount(3)
        self.table_widget.setHorizontalHeaderLabels(['序号', '原始商品编码', '处理状态'])
        
        preview_layout.addWidget(self.table_widget)
        preview_frame.setLayout(preview_layout)
        layout.addWidget(preview_frame)
        
        # 状态栏
        self.statusBar().showMessage('就绪')
        
        # 关闭启动画面
        self.splash.finish(self)
        
    def load_databases_async(self):
        """异步加载数据库"""
        self.product_db_label.setText("商品库: 加载中...")
        self.combo_db_label.setText("组合库: 加载中...")
        
        self.load_thread = LoadDatabasesThread(self.product_db_path, self.combo_db_path)
        self.load_thread.finished.connect(self.on_databases_loaded)
        self.load_thread.error.connect(self.on_database_error)
        self.load_thread.start()
        
    def on_databases_loaded(self, product_df, combo_df):
        """数据库加载完成"""
        self.product_df = product_df
        self.combo_df = combo_df
        
        # 更新显示
        if product_df is not None:
            self.product_db_label.setText(f"商品库: {len(product_df)}条")
            self.product_db_label.setStyleSheet("color: #27ae60;")
        else:
            self.product_db_label.setText("商品库: 缺失")
            self.product_db_label.setStyleSheet("color: #e74c3c;")
            
        if combo_df is not None:
            self.combo_db_label.setText(f"组合库: {len(combo_df)}条")
            self.combo_db_label.setStyleSheet("color: #27ae60;")
        else:
            self.combo_db_label.setText("组合库: 缺失")
            self.combo_db_label.setStyleSheet("color: #e74c3c;")
            
        self.statusBar().showMessage('数据库加载完成')
        
    def on_database_error(self, error_msg):
        """数据库加载错误"""
        QMessageBox.warning(self, "数据库加载错误", error_msg)
        self.product_db_label.setText("商品库: 错误")
        self.combo_db_label.setText("组合库: 错误")
        
    def dragEnterEvent(self, event):
        """拖拽进入"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.import_frame.setStyleSheet("""
                QFrame {
                    border: 2px solid #27ae60;
                    border-radius: 6px;
                    background-color: #ebf5fb;
                    min-height: 80px;
                }
            """)
            
    def dragLeaveEvent(self, event):
        """拖拽离开"""
        self.import_frame.setStyleSheet("""
            QFrame {
                border: 2px dashed #3498db;
                border-radius: 6px;
                background-color: #f8f9fa;
                min-height: 80px;
            }
            QFrame:hover { background-color: #ebf5fb; }
        """)
        
    def dropEvent(self, event):
        """释放拖拽"""
        self.import_frame.setStyleSheet("""
            QFrame {
                border: 2px dashed #3498db;
                border-radius: 6px;
                background-color: #f8f9fa;
                min-height: 80px;
            }
        """)
        
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        if files and files[0].endswith(('.xlsx', '.xls')):
            self.load_file(files[0])
            
    def browse_file(self):
        """浏览文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择Excel文件", "", "Excel文件 (*.xlsx *.xls)"
        )
        if file_path:
            self.load_file(file_path)
            
    def load_file(self, file_path):
        """加载文件"""
        try:
            self.input_df = pd.read_excel(file_path)
            
            if '原始商品编码' not in self.input_df.columns:
                QMessageBox.warning(self, "警告", "文件中没有找到'原始商品编码'字段")
                return
                
            # 显示文件名
            self.file_name_label.setText(f"已加载: {os.path.basename(file_path)}")
            
            # 预览数据（只显示关键信息）
            preview_df = self.input_df[['原始商品编码']].head(100).copy()
            preview_df.insert(0, '序号', range(1, len(preview_df)+1))
            preview_df['处理状态'] = '待处理'
            
            self.display_preview(preview_df)
            
            # 更新步骤和按钮状态
            self.step1_label.setStyleSheet("color: #27ae60; font-weight: bold;")
            self.step2_label.setStyleSheet("color: #3498db;")
            self.step3_label.setStyleSheet("")
            
            self.process_btn.setEnabled(True)
            self.export_btn.setEnabled(False)
            
            self.status_label.setText(f"已加载 {len(self.input_df)} 条数据，点击【开始分析】处理")
            self.statusBar().showMessage(f'文件加载成功: {os.path.basename(file_path)}')
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"读取文件失败：{str(e)}")
            
    def display_preview(self, df):
        """显示预览数据"""
        self.table_widget.clear()
        self.table_widget.setRowCount(min(len(df), 100))
        self.table_widget.setColumnCount(len(df.columns))
        self.table_widget.setHorizontalHeaderLabels(df.columns.tolist())
        
        for row in range(self.table_widget.rowCount()):
            for col in range(self.table_widget.columnCount()):
                item = QTableWidgetItem(str(df.iat[row, col]))
                item.setTextAlignment(Qt.AlignCenter)
                self.table_widget.setItem(row, col, item)
                
        self.table_widget.resizeColumnsToContents()
        
    def process_data(self):
        """处理数据（带进度条）"""
        if self.input_df is None:
            return
        if self.product_df is None or self.combo_df is None:
            QMessageBox.warning(self, "警告", "数据库未加载，请点击【刷新】重试")
            return
            
        # 显示进度条
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.status_label.setText("分析中...")
        self.process_btn.setEnabled(False)
        QApplication.processEvents()
        
        try:
            df = self.input_df.copy()
            
            # 步骤1: 去重和过滤 (10%)
            self.progress_bar.setValue(10)
            df = df.drop_duplicates(subset=['原始商品编码'])
            df = df[df['原始商品编码'].astype(str).str.contains('*', regex=False, na=False)]
            
            # 步骤2: 检查单品和组合 (30%)
            self.progress_bar.setValue(30)
            product_codes = self.product_df['商品编码'].astype(str).tolist()
            combo_codes = self.combo_df['组合商品编码'].astype(str).tolist()
            
            df['单品存在'] = df['原始商品编码'].astype(str).apply(lambda x: x in product_codes)
            df['组合存在'] = df['原始商品编码'].astype(str).apply(lambda x: x in combo_codes)
            df = df[~(df['单品存在']) & ~(df['组合存在'])]
            
            # 步骤3: 拆分编码 (50%)
            self.progress_bar.setValue(50)
            df['商品编码'] = df['原始商品编码'].astype(str).apply(
                lambda x: x.rsplit('*', 1)[0] if '*' in x else ''
            )
            df['数量'] = df['原始商品编码'].astype(str).apply(
                lambda x: x.rsplit('*', 1)[1] if '*' in x and len(x.rsplit('*', 1)) > 1 else ''
            )
            
            # 步骤4: 查找商品名称 (70%)
            self.progress_bar.setValue(70)
            def find_product_name(code):
                if not code:
                    return ''
                match = self.product_df[self.product_df['商品编码'].astype(str) == code]
                if not match.empty:
                    return str(match.iloc[0].get('商品名称', ''))
                return ''
            
            df['临时名称'] = df['商品编码'].apply(find_product_name)
            
            # 步骤5: 提取中文字符 (85%)
            self.progress_bar.setValue(85)
            def extract_chinese(text):
                if not isinstance(text, str):
                    return ''
                reversed_text = text[::-1]
                match = re.match(r'^[^\d\u0030-\u0039]*', reversed_text)
                if match:
                    return match.group()[::-1]
                return text
            
            df['提取后'] = df['临时名称'].apply(extract_chinese)
            
            # 步骤6: 生成组合商品名称（用于内部，但不导出）(95%)
            self.progress_bar.setValue(95)
            def create_combo_name(row):
                try:
                    temp_name = str(row['临时名称']) if pd.notna(row['临时名称']) else ''
                    code = str(row['商品编码'])
                    match = self.product_df[self.product_df['商品编码'].astype(str) == code]
                    
                    size = ''
                    pcs = ''
                    color = ''
                    
                    if not match.empty:
                        if '尺寸规格(mm)' in match.columns:
                            size = str(match.iloc[0].get('尺寸规格(mm)', ''))
                        if '数量(pcs)' in match.columns:
                            pcs_str = str(match.iloc[0].get('数量(pcs)', ''))
                            try:
                                pcs_value = float(pcs_str) if pcs_str else 0
                                multiple = float(row['数量']) if row['数量'] else 0
                                pcs = str(int(pcs_value * multiple))
                            except:
                                pcs = ''
                        if '颜色' in match.columns:
                            color = str(match.iloc[0].get('颜色', ''))
                    
                    return f"{temp_name}{size}{pcs}{color}"
                except:
                    return ''
            
            df['组合商品名称'] = df.apply(create_combo_name, axis=1)
            
            # 保存处理结果
            self.processed_df = df[['原始商品编码', '商品编码', '数量']].copy()
            self.processed_df.rename(columns={'原始商品编码': '组合商品编码'}, inplace=True)
            
            # 完成
            self.progress_bar.setValue(100)
            
            # 更新预览
            preview_df = self.processed_df.copy()
            preview_df.insert(0, '序号', range(1, len(preview_df)+1))
            self.display_preview(preview_df)
            
            # 更新状态
            self.export_btn.setEnabled(True)
            self.step2_label.setStyleSheet("color: #27ae60; font-weight: bold;")
            self.step3_label.setStyleSheet("color: #e67e22;")
            self.status_label.setText(f"分析完成！共 {len(self.processed_df)} 条记录，可导出结果")
            self.statusBar().showMessage('数据分析完成')
            
            # 隐藏进度条
            QTimer.singleShot(1000, lambda: self.progress_bar.setVisible(False))
            
        except Exception as e:
            self.progress_bar.setVisible(False)
            self.process_btn.setEnabled(True)
            QMessageBox.critical(self, "错误", f"分析失败：{str(e)}")
            self.status_label.setText("分析失败")
            
    def export_data(self):
        """导出数据（不包含组合商品名称）"""
        if self.processed_df is None:
            return
            
        try:
            # 导出数据 - 不包含组合商品名称
            export_df = self.processed_df[['组合商品编码', '商品编码', '数量']].copy()
            
            # 生成文件名
            current_date = datetime.now().strftime("%m%d")
            default_filename = f"组合装数据{current_date}.xlsx"
            
            file_path, _ = QFileDialog.getSaveFileName(
                self, "保存文件", default_filename, "Excel文件 (*.xlsx)"
            )
            
            if file_path:
                if not file_path.endswith('.xlsx'):
                    file_path += '.xlsx'
                    
                export_df.to_excel(file_path, index=False)
                
                QMessageBox.information(self, "成功", f"数据导出成功！\n共 {len(export_df)} 条记录")
                self.statusBar().showMessage('导出成功')
                
                # 打开文件夹
                os.startfile(os.path.dirname(file_path))
                
        except Exception as e:
            QMessageBox.critical(self, "错误", f"导出失败：{str(e)}")

def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    # 设置全局字体
    font = QFont()
    font.setPointSize(10)
    app.setFont(font)
    
    processor = ExcelDataProcessor()
    processor.show()
    
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
