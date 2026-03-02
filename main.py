import sys
import os
import pandas as pd
from datetime import datetime
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import re

class ExcelDataProcessor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.product_db_path = r'\\Desktop-inv4qoc\图片数据\Temu_半托项目组\倒表格\数据\组合分析\商品资料.xlsx'
        self.combo_db_path = r'\\Desktop-inv4qoc\图片数据\Temu_半托项目组\倒表格\数据\组合分析\组合资料.xlsx'
        self.product_df = None
        self.combo_df = None
        self.input_df = None
        self.initUI()
        self.load_databases()
        
    def initUI(self):
        self.setWindowTitle('组合装数据分析工具')
        self.setGeometry(100, 100, 1000, 700)
        
        # 中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # 标题
        title_label = QLabel('组合装数据分析处理工具')
        title_label.setStyleSheet("""
            QLabel {
                font-size: 24px;
                font-weight: bold;
                color: #2c3e50;
                padding: 20px;
                text-align: center;
            }
        """)
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)
        
        # 数据库状态显示
        db_frame = QGroupBox("数据库状态")
        db_layout = QHBoxLayout()
        
        self.product_db_label = QLabel("商品资料库: 未加载")
        self.product_db_label.setStyleSheet("color: #e74c3c; font-weight: bold;")
        
        self.combo_db_label = QLabel("组合资料库: 未加载")
        self.combo_db_label.setStyleSheet("color: #e74c3c; font-weight: bold;")
        
        reload_btn = QPushButton("重新加载数据库")
        reload_btn.clicked.connect(self.load_databases)
        reload_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                padding: 8px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        
        db_layout.addWidget(self.product_db_label)
        db_layout.addWidget(self.combo_db_label)
        db_layout.addWidget(reload_btn)
        db_layout.addStretch()
        db_frame.setLayout(db_layout)
        layout.addWidget(db_frame)
        
        # 文件导入区域
        import_frame = QGroupBox("数据导入")
        import_layout = QVBoxLayout()
        
        file_layout = QHBoxLayout()
        self.file_path_edit = QLineEdit()
        self.file_path_edit.setPlaceholderText("请选择要处理的Excel文件...")
        browse_btn = QPushButton("浏览文件")
        browse_btn.clicked.connect(self.browse_file)
        file_layout.addWidget(self.file_path_edit)
        file_layout.addWidget(browse_btn)
        
        self.preview_btn = QPushButton("预览数据")
        self.preview_btn.clicked.connect(self.preview_data)
        self.preview_btn.setEnabled(False)
        self.preview_btn.setStyleSheet("""
            QPushButton {
                background-color: #2ecc71;
                color: white;
                padding: 10px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:disabled {
                background-color: #95a5a6;
            }
        """)
        
        import_layout.addLayout(file_layout)
        import_layout.addWidget(self.preview_btn)
        import_frame.setLayout(import_layout)
        layout.addWidget(import_frame)
        
        # 数据处理按钮
        process_btn = QPushButton("开始处理数据")
        process_btn.clicked.connect(self.process_data)
        process_btn.setStyleSheet("""
            QPushButton {
                background-color: #9b59b6;
                color: white;
                padding: 12px;
                border-radius: 4px;
                font-size: 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #8e44ad;
            }
        """)
        layout.addWidget(process_btn)
        
        # 进度显示
        self.progress_label = QLabel("就绪")
        self.progress_label.setAlignment(Qt.AlignCenter)
        self.progress_label.setStyleSheet("font-size: 14px; color: #7f8c8d;")
        layout.addWidget(self.progress_label)
        
        # 数据预览表格
        preview_frame = QGroupBox("数据预览")
        preview_layout = QVBoxLayout()
        
        self.table_widget = QTableWidget()
        self.table_widget.setAlternatingRowColors(True)
        self.table_widget.setStyleSheet("""
            QTableWidget {
                gridline-color: #bdc3c7;
                font-size: 12px;
            }
            QTableWidget::item {
                padding: 5px;
            }
        """)
        
        preview_layout.addWidget(self.table_widget)
        preview_frame.setLayout(preview_layout)
        layout.addWidget(preview_frame)
        
        # 导出按钮
        export_btn = QPushButton("导出结果")
        export_btn.clicked.connect(self.export_data)
        export_btn.setStyleSheet("""
            QPushButton {
                background-color: #e67e22;
                color: white;
                padding: 12px;
                border-radius: 4px;
                font-size: 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #d35400;
            }
        """)
        layout.addWidget(export_btn)
        
        # 状态栏
        self.statusBar().showMessage('就绪')
        
    def load_databases(self):
        """加载商品资料和组合资料数据库"""
        try:
            # 加载商品资料
            if os.path.exists(self.product_db_path):
                self.product_df = pd.read_excel(self.product_db_path)
                self.product_db_label.setText(f"商品资料库: 已加载 ({len(self.product_df)} 条记录)")
                self.product_db_label.setStyleSheet("color: #27ae60; font-weight: bold;")
            else:
                self.product_db_label.setText("商品资料库: 文件不存在")
                self.product_db_label.setStyleSheet("color: #e74c3c; font-weight: bold;")
                
            # 加载组合资料
            if os.path.exists(self.combo_db_path):
                self.combo_df = pd.read_excel(self.combo_db_path)
                self.combo_db_label.setText(f"组合资料库: 已加载 ({len(self.combo_df)} 条记录)")
                self.combo_db_label.setStyleSheet("color: #27ae60; font-weight: bold;")
            else:
                self.combo_db_label.setText("组合资料库: 文件不存在")
                self.combo_db_label.setStyleSheet("color: #e74c3c; font-weight: bold;")
                
            if self.product_df is not None and self.combo_df is not None:
                self.statusBar().showMessage('数据库加载成功')
            else:
                self.statusBar().showMessage('警告：部分数据库加载失败')
                
        except Exception as e:
            QMessageBox.critical(self, "错误", f"加载数据库时出错：{str(e)}")
            self.statusBar().showMessage('数据库加载失败')
            
    def browse_file(self):
        """浏览选择Excel文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择Excel文件", "", "Excel文件 (*.xlsx *.xls)"
        )
        if file_path:
            self.file_path_edit.setText(file_path)
            self.preview_btn.setEnabled(True)
            
    def preview_data(self):
        """预览数据"""
        file_path = self.file_path_edit.text()
        if not file_path:
            QMessageBox.warning(self, "警告", "请先选择文件")
            return
            
        try:
            self.input_df = pd.read_excel(file_path)
            
            if '原始商品编码' not in self.input_df.columns:
                QMessageBox.warning(self, "警告", "文件中没有找到'原始商品编码'字段")
                return
                
            # 显示数据
            self.display_data(self.input_df)
            self.statusBar().showMessage(f'已加载 {len(self.input_df)} 条记录')
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"读取文件时出错：{str(e)}")
            
    def display_data(self, df):
        """在表格中显示数据"""
        self.table_widget.clear()
        self.table_widget.setRowCount(df.shape[0])
        self.table_widget.setColumnCount(df.shape[1])
        
        # 设置表头
        self.table_widget.setHorizontalHeaderLabels(df.columns.tolist())
        
        # 填充数据
        for row in range(df.shape[0]):
            for col in range(df.shape[1]):
                item = QTableWidgetItem(str(df.iat[row, col]))
                item.setTextAlignment(Qt.AlignCenter)
                self.table_widget.setItem(row, col, item)
                
        # 调整列宽
        self.table_widget.resizeColumnsToContents()
        
    def process_data(self):
        """处理数据"""
        if self.input_df is None:
            QMessageBox.warning(self, "警告", "请先导入数据")
            return
            
        if self.product_df is None or self.combo_df is None:
            QMessageBox.warning(self, "警告", "数据库未正确加载")
            return
            
        try:
            self.progress_label.setText("正在处理数据...")
            QApplication.processEvents()
            
            # 步骤1: 基础处理
            df = self.input_df.copy()
            
            # 1).去重
            df = df.drop_duplicates(subset=['原始商品编码'])
            
            # 2).删除不含"*"的行
            df = df[df['原始商品编码'].astype(str).str.contains('*', regex=False, na=False)]
            
            # 3).新增字段
            df['单品是否存在'] = ''
            df['组合是否存在'] = ''
            df['倍数前'] = ''
            df['倍数后'] = ''
            df['组合商品名称'] = ''
            df['临时名称'] = ''
            df['提取后'] = ''
            
            # 步骤2: 计算过程
            # 1).检查单品是否存在
            product_codes = self.product_df['商品编码'].astype(str).tolist()
            df['单品是否存在'] = df['原始商品编码'].astype(str).apply(
                lambda x: '存在' if x in product_codes else ''
            )
            
            # 2).检查组合是否存在
            combo_codes = self.combo_df['组合商品编码'].astype(str).tolist()
            df['组合是否存在'] = df['原始商品编码'].astype(str).apply(
                lambda x: '存在' if x in combo_codes else ''
            )
            
            # 3).删除存在的行
            df = df[~(df['单品是否存在'] == '存在') & ~(df['组合是否存在'] == '存在')]
            
            # 4).提取"*"之前的内容（从右往左）
            df['倍数前'] = df['原始商品编码'].astype(str).apply(
                lambda x: x.rsplit('*', 1)[0] if '*' in x else ''
            )
            
            # 5).提取"*"之后的内容（从右往左）
            df['倍数后'] = df['原始商品编码'].astype(str).apply(
                lambda x: x.rsplit('*', 1)[1] if '*' in x and len(x.rsplit('*', 1)) > 1 else ''
            )
            
            # 6).查找商品名称
            def find_product_name(code):
                if not code:
                    return ''
                # 在商品资料中查找
                match = self.product_df[self.product_df['商品编码'].astype(str) == code]
                if not match.empty:
                    return match.iloc[0]['商品名称'] if '商品名称' in match.columns else ''
                return ''
            
            df['临时名称'] = df['倍数前'].apply(find_product_name)
            
            # 7).提取中文字符（从右往左提取数字之前的所有中文字符）
            def extract_chinese(text):
                if not isinstance(text, str):
                    return ''
                # 从右往左查找第一个数字，然后提取该数字之前的所有中文字符
                text = str(text)
                # 反转字符串以便从右往左处理
                reversed_text = text[::-1]
                
                # 使用正则表达式匹配从右往左直到遇到数字
                # 匹配非数字字符（主要是中文字符）
                match = re.match(r'^[^\d\u0030-\u0039]*', reversed_text)
                
                if match:
                    # 提取匹配的部分并反转回正常顺序
                    extracted = match.group()[::-1]
                    return extracted
                return text  # 如果没有找到数字，返回整个字符串
            
            df['提取后'] = df['临时名称'].apply(extract_chinese)
            
            # 8).组合商品名称
            def create_combo_name(row):
                try:
                    temp_name = str(row['临时名称']) if pd.notna(row['临时名称']) else ''
                    # 查找尺寸规格
                    code = str(row['倍数前'])
                    match = self.product_df[self.product_df['商品编码'].astype(str) == code]
                    
                    size = ''
                    pcs = ''
                    color = ''
                    
                    if not match.empty:
                        if '尺寸规格(mm)' in match.columns:
                            size = str(match.iloc[0]['尺寸规格(mm)']) if pd.notna(match.iloc[0]['尺寸规格(mm)']) else ''
                        if '数量(pcs)' in match.columns:
                            pcs_str = str(match.iloc[0]['数量(pcs)']) if pd.notna(match.iloc[0]['数量(pcs)']) else ''
                            try:
                                pcs_value = float(pcs_str) if pcs_str else 0
                                multiple = float(row['倍数后']) if row['倍数后'] else 0
                                pcs = str(int(pcs_value * multiple))
                            except:
                                pcs = ''
                        if '颜色' in match.columns:
                            color = str(match.iloc[0]['颜色']) if pd.notna(match.iloc[0]['颜色']) else ''
                    
                    return f"{temp_name}{size}{pcs}{color}"
                except:
                    return ''
            
            df['组合商品名称'] = df.apply(create_combo_name, axis=1)
            
            # 保存处理后的数据
            self.processed_df = df
            
            # 显示处理后的数据
            self.display_data(df)
            
            self.progress_label.setText("数据处理完成")
            self.statusBar().showMessage(f'处理完成，剩余 {len(df)} 条记录')
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"处理数据时出错：{str(e)}")
            self.progress_label.setText("数据处理失败")
            
    def export_data(self):
        """导出数据"""
        if not hasattr(self, 'processed_df') or self.processed_df is None:
            QMessageBox.warning(self, "警告", "请先处理数据")
            return
            
        try:
            # 准备导出数据
            export_df = pd.DataFrame()
            export_df['组合商品编码'] = self.processed_df['原始商品编码']
            export_df['组合商品名称'] = self.processed_df['组合商品名称']
            export_df['商品编码'] = self.processed_df['倍数前']
            export_df['数量'] = self.processed_df['倍数后']
            
            # 生成文件名
            current_date = datetime.now().strftime("%m%d")
            default_filename = f"组合装数据{current_date}.xlsx"
            
            # 选择保存位置
            file_path, _ = QFileDialog.getSaveFileName(
                self, "保存文件", default_filename, "Excel文件 (*.xlsx)"
            )
            
            if file_path:
                # 确保文件以.xlsx结尾
                if not file_path.endswith('.xlsx'):
                    file_path += '.xlsx'
                    
                # 导出数据
                export_df.to_excel(file_path, index=False)
                
                QMessageBox.information(self, "成功", f"数据已成功导出到：\n{file_path}")
                self.statusBar().showMessage('数据导出成功')
                
                # 打开文件所在文件夹
                os.startfile(os.path.dirname(file_path))
                
        except Exception as e:
            QMessageBox.critical(self, "错误", f"导出数据时出错：{str(e)}")
            
def main():
    app = QApplication(sys.argv)
    
    # 设置应用程序样式
    app.setStyle('Fusion')
    
    # 创建并显示主窗口
    processor = ExcelDataProcessor()
    processor.show()
    
    sys.exit(app.exec_())
    
if __name__ == '__main__':
    main()
