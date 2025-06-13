import sys
import os
import pandas as pd
import re
from datetime import datetime
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import platform
import subprocess
import traceback
import uuid
from tkinter.filedialog import asksaveasfilename

# AQL 抽样表 (constant, kept global for simplicity)
AQL_SAMPLE_SIZE = {
    1.0: {(2, 8): 8, (9, 15): 13, (16, 25): 13, (26, 50): 13, (51, 90): 13, (91, 150): 19, (151, 280): 29,
          (281, 500): 29, (501, 1200): 34, (1201, 3200): 42, (3201, 10000): 50},
    1.5: {(2, 8): 8, (9, 15): 8, (16, 25): 8, (26, 50): 8, (51, 90): 13, (91, 150): 19, (151, 280): 19, (281, 500): 21,
          (501, 1200): 27, (1201, 3200): 35, (3201, 10000): 38},
    2.5: {(2, 8): 5, (9, 15): 5, (16, 25): 5, (26, 50): 7, (51, 90): 11, (91, 150): 11, (151, 280): 13, (281, 500): 16,
          (501, 1200): 19, (1201, 3200): 23, (3201, 10000): 29},
    4.0: {(2, 8): 3, (9, 15): 3, (16, 25): 3, (26, 50): 7, (51, 90): 8, (91, 150): 9, (151, 280): 10, (281, 500): 11,
          (501, 1200): 15, (1201, 3200): 18, (3201, 10000): 22},
    10.0: {(2, 8): 3, (9, 15): 3, (16, 25): 3, (26, 50): 3, (51, 90): 4, (91, 150): 5, (151, 280): 6, (281, 500): 7,
           (501, 1200): 8, (1201, 3200): 9, (3201, 10000): 9}
}

def resource_path(relative_path):
    """获取运行时文件的绝对路径（支持 PyInstaller 打包）"""
    try:
        # PyInstaller 创建的临时文件夹路径
        base_path = sys._MEIPASS
    except AttributeError:
        # 非打包环境，直接使用当前目录
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class BatchLaborTimeProcessor:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("批次检验工时计算")
        self.root.geometry("1000x980")
        self.root.resizable(False, False)
        self.root.configure(bg="#f0f0f0")

        # Initialize instance variables
        self.current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        self.debug_info = []
        self.output_file_path = None
        self.df_test = None
        self.tool_to_hours_base = None

        # Set up ttk style
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TButton", font=("Microsoft YaHei", 15), padding=12)
        style.configure("TLabel", font=("Microsoft YaHei", 18), background="#f0f0f0", foreground="#333333")
        style.configure("TEntry", font=("Microsoft YaHei", 15))
        style.configure("Green.TButton", background="#28a745", foreground="white")
        style.configure("Blue.TButton", background="#0078d7", foreground="white")

        # Main frame
        self.main_frame = tk.Frame(self.root, bg="#f0f0f0", relief="groove", borderwidth=2)
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=10)

        # Title
        ttk.Label(self.main_frame, text="批次检验工时计算", font=("Microsoft YaHei", 18, "bold")).grid(row=0, column=0,
                                                                                             columnspan=3, pady=10)

        # File path input
        ttk.Label(self.main_frame, text="检验卡文件路径：").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.file_entry = ttk.Entry(self.main_frame, width=60)
        self.file_entry.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        ttk.Button(self.main_frame, text="浏览", command=self.browse_file).grid(row=2, column=2, padx=5, pady=5)

        # Batch size input
        ttk.Label(self.main_frame, text="批量大小：").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.batch_entry = ttk.Entry(self.main_frame, width=20)
        self.batch_entry.grid(row=4, column=0, sticky=tk.W, pady=5)
        ttk.Label(self.main_frame, text="（留空将使用历史开单批次大小的中位数计算默认批量大小）",
                  font=("Microsoft YaHei", 12)).grid(
            row=4, column=1, columnspan=2, sticky=tk.W, pady=5)

        # Progress bar
        self.progress_bar = ttk.Progressbar(self.main_frame, mode="indeterminate", length=400)
        self.progress_bar.grid(row=5, column=0, columnspan=3, pady=10)
        self.progress_bar.grid_remove()

        # Status label
        self.status_label = ttk.Label(self.main_frame, text="请输入文件路径和批量大小，然后点击提交。", wraplength=700,
                                      font=("Microsoft YaHei", 12))
        self.status_label.grid(row=6, column=0, columnspan=3, pady=10)

        # Buttons
        self.submit_button = ttk.Button(self.main_frame, text="提交", command=self.submit, style="Blue.TButton")
        self.submit_button.grid(row=7, column=0, padx=10, pady=10)
        self.open_button = ttk.Button(self.main_frame, text="打开输出文件", command=self.open_output,
                                      style="Green.TButton")
        self.open_button.grid(row=7, column=1, padx=10, pady=10)
        self.open_button.grid_remove()
        ttk.Button(self.main_frame, text="退出", command=self.root.quit).grid(row=7, column=2, padx=10, pady=10)

        # Rules display
        ttk.Label(self.main_frame,
                  text="检验规则说明：",
                  font=("Microsoft YaHei", 15, "bold"),
                  foreground="#e67e22").grid(
            row=8, column=0, columnspan=3, sticky=tk.W, pady=(10, 5)
        )

        rules_lines = self.get_rules_text().split('\n')
        row_index = 9
        for line in rules_lines:
            if not line.strip():
                ttk.Label(self.main_frame, text=" ").grid(row=row_index, column=0, sticky="w")
            elif line.startswith("⚠️"):
                ttk.Label(self.main_frame, text=line, font=("Microsoft YaHei", 15, "bold"), foreground="#e67e22").grid(
                    row=row_index, column=0, sticky="w")
            elif line.startswith("1.") or line.startswith("2."):
                ttk.Label(self.main_frame, text=line, font=("Microsoft YaHei", 12, "bold"), foreground="#2f3640").grid(
                    row=row_index, column=0, sticky="w")
            elif line.startswith("   •") or line.startswith("   →"):
                ttk.Label(self.main_frame, text=line, font=("Microsoft YaHei", 12), foreground="#555555").grid(
                    row=row_index, column=0, sticky="w")
            else:
                ttk.Label(self.main_frame, text=line, font=("Microsoft YaHei", 12), foreground="#333333").grid(
                    row=row_index, column=0, sticky="w")
            row_index += 1
        self.add_footer(self.main_frame, row=row_index)
        self.debug_info.append({
            '时间戳': self.current_time,
            '类别': '信息',
            '信息': 'GUI初始化成功，包含检验规则说明'
        })

    def add_footer(self, frame, row):
        footer_text = (
            "© 2025 批次检验工时计算参考 V1.0 | 开发者: Chen Jinzhuo  "
            "| 最后更新: 2025-06-10 | 问题反馈: Jinzhuo.Chen@Zimmerbiomet.com"
        )

        footer_frame = ttk.Frame(frame)
        footer_frame.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(10, 5))

        label = ttk.Label(
            footer_frame,
            text=footer_text,
            font=("Microsoft YaHei", 9),
            foreground="#666666"
        )
        label.pack(side="left")

    def get_rules_text(self) -> str:
        """Return the formatted rules text for display in the GUI."""
        return """未包含计算逻辑如下，需自行判断

        1.批次准备工时 （+0.1H/ lot）
          1.1 投影仪多个视角摆放
          1.2 终检/外协拆包装检验（如Web3-FEM PRO 42-5047-066-02）
          1.3 托盘包装的DRILL产品

        2. 检验时间（自行调整）
          2.1除CMM等非专员操作检验外的其他自动测量程序
          - OGP程序测量 
          - 闪测仪程序测量
          - 轮廓仪程序测量
          - 粗糙度仪
          2.2属性检验项
          - 孔检验数量 
          - 螺纹检验数量
          - 单份投影纸检验视图数量 
          - 功能规:每次检验覆盖的检验项数
          """

    def validate_file(self, file_path: str, description: str) -> tuple[bool, str]:
        """Validate if a file exists and is readable."""
        try:
            file_path = resource_path(file_path)  # 使用动态路径
            if not os.path.isfile(file_path):
                raise FileNotFoundError(f"{description} 文件 {file_path} 不存在。")
            if not os.access(file_path, os.R_OK):
                raise PermissionError(f"无权限读取 {description} 文件 {file_path}。")
            return True, ""
        except (FileNotFoundError, PermissionError) as e:
            self.debug_info.append({
                '时间戳': self.current_time,
                '类别': '错误',
                '信息': str(e)
            })
            return False, str(e)

    def check_unmatched_records(self, df: pd.DataFrame, key_column: str, key_dict: dict, display_columns: list,
                                unmatched_message: str, matched_message: str, default_value=None) -> pd.DataFrame:
        """Check for unmatched records in a DataFrame and log results."""
        try:
            key_set = set(key_dict.keys())
            unmatched_df = df[~df[key_column].isin(key_set)]
            if not unmatched_df.empty:
                message = unmatched_message.format(default_value) if default_value is not None else unmatched_message
                self.debug_info.append({
                    '时间戳': self.current_time,
                    '类别': '警告',
                    '信息': f"{message}\n{unmatched_df[display_columns].to_string(index=False)}"
                })
            else:
                self.debug_info.append({
                    '时间戳': self.current_time,
                    '类别': '信息',
                    '信息': matched_message
                })
            return unmatched_df
        except Exception as e:
            self.debug_info.append({
                '时间戳': self.current_time,
                '类别': '错误',
                '信息': f"检查未匹配记录时出错：{e}"
            })
            raise

    def get_aql_sample_size(self, aql_level: float, lot_size: int) -> int:
        """Determine AQL sample size based on lot size and AQL level."""
        try:
            if aql_level not in AQL_SAMPLE_SIZE:
                available_levels = list(AQL_SAMPLE_SIZE.keys())
                aql_level = min(available_levels, key=lambda x: abs(x - aql_level))
                self.debug_info.append({
                    '时间戳': self.current_time,
                    '类别': '警告',
                    '信息': f"AQL级别 {aql_level} 未在抽样表中，使用最近的AQL级别 {aql_level}"
                })
            for (min_lot, max_lot), sample_size in AQL_SAMPLE_SIZE[aql_level].items():
                if min_lot <= lot_size <= max_lot:
                    return min(lot_size, sample_size)
            return 0
        except Exception as e:
            self.debug_info.append({
                '时间戳': self.current_time,
                '类别': '错误',
                '信息': f"获取AQL抽样量时出错：{e}"
            })
            return 0

    def open_output_file(self, file_path: str) -> tuple[bool, str]:
        """Open the output file using the appropriate system command."""
        try:
            if platform.system() == "Windows":
                os.startfile(file_path)
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["open", file_path], check=True)
            else:  # Linux
                subprocess.run(["xdg-open", file_path], check=True)
            self.debug_info.append({
                '时间戳': self.current_time,
                '类别': '信息',
                '信息': f"用户打开输出文件：{file_path}"
            })
            return True, "文件已打开"
        except Exception as e:
            self.debug_info.append({
                '时间戳': self.current_time,
                '类别': '错误',
                '信息': f"打开输出文件 {file_path} 时出错：{e}"
            })
            return False, f"错误：无法打开输出文件 {file_path}：{e}"

    def load_and_validate_batch_data(self) -> tuple[pd.DataFrame, dict, float]:
        """Load and validate batch_num.xlsx, returning DataFrame, median lot sizes, and default lot size."""
        df_batch = pd.read_excel(resource_path('batch_num.xlsx'), dtype={'Material Number': str})
        if df_batch.empty:
            raise ValueError("batch_num.xlsx 文件为空。")
        batch_required_columns = ['Material Number', 'Batch', 'Order quantity (GMEIN)']
        if not all(col in df_batch.columns for col in batch_required_columns):
            missing = [col for col in batch_required_columns if col not in df_batch.columns]
            raise ValueError(f"batch_num.xlsx 中缺少必要列：{missing}")

        df_batch = df_batch.rename(columns={'Order quantity (GMEIN)': '开单数量'})
        df_batch['开单数量'] = pd.to_numeric(df_batch['开单数量'], errors='coerce').fillna(0)
        lot_size_median = df_batch.groupby('Material Number')['开单数量'].median().to_dict()
        default_lot_size = df_batch['开单数量'].median()
        if default_lot_size == 0 or pd.isna(default_lot_size):
            raise ValueError("batch_num.xlsx 开单数量数据无效或全为0，默认批量大小无法计算。")

        self.debug_info.append({
            '时间戳': self.current_time,
            '类别': '分布信息',
            '信息': f"开单数量分布（描述性统计）：\n{df_batch['开单数量'].describe().to_string()}"
        })
        return df_batch, lot_size_median, default_lot_size

    def load_and_validate_database(self) -> dict:
        """Load and validate database.xlsx, returning tool-to-hours mapping."""
        df_database = pd.read_excel(resource_path('database.xlsx'))
        if df_database.empty:
            raise ValueError("database.xlsx 文件为空。")
        df_database.iloc[:, 0] = df_database.iloc[:, 0].astype(str).str.strip().str.upper().str.replace(r'\s+', ' ',
                                                                                                        regex=True)
        df_database.iloc[:, 2] = pd.to_numeric(df_database.iloc[:, 2], errors='coerce').fillna(0)
        tool_to_hours_base = dict(zip(df_database.iloc[:, 0], df_database.iloc[:, 2]))
        self.debug_info.append({
            '时间戳': self.current_time,
            '类别': '信息',
            '信息': f"tool_to_hours_base 键：{list(tool_to_hours_base.keys())}"
        })
        zero_hours_tools = {k: v for k, v in tool_to_hours_base.items() if v == 0}
        if zero_hours_tools:
            self.debug_info.append({
                '时间戳': self.current_time,
                '类别': '警告',
                '信息': f"以下量具编码的工时值为0：{zero_hours_tools}"
            })
        return tool_to_hours_base

    def load_and_validate_test_data(self, test_file_path: str) -> pd.DataFrame:
        """Load and validate test data from the specified file."""
        valid, msg = self.validate_file(test_file_path, "Test Data")
        if not valid:
            raise FileNotFoundError(msg)
        df_test = pd.read_excel(test_file_path, dtype={'产品编码': str})
        if df_test.empty:
            raise ValueError("whole_rawdata.xlsx 文件为空。")
        required_columns = ['产品编码', '检验卡编号', '版本', '工艺编号', '工序号', '工序名称', '量具1层编码',
                            '量具2层编码', '抽样频率', '是否检验', '尺寸内容']
        if not all(col in df_test.columns for col in required_columns):
            missing = [col for col in required_columns if col not in df_test.columns]
            raise ValueError(f"{test_file_path} 中缺少必要列：{missing}")

        df_test['量具1层编码'] = df_test['量具1层编码'].astype(str).str.strip().str.upper().str.replace(r'\s+', ' ',
                                                                                                        regex=True)
        df_test['工序号'] = df_test['工序号'].astype(str).str.zfill(4)
        df_test['工艺编号'] = df_test['工艺编号'].astype(str).str.strip().str.upper()
        df_test['抽样频率'] = df_test['抽样频率'].astype(str).str.strip().str.upper()
        duplicates = df_test.duplicated(subset=['工艺编号', '工序号'], keep=False)
        df_test = df_test.drop_duplicates(subset=['工艺编号', '工序号'], keep='first')
        self.debug_info.append({
            '时间戳': self.current_time,
            '类别': '信息',
            '信息': f"量具1层编码唯一值：{list(df_test['量具1层编码'].unique())}"
        })
        return df_test

    def load_and_validate_bom(self) -> pd.DataFrame:
        """Load and validate BOM.xlsx."""
        df_bom = pd.read_excel(resource_path('BOM.xlsx'))
        if df_bom.empty:
            raise ValueError("BOM.xlsx 文件为空。")
        bom_required_columns = ['Description', 'Operation/Activity N', 'Production Version', 'Setup Personal time',
                                'Labor time']
        if not all(col in df_bom.columns for col in bom_required_columns):
            missing = [col for col in bom_required_columns if col not in df_bom.columns]
            raise ValueError(f"BOM.xlsx 中缺少必要列：{missing}")

        df_bom['Description'] = df_bom['Description'].astype(str).str.strip().str.upper()
        df_bom['Operation/Activity N'] = df_bom['Operation/Activity N'].astype(str).str.zfill(4)
        df_bom = df_bom.drop_duplicates(subset=['Material',
                                                'Production Version',
                                                'Description',
                                                'Operation/Activity N',
                                                'Operation short text',
                                                'Setup Personal time',
                                                'Labor time'], keep='first')
        return df_bom

    def assign_batch_size(self, df_test: pd.DataFrame, lot_size_median: dict, default_lot_size: float,
                          user_batch_size: int | None) -> pd.DataFrame:
        """Assign batch size to test data."""
        if user_batch_size is not None:
            df_test['lot_size_median'] = int(float(user_batch_size))
            self.debug_info.append({
                '时间戳': self.current_time,
                '类别': '信息',
                '信息': f"使用用户提供的批量大小：{user_batch_size}"
            })
        else:
            df_test['lot_size_median'] = df_test['产品编码'].map(lot_size_median).fillna(default_lot_size)
            self.debug_info.append({
                '时间戳': self.current_time,
                '类别': '信息',
                '信息': f"用户未提供批量大小，使用默认批量大小：{default_lot_size}"
            })
        return df_test

    def calculate_sampling_quantity(self, df_test: pd.DataFrame) -> pd.DataFrame:
        """Calculate sampling quantity based on sampling frequency."""
        df_test['抽样数量'] = 0
        no_inspection = df_test['是否检验'] != '是'
        is_first_last_piece = df_test['抽样频率'] == '首末件'
        is_piece_per_lot = df_test['抽样频率'].str.contains('1件', case=False, na=False)
        is_full_inspection = df_test['抽样频率'].str.contains('100%', case=False, na=False)
        is_aql = df_test['抽样频率'].str.contains('AQL', case=False, na=False)

        df_test.loc[no_inspection, '抽样数量'] = 0
        df_test.loc[is_first_last_piece, '抽样数量'] = 2
        df_test.loc[is_piece_per_lot, '抽样数量'] = 1
        df_test.loc[is_full_inspection & (~is_piece_per_lot) & (~is_first_last_piece), '抽样数量'] = df_test[
            'lot_size_median']

        aql_rows = (~no_inspection) & (~is_full_inspection) & (~is_piece_per_lot) & (~is_first_last_piece) & is_aql
        if aql_rows.any():
            aql_levels = df_test.loc[aql_rows, '抽样频率'].str.extract(r'AQL\s*(\d*\.?\d+)', expand=False).astype(float)
            aql_rows_invalid = aql_rows & aql_levels.isna()
            if aql_rows_invalid.any():
                invalid_freq = df_test.loc[aql_rows_invalid, '抽样频率'].iloc[0]
                self.debug_info.append({
                    '时间戳': self.current_time,
                    '类别': '警告',
                    '信息': f"警告：无法提取有效的AQL级别，抽样频率 {invalid_freq}，抽样数量设为0 | "
                            f"检验卡编号: {df_test[aql_rows_invalid]['检验卡编号'].iloc[0]}, "
                            f"工序名称: {df_test[aql_rows_invalid]['工序名称'].iloc[0]}, "
                            f"量具1层编码: {df_test[aql_rows_invalid]['量具1层编码'].iloc[0] if pd.notna(df_test[aql_rows_invalid]['量具1层编码'].iloc[0]) else '无量具'}"
                })
                df_test.loc[aql_rows_invalid, '抽样数量'] = 0

            aql_rows_valid = aql_rows & aql_levels.notna()
            if aql_rows_valid.any():
                lot_sizes = df_test.loc[aql_rows_valid, 'lot_size_median'].round(0).astype(int)
                sample_sizes = pd.Series(0, index=df_test.index[aql_rows_valid])
                for aql_level in aql_levels[aql_rows_valid].unique():
                    mask = (aql_levels == aql_level) & aql_rows_valid
                    for (min_lot, max_lot), size in AQL_SAMPLE_SIZE.get(aql_level, {}).items():
                        lot_mask = (lot_sizes >= min_lot) & (lot_sizes <= max_lot)
                        sample_sizes.loc[mask & lot_mask] = np.minimum(lot_sizes[lot_mask], size).astype(int)

                out_of_range = (sample_sizes == 0) & aql_rows_valid
                if out_of_range.any():
                    self.debug_info.append({
                        '时间戳': self.current_time,
                        '类别': '警告',
                        '信息': f"警告：批次大小 {lot_sizes[out_of_range].iloc[0]} 超出范围，返回0 | "
                                f"检验卡编号: {df_test[out_of_range]['检验卡编号'].iloc[0]}, "
                                f"工序名称: {df_test[out_of_range]['工序名称'].iloc[0]}, "
                                f"量具1层编码: {df_test[out_of_range]['量具1层编码'].iloc[0] if pd.notna(df_test[out_of_range]['量具1层编码'].iloc[0]) else '无量具'}"
                    })
                df_test.loc[aql_rows_valid, '抽样数量'] = sample_sizes

        invalid_freq_rows = (~no_inspection) & (~is_full_inspection) & (~is_piece_per_lot) & (
            ~is_first_last_piece) & (~is_aql)
        if invalid_freq_rows.any():
            invalid_freq = df_test.loc[invalid_freq_rows, '抽样频率'].iloc[0]
            self.debug_info.append({
                '时间戳': self.current_time,
                '类别': '警告',
                '信息': f"警告：无效的抽样频率 {invalid_freq}，抽样数量设为0 | "
                        f"检验卡编号: {df_test[invalid_freq_rows]['检验卡编号'].iloc[0]}, "
                        f"工序名称: {df_test[invalid_freq_rows]['工序名称'].iloc[0]}, "
                        f"量具1层编码: {df_test[invalid_freq_rows]['量具1层编码'].iloc[0] if pd.notna(df_test[invalid_freq_rows]['量具1层编码'].iloc[0]) else '无量具'}"
            })
            df_test.loc[invalid_freq_rows, '抽样数量'] = 0
        return df_test

    def create_sheet1_data(self, df_test: pd.DataFrame, user_batch_size: int | None) -> pd.DataFrame:
        """Create data for Sheet1 from processed test data, including setup time calculation."""
        sheet1_columns = ['产品编码', '检验卡编号', '版本', '工艺编号', 'Production Version', '工序号', '工序名称',
                         '变更前Setup Personal time批量准备工时', '变更前Labor time单件工时', '变更后Setup Personal time批量准备工时', '变更后Labor time单件工时',
                         '批次工时', '批次大小']

        special_tools = [
            'OPTICAL MEASURING INSTRUMENT光学影像仪',
            'MYLAR / OVERLAY数字投影纸',
            'MYLAR / OVERLAY投影纸',
            'AIR COLUMN / AIR GAGE气动量具',
            'ROUGHNESS TESTER粗糙度仪',
            'CONTOURGRAPH轮廓仪'
        ]
        keywords = {
            '光学影像仪', '数字投影纸', '投影纸', '气动量具', '粗糙度仪', '轮廓仪'
        }

        grouped_hours = df_test.groupby(['产品编码', '检验卡编号', '版本', '工艺编号', '工序号', '工序名称']).agg({
            '工时': 'sum',
            '批次大小': 'first',
            'Production Version': 'first',
            'Setup Personal time': 'first',
            'Labor time': 'first',
            '量具1层编码': lambda x: list(x)
        }).reset_index().rename(columns={
            '工时': '批次工时',
            'Setup Personal time': '变更前Setup Personal time批量准备工时',
            'Labor time': '变更前Labor time单件工时',
            '平均单件工时': '变更后Labor time单件工时'
        })

        def calculate_setup_time(row):
            tools = row['量具1层编码']
            operation_name = row['工序名称']

            if not isinstance(tools, list):
                setup_time = 0.1
            else:
                unique_tools = set(tool.strip() for tool in tools if tool and isinstance(tool, str))
                count = 0
                for tool in unique_tools:
                    tool_lower = tool.lower()
                    if tool in special_tools or any(keyword.lower() in tool_lower for keyword in keywords):
                        count += 1
                setup_time = 0.1 + 0.1 * count

            if isinstance(operation_name, str) and '外协' in operation_name:
                setup_time += 0.1

            return setup_time

        grouped_hours['变更后Setup Personal time批量准备工时'] = grouped_hours.apply(calculate_setup_time, axis=1)

        self.debug_info.append({
            '时间戳': self.current_time,
            '类别': '信息',
            '信息': f"准备工时计算完成，基于特殊量具：{special_tools}，并为包含‘外协’的工序名称增加0.1工时"
        })

        if user_batch_size is not None:
            grouped_hours['批次大小'] = int(float(user_batch_size))
            self.debug_info.append({
                '时间戳': self.current_time,
                '类别': '信息',
                '信息': f"使用用户提供的批量大小 {user_batch_size} 计算'变更后Labor time单件工时'"
            })

        grouped_hours['变更后Labor time单件工时'] = (grouped_hours['批次工时'] / grouped_hours['批次大小'].replace(0, 1)).round(3)
        grouped_hours.loc[grouped_hours['批次大小'] == 0, '变更后Labor time单件工时'] = 0

        grouped_hours = grouped_hours.drop(columns=['量具1层编码'])
        return grouped_hours[sheet1_columns]

    def create_sheet2_data(self, df_test: pd.DataFrame) -> pd.DataFrame:
        """Create data for Sheet2 from processed test data."""
        sheet2_columns = ['产品编码', '检验卡编号', '版本', '工艺编号', 'Production Version', '工序号', '工序名称',
                          '尺寸内容', '量具1层编码', '量具2层编码', '批次大小', '抽样频率', '是否检验', '抽样数量',
                          '单件工时', '工时']
        return df_test[sheet2_columns]

    def process_data(self, test_file_path: str, user_batch_size: int | None) -> tuple[bool, str, pd.DataFrame, dict]:
        """Process input data, calculate labor hours, and save results."""
        try:
            for file, desc in [("batch_num.xlsx", "Batch Size"), ("database.xlsx", "Database"), ("BOM.xlsx", "BOM")]:
                valid, msg = self.validate_file(file, desc)
                if not valid:
                    return False, msg, None, None

            df_batch, lot_size_median, default_lot_size = self.load_and_validate_batch_data()
            self.tool_to_hours_base = self.load_and_validate_database()
            df_test = self.load_and_validate_test_data(test_file_path)
            df_bom = self.load_and_validate_bom()

            df_test = self.assign_batch_size(df_test, lot_size_median, default_lot_size, user_batch_size)

            empty_tools = df_test[df_test['量具1层编码'].isna() | (df_test['量具1层编码'] == '')][
                ['检验卡编号', '工序名称']]
            if not empty_tools.empty:
                self.debug_info.append({
                    '时间戳': self.current_time,
                    '类别': '警告',
                    '信息': f"以下记录的量具1层编码为空：\n{empty_tools.to_string(index=False)}"
                })

            unmatched_tools = df_test[~df_test['量具1层编码'].isin(self.tool_to_hours_base.keys()) &
                                      df_test['抽样频率'].isin(['AQL1.0 C=0', 'AQL2.5 C=0', 'AQL4.0 C=0'])][
                ['产品编码', '检验卡编号', '工序名称', '量具1层编码']]
            if not unmatched_tools.empty:
                self.debug_info.append({
                    '时间戳': self.current_time,
                    '类别': '警告',
                    '信息': f"以下记录的量具1层编码未在 tool_to_hours_base 中找到：\n{unmatched_tools.to_string(index=False)}"
                })

            df_test = self.calculate_sampling_quantity(df_test)

            df_test['单件工时'] = df_test['量具1层编码'].map(self.tool_to_hours_base).fillna(0)
            df_test['工时'] = df_test['抽样数量'] * df_test['单件工时']
            zero_hours_records = df_test[df_test['工时'] == 0][
                ['产品编码', '检验卡编号', '工序名称', '量具1层编码', '抽样数量', '单件工时', '工时']]
            if not zero_hours_records.empty:
                self.debug_info.append({
                    '时间戳': self.current_time,
                    '类别': '警告',
                    '信息': f"以下记录的工时为0：\n{zero_hours_records.to_string(index=False)}"
                })

            df_test = pd.merge(
                df_test,
                df_bom[
                    ['Description', 'Operation/Activity N', 'Production Version', 'Setup Personal time', 'Labor time']],
                how='left',
                left_on=['工艺编号', '工序号'],
                right_on=['Description', 'Operation/Activity N']
            ).drop(columns=['Description'])
            df_test['Production Version'] = df_test['Production Version'].fillna('N/A')
            df_test['Setup Personal time'] = df_test['Setup Personal time'].fillna('N/A')
            df_test['Labor time'] = df_test['Labor time'].fillna('N/A')

            self.check_unmatched_records(
                df_test, '产品编码', lot_size_median if user_batch_size is None else {},
                ['产品编码', '检验卡编号', '工序名称', '量具1层编码'],
                "警告：发现未匹配的产品编码（使用默认批次大小 {}）：", "所有产品编码均已匹配，无未匹配产品编码。",
                default_value=user_batch_size if user_batch_size is not None else default_lot_size
            )
            self.check_unmatched_records(
                df_test, '量具1层编码', self.tool_to_hours_base,
                ['产品编码', '检验卡编号', '工序名称', '量具1层编码'],
                "警告：发现未匹配的量具编码：", "所有量具编码均已匹配，无未匹配量具编码。"
            )

            df_test = df_test.rename(columns={'lot_size_median': '批次大小'})

            sheet1_data = self.create_sheet1_data(df_test, user_batch_size)
            sheet2_data = self.create_sheet2_data(df_test)

            unmatched_bom = sheet1_data[sheet1_data['Production Version'].isna()][['工艺编号']]
            if not unmatched_bom.empty:
                self.debug_info.append({
                    '时间戳': self.current_time,
                    '类别': '警告',
                    '信息': f"以下记录未在 BOM.xlsx 中找到匹配：\n{unmatched_bom.to_string(index=False)}"
                })

            zero_lot_size = sheet1_data[sheet1_data['批次大小'] == 0][
                ['产品编码', '检验卡编号', '工序名称', '批次工时', '批次大小']]
            if not zero_lot_size.empty:
                self.debug_info.append({
                    '时间戳': self.current_time,
                    '类别': '警告',
                    '信息': f"以下记录的默认批次大小为0，变更后Labor time单件工时设为0：\n{zero_lot_size.to_string(index=False)}"
                })

            output_file_path = asksaveasfilename(
                defaultextension='.xlsx',
                filetypes=[('Excel files', '*.xlsx'), ('All files', '*.*')],
                title="请选择保存路径"
            )

            if not output_file_path:
                print("用户取消了保存操作")
                return False, "用户取消了保存操作", None, None

            if os.path.exists(output_file_path):
                self.debug_info.append({
                    '时间戳': self.current_time,
                    '类别': '警告',
                    '信息': f"输出文件 {output_file_path} 已存在，将被覆盖"
                })
            if not os.access(os.path.dirname(output_file_path) or '.', os.W_OK):
                raise PermissionError(f"无权限写入文件 {output_file_path}")
            with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
                sheet1_data.to_excel(writer, sheet_name='Sheet1', index=False)
                sheet2_data.to_excel(writer, sheet_name='Sheet2', index=False)
                pd.DataFrame(self.debug_info).to_excel(writer, sheet_name='DebugInfo', index=False)

                for sheet_name in ['Sheet1', 'Sheet2']:
                    worksheet = writer.sheets[sheet_name]
                    col_idx = sheet1_data.columns.get_loc(
                        '工序号') if sheet_name == 'Sheet1' else sheet2_data.columns.get_loc('工序号')
                    text_format = writer.book.add_format({'num_format': '@'})
                    worksheet.set_column(0, 0, None, text_format)

            self.debug_info.append({
                '时间戳': self.current_time,
                '类别': '信息',
                '信息': f"处理完成，结果已保存至 {output_file_path}"
            })
            return True, output_file_path, df_test, self.tool_to_hours_base

        except Exception as e:
            self.debug_info.append({
                '时间戳': self.current_time,
                '类别': '错误',
                '信息': f"处理数据时出错：{e}\n{traceback.format_exc()}"
            })
            return False, f"错误：处理数据失败：{e}", None, None

    def browse_file(self):
        """Handle file selection via dialog."""
        try:
            file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
            if file_path:
                self.file_entry.delete(0, tk.END)
                self.file_entry.insert(0, file_path)
                self.status_label.config(text="文件已选择，请输入批量大小或点击提交。", foreground="#333333")
                self.debug_info.append({
                    '时间戳': self.current_time,
                    '类别': '信息',
                    '信息': f"用户通过文件对话框选择文件路径：{file_path}"
                })
        except Exception as e:
            self.status_label.config(text=f"错误：无法选择文件：{e}", foreground="#dc3545")
            self.debug_info.append({
                '时间戳': self.current_time,
                '类别': '错误',
                '信息': f"选择文件时出错：{e}\n{traceback.format_exc()}"
            })

    def submit(self):
        """Handle submit button action."""
        try:
            self.submit_button.config(state='disabled')
            self.file_entry.config(state='disabled')
            self.batch_entry.config(state='disabled')
            self.open_button.grid_remove()
            self.status_label.config(text="正在处理数据，请稍候...", foreground="#333333")
            self.progress_bar.grid()
            self.progress_bar.start()
            self.root.update()

            file_path = self.file_entry.get().strip()
            batch_input = self.batch_entry.get().strip()

            if not file_path:
                raise ValueError("文件路径不能为空。")
            if not file_path.lower().endswith('.xlsx'):
                raise ValueError("文件必须是 .xlsx 格式。")
            valid, _ = self.validate_file(file_path, "Test Data")
            if not valid:
                raise FileNotFoundError(f"文件 {file_path} 不存在或无法访问。")

            user_batch_size = None
            if batch_input:
                try:
                    user_batch_size = float(batch_input)
                    if user_batch_size <= 0:
                        raise ValueError("批量大小必须是正数。")
                except ValueError as ve:
                    raise ValueError(f"批量大小必须是数字：{ve}")

            success, result, self.df_test, self.tool_to_hours_base = self.process_data(file_path, user_batch_size)
            if success:
                unmatched_tools_df = self.df_test[~self.df_test['量具1层编码'].isin(self.tool_to_hours_base.keys())][
                    ['产品编码', '检验卡编号', '工序名称', '量具1层编码']]
                unmatched_count = len(unmatched_tools_df)
                zero_hours_count = len(self.df_test[self.df_test['工时'] == 0])
                self.status_label.config(
                    text=f"处理完成，结果已保存至 {result}。未匹配量具数量：{unmatched_count}，工时为0的记录数：{zero_hours_count}。点击“打开输出文件”查看。",
                    foreground="#28a745")
                self.output_file_path = result
                self.open_button.grid()
            else:
                raise RuntimeError(result)

        except Exception as e:
            self.status_label.config(text=f"错误：{e}", foreground="#dc3545")
            self.debug_info.append({
                '时间戳': self.current_time,
                '类别': '错误',
                '信息': f"提交处理时出错：{e}\n{traceback.format_exc()}"
            })
            messagebox.showerror("错误", str(e))
        finally:
            self.progress_bar.stop()
            self.progress_bar.grid_remove()
            self.submit_button.config(state='normal')
            self.file_entry.config(state='normal')
            self.batch_entry.config(state='normal')
            self.root.update()

    def open_output(self):
        """Handle open output file button action."""
        try:
            if not self.output_file_path:
                raise ValueError("未生成输出文件。")
            success, message = self.open_output_file(self.output_file_path)
            if not success:
                raise RuntimeError(message)
        except Exception as e:
            self.status_label.config(text=f"错误：{e}", foreground="#dc3545")
            self.debug_info.append({
                '时间戳': self.current_time,
                '类别': '错误',
                '信息': f"打开输出文件时出错：{e}\n{traceback.format_exc()}"
            })
            messagebox.showerror("错误", str(e))

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = BatchLaborTimeProcessor(root)
        root.mainloop()
    except Exception as e:
        print(f"错误：启动程序失败：{e}")
        pd.DataFrame(app.debug_info).to_excel('debug_log.xlsx', index=False)
        exit(1)