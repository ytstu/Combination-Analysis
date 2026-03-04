import argparse
import math
import os
import re
from datetime import datetime
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_DATA_DIR = BASE_DIR / "data" / "input"
DEFAULT_OUTPUT_DIR = BASE_DIR / "data" / "output"
DEFAULT_PRODUCT_DB_PATH = DEFAULT_DATA_DIR / "商品资料.xlsx"
DEFAULT_COMBO_DB_PATH = DEFAULT_DATA_DIR / "组合资料.xlsx"
PROCESS_COLUMNS = [
    "单品是否存在",
    "组合是否存在",
    "倍数前",
    "倍数后",
    "组合商品名称",
    "临时名称",
    "提取后",
]


class ExcelDataService:

    def __init__(self):
        self.product_db_path = DEFAULT_PRODUCT_DB_PATH
        self.combo_db_path = DEFAULT_COMBO_DB_PATH
        self.product_df = None
        self.combo_df = None

    @staticmethod
    def _get_first_row(value):
        if isinstance(value, pd.DataFrame):
            return value.iloc[0]
        return value

    @staticmethod
    def _string_or_empty(value):
        return str(value) if pd.notna(value) else ""

    def _load_database(self, file_path, label):
        result = {"loaded": False, "message": f"{label}: 未加载", "count": 0}
        if not os.path.exists(file_path):
            result["message"] = f"{label}: 文件不存在"
            return None, result

        dataframe = pd.read_excel(file_path)
        count = len(dataframe)
        result["loaded"] = True
        result["message"] = f"{label}: 已加载 ({count} 条记录)"
        result["count"] = count
        return dataframe, result

    def load_databases(self):
        self.product_df, product_result = self._load_database(
            self.product_db_path,
            "商品资料库",
        )
        self.combo_df, combo_result = self._load_database(
            self.combo_db_path,
            "组合资料库",
        )
        return {
            "product": product_result,
            "combo": combo_result,
            "all_loaded": bool(product_result["loaded"] and combo_result["loaded"]),
        }

    def load_input_file(self, file_path):
        df = pd.read_excel(file_path)
        if "原始商品编码" not in df.columns:
            raise ValueError("文件中没有找到'原始商品编码'字段")
        return df

    def process_data(self, input_df):
        if input_df is None:
            raise ValueError("请先导入数据")
        if self.product_df is None or self.combo_df is None:
            raise ValueError("数据库未正确加载")

        df = input_df.copy().drop_duplicates(subset=["原始商品编码"])
        source_codes = df["原始商品编码"].astype(str)
        df = df[source_codes.str.contains("*", regex=False, na=False)].copy()

        for column in PROCESS_COLUMNS:
            df[column] = ""

        source_codes = df["原始商品编码"].astype(str)
        product_codes = set(self.product_df["商品编码"].astype(str))
        combo_codes = set(self.combo_df["组合商品编码"].astype(str))

        df["单品是否存在"] = source_codes.apply(
            lambda code: "存在" if code in product_codes else ""
        )
        df["组合是否存在"] = source_codes.apply(
            lambda code: "存在" if code in combo_codes else ""
        )

        missing_mask = (df["单品是否存在"] != "存在") & (df["组合是否存在"] != "存在")
        df = df[missing_mask].copy()

        code_parts = (
            df["原始商品编码"].astype(str).apply(lambda code: code.rsplit("*", 1))
        )
        df["倍数前"] = code_parts.str[0]
        df["倍数后"] = code_parts.str[1].fillna("")

        product_index = self.product_df.copy()
        product_index["商品编码"] = product_index["商品编码"].astype(str)
        product_index = product_index.set_index("商品编码", drop=False)

        def get_product_item(code):
            if not code or code not in product_index.index:
                return None
            return self._get_first_row(product_index.loc[code])

        def get_item_text(item, column):
            if item is None or column not in item.index:
                return ""
            return self._string_or_empty(item[column])

        def find_product_name(code):
            return get_item_text(get_product_item(code), "商品名称")

        df["临时名称"] = df["倍数前"].apply(find_product_name)

        def extract_chinese(text):
            if not isinstance(text, str):
                return ""
            match = re.search(r"[^\d\u0030-\u0039]*$", text)
            if match:
                return match.group()
            return text

        df["提取后"] = df["临时名称"].apply(extract_chinese)

        def calculate_pcs(item, multiple):
            if item is None or "数量(pcs)" not in item.index:
                return ""

            pcs_str = self._string_or_empty(item["数量(pcs)"])
            try:
                pcs_value = float(pcs_str) if pcs_str else 0
                multiple_value = float(multiple) if multiple else 0
            except (TypeError, ValueError):
                return ""

            total_pcs = pcs_value * multiple_value
            if not math.isfinite(total_pcs):
                return None
            return str(int(total_pcs))

        def create_combo_name(row):
            item = get_product_item(str(row["倍数前"]))
            temp_name = self._string_or_empty(row["临时名称"])
            size = get_item_text(item, "尺寸规格(mm)")
            pcs = calculate_pcs(item, row["倍数后"])
            if pcs is None:
                return ""
            color = get_item_text(item, "颜色")
            return f"{temp_name}{size}{pcs}{color}"

        df["组合商品名称"] = df.apply(create_combo_name, axis=1)
        return df

    @staticmethod
    def build_export_df(processed_df):
        return pd.DataFrame(
            {
                "组合商品编码": processed_df["原始商品编码"],
                "组合商品名称": processed_df["组合商品名称"],
                "商品编码": processed_df["倍数前"],
                "数量": processed_df["倍数后"],
            }
        )


def parse_args():
    parser = argparse.ArgumentParser(description="组合装拆解分析后端程序")
    parser.add_argument("-i", "--input", required=True, help="待处理的 Excel 文件路径")
    return parser.parse_args()


def resolve_project_path(path_value):
    path = Path(path_value)
    if path.is_absolute():
        return path
    return BASE_DIR / path


def resolve_output_path():
    return DEFAULT_OUTPUT_DIR / f"组合装数据{datetime.now().strftime('%m%d')}.xlsx"


def run():
    args = parse_args()
    input_path = resolve_project_path(args.input)
    if not input_path.exists():
        raise FileNotFoundError(f"输入文件不存在: {input_path}")

    service = ExcelDataService()

    db_status = service.load_databases()
    if not db_status["all_loaded"]:
        raise RuntimeError("数据库未完整加载，请检查路径后重试")

    input_df = service.load_input_file(str(input_path))
    processed_df = service.process_data(input_df)
    export_df = service.build_export_df(processed_df)
    output_path = resolve_output_path()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    export_df.to_excel(output_path, index=False)


def main():
    run()


if __name__ == "__main__":
    main()
