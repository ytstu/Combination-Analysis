import math
from datetime import datetime
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_DATA_DIR = BASE_DIR / "data" / "input"
DEFAULT_OUTPUT_DIR = BASE_DIR / "data" / "output"
DEFAULT_INPUT_FILE_PATH = DEFAULT_DATA_DIR / "模拟-i输入-讲解.xlsx"
DEFAULT_PRODUCT_DB_PATH = DEFAULT_DATA_DIR / "商品资料.xlsx"
DEFAULT_COMBO_DB_PATH = DEFAULT_DATA_DIR / "组合资料.xlsx"
PROCESS_COLUMNS = [
    "单品是否存在",
    "组合是否存在",
    "倍数前",
    "倍数后",
    "组合商品名称",
]


class ExcelDataService:

    def __init__(self):
        self.product_db_path = DEFAULT_PRODUCT_DB_PATH
        self.combo_db_path = DEFAULT_COMBO_DB_PATH
        self.product_df = None
        self.combo_df = None

    @staticmethod
    def _string_or_empty(value):
        return str(value) if pd.notna(value) else ""

    @staticmethod
    def _parse_multiplier_codes(code_series):
        source_codes = code_series.astype(str)
        code_parts = source_codes.str.extract(r"^(.*)\*(\d+)$")
        has_multiplier = code_parts[1].notna()
        return (
            source_codes,
            has_multiplier,
            code_parts[0].fillna(""),
            code_parts[1].fillna(""),
        )

    def _load_database(self, file_path):
        if not file_path.exists():
            return None
        return pd.read_excel(file_path)

    def load_databases(self):
        self.product_df = self._load_database(self.product_db_path)
        self.combo_df = self._load_database(self.combo_db_path)

    def load_input_file(self, file_path):
        return pd.read_excel(file_path)

    def process_data(self, input_df):
        df = input_df.copy().drop_duplicates(subset=["原始商品编码"])
        source_codes, has_multiplier, base_codes, multiplier_codes = (
            self._parse_multiplier_codes(df["原始商品编码"])
        )
        df = df[has_multiplier].copy()

        for column in PROCESS_COLUMNS:
            df[column] = ""

        source_codes = source_codes[has_multiplier]
        product_lookup = self.product_df.copy()
        product_lookup["商品编码"] = product_lookup["商品编码"].astype(str)
        product_lookup = product_lookup.drop_duplicates(
            subset=["商品编码"],
            keep="first",
        ).set_index("商品编码")

        product_codes = product_lookup.index
        combo_codes = self.combo_df["组合商品编码"].astype(str)

        product_exists = source_codes.isin(product_codes)
        combo_exists = source_codes.isin(combo_codes)

        df["单品是否存在"] = product_exists.map({True: "存在", False: ""})
        df["组合是否存在"] = combo_exists.map({True: "存在", False: ""})

        missing_mask = ~product_exists & ~combo_exists
        df = df[missing_mask].copy()
        if df.empty:
            return df

        df["倍数前"] = base_codes.loc[df.index]
        df["倍数后"] = multiplier_codes.loc[df.index]

        base_codes = df["倍数前"]
        has_product = base_codes.isin(product_codes)
        has_quantity_column = "数量(pcs)" in product_lookup.columns

        def get_product_text(column_name):
            if column_name not in product_lookup.columns:
                return pd.Series("", index=df.index, dtype="object")

            values = base_codes.map(product_lookup[column_name])
            return values.where(values.notna(), "").astype(str)

        size_values = get_product_text("尺寸规格(mm)")
        color_values = get_product_text("颜色")
        if has_quantity_column:
            quantity_values = base_codes.map(product_lookup["数量(pcs)"])
        else:
            quantity_values = pd.Series(None, index=df.index, dtype="object")

        def calculate_pcs_value(found_product, quantity, multiple):
            if not found_product or not has_quantity_column:
                return ""

            pcs_str = self._string_or_empty(quantity)
            multiple_str = self._string_or_empty(multiple)
            try:
                pcs_value = float(pcs_str) if pcs_str else 0
                multiple_value = float(multiple_str) if multiple_str else 0
            except (TypeError, ValueError):
                return ""

            total_pcs = pcs_value * multiple_value
            if not math.isfinite(total_pcs):
                return None
            return str(int(total_pcs))

        pcs_values = [
            calculate_pcs_value(found_product, quantity, multiple)
            for found_product, quantity, multiple in zip(
                has_product.tolist(),
                quantity_values.tolist(),
                df["倍数后"].tolist(),
            )
        ]
        df["组合商品名称"] = [
            "" if pcs is None else f"{size}{pcs}{color}"
            for size, pcs, color in zip(
                size_values.tolist(),
                pcs_values,
                color_values.tolist(),
            )
        ]
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


def resolve_output_path():
    return DEFAULT_OUTPUT_DIR / f"组合装数据{datetime.now().strftime('%m%d')}.xlsx"


def run():
    input_path = DEFAULT_INPUT_FILE_PATH
    service = ExcelDataService()

    service.load_databases()
    input_df = service.load_input_file(input_path)
    processed_df = service.process_data(input_df)
    export_df = service.build_export_df(processed_df)
    output_path = resolve_output_path()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    export_df.to_excel(output_path, index=False)


def main():
    run()


if __name__ == "__main__":
    main()
