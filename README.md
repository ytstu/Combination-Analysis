# Combination-Analysis
组合装拆解分析

## 运行方式

1. 安装 Python 3.9+。
2. 安装依赖：
```bash
pip install pandas openpyxl
```
3. 启动程序：
```bash
python3 main.py -i "待处理文件.xlsx"
```

## 使用流程

1. 准备待处理的 Excel，确认包含 `原始商品编码` 列。
2. 执行命令进行处理并导出：
```bash
python3 main.py -i "input.xlsx" -o "output.xlsx"
```
3. 不传 `-o` 时，会在输入文件同目录自动生成 `组合装数据MMDD.xlsx`。

## 参数说明

```bash
python3 main.py -h
```

- `-i, --input`：待处理 Excel 文件路径（必填）
- `-o, --output`：导出文件路径（可选）
- `--product-db`：商品资料库路径（可选）
- `--combo-db`：组合资料库路径（可选）

## 数据库路径说明

默认数据库路径定义在 `main.py` 的 `ExcelDataService` 中。  
如果运行环境不同，建议通过命令行参数覆盖：

```bash
python3 main.py -i "input.xlsx" --product-db "商品资料.xlsx" --combo-db "组合资料.xlsx"
```
