"""
Google Play 评论 CSV → xlsx 预处理脚本

功能：读取 Google Play 导出的评论 CSV，完成清洗、格式转换、拆分，
输出带机翻公式的 xlsx 文件（二级分类列留空，由 Claude 或 API 填充）。

用法：
    python review_csv_to_xlsx.py <input.csv> [output.xlsx]

依赖：
    pip install pandas openpyxl
"""

import sys
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from pathlib import Path


# ── 配置 ──────────────────────────────────────────────

SOURCE_COLUMNS = [
    "Review Text",
    "Star Rating",
    "Reviewer Language",
    "Review Submit Date and Time",
    "Review Link",
]

OUTPUT_HEADERS = [
    "二级分类",
    "机翻 (Translation)",
    "原文 (Review Text)",
    "评分 (Star Rating)",
    "语言代码 (Language)",
    "发布时间 (Submit Date)",
    "Review Link",
]

COLUMN_WIDTHS = [30, 45, 45, 12, 18, 18, 80]

HEADER_FILL = PatternFill("solid", fgColor="4472C4")
HEADER_FONT = Font(bold=True, color="FFFFFF", name="Arial", size=11)
BODY_FONT = Font(name="Arial", size=11)

FIVE_STAR_SHORT_THRESHOLD = 50  # 原文 < 此字符数的 5星评论自动标为纯好评


# ── Step 1: 读取 CSV ─────────────────────────────────

def read_csv(path: str) -> pd.DataFrame:
    """尝试 UTF-16，失败则回退 UTF-8。"""
    for enc in ("utf-16", "utf-8", "utf-8-sig", "cp1252"):
        try:
            df = pd.read_csv(path, encoding=enc)
            if all(col in df.columns for col in SOURCE_COLUMNS):
                return df[SOURCE_COLUMNS]
            missing = [c for c in SOURCE_COLUMNS if c not in df.columns]
            raise ValueError(f"CSV 缺少必需列: {missing}\n现有列: {list(df.columns)}")
        except (UnicodeDecodeError, UnicodeError):
            continue
    raise ValueError(f"无法以 UTF-16/UTF-8 解码文件: {path}")


# ── Step 2: 数据清洗 ─────────────────────────────────

def clean(df: pd.DataFrame) -> pd.DataFrame:
    df = df[df["Review Text"].notna() & (df["Review Text"].str.strip() != "")]
    df = df.copy()
    df["Review Submit Date and Time"] = (
        pd.to_datetime(df["Review Submit Date and Time"], errors="coerce")
        .dt.strftime("%Y-%m-%d")
    )
    return df


# ── Step 3: 拆分 ─────────────────────────────────────

def split_by_rating(df: pd.DataFrame):
    five_star = df[df["Star Rating"] == 5].reset_index(drop=True)
    low_star = df[df["Star Rating"] < 5].reset_index(drop=True)
    five_star_long = five_star[
        five_star["Review Text"].str.len() >= FIVE_STAR_SHORT_THRESHOLD
    ].reset_index(drop=True)
    return five_star, five_star_long, low_star


# ── Step 4 + 6: 写入 xlsx ────────────────────────────

def write_sheet(ws, df: pd.DataFrame, is_five_star: bool = False):
    """向 worksheet 写入表头 + 数据行 + 机翻公式 + 样式。"""
    # 表头
    for col_idx, header in enumerate(OUTPUT_HEADERS, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # 数据行
    for row_idx, (_, row) in enumerate(df.iterrows(), 2):
        # A: 二级分类
        if is_five_star and len(str(row["Review Text"]).strip()) < FIVE_STAR_SHORT_THRESHOLD:
            ws.cell(row=row_idx, column=1, value="2-01 5星纯好评")
        else:
            ws.cell(row=row_idx, column=1, value="")
        # B: 机翻公式
        formula = f'=IF(C{row_idx}="","",GOOGLETRANSLATE(C{row_idx},E{row_idx},"zh-CN"))'
        ws.cell(row=row_idx, column=2, value=formula)
        # C: 原文
        ws.cell(row=row_idx, column=3, value=row["Review Text"])
        # D: 评分
        ws.cell(row=row_idx, column=4, value=row["Star Rating"])
        # E: 语言代码
        ws.cell(row=row_idx, column=5, value=row["Reviewer Language"])
        # F: 发布时间
        ws.cell(row=row_idx, column=6, value=row["Review Submit Date and Time"])
        # G: Review Link
        ws.cell(row=row_idx, column=7, value=row["Review Link"])

    # 样式：列宽
    for col_idx, width in enumerate(COLUMN_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # 字体
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(OUTPUT_HEADERS)):
        for cell in row:
            cell.font = BODY_FONT

    # 冻结首行 + 自动筛选
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = f"A1:G{ws.max_row}"


def build_xlsx(five_star, five_star_long, low_star, output_path: str):
    wb = Workbook()

    # Sheet 1: 1-4星评论
    ws_low = wb.active
    ws_low.title = "1-4星评论"
    write_sheet(ws_low, low_star)

    # Sheet 2: 5星评论（全部）
    ws_five = wb.create_sheet("5星评论")
    write_sheet(ws_five, five_star, is_five_star=True)

    # Sheet 3: 5星长评（≥50字符，A列留空）
    ws_five_long = wb.create_sheet("5星长评")
    write_sheet(ws_five_long, five_star_long)

    wb.save(output_path)


# ── 主流程 ────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print("用法: python review_csv_to_xlsx.py <input.csv> [output.xlsx]")
        sys.exit(1)

    input_path = sys.argv[1]
    if len(sys.argv) >= 3:
        output_path = sys.argv[2]
    else:
        output_path = str(Path(input_path).with_suffix(".xlsx"))

    print(f"读取: {input_path}")
    df = read_csv(input_path)
    print(f"  原始行数: {len(df)}")

    df = clean(df)
    print(f"  清洗后行数: {len(df)}")

    five_star, five_star_long, low_star = split_by_rating(df)
    five_star_short = len(five_star) - len(five_star_long)
    print(f"  5星: {len(five_star)} 条 (短评 {five_star_short}, 长评 {len(five_star_long)}), 1-4星: {len(low_star)} 条")

    build_xlsx(five_star, five_star_long, low_star, output_path)
    print(f"输出: {output_path}")
    print(f"  ✓ Sheet 1: 1-4星评论 ({len(low_star)} 条)")
    print(f"  ✓ Sheet 2: 5星评论 ({len(five_star)} 条，短评已标为 2-01)")
    print(f"  ✓ Sheet 3: 5星长评 ({len(five_star_long)} 条，等待分类)")
    print("  ✓ 机翻公式已写入，上传 Google Sheets 后生效")


if __name__ == "__main__":
    main()
