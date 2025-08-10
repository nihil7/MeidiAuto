import os
import re
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.chart import LineChart, Reference, Series

# ======================== 配置区域 ========================
folder_path = r'C:\Users\ishel\Desktop\美的发货\月度汇总'  # 替换为你的文件夹路径
start_month = 2502  # 起始月份
end_month = 2507    # 结束月份
output_filename = f'月汇总表{start_month}-{end_month}.xlsx'
# ==========================================================

def extract_month(filename):
    match = re.search(r'(\d{4})月底', filename)
    if match:
        return int(match.group(1))
    return None

def read_monthly_data(file_path, month_str):
    wb = load_workbook(file_path, data_only=True)
    if '库存表' not in wb.sheetnames:
        print(f"⚠️ 文件 {os.path.basename(file_path)} 未找到 '库存表'，跳过")
        return None

    ws = wb['库存表']
    headers = [cell.value for cell in ws[4]]

    if '编号' not in headers or '外仓入库总量' not in headers:
        print(f"⚠️ 文件 {os.path.basename(file_path)} 缺少必要列，跳过")
        return None

    idx_code = headers.index('编号') + 1
    idx_in_qty = headers.index('外仓入库总量') + 1

    records = []
    for row in ws.iter_rows(min_row=5, values_only=True):
        code = row[idx_code - 1]
        in_qty = row[idx_in_qty - 1]
        if code is not None:
            records.append({'编号': code, f'{month_str}外仓入库总量': in_qty})

    return pd.DataFrame(records) if records else None

def create_excel_chart(excel_path):
    wb = load_workbook(excel_path)
    ws = wb.active

    # 获取列数和行数
    max_col = ws.max_column
    max_row = ws.max_row

    # 自动创建折线图
    chart = LineChart()
    chart.title = "每个编号的月度外仓入库量趋势"
    chart.style = 13
    chart.y_axis.title = '外仓入库总量'
    chart.x_axis.title = '月份'

    # 横轴标签（月份）
    months = [ws.cell(row=1, column=col).value for col in range(2, max_col + 1)]
    chart.x_axis.number_format = 'General'
    chart.x_axis.majorTickMark = "out"

    # 为每个“编号”添加一条线
    for row in range(2, max_row + 1):
        series = Series(
            Reference(ws, min_col=2, max_col=max_col, min_row=row, max_row=row),
            title=str(ws.cell(row=row, column=1).value)
        )
        chart.series.append(series)

    # 设置横轴分类标签（月份）
    cats = Reference(ws, min_col=2, max_col=max_col, min_row=1)
    chart.set_categories(cats)

    # 插入图表到表中位置（可调整）
    ws.add_chart(chart, f"B{max_row + 3}")

    wb.save(excel_path)
    print(f"✅ 已在 {excel_path} 中插入 Excel 原生折线图")

def main():
    monthly_dfs = []
    processed_months = []

    for file in os.listdir(folder_path):
        if file.endswith('.xlsx'):
            month = extract_month(file)
            if month and start_month <= month <= end_month:
                if month in processed_months:
                    continue
                processed_months.append(month)

                file_path = os.path.join(folder_path, file)
                df = read_monthly_data(file_path, str(month))
                if df is not None:
                    monthly_dfs.append(df)

    if not monthly_dfs:
        print("⚠️ 未获取到任何可用数据，未生成汇总文件")
        return

    result_df = monthly_dfs[0]
    for df in monthly_dfs[1:]:
        result_df = pd.merge(result_df, df, on='编号', how='outer')
    result_df = result_df.fillna(0)

    output_path = os.path.join(folder_path, output_filename)
    result_df.to_excel(output_path, index=False)
    print(f"✅ 已生成汇总文件: {output_path}")
    print(f"📊 汇总表行数: {len(result_df)}，列数: {len(result_df.columns)}")

    create_excel_chart(output_path)

if __name__ == "__main__":
    main()
