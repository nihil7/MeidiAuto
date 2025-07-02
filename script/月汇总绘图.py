# -*- coding: utf-8 -*-
"""
【在折线上直接显示“编号: 数量”】
自动生成每个编号月度入库量折线图（含平均线 MarkLine）
"""

import pandas as pd
from pyecharts import options as opts
from pyecharts.charts import Line
from pyecharts.commons.utils import JsCode
import webbrowser
import os

# ========== 配置区 ==========
EXCEL_PATH = r"C:\Users\ishel\Desktop\美的发货\月度汇总\月汇总表2502-2506.xlsx"
OUTPUT_HTML = r"C:\Users\ishel\Desktop\美的发货\月度汇总\月度入库量折线图_显示编号和数量.html"
# ============================

def generate_line_chart():
    # 读取数据
    df = pd.read_excel(EXCEL_PATH)

    # 提取月份列和月份列表
    month_cols = [col for col in df.columns if "外仓入库总量" in col]
    months = [col.replace("外仓入库总量", "") for col in month_cols]

    # 创建图表对象
    line = (
        Line(init_opts=opts.InitOpts(width="1400px", height="2400px"))
        .add_xaxis(months)
        .set_global_opts(
            title_opts=opts.TitleOpts(title="折线图"),
            tooltip_opts=opts.TooltipOpts(
                trigger="item",
                formatter=JsCode(
                    "function(params){return params.seriesName + ': ' + params.value.toLocaleString();}"
                )
            ),
            toolbox_opts=opts.ToolboxOpts(),
            datazoom_opts=[
                opts.DataZoomOpts(type_="inside"),
                opts.DataZoomOpts(type_="slider"),
            ],
            yaxis_opts=opts.AxisOpts(
                type_="value",
                axislabel_opts=opts.LabelOpts(formatter="{value}"),
            ),
            xaxis_opts=opts.AxisOpts(
                type_="category",
                boundary_gap=False,
                axislabel_opts=opts.LabelOpts(rotate=0),
            ),
        )
    )

    # 为每个编号添加折线
    for idx, row in df.iterrows():
        code = str(row["编号"])
        qtys = [int(row[col]) if pd.notna(row[col]) else 0 for col in month_cols]

        # 如果全为 0 则不绘制
        if sum(qtys) == 0:
            continue

        line.add_yaxis(
            series_name=code,
            y_axis=qtys,
            is_symbol_show=True,
            label_opts=opts.LabelOpts(
                is_show=True,
                formatter=JsCode(
                    "function(params){return params.seriesName + ': ' + params.value.toLocaleString();}"
                ),
                position="top",
            ),
            markline_opts=opts.MarkLineOpts(
                data=[opts.MarkLineItem(type_="average", name="平均值")]
            ),
            linestyle_opts=opts.LineStyleOpts(width=2),
        )

    # 输出并自动打开浏览器预览
    line.render(OUTPUT_HTML)
    print(f"✅ 已生成交互式折线图（在折线上显示编号和数量）：{OUTPUT_HTML}")
    webbrowser.open("file://" + os.path.realpath(OUTPUT_HTML))

if __name__ == "__main__":
    generate_line_chart()
