"""Build the Excel report. Every summary number is an Excel formula pointing at the
clean data sheet, so the owner can edit or add a sale and the report updates itself."""
from __future__ import annotations

from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from .catalog import CATALOG

FONT = "Arial"
MARE, CORAL, AREIA = "0F4C5C", "E07A5F", "F2E8CF"
BRL = '"R$" #,##0.00'
PCT = "0.0%"
DATE = "dd/mm/yyyy"
H1 = Font(name=FONT, size=16, bold=True, color=MARE)
H2 = Font(name=FONT, size=11, bold=True, color=MARE)
HEAD = Font(name=FONT, size=10, bold=True, color="FFFFFF")
BODY = Font(name=FONT, size=10)
NOTE = Font(name=FONT, size=9, italic=True, color="6B5A52")
KPI = Font(name=FONT, size=14, bold=True, color=MARE)
HEAD_FILL = PatternFill("solid", fgColor=MARE)
KPI_FILL = PatternFill("solid", fgColor=AREIA)
THIN = Side(style="thin", color="E5D7B8")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _header(ws, row, headers, col=1):
    for i, h in enumerate(headers):
        c = ws.cell(row=row, column=col + i, value=h)
        c.font, c.fill, c.alignment = HEAD, HEAD_FILL, Alignment(horizontal="center", vertical="center", wrap_text=True)


def _widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _style_chart(chart, title, x_title=None, y_title=None):
    chart.title = title
    chart.roundedCorners = False
    chart.legend = None
    if x_title:
        chart.x_axis.title = x_title
    if y_title:
        chart.y_axis.title = y_title
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.height, chart.width = 8, 17


def build_report(df: pd.DataFrame, log: pd.DataFrame, path: str) -> dict:
    wb = Workbook()
    resumo = wb.active
    resumo.title = "Resumo"
    ws_prod = wb.create_sheet("Por produto")
    ws_day = wb.create_sheet("Por dia")
    ws_hour = wb.create_sheet("Por hora")
    ws_data = wb.create_sheet("Vendas")
    ws_log = wb.create_sheet("Correções")

    # ---------------- Vendas (clean data) ----------------
    cols = list(df.columns)
    _header(ws_data, 1, cols)
    for r, rec in enumerate(df.itertuples(index=False), start=2):
        for c, v in enumerate(rec, start=1):
            cell = ws_data.cell(row=r, column=c, value=v)
            cell.font = BODY
    n = len(df) + 1  # last data row
    fmt = {"Data": DATE, "Data e hora": "dd/mm/yyyy hh:mm", "Preço unitário": BRL, "Total": BRL}
    for name, f in fmt.items():
        col = cols.index(name) + 1
        for r in range(2, n + 1):
            ws_data.cell(row=r, column=col).number_format = f
    table = Table(displayName="Vendas", ref=f"A1:{get_column_letter(len(cols))}{n}")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
    ws_data.add_table(table)
    ws_data.freeze_panes = "A2"
    _widths(ws_data, [12, 7, 13, 17, 26, 11, 11, 14, 12, 15, 10, 22])

    L = {name: get_column_letter(cols.index(name) + 1) for name in cols}
    rng = lambda name: f"Vendas!${L[name]}$2:${L[name]}${n}"  # noqa: E731
    TOTAL, QTY, PROD, DAYC, HOURC, CANAL, PAG = (rng("Total"), rng("Quantidade"), rng("Produto"),
                                                 rng("Data"), rng("Hora"), rng("Canal"), rng("Pagamento"))

    # ---------------- Por produto ----------------
    ws_prod["A1"], ws_prod["A1"].font = "Vendas por produto", H1
    _header(ws_prod, 3, ["Produto", "Categoria", "Itens vendidos", "Faturamento", "% do faturamento"])
    order = df.groupby("Produto")["Total"].sum().sort_values(ascending=False).index.tolist()
    first, last = 4, 3 + len(order)
    for i, prod in enumerate(order, start=first):
        ws_prod.cell(row=i, column=1, value=prod)
        ws_prod.cell(row=i, column=2, value=CATALOG.get(prod, ("Outros", 0))[0])
        ws_prod.cell(row=i, column=3, value=f'=SUMIFS({QTY},{PROD},A{i})')
        ws_prod.cell(row=i, column=4, value=f'=SUMIFS({TOTAL},{PROD},A{i})').number_format = BRL
        ws_prod.cell(row=i, column=5, value=f"=D{i}/SUM($D${first}:$D${last})").number_format = PCT
        for c in range(1, 6):
            ws_prod.cell(row=i, column=c).font = BODY
            ws_prod.cell(row=i, column=c).border = BOX
    _widths(ws_prod, [28, 12, 15, 16, 17])
    chart = BarChart()
    chart.type = "bar"
    chart.add_data(Reference(ws_prod, min_col=4, min_row=3, max_row=last), titles_from_data=True)
    chart.set_categories(Reference(ws_prod, min_col=1, min_row=first, max_row=last))
    _style_chart(chart, "Faturamento por produto")
    chart.y_axis.numFmt = '"R$" #,##0'
    chart.x_axis.scaling.orientation = "maxMin"  # biggest on top
    chart.series[0].graphicalProperties.solidFill = MARE
    chart.height = 10
    ws_prod.add_chart(chart, "G3")

    # ---------------- Por dia ----------------
    ws_day["A1"], ws_day["A1"].font = "Vendas por dia", H1
    _header(ws_day, 3, ["Data", "Dia da semana", "Faturamento", "Itens vendidos"])
    days = sorted(df["Data"].unique())
    dnames = df.drop_duplicates("Data").set_index("Data")["Dia da semana"]
    d_first, d_last = 4, 3 + len(days)
    for i, d in enumerate(days, start=d_first):
        ws_day.cell(row=i, column=1, value=d).number_format = DATE
        ws_day.cell(row=i, column=2, value=dnames[d])
        ws_day.cell(row=i, column=3, value=f"=SUMIFS({TOTAL},{DAYC},A{i})").number_format = BRL
        ws_day.cell(row=i, column=4, value=f"=SUMIFS({QTY},{DAYC},A{i})")
        for c in range(1, 5):
            ws_day.cell(row=i, column=c).font = BODY
            ws_day.cell(row=i, column=c).border = BOX
    _widths(ws_day, [13, 15, 15, 15])
    line = LineChart()
    line.add_data(Reference(ws_day, min_col=3, min_row=3, max_row=d_last), titles_from_data=True)
    line.set_categories(Reference(ws_day, min_col=1, min_row=d_first, max_row=d_last))
    _style_chart(line, "Faturamento diário")
    line.y_axis.numFmt = '"R$" #,##0'
    line.x_axis.number_format = "dd/mm"
    line.series[0].graphicalProperties.line.solidFill = CORAL
    line.series[0].graphicalProperties.line.width = 28000
    line.series[0].smooth = False  # straight segments: smoothing would invent values between days
    ws_day.add_chart(line, "F3")

    # ---------------- Por hora ----------------
    ws_hour["A1"], ws_hour["A1"].font = "Vendas por horário", H1
    _header(ws_hour, 3, ["Hora", "Faturamento", "Itens vendidos", "Média por dia aberto"])
    hours = sorted(df["Hora"].unique())
    h_first, h_last = 4, 3 + len(hours)
    for i, h in enumerate(hours, start=h_first):
        ws_hour.cell(row=i, column=1, value=int(h))
        ws_hour.cell(row=i, column=2, value=f"=SUMIFS({TOTAL},{HOURC},A{i})").number_format = BRL
        ws_hour.cell(row=i, column=3, value=f"=SUMIFS({QTY},{HOURC},A{i})")
        ws_hour.cell(row=i, column=4, value=f"=B{i}/SUMPRODUCT(({HOURC}=A{i})/COUNTIFS({DAYC},{DAYC},{HOURC},{HOURC}))").number_format = BRL
        for c in range(1, 5):
            ws_hour.cell(row=i, column=c).font = BODY
            ws_hour.cell(row=i, column=c).border = BOX
    ws_hour.cell(row=h_last + 2, column=1, value="Hora = início da faixa (7 = das 7h às 7h59). "
                 "Média por dia aberto = faturamento da faixa ÷ dias em que houve venda nessa faixa.").font = NOTE
    _widths(ws_hour, [9, 15, 15, 20])
    bar = BarChart()
    bar.add_data(Reference(ws_hour, min_col=2, min_row=3, max_row=h_last), titles_from_data=True)
    bar.set_categories(Reference(ws_hour, min_col=1, min_row=h_first, max_row=h_last))
    _style_chart(bar, "Faturamento por horário", x_title="Hora")
    bar.y_axis.numFmt = '"R$" #,##0'
    bar.series[0].graphicalProperties.solidFill = MARE
    ws_hour.add_chart(bar, "F3")

    # ---------------- Correções ----------------
    lcols = list(log.columns)
    _header(ws_log, 1, lcols)
    for r, rec in enumerate(log.itertuples(index=False), start=2):
        for c, v in enumerate(rec, start=1):
            ws_log.cell(row=r, column=c, value=v).font = BODY
    ws_log.freeze_panes = "A2"
    ws_log.auto_filter.ref = f"A1:{get_column_letter(len(lcols))}{max(len(log) + 1, 2)}"
    _widths(ws_log, [22, 8, 15, 32, 28, 42])

    # ---------------- Resumo ----------------
    start, end = df["Data"].min(), df["Data"].max()
    resumo["A1"], resumo["A1"].font = "Relatório de vendas · Quebra-Mar Café & Padaria", H1
    resumo["A2"] = f"Período: {start:%d/%m/%Y} a {end:%d/%m/%Y} · gerado em {datetime.now():%d/%m/%Y %H:%M}"
    resumo["A2"].font = NOTE
    kpis = [
        ("Faturamento líquido", f"=SUM({TOTAL})", BRL),
        ("Itens vendidos", f"=SUM({QTY})", "#,##0"),
        ("Preço médio por item", "=B4/B5", BRL),
        ("Faturamento no balcão", f'=SUMIFS({TOTAL},{CANAL},"Balcão")', BRL),
        ("Faturamento no delivery", f'=SUMIFS({TOTAL},{CANAL},"Delivery")', BRL),
        ("Estornos", f'=SUMIFS({TOTAL},{QTY},"<0")', BRL),
        ("Melhor dia", f"=INDEX('Por dia'!$A${d_first}:$A${d_last},MATCH(MAX('Por dia'!$C${d_first}:$C${d_last}),'Por dia'!$C${d_first}:$C${d_last},0))", DATE),
        ("Faturamento do melhor dia", f"=MAX('Por dia'!$C${d_first}:$C${d_last})", BRL),
        ("Produto campeão (faturamento)", f"=INDEX('Por produto'!$A${first}:$A${last},MATCH(MAX('Por produto'!$D${first}:$D${last}),'Por produto'!$D${first}:$D${last},0))", "@"),
        ("Horário mais forte", f"=INDEX('Por hora'!$A${h_first}:$A${h_last},MATCH(MAX('Por hora'!$B${h_first}:$B${h_last}),'Por hora'!$B${h_first}:$B${h_last},0))&\"h\"", "@"),
        ("Correções feitas nos dados", f"=COUNTA('Correções'!$A$2:$A${max(len(log) + 1, 2)})", "#,##0"),
    ]
    resumo["A3"], resumo["A3"].font = "Indicadores", H2
    for i, (label, formula, f) in enumerate(kpis, start=4):
        a = resumo.cell(row=i, column=1, value=label)
        b = resumo.cell(row=i, column=2, value=formula)
        a.font, b.font, b.number_format = BODY, KPI, f
        a.fill = b.fill = KPI_FILL
        a.border = b.border = BOX
        b.alignment = Alignment(horizontal="right")
        resumo.row_dimensions[i].height = 22
        a.alignment = Alignment(vertical="center")
        b.alignment = Alignment(horizontal="right", vertical="center")
    pay_row = 4 + len(kpis) + 1
    resumo.cell(row=pay_row, column=1, value="Formas de pagamento").font = H2
    _header(resumo, pay_row + 1, ["Forma", "Faturamento", "% do total"])
    methods = df.groupby("Pagamento")["Total"].sum().sort_values(ascending=False).index.tolist()
    for i, m in enumerate(methods, start=pay_row + 2):
        resumo.cell(row=i, column=1, value=m).font = BODY
        c = resumo.cell(row=i, column=2, value=f"=SUMIFS({TOTAL},{PAG},A{i})")
        c.number_format, c.font = BRL, BODY
        p = resumo.cell(row=i, column=3, value=f"=B{i}/$B$4")
        p.number_format, p.font = PCT, BODY
        for col in (1, 2, 3):
            resumo.cell(row=i, column=col).border = BOX
    note_row = pay_row + 2 + len(methods) + 1
    resumo.cell(row=note_row, column=1, value="Todos os números são fórmulas ligadas à aba Vendas: corrija ou acrescente uma venda lá e o relatório se atualiza.").font = NOTE
    resumo.cell(row=note_row + 1, column=1, value="Estornos aparecem como quantidade negativa e já estão descontados do faturamento líquido.").font = NOTE
    _widths(resumo, [32, 22, 14])
    resumo.sheet_view.showGridLines = False

    wb.save(path)
    return {"rows": len(df), "corrections": len(log), "products": len(order), "days": len(days)}
