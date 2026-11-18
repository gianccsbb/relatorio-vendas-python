"""Generate one month of realistic, *messy* sales exports for the demo.

Three sources, each messy in its own way, like real small-business data:
  - caixa1 (POS register 1): ';' separator, dd/mm/yyyy, prices like "R$ 6,00",
    inconsistent capitalization, blank lines, header repeated mid-file
  - caixa2 (POS register 2): ',' separator, ISO dates, prices like "6.0",
    typos in product names, duplicated rows, refunds as negative quantities
  - delivery app (Excel): one datetime column, line total instead of unit price

Run:  python scripts/make_sample_data.py
"""
import csv
import random
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from sales_report.catalog import CATALOG  # noqa: E402

RAW = ROOT / "data" / "raw"
random.seed(42)

# How each register spells products (messy on purpose)
SPELLINGS = {
    "Café coado": ["Café Coado", "cafe coado", "CAFÉ COADO ", "Cafe Coado", "Café coado"],
    "Espresso": ["Espresso", "Expresso", "espresso"],
    "Cappuccino": ["Cappuccino", "Capuccino", "capuccino"],
    "Café gelado com rapadura": ["Café gelado com rapadura", "cafe gelado", "Gelado rapadura"],
    "Pão na chapa": ["Pão na chapa", "pao chapa", "Pao na Chapa", "PÃO NA CHAPA"],
    "Tapioca de queijo coalho": ["Tapioca de queijo coalho", "Tapioca queijo coalho", "tapioca coalho"],
    "Cuscuz com ovo": ["Cuscuz com ovo", "cuscuz c/ ovo", "Cuscus com ovo"],
    "Croissant": ["Croissant", "croissant", "Croisant"],
    "Bolo de rolo (fatia)": ["Bolo de rolo (fatia)", "bolo de rolo", "Bolo de Rolo fatia"],
    "Bolo Souza Leão": ["Bolo Souza Leão", "bolo souza leao", "Souza Leao"],
    "Cartola": ["Cartola", "cartola"],
    "Brigadeiro": ["Brigadeiro", "brigadeiro un", "Brigadeiro unid"],
}
# Popularity by hour block (morning favours coffee and bread)
MORNING = {"Café coado": 30, "Pão na chapa": 22, "Cuscuz com ovo": 12, "Tapioca de queijo coalho": 12,
           "Espresso": 8, "Cappuccino": 8, "Croissant": 6, "Bolo de rolo (fatia)": 4}
AFTERNOON = {"Café coado": 14, "Cappuccino": 12, "Café gelado com rapadura": 14, "Bolo de rolo (fatia)": 14,
             "Bolo Souza Leão": 8, "Cartola": 9, "Brigadeiro": 12, "Croissant": 6, "Espresso": 7, "Tapioca de queijo coalho": 4}


def pick(weights):
    names, w = zip(*weights.items())
    return random.choices(names, weights=w)[0]


def sales_for_day(d: date):
    sunday = d.weekday() == 6
    close = 12 if sunday else 19
    weekend_boost = 1.35 if d.weekday() >= 5 else 1.0
    out = []
    for hour in range(7, close):
        base = {7: 9, 8: 12, 9: 9, 10: 6, 11: 4, 12: 4, 13: 3, 14: 4, 15: 7, 16: 8, 17: 7, 18: 4}[hour]
        for _ in range(int(random.gauss(base * weekend_boost, 2)) + 1):
            prod = pick(MORNING if hour < 11 else AFTERNOON)
            qty = random.choices([1, 2, 3], weights=[80, 16, 4])[0]
            ts = datetime.combine(d, datetime.min.time()) + timedelta(hours=hour, minutes=random.randint(0, 59))
            out.append((ts, prod, qty))
    return out


def main():
    RAW.mkdir(parents=True, exist_ok=True)
    month = [date(2026, 10, 1) + timedelta(days=i) for i in range(31)]
    c1, c2, delivery = [], [], []
    for d in month:
        for ts, prod, qty in sales_for_day(d):
            r = random.random()
            if r < 0.55:
                c1.append((ts, prod, qty))
            elif r < 0.85:
                c2.append((ts, prod, qty))
            elif 10 <= ts.hour < 19:
                delivery.append((ts, prod, qty))
            else:
                c1.append((ts, prod, qty))

    pay1 = ["PIX", "Pix", "pix", "Cartão Débito", "Cartao Debito", "Cartão Crédito", "Dinheiro", "dinheiro"]
    with open(RAW / "caixa1_outubro.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter=";")
        header = ["Data", "Hora", "Produto", "Qtd", "Valor Unit.", "Pagamento"]
        w.writerow(header)
        for i, (ts, prod, qty) in enumerate(c1):
            if i in (412, 1260):  # export glitch: header repeated
                w.writerow(header)
            if random.random() < 0.01:
                w.writerow([])
            price = CATALOG[prod][1] / 100
            w.writerow([ts.strftime("%d/%m/%Y"), ts.strftime("%H:%M"), random.choice(SPELLINGS[prod]), qty,
                        f"R$ {price:.2f}".replace(".", ","), random.choice(pay1)])

    pay2 = ["pix", "debito", "credito", "dinheiro", "PIX"]
    with open(RAW / "caixa2_outubro.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["sale_id", "date", "time", "item", "quantity", "unit_price", "payment"])
        for sale_id, (ts, prod, qty) in enumerate(c2, start=50001):
            price = CATALOG[prod][1] / 100
            row = [sale_id, ts.strftime("%Y-%m-%d"), ts.strftime("%H:%M:%S"), random.choice(SPELLINGS[prod]), qty,
                   f"{price:g}" if random.random() < .5 else f"{price:.1f}", random.choice(pay2)]
            w.writerow(row)
            if random.random() < 0.02:   # duplicated by the POS
                w.writerow(row)
            if random.random() < 0.008:  # refund later the same hour
                w.writerow([f"{sale_id}-E", row[1], (ts + timedelta(minutes=5)).strftime("%H:%M:%S"), row[3], -1, row[5], row[6]])

    wb = Workbook()
    ws = wb.active
    ws.title = "Pedidos"
    ws.append(["Data do pedido", "Item", "Qtd", "Total", "Taxa app"])
    for ts, prod, qty in delivery:
        total = round(CATALOG[prod][1] * 1.15 / 100 * qty, 2)  # delivery menu is 15% higher
        ws.append([ts, random.choice(SPELLINGS[prod]), qty, total, round(total * 0.12, 2)])
        ws.cell(row=ws.max_row, column=1).number_format = "dd/mm/yyyy hh:mm"
    wb.save(RAW / "delivery_outubro.xlsx")

    print(f"caixa1: {len(c1)} sales | caixa2: {len(c2)} sales | delivery: {len(delivery)} orders")


if __name__ == "__main__":
    main()
