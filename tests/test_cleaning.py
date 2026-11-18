from datetime import date, time

import pandas as pd
import pytest

from sales_report.cleaning import (UNKNOWN_PREFIX, clean, normalize_payment, normalize_product,
                                   parse_date, parse_price, parse_time)


@pytest.mark.parametrize("raw, cents", [
    ("R$ 6,00", 600), ("6.00", 600), ("6", 600), ("6.0", 600), ("R$ 1.234,50", 123450),
    (" 12,5 ", 1250), (14.95, 1495), (7, 700),
])
def test_parse_price(raw, cents):
    assert parse_price(raw) == cents


def test_parse_price_rejects_empty():
    with pytest.raises(ValueError):
        parse_price("  ")


@pytest.mark.parametrize("raw", ["05/10/2026", "2026-10-05", "05-10-2026", "05/10/26"])
def test_parse_date_formats(raw):
    assert parse_date(raw) == date(2026, 10, 5)


def test_parse_date_rejects_garbage():
    with pytest.raises(ValueError):
        parse_date("ontem")


@pytest.mark.parametrize("raw, expected", [("07:05", time(7, 5)), ("07:05:00", time(7, 5)), ("7h30", time(7, 30))])
def test_parse_time(raw, expected):
    assert parse_time(raw) == expected


@pytest.mark.parametrize("raw, expected, how", [
    ("Café coado", "Café coado", "exact"),
    ("CAFÉ COADO ", "Café coado", "exact"),      # case, accents, spaces
    ("Expresso", "Espresso", "alias"),           # common misspelling
    ("tapioca coalho", "Tapioca de queijo coalho", "alias"),
    ("Croisant", "Croissant", "fuzzy"),          # typo
    ("Cuscus com ovo", "Cuscuz com ovo", "fuzzy"),
])
def test_normalize_product(raw, expected, how):
    assert normalize_product(raw) == (expected, how)


def test_unknown_product_is_flagged_not_dropped():
    name, how = normalize_product("Pastel de vento")
    assert how == "unknown" and name == UNKNOWN_PREFIX + "Pastel de vento"


@pytest.mark.parametrize("raw, expected", [("pix", "PIX"), ("Cartão Débito", "Débito"), ("credito", "Crédito"), ("dinheiro", "Dinheiro")])
def test_normalize_payment(raw, expected):
    assert normalize_payment(raw) == expected


def test_clean_end_to_end(tmp_path):
    a = tmp_path / "caixa1.csv"
    a.write_text(
        "Data;Hora;Produto;Qtd;Valor Unit.;Pagamento\n"
        "01/10/2026;07:10;CAFÉ COADO ;2;R$ 6,00;pix\n"
        "\n"
        "Data;Hora;Produto;Qtd;Valor Unit.;Pagamento\n"   # repeated header
        "01/10/2026;08:00;Croisant;1;R$ 11,00;Dinheiro\n",
        encoding="utf-8")
    b = tmp_path / "caixa2.csv"
    b.write_text(
        "sale_id,date,time,item,quantity,unit_price,payment\n"
        "1,2026-10-01,09:00:00,Expresso,1,7.0,debito\n"
        "1,2026-10-01,09:00:00,Expresso,1,7.0,debito\n"       # duplicate (same sale id)
        "2,2026-10-01,09:00:00,Expresso,1,7.0,debito\n"       # same content, different sale: kept
        "2-E,2026-10-01,09:05:00,Expresso,-1,7.0,debito\n",   # refund
        encoding="utf-8")
    c = tmp_path / "delivery.xlsx"
    pd.DataFrame({"Data do pedido": [pd.Timestamp("2026-10-01 12:30")], "Item": ["bolo de rolo"],
                  "Qtd": [2], "Total": [23.0]}).to_excel(c, index=False)

    df, log = clean([a, b, c])
    actions = log.to_frame()["Ação"].tolist()

    assert len(df) == 6
    assert df["Total"].sum() == pytest.approx(12 + 11 + 7 + 7 - 7 + 23)
    assert set(df["Produto"]) == {"Café coado", "Croissant", "Espresso", "Bolo de rolo (fatia)"}
    assert df.loc[df["Canal"] == "Delivery", "Preço unitário"].item() == pytest.approx(11.5)
    assert "Venda duplicada removida" in actions
    assert "Cabeçalho duplicado removido" in actions
    assert "Linha vazia removida" in actions
    assert "Estorno mantido como valor negativo" in actions
