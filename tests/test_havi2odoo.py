# -*- coding: utf-8 -*-
"""Tests de havi2odoo.py sobre datos sintéticos (tests/fixture_havi.py).
Ejecutar desde la raíz del repo: pytest"""
import io

import pandas as pd
import pytest

from fixture_havi import xlsx_havi_sintetico
from havi2odoo import (DEFAULT_DEBTOR_MAP, DEFAULT_PRODUCT_MAP,
                       DEFAULT_TRANSPORT_MAP, config_xlsx_a_mapeos,
                       exportar_xlsx, leer_havi, mapeos_a_config_xlsx,
                       procesar)

COLUMNAS_EXPORT = [
    "partner_id", "client_order_ref", "origin", "date_order",
    "order_line/product_id", "order_line/product_uom_qty",
    "order_line/product_uom_id",
]


@pytest.fixture(scope="module")
def resultado():
    df = leer_havi(xlsx_havi_sintetico())
    return procesar(df, DEFAULT_PRODUCT_MAP, DEFAULT_DEBTOR_MAP,
                    DEFAULT_TRANSPORT_MAP)


def _pedido(resultado, origin):
    encontrados = [p for p in resultado.pedidos if p["origin"] == origin]
    assert len(encontrados) == 1, f"esperaba un único pedido con origin {origin!r}"
    return encontrados[0]


# ---------------------------------------------------------------------------
# leer_havi
# ---------------------------------------------------------------------------

def test_fila_totales_ignorada():
    df = leer_havi(xlsx_havi_sintetico())
    # 11 filas en la fixture, la de totales (sin fecha ni artículo) se elimina
    assert len(df) == 10
    assert not (df["Cantidad Entregada"] == 999).any()


# ---------------------------------------------------------------------------
# procesar: agrupaciones y exclusiones
# ---------------------------------------------------------------------------

def test_numero_y_orden_de_pedidos(resultado):
    assert [p["origin"] for p in resultado.pedidos] == [
        "5001", "5002", "SIN Nº PEDIDO", "SIN Nº PEDIDO"]


def test_bloques_no_contiguos_un_solo_pedido(resultado):
    # 5001 viene en dos bloques separados por 5002: debe salir UN pedido
    p = _pedido(resultado, "5001")
    productos = [ln["product_id"] for ln in p["lineas"]]
    # referencias internas (Atún, Salsa Chimichurri, Alfajor) + transporte
    assert productos == ["PA00025", "PA00043", "ME00043",
                         "Transporte Península"]
    assert p["partner_id"] == "GRUPO CANTALAR, S.L"
    assert p["client_order_ref"] == "9001 - Tienda Cantalar Centro"
    assert p["date_order"] == "2026-08-10"
    assert p["revisar"] is False


def test_linea_qty_cero_descartada(resultado):
    p = _pedido(resultado, "5001")
    assert "PA00034" not in [ln["product_id"] for ln in p["lineas"]]  # Pollo Asado
    # informativo: la línea a 0 de 5001 y la de 5003
    assert len(resultado.incidencias.qty_cero) == 2


def test_pedido_todo_a_cero_no_se_genera(resultado):
    assert resultado.incidencias.pedidos_vacios == ["5003"]
    assert not any(p["origin"] == "5003" for p in resultado.pedidos)


def test_sin_pedido_agrupado_por_nota(resultado):
    sin_num = [p for p in resultado.pedidos if p["origin"] == "SIN Nº PEDIDO"]
    assert len(sin_num) == 2
    por_nota = [p for p in sin_num
                if p["client_order_ref"] == "9100 - Muns Valles Tienda"]
    assert len(por_nota) == 1
    p = por_nota[0]
    assert p["partner_id"] == "MUNS VALLES, S.L."
    # referencias internas (Atún, Ternera suave) + transporte
    assert [ln["product_id"] for ln in p["lineas"]] == [
        "PA00025", "PA00009", "Transporte Barcelona"]
    assert p["revisar"] is True


def test_sin_pedido_ni_nota_agrupado_por_fecha_tienda(resultado):
    sueltos = [p for p in resultado.pedidos
               if p["origin"] == "SIN Nº PEDIDO"
               and p["client_order_ref"] == "Pirepona Tienda"]
    assert len(sueltos) == 1
    p = sueltos[0]
    assert p["partner_id"] == "PIREPONA, S.L."
    assert p["revisar"] is True
    # las 3 líneas sin nº de pedido quedan reportadas como incidencia
    assert len(resultado.incidencias.sin_pedido) == 3


def test_placeres_muns_siempre_excluido(resultado):
    assert not any(p["debtor_havi"] == "PLACERES MUNS SL"
                   for p in resultado.pedidos)
    assert not any(p["origin"] == "5004" for p in resultado.pedidos)
    assert len(resultado.incidencias.excluidos) == 1


def test_debtor_grafia_distinta_resuelve_por_normalizacion(resultado):
    # la fixture usa "GRUPO CANTALAR, S.L." y el mapeo "Grupo Cantalar S.L."
    assert _pedido(resultado, "5001")["partner_id"] == "GRUPO CANTALAR, S.L"
    assert resultado.incidencias.debtors_sin_mapeo == []


def test_sin_incidencias_de_mapeo(resultado):
    assert resultado.incidencias.productos_sin_mapeo == []
    assert resultado.incidencias.debtors_sin_transporte == []


# ---------------------------------------------------------------------------
# procesar: cantidades y transporte
# ---------------------------------------------------------------------------

def test_salsa_chimichurri_cantidad_en_kg(resultado):
    p = _pedido(resultado, "5001")
    salsa = [ln for ln in p["lineas"] if ln["product_id"] == "PA00043"][0]
    # Kg Entregados (6.5), NO Cantidad Entregada (4 cajas)
    assert salsa["product_uom_qty"] == 6.5
    assert salsa["product_uom_id"] == "kg"


def test_linea_transporte_suma_kg_y_udm_en_blanco(resultado):
    p = _pedido(resultado, "5001")
    transporte = p["lineas"][-1]
    assert transporte["product_id"] == "Transporte Península"
    assert transporte["product_uom_qty"] == 20.0  # 10 + 6.5 + 3.5
    assert transporte["product_uom_id"] == ""
    assert p["total_kg"] == 20.0


def test_transporte_no_aplica_sin_linea(resultado):
    p = _pedido(resultado, "5002")  # AREAS, SAU -> NO APLICA
    assert [ln["product_id"] for ln in p["lineas"]] == ["PA00001"]  # Jamón y queso
    assert not any("Transporte" in ln["product_id"] for ln in p["lineas"])


# ---------------------------------------------------------------------------
# df_import: formato one2many y export
# ---------------------------------------------------------------------------

def test_columnas_exactas_del_export(resultado):
    assert list(resultado.df_import.columns) == COLUMNAS_EXPORT


def test_one2many_cabecera_solo_en_primera_linea(resultado):
    df = resultado.df_import
    # 4 (5001) + 1 (5002) + 3 (nota 9100) + 2 (suelto) = 10 filas
    assert len(df) == 10
    con_cabecera = df.index[df["partner_id"] != ""].tolist()
    assert con_cabecera == [0, 4, 5, 8]  # primera línea de cada pedido
    cab = ["partner_id", "client_order_ref", "origin", "date_order"]
    for i in df.index:
        if i in con_cabecera:
            assert df.loc[i, "partner_id"] != ""
            assert df.loc[i, "origin"] != ""
            assert df.loc[i, "date_order"] != ""
        else:
            assert all(df.loc[i, c] == "" for c in cab)
    # todas las filas llevan línea de producto
    assert (df["order_line/product_id"] != "").all()


def test_exportar_xlsx_legible(resultado):
    contenido = exportar_xlsx(resultado.df_import)
    df = pd.read_excel(io.BytesIO(contenido))
    assert list(df.columns) == COLUMNAS_EXPORT
    assert len(df) == len(resultado.df_import)
    assert (df["order_line/product_uom_qty"] == 6.5).any()  # la salsa en kg


# ---------------------------------------------------------------------------
# Config xlsx: round-trip de mapeos
# ---------------------------------------------------------------------------

def test_roundtrip_config_xlsx():
    contenido = mapeos_a_config_xlsx(DEFAULT_PRODUCT_MAP, DEFAULT_DEBTOR_MAP,
                                     DEFAULT_TRANSPORT_MAP)
    pmap, dmap, tmap = config_xlsx_a_mapeos(io.BytesIO(contenido))
    assert pmap == DEFAULT_PRODUCT_MAP
    assert dmap == DEFAULT_DEBTOR_MAP
    assert tmap == DEFAULT_TRANSPORT_MAP


def test_config_sin_hoja_transporte_usa_defaults():
    productos = pd.DataFrame(
        [("EMPANADA ATÚN", "PA00025", "Caja 40 Uds")],
        columns=["Desc Artículo HAVI", "Producto Odoo", "UdM Odoo"])
    clientes = pd.DataFrame(
        [("AREAS, SAU", "AREAS, SAU")],
        columns=["Debtor HAVI", "Cliente Odoo"])
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        productos.to_excel(writer, index=False, sheet_name="Productos")
        clientes.to_excel(writer, index=False, sheet_name="Clientes")
    buf.seek(0)
    pmap, dmap, tmap = config_xlsx_a_mapeos(buf)
    assert pmap == {"EMPANADA ATÚN": ("PA00025", "Caja 40 Uds")}
    assert dmap == {"AREAS, SAU": "AREAS, SAU"}
    assert tmap == DEFAULT_TRANSPORT_MAP  # sin hoja Transporte -> embebidos
