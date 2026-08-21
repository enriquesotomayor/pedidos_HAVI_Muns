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
    # 13 filas en la fixture, la de totales (sin fecha ni artículo) se elimina
    assert len(df) == 12
    assert not (df["Cantidad Entregada"] == 999).any()


# ---------------------------------------------------------------------------
# procesar: agrupaciones y exclusiones
# ---------------------------------------------------------------------------

def test_numero_y_orden_de_pedidos(resultado):
    assert [p["origin"] for p in resultado.pedidos] == [
        "5001", "5002", "SIN Nº PEDIDO", "SIN Nº PEDIDO", "5005"]


def test_bloques_no_contiguos_un_solo_pedido(resultado):
    # 5001 viene en dos bloques separados por 5002: debe salir UN pedido
    p = _pedido(resultado, "5001")
    productos = [ln["product_id"] for ln in p["lineas"]]
    # referencias internas (Atún, Salsa Chimichurri, Alfajor) + transporte
    assert productos == ["PA00025", "PA00043", "ME00043",
                         "Transporte Península"]
    assert p["partner_id"] == "GRUPO CANTALAR, S.L"
    # nota - cliente - nº pedido HAVI (el nº viaja así hasta la factura)
    assert p["client_order_ref"] == "9001 - Tienda Cantalar Centro - 5001"
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
    # sin nº de pedido HAVI: la referencia NO lleva sufijo
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


def test_pedido_sin_nota_ref_cliente_y_numero(resultado):
    # con nº de pedido HAVI pero sin nota: referencia = "cliente - nºpedido"
    p = _pedido(resultado, "5005")
    assert p["client_order_ref"] == "Muns DLG Tienda - 5005"
    assert p["partner_id"] == "MUNS DLG, S.L"
    assert p["revisar"] is False
    assert [ln["product_id"] for ln in p["lineas"]] == [
        "PA00039", "Transporte Barcelona"]  # Tüna + transporte (4.0 kg)


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

def test_salsa_chimichurri_factor_bolsas(resultado):
    p = _pedido(resultado, "5001")
    salsa = [ln for ln in p["lineas"] if ln["product_id"] == "PA00043"][0]
    # 4 cajas HAVI × factor 3 = 12 bolsas de 2 kg (NO los 6.5 Kg Entregados)
    assert salsa["product_uom_qty"] == 12.0
    assert salsa["product_uom_id"] == "Bolsa 2kg"


def test_empanadas_en_unidades_factor_40(resultado):
    p = _pedido(resultado, "5001")
    atun = [ln for ln in p["lineas"] if ln["product_id"] == "PA00025"][0]
    # 2 cajas HAVI × factor 40 = 80 unidades sueltas
    assert atun["product_uom_qty"] == 80.0
    assert atun["product_uom_id"] == "Unidades"


def test_factor_default_1_no_altera_cantidades(resultado):
    p = _pedido(resultado, "5001")
    alfajor = [ln for ln in p["lineas"] if ln["product_id"] == "ME00043"][0]
    # alfajor sigue con factor 1: cantidad = Cantidad Entregada tal cual
    assert alfajor["product_uom_qty"] == 1.0
    assert alfajor["product_uom_id"] == "Caja de 27"


def test_linea_transporte_suma_kg_y_udm_kg(resultado):
    p = _pedido(resultado, "5001")
    transporte = p["lineas"][-1]
    assert transporte["product_id"] == "Transporte Península"
    # Σ Kg Entregados, ajeno al factor: los 6.5 kg de la salsa siguen sumando
    assert transporte["product_uom_qty"] == 20.0  # 10 + 6.5 + 3.5
    assert transporte["product_uom_id"] == "kg"  # explícito: el import no hereda
    assert p["total_kg"] == 20.0


def test_transporte_no_aplica_sin_linea(resultado):
    p = _pedido(resultado, "5002")  # AREAS, SAU -> NO APLICA
    # Jamón y queso + servilletas, sin línea de transporte
    assert [ln["product_id"] for ln in p["lineas"]] == ["PA00001", "MP00130"]
    assert not any("Transporte" in ln["product_id"] for ln in p["lineas"])


def test_servilletas_factor_1_pack_4800(resultado):
    p = _pedido(resultado, "5002")
    serv = [ln for ln in p["lineas"] if ln["product_id"] == "MP00130"][0]
    # factor 1: 2 packs entregados por HAVI -> 2 en Odoo
    assert serv["product_uom_qty"] == 2.0
    assert serv["product_uom_id"] == "Pack 4800"


# ---------------------------------------------------------------------------
# df_import: formato one2many y export
# ---------------------------------------------------------------------------

def test_columnas_exactas_del_export(resultado):
    assert list(resultado.df_import.columns) == COLUMNAS_EXPORT


def test_one2many_cabecera_solo_en_primera_linea(resultado):
    df = resultado.df_import
    # 4 (5001) + 2 (5002) + 3 (nota 9100) + 2 (suelto) + 2 (5005) = 13 filas
    assert len(df) == 13
    con_cabecera = df.index[df["partner_id"] != ""].tolist()
    assert con_cabecera == [0, 4, 6, 9, 11]  # primera línea de cada pedido
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
    assert (df["order_line/product_uom_qty"] == 12.0).any()  # la salsa en bolsas


# ---------------------------------------------------------------------------
# Config xlsx: round-trip de mapeos
# ---------------------------------------------------------------------------

def test_roundtrip_config_xlsx():
    contenido = mapeos_a_config_xlsx(DEFAULT_PRODUCT_MAP, DEFAULT_DEBTOR_MAP,
                                     DEFAULT_TRANSPORT_MAP)
    pmap, dmap, tmap = config_xlsx_a_mapeos(io.BytesIO(contenido))
    assert pmap == DEFAULT_PRODUCT_MAP  # incluye el factor (salsa: 3)
    assert dmap == DEFAULT_DEBTOR_MAP
    assert tmap == DEFAULT_TRANSPORT_MAP


def test_config_xlsx_escribe_columna_factor():
    contenido = mapeos_a_config_xlsx(DEFAULT_PRODUCT_MAP, DEFAULT_DEBTOR_MAP,
                                     DEFAULT_TRANSPORT_MAP)
    dfp = pd.read_excel(io.BytesIO(contenido), sheet_name="Productos")
    assert "Factor" in dfp.columns
    salsa = dfp[dfp["Producto Odoo"] == "PA00043"]
    assert salsa["Factor"].tolist() == [3]
    assert salsa["UdM Odoo"].tolist() == ["Bolsa 2kg"]


def test_config_antigua_3_columnas_factor_1():
    # configs guardadas antes de existir la columna Factor: siguen cargando
    productos = pd.DataFrame(
        [("EMPANADA ATÚN", "PA00025", "Caja 40 Uds"),
         ("SALSA CHIMICHURRI", "PA00043", "Bolsa 2kg")],
        columns=["Desc Artículo HAVI", "Producto Odoo", "UdM Odoo"])
    clientes = pd.DataFrame(
        [("AREAS, SAU", "AREAS, SAU")],
        columns=["Debtor HAVI", "Cliente Odoo"])
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        productos.to_excel(writer, index=False, sheet_name="Productos")
        clientes.to_excel(writer, index=False, sheet_name="Clientes")
    buf.seek(0)
    pmap, _, _ = config_xlsx_a_mapeos(buf)
    assert pmap == {"EMPANADA ATÚN": ("PA00025", "Caja 40 Uds", 1),
                    "SALSA CHIMICHURRI": ("PA00043", "Bolsa 2kg", 1)}


def test_config_factor_vacio_o_no_numerico_es_1():
    productos = pd.DataFrame(
        [("EMPANADA ATÚN", "PA00025", "Caja 40 Uds", ""),
         ("SALSA CHIMICHURRI", "PA00043", "Bolsa 2kg", "x3"),
         ("ALFAJOR", "ME00043", "Caja de 27", 2)],
        columns=["Desc Artículo HAVI", "Producto Odoo", "UdM Odoo", "Factor"])
    clientes = pd.DataFrame(
        [("AREAS, SAU", "AREAS, SAU")],
        columns=["Debtor HAVI", "Cliente Odoo"])
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        productos.to_excel(writer, index=False, sheet_name="Productos")
        clientes.to_excel(writer, index=False, sheet_name="Clientes")
    buf.seek(0)
    pmap, _, _ = config_xlsx_a_mapeos(buf)
    assert pmap["EMPANADA ATÚN"][2] == 1     # celda vacía
    assert pmap["SALSA CHIMICHURRI"][2] == 1  # no numérico
    assert pmap["ALFAJOR"][2] == 2            # numérico se respeta


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
    assert pmap == {"EMPANADA ATÚN": ("PA00025", "Caja 40 Uds", 1)}
    assert dmap == {"AREAS, SAU": "AREAS, SAU"}
    assert tmap == DEFAULT_TRANSPORT_MAP  # sin hoja Transporte -> embebidos
