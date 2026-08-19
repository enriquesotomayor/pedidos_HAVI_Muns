# -*- coding: utf-8 -*-
"""
Fixture sintética con la estructura real del Excel de ventas de HAVI.
Datos inventados (sin información de cliente): NO usar ficheros reales aquí.

Casos límite que cubre (los tests dependen de estas filas; si se añaden
casos, añadirlos al final y ajustar los conteos en test_havi2odoo.py):

- Pedido 5001 en DOS bloques no contiguos (5002 en medio) -> un solo pedido.
- Línea con Cantidad Entregada = 0 dentro de 5001 -> se descarta.
- Pedido 5003 con TODAS las líneas a 0 -> no genera pedido (pedidos_vacios).
- Dos líneas sin Nº Pedido con la misma Nota de Entrega (9100) -> un pedido
  SIN Nº PEDIDO; una tercera sin pedido NI nota -> otro pedido (fecha+tienda).
- Línea de PLACERES MUNS SL -> excluida siempre.
- Debtor "GRUPO CANTALAR, S.L." (grafía distinta de la clave del mapeo
  "Grupo Cantalar S.L.") -> debe resolver a "GRUPO CANTALAR, S.L".
- SALSA CHIMICHURRI (UdM kg): cantidad = Kg Entregados (6.5), no cajas (4).
- AREAS, SAU tiene transporte NO APLICA (sin línea); Cantalar/Muns Vallès
  llevan línea de transporte con la suma de kg y UdM en blanco.
- Fila de totales al final (sin fecha ni artículo) -> se ignora.
"""
import io

import pandas as pd

from havi2odoo import (COL_CIF, COL_CLIENTE, COL_DEBTOR, COL_DESC, COL_FECHA,
                       COL_KG, COL_NOTA, COL_PEDIDO, COL_QTY, COL_SPNR)

FECHA = pd.Timestamp("2026-08-10")


def df_havi_sintetico() -> pd.DataFrame:
    """DataFrame en memoria con las columnas reales del fichero de HAVI."""
    filas = [
        # fecha, cliente(tienda), desc artículo, spnr, debtor, cif, nota, pedido, qty, kg
        # --- Pedido 5001, bloque 1 (Cantalar, grafía distinta del mapeo) ---
        (FECHA, "Tienda Cantalar Centro", "EMPANADA ATÚN", 5268589,
         "GRUPO CANTALAR, S.L.", "B11111111", 9001, 5001, 2, 10.0),
        (FECHA, "Tienda Cantalar Centro", "SALSA CHIMICHURRI", 5268590,
         "GRUPO CANTALAR, S.L.", "B11111111", 9001, 5001, 4, 6.5),
        # --- Pedido 5002 intercalado (AREAS: transporte NO APLICA) ---
        (FECHA, "Areas T1 Aeropuerto", "EMPANADA JAMÓN Y QUESO", 5268591,
         "AREAS, SAU", "A22222222", 9002, 5002, 3, 9.0),
        # --- Pedido 5001, bloque 2 (no contiguo) ---
        (FECHA, "Tienda Cantalar Centro", "ALFAJOR", 5268592,
         "GRUPO CANTALAR, S.L.", "B11111111", 9001, 5001, 1, 3.5),
        # --- Línea a 0 dentro de 5001: se descarta ---
        (FECHA, "Tienda Cantalar Centro", "EMPANADA POLLO ASADO", 5268593,
         "GRUPO CANTALAR, S.L.", "B11111111", 9001, 5001, 0, 0.0),
        # --- Pedido 5003 con todas las líneas a 0: no se genera ---
        (FECHA, "Muns DLG Tienda", "EMPANADA ATÚN", 5268589,
         "MUNS DLG, S.L", "B33333333", 9003, 5003, 0, 0.0),
        # --- Sin Nº Pedido, misma Nota de Entrega (9100): un solo pedido ---
        (FECHA, "Muns Valles Tienda", "EMPANADA ATÚN", 5268589,
         "MUNS VALLES, S.L.", "B44444444", 9100, None, 2, 4.0),
        (FECHA, "Muns Valles Tienda", "EMPANADA TERNERA SUAVE", 5268594,
         "MUNS VALLES, S.L.", "B44444444", 9100, None, 1, 2.0),
        # --- Sin Nº Pedido NI nota: pedido aparte por fecha+tienda ---
        (FECHA, "Pirepona Tienda", "EMPANADA POLLO AL CURRY", 5268595,
         "PIREPONA, S.L.", "B55555555", None, None, 1, 2.0),
        # --- PLACERES MUNS SL: excluido siempre ---
        (FECHA, "Placeres Muns Obrador", "EMPANADA ATÚN", 5268589,
         "PLACERES MUNS SL", "B66666666", 9200, 5004, 5, 10.0),
        # --- Fila de totales (sin fecha ni artículo): debe ignorarse ---
        (None, None, None, None, None, None, None, None, 999, 999.0),
    ]
    return pd.DataFrame(filas, columns=[
        COL_FECHA, COL_CLIENTE, COL_DESC, COL_SPNR, COL_DEBTOR, COL_CIF,
        COL_NOTA, COL_PEDIDO, COL_QTY, COL_KG,
    ])


def xlsx_havi_sintetico() -> io.BytesIO:
    """El mismo DataFrame como xlsx en memoria, para pasar por leer_havi()
    igual que un fichero real (incluida la cabecera 'Nº Pedido ' con espacio)."""
    buf = io.BytesIO()
    df_havi_sintetico().to_excel(buf, index=False, engine="openpyxl")
    buf.seek(0)
    return buf
