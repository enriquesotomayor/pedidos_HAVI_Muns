# -*- coding: utf-8 -*-
"""
Lógica de transformación: Excel de ventas HAVI -> Excel importable en Odoo
como PEDIDOS DE VENTA (sale.order). La facturación se hace después en Odoo
desde el pedido, para que precios y descuentos salgan de las tarifas.

Sin estado: todo entra y sale por parámetros. La UI (Streamlit) vive en app.py.
"""
from __future__ import annotations

import io
import re
from dataclasses import dataclass, field

import pandas as pd

# Columnas esperadas en el fichero de HAVI
COL_FECHA = "Fecha Entrega"
COL_CLIENTE = "Cliente"            # punto de entrega (tienda)
COL_DESC = "Desc Artículo"
COL_SPNR = "SPNR"
COL_DEBTOR = "Debtor"
COL_CIF = "Debtor CIF"
COL_NOTA = "Nota de Entrega"
COL_PEDIDO = "Nº Pedido "          # ojo: espacio final en el original
COL_QTY = "Cantidad Entregada"
COL_KG = "Kg Entregados"

COLUMNAS_REQUERIDAS = [COL_FECHA, COL_CLIENTE, COL_DESC, COL_DEBTOR,
                       COL_NOTA, COL_PEDIDO, COL_QTY, COL_KG]

# Debtors del propio grupo que NUNCA se facturan por este circuito
DEBTORS_EXCLUIDOS = {"PLACERES MUNS SL"}

# Marcador en el mapeo de transporte para clientes sin cargo
NO_APLICA = "NO APLICA"

# ---------------------------------------------------------------------------
# Mapeos por defecto (editables en la UI / sobreescribibles por config xlsx)
# ---------------------------------------------------------------------------

# HAVI (Desc Artículo, en MAYÚSCULAS normalizadas)
#   -> (producto Odoo, UdM Odoo, factor)
# Producto Odoo = REFERENCIA INTERNA (default_code), no traducible: blinda el
# import frente al idioma del usuario. Las UdM van por nombre en español.
# Factor: cantidad Odoo = Cantidad Entregada HAVI × factor. Empanadas en
# unidades sueltas (1 caja HAVI = 40 Unidades); salsa en bolsas (1 caja
# HAVI = 3 bolsas de 2 kg).
DEFAULT_PRODUCT_MAP: dict[str, tuple[str, str, int | float]] = {
    "EMPANADA ATÚN": ("PA00025", "Unidades", 40),
    "EMPANADA CEBOLLA CARAMELIZADA": ("PA00003", "Unidades", 40),
    "EMPANADA CHEESEBURGUER": ("PA00030", "Unidades", 40),
    "EMPANADA CHOCO PLÁTANO": ("PA00016", "Unidades", 40),
    "EMPANADA ESPINACA Y EMMENTAL": ("PA00004", "Unidades", 40),
    "EMPANADA JAMÓN Y QUESO": ("PA00001", "Unidades", 40),
    "EMPANADA MANZANA Y CANELA": ("PA00012", "Unidades", 40),
    "EMPANADA MOZZARELLA Y OLIVADA": ("PA00005", "Unidades", 40),
    "EMPANADA POLLO AL CURRY": ("PA00008", "Unidades", 40),
    "EMPANADA POLLO ASADO": ("PA00034", "Unidades", 40),
    "EMPANADA POLLO THAI": ("PA00010", "Unidades", 40),
    "EMPANADA PROVOLONE Y TOMATE": ("PA00006", "Unidades", 40),
    "EMPANADA PULLED PORK XXL": ("PA00042", "Unidades", 40),
    "EMPANADA SETAS Y CAMEMBERT": ("PA00015", "Unidades", 40),
    "EMPANADA TERNERA PICANTE": ("PA00011", "Unidades", 40),
    "EMPANADA TERNERA ROYALE": ("PA00035", "Unidades", 40),
    "EMPANADA TERNERA SUAVE": ("PA00009", "Unidades", 40),
    "EMPANADA TOMATE Y ALBAHACA": ("PA00002", "Unidades", 40),
    "EMPANADA TÜNA": ("PA00039", "Unidades", 40),
    "SALSA CHIMICHURRI": ("PA00043", "Bolsa 2kg", 3),
    "ALFAJOR": ("ME00043", "Caja de 27", 1),
    "CAJA 4 MUNS": ("MP00122", "Pack 100", 1),
    "CAJA 8 MUNS": ("MP00123", "Pack 100", 1),
    "CAJA 12 MUNS": ("MP00121", "Pack 100", 1),
    "SEPARADOR CAJA 4": ("MP00128", "Pack 200", 1),
    "SEPARADOR CAJA 8": ("MP00129", "Pack 100", 1),
    "SEPARADOR CAJA 12": ("MP00127", "Pack 100", 1),
    "PAPEL ANTIADHERENTE GRANDE": ("MP00124", "Pack 1000", 1),
    "PAPEL ANTIADHERENTE MEDIANO": ("MP00125", "Pack 1000", 1),
}

# Debtor HAVI -> nombre exacto del cliente en Odoo (res.partner).
# Nombres VERIFICADOS contra producción el 18/08/2026.
DEFAULT_DEBTOR_MAP: dict[str, str] = {
    "AREAS, SAU": "AREAS, SAU",
    "Amigos de Muns SL": "Amigos de Muns SL",
    "BAIRESFOODIE DONOSTI SL.": "BAIRESFOODIE DONOSTI SL.",
    "BANZAI FOOD, UNIPESSOAL, LDA": "BANZAI FOOD, UNIPESSOAL, LDA",
    "BELASHARK, S.L.": "BELASHARK, S.L.",
    "BOLDER CORPORATE, S.L.": "BOLDER CORPORATE, S.L.",
    "CANALLA CAPITAL, S.L.": "CANALLA CAPITAL, S.L.",
    "GONZALO LEGALLAIS": "GONZALO LEGALLAIS",
    "Grupo Cantalar S.L.": "GRUPO CANTALAR, S.L",
    "MUNS DLG, S.L": "MUNS DLG, S.L",
    "MUNS VALLES, S.L.": "MUNS VALLES, S.L.",
    "NUTRIM, S.L.": "NUTRIM, S.L.",
    "PIREPONA, S.L.": "PIREPONA, S.L.",
    "PROALDAMA INVESTMENTS, S.L": "PROALDAMA INVESTMENTS, S.L",
    "SCALO BCN BRAND SERVICES, SL": "SCALO BCN BRAND SERVICES, S.L",
    "SURPRESA PROFICUA, LDA": "SURPRESA PROFICUA, LDA",
    "TOP HILL INVESTMENTS, S.L.": "TOP HILL INVESTMENTS, S.L.",
    "VALVI ALIMENTACIÓ I SERVEIS, S.L.": "VALVI ALIMENTACIÓ I SERVEIS, S.L.",
}

# Debtor -> servicio de transporte en Odoo (o NO_APLICA).
# Origen: tabla Transporte_HAVI_Odoo del cliente (18/08/2026).
DEFAULT_TRANSPORT_MAP: dict[str, str] = {
    "AREAS, SAU": NO_APLICA,
    "Amigos de Muns SL": "Transporte Barcelona",
    "BAIRESFOODIE DONOSTI SL.": "Transporte Península",
    "BANZAI FOOD, UNIPESSOAL, LDA": "Transporte Portugal, Andorra e Islas",
    "BELASHARK, S.L.": "Transporte Portugal, Andorra e Islas",
    "BOLDER CORPORATE, S.L.": "Transporte Península",
    "CANALLA CAPITAL, S.L.": "Transporte Península",
    "GONZALO LEGALLAIS": NO_APLICA,
    "Grupo Cantalar S.L.": "Transporte Península",
    "MUNS DLG, S.L": "Transporte Barcelona",
    "MUNS VALLES, S.L.": "Transporte Barcelona",
    "NUTRIM, S.L.": "Transporte Portugal, Andorra e Islas",
    "PIREPONA, S.L.": "Transporte Península",
    "PROALDAMA INVESTMENTS, S.L": "Transporte Portugal, Andorra e Islas",
    "SCALO BCN BRAND SERVICES, SL": "Transporte Península",
    "SURPRESA PROFICUA, LDA": "Transporte Portugal, Andorra e Islas",
    "TOP HILL INVESTMENTS, S.L.": "Transporte Barcelona",
    "VALVI ALIMENTACIÓ I SERVEIS, S.L.": NO_APLICA,
}


def _norm(s: str) -> str:
    """Normaliza razones sociales para lookup tolerante:
    mayúsculas, sin puntos/comas, espacios colapsados.
    'Grupo Cantalar S.L.' y 'GRUPO CANTALAR, S.L' -> misma clave."""
    s = re.sub(r"[.,]", "", str(s).upper())
    return re.sub(r"\s+", " ", s).strip()


def parse_factor(v) -> int | float:
    """Factor multiplicador de cantidad. Celda vacía, ausente o no numérica
    (y valores <= 0, que dejarían las líneas a cero) -> 1. Admite coma decimal."""
    try:
        f = float(str(v).strip().replace(",", "."))
    except (TypeError, ValueError):
        return 1
    if pd.isna(f) or f <= 0:
        return 1
    return int(f) if f.is_integer() else f


# ---------------------------------------------------------------------------
# Estructuras de resultado
# ---------------------------------------------------------------------------

@dataclass
class Incidencias:
    sin_pedido: pd.DataFrame = None          # qty > 0 pero sin Nº Pedido
    qty_cero: pd.DataFrame = None            # líneas descartadas por qty 0
    productos_sin_mapeo: list = field(default_factory=list)
    debtors_sin_mapeo: list = field(default_factory=list)
    debtors_sin_transporte: list = field(default_factory=list)
    pedidos_vacios: list = field(default_factory=list)
    excluidos: pd.DataFrame = None           # líneas de debtors excluidos


@dataclass
class Resultado:
    pedidos: list                 # lista de dicts (cabecera + lineas)
    incidencias: Incidencias
    df_import: pd.DataFrame       # dataframe listo para exportar a Odoo


# ---------------------------------------------------------------------------
# Funciones
# ---------------------------------------------------------------------------

def leer_havi(archivo) -> pd.DataFrame:
    """Lee el Excel de HAVI y valida columnas."""
    df = pd.read_excel(archivo)
    df.columns = [str(c) for c in df.columns]
    if COL_PEDIDO not in df.columns:
        for c in df.columns:
            if c.strip() == COL_PEDIDO.strip():
                df = df.rename(columns={c: COL_PEDIDO})
                break
    faltan = [c for c in COLUMNAS_REQUERIDAS if c not in df.columns]
    if faltan:
        raise ValueError(f"Faltan columnas en el fichero: {', '.join(faltan)}")
    # Eliminar filas de totales / vacías (sin fecha ni artículo)
    df = df[~(df[COL_FECHA].isna() & df[COL_DESC].isna())].copy()
    df[COL_QTY] = pd.to_numeric(df[COL_QTY], errors="coerce").fillna(0)
    df[COL_KG] = pd.to_numeric(df[COL_KG], errors="coerce").fillna(0)
    return df


def _fmt_num(v) -> str:
    """5268589.0 -> '5268589'; deja strings tal cual."""
    if pd.isna(v):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def procesar(df: pd.DataFrame,
             product_map: dict[str, tuple[str, str]],
             debtor_map: dict[str, str],
             transport_map: dict[str, str],
             debtors_incluidos: list[str] | None = None) -> Resultado:
    """Transforma el dataframe HAVI en pedidos de venta agrupados por Nº Pedido."""
    inc = Incidencias()

    # Lookups normalizados (tolerantes a puntuación/mayúsculas)
    debtor_lk = {_norm(k): v for k, v in debtor_map.items()}
    transp_lk = {_norm(k): v for k, v in transport_map.items()}

    # Líneas con cantidad > 0 pero sin Nº Pedido: se generan igualmente como
    # pedidos (Origen = "SIN Nº PEDIDO") para revisión manual
    mask_sin_pedido = df[COL_PEDIDO].isna() & (df[COL_QTY] > 0)
    inc.sin_pedido = df.loc[mask_sin_pedido,
                            [COL_FECHA, COL_CLIENTE, COL_DESC, COL_DEBTOR,
                             COL_NOTA, COL_QTY]].copy()

    # Informativo: líneas con qty 0
    inc.qty_cero = df.loc[df[COL_QTY] == 0,
                          [COL_FECHA, COL_CLIENTE, COL_DESC, COL_DEBTOR,
                           COL_PEDIDO]].copy()

    trabajo = df[df[COL_PEDIDO].notna() | mask_sin_pedido].copy()

    # Exclusión fija del propio grupo (Placeres Muns no se factura a sí misma)
    mask_excl = trabajo[COL_DEBTOR].astype(str).map(_norm).isin(
        {_norm(d) for d in DEBTORS_EXCLUIDOS})
    inc.excluidos = trabajo.loc[mask_excl & (trabajo[COL_QTY] > 0),
                                [COL_FECHA, COL_CLIENTE, COL_DESC,
                                 COL_DEBTOR, COL_PEDIDO, COL_QTY]].copy()
    trabajo = trabajo[~mask_excl]

    if debtors_incluidos is not None:
        trabajo = trabajo[trabajo[COL_DEBTOR].isin(debtors_incluidos)]

    pedidos = []
    prod_sin_mapeo, debt_sin_mapeo, debt_sin_transp = set(), set(), set()

    # Agrupación por VALOR de Nº Pedido (orden de primera aparición).
    # En los ficheros reales de HAVI el mismo pedido puede aparecer en bloques
    # no contiguos (ordenan por proveedor). Las líneas sin Nº Pedido se agrupan
    # por Nota de Entrega (o por fecha+tienda si tampoco hay nota) y generan
    # pedidos con Origen "SIN Nº PEDIDO" para revisión manual.
    def _clave_grupo(row):
        if pd.notna(row[COL_PEDIDO]):
            return ("PED", _fmt_num(row[COL_PEDIDO]))
        if pd.notna(row[COL_NOTA]):
            return ("NOTA", _fmt_num(row[COL_NOTA]))
        return ("SUELTA", f"{row[COL_FECHA]}|{row[COL_DEBTOR]}|{row[COL_CLIENTE]}")

    trabajo["_grp"] = trabajo.apply(_clave_grupo, axis=1)

    for clave, g in trabajo.groupby("_grp", sort=False):
        sin_num = clave[0] != "PED"
        cab = g.iloc[0]
        lineas = g[g[COL_QTY] > 0]
        num_pedido = "SIN Nº PEDIDO" if sin_num else clave[1]
        if lineas.empty:
            inc.pedidos_vacios.append(num_pedido)
            continue

        debtor_havi = str(cab[COL_DEBTOR]).strip()
        partner = debtor_lk.get(_norm(debtor_havi))
        if not partner:
            debt_sin_mapeo.add(debtor_havi)
            partner = debtor_havi  # fallback: nombre tal cual

        fecha = pd.to_datetime(cab[COL_FECHA])
        # Referencia de cliente: "nota - cliente" y, si hay nº de pedido HAVI,
        # con él como sufijo — el origin no se propaga del pedido a la factura
        # (Odoo pone el nombre del sale.order), la referencia sí.
        ref = (f"{_fmt_num(cab[COL_NOTA])} - {str(cab[COL_CLIENTE]).strip()}"
               if pd.notna(cab[COL_NOTA]) else str(cab[COL_CLIENTE]).strip())
        if not sin_num:
            ref = f"{ref} - {num_pedido}"
        pedido = {
            "partner_id": partner,
            "client_order_ref": ref,
            "origin": num_pedido,
            "date_order": fecha.strftime("%Y-%m-%d"),
            "debtor_havi": debtor_havi,
            "revisar": sin_num,
            "lineas": [],
        }
        for _, ln in lineas.iterrows():
            desc = str(ln[COL_DESC]).strip()
            clave = desc.upper()
            if clave in product_map:
                nombre_odoo, udm, factor = product_map[clave]
            else:
                prod_sin_mapeo.add(desc)
                nombre_odoo, udm, factor = desc, "", 1
            # cantidad Odoo = Cantidad Entregada HAVI × factor del mapeo
            # (p. ej. salsa: 1 caja HAVI = 3 × Bolsa 2kg)
            cantidad = round(float(ln[COL_QTY]) * factor, 2)
            pedido["lineas"].append({
                "product_id": nombre_odoo,
                "product_uom_qty": cantidad,
                "product_uom_id": udm,
                "desc_havi": desc,
                "spnr": _fmt_num(ln.get(COL_SPNR, "")),
            })

        # Línea de transporte: suma de Kg Entregados del pedido
        total_kg = round(float(lineas[COL_KG].sum()), 2)
        servicio = transp_lk.get(_norm(debtor_havi))
        if servicio is None:
            debt_sin_transp.add(debtor_havi)
        elif _norm(servicio).startswith(_norm(NO_APLICA)):
            pass  # cliente sin cargo de transporte
        elif total_kg > 0:
            pedido["lineas"].append({
                "product_id": servicio,
                "product_uom_qty": total_kg,
                # "kg" explícito: el importador de Odoo NO hereda la UdM del
                # servicio si la columna va mapeada con la celda vacía
                "product_uom_id": "kg",
                "desc_havi": f"— transporte ({total_kg} kg)",
                "spnr": "",
            })
        pedido["total_kg"] = total_kg
        pedidos.append(pedido)

    inc.productos_sin_mapeo = sorted(prod_sin_mapeo)
    inc.debtors_sin_mapeo = sorted(debt_sin_mapeo)
    inc.debtors_sin_transporte = sorted(debt_sin_transp)

    # Dataframe de importación Odoo (one2many: cabecera solo en la 1ª línea)
    filas = []
    for p in pedidos:
        for i, ln in enumerate(p["lineas"]):
            filas.append({
                "partner_id": p["partner_id"] if i == 0 else "",
                "client_order_ref": p["client_order_ref"] if i == 0 else "",
                "origin": p["origin"] if i == 0 else "",
                "date_order": p["date_order"] if i == 0 else "",
                "order_line/product_id": ln["product_id"],
                "order_line/product_uom_qty": ln["product_uom_qty"],
                "order_line/product_uom_id": ln["product_uom_id"],
            })
    df_import = pd.DataFrame(filas)
    return Resultado(pedidos=pedidos, incidencias=inc, df_import=df_import)


def exportar_xlsx(df_import: pd.DataFrame) -> bytes:
    """Genera el xlsx importable en Odoo (una sola hoja)."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_import.to_excel(writer, index=False, sheet_name="Pedidos")
        ws = writer.sheets["Pedidos"]
        anchos = {"A": 32, "B": 36, "C": 12, "D": 12, "E": 38, "F": 12, "G": 14}
        for col, w in anchos.items():
            ws.column_dimensions[col].width = w
    return buf.getvalue()


def mapeos_a_config_xlsx(product_map: dict, debtor_map: dict,
                         transport_map: dict) -> bytes:
    """Exporta los mapeos actuales a un xlsx de configuración reutilizable."""
    buf = io.BytesIO()
    dfp = pd.DataFrame(
        [(k, v[0], v[1], v[2]) for k, v in product_map.items()],
        columns=["Desc Artículo HAVI", "Producto Odoo", "UdM Odoo", "Factor"])
    dfd = pd.DataFrame(list(debtor_map.items()),
                       columns=["Debtor HAVI", "Cliente Odoo"])
    dft = pd.DataFrame(list(transport_map.items()),
                       columns=["Debtor HAVI", "Servicio transporte Odoo"])
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        dfp.to_excel(writer, index=False, sheet_name="Productos")
        dfd.to_excel(writer, index=False, sheet_name="Clientes")
        dft.to_excel(writer, index=False, sheet_name="Transporte")
    return buf.getvalue()


def config_xlsx_a_mapeos(archivo) -> tuple[dict, dict, dict]:
    """Lee un xlsx de configuración (hojas Productos, Clientes y,
    opcionalmente, Transporte). Compatible con Embalaje_HAVI_odoo.xlsx.
    La columna Factor de Productos es opcional (configs antiguas de
    3 columnas): si falta, o la celda está vacía/no numérica, factor = 1."""
    xl = pd.ExcelFile(archivo)
    dfp = xl.parse("Productos").fillna("")
    dfd = xl.parse("Clientes").fillna("")
    con_factor = "Factor" in dfp.columns
    product_map = {
        str(r["Desc Artículo HAVI"]).strip().upper():
            (str(r["Producto Odoo"]).strip(), str(r["UdM Odoo"]).strip(),
             parse_factor(r["Factor"]) if con_factor else 1)
        for _, r in dfp.iterrows() if str(r["Desc Artículo HAVI"]).strip()
    }
    debtor_map = {
        str(r["Debtor HAVI"]).strip(): str(r["Cliente Odoo"]).strip()
        for _, r in dfd.iterrows() if str(r["Debtor HAVI"]).strip()
    }
    transport_map = dict(DEFAULT_TRANSPORT_MAP)
    if "Transporte" in xl.sheet_names:
        dft = xl.parse("Transporte").fillna("")
        # tolerar cabeceras alternativas (fichero Transporte_HAVI del cliente)
        cols = list(dft.columns)
        c_deb = next((c for c in cols if "debtor" in str(c).lower()
                      or "razon" in str(c).lower()), cols[0])
        c_srv = next((c for c in cols if c != c_deb), cols[-1])
        transport_map = {
            str(r[c_deb]).strip(): str(r[c_srv]).strip()
            for _, r in dft.iterrows() if str(r[c_deb]).strip()
        }
    return product_map, debtor_map, transport_map
