# -*- coding: utf-8 -*-
"""
HAVI → Odoo · Generador de pedidos de venta (Las Muns)

App Streamlit sin persistencia: se sube el Excel de ventas de HAVI y se
descarga un Excel importable en Odoo (Ventas → Pedidos → Importar registros).
La facturación se hace en Odoo desde los pedidos, de modo que precios y
descuentos salen de las tarifas de cada cliente. No se guarda ningún dato.
"""
import pandas as pd
import streamlit as st

import havi2odoo as h2o

st.set_page_config(page_title="HAVI → Odoo · Pedidos", page_icon="🥟",
                   layout="wide")

st.title("🥟 HAVI → Odoo · Pedidos de venta")
st.caption(
    "Sube el Excel de ventas de HAVI y descarga un Excel importable en Odoo "
    "(`sale.order`). Después, en Odoo: confirmar los pedidos y facturar — "
    "precios y descuentos los aplica Odoo según la tarifa de cada cliente. "
    "La app no almacena datos: todo se procesa en memoria durante la sesión."
)

# ---------------------------------------------------------------------------
# Estado de sesión: mapeos editables
# ---------------------------------------------------------------------------
if "product_map" not in st.session_state:
    st.session_state.product_map = dict(h2o.DEFAULT_PRODUCT_MAP)
if "debtor_map" not in st.session_state:
    st.session_state.debtor_map = dict(h2o.DEFAULT_DEBTOR_MAP)
if "transport_map" not in st.session_state:
    st.session_state.transport_map = dict(h2o.DEFAULT_TRANSPORT_MAP)

# ---------------------------------------------------------------------------
# Barra lateral: configuración
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ Configuración")
    st.markdown(
        "**Reglas fijas**\n"
        "- Fecha del pedido = *Fecha Entrega*\n"
        "- Referencia de cliente = *Nota de Entrega* − *Cliente* "
        "(pasa a la Referencia de la factura)\n"
        "- Documento origen = *Nº Pedido*\n"
        "- Mismo *Nº Pedido* ⇒ mismo pedido de venta\n"
        "- Líneas con cantidad 0 se descartan\n"
        "- **PLACERES MUNS SL siempre excluido** (no se factura a sí misma)\n"
        "- Línea de transporte: suma de *Kg Entregados* × servicio del "
        "cliente (precio/kg registrado en Odoo)\n"
        "- Precios, descuentos y plazo de pago (30 días) los aplica Odoo "
        "según tarifa y ficha del cliente"
    )
    st.divider()
    st.subheader("Config reutilizable")
    st.caption(
        "La app no guarda nada entre sesiones. Los mapeos por defecto van "
        "embebidos; puedes ajustarlos abajo y descargarlos como xlsx, o "
        "subir aquí un xlsx modificado (hojas `Productos`, `Clientes` y "
        "opcionalmente `Transporte` — compatible con Embalaje_HAVI_odoo)."
    )
    cfg_up = st.file_uploader("Subir configuración (.xlsx)", type=["xlsx"],
                              key="cfg")
    if cfg_up is not None:
        try:
            pm, dm, tm = h2o.config_xlsx_a_mapeos(cfg_up)
            st.session_state.product_map = pm
            st.session_state.debtor_map = dm
            st.session_state.transport_map = tm
            st.success(f"Config cargada: {len(pm)} productos, {len(dm)} "
                       f"clientes, {len(tm)} transportes.")
        except Exception as e:
            st.error(f"No se pudo leer la configuración: {e}")

    st.download_button(
        "⬇️ Descargar configuración actual",
        data=h2o.mapeos_a_config_xlsx(st.session_state.product_map,
                                      st.session_state.debtor_map,
                                      st.session_state.transport_map),
        file_name="config_havi_odoo.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ---------------------------------------------------------------------------
# Editores de mapeos
# ---------------------------------------------------------------------------
with st.expander("🗺️ Mapeo de productos (HAVI → Odoo + UdM)"):
    dfp = pd.DataFrame(
        [(k, v[0], v[1]) for k, v in sorted(st.session_state.product_map.items())],
        columns=["Desc Artículo HAVI", "Producto Odoo", "UdM Odoo"])
    dfp_edit = st.data_editor(dfp, num_rows="dynamic", width="stretch",
                              key="edit_prod")
    st.session_state.product_map = {
        str(r["Desc Artículo HAVI"]).strip().upper():
            (str(r["Producto Odoo"]).strip(), str(r["UdM Odoo"]).strip())
        for _, r in dfp_edit.iterrows()
        if str(r["Desc Artículo HAVI"]).strip() not in ("", "None", "nan")
    }
    st.caption(
        "El **Producto Odoo** debe identificar UNÍVOCAMENTE el producto en "
        "Odoo (nombre exacto o referencia interna). Ojo con las empanadas "
        "con variantes de Mercado: el nombre de plantilla coincide con las "
        "4 variantes y el import lo rechaza — usar la referencia de la "
        "variante cuando estén cargadas (p. ej. `PA00001-ESP`)."
    )

with st.expander("👥 Mapeo de clientes (Debtor HAVI → Cliente Odoo)"):
    dfd = pd.DataFrame(sorted(st.session_state.debtor_map.items()),
                       columns=["Debtor HAVI", "Cliente Odoo"])
    dfd_edit = st.data_editor(dfd, num_rows="dynamic", width="stretch",
                              key="edit_debt")
    st.session_state.debtor_map = {
        str(r["Debtor HAVI"]).strip(): str(r["Cliente Odoo"]).strip()
        for _, r in dfd_edit.iterrows()
        if str(r["Debtor HAVI"]).strip() not in ("", "None", "nan")
    }
    st.caption("Nombres verificados contra producción (18/08/2026). El "
               "matching tolera diferencias de puntos, comas y mayúsculas "
               "en la columna Debtor del fichero de HAVI.")

with st.expander("🚚 Mapeo de transporte (Debtor → servicio Odoo)"):
    dft = pd.DataFrame(sorted(st.session_state.transport_map.items()),
                       columns=["Debtor HAVI", "Servicio transporte Odoo"])
    dft_edit = st.data_editor(dft, num_rows="dynamic", width="stretch",
                              key="edit_transp")
    st.session_state.transport_map = {
        str(r["Debtor HAVI"]).strip():
            str(r["Servicio transporte Odoo"]).strip()
        for _, r in dft_edit.iterrows()
        if str(r["Debtor HAVI"]).strip() not in ("", "None", "nan")
    }
    st.caption(
        f"Escribe `{h2o.NO_APLICA}` para clientes sin cargo de transporte. "
        "El servicio debe existir en Odoo como producto de tipo servicio "
        "con UdM kg y su precio por kg — la app solo pone el servicio y "
        "los kg; el importe lo calcula Odoo."
    )

st.divider()

# ---------------------------------------------------------------------------
# Carga del fichero de ventas
# ---------------------------------------------------------------------------
archivo = st.file_uploader("📤 Excel de ventas de HAVI", type=["xlsx", "xls"])

if archivo is None:
    st.info("Sube el fichero de ventas para empezar.")
    st.stop()

try:
    df = h2o.leer_havi(archivo)
except Exception as e:
    st.error(f"Error leyendo el fichero: {e}")
    st.stop()

st.success(f"Fichero leído: **{len(df)}** líneas.")

# Filtro de debtors (los excluidos fijos ni se ofrecen)
debtors = sorted(
    d for d in df[h2o.COL_DEBTOR].dropna().unique().tolist()
    if h2o._norm(d) not in {h2o._norm(x) for x in h2o.DEBTORS_EXCLUIDOS})
sel_debtors = st.multiselect(
    "Clientes (Debtor) a incluir", options=debtors, default=debtors,
    help="PLACERES MUNS SL queda excluido siempre de forma automática.")

resultado = h2o.procesar(df, st.session_state.product_map,
                         st.session_state.debtor_map,
                         st.session_state.transport_map,
                         debtors_incluidos=sel_debtors)
inc = resultado.incidencias

# ---------------------------------------------------------------------------
# Incidencias
# ---------------------------------------------------------------------------
problemas = []
if inc.productos_sin_mapeo:
    problemas.append(
        ("🟥 Productos sin mapeo (se exportan con el nombre HAVI tal cual y "
         "sin UdM — la importación fallará si no coinciden en Odoo):",
         pd.DataFrame({"Desc Artículo HAVI": inc.productos_sin_mapeo})))
if inc.debtors_sin_mapeo:
    problemas.append(
        ("🟥 Debtors sin mapeo de cliente (se exporta el nombre HAVI tal "
         "cual):", pd.DataFrame({"Debtor HAVI": inc.debtors_sin_mapeo})))
if inc.debtors_sin_transporte:
    problemas.append(
        ("🟧 Debtors sin entrada en el mapeo de transporte — sus pedidos "
         "salen SIN línea de transporte:",
         pd.DataFrame({"Debtor HAVI": inc.debtors_sin_transporte})))
if inc.sin_pedido is not None and len(inc.sin_pedido):
    problemas.append(
        ("🟧 Líneas con cantidad entregada pero SIN Nº Pedido — SE INCLUYEN "
         "como pedidos con Origen `SIN Nº PEDIDO` (agrupadas por nota de "
         "entrega o fecha+tienda). Revisarlos a mano tras importar:",
         inc.sin_pedido))
if inc.pedidos_vacios:
    problemas.append(
        ("🟨 Pedidos descartados por quedarse sin líneas (todas a cantidad "
         "0):", pd.DataFrame({"Nº Pedido": inc.pedidos_vacios})))

if problemas:
    st.subheader("Incidencias")
    for titulo, tabla in problemas:
        st.markdown(titulo)
        st.dataframe(tabla, width="stretch", hide_index=True)

col1, col2 = st.columns(2)
with col1:
    with st.expander(f"ℹ️ Líneas descartadas por cantidad 0 "
                     f"({len(inc.qty_cero)})"):
        st.dataframe(inc.qty_cero, width="stretch", hide_index=True)
with col2:
    n_excl = len(inc.excluidos) if inc.excluidos is not None else 0
    with st.expander(f"ℹ️ Líneas de PLACERES MUNS SL excluidas ({n_excl})"):
        if n_excl:
            st.dataframe(inc.excluidos, width="stretch", hide_index=True)
        else:
            st.write("Ninguna en este fichero.")

# ---------------------------------------------------------------------------
# Vista previa de pedidos
# ---------------------------------------------------------------------------
st.subheader(f"Pedidos de venta a generar: {len(resultado.pedidos)}")

for p in resultado.pedidos:
    marca = "🟧 REVISAR · " if p.get("revisar") else ""
    with st.expander(
            f"{marca}**{p['partner_id']}** · Ref cliente `{p['client_order_ref']}` "
            f"· Origen `{p['origin']}` · {p['date_order']} · "
            f"{len(p['lineas'])} líneas · {p['total_kg']:g} kg"):
        st.dataframe(
            pd.DataFrame(p["lineas"])[
                ["spnr", "desc_havi", "product_id", "product_uom_qty",
                 "product_uom_id"]
            ].rename(columns={
                "spnr": "SPNR", "desc_havi": "Artículo HAVI",
                "product_id": "Producto Odoo",
                "product_uom_qty": "Cantidad", "product_uom_id": "UdM"}),
            width="stretch", hide_index=True)

# ---------------------------------------------------------------------------
# Descarga del fichero de importación
# ---------------------------------------------------------------------------
st.divider()
if len(resultado.df_import):
    st.download_button(
        "⬇️ Descargar Excel importable en Odoo (pedidos de venta)",
        data=h2o.exportar_xlsx(resultado.df_import),
        file_name="import_pedidos_havi_odoo.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
    st.markdown(
        "**Cómo seguir en Odoo:** Ventas → Pedidos → ⚙️ Importar registros "
        "(las cabeceras técnicas se mapean solas). Los pedidos quedan en "
        "borrador con precios y descuentos calculados por la tarifa de cada "
        "cliente. Revisar → Confirmar → Crear factura (se puede en lote). "
        "La Referencia de cliente pasa a la Referencia de la factura y el "
        "vencimiento sale del plazo de pago (30 días) de la ficha del "
        "cliente. Probar primero en `muns-pruebas`."
    )
else:
    st.error("No hay líneas facturables con los filtros actuales.")
