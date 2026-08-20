# CLAUDE.md — pedidos_HAVI_Muns

App Streamlit **sin persistencia** que convierte el Excel de ventas del operador
logístico HAVI en un Excel importable en Odoo 19 como **pedidos de venta**
(`sale.order`). Las Muns factura después desde los pedidos, de modo que precios,
descuentos, impuestos y vencimientos los aplica Odoo (tarifas, posición fiscal y
plazo de pago de cada cliente). La app NO habla con Odoo: solo transforma ficheros.

## Arquitectura

- `app.py` — UI Streamlit. No contiene lógica de negocio.
- `havi2odoo.py` — lógica pura (pandas): parseo, agrupación, mapeos, export.
  Testeable sin UI. Los mapeos por defecto (`DEFAULT_*`) van embebidos aquí
  y son la fuente de verdad; la config xlsx que sube el usuario los pisa
  solo durante la sesión.
- `requirements.txt` — streamlit, pandas, openpyxl.

Despliegue: Streamlit Community Cloud conectado a este repo; **cada push a
`main` redespliega automáticamente**. No hay staging: probar en local
(`streamlit run app.py`) antes de pushear.

## Reglas de negocio (NO cambiar sin decisión de Enrique)

- Agrupación por VALOR de `Nº Pedido ` (ojo: espacio final en la cabecera
  real). El mismo pedido puede venir en bloques no contiguos (HAVI ordena
  por proveedor): jamás agrupar secuencialmente.
- Líneas con cantidad > 0 y SIN nº de pedido: se generan como pedidos con
  Origen `SIN Nº PEDIDO`, agrupadas por Nota de Entrega o, si tampoco hay,
  por fecha+tienda. Se marcan para revisión manual.
- Cantidad 0 ⇒ línea descartada; pedido sin líneas ⇒ no se genera.
- `PLACERES MUNS SL` SIEMPRE excluido (no se autofactura). Amigos de Muns
  es cliente normal.
- Referencia de cliente = `Nota de Entrega - Cliente - Nº Pedido HAVI`:
  el nº de pedido viaja así hasta la factura (Odoo no propaga el `origin`
  del pedido; en la factura pone el nombre del sale.order). Los pedidos
  `SIN Nº PEDIDO` van sin sufijo: `Nota - Cliente` o solo `Cliente`.
- Cada producto del mapeo lleva un FACTOR multiplicador: cantidad Odoo =
  Cantidad Entregada HAVI × factor (default 1). Las empanadas se facturan
  en UNIDADES sueltas: UdM "Unidades", factor 40 (1 caja HAVI = 40 uds; la
  caja es solo el dato de HAVI). Salsa Chimichurri: UdM "Bolsa 2kg",
  factor 3 (1 caja HAVI = 3 bolsas). Alfajor y embalaje siguen a factor 1.
- Línea de transporte por pedido: Σ Kg Entregados × servicio según mapeo de
  transporte del debtor (`NO APLICA` = sin línea). La app pone servicio y kg;
  el precio/kg vive en Odoo. UdM de esa línea = "kg" EXPLÍCITO: el
  importador de Odoo no hereda la UdM del servicio si la columna va mapeada
  con la celda vacía (llegaba sin UdM y con precio 0).
- El xlsx de salida usa nombres técnicos de campo de `sale.order`
  (`partner_id`, `client_order_ref`, `origin`, `date_order`,
  `order_line/product_id`, `order_line/product_uom_qty`,
  `order_line/product_uom_id`) con formato one2many: cabecera solo en la
  primera línea de cada pedido. Se importa desde Ventas → Pedidos, NUNCA
  desde Compras ni Contabilidad.

## Datos verificados contra producción lasmuns (18/08/2026)

- `DEFAULT_PRODUCT_MAP` casa producto por REFERENCIA INTERNA
  (`default_code`, p. ej. `PA00025` = Empanada Atún), no traducible:
  inmune a las traducciones de nombre y al idioma del usuario que importa.
  Las UdM sí van por nombre en español (p. ej. "Unidades", "Pack 200").
- Nombres de cliente de `DEFAULT_DEBTOR_MAP` verificados (p. ej. HAVI
  escribe "Grupo Cantalar S.L." y en Odoo es "GRUPO CANTALAR, S.L").
  El lookup normaliza puntos/comas/mayúsculas del lado HAVI.
- Sin variantes de Mercado en producción: el matching por nombre de
  plantilla es unívoco HOY. Si algún día se despliegan variantes
  (ESP/ENG/DE/MULTI), cambiar la columna "Producto Odoo" de la config a
  referencias de variante — no requiere tocar código.
- Servicios de transporte esperados en Odoo: `Transporte Barcelona`,
  `Transporte Península`, `Transporte Portugal, Andorra e Islas`, tipo
  servicio, UdM kg, precio €/kg.

## Cautelas

- Cambios de formato de la config xlsx (hojas `Productos`, `Clientes`,
  `Transporte`) rompen ficheros guardados por los usuarios: mantener
  retrocompatibilidad (la hoja Transporte y la columna Factor de Productos
  son opcionales).
- Probar SIEMPRE contra un fichero real de HAVI antes de pushear; validar
  nº de pedidos, los `SIN Nº PEDIDO`, las empanadas en unidades
  (cajas × 40), la línea de salsa en bolsas (cajas × 3) y los kg de
  transporte.
- No añadir dependencias pesadas ni llamadas de red: la app debe seguir
  siendo un transformador puro de ficheros.
- Decisión pendiente (Enrique/cliente): confirmar que el precio de tarifa
  de la salsa casa con la UdM "Bolsa 2kg" (19/08/2026: la salsa deja de ir
  en kg y pasa a factor 3); discrepancia conocida Separador Caja 4 (tabla
  cliente Pack 100 vs producto Odoo Pack 200 — 19/08/2026: el mapeo usa
  "Pack 200", el nombre español real de la UdM del producto en Odoo).
