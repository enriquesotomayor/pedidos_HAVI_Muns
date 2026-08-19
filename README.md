# HAVI → Odoo · Generador de pedidos de venta (Las Muns)

App Streamlit **sin persistencia**: sube el Excel de ventas que envía HAVI y
descarga un Excel importable en Odoo que crea los **pedidos de venta**
(`sale.order`). La facturación se hace después en Odoo desde los pedidos,
para que precios y descuentos salgan de las tarifas de cada cliente.

## Reglas de transformación

| Excel HAVI | Odoo (sale.order) |
|---|---|
| Fecha Entrega | Fecha del pedido (`date_order`) |
| Nota de Entrega − Cliente | Referencia de cliente (`client_order_ref`) → pasa a Referencia de la factura |
| Nº Pedido | Documento origen (`origin`) |
| Debtor | Cliente (`partner_id`, vía mapeo, verificado contra producción) |
| Desc Artículo | Producto de línea (vía mapeo) |
| Cantidad Entregada | Cantidad de línea |
| (tabla Embalaje) | UdM de línea (`Caja 40 Uds`, `Pack 100`…) |
| Σ Kg Entregados del pedido | Línea de servicio de transporte (qty en kg; precio/kg lo pone Odoo) |

- Agrupación por valor de **Nº Pedido** (mismo pedido = mismo sale.order,
  aunque las líneas no sean contiguas en el fichero).
- Líneas con cantidad 0 se descartan; pedidos vacíos no se generan.
- Líneas con cantidad pero sin Nº Pedido se reportan como incidencia.
- **PLACERES MUNS SL siempre excluido** (no se factura a sí misma).
  Amigos de Muns SL se factura como un cliente normal.
- Transporte por cliente según mapeo (tabla Transporte_HAVI); `NO APLICA`
  = sin línea. El servicio debe existir en Odoo con UdM kg y precio/kg.
- Precio, descuento y plazo de pago (30 días) los aplica Odoo: tarifa y
  ficha del cliente.

## Estructura

- `app.py` — interfaz Streamlit
- `havi2odoo.py` — lógica pura (parseo, agrupación, export), testeable sin UI
- `tests/` — suite pytest sobre datos sintéticos (sin ficheros reales de HAVI)
- `requirements.txt` — dependencias de la app (las que instala Streamlit Cloud)
- `requirements-dev.txt` — dependencias solo de desarrollo (pytest)

## Ejecutar en local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Tests

Los tests no usan ficheros reales de HAVI (datos de cliente): trabajan
sobre un DataFrame sintético generado en memoria (`tests/fixture_havi.py`)
que cubre los casos límite (bloques no contiguos, líneas a 0, pedidos
`SIN Nº PEDIDO`, exclusión de PLACERES MUNS, salsa en kg, transporte,
fila de totales…). Para lanzarlos, desde la raíz del repo:

```bash
pip install -r requirements-dev.txt
pytest
```

`pytest` va solo en `requirements-dev.txt` a propósito: Streamlit Cloud
instala `requirements.txt` y no debe cargar dependencias de test.

## Publicar en Streamlit Community Cloud

1. Subir estos tres ficheros a un repo de GitHub (público o privado).
2. En https://share.streamlit.io → "Create app" → repo, rama, `app.py`.
3. Deploy. Cada push redespliega automáticamente.

## Configuración reutilizable

Mapeos editables en la propia app (productos, clientes, transporte) y
exportables/importables como un único xlsx con hojas `Productos`,
`Clientes` y `Transporte` (compatible con `Embalaje_HAVI_odoo.xlsx`; si
falta la hoja Transporte se usan los valores embebidos). Para fijar
cambios permanentes, editar los `DEFAULT_*` en `havi2odoo.py`.

## Importación en Odoo

Ventas → Pedidos → ⚙️ Importar registros. Luego, en lote: Confirmar →
Crear factura. Requisitos:

- Clientes: nombres verificados contra producción el 18/08/2026.
- Productos: la celda "Producto Odoo" debe identificar unívocamente el
  producto. Las empanadas con variantes de Mercado dan "varias
  coincidencias" si se usa el nombre de plantilla → usar la referencia de
  variante (p. ej. `PA00001-ESP`) cuando estén cargadas.
- UdM `Caja 40 Uds`, `Pack 100`, `Pack 1000`, `Caja de 27` en Unidades y
  embalajes.
- Servicios de transporte (`Transporte Península`, `Transporte Barcelona`,
  `Transporte Portugal, Andorra e Islas`) dados de alta con UdM kg y
  precio por kg.

Probar primero en `muns-pruebas`.
