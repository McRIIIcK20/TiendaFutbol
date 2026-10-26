# FASE 4 — Vistas Analíticas, Medidas DAX y Diseño Visual en Power BI

## 1. Arquitectura de Datos para Power BI

El script `scripts/crear_vistas_powerbi.py` genera **5 vistas analíticas (`QueryDefs`)** directamente dentro de `access/tienda_futbol.accdb` y, además, las exporta como archivos CSV con codificación `UTF-8-SIG` en `powerbi/datos_csv/` para que puedas cargarlas en Power BI Desktop en segundos (incluso sin instalar el conector OLEDB de Access):

| Vista / Archivo CSV | Filas esperadas | Propósito en el Modelo de Datos |
| :--- | :--- | :--- |
| **`vw_ventas_detalle`** | `2.146` líneas | **Tabla de Hechos Principal (Fact Ventas)**: combina cabecera de venta, correlativo de boleta, vendedor, cliente, producto, categoría, marca, cantidad, costo unitario, precio unitario, ingreso neto, costo total y margen bruto por línea |
| **`vw_resumen_mensual`** | `12–13` meses | Evolución mensual de transacciones, ventas netas, IVA (19%) y ventas totales (muestra el pico estacional de diciembre con `164` ventas) |
| **`vw_rentabilidad_productos`** | `60` SKUs | Catálogo maestro con margen unitario ($), margen porcentual (`20%–60%`), stock actual, stock mínimo, bandera `estado_stock` (`Critico` / `Normal`) y valorización de inventario |
| **`vw_movimientos_inventario`** | `2.367` movs. | Historial completo de Kardex (`Entrada`, `Salida`, `Ajuste`) con proveedor, usuario y motivo |
| **`vw_ranking_clientes`** | `150` clientes | Frecuencia de compra, gasto neto acumulado, gasto con IVA, ticket promedio y fecha de última compra |

## 2. Tabla Calendario y Medidas DAX (Listas para Copiar y Pegar)

En **Power BI Desktop**, ve a **Modelado $\rightarrow$ Nueva tabla** y pega esta definición para la tabla calendario:

```dax
Calendario = 
ADDCOLUMNS (
    CALENDAR ( MIN ( vw_ventas_detalle[fecha] ), MAX ( vw_ventas_detalle[fecha] ) ),
    "Anio", YEAR ( [Date] ),
    "MesNum", MONTH ( [Date] ),
    "MesNombre", FORMAT ( [Date], "mmm", "es-CL" ),
    "Periodo", FORMAT ( [Date], "yyyy-mm" ),
    "Trimestre", "T" & FORMAT ( [Date], "q" ),
    "DiaSemana", FORMAT ( [Date], "dddd", "es-CL" )
)
```

Relaciona `Calendario[Date]` (1) $\rightarrow$ `vw_ventas_detalle[fecha]` (*) y crea estas **10 medidas DAX** en una tabla o carpeta de medidas:

```dax
// 1. Ingresos Netos (sin IVA)
Ventas Netas ($) = SUM ( vw_ventas_detalle[ingreso_neto_linea] )

// 2. Ingresos Brutos (con IVA 19%)
Ventas con IVA ($) = [Ventas Netas ($)] * 1.19

// 3. Costo Total de lo Vendido (COGS)
Costo de Ventas ($) = SUM ( vw_ventas_detalle[costo_total_linea] )

// 4. Margen Bruto en Pesos
Margen Bruto ($) = [Ventas Netas ($)] - [Costo de Ventas ($)]

// 5. Margen Bruto Porcentual (%)
Margen Bruto (%) = DIVIDE ( [Margen Bruto ($)], [Ventas Netas ($)], 0 )

// 6. Cantidad de Ventas (Boletas emitidas)
Total Transacciones = DISTINCTCOUNT ( vw_ventas_detalle[id_venta] )

// 7. Ticket Promedio Neto por Venta
Ticket Promedio ($) = DIVIDE ( [Ventas Netas ($)], [Total Transacciones], 0 )

// 8. Total de Unidades Vendidas
Unidades Vendidas = SUM ( vw_ventas_detalle[cantidad] )

// 9. Crecimiento Mensual (MoM %)
Ventas Mes Anterior ($) = CALCULATE ( [Ventas Netas ($)], DATEADD ( Calendario[Date], -1, MONTH ) )

Crecimiento MoM (%) = DIVIDE ( [Ventas Netas ($)] - [Ventas Mes Anterior ($)], [Ventas Mes Anterior ($)], 0 )

// 10. Alertas de Stock Crítico (SKUs con stock <= stock_minimo)
SKUs en Stock Critico = 
CALCULATE (
    COUNTROWS ( vw_rentabilidad_productos ),
    vw_rentabilidad_productos[estado_stock] = "Critico"
)
```

## 3. Paleta Visual y Diseño Ejecutivo del Dashboard (16:9)

Para que el dashboard en Power BI tenga exactamente la misma identidad visual moderna que los `UserForms` y la Boleta Electrónica de Excel:

| Elemento Visual | Código Hex | Uso en el Lienzo de Power BI |
| :--- | :--- | :--- |
| **Fondo del Lienzo** | `#F8FAFC` | Fondo general gris pizarra claro (transparencia 0%) |
| **Encabezado / Tarjetas KPI** | `#0F172A` | Azul marino profundo para la barra superior y títulos principales |
| **Acento Primario (Éxito / Margen)** | `#10B981` | Verde cancha esmeralda para ingresos, margen bruto y líneas de tendencia |
| **Acento Secundario (Volumen)** | `#2563EB` | Azul real para barras de categorías (`Camisetas`, `Botines`, `Balones`) |
| **Alerta de Quiebre de Stock** | `#EF4444` / `#F59E0B` | Rojo coral y ámbar para formato condicional en productos con `estado_stock = "Critico"` |

### Distribución recomendada del Lienzo (3 bandas horizontales)

1. **Banda Superior (Cabecera + 5 Tarjetas KPI)**:
   - Título a la izquierda (`TIENDA DE FÚTBOL — DASHBOARD COMERCIAL Y DE INVENTARIO`) + Segmentadores (*Slicers*) de `Año-Mes`, `Categoría` y `Marca`.
   - **5 Tarjetas KPI**: `Ventas Netas ($)`, `Margen Bruto ($)` (con subtítulo `Margen Bruto (%)`), `Total Transacciones` (`850`), `Ticket Promedio ($)` y `SKUs en Stock Crítico`.
2. **Banda Central (Tendencia Estacional + Rentabilidad por Categoría)**:
   - **Izquierda (60% del ancho) — Gráfico de líneas y columnas agrupadas**: Eje X = `Calendario[Periodo]`, Columnas = `Ventas Netas ($)`, Línea = `Margen Bruto (%)` (destaca claramente el pico estacional de diciembre con 164 ventas).
   - **Derecha (40% del ancho) — Gráfico de barras horizontales**: Eje Y = `categoria`, Eje X = `Ventas Netas ($)` y *Tooltip* con `Unidades Vendidas` y `Margen Bruto (%)`.
3. **Banda Inferior (Control de Stock Crítico + Top 10 Clientes)**:
   - **Izquierda (55% del ancho) — Matriz / Tabla de Control de Inventario (`vw_rentabilidad_productos`)**: Columnas `sku`, `producto`, `categoria`, `stock`, `stock_minimo`, `estado_stock` y `margen_porcentaje`, aplicando formato condicional de fondo rojo suave a las filas donde `estado_stock = "Critico"`.
   - **Derecha (45% del ancho) — Tabla Top 10 Clientes (`vw_ranking_clientes`)**: `cliente`, `rut`, `cantidad_compras`, `gasto_total_con_iva` y `ticket_promedio`.
