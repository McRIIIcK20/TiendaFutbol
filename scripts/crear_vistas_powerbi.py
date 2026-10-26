#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
crear_vistas_powerbi.py (FASE 4)
Crea las 5 vistas analíticas (QueryDefs) en access/tienda_futbol.accdb usando
DAO.DBEngine.120 (sin requerir proveedor ACE OLEDB en Python) y exporta además
las vistas a CSV en powerbi/datos_csv/ para carga directa en Power BI Desktop.
"""

import csv
import os
import sys
import win32com.client

RUTA_ACCDB = r"D:\Proyectos\TiendaFutbol\access\tienda_futbol.accdb"
CARPETA_CSV = r"D:\Proyectos\TiendaFutbol\powerbi\datos_csv"

VISTAS_SQL = {
    "vw_ventas_detalle": """
        SELECT
            v.id AS id_venta,
            b.numero_correlativo,
            v.fecha,
            Year(v.fecha) AS anio,
            Month(v.fecha) AS mes,
            v.estado,
            u.nombre AS vendedor,
            u.rol AS rol_vendedor,
            c.id AS id_cliente,
            c.rut AS rut_cliente,
            c.nombre & ' ' & c.apellido AS cliente,
            d.sku_producto AS sku,
            p.nombre AS producto,
            cat.nombre AS categoria,
            p.marca,
            p.talla,
            d.cantidad,
            p.precio_costo AS costo_unitario,
            d.precio_unitario AS precio_venta_unitario,
            d.subtotal AS ingreso_neto_linea,
            (d.cantidad * p.precio_costo) AS costo_total_linea,
            (d.subtotal - (d.cantidad * p.precio_costo)) AS margen_bruto_linea
        FROM (((((VENTAS v
            INNER JOIN DETALLE_VENTAS d ON v.id = d.id_venta)
            INNER JOIN PRODUCTOS p ON d.sku_producto = p.sku)
            INNER JOIN CATEGORIAS cat ON p.id_categoria = cat.id)
            INNER JOIN CLIENTES c ON v.id_cliente = c.id)
            INNER JOIN USUARIOS u ON v.id_usuario = u.id)
            LEFT JOIN BOLETAS b ON v.id = b.id_venta
    """,
    "vw_resumen_mensual": """
        SELECT
            Year(v.fecha) AS anio,
            Month(v.fecha) AS mes,
            Year(v.fecha) & '-' & Right('0' & CStr(Month(v.fecha)), 2) AS periodo_ym,
            COUNT(v.id) AS total_transacciones,
            SUM(v.subtotal) AS ventas_netas,
            SUM(v.iva) AS iva_total,
            SUM(v.total) AS ventas_brutas_con_iva
        FROM VENTAS v
        WHERE v.estado = 'Completada'
        GROUP BY Year(v.fecha), Month(v.fecha)
    """,
    "vw_rentabilidad_productos": """
        SELECT
            p.sku,
            p.nombre AS producto,
            c.nombre AS categoria,
            p.marca,
            p.talla,
            p.precio_costo,
            p.precio_venta,
            (p.precio_venta - p.precio_costo) AS margen_unitario,
            Round((p.precio_venta / p.precio_costo - 1) * 100, 2) AS margen_porcentaje,
            p.stock,
            p.stock_minimo,
            IIf(p.stock <= p.stock_minimo, 'Critico', 'Normal') AS estado_stock,
            (p.stock * p.precio_costo) AS valor_inventario_costo
        FROM PRODUCTOS p
        INNER JOIN CATEGORIAS c ON p.id_categoria = c.id
    """,
    "vw_movimientos_inventario": """
        SELECT
            m.id AS id_movimiento,
            m.fecha,
            m.tipo_movimiento,
            IIf(m.tipo_movimiento='E', 'Entrada', IIf(m.tipo_movimiento='S', 'Salida', 'Ajuste')) AS desc_tipo,
            m.sku_producto AS sku,
            p.nombre AS producto,
            c.nombre AS categoria,
            m.cantidad,
            u.nombre AS usuario_registro,
            prov.nombre AS proveedor,
            m.motivo
        FROM (((MOVIMIENTOS_INVENTARIO m
            INNER JOIN PRODUCTOS p ON m.sku_producto = p.sku)
            INNER JOIN CATEGORIAS c ON p.id_categoria = c.id)
            INNER JOIN USUARIOS u ON m.id_usuario = u.id)
            LEFT JOIN PROVEEDORES prov ON m.id_proveedor = prov.id
    """,
    "vw_ranking_clientes": """
        SELECT
            c.id AS id_cliente,
            c.rut,
            c.nombre & ' ' & c.apellido AS cliente,
            c.correo,
            c.fecha_registro,
            COUNT(v.id) AS cantidad_compras,
            SUM(v.subtotal) AS gasto_neto_total,
            SUM(v.total) AS gasto_total_con_iva,
            Round(AVG(v.total), 0) AS ticket_promedio,
            MAX(v.fecha) AS ultima_compra
        FROM CLIENTES c
        LEFT JOIN VENTAS v ON c.id = v.id_cliente
        GROUP BY c.id, c.rut, c.nombre, c.apellido, c.correo, c.fecha_registro
    """,
}


def abrir_bd_dao(ruta):
  """Abre la base Access mediante DAO.DBEngine.120."""
  dbe = win32com.client.Dispatch("DAO.DBEngine.120")
  return dbe.OpenDatabase(ruta)


def crear_y_exportar_vistas(ruta_accdb, carpeta_csv):
  if not os.path.exists(ruta_accdb):
    print(f"ERROR: No existe la base de datos en {ruta_accdb}")
    sys.exit(1)

  os.makedirs(carpeta_csv, exist_ok=True)
  db = abrir_bd_dao(ruta_accdb)

  print("=" * 64)
  print("FASE 4 - Creación de Vistas SQL (QueryDefs) y Exportación CSV")
  print("=" * 64)

  try:
    nombres_existentes = [q.Name for q in db.QueryDefs]
    for nombre_vista, sql in VISTAS_SQL.items():
      sql_limpio = " ".join(
          line.strip() for line in sql.strip().splitlines() if line.strip()
      )
      if nombre_vista in nombres_existentes:
        db.QueryDefs.Delete(nombre_vista)

      db.CreateQueryDef(nombre_vista, sql_limpio)

      # Ejecutar la vista recién creada y volcarla a CSV
      rs = db.OpenRecordset(nombre_vista)
      columnas = [rs.Fields(i).Name for i in range(rs.Fields.Count)]
      ruta_csv = os.path.join(carpeta_csv, f"{nombre_vista}.csv")

      n_filas = 0
      with open(ruta_csv, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(columnas)
        while not rs.EOF:
          fila = [rs.Fields(i).Value for i in range(rs.Fields.Count)]
          writer.writerow(fila)
          n_filas += 1
          rs.MoveNext()
      rs.Close()
      print(
          f"  OK | Vista '{nombre_vista}': {n_filas} filas"
          f" ({len(columnas)} columnas) -> {ruta_csv}"
      )

  finally:
    db.Close()

  print("\nRESULTADO: 5/5 vistas creadas en Access y exportadas a CSV.")


if __name__ == "__main__":
  crear_y_exportar_vistas(RUTA_ACCDB, CARPETA_CSV)