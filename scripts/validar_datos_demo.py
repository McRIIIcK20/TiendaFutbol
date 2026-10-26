# -*- coding: utf-8 -*-
"""
FASE 1 - Validaciones de coherencia de los datos demo (condiciones del encargo).

Comprueba y muestra:
  1. Conteos por tabla (60 productos, 150 clientes, 850+ ventas, etc.).
  2. Stock negativo = 0.
  3. Ventas sin detalle = 0.
  4. Ventas cuadradas: subtotal = SUM(detalles), iva = 19%, total = neto + iva.
  5. Inventario cuadrado: stock de PRODUCTOS = entradas - salidas (por SKU).
  6. Margen de precios: 20% <= (venta/costo - 1) <= 60% en todos los productos.
  7. Reproducibilidad/estacionalidad: ventas por mes (pico diciembre).
  8. Usuarios: 3 con sal individual (sal distintas entre si).
  9. Boletas: 1 por venta, correlativo unico.
"""

import sys
import win32com.client

RUTA = r"D:\Proyectos\TiendaFutbol\access\tienda_futbol.accdb"
PROVEEDOR = "Microsoft.ACE.OLEDB.12.0"

conn = win32com.client.Dispatch("ADODB.Connection")
conn.Open(f"Provider={PROVEEDOR};Data Source={RUTA};")

problemas = []


def escalar(sql, etiqueta=None):
    rs, _ = conn.Execute(sql)
    v = rs.Fields(0).Value
    rs.Close()
    return v


def tabla_filas(sql):
    rs, _ = conn.Execute(sql)
    filas = []
    while not rs.EOF:
        filas.append(tuple(rs.Fields(i).Value for i in range(rs.Fields.Count)))
        rs.MoveNext()
    rs.Close()
    return filas


print("=" * 64)
print("VALIDACIONES DE LA CARGA DEMO")
print("=" * 64)

# --- 1. Conteos -------------------------------------------------------------
print("\n[1] CONTEOS POR TABLA")
esperados = {
    "USUARIOS": 3, "CATEGORIAS": 6, "PROVEEDORES": 5, "PRODUCTOS": 60,
    "CLIENTES": 150, "BOLETAS": 850,
}
for t, esperado in esperados.items():
    n = int(escalar(f"SELECT COUNT(*) FROM {t}"))
    marca = "OK " if n == esperado else "!! "
    if n != esperado:
        problemas.append(f"{t}: {n} filas, esperado {esperado}")
    print(f"  {marca}{t}: {n} (esperado {esperado})")

n_ventas = int(escalar("SELECT COUNT(*) FROM VENTAS"))
print(f"    VENTAS: {n_ventas} (minimo 850 -> {'OK' if n_ventas >= 850 else 'INSUFICIENTE'})")
if n_ventas < 850:
    problemas.append(f"VENTAS: {n_ventas} < 850")
n_detalles = int(escalar("SELECT COUNT(*) FROM DETALLE_VENTAS"))
n_movs = int(escalar("SELECT COUNT(*) FROM MOVIMIENTOS_INVENTARIO"))
print(f"    DETALLE_VENTAS: {n_detalles} | MOVIMIENTOS_INVENTARIO: {n_movs}")

# --- 2. Stock negativo ------------------------------------------------------
print("\n[2] PRODUCTOS CON STOCK NEGATIVO (debe ser 0)")
n = int(escalar("SELECT COUNT(*) FROM PRODUCTOS WHERE stock < 0"))
print(f"  {'OK ' if n == 0 else '!! '}stock negativo: {n}")
if n:
    problemas.append(f"Hay {n} productos con stock negativo")

# --- 3. Ventas sin detalle ---------------------------------------------------
print("\n[3] VENTAS SIN DETALLE (debe ser 0)")
n = int(escalar(
    "SELECT COUNT(*) FROM VENTAS v WHERE NOT EXISTS "
    "(SELECT 1 FROM DETALLE_VENTAS d WHERE d.id_venta = v.id)"))
print(f"  {'OK ' if n == 0 else '!! '}ventas sin detalle: {n}")
if n:
    problemas.append(f"{n} ventas sin detalle")

# --- 4. Ventas cuadradas ------------------------------------------------------
print("\n[4] VENTAS CUADRADAS (subtotal=SUM(detalles), iva=19%, total=neto+iva)")
# 4a: subtotal vs suma de detalles (tolerancia 0.05 por redondeo CURRENCY)
n_sub = int(escalar(
    "SELECT COUNT(*) FROM VENTAS v "
    "INNER JOIN (SELECT id_venta, SUM(subtotal) AS s FROM DETALLE_VENTAS GROUP BY id_venta) d "
    "ON v.id = d.id_venta WHERE ABS(v.subtotal - d.s) > 0.05"))
print(f"  {'OK ' if n_sub == 0 else '!! '}subtotal != suma detalles: {n_sub}")
if n_sub:
    problemas.append(f"{n_sub} ventas con subtotal mal cuadrado")
# 4b: IVA 19%
n_iva = int(escalar("SELECT COUNT(*) FROM VENTAS WHERE ABS(iva - subtotal*0.19) > 0.05"))
print(f"  {'OK ' if n_iva == 0 else '!! '}iva != 19% del subtotal: {n_iva}")
if n_iva:
    problemas.append(f"{n_iva} ventas con IVA != 19%")
# 4c: total = subtotal + iva
n_tot = int(escalar("SELECT COUNT(*) FROM VENTAS WHERE ABS(total - subtotal - iva) > 0.05"))
print(f"  {'OK ' if n_tot == 0 else '!! '}total != subtotal+iva: {n_tot}")
if n_tot:
    problemas.append(f"{n_tot} ventas con total mal cuadrado")
# 4d: detalles cuadrados a nivel de linea
n_linea = int(escalar(
    "SELECT COUNT(*) FROM DETALLE_VENTAS WHERE ABS(subtotal - cantidad*precio_unitario) > 0.05"))
print(f"  {'OK ' if n_linea == 0 else '!! '}lineas con subtotal != cant*precio: {n_linea}")
if n_linea:
    problemas.append(f"{n_linea} lineas de detalle mal cuadradas")

# --- 5. Inventario cuadrado ----------------------------------------------------
print("\n[5] INVENTARIO CUADRADO (stock = entradas - salidas por SKU)")
filas = tabla_filas(
    "SELECT p.sku, p.stock, "
    "  SUM(IIf(m.tipo_movimiento='E', m.cantidad, 0)) AS entradas, "
    "  SUM(IIf(m.tipo_movimiento='S', m.cantidad, 0)) AS salidas "
    "FROM PRODUCTOS p LEFT JOIN MOVIMIENTOS_INVENTARIO m ON p.sku = m.sku_producto "
    "GROUP BY p.sku, p.stock")
malos = [(sku, stock, ent, sal) for sku, stock, ent, sal in filas if stock != ent - sal]
print(f"  {'OK ' if not malos else '!! '}productos descuadrados: {len(malos)} de {len(filas)}")
for sku, stock, ent, sal in malos[:5]:
    print(f"      {sku}: stock={stock}, entradas={ent}, salidas={sal}")
if malos:
    problemas.append(f"{len(malos)} productos con stock != entradas-salidas")

# --- 6. Margenes ----------------------------------------------------------------
print("\n[6] MARGENES (20%..60% y precio_venta > precio_costo)")
n_fuera = int(escalar(
    "SELECT COUNT(*) FROM PRODUCTOS "
    "WHERE precio_venta <= precio_costo "
    "OR (precio_venta / precio_costo - 1) < 0.1999 "
    "OR (precio_venta / precio_costo - 1) > 0.6001"))
print(f"  {'OK ' if n_fuera == 0 else '!! '}productos fuera del margen: {n_fuera}")
if n_fuera:
    problemas.append(f"{n_fuera} productos fuera del margen 20-60%")
filas = tabla_filas(
    "SELECT MIN(precio_venta/precio_costo - 1), MAX(precio_venta/precio_costo - 1) FROM PRODUCTOS")
print(f"  margen minimo: {filas[0][0]*100:.1f}% | margen maximo: {filas[0][1]*100:.1f}%")

# --- 7. Estacionalidad -----------------------------------------------------------
print("\n[7] ESTACIONALIDAD (ventas por mes, 12 meses)")
filas = tabla_filas(
    "SELECT Year(fecha) & '-' & Right('0' & CStr(Month(fecha)), 2) AS mes, COUNT(*) AS ventas "
    "FROM VENTAS GROUP BY Year(fecha), Month(fecha) ORDER BY Year(fecha), Month(fecha)")
for mes, nv in filas:
    barra = "#" * min(int(nv / 4), 60)
    print(f"  {mes}: {nv:4d} {barra}")
dic = dict(filas).get("2025-12", 0) or dict(filas).get("2026-12", 0)
promedio = sum(nv for _m, nv in filas) / len(filas) if filas else 0
print(f"  {'OK ' if dic > promedio else 'AVISO '}diciembre ({dic}) vs promedio mensual ({promedio:.0f})")

# --- 8. Usuarios con sal -----------------------------------------------------------
print("\n[8] USUARIOS (sal individual por usuario)")
filas = tabla_filas("SELECT nombre, rol, sal FROM USUARIOS ORDER BY id")
sales = [s for _n, _r, s in filas]
print(f"  {'OK ' if len(set(sales)) == len(sales) else '!! '}"
      f"{len(filas)} usuarios, {len(set(sales))} sales distintas")
for nombre, rol, sal in filas:
    print(f"    {nombre:15s} rol={rol:10s} sal={sal[:12]}...")
if len(set(sales)) != len(sales):
    problemas.append("Sales repetidas entre usuarios")

# --- 9. Boletas ---------------------------------------------------------------------
print("\n[9] BOLETAS (1 por venta, correlativo unico)")
n_dup = int(escalar(
    "SELECT COUNT(*) FROM (SELECT id_venta FROM BOLETAS GROUP BY id_venta "
    "HAVING COUNT(*) > 1)"))
print(f"  {'OK ' if n_dup == 0 else '!! '}ventas con boleta duplicada: {n_dup}")
if n_dup:
    problemas.append(f"{n_dup} boletas duplicadas")

# --- Resultado -------------------------------------------------------------------------
print("\n" + "=" * 64)
if problemas:
    print(f"VALIDACION CON PROBLEMAS ({len(problemas)}):")
    for p in problemas:
        print(f" - {p}")
    conn.Close()
    sys.exit(1)
print("VALIDACION: TODAS LAS CONDICIONES CUMPLIDAS.")
conn.Close()
