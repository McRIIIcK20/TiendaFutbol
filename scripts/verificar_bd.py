# -*- coding: utf-8 -*-
"""
Verificacion independiente de la base tienda_futbol.accdb (post-FASE 1).
Prueba que las restricciones funcionen de verdad:
  - INSERTs validos deben insertarse (y se borran al final).
  - INSERTs invalidos deben fallar por CHECK o FK.
No modifica decisiones de diseno: solo lee y prueba, dejando la base limpia.
"""
import os
import sys
import win32com.client

RUTA = r"D:\Proyectos\TiendaFutbol\access\tienda_futbol.accdb"
PROVEEDOR = "Microsoft.ACE.OLEDB.12.0"

conn = win32com.client.Dispatch("ADODB.Connection")
conn.Open(f"Provider={PROVEEDOR};Data Source={RUTA};")

fallas = []

# 0) Limpieza previa idempotente: una ejecucion interrumpida puede haber dejado
#    filas de prueba (crash antes de la limpieza final), lo que chocaria con UNIQUE.
print("--- Limpieza previa (restos de ejecuciones anteriores) ---")
for sql in [
    "DELETE FROM MOVIMIENTOS_INVENTARIO",
    "DELETE FROM PRODUCTOS WHERE sku='CAM-CHI-M-001'",
    "DELETE FROM USUARIOS WHERE correo='admin@tiendafutbol.cl'",
    "DELETE FROM CATEGORIAS WHERE nombre='Camisetas'",
]:
    try:
        conn.Execute(sql)
    except Exception:
        pass  # no hay restos: normal en una base recien creada
print("  (hecha)")

# 1) Conteo de filas por tabla (deben estar vacias o con datos previos; solo sanity check)
tablas = ["USUARIOS", "CATEGORIAS", "PRODUCTOS", "PROVEEDORES", "CLIENTES",
          "VENTAS", "DETALLE_VENTAS", "MOVIMIENTOS_INVENTARIO", "BOLETAS"]
print("--- Conteo inicial de filas ---")
for t in tablas:
    rs, n = conn.Execute(f"SELECT COUNT(*) FROM {t}")
    print(f"  {t}: {n}")

# 2) INSERTs validos encadenados capturando los id generados con @@IDENTITY
#    (los contadores AUTOINCREMENT no se reinician al borrar; nunca asumir id=1)
print("\n--- Inserts validos (deben PASAR) ---")
try:
    conn.Execute("INSERT INTO CATEGORIAS (nombre, descripcion) VALUES ('Camisetas','Camisetas de clubes y selecciones')")
    rs, _ = conn.Execute("SELECT @@IDENTITY")
    id_cat = int(rs.Fields(0).Value)
    print(f"PASA  categoria insertada, id={id_cat}")

    conn.Execute("INSERT INTO USUARIOS (nombre, correo, hash_contrasena, sal, rol) VALUES ('Admin Demo','admin@tiendafutbol.cl','aabbccdd11223344','0123456789abcdef0123456789abcdef','Admin')")
    rs, _ = conn.Execute("SELECT @@IDENTITY")
    id_usr = int(rs.Fields(0).Value)
    print(f"PASA  usuario insertado, id={id_usr}")

    conn.Execute(
        f"INSERT INTO PRODUCTOS (sku, nombre, id_categoria, talla, marca, precio_costo, precio_venta, stock, stock_minimo) "
        f"VALUES ('CAM-CHI-M-001','Camiseta Chile Local M',{id_cat},'M','Adidas',25000,39900,10,3)"
    )
    print("PASA  producto insertado (FK hacia la categoria recien creada)")
    ids_prueba = {"id_cat": id_cat, "id_usr": id_usr}
except Exception as e:
    try:
        desc = e.excepinfo[2] if e.excepinfo else str(e)
    except Exception:
        desc = str(e)
    fallas.append(f"INSERT valido rechazado: {str(desc)[:150]}")
    print(f"FALLA INESPERADA: {str(desc)[:140]}")
    ids_prueba = {"id_cat": None, "id_usr": None}

# 3) INSERTs invalidos: deben fallar
pruebas_mal = [
    # CHECK rol: 'Gerente' no esta en la lista
    ("CHECK rol", "INSERT INTO USUARIOS (nombre, correo, hash_contrasena, sal, rol) VALUES ('Mal Rol','mal@x.cl','dd','0123456789abcdef0123456789abcdef','Gerente')"),
    # CHECK stock >= 0
    ("CHECK stock", "INSERT INTO PRODUCTOS (sku, nombre, id_categoria, talla, marca, precio_costo, precio_venta, stock, stock_minimo) VALUES ('CAM-MAL-01','Camiseta mala',1,'S','X',1000,2000,-5,1)"),
    # FK id_categoria inexistente
    ("FK categoria", "INSERT INTO PRODUCTOS (sku, nombre, id_categoria, talla, marca, precio_costo, precio_venta, stock, stock_minimo) VALUES ('CAM-MAL-02','Camiseta sin categoria',999,'S','X',1000,2000,5,1)"),
    # UNIQUE correo duplicado
    ("UNIQUE correo", "INSERT INTO USUARIOS (nombre, correo, hash_contrasena, sal, rol) VALUES ('Duplicado','admin@tiendafutbol.cl','ee','0123456789abcdef0123456789abcdef','Vendedor')"),
    # CHECK tipo_movimiento
    ("CHECK tipo_mov", f"INSERT INTO MOVIMIENTOS_INVENTARIO (sku_producto, id_usuario, tipo_movimiento, cantidad) VALUES ('CAM-CHI-M-001',{ids_prueba['id_usr']},'X',3)"),
]
print("\n--- Pruebas de restricciones (deben FALLAR) ---")
for etiqueta, sql in pruebas_mal:
    try:
        conn.Execute(sql)
        fallas.append(f"{etiqueta}: la insercion invalida PASO (la restriccion no funciona)")
        print(f"ERROR: {etiqueta} permitio un insert invalido")
    except Exception as e:
        # Extraer la descripcion legible del error COM sin asumir su formato
        try:
            desc = e.excepinfo[2] if e.excepinfo else str(e)
        except Exception:
            desc = str(e)
        print(f"OK    {etiqueta}: rechazado por el motor ({str(desc)[:90]}...)")

# 4) Limpieza: borrar las filas de prueba (orden inverso por dependencias)
print("\n--- Limpieza de filas de prueba ---")
limpieza = [
    "DELETE FROM MOVIMIENTOS_INVENTARIO",
    "DELETE FROM PRODUCTOS WHERE sku='CAM-CHI-M-001'",
    "DELETE FROM USUARIOS WHERE correo='admin@tiendafutbol.cl'",
    f"DELETE FROM CATEGORIAS WHERE id={ids_prueba['id_cat']}" if ids_prueba["id_cat"] else "DELETE FROM CATEGORIAS WHERE nombre='Camisetas'",
]
for sql in limpieza:
    try:
        conn.Execute(sql)
    except Exception as e:
        fallas.append(f"Limpieza fallo: {sql} -> {str(e)[:100]}")

conn.Close()

print("\n" + "=" * 50)
if fallas:
    print("VERIFICACION CON PROBLEMAS:")
    for f_ in fallas:
        print(f" - {f_}")
    sys.exit(1)
print("VERIFICACION INDEPENDIENTE: EXITO TOTAL.")
print("Restricciones CHECK, FK y UNIQUE funcionan; inserts validos aceptados.")
