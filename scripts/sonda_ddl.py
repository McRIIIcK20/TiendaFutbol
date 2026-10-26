# -*- coding: utf-8 -*-
"""
Sonda de capacidades DDL de Access/Jet-ACE.
Prueba qué sintaxis SQL acepta el motor (vía ADOX + ADO con pywin32)
antes de reescribir create_access_db.py. Diagnóstico de la FASE 1.
"""
import os
import sys
import win32com.client

# Carpeta temporal de trabajo (scratch de Hermes o TEMP del sistema)
TMP = os.environ.get("TMPDIR") or os.environ.get("TEMP") or r"C:\Users\mc18m\AppData\Local\Temp"
scratch = os.path.join(TMP, "sonda_ddl.accdb")

# Partir limpio: borrar la base de prueba si quedó de una sonda anterior
if os.path.exists(scratch):
    try:
        os.remove(scratch)
        print(f"Base de prueba anterior eliminada: {scratch}")
    except Exception as e:
        print(f"No se pudo borrar la base anterior (¿bloqueada?): {e}")
        sys.exit(1)

# 1) Crear el archivo .accdb con ADOX probando los proveedores ACE disponibles
cat = win32com.client.Dispatch("ADOX.Catalog")
proveedor = None
for p in ["Microsoft.ACE.OLEDB.12.0", "Microsoft.ACE.OLEDB.16.0"]:
    try:
        cat.Create(f"Provider={p};Data Source={scratch};")
        proveedor = p
        print(f"PROVEEDOR OK   : {p}")
        break
    except Exception as e:
        print(f"PROVEEDOR FALLA: {p} -> {str(e)[:100]}")

if not proveedor:
    print("SIN PROVEEDOR ACE DISPONIBLE: no se puede crear la base por ADOX.")
    sys.exit(2)

# Cerrar la conexión que ADOX dejó abierta para liberar el bloqueo del archivo
try:
    cat.ActiveConnection.Close()
    print("Conexión de ADOX cerrada.")
except Exception as e:
    print(f"Aviso: no se pudo cerrar la conexión de ADOX: {str(e)[:100]}")

# 2) Reabrir con ADO y probar sentencias DDL una por una
conn = win32com.client.Dispatch("ADODB.Connection")
conn.Open(f"Provider={proveedor};Data Source={scratch};")
print("Conexión ADO abierta para pruebas DDL.")

pruebas = [
    ("AUTOINCREMENT + PK",        "CREATE TABLE t01 (id AUTOINCREMENT PRIMARY KEY, nombre TEXT(50) NOT NULL)"),
    ("LONG",                      "CREATE TABLE t02 (id AUTOINCREMENT PRIMARY KEY, id_cat LONG NOT NULL)"),
    ("CURRENCY + DEFAULT 0",      "CREATE TABLE t03 (id AUTOINCREMENT PRIMARY KEY, precio CURRENCY NOT NULL DEFAULT 0)"),
    ("DATE + DEFAULT DATE()",     "CREATE TABLE t04 (id AUTOINCREMENT PRIMARY KEY, fecha DATE DEFAULT DATE())"),
    ("DATE + DEFAULT NOW()",      "CREATE TABLE t05 (id AUTOINCREMENT PRIMARY KEY, fecha DATE DEFAULT NOW())"),
    ("TEXT + DEFAULT literal",    "CREATE TABLE t06 (id AUTOINCREMENT PRIMARY KEY, estado TEXT(20) DEFAULT 'Pendiente')"),
    ("CHECK con IN",              "CREATE TABLE t07 (id AUTOINCREMENT PRIMARY KEY, rol TEXT(20) NOT NULL, CONSTRAINT chk_rol CHECK (rol IN ('Admin','Vendedor')))"),
    ("CHECK con >=",              "CREATE TABLE t08 (id AUTOINCREMENT PRIMARY KEY, stock LONG NOT NULL, CONSTRAINT chk_stock CHECK (stock >= 0))"),
    ("FOREIGN KEY",               "CREATE TABLE t09 (id AUTOINCREMENT PRIMARY KEY, id_cat LONG NOT NULL, CONSTRAINT fk_probe FOREIGN KEY (id_cat) REFERENCES t02(id))"),
    ("CHAR(1) + DEFAULT",         "CREATE TABLE t10 (id AUTOINCREMENT PRIMARY KEY, flag CHAR(1) DEFAULT 'N')"),
    ("VARCHAR + UNIQUE",          "CREATE TABLE t11 (id AUTOINCREMENT PRIMARY KEY, rut VARCHAR(12) NOT NULL UNIQUE)"),
    ("DEFAULT + CHECK juntos",    "CREATE TABLE t12 (id AUTOINCREMENT PRIMARY KEY, estado TEXT(20) NOT NULL DEFAULT 'Pendiente', CONSTRAINT chk_est CHECK (estado IN ('Pendiente','Completada')))"),
    ("CREATE INDEX",              "CREATE INDEX idx_probe ON t01 (nombre)"),
]

resultados = {}
for nombre, sql in pruebas:
    try:
        conn.Execute(sql)
        resultados[nombre] = True
        print(f"PASA  : {nombre}")
    except Exception as e:
        resultados[nombre] = False
        # Extraer la descripción legible del error COM
        detalle = str(e).replace("\n", " ")
        print(f"FALLA : {nombre} -> {detalle[:140]}")

conn.Close()

# 3) Limpieza: intentar borrar la base de prueba (si queda bloqueada, se reporta)
try:
    os.remove(scratch)
    print("Base de prueba eliminada (limpieza OK).")
except Exception as e:
    print(f"Aviso: la base de prueba quedó en {scratch} ({str(e)[:80]})")

print("SONDA COMPLETA")
