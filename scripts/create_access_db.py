# -*- coding: utf-8 -*-
"""
FASE 1 - Creacion de la base de datos Access para TiendaFutbol.

Estrategia (validada con scripts/sonda_ddl.py):
  - ADOX.Catalog para crear el archivo .accdb (no requiere Access instalado).
  - ADODB.Connection para ejecutar el DDL (modo ANSI-92: soporta DEFAULT y CHECK).
  - El DDL antiguo fallaba por "NUMBER(LONG)" (no existe en Access SQL; es LONG)
    y por la via DAO de Access.Application.

Al terminar, el script se auto-verifica: lista las 9 tablas y sus llaves
primarias/foraneas leidas desde el motor, no desde el codigo.

Uso:
    python create_access_db.py [ruta_base]
    (por defecto: D:/Proyectos/TiendaFutbol/access/tienda_futbol.accdb)
"""

import os
import sys
import win32com.client

# --- Configuracion ---------------------------------------------------------
RUTA_BASE_DEFAULT = r"D:\Proyectos\TiendaFutbol\access\tienda_futbol.accdb"
PROVEEDOR = "Microsoft.ACE.OLEDB.12.0"

# --- DDL de las 9 tablas (nombres en espanol, sin tildes ni espacios) ------
# Orden importado: las tablas referenciadas deben existir antes que sus FK.
# NOTA: los CHECK van como CONSTRAINT nombradas a nivel de tabla; el CHECK
# inline en la columna produce "Error de sintaxis en la definicion del campo".
DDL_TABLAS = [
    # 1. Usuarios del sistema (hash SHA-256 con sal individual: hash=SHA256(sal+clave))
    """
    CREATE TABLE USUARIOS (
        id AUTOINCREMENT PRIMARY KEY,
        nombre TEXT(100) NOT NULL,
        correo TEXT(255) NOT NULL UNIQUE,
        hash_contrasena TEXT(255) NOT NULL,
        sal TEXT(64) NOT NULL,
        rol TEXT(20) NOT NULL,
        CONSTRAINT chk_usuarios_rol CHECK (rol IN ('Admin','Vendedor','Cliente'))
    )
    """,
    # 2. Categorias de productos
    """
    CREATE TABLE CATEGORIAS (
        id AUTOINCREMENT PRIMARY KEY,
        nombre TEXT(100) NOT NULL UNIQUE,
        descripcion TEXT(255)
    )
    """,
    # 3. Catalogo de productos (SKU como llave natural)
    """
    CREATE TABLE PRODUCTOS (
        sku TEXT(50) PRIMARY KEY,
        nombre TEXT(200) NOT NULL,
        id_categoria LONG NOT NULL,
        talla TEXT(50),
        marca TEXT(100),
        precio_costo CURRENCY NOT NULL,
        precio_venta CURRENCY NOT NULL,
        stock LONG NOT NULL DEFAULT 0,
        stock_minimo LONG NOT NULL DEFAULT 0,
        CONSTRAINT fk_productos_categoria FOREIGN KEY (id_categoria) REFERENCES CATEGORIAS(id),
        CONSTRAINT chk_productos_costo CHECK (precio_costo >= 0),
        CONSTRAINT chk_productos_venta CHECK (precio_venta >= 0),
        CONSTRAINT chk_productos_stock CHECK (stock >= 0)
    )
    """,
    # 4. Proveedores
    """
    CREATE TABLE PROVEEDORES (
        id AUTOINCREMENT PRIMARY KEY,
        nombre TEXT(200) NOT NULL,
        contacto TEXT(100),
        telefono TEXT(50),
        correo TEXT(255),
        direccion TEXT(255)
    )
    """,
    # 5. Clientes (RUT chileno con digito verificador, unico)
    """
    CREATE TABLE CLIENTES (
        id AUTOINCREMENT PRIMARY KEY,
        nombre TEXT(100) NOT NULL,
        apellido TEXT(100) NOT NULL,
        correo TEXT(255) NOT NULL UNIQUE,
        telefono TEXT(50),
        direccion TEXT(255),
        rut VARCHAR(12) NOT NULL UNIQUE,
        fecha_registro DATE NOT NULL DEFAULT DATE()
    )
    """,
    # 6. Ventas (cabecera)
    """
    CREATE TABLE VENTAS (
        id AUTOINCREMENT PRIMARY KEY,
        id_usuario LONG NOT NULL,
        id_cliente LONG NOT NULL,
        fecha DATE NOT NULL DEFAULT DATE(),
        subtotal CURRENCY NOT NULL DEFAULT 0,
        iva CURRENCY NOT NULL DEFAULT 0,
        total CURRENCY NOT NULL DEFAULT 0,
        estado TEXT(20) NOT NULL DEFAULT 'Pendiente',
        CONSTRAINT fk_ventas_usuario FOREIGN KEY (id_usuario) REFERENCES USUARIOS(id),
        CONSTRAINT fk_ventas_cliente FOREIGN KEY (id_cliente) REFERENCES CLIENTES(id),
        CONSTRAINT chk_ventas_estado CHECK (estado IN ('Pendiente','Completada','Anulada'))
    )
    """,
    # 7. Detalle de ventas (lineas de cada venta)
    """
    CREATE TABLE DETALLE_VENTAS (
        id AUTOINCREMENT PRIMARY KEY,
        id_venta LONG NOT NULL,
        sku_producto TEXT(50) NOT NULL,
        cantidad LONG NOT NULL,
        precio_unitario CURRENCY NOT NULL,
        subtotal CURRENCY NOT NULL DEFAULT 0,
        CONSTRAINT fk_detalle_venta FOREIGN KEY (id_venta) REFERENCES VENTAS(id),
        CONSTRAINT fk_detalle_producto FOREIGN KEY (sku_producto) REFERENCES PRODUCTOS(sku),
        CONSTRAINT chk_detalle_cantidad CHECK (cantidad > 0),
        CONSTRAINT chk_detalle_precio CHECK (precio_unitario >= 0)
    )
    """,
    # 8. Movimientos de inventario (E: entrada, S: salida, A: ajuste)
    """
    CREATE TABLE MOVIMIENTOS_INVENTARIO (
        id AUTOINCREMENT PRIMARY KEY,
        sku_producto TEXT(50) NOT NULL,
        id_usuario LONG NOT NULL,
        id_proveedor LONG,
        tipo_movimiento CHAR(1) NOT NULL,
        cantidad LONG NOT NULL,
        fecha DATE NOT NULL DEFAULT DATE(),
        motivo TEXT(255),
        CONSTRAINT fk_mov_producto FOREIGN KEY (sku_producto) REFERENCES PRODUCTOS(sku),
        CONSTRAINT fk_mov_usuario FOREIGN KEY (id_usuario) REFERENCES USUARIOS(id),
        CONSTRAINT fk_mov_proveedor FOREIGN KEY (id_proveedor) REFERENCES PROVEEDORES(id),
        CONSTRAINT chk_mov_tipo CHECK (tipo_movimiento IN ('E','S','A')),
        CONSTRAINT chk_mov_cantidad CHECK (cantidad <> 0)
    )
    """,
    # 9. Boletas (PDF y estado de envio por correo)
    """
    CREATE TABLE BOLETAS (
        id AUTOINCREMENT PRIMARY KEY,
        id_venta LONG NOT NULL UNIQUE,
        numero_correlativo TEXT(20) NOT NULL UNIQUE,
        ruta_pdf TEXT(255),
        fecha_emision DATE DEFAULT DATE(),
        enviada_correo CHAR(1) DEFAULT 'N',
        CONSTRAINT fk_boletas_venta FOREIGN KEY (id_venta) REFERENCES VENTAS(id),
        CONSTRAINT chk_boletas_enviada CHECK (enviada_correo IN ('S','N'))
    )
    """,
]

# Indices para acelerar las consultas mas frecuentes (reportes y Power BI)
DDL_INDICES = [
    "CREATE INDEX idx_ventas_fecha ON VENTAS (fecha)",
    "CREATE INDEX idx_movimientos_fecha ON MOVIMIENTOS_INVENTARIO (fecha)",
    "CREATE INDEX idx_detalle_venta ON DETALLE_VENTAS (id_venta, sku_producto)",
    "CREATE INDEX idx_productos_categoria ON PRODUCTOS (id_categoria)",
    "CREATE INDEX idx_boletas_enviada ON BOLETAS (enviada_correo)",
    "CREATE INDEX idx_clientes_rut ON CLIENTES (rut)",
]

# Verificacion esperada: columna PK por tabla + lista de FKs (nombre, tabla referenciada)
LLAVES_ESPERADAS = {
    "USUARIOS":               {"pk": "id",  "fks": []},
    "CATEGORIAS":             {"pk": "id",  "fks": []},
    "PRODUCTOS":              {"pk": "sku", "fks": [("fk_productos_categoria", "CATEGORIAS")]},
    "PROVEEDORES":            {"pk": "id",  "fks": []},
    "CLIENTES":               {"pk": "id",  "fks": []},
    "VENTAS":                 {"pk": "id",  "fks": [("fk_ventas_usuario", "USUARIOS"),
                                                     ("fk_ventas_cliente", "CLIENTES")]},
    "DETALLE_VENTAS":         {"pk": "id",  "fks": [("fk_detalle_venta", "VENTAS"),
                                                     ("fk_detalle_producto", "PRODUCTOS")]},
    "MOVIMIENTOS_INVENTARIO": {"pk": "id",  "fks": [("fk_mov_producto", "PRODUCTOS"),
                                                     ("fk_mov_usuario", "USUARIOS"),
                                                     ("fk_mov_proveedor", "PROVEEDORES")]},
    "BOLETAS":                {"pk": "id",  "fks": [("fk_boletas_venta", "VENTAS")]},
}


def eliminar_base_si_existe(ruta):
    """Borra la base anterior para poder reejecutar el script (idempotente)."""
    if not os.path.exists(ruta):
        return
    # El archivo de bloqueo (.laccdb) indica que Access/ADO tiene la base abierta
    bloqueo = os.path.splitext(ruta)[0] + ".laccdb"
    if os.path.exists(bloqueo):
        print(f"ERROR: la base parece estar abierta (existe {bloqueo}). Cierre Access y reejecute.")
        sys.exit(1)
    os.remove(ruta)
    print(f"Base anterior eliminada: {ruta}")


def crear_base(ruta):
    """Crea el archivo .accdb vacio con ADOX."""
    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    cat = win32com.client.Dispatch("ADOX.Catalog")
    cat.Create(f"Provider={PROVEEDOR};Data Source={ruta};")
    # Cerrar la conexion que ADOX deja abierta para liberar el bloqueo del archivo
    try:
        cat.ActiveConnection.Close()
    except Exception:
        pass
    print(f"Archivo creado: {ruta}")


def ejecutar_ddl(ruta):
    """Abre la base con ADO y ejecuta tablas + indices."""
    conn = win32com.client.Dispatch("ADODB.Connection")
    conn.Open(f"Provider={PROVEEDOR};Data Source={ruta};")
    n_ok = 0
    for i, sql in enumerate(DDL_TABLAS, start=1):
        try:
            conn.Execute(sql)
            n_ok += 1
            print(f"Tabla {i}/{len(DDL_TABLAS)} creada.")
        except Exception as e:
            detalle = str(e).replace("\n", " ")[:200]
            print(f"ERROR creando tabla {i}: {detalle}")
            print(f"SQL: {sql.strip()}")
            conn.Close()
            raise
    for sql in DDL_INDICES:
        nombre_idx = sql.split()[2]
        try:
            conn.Execute(sql)
            print(f"Indice creado: {nombre_idx}")
        except Exception as e:
            detalle = str(e).replace("\n", " ")[:120]
            print(f"AVISO indice {nombre_idx}: {detalle}")
    conn.Close()
    return n_ok


def verificar_base(ruta):
    """Lee tablas y llaves desde el motor (ADOX) y compara con lo esperado.

    Notas ADOX:
      - Tipos de llave: 1 = Primary (PK), 2 = Foreign (FK), 3 = Unique.
      - El nombre de la PK es el del indice autogenerado (p.ej. Index_415C...),
        por eso la PK se valida por su COLUMNA, no por nombre.
      - Los UNIQUE (tipo 3) se ignoran: existen por los UNIQUE del DDL.
    """
    cat = win32com.client.Dispatch("ADOX.Catalog")
    cat.ActiveConnection = f"Provider={PROVEEDOR};Data Source={ruta};"

    tablas_encontradas = []
    errores = []

    for tabla in cat.Tables:
        if tabla.Type != "TABLE":
            continue
        nombre = tabla.Name
        tablas_encontradas.append(nombre)

        pk_col = None
        fks = []
        for k in tabla.Keys:
            if k.Type == 1:  # PK: tomar la primera columna de la llave
                pk_col = k.Columns(0).Name
            elif k.Type == 2:  # FK: nombre del constraint y tabla referenciada
                fks.append((k.Name, k.RelatedTable))

        esperado = LLAVES_ESPERADAS.get(nombre)
        if esperado is None:
            errores.append(f"{nombre}: tabla inesperada (no esta en LLAVES_ESPERADAS)")
            continue
        if pk_col != esperado["pk"]:
            errores.append(f"{nombre}: PK esperada en columna '{esperado['pk']}', obtenida '{pk_col}'")
        if sorted(fks) != sorted(esperado["fks"]):
            errores.append(f"{nombre}: FKs esperadas {sorted(esperado['fks'])} != obtenidas {sorted(fks)}")

        fk_txt = ", ".join(f"{n} -> {t}" for n, t in fks) or "(sin FK)"
        print(f"  {nombre}: PK({pk_col}) | FKs: {fk_txt}")

    try:
        cat.ActiveConnection.Close()
    except Exception:
        pass

    print(f"\nTablas encontradas: {len(tablas_encontradas)}/9 -> {sorted(tablas_encontradas)}")
    if errores:
        for e in errores:
            print(f"ERROR DE VERIFICACION: {e}")
        return False, len(tablas_encontradas)
    faltan = set(LLAVES_ESPERADAS) - set(tablas_encontradas)
    if faltan:
        print(f"ERROR: faltan tablas: {sorted(faltan)}")
        return False, len(tablas_encontradas)
    return True, len(tablas_encontradas)


def main():
    ruta = sys.argv[1] if len(sys.argv) > 1 else RUTA_BASE_DEFAULT
    print("=" * 60)
    print("FASE 1 - Creacion de base de datos Access (TiendaFutbol)")
    print("=" * 60)

    eliminar_base_si_existe(ruta)
    crear_base(ruta)
    n_tablas = ejecutar_ddl(ruta)

    print("\n--- VERIFICACION (leida desde el motor) ---")
    ok, total = verificar_base(ruta)

    if ok and total == 9 and n_tablas == 9:
        print("\nRESULTADO: EXITO. 9 tablas creadas con sus llaves foraneas verificadas.")
        print(f"Base lista en: {ruta}")
        return 0
    print(f"\nRESULTADO: FALLO (tablas ejecutadas={n_tablas}, verificadas={total}).")
    return 1


if __name__ == "__main__":
    sys.exit(main())
