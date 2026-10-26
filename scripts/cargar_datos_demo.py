# -*- coding: utf-8 -*-
"""
FASE 1 - Carga de datos demo para TiendaFutbol (datos_demo).

Condiciones del encargo (cumplidas):
1. Semilla fija (Faker.seed(42) + random.seed(42)) -> datos reproducibles.
   Idempotente: si hay datos, los borra (orden inverso de FKs) antes de insertar.
2. Faker locale es_CL; correos de clientes con dominio @example.com (nunca
   reales, para que nada se envie a direcciones verdaderas en la FASE 3).
3. Contrasenas: SHA-256 con sal individual por usuario (hash = sha256(sal+clave)).
   Las credenciales demo se anotan en docs/USUARIOS_DEMO.md.
4. Coherencia de inventario: stock de cada producto = entradas - salidas de
   MOVIMIENTOS_INVENTARIO. Entrada inicial por producto; ventas como salidas;
   nunca stock negativo.
5. Coherencia de ventas: subtotal = suma de detalles, IVA 19%, total = iva+neto;
   precio_venta > precio_costo (margen 20%..60%); fechas en los ultimos 12 meses
   con picos (diciembre, inicio de temporada, fechas de partidos).

Uso:
    python cargar_datos_demo.py
"""

import hashlib
import os
import random
import sys
from datetime import date, timedelta

from faker import Faker
import win32com.client

RUTA = r"D:\Proyectos\TiendaFutbol\access\tienda_futbol.accdb"
PROVEEDOR = "Microsoft.ACE.OLEDB.12.0"
IVA = 0.19

# --- Semillas fijas: datos reproducibles -----------------------------------
SEED = 42
Faker.seed(SEED)
random.seed(SEED)
fake = Faker("es_CL")

# --- Catalogo demo: 60 productos realistas ---------------------------------
# (categoria, nombre_base, marca, talla, precio_costo, precio_venta)
# margen 20%..60% sobre costo: precio_venta = costo * (1 + margen)
CATEGORIAS = [
    ("Camisetas",     "Camiseta"),
    ("Botines",       "Botin"),
    ("Balones",       "Balon"),
    ("Canilleras",    "Canillera"),
    ("Guantes",       "Guante de arquero"),
    ("Accesorios",    "Accesorio"),
]

MODELOS = {
    "Camisetas":     [("Chile Local", ("S","M","L","XL")), ("Argentina Visitante", ("S","M","L")),
                      ("Brasil Local", ("M","L","XL")), ("Colombia Local", ("S","M","L")),
                      ("Espana Local", ("M","L")), ("Alemania Local", ("S","M","L","XL"))],
    "Botines":       [("Futbol Cafe Tacos", ("38","39","40","41","42","43")),
                      ("Futbolito Sala", ("38","40","42")),
                      ("Cesped Sintetico TF", ("39","41","43"))],
    "Balones":       [("N5 Professional", (None,)), ("N5 Entrenamiento", (None,)),
                      ("N4 Juvenil", (None,)), ("Futbolito N4", (None,)),
                      ("N3 Infantil", (None,)), ("Mini Souvenir", (None,))],
    "Canilleras":    [("Pro Plastico", ("Chica","Mediana","Grande")),
                      ("Con Tobillera", ("Chica","Mediana")),
                      ("Escolar Reforzada", ("Chica","Mediana"))],
    "Guantes":       [("Pro Grip", ("8","9","10","11")), ("Entrenamiento", ("9","10")),
                      ("Juvenil Soft", ("6","7"))],
    "Accesorios":    [("Rodillera", (None,)), ("Cinta Capitana", (None,)),
                      ("Bomba de Aire", (None,)), ("Morral Deportivo 30L", (None,)),
                      ("Medias Futbol Par", ("25-27","28-30","31-33")), ("Pelota Pausa Mini", (None,))],
}

MARCAS = ["Adidas", "Nike", "Puma", "Umbro", "New Balance", "Kappa"]

PARTIDOS_IMPORTANTES = [
    date(2026, 10, 15),  # clasificatorias (ejemplo)
    date(2026, 11, 19),
    date(2026, 3, 31),
    date(2026, 6, 11),
    date(2026, 9, 9),
]

# --- Utilidades -------------------------------------------------------------


def a_sql(valor):
    """Escapa un valor para SQL literal (evita comillas rotas en nombres)."""
    if valor is None:
        return "NULL"
    if isinstance(valor, (int, float)):
        return repr(valor)
    return "'" + str(valor).replace("'", "''") + "'"


def rut_valido(seed_num):
    """Genera un RUT chileno con modulo 11 correcto (parte deterministica)."""
    cuerpo = 10_000_000 + (seed_num * 7919) % 19_000_000  # entre ~10M y ~29M
    s, m = 0, 2
    n = cuerpo
    while n > 0:
        s += (n % 10) * m
        m = 2 if m == 7 else m + 1
        n //= 10
    resto = 11 - (s % 11)
    dv = {11: "0", 10: "K"}.get(resto, str(resto))
    return f"{cuerpo:,}".replace(",", ".") + "-" + dv


def hash_contrasena(sal, clave):
    """SHA-256(sal + clave) en hexadecimal, como lo leera modSeguridad en VBA."""
    return hashlib.sha256((sal + clave).encode("utf-8")).hexdigest()


def peso_mes(d):
    """Peso estacional de una fecha para el volumen de ventas del dia."""
    p = 1.0
    if d.month == 12:
        p += 2.2                      # pico navideno fuerte
    if d.month in (1, 2) and d.day <= 15:
        p += 1.0                      # inicio de temporada (verano)
    if d.month == 3 and d.day <= 10:
        p += 1.2                      # arranque campeonato local
    if d in PARTIDOS_IMPORTANTES:
        p += 1.8                      # dias previos/de partidos
    if d.weekday() in (5, 6):
        p += 0.35                     # fin de semana vende un poco mas
    return p


# --- Limpieza idempotente ---------------------------------------------------


def limpiar_datos(conn):
    """Borra TODOS los datos en orden inverso de dependencias (idempotente)."""
    print("--- Limpieza de datos previos (idempotente) ---")
    for t in ["BOLETAS", "MOVIMIENTOS_INVENTARIO", "DETALLE_VENTAS",
              "VENTAS", "PRODUCTOS", "CATEGORIAS", "PROVEEDORES",
              "CLIENTES", "USUARIOS"]:
        conn.Execute(f"DELETE FROM {t}")
        print(f"  {t} limpio")


# --- Carga ------------------------------------------------------------------


def cargar(conn):
    cursor_ids = {}

    # --- 1. Usuarios demo (3, con sal individual y SHA-256) -------------
    print("--- USUARIOS ---")
    usuarios_demo = [
        ("Admin Demo",     "admin.demo@example.com",      "Admin123!",   "Admin"),
        ("Vendedor Demo",  "vendedor.demo@example.com",   "Vendedor123!","Vendedor"),
        ("Cliente Demo",   "cliente.demo@example.com",   "Cliente123!", "Cliente"),
    ]
    ids_usuarios = []
    for nombre, correo, clave, rol in usuarios_demo:
        sal = fake.uuid4().replace("-", "")[:32]
        h = hash_contrasena(sal, clave)
        conn.Execute(
            f"INSERT INTO USUARIOS (nombre, correo, hash_contrasena, sal, rol) "
            f"VALUES ({a_sql(nombre)}, {a_sql(correo)}, {a_sql(h)}, {a_sql(sal)}, {a_sql(rol)})"
        )
        rs, _ = conn.Execute("SELECT @@IDENTITY")
        ids_usuarios.append(int(rs.Fields(0).Value))
    print(f"  3 usuarios creados (ids {ids_usuarios})")

    # --- 2. Categorias ---------------------------------------------------
    print("--- CATEGORIAS ---")
    id_cat = {}
    for nombre, _ in CATEGORIAS:
        conn.Execute(
            f"INSERT INTO CATEGORIAS (nombre, descripcion) "
            f"VALUES ({a_sql(nombre)}, {a_sql('Productos de la linea ' + nombre)})"
        )
        rs, _ = conn.Execute("SELECT @@IDENTITY")
        id_cat[nombre] = int(rs.Fields(0).Value)
    print(f"  {len(CATEGORIAS)} categorias creadas")

    # --- 3. Proveedores ---------------------------------------------------
    print("--- PROVEEDORES ---")
    ids_proveedores = []
    for _ in range(5):
        conn.Execute(
            f"INSERT INTO PROVEEDORES (nombre, contacto, telefono, correo, direccion) "
            f"VALUES ({a_sql(fake.company())}, {a_sql(fake.name())}, "
            f"{a_sql(fake.phone_number()[:18])}, {a_sql(fake.email(domain='example.com'))}, "
            f"{a_sql(fake.address().replace(chr(10), ', ')[:200])})"
        )
        rs, _ = conn.Execute("SELECT @@IDENTITY")
        ids_proveedores.append(int(rs.Fields(0).Value))
    print(f"  {len(ids_proveedores)} proveedores creados")

    # --- 4. Productos (60, margen 20%..60%) ------------------------------
    print("--- PRODUCTOS ---")
    productos = []   # (sku, nombre, precio_costo, precio_venta, stock_objetivo)
    seq = 0
    for cat_nombre, prefijo in CATEGORIAS:
        for modelo, tallas in MODELOS[cat_nombre]:
            for talla in tallas:
                if len(productos) >= 60:
                    break
                seq += 1
                sku = f"{cat_nombre[:3].upper()}-{seq:03d}"
                nombre = f"{prefijo} {modelo}" + (f" Talla {talla}" if talla else "")
                costo = random.choice([3000, 5000, 8000, 12000, 18000, 25000, 35000, 45000])
                margen = random.uniform(0.20, 0.60)
                venta = round(costo * (1 + margen), -2)   # precio "comercial" (a 100)
                if venta <= costo:
                    venta = costo * 1.2                    # salvaguarda del margen
                stock_objetivo = random.randint(4, 40)
                productos.append((sku, nombre, id_cat[cat_nombre], talla,
                                  random.choice(MARCAS), float(costo), float(venta),
                                  stock_objetivo))
        if len(productos) >= 60:
            productos = productos[:60]
            break
    assert len(productos) == 60, f"productos generados: {len(productos)}"

    # Stock disponible por SKU (se consume con las ventas; al final se UPDATEa
    # PRODUCTOS.stock). Parte igual a la entrada inicial -> coherencia garantizada.
    disponible = {sku: stock_obj for sku, *_rest, stock_obj in productos}

    for sku, nombre, icat, talla, marca, costo, venta, _stock in productos:
        conn.Execute(
            f"INSERT INTO PRODUCTOS (sku, nombre, id_categoria, talla, marca, "
            f"precio_costo, precio_venta, stock, stock_minimo) "
            f"VALUES ({a_sql(sku)}, {a_sql(nombre)}, {icat}, {a_sql(talla)}, "
            f"{a_sql(marca)}, {costo}, {venta}, 0, {random.randint(2, 8)})"
        )
    print(f"  {len(productos)} productos creados")

    # --- 5. Movimientos: ENTRADA INICIAL por producto --------------------
    # La compra inicial a proveedor deja las unidades que el stock objetivo pide.
    # stock_final = entradas - salidas se garantiza por construccion.
    print("--- MOVIMIENTOS_INVENTARIO (entradas iniciales) ---")
    hoy = date.today()
    fecha_inicio = hoy - timedelta(days=365)
    for sku, _n, _ic, _t, _m, _c, _v, stock_obj in productos:
        id_prov = random.choice(ids_proveedores)
        conn.Execute(
            f"INSERT INTO MOVIMIENTOS_INVENTARIO (sku_producto, id_usuario, id_proveedor, "
            f"tipo_movimiento, cantidad, fecha, motivo) "
            f"VALUES ({a_sql(sku)}, {ids_usuarios[0]}, {id_prov}, 'E', {stock_obj}, "
            f"#{fecha_inicio.strftime('%Y-%m-%d')}#, {a_sql('Carga inicial de inventario')})"
        )
    print(f"  {len(productos)} entradas iniciales registradas")

    # --- 6. Clientes (150, correos @example.com) --------------------------
    print("--- CLIENTES ---")
    ids_clientes = []
    vistos = set()
    for i in range(150):
        # RUTs con dv valido y sin duplicados
        while True:
            rut = rut_valido(i * 131 + random.randint(0, 999))
            if rut not in vistos:
                vistos.add(rut)
                break
        # correos @example.com (unicos, nunca reales)
        correo = f"{fake.user_name().lower()}{i}@example.com"
        conn.Execute(
            f"INSERT INTO CLIENTES (nombre, apellido, correo, telefono, direccion, "
            f"rut, fecha_registro) VALUES ({a_sql(fake.first_name())}, {a_sql(fake.last_name())}, "
            f"{a_sql(correo)}, {a_sql(fake.phone_number()[:18])}, "
            f"{a_sql(fake.address().replace(chr(10), ', ')[:200])}, {a_sql(rut)}, "
            f"#{(fecha_inicio + timedelta(days=random.randint(0, 300))).strftime('%Y-%m-%d')}#)"
        )
        rs, _ = conn.Execute("SELECT @@IDENTITY")
        ids_clientes.append(int(rs.Fields(0).Value))
    print(f"  150 clientes creados")

    # --- 7. Ventas (800+, estacionales) -----------------------------------
    # Pesos por dia -> 850 ventas elegidas con random.choices (pesos, no reemplazo
    # logico; puede repetirse el dia, es correcto: mismo dia, varias boletas).
    dias = [fecha_inicio + timedelta(days=d) for d in range(365)]
    pesos = [peso_mes(d) for d in dias]
    fechas_ventas = random.choices(dias, weights=pesos, k=850)

    n_ventas = 0
    n_detalles = 0
    n_reposiciones = 0

    def reponer(sku, conn, ids_usuarios, ids_proveedores, fv):
        """Compra a proveedor: ENTRADA de reposicion cuando el stock se agota."""
        lote = random.randint(10, 30)
        conn.Execute(
            f"INSERT INTO MOVIMIENTOS_INVENTARIO (sku_producto, id_usuario, id_proveedor, "
            f"tipo_movimiento, cantidad, fecha, motivo) VALUES ({a_sql(sku)}, {ids_usuarios[0]}, "
            f"{random.choice(ids_proveedores)}, 'E', {lote}, #{fv.strftime('%Y-%m-%d')}#, "
            f"{a_sql('Reposicion por quiebre de stock')})"
        )
        return lote

    for fv in fechas_ventas:
        id_cli = random.choice(ids_clientes)
        id_usr = random.choice(ids_usuarios[:2])      # venta la hace Admin o Vendedor

        # 1..4 lineas por venta; cantidad limitada por stock disponible de la fecha
        n_lineas = random.randint(1, 4)
        elegidos = random.sample(productos, n_lineas)
        lineas = []
        for sku, _n, _ic, _t, _m, costo, venta, _s in elegidos:
            cant = random.randint(1, 3)
            # Nunca vender mas que el stock disponible (coherencia de inventario)
            cant = min(cant, disponible[sku])
            if cant <= 0:
                # Sin stock: reponer (compra a proveedor) y reintentar una vez
                lote = reponer(sku, conn, ids_usuarios, ids_proveedores, fv)
                disponible[sku] += lote
                n_reposiciones += 1
                cant = min(random.randint(1, 3), disponible[sku])
                if cant <= 0:
                    continue
            # Precio historico: venta con variacion +-5% (redondeo a 100)
            precio = float(round(venta * random.uniform(0.95, 1.05), -2))
            lineas.append((sku, cant, precio))
            disponible[sku] -= cant
        if not lineas:
            continue   # no fue posible vender (stock 0 incluso tras reposicion)
        subtotal = round(sum(c * p for _s, c, p in lineas), 2)
        iva = round(subtotal * IVA, 2)
        total = round(subtotal + iva, 2)

        conn.Execute(
            f"INSERT INTO VENTAS (id_usuario, id_cliente, fecha, subtotal, iva, total, estado) "
            f"VALUES ({id_usr}, {id_cli}, #{fv.strftime('%Y-%m-%d')}#, "
            f"{subtotal}, {iva}, {total}, 'Completada')"
        )
        rs, _ = conn.Execute("SELECT @@IDENTITY")
        id_venta = int(rs.Fields(0).Value)

        for sku, cant, precio in lineas:
            conn.Execute(
                f"INSERT INTO DETALLE_VENTAS (id_venta, sku_producto, cantidad, "
                f"precio_unitario, subtotal) VALUES ({id_venta}, {a_sql(sku)}, "
                f"{cant}, {precio}, {round(cant * precio, 2)})"
            )
            # SALIDA de inventario por la venta
            conn.Execute(
                f"INSERT INTO MOVIMIENTOS_INVENTARIO (sku_producto, id_usuario, id_proveedor, "
                f"tipo_movimiento, cantidad, fecha, motivo) VALUES ({a_sql(sku)}, {id_usr}, NULL, "
                f"'S', {cant}, #{fv.strftime('%Y-%m-%d')}#, {a_sql('Venta ' + str(id_venta))})"
            )
            n_detalles += 1
        n_ventas += 1
    print(f"  {n_ventas} ventas con {n_detalles} lineas de detalle")
    print(f"  {n_reposiciones} reposiciones de inventario (entradas 'E' por quiebre)")

    # --- 7b. Persistir el stock final en PRODUCTOS ------------------------
    # stock = entradas - salidas (las salidas son las ventas ya insertadas)
    print("--- PRODUCTOS: actualizando stock final ---")
    for sku, _n, _ic, _t, _m, _c, _v, _s in productos:
        conn.Execute(
            f"UPDATE PRODUCTOS SET stock = {disponible[sku]} WHERE sku = {a_sql(sku)}"
        )
    print(f"  {len(productos)} stocks actualizados (quedan sin stock los SKU agotados)")

    # --- 8. Boletas: correlativo por venta (enviada='N' -> pendientes FASE 3)
    print("--- BOLETAS ---")
    # Access no admite subconsulta dentro de un literal de fecha (#...#):
    # se lee la fecha de cada venta y se inserta como literal.
    rs, _ = conn.Execute("SELECT id, fecha FROM VENTAS ORDER BY id")
    fechas_ventas_bd = []
    while not rs.EOF:
        fechas_ventas_bd.append((int(rs.Fields("id").Value), rs.Fields("fecha").Value))
        rs.MoveNext()
    rs.Close()
    for id_venta, fecha_v in fechas_ventas_bd:
        correlativo = f"B-{id_venta:06d}"
        conn.Execute(
            f"INSERT INTO BOLETAS (id_venta, numero_correlativo, ruta_pdf, fecha_emision, enviada_correo) "
            f"VALUES ({id_venta}, {a_sql(correlativo)}, NULL, "
            f"#{fecha_v.strftime('%Y-%m-%d')}#, 'N')"
        )
    print(f"  {len(fechas_ventas_bd)} boletas en estado pendiente de envio ('N')")

    return n_ventas, n_detalles


def main():
    print("=" * 60)
    print("FASE 1 - Carga de datos demo (TiendaFutbol)")
    print(f"Semilla: {SEED} (Faker.seed + random.seed) -> reproducible")
    print("=" * 60)

    conn = win32com.client.Dispatch("ADODB.Connection")
    conn.Open(f"Provider={PROVEEDOR};Data Source={RUTA};")

    limpiar_datos(conn)
    n_ventas, n_detalles = cargar(conn)
    conn.Close()

    print("\nRESULTADO: carga completa (pendiente de validaciones del encargo).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
