Eliminado archivo previo: D:\Proyectos\TiendaFutbol\excel\TiendaFutbol.xlsm
Iniciando Excel...
Rango 'RutaBaseDatos' configurado -> D:\Proyectos\TiendaFutbol\access\tienda_futbol.accdb
Importando módulos .bas...
  Importado módulo: modConexion
  Importado módulo: modSeguridad
  Importado módulo: modValidaciones
  Importado módulo: modClientes
  Importado módulo: modUsuarios
  Importado módulo: modInventario
  Importado módulo: modVentas
  Importado módulo: modSesion
  Importado módulo: modInicio
Construyendo UserForms desde D:\Proyectos\TiendaFutbol\excel\codigo_vba\formularios
  UserForm construido en el libro: frmLogin (8 controles)
  UserForm construido en el libro: frmMenu (8 controles)
  UserForm construido en el libro: frmRegistroCliente (16 controles)
  UserForm construido en el libro: frmRegistroUsuario (14 controles)


Ejecutando suite VBA modPruebaVentas.ProbarVentasEInventario...
Módulo temporal modPruebaVentas retirado del libro de producción.

--- SALIDA DE PRUEBAS FASE 2C-1 ---
Antes: VENTAS=850 | DETALLE_VENTAS=2146 | MOVIMIENTOS=2367 | BOLETAS=850
--- 1. Catalogo y control de costos por rol (modInventario) ---
OK    | ExisteProducto con SKU real (ACC-053)
OK    | ExisteProducto en minusculas
OK    | ExisteProducto con SKU inexistente
OK    | ObtenerStock con SKU inexistente devuelve -1
OK    | ObtenerProducto como Admin devuelve True
OK    | Admin ve precio_costo > 0
OK    | precio_venta > precio_costo
OK    | ObtenerProducto como Vendedor devuelve True
OK    | Vendedor NO ve precio_costo (queda 0)
OK    | ObtenerProducto como Cliente devuelve True
OK    | Cliente NO ve precio_costo (queda 0)
OK    | Catalogo Admin tiene 60 productos
OK    | Catalogo Admin incluye columna precio_costo (9 cols)
OK    | Catalogo Vendedor omite columna precio_costo (8 cols)
--- 2. Movimientos de inventario (E, S, A y rechazo de stock negativo) ---
OK    | Entrada 'E' (+10) devuelve id > 0
OK    | Stock de ACC-053 sube en 10
OK    | Entrada 'E' (+5) en ACC-054 devuelve id > 0
OK    | Stock de ACC-054 sube en 5
OK    | Salida 'S' (-4) devuelve id > 0
OK    | Stock de ACC-053 queda en inicial + 6
OK    | Ajuste 'A' (-2) devuelve id > 0
OK    | Stock de ACC-053 queda en inicial + 4
OK    | Salida mayor al stock se rechaza (id=0)
     motivo: Stock insuficiente para SKU ACC-053 (disponible: 14).
OK    | Stock de ACC-053 no cambio tras rechazo
OK    | Movimiento con cantidad 0 se rechaza
OK    | Movimiento con tipo invalido 'X' se rechaza
--- 3. Venta transaccional multi-item (modVentas.RegistrarVenta) ---
OK    | RegistrarVenta valida devuelve idVenta > 0
OK    | RegistrarVenta no deja mensaje de error
OK    | Genera boleta 1:1 con idBoleta > 0
OK    | Correlativo tiene formato B-XXXXXX
     idVenta=2934 | correlativo=B-002934
OK    | VENTAS.subtotal cuadra con suma de lineas
OK    | VENTAS.iva cuadra con 19%
OK    | VENTAS.total cuadra con subtotal + iva
OK    | VENTAS.estado es 'Completada'
OK    | DETALLE_VENTAS tiene 2 lineas para la venta
OK    | BOLETAS registra enviada_correo = 'N'
OK    | Stock de ACC-053 descuento 2 unidades
OK    | Stock de ACC-054 descuento 3 unidades
--- 4. Atomicidad ante quiebre de stock en 2da linea del carrito ---
OK    | Venta con quiebre en 2do item se rechaza (id=0)
     motivo: Stock insuficiente para SKU ACC-054 (solicitado: 1001, disponible: 4).
OK    | Stock del 1er item (ACC-053) NO se desconto
OK    | No se agregaron ventas tras el rechazo
--- 5. Rechazo por rol Cliente y validaciones de entrada ---
OK    | Rol 'Cliente' no puede registrar ventas
     motivo: Su rol (Cliente) no tiene permiso para registrar ventas.
OK    | Cliente inexistente se rechaza
OK    | Cantidad 0 en venta se rechaza
--- 6. Anulacion de venta y restitucion de stock (modVentas.AnularVenta) ---
OK    | AnularVenta devuelve True
OK    | Estado de la venta pasa a 'Anulada'
OK    | Stock de ACC-053 restituido tras anulacion
OK    | Stock de ACC-054 restituido tras anulacion
OK    | Anular una venta ya anulada se rechaza
--- 7. Rollback de la transaccion de pruebas ---
OK    | VENTAS vuelve al conteo inicial (850)
OK    | DETALLE_VENTAS vuelve al conteo inicial (2146)
OK    | MOVIMIENTOS vuelve al conteo inicial (2367)
OK    | BOLETAS vuelve al conteo inicial (850)
OK    | Stock de ACC-053 vuelve a 10
OK    | Stock de ACC-054 vuelve a 2
=== RESULTADO: OK=55 | FALLAS=0 ===


ÉXITO: Libro guardado en D:\Proyectos\TiendaFutbol\excel\TiendaFutbol.xlsm
Componentes VBA en el libro:
  - ThisWorkbook (Tipo=100)
  - Hoja1 (Tipo=100)
  - Hoja2 (Tipo=100)
  - Hoja3 (Tipo=100)
  - Hoja4 (Tipo=100)
  - Hoja5 (Tipo=100)
  - Hoja6 (Tipo=100)
  - Hoja7 (Tipo=100)
  - Hoja8 (Tipo=100)
  - modConexion (Tipo=1)
  - modSeguridad (Tipo=1)
  - modValidaciones (Tipo=1)
  - modClientes (Tipo=1)
  - modUsuarios (Tipo=1)
  - modInventario (Tipo=1)
  - modVentas (Tipo=1)
  - modSesion (Tipo=1)
  - modInicio (Tipo=1)
  - frmLogin (Tipo=3)
  - frmMenu (Tipo=3)
  - frmRegistroCliente (Tipo=3)
  - frmRegistroUsuario (Tipo=3)
Excel cerrado correctamente.