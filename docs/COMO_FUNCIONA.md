# Cómo funciona TiendaFutbol

## 1. Visión general

```
 Excel (TiendaFutbol.xlsm)           Access (tienda_futbol.accdb)       Carpetas
 ┌────────────────────────┐   ADODB  ┌─────────────────────────┐   ┌────────────────────┐
 │ Hoja "Inicio" (portada) │ <------> │ 9 tablas + 5 vistas     │   │ imagenes_productos │
 │ Formularios (UserForms) │  ACE     │ stock, ventas, usuarios │   │ boletas (PDF)      │
 │ Módulos VBA (lógica)    │  OLEDB   └─────────────────────────┘   │ assets (logo)      │
 │ Hojas internas ocultas  │                                         └────────────────────┘
 └────────────────────────┘
```

* **Los datos viven en Access.** Excel solo es la interfaz y el motor de lógica (VBA).
* Las hojas del libro (`Catalogo`, `Ventas`, `Inventario`, `Configuracion`, `PlantillaBoleta`…)
  son internas: están *muy ocultas* y el usuario nunca las ve. Se refrescan solas tras cada login/venta.
* Si alguna vez necesita verlas: macro `MostrarHojasAdmin` (Alt+F8). Pide credenciales de un Admin.
  Para volver al modo tienda: macro `PrepararPortada`.

## 2. Flujo de uso

1. Se abre el libro (**Habilitar contenido**) → aparece la portada con el logo.
2. **ABRIR TIENDA** → `frmLogin` (correo + contraseña; SHA‑256 con sal por usuario).
3. `frmMenu` muestra solo los botones que el rol puede usar.
4. **SALIR** (portada) restaura Excel y cierra el libro. "Volver al inicio" (menú) regresa a la portada.

## 3. Roles

| Función | Cliente | Vendedor | Admin |
|---|:--:|:--:|:--:|
| Catálogo con fotos y compra | ✔ | ✔ (vende a un cliente elegido) | ✔ |
| Punto de venta (`frmVenta`) | – | ✔ | ✔ |
| Inventario y movimientos | – | ✔ (sin costos) | ✔ (con costos) |
| Registrar clientes | – | ✔ | ✔ |
| Gestión de usuarios | – | – | ✔ |
| Configuración de correo | – | – | ✔ |

Los permisos se definen en `modSesion` (`PuedeVender`, `PuedeVerInventario`, `PuedeAdministrarUsuarios`…)
y los formularios ocultan lo que no corresponde. Los costos nunca se cargan para Cliente/Vendedor.

## 4. Catálogo y compra (`frmCatalogo`)

* Cuadrícula de 6 productos por página con foto (`imagenes_productos/<SKU>.jpg`), precio y stock.
* Filtro por categoría y búsqueda por nombre, SKU o marca.
* Carrito con `+`, `-`, Quitar y Vaciar; el stock mostrado descuenta lo que ya está en el carrito.
* **Confirmar** → `modVentas.RegistrarVenta` (transacción: descuenta stock, registra movimiento,
  crea venta y boleta) → PDF (`modBoletas.GenerarBoletaPDF`) → correo si hay App Password.
* Cliente: la compra queda a su nombre (si no tiene ficha en CLIENTES se crea una automáticamente
  con un RUT de reemplazo de la serie 90.000.000). Vendedor/Admin: eligen el cliente en la lista.
* Si el correo falla, se muestra el motivo exacto del error (antes quedaba silencioso).

## 5. Administración de usuarios (`frmUsuarios`, solo Admin)

* **Nuevo usuario**: abre `frmRegistroUsuario` (valida correo y clave ≥ 8).
* **Cambiar rol**, **Restablecer clave** (sal nueva) y **Bloquear acceso** (invalida la clave; se
  revierte restableciéndola).
* **Eliminar**: solo si el usuario no tiene ventas/movimientos (para no perder historial); si los tiene,
  se sugiere bloquear. No se puede eliminar/degradar la propia cuenta ni al último Admin activo.

## 6. Configuración de correo (`frmConfiguracion`, solo Admin)

Guarda en `Configuracion` (A3 correo de prueba, A4 Gmail emisor, A5 contraseña de aplicación) y
guarda el libro. Cree la contraseña en https://myaccount.google.com/apppasswords (requiere
verificación en 2 pasos). **No suba el .xlsm a Git con la clave escrita.**
Los clientes `@example.com` reciben las boletas en el correo de prueba.

## 7. Imágenes de producto

`scripts/generar_imagenes.py` dibuja cada foto según la especificación del producto: colores
oficiales de cada selección (Chile local rojo, Argentina visitante azul marino, Brasil amarillo/verde,
Colombia amarillo/azul/rojo, España rojo/amarillo, Alemania blanco/negro/rojo/dorado), botines café con
tacos / sala con suela plana / césped sintético con tacos cortos, balones por tamaño (N5, N4, N3, mini)
y modelo, canilleras por talla y modelo, guantes por talla, y accesorios. Para usar fotos reales,
reemplace el archivo `<SKU>.jpg` (mismo nombre). Logo: `assets/logo_tienda.png` + `python scripts/generar_logo.py`.

## 8. Puesta en marcha

1. Windows + Excel de escritorio + Access Database Engine (ACE OLEDB 12.0) del mismo bitness que Excel.
2. `pip install pywin32 pillow`; en Excel activar *Confiar en el acceso al modelo de objetos de proyectos de VBA*.
3. `python scripts/generar_imagenes.py` y `python scripts/generar_logo.py` (si no existen las imágenes).
4. `python excel/construir_xlsm.py` → crea `excel/TiendaFutbol.xlsm`.
5. Abrir el .xlsm → Habilitar contenido → ABRIR TIENDA.
Si mueve el proyecto de carpeta, las rutas se resuelven solas (relativas al libro) aunque la hoja
Configuracion conserve la ruta antigua.

## 9. Archivos principales

| Archivo | Rol |
|---|---|
| `modInicio` | Portada, modo tienda (oculta la cinta), login → menú |
| `modEstilo` | Paleta, fuentes, estilo de botones/cajas de todos los formularios |
| `modSesion` / `modSeguridad` | Sesión y permisos / hash SHA‑256 con sal |
| `modVentas`, `modInventario`, `modBoletas` | Reglas de negocio, stock, PDF y correo |
| `modUsuarios`, `modClientes` | Altas, bajas, bloqueo y cambios de rol / clientes |
| `frmCatalogo`, `frmUsuarios`, `frmConfiguracion`, `frmMenu`, … | Interfaz |
| `ThisWorkbook.vba` | Eventos del libro (abrir, activar, cerrar) |
