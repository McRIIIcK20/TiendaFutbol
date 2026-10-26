# TiendaFutbol

Sistema de gestión para una tienda de fútbol: ventas con control de inventario
transaccional, emisión de boletas PDF y su envío por correo, con análisis en
Power BI. Datos en Access, aplicación en Excel VBA.

## Arquitectura

- **Access** (`access/tienda_futbol.accdb`): base de datos con 9 tablas
  (USUARIOS, CATEGORIAS, PRODUCTOS, PROVEEDORES, CLIENTES, VENTAS,
  DETALLE_VENTAS, MOVIMIENTOS_INVENTARIO, BOLETAS), integridad con PK/FK/CHECK/
  UNIQUE, y 5 vistas analíticas para Power BI (FASE 4).
- **Excel VBA** (`excel/TiendaFutbol.xlsm`): aplicación completa — login
  SHA-256 con roles (Admin/Vendedor/Cliente), CRUD de clientes/usuarios,
  inventario, ventas transaccionales ACID, UserForms y boletas.
- **Boletas PDF** (`modBoletas`): llena la hoja `PlantillaBoleta` con diseño
  ejecutivo y la exporta con `ExportAsFixedFormat` a `boletas/<correlativo>.pdf`;
  guarda `ruta_pdf` en BOLETAS.
- **Gmail SMTP (CDO) / Outlook** (`modBoletas.EnviarBoletaPorCorreo`): envía la
  boleta al cliente; los correos demo `@example.com` se redirigen a
  `CorreoPrueba` y hay bloqueos anti-spam; marca `enviada_correo='S'`.
- **Power BI** (`powerbi/`): vistas SQL listas para importar
  (`vw_ventas_detalle`, `vw_resumen_mensual`, `vw_rentabilidad_productos`,
  `vw_movimientos_inventario`, `vw_ranking_clientes`) + guía de medidas DAX y
  diseño (`powerbi/medidas_dax_y_diseno.md`).

## Estructura de carpetas

```
access/          tienda_futbol.accdb (9 tablas + 5 vistas analíticas)
excel/           TiendaFutbol.xlsm y construir_xlsm.py (reconstruye el libro)
excel/codigo_vba/    módulos .bas de producción
excel/codigo_vba/pruebas_vba/   suites de pruebas VBA (temporales en el build)
excel/codigo_vba/formularios/   UserForms .frm
docs/            PROGRESO.md, PRUEBAS.md, PRUEBAS_*.md, USUARIOS_DEMO.md
scripts/         crear Access, carga demo, validaciones, vistas Power BI
powerbi/         medidas_dax_y_diseno.md y datos_csv/ (5 CSV exportados)
boletas/         PDFs de boletas generados (carpeta de salida RutaBoletas)
imagenes_productos/   60 fotos <SKU>.jpg acordes a la especificación de cada producto
assets/          logo_tienda.png, logo_form.jpg y portada.jpg (portada de la hoja Inicio)
```

## Credenciales demo

Ver **docs/USUARIOS_DEMO.md** (contraseñas, roles y sales). Resumen:

| Usuario | Rol | Correo |
|---------|-----|--------|
| Admin Demo | Admin | admin.demo@example.com |
| Vendedor Demo | Vendedor | vendedor.demo@example.com |
| Cliente Demo | Cliente | cliente.demo@example.com |

## Reconstrucción del libro

```
python excel/construir_xlsm.py
```

Reconstruye `excel/TiendaFutbol.xlsm` desde los fuentes (.bas/.frm). Requiere Windows,
Excel de escritorio, `pip install pywin32` y activar en Excel *Confiar en el acceso al modelo
de objetos de proyectos de VBA*. Antes, si falta algo gráfico:
`python scripts/generar_imagenes.py` y `python scripts/generar_logo.py` (necesitan Pillow).

## Configuración de la hoja Configuracion (nombres definidos)

| Nombre | Celda | Contenido |
|--------|-------|-----------|
| RutaBaseDatos | A1 | ruta de `access/tienda_futbol.accdb` |
| RutaBoletas | A2 | carpeta de salida de los PDF (`boletas/`) |
| CorreoPrueba | A3 | buzón real que recibe los correos de clientes `@example.com` |
| GmailUsuario | A4 | cuenta Gmail remitente |
| GmailClaveApp | A5 | **App Password de Google** (16 caracteres) para el envío real |

La App Password se crea en https://myaccount.google.com/apppasswords
(requiere verificación en 2 pasos). Sin ella, el envío queda en modo simulado.

## Conexión de Power BI

- **Directo a Access**: Obtener datos → Base de datos de Access →
  `access/tienda_futbol.accdb` → seleccionar las vistas `vw_*`.
- **Vía CSV**: Obtener datos → Texto/CSV → archivos de `powerbi/datos_csv/`
  (re-exportables con `python scripts/crear_vistas_powerbi.py`).

Medidas DAX sugeridas y diseño de la página en
`powerbi/medidas_dax_y_diseno.md`.


## Interfaz de tienda (portada + formularios por rol)

Al abrir el libro solo se ve la hoja **Inicio** (logo y botones **ABRIR TIENDA** / **SALIR**); las
demás hojas están *muy ocultas* y la cinta de Excel se contrae mientras la app está abierta.
Todo se hace en formularios con tema único (`modEstilo`: Georgia + Segoe UI, azul marino y esmeralda):

| Rol | Puede |
|-----|-------|
| Cliente | Ver catálogo con fotos, buscar/filtrar y comprar (boleta PDF + correo) |
| Vendedor | Lo anterior + punto de venta, inventario y registrar clientes |
| Admin | Todo + gestión de usuarios (registrar, cambiar rol, restablecer clave, bloquear, eliminar) y configuración de correo |

Explicación completa del funcionamiento: **docs/COMO_FUNCIONA.md**.
