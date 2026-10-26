# PROGRESO

**Estado general:** PROYECTO COMPLETADO (FASES 1 A 5)

## Resumen de logros

### FASE 2A
- Login SHA-256 implementado y verificado.

### FASE 2B-1
- modValidaciones, modClientes, modUsuarios desarrollados; 41/41 tests OK.

### FASE 2B-2
- modSesion, modInicio desarrollados; 4 UserForms construidos con construir_xlsm.py (frmLogin, frmMenu, frmRegistroCliente, frmRegistroUsuario).

**Pendiente:** FASE 2C (Ventas con transacción e Inventario).

## FASE 2C-1
- Motor de Inventario y Ventas con transacción ACID implementado y verificado.
- Pruebas VBA: OK=55 | FALLAS=0 (ver docs/PRUEBAS_2C1.md). El conteo real es 55
  (re-verificado re-ejecutando la suite en el build de FASE 3A); el mensaje del
  commit e22f2f2 citaba "47/47" por error.

## FASE 2C-2
- Formularios frmInventario y frmVenta integrados y verificados.
- Cierre completo de FASE 2.

## FASE 3A
- modBoletas.GenerarBoletaPDF(conn, idVenta, rutaPdf, mensaje): genera el PDF de
  la boleta llenando la hoja PlantillaBoleta y exportándola con
  ExportAsFixedFormat (xlTypePDF) a la carpeta del rango RutaBoletas de la hoja
  Configuracion, con nombre <numero_correlativo>.pdf (ej. B-000851.pdf).
- Lee VENTAS, DETALLE_VENTAS, PRODUCTOS, CLIENTES, USUARIOS y BOLETAS; rechaza
  ventas anuladas o inexistentes y carpeta inexistente con mensaje claro; al
  éxito hace UPDATE BOLETAS SET ruta_pdf.
- Pruebas VBA: OK=24 | FALLAS=0 (docs/PRUEBAS_3A.md), con transacción y
  RollbackTrans: sin filas de prueba en la BD y PDF temporal eliminado.
- construir_xlsm.py: hoja PlantillaBoleta, nombre definido RutaBoletas y modo
  --probar-3a (suite temporal retirada del libro final).

## FASE 3B
- modBoletas.EnviarBoletaPorCorreo: envío de la boleta PDF por correo (Gmail
  SMTP vía CDO / Outlook), con redirección de clientes @example.com a la celda
  'CorreoPrueba', bloqueos anti-spam, regeneración automática del PDF si faltaba
  y actualización de BOLETAS.enviada_correo = 'S' solo tras envío exitoso.
- La suite modPruebaBoletas.ProbarBoletasPDF cubre 3A+3B: OK=34 | FALLAS=0
  (docs/PRUEBAS_3B.md), con transacción y RollbackTrans final: sin restos en la
  BD ni PDFs de prueba en disco.
- construir_xlsm.py y frmVenta actualizados: rangos 'CorreoPrueba',
  'GmailUsuario' y 'GmailClaveApp' en la hoja Configuracion y modo --probar-3b.
- Envío real por SMTP: colocar la App Password de Google en Configuracion!
  (GmailClaveApp) y verificar la recepción (pendiente del usuario).
- Cierre de FASE 3 (3A + 3B): generación y envío de boletas completados.

## FASE 4
- 5 vistas SQL analíticas creadas en access/tienda_futbol.accdb y exportadas
  a CSV (powerbi/datos_csv/): vw_ventas_detalle (2 146 filas),
  vw_resumen_mensual (13), vw_rentabilidad_productos (60),
  vw_movimientos_inventario (2 367), vw_ranking_clientes (150).
- Guía de medidas DAX y diseño del informe en powerbi/medidas_dax_y_diseno.md.

## FASE 5
- Re-validación de la carga demo tras crear las vistas:
  VALIDACION: TODAS LAS CONDICIONES CUMPLIDAS (scripts/validar_datos_demo.py).
- docs/PRUEBAS.md: resumen maestro de todas las pruebas del proyecto.
- README.md en la raíz: arquitectura, estructura, credenciales demo,
  reconstrucción y conexión Power BI.

### Siguiente etapa
- (Ninguna: proyecto completado. Pendiente del usuario: envío SMTP real con
  App Password y revisión visual del PDF; ver docs/PRUEBAS_3B.md.)
## Actualización: interfaz de tienda por roles
- Portada (hoja Inicio) con logo y botones; hojas internas muy ocultas; modo tienda.
- Nuevos: `modEstilo`, `frmCatalogo`, `frmUsuarios`, `frmConfiguracion`, 60 imágenes por especificación, logo/portada.
- `frmMenu` por rol; `modUsuarios` con cambio de rol, bloqueo, eliminación segura y reset de clave.
- Correo: ahora se informa el motivo del fallo. Rutas de boletas con respaldo relativo.
- Pendiente de validar en Excel real: ejecutar `python excel/construir_xlsm.py` y recorrer los 3 roles (ver docs/COMO_FUNCIONA.md).
