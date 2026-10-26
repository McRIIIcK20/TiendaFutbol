# PRUEBAS — Resumen maestro del proyecto

Todas las pruebas del proyecto TiendaFutbol superadas, por fase. Cada fase tiene
su documento con la salida literal completa en `docs/`.

| Fase | Ámbito | Resultado | Detalle literal |
|------|--------|-----------|-----------------|
| 1 | 9 tablas Access (PK, FK, CHECK, UNIQUE) + carga demo reproducible (60 productos, 150 clientes, 850+ ventas, 2 367 movimientos) | Verificación independiente OK (`scripts/verificar_bd.py` post-creación, `scripts/validar_datos_demo.py` re-ejecutado en Fase 5) | PROGRESO.md (commit 977017b, fa93e93) |
| 2A | Login con SHA-256 (Python + VBA idénticos), sesión y roles | 100% OK | docs/PRUEBAS_2A.md |
| 2B-1 | modValidaciones, modClientes, modUsuarios | 39/39 validaciones, 150/150 RUTs demo, 41/41 clientes/usuarios | docs/PRUEBAS_2B1.md |
| 2B-2 | 4 UserForms por rol (frmLogin, frmMenu, frmRegistroCliente, frmRegistroUsuario) | 100% OK | docs/PRUEBAS_2B2.md |
| 2C-1 | Motor de inventario y ventas transaccionales ACID | OK=55 \| FALLAS=0 | docs/PRUEBAS_2C1.md |
| 2C-2 | Formularios frmInventario y frmVenta integrados | 100% OK | docs/PRUEBAS_2C2.md |
| 3A | Generación de boletas PDF (diseño ejecutivo, carpeta `RutaBoletas`, `ruta_pdf` en BOLETAS) | OK=34 \| FALLAS=0 | docs/PRUEBAS_3A.md |
| 3B | Envío de boletas por correo (Gmail SMTP vía CDO / Outlook, redirección @example.com a `CorreoPrueba`, anti-spam, `enviada_correo='S'`) | OK=34 \| FALLAS=0 (misma suite 3A+3B) | docs/PRUEBAS_3B.md |
| 4 | 5 vistas SQL analíticas en Access + exportación CSV para Power BI | 5/5 vistas creadas y CSV exportados (re-verificado vía ADOX) | powerbi/medidas_dax_y_diseno.md |
| 5 | Cierre: datos validados tras crear las vistas + documentación y README | VALIDACION: TODAS LAS CONDICIONES CUMPLIDAS | docs/PROGRESO.md |

Notas:
- Las suites VBA (2C-1 y 3A/3B) se ejecutan importando el módulo temporal
  `modPruebaBoletas`/`modPruebaVentas` y retirándolo del libro final; usan
  transacción con `RollbackTrans` final: la BD queda sin filas de prueba y los
  PDFs temporales se eliminan del disco.
- El conteo real de 2C-1 es 55/55 (re-ejecutado en el build de FASE 3A); el
  mensaje del commit e22f2f2 citaba "47/47" por error.
- Verificación pendiente del usuario (documentada en PRUEBAS_3B.md): envío SMTP
  real con App Password y revisión visual del PDF.
- `scripts/verificar_bd.py` es solo para la BD recién creada (FASE 1, vacía);
  NO ejecutarlo contra la base con datos demo (su limpieza previa borra
  MOVIMIENTOS_INVENTARIO).
