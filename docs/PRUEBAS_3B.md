# PRUEBAS FASE 3B — Envío de boletas por correo (modBoletas.EnviarBoletaPorCorreo)

Ejecutado el 14-oct-2026 con `python excel/construir_xlsm.py --probar-3b` sobre
`D:\Proyectos\TiendaFutbol`. La suite `modPruebaBoletas.ProbarBoletasPDF` cubre
FASE 3A + FASE 3B (34 verificaciones) y se retira del libro de producción tras la
ejecución. Todo lo que modifica va dentro de una transacción con `RollbackTrans`
final; el PDF de prueba se elimina del disco y la BD queda sin restos.

Resultado: **OK=34 | FALLAS=0**.

## Salida literal

```
Antes: VENTAS=850 | BOLETAS=850 | con ruta_pdf=0 | enviadas='S'=0
Carpeta RutaBoletas: D:\Proyectos\TiendaFutbol\boletas
OK    | Precondicion: venta de prueba creada (id=2934, B-002934)
--- 1. FASE 3A - Venta valida: generacion de PDF y UPDATE en BOLETAS ---
OK    | GenerarBoletaPDF en venta valida devuelve True
OK    | GenerarBoletaPDF no deja mensaje de error
OK    | Ruta devuelta termina en B-002934.pdf
OK    | El archivo PDF existe en disco
OK    | El archivo PDF tiene tamano > 0 bytes
OK    | BOLETAS.ruta_pdf se actualizo en la BD
     PDF generado: D:\Proyectos\TiendaFutbol\boletas\B-002934.pdf
--- 2. FASE 3A - Regenerar una boleta ya generada ---
OK    | Regenerar boleta existente devuelve True
OK    | Regenerar mantiene la misma ruta B-002934.pdf
OK    | El PDF regenerado existe y pesa > 0 bytes
--- 3. FASE 3A - Carpeta de destino inexistente ---
OK    | Carpeta inexistente devuelve False
OK    | Carpeta inexistente deja rutaPdf vacia
OK    | Carpeta inexistente informa motivo claro
     motivo: La carpeta de destino para boletas no existe: D:\Proyectos\TiendaFutbol\carpeta_inexistente_3a_test
--- 4. FASE 3B - Envio por correo, redireccion @example.com y proteccion anti-spam ---
OK    | EnviarBoletaPorCorreo con CorreoPrueba valido devuelve True
OK    | Cliente @example.com se redirige a CorreoPrueba (m.coroseo.ramos@gmail.com)
     envio: Boleta B-002934 enviada a m.coroseo.ramos@gmail.com (Simulado).
OK    | BOLETAS.enviada_correo se actualizo a 'S'
OK    | Cliente @example.com sin CorreoPrueba se bloquea (False)
OK    | Bloqueo deja destinatario vacio
     motivo bloqueo vacio: El cliente usa dominio de prueba (@example.com) y la celda 'CorreoPrueba' en Configuracion esta vacia.
OK    | CorreoPrueba con dominio @example.com se bloquea (False)
     motivo bloqueo @example.com: La celda 'CorreoPrueba' no puede usar el dominio reservado @example.com.
OK    | EnviarBoletaPorCorreo regenera automaticamente el PDF si faltaba
--- 5. FASE 3A/3B - Venta anulada y venta inexistente ---
OK    | Anular la venta de prueba devuelve True
OK    | GenerarBoletaPDF en venta anulada devuelve False
OK    | Venta anulada deja rutaPdf vacia
OK    | Venta anulada informa que esta anulada
OK    | EnviarBoletaPorCorreo en venta anulada devuelve False
OK    | GenerarBoletaPDF en venta inexistente devuelve False
OK    | Venta inexistente deja rutaPdf vacia
OK    | Venta inexistente informa que no existe
OK    | EnviarBoletaPorCorreo en venta inexistente devuelve False
--- 6. Limpieza de PDF de prueba y RollbackTrans ---
OK    | PDF temporal de prueba eliminado del disco
OK    | VENTAS vuelve al conteo inicial (850)
OK    | BOLETAS vuelve al conteo inicial (850)
OK    | BOLETAS con ruta_pdf vuelve a 0
OK    | BOLETAS con enviada_correo='S' vuelve a 0
=== RESULTADO: OK=34 | FALLAS=0 ===
```

## Qué cubre la suite (resumen)

- **3A**: generación de PDF con diseño ejecutivo en la hoja PlantillaBoleta,
  archivo `<numero_correlativo>.pdf` en la carpeta del rango `RutaBoletas`,
  regeneración de boleta ya generada, carpeta de destino inexistente, venta
  anulada y venta inexistente; `UPDATE BOLETAS SET ruta_pdf` al éxito.
- **3B**: envío por correo con redirección de clientes `@example.com` hacia la
  celda `CorreoPrueba`, bloqueos anti-spam (CorreoPrueba vacía o con dominio
  `@example.com`), regeneración automática del PDF si faltaba y actualización de
  `BOLETAS.enviada_correo = 'S'` solo tras un envío exitoso.

## Envío real (pendiente de verificar por el usuario)

La suite ejecuta el envío en **modo simulado**: valida todo el flujo pero no
conecta a SMTP real. Para un envío real:

1. Activar la verificación en 2 pasos en la cuenta Google y crear una App
   Password en https://myaccount.google.com/apppasswords.
2. En la hoja **Configuracion**:
   - `GmailUsuario` (Configuracion!A4): cuenta Gmail remitente.
   - `GmailClaveApp` (Configuracion!A5): la App Password de Google de
     16 caracteres (sin espacios).
   - `CorreoPrueba` (Configuracion!A3): buzón real de prueba al que se
     redirigen los clientes `@example.com`.
3. Registrar una venta en la aplicación, enviar su boleta y confirmar la
   recepción del correo con el PDF adjunto en el buzón de destino.
