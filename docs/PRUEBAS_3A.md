# PRUEBAS FASE 3A — Generación de boletas PDF (modBoletas.GenerarBoletaPDF)

Ejecutado el 14-oct-2026 con `python excel/construir_xlsm.py --probar-2c --probar-3a`
sobre `D:\Proyectos\TiendaFutbol`. La suite `modPruebaBoletas.ProbarBoletasPDF`
se importa como módulo temporal y se retira del libro de producción tras la ejecución.

La suite crea su venta de prueba dentro de una transacción con `RollbackTrans`
final y elimina el PDF temporal del disco: la base y la carpeta `boletas/` quedan
sin restos (verificado: VENTAS=850, BOLETAS=850, ruta_pdf=0, carpeta vacía).

Resultado: **OK=24 | FALLAS=0**.

## Salida literal

```
Antes: VENTAS=850 | BOLETAS=850 | BOLETAS con ruta_pdf=0
Carpeta RutaBoletas: D:\Proyectos\TiendaFutbol\boletas
OK    | Precondicion: venta de prueba creada (id=2935, B-002935)
--- 1. Venta valida: generacion de PDF y UPDATE en BOLETAS ---
OK    | GenerarBoletaPDF en venta valida devuelve True
OK    | GenerarBoletaPDF no deja mensaje de error
OK    | Ruta devuelta termina en B-002935.pdf
OK    | El archivo PDF existe en disco
OK    | El archivo PDF tiene tamano > 0 bytes
OK    | BOLETAS.ruta_pdf se actualizo en la BD
     PDF generado: D:\Proyectos\TiendaFutbol\boletas\B-002935.pdf
--- 2. Regenerar una boleta ya generada ---
OK    | Regenerar boleta existente devuelve True
OK    | Regenerar mantiene la misma ruta B-002935.pdf
OK    | El PDF regenerado existe y pesa > 0 bytes
--- 3. Carpeta de destino inexistente ---
OK    | Carpeta inexistente devuelve False
OK    | Carpeta inexistente deja rutaPdf vacia
OK    | Carpeta inexistente informa motivo claro
     motivo: La carpeta de destino para boletas no existe: D:\Proyectos\TiendaFutbol\carpeta_inexistente_3a_test
--- 4. Venta anulada ---
OK    | Anular la venta de prueba devuelve True
OK    | Venta anulada devuelve False
OK    | Venta anulada deja rutaPdf vacia
OK    | Venta anulada informa que esta anulada
     motivo: No se puede generar boleta de una venta anulada.
--- 5. Venta inexistente ---
OK    | Venta inexistente devuelve False
OK    | Venta inexistente deja rutaPdf vacia
OK    | Venta inexistente informa que no existe
     motivo: La venta indicada no existe.
--- 6. Limpieza de PDF de prueba y RollbackTrans ---
OK    | PDF temporal de prueba eliminado del disco
OK    | VENTAS vuelve al conteo inicial (850)
OK    | BOLETAS vuelve al conteo inicial (850)
OK    | BOLETAS con ruta_pdf vuelve a 0
=== RESULTADO: OK=24 | FALLAS=0 ===
```

## Contexto del build (misma ejecución)

- Hojas del libro: Configuracion, Ventas, Inventario, Productos, Clientes,
  Proveedores, Usuarios, Boletas, PlantillaBoleta.
- Nombres definidos: `RutaBaseDatos` (Configuracion!$A$1) y `RutaBoletas`
  (Configuracion!$A$2 → `D:\Proyectos\TiendaFutbol\boletas`).
- En la misma ejecución se re-corrió la suite 2C-1 (`ProbarVentasEInventario`):
  RESULTADO: OK=55 | FALLAS=0 (coincide con docs/PRUEBAS_2C1.md).
- El libro quedó guardado sin módulos de prueba (listado final de componentes:
  modConexion, modSeguridad, modValidaciones, modClientes, modUsuarios,
  modInventario, modVentas, modBoletas, modSesion, modInicio + 6 UserForms).

## Pendiente de verificar por el usuario

- El contenido visual del PDF (la suite verifica existencia y tamaño > 0, no el
  renderizado): abrir `boletas/B-XXXXXX.pdf` generado por una venta real y
  revisar encabezado, datos del cliente, tabla de ítems, subtotal/IVA/total.
