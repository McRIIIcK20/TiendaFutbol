# USUARIOS DEMO - SOLO PARA PRUEBAS

> **ADVERTENCIA**: Estas credenciales existen únicamente para probar el
> portafolio (login VBA, roles, flujos de venta). **NUNCA** usarlas en un
> entorno real o productivo. Las contraseñas están en claro aquí a propósito
> para poder probar el login; en la base solo se guarda `SHA-256(sal + clave)`
> con una sal individual de 32 caracteres hexadecimales por usuario
> (columna `sal`). El hash nunca se calcula desde la contraseña sola.

## Credenciales

| Usuario        | Correo                     | Contraseña     | Rol      |
|----------------|----------------------------|----------------|----------|
| Admin Demo     | admin.demo@example.com     | `Admin123!`    | Admin    |
| Vendedor Demo  | vendedor.demo@example.com  | `Vendedor123!` | Vendedor |
| Cliente Demo   | cliente.demo@example.com   | `Cliente123!`  | Cliente  |

## Notas de seguridad implementadas

1. **Hash con sal**: `hash_contrasena = SHA-256(sal + contrasena)`. La sal es
   única por usuario (generada con Faker UUID truncado a 32 hex). Dos usuarios
   con la misma contraseña tendrían hashes distintos.
2. **Verificación en VBA** (FASE 2, `modSeguridad`): el login calculará
   `SHA-256(sal_leida + clave_ingresada)` y comparará contra `hash_contrasena`.
   La sal se lee de la base, nunca se hardcodea.
3. **Roles**: Admin ve todo (costos, usuarios, inventario); Vendedor no ve
   costos ni administración de usuarios; Cliente solo consulta catálogo.
4. **Correos de clientes**: todos los clientes demo usan dominio
   `@example.com` (reservado por IANA para pruebas; no enruta correo real).
   El único correo externo de prueba es el que el usuario indique en la hoja
   `Configuración` (FASE 3).
5. **Cambio de contraseñas**: en producción, exigir longitud >= 8 y rotación;
   no reutilizar estas credenciales en ningún sistema real.

## Cómo se generaron

Script: `scripts/cargar_datos_demo.py` (semilla fija 42 → reproducible).
Volver a ejecutarlo regenera exactamente los mismos usuarios, sal y hashes.
