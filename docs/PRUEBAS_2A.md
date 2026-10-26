# PRUEBAS FASE 2A — Login y verificación de credenciales

## Fuente (c) — Entorno

- Excel de 64 bits; conexión ADODB + `Microsoft.ACE.OLEDB.12.0` funcionando desde Excel (las pruebas VBA se ejecutaron correctamente con ese proveedor).
- Python de 64 bits (`C:\Users\mc18m\AppData\Local\hermes\tools\python-3.14.7+...-win32-x64\python.exe`) SIN el proveedor ACE registrado (error COM "Clase no registrada"): los scripts Python que accedan a la base deben usar DAO (`DAO.DBEngine.120`).
- Base: `access/tienda_futbol.accdb`, tabla `USUARIOS`. Columnas reales: `id`, `nombre`, `correo`, `hash_contrasena`, `sal`, `rol` (no existen `email` ni `clave`).

## Fuente (a) — Parte A en Python (SHA-256(sal + clave), UTF-8, hexdigest; solo se muestran los primeros 12 caracteres de sal/hash; no se imprimen contraseñas)

```
PYTHON: C:\Users\mc18m\AppData\Local\hermes\tools\python-3.14.7+202****0901-win32-x64\python.exe 64 bits
usuarios en tabla: 3
--- pruebas con clave CORRECTA ---
usuario=Admin Demo | rol=Admin | sal12=bdd640fb0667 | esperado12=b75174f04403 | calculado12=b75174f04403 | COINCIDE: True
usuario=Vendedor Demo | rol=Vendedor | sal12=23b8c1e93924 | esperado12=b83fca712905 | calculado12=b83fca712905 | COINCIDE: True
usuario=Cliente Demo | rol=Cliente | sal12=bd9c66b3ad3c | esperado12=4a75a236536d | calculado12=4a75a236536d | COINCIDE: True
--- prueba con clave INCORRECTA (correo real, clave equivocada) ---
usuario=Admin Demo | rol=Admin | sal12=bdd640fb0667 | esperado12=b75174f04403 | calculado12=e9ab760ad4e8 | COINCIDE: False
--- prueba con correo INEXISTENTE ---
correo=noexiste@example.c... | rol=N/A | sal=N/A | esperado=N/A | calculado=3973e022e932 | COINCIDE: NO (correo inexistente)
```

Resultado: 3/3 hashes coinciden con la base; 2 casos negativos rechazados como debe.

## Fuente (b) — Prueba VBA ejecutada en Excel (64 bits, ADODB + Microsoft.ACE.OLEDB.12.0)

```
Hash de 'test': 9f86d081884c...
Clave correcta | Admin | acceso=Verdadero | rol=Admin | id=15
Clave correcta | Vendedor | acceso=Verdadero | rol=Vendedor | id=16
Clave correcta | Cliente | acceso=Verdadero | rol=Cliente | id=17
Clave incorrecta | Admin | acceso=Falso
Correo inexistente | acceso=Falso
```

Resultado: 3/3 logins correctos con id y rol correctos (id 15/16/17); rechazo correcto de clave incorrecta y de correo inexistente.
