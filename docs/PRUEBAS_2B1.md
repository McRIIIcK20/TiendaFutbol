# Pruebas FASE 2B-1: validaciones, clientes y usuarios

Fecha: 2026-10-13
Ejecutadas a mano en Excel (64 bits), desde la ventana Inmediato del editor de VBA.
Conexión: ADODB con `Microsoft.ACE.OLEDB.12.0` sobre `access\tienda_futbol.accdb`.

Las salidas están copiadas tal como aparecieron en Inmediato, sin resumir.
`ProbarValidaciones`, `ListarEsquema` y `ProbarRutsDemo` se ejecutaron en un libro donde
estaba importado `probarValidaciones.bas`; `ProbarClientes` se ejecutó después, en otro libro,
con `modValidaciones`, `modClientes`, `modUsuarios`, `modConexion` y `modSeguridad` importados.

| Macro | Módulo de prueba | Resultado |
|---|---|---|
| `ProbarValidaciones` | `probarValidaciones.bas` | OK=39, FALLAS=0 |
| `ListarEsquema` | `probarValidaciones.bas` | Esquema de USUARIOS y CLIENTES (informativo) |
| `ProbarRutsDemo` | `probarValidaciones.bas` | 150 válidos, 0 inválidos, 0 repetidos |
| `ProbarClientes` | `probarClientes.bas` | OK=41, FALLAS=0 (todo dentro de una transacción con rollback) |

## 1. ProbarValidaciones

```
--- RUT validos ---
OK    | RUT 12.345.678-5
OK    | RUT 12345678-5
OK    | RUT 123456785 (sin guion)
OK    | RUT con espacios
OK    | RUT con K 10.000.013-K
OK    | RUT con k minuscula
OK    | RUT con DV 0 10.000.004-0
OK    | RUT 11.111.111-1
--- RUT invalidos ---
OK    | DV incorrecto 12.345.678-9
OK    | vacio
OK    | texto abc
OK    | muy corto 1-9
OK    | sin DV 12.345.678
OK    | letra en cuerpo 12.345.67X-5
OK    | demasiado largo
OK    | K incorrecta 12.345.678-K
--- Digito verificador y formato ---
OK    | DV de 12345678 es 5
OK    | Formato 123456785
OK    | Formato 1234567-4
OK    | Formato de RUT invalido devuelve vacio
--- Correos ---
OK    | juan@example.com
OK    | a.b@sub.example.cl
OK    | vacio
OK    | sin arroba
OK    | sin parte local
OK    | dominio sin punto a@b
OK    | con espacio
OK    | doble arroba
OK    | puntos seguidos
OK    | dominio empieza con punto
OK    | termina en punto
OK    | parte local empieza con punto
--- Claves y campos obligatorios ---
OK    | clave Admin123!
OK    | clave 12345678
OK    | clave corta
OK    | clave vacia
OK    | campo con texto
OK    | campo con espacios
OK    | campo vacio
=== RESULTADO: OK=39 | FALLAS=0 ===
```

## 2. ListarEsquema

```
=== USUARIOS : columnas ===
  3. correo | tipo=Texto | largo=255 | acepta nulos=Falso
  4. hash_contrasena | tipo=Texto | largo=255 | acepta nulos=Falso
  1. id | tipo=Entero largo | largo=- | acepta nulos=Falso
  2. nombre | tipo=Texto | largo=100 | acepta nulos=Falso
  6. rol | tipo=Texto | largo=20 | acepta nulos=Falso
  5. sal | tipo=Texto | largo=64 | acepta nulos=Falso
=== USUARIOS : indices ===
  Index_16BB581D_98BC_462C | columna=correo | primaria=Falso | unico=Verdadero
  Index_CE3391C5_76AB_436C | columna=id | primaria=Verdadero | unico=Verdadero
=== CLIENTES : columnas ===
  3. apellido | tipo=Texto | largo=100 | acepta nulos=Falso
  4. correo | tipo=Texto | largo=255 | acepta nulos=Falso
  6. direccion | tipo=Texto | largo=255 | acepta nulos=Verdadero
  8. fecha_registro | tipo=Fecha | largo=- | acepta nulos=Falso
  1. id | tipo=Entero largo | largo=- | acepta nulos=Falso
  2. nombre | tipo=Texto | largo=100 | acepta nulos=Falso
  7. rut | tipo=Texto | largo=12 | acepta nulos=Falso
  5. telefono | tipo=Texto | largo=50 | acepta nulos=Verdadero
=== CLIENTES : indices ===
  Index_AA710E55_2116_4E1E | columna=correo | primaria=Falso | unico=Verdadero
  Index_B951EC73_7F6B_42CA | columna=id | primaria=Verdadero | unico=Verdadero
  Index_E1AA922C_0E91_49A1 | columna=rut | primaria=Falso | unico=Verdadero
  idx_clientes_rut | columna=rut | primaria=Falso | unico=Falso
=== CLIENTES : primeras 3 filas (datos demo) ===
  id=451 | nombre=María | apellido=Díaz | correo=joseneira0@example.com | telefono=+56 2 2451 8347 | direccion=Avda. Michelle Flores 7656 Piso 0, Andacollo, Región de Coquimbo | rut=10.356.355-0 | fecha_registro=13-10-2025 | 
  id=452 | nombre=Eugenia | apellido=Valenzuela | correo=zlobos1@example.com | telefono=+56 2 3887 2624 | direccion=Las Mosquetas 28, Vallenar, Región de Atacama, 8013260 | rut=13.737.768-3 | fecha_registro=17-12-2025 | 
  id=453 | nombre=Ítalo | apellido=Martínez | correo=zjimenez2@example.com | telefono=+56 2 2864 7468 | direccion=Psje. Blanca Sanhueza 10 Of. 978, La Serena, Región de Coquimbo, 8121910 | rut=17.237.966-4 | fecha_registro=23-02-2026 | 
```

## 3. ProbarRutsDemo

```
Columna de RUT detectada: rut
  guardado: 10.356.355-0  ->  10.356.355-0
  guardado: 13.737.768-3  ->  13.737.768-3
  guardado: 17.237.966-4  ->  17.237.966-4
  guardado: 14.418.802-0  ->  14.418.802-0
  guardado: 18.615.872-5  ->  18.615.872-5
RUT repetidos en CLIENTES: 0
=== RESULTADO: total=150 | validos=150 | invalidos=0 ===
```

## 4. ProbarClientes

```
Antes: CLIENTES=150 | USUARIOS=3
--- Precondiciones (los datos de prueba no deben existir) ---
OK    | RUT 10.000.013-K libre
OK    | RUT 10.000.004-0 libre
OK    | Correo prueba.temporal@example.com libre
OK    | Usuario usuario.prueba@example.com libre
--- Registro valido ---
OK    | Registro valido devuelve id > 0
OK    | Registro valido no deja mensaje
     id asignado: 601
OK    | Existe por RUT (10.000.013-K)
OK    | Existe por RUT en otro formato (10000013-k)
OK    | Existe por correo en mayusculas
OK    | ExisteCliente con solo el RUT repetido
OK    | ExisteCliente con solo el correo repetido
OK    | ExisteCliente con RUT y correo nuevos
OK    | RUT guardado como 10.000.013-K
OK    | Correo guardado en minusculas
OK    | Fecha de registro es hoy
--- Duplicados ---
OK    | RUT repetido se rechaza
     motivo: Ya existe un cliente con ese RUT.
OK    | RUT repetido explica el motivo
OK    | Correo repetido se rechaza
     motivo: Ya existe un cliente con ese correo.
OK    | Correo repetido explica el motivo
--- Datos invalidos ---
OK    | RUT con digito incorrecto se rechaza
     motivo: El RUT no es valido.
OK    | Correo invalido se rechaza
     motivo: El correo no es valido.
OK    | Nombre vacio se rechaza
     motivo: El nombre es obligatorio.
OK    | Apellido vacio se rechaza
     motivo: El apellido es obligatorio.
OK    | Telefono demasiado largo se rechaza
     motivo: El telefono admite hasta 50 caracteres.
OK    | Ningun dato invalido llego a la base
--- Campos opcionales vacios ---
OK    | Cliente sin telefono ni direccion se registra
OK    | Telefono vacio queda NULL
OK    | Direccion vacia queda NULL
--- Usuarios (modUsuarios) ---
OK    | Registro de usuario valido devuelve id > 0
     id asignado: 18
OK    | Login del usuario recien creado
OK    | El id del login coincide con el del registro
OK    | Rol guardado como Vendedor
OK    | Login con clave incorrecta falla
OK    | Correo de usuario repetido se rechaza
     motivo: Ya existe un usuario con ese correo.
OK    | Clave corta se rechaza
     motivo: La clave debe tener al menos 8 caracteres.
OK    | Rol invalido se rechaza
     motivo: El rol no es valido. Use Admin, Vendedor o Cliente.
OK    | Correo de usuario invalido se rechaza
     motivo: El correo no es valido.
--- Deshacer la transaccion ---
OK    | CLIENTES vuelve al conteo inicial
OK    | USUARIOS vuelve al conteo inicial
OK    | El cliente de prueba ya no existe
OK    | El usuario de prueba ya no existe
=== RESULTADO: OK=41 | FALLAS=0 ===
```
