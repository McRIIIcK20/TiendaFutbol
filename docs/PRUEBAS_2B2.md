# PRUEBAS 2B-2

## Formularios probados

- frmLogin: 8 controles. Valida campos vacíos, autentica con SHA-256(sal + clave), inicia sesión en modSesion y cierra al superar 3 intentos fallidos.
- frmMenu: 8 controles. Muestra nombre y rol conectado; para Admin habilita Registrar cliente y muestra Registrar usuario; para Vendedor habilita Registrar cliente y oculta Registrar usuario; para Cliente deshabilita Registrar cliente, oculta Registrar usuario y muestra aviso de acceso solo a catálogo; Cerrar sesion regresa a frmLogin.
- frmRegistroCliente: 16 controles. Formatea RUT automáticamente en txtRut_AfterUpdate, valida obligatorios/RUT/correo/duplicados vía modClientes.RegistrarCliente y muestra el ID creado.
- frmRegistroUsuario: 14 controles. Exclusivo para Admin, verifica coincidencia de claves (txtClave = txtClave2), valida longitud ≥ 8 y correo único vía modUsuarios.RegistrarUsuario.

## Permisos por rol verificados

- Admin: acceso a Registrar cliente y Registrar usuario.
- Vendedor: acceso solo a Registrar cliente.
- Cliente: solo aviso de catálogo, sin acceso a registros.

## Estructura de componentes VBA en TiendaFutbol.xlsm

- ThisWorkbook (Tipo=100)
- Hoja1..Hoja8 (Tipo=100)
- Módulos: modConexion, modSeguridad, modValidaciones, modClientes, modUsuarios, modSesion, modInicio (Tipo=1)
- UserForms: frmLogin, frmMenu, frmRegistroCliente, frmRegistroUsuario (Tipo=3)
