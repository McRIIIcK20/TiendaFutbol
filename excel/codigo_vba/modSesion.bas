Attribute VB_Name = "modSesion"
Option Explicit

' ============================================================================
' modSesion - Quien esta conectado y que puede hacer (FASE 2B-2)
' Solo guarda el estado de la sesion; no mantiene conexiones abiertas
' (cada formulario abre y cierra su propia conexion con modConexion).
' ============================================================================

Public gSesionIniciada As Boolean
Public gUsuarioId As Long
Public gUsuarioNombre As String
Public gUsuarioCorreo As String
Public gUsuarioRol As String
Public gReiniciar As Boolean       ' True si el menu se cerro para volver al login

' Registra al usuario autenticado. Intenta leer su nombre; si falla, usa el correo.
Public Sub IniciarSesion(conn As Object, ByVal idUsuario As Long, _
                         ByVal correo As String, ByVal rol As String)
    Dim cmd As Object, rs As Object

    gUsuarioId = idUsuario
    gUsuarioCorreo = correo
    gUsuarioRol = rol
    gUsuarioNombre = correo

    On Error Resume Next
    Set cmd = modConexion.NuevoComando(conn, "SELECT nombre FROM USUARIOS WHERE id = ?")
    modConexion.ParamEntero cmd, "pId", idUsuario
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open cmd, , 3, 1
    If Not rs.EOF Then gUsuarioNombre = CStr(rs.Fields(0).Value)
    rs.Close
    On Error GoTo 0

    gSesionIniciada = True
End Sub

Public Sub CerrarSesion()
    gSesionIniciada = False
    gUsuarioId = 0
    gUsuarioNombre = ""
    gUsuarioCorreo = ""
    gUsuarioRol = ""
End Sub

' --- Permisos por rol -------------------------------------------------------

Public Function PuedeRegistrarClientes() As Boolean
    PuedeRegistrarClientes = gSesionIniciada And (gUsuarioRol = "Admin" Or gUsuarioRol = "Vendedor")
End Function

Public Function PuedeVender() As Boolean
    PuedeVender = gSesionIniciada And (gUsuarioRol = "Admin" Or gUsuarioRol = "Vendedor")
End Function

Public Function PuedeVerCostos() As Boolean
    PuedeVerCostos = gSesionIniciada And modSeguridad.PuedeVerCostos(gUsuarioRol)
End Function

Public Function PuedeAdministrarUsuarios() As Boolean
    PuedeAdministrarUsuarios = gSesionIniciada And modSeguridad.PuedeAdministrarUsuarios(gUsuarioRol)
End Function

' --- Permisos adicionales (menu por rol) -------------------------------------

' Cliente: solo catalogo y compra. Vendedor y Admin: ademas inventario, punto de venta y clientes.
Public Function PuedeVerInventario() As Boolean
    PuedeVerInventario = gSesionIniciada And (gUsuarioRol = "Admin" Or gUsuarioRol = "Vendedor")
End Function

Public Function PuedeComprarEnCatalogo() As Boolean
    PuedeComprarEnCatalogo = gSesionIniciada
End Function

Public Function EsAdmin() As Boolean
    EsAdmin = gSesionIniciada And (gUsuarioRol = "Admin")
End Function

Public Function PuedeConfigurar() As Boolean
    PuedeConfigurar = EsAdmin()
End Function

