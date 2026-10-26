Attribute VB_Name = "modUsuarios"
Option Explicit

' Marca de usuario bloqueado (reemplaza al hash de la clave)
Public Const CLAVE_BLOQUEADA As String = "BLOQUEADO"

' ============================================================================
' modUsuarios - Registro de usuarios del sistema (tabla USUARIOS)
' Columnas: id (autonumerico), nombre, correo (unico), hash_contrasena, sal, rol
' Reutiliza modSeguridad.CrearUsuario (sal aleatoria + SHA-256(sal + clave)).
' Importante: quien llame a RegistrarUsuario desde la interfaz debe comprobar
' antes que el usuario conectado sea Admin (modSeguridad.PuedeAdministrarUsuarios).
' ============================================================================

' True si ya existe un usuario con ese correo.
Public Function ExisteUsuarioPorCorreo(ByVal conn As Object, ByVal correo As String) As Boolean
    On Error GoTo Manejador
    Dim cmd As Object, rs As Object

    Set cmd = modConexion.NuevoComando(conn, "SELECT COUNT(*) FROM USUARIOS WHERE correo = ?")
    modConexion.ParamTexto cmd, "pCorreo", LCase$(Trim$(correo)), 255
    Set rs = cmd.Execute
    ExisteUsuarioPorCorreo = (CLng(rs.Fields(0).Value) > 0)
    rs.Close
    Exit Function
Manejador:
    Err.Raise Err.Number, Err.Source, "Error al buscar el usuario por correo: " & Err.Description
End Function

' Devuelve el rol con la mayuscula exacta que guarda la base (Admin, Vendedor, Cliente),
' o "" si no es un rol permitido. Acepta cualquier combinacion de mayusculas.
Public Function NormalizarRol(ByVal rol As String) As String
    Select Case LCase$(Trim$(rol))
        Case "admin"
            NormalizarRol = "Admin"
        Case "vendedor"
            NormalizarRol = "Vendedor"
        Case "cliente"
            NormalizarRol = "Cliente"
        Case Else
            NormalizarRol = ""
    End Select
End Function

' Registra un usuario. Devuelve el id nuevo (> 0), o 0 si los datos no son validos;
' en ese caso "mensaje" explica el motivo. Los errores de base de datos se propagan.
Public Function RegistrarUsuario(ByVal conn As Object, ByVal nombre As String, _
                                 ByVal correo As String, ByVal clave As String, _
                                 ByVal rol As String, ByRef mensaje As String) As Long
    On Error GoTo Manejador
    Dim rs As Object, rolN As String, correoN As String

    RegistrarUsuario = 0
    mensaje = ""

    If Not modValidaciones.ValidarCampoObligatorio(nombre) Then
        mensaje = "El nombre es obligatorio."
        Exit Function
    End If
    If Len(Trim$(nombre)) > 100 Then
        mensaje = "El nombre no puede superar los 100 caracteres."
        Exit Function
    End If
    If Not modValidaciones.ValidarCorreo(correo) Then
        mensaje = "El correo no es valido."
        Exit Function
    End If
    If Not modValidaciones.ValidarClave(clave) Then
        mensaje = "La clave debe tener al menos 8 caracteres."
        Exit Function
    End If

    rolN = NormalizarRol(rol)
    If Len(rolN) = 0 Then
        mensaje = "El rol no es valido. Use Admin, Vendedor o Cliente."
        Exit Function
    End If

    correoN = LCase$(Trim$(correo))
    If ExisteUsuarioPorCorreo(conn, correoN) Then
        mensaje = "Ya existe un usuario con ese correo."
        Exit Function
    End If

    modSeguridad.CrearUsuario conn, Trim$(nombre), correoN, clave, rolN

    ' @@IDENTITY devuelve el ultimo autonumerico generado en ESTA conexion.
    Set rs = conn.Execute("SELECT @@IDENTITY")
    RegistrarUsuario = CLng(rs.Fields(0).Value)
    rs.Close
    Exit Function
Manejador:
    Err.Raise Err.Number, Err.Source, "Error al registrar el usuario: " & Err.Description
End Function

' ============================================================================
' Administracion de usuarios (solo Admin; el formulario debe comprobarlo)
' ============================================================================


' Recordset con id, nombre, correo, rol y estado ("Activo"/"Bloqueado").
Public Function ListarUsuarios(ByVal conn As Object) As Object
    Set ListarUsuarios = modConexion.Consultar(conn, _
        "SELECT id, nombre, correo, rol, hash_contrasena FROM USUARIOS ORDER BY rol, nombre")
End Function

Private Function ContarAdmins(ByVal conn As Object) As Long
    Dim rs As Object
    Set rs = conn.Execute("SELECT COUNT(*) FROM USUARIOS WHERE rol = 'Admin' AND hash_contrasena <> '" & CLAVE_BLOQUEADA & "'")
    ContarAdmins = CLng(rs.Fields(0).Value)
    rs.Close
End Function

Private Function RolDeUsuario(ByVal conn As Object, ByVal idUsuario As Long) As String
    Dim cmd As Object, rs As Object
    Set cmd = modConexion.NuevoComando(conn, "SELECT rol FROM USUARIOS WHERE id = ?")
    modConexion.ParamEntero cmd, "pId", idUsuario
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open cmd, , 3, 1
    If Not rs.EOF Then RolDeUsuario = CStr(rs.Fields(0).Value)
    rs.Close
End Function

' True si el usuario tiene ventas o movimientos de inventario (no se puede borrar sin perder historial).
Public Function UsuarioTieneHistorial(ByVal conn As Object, ByVal idUsuario As Long) As Boolean
    Dim cmd As Object, rs As Object
    Set cmd = modConexion.NuevoComando(conn, "SELECT COUNT(*) FROM VENTAS WHERE id_usuario = ?")
    modConexion.ParamEntero cmd, "pId", idUsuario
    Set rs = cmd.Execute
    If CLng(rs.Fields(0).Value) > 0 Then UsuarioTieneHistorial = True
    rs.Close
    If UsuarioTieneHistorial Then Exit Function
    Set cmd = modConexion.NuevoComando(conn, "SELECT COUNT(*) FROM MOVIMIENTOS_INVENTARIO WHERE id_usuario = ?")
    modConexion.ParamEntero cmd, "pId", idUsuario
    Set rs = cmd.Execute
    If CLng(rs.Fields(0).Value) > 0 Then UsuarioTieneHistorial = True
    rs.Close
End Function

' Valida que el usuario objetivo pueda ser eliminado/bloqueado/degradado.
Private Function ValidarOperacionSobreUsuario(ByVal conn As Object, ByVal idObjetivo As Long, _
        ByVal idActual As Long, ByRef mensaje As String, ByVal quitaAdmin As Boolean) As Boolean
    If idObjetivo <= 0 Then
        mensaje = "Seleccione un usuario de la lista."
        Exit Function
    End If
    If idObjetivo = idActual Then
        mensaje = "No puede modificar ni eliminar su propia cuenta mientras la esta usando."
        Exit Function
    End If
    If quitaAdmin And StrComp(RolDeUsuario(conn, idObjetivo), "Admin", vbTextCompare) = 0 Then
        If ContarAdmins(conn) <= 1 Then
            mensaje = "Debe existir al menos un Admin activo."
            Exit Function
        End If
    End If
    ValidarOperacionSobreUsuario = True
End Function

' Elimina un usuario SIN historial. Devuelve False y deja el motivo en mensaje.
Public Function EliminarUsuario(ByVal conn As Object, ByVal idObjetivo As Long, _
        ByVal idActual As Long, ByRef mensaje As String) As Boolean
    Dim cmd As Object
    mensaje = ""
    If Not ValidarOperacionSobreUsuario(conn, idObjetivo, idActual, mensaje, True) Then Exit Function
    If UsuarioTieneHistorial(conn, idObjetivo) Then
        mensaje = "El usuario tiene ventas o movimientos registrados; no se puede eliminar sin perder el historial."
        Exit Function
    End If
    Set cmd = modConexion.NuevoComando(conn, "DELETE FROM USUARIOS WHERE id = ?")
    modConexion.ParamEntero cmd, "pId", idObjetivo
    cmd.Execute
    EliminarUsuario = True
End Function

' Bloquea el acceso (invalida la clave) conservando el historial. Se revierte con RestablecerClave.
Public Function BloquearUsuario(ByVal conn As Object, ByVal idObjetivo As Long, _
        ByVal idActual As Long, ByRef mensaje As String) As Boolean
    Dim cmd As Object
    mensaje = ""
    If Not ValidarOperacionSobreUsuario(conn, idObjetivo, idActual, mensaje, True) Then Exit Function
    Set cmd = modConexion.NuevoComando(conn, "UPDATE USUARIOS SET hash_contrasena = ? WHERE id = ?")
    modConexion.ParamTexto cmd, "pHash", CLAVE_BLOQUEADA, 255
    modConexion.ParamEntero cmd, "pId", idObjetivo
    cmd.Execute
    BloquearUsuario = True
End Function

Public Function CambiarRolUsuario(ByVal conn As Object, ByVal idObjetivo As Long, _
        ByVal idActual As Long, ByVal rolNuevo As String, ByRef mensaje As String) As Boolean
    Dim cmd As Object, rolN As String
    mensaje = ""
    rolN = NormalizarRol(rolNuevo)
    If Len(rolN) = 0 Then
        mensaje = "Seleccione un rol valido (Admin, Vendedor o Cliente)."
        Exit Function
    End If
    If Not ValidarOperacionSobreUsuario(conn, idObjetivo, idActual, mensaje, (rolN <> "Admin")) Then Exit Function
    Set cmd = modConexion.NuevoComando(conn, "UPDATE USUARIOS SET rol = ? WHERE id = ?")
    modConexion.ParamTexto cmd, "pRol", rolN, 20
    modConexion.ParamEntero cmd, "pId", idObjetivo
    cmd.Execute
    CambiarRolUsuario = True
End Function

' Asigna una clave nueva (sal nueva + SHA-256). Tambien reactiva a un usuario bloqueado.
Public Function RestablecerClave(ByVal conn As Object, ByVal idObjetivo As Long, _
        ByVal claveNueva As String, ByRef mensaje As String) As Boolean
    Dim cmd As Object, sal As String
    mensaje = ""
    If idObjetivo <= 0 Then
        mensaje = "Seleccione un usuario de la lista."
        Exit Function
    End If
    If Not modValidaciones.ValidarClave(claveNueva) Then
        mensaje = "La clave debe tener al menos 8 caracteres."
        Exit Function
    End If
    sal = modSeguridad.GenerarSal()
    Set cmd = modConexion.NuevoComando(conn, "UPDATE USUARIOS SET hash_contrasena = ?, sal = ? WHERE id = ?")
    modConexion.ParamTexto cmd, "pHash", modSeguridad.SHA256Hex(sal & claveNueva), 255
    modConexion.ParamTexto cmd, "pSal", sal, 64
    modConexion.ParamEntero cmd, "pId", idObjetivo
    cmd.Execute
    RestablecerClave = True
End Function

