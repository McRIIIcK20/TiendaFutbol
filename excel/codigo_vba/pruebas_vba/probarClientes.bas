Attribute VB_Name = "modPruebaClientes"
Option Explicit

' ============================================================================
' modPruebaClientes - SOLO PRUEBAS. No va en el libro de produccion.
' Macro: ProbarClientes  (escribela en la ventana Inmediato, Ctrl+G, y Enter)
'
' Todo lo que escribe va dentro de una transaccion que se DESHACE al final,
' asi que no deja clientes ni usuarios de prueba en la base. Los autonumericos
' si avanzan (Access no los reutiliza), por eso los proximos ids tendran huecos.
' ============================================================================

Private Const CADENA As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Proyectos\TiendaFutbol\access\tienda_futbol.accdb;"

Private mOk As Long
Private mFalla As Long

Private Sub Chequear(ByVal etiqueta As String, ByVal obtenido As Boolean, ByVal esperado As Boolean)
    If obtenido = esperado Then
        mOk = mOk + 1
        Debug.Print "OK    | " & etiqueta
    Else
        mFalla = mFalla + 1
        Debug.Print "FALLA | " & etiqueta & " | esperado=" & esperado & " obtenido=" & obtenido
    End If
End Sub

Private Function Contar(ByVal conn As Object, ByVal tabla As String) As Long
    Dim rs As Object
    Set rs = conn.Execute("SELECT COUNT(*) FROM " & tabla)
    Contar = CLng(rs.Fields(0).Value)
    rs.Close
End Function

Public Sub ProbarClientes()
    On Error GoTo Fallo
    Dim conn As Object, rs As Object
    Dim msg As String, idC As Long, idC2 As Long
    Dim rolOut As String, idU As Long, idU2 As Long, ok As Boolean
    Dim abierta As Boolean
    Dim nClientesAntes As Long, nUsuariosAntes As Long

    mOk = 0
    mFalla = 0

    Set conn = CreateObject("ADODB.Connection")
    conn.Open CADENA

    nClientesAntes = Contar(conn, "CLIENTES")
    nUsuariosAntes = Contar(conn, "USUARIOS")
    Debug.Print "Antes: CLIENTES=" & nClientesAntes & " | USUARIOS=" & nUsuariosAntes

    Debug.Print "--- Precondiciones (los datos de prueba no deben existir) ---"
    Chequear "RUT 10.000.013-K libre", Not modClientes.ExisteClientePorRut(conn, "10.000.013-K"), True
    Chequear "RUT 10.000.004-0 libre", Not modClientes.ExisteClientePorRut(conn, "10.000.004-0"), True
    Chequear "Correo prueba.temporal@example.com libre", Not modClientes.ExisteClientePorCorreo(conn, "prueba.temporal@example.com"), True
    Chequear "Usuario usuario.prueba@example.com libre", Not modUsuarios.ExisteUsuarioPorCorreo(conn, "usuario.prueba@example.com"), True
    If mFalla > 0 Then
        Debug.Print "Hay datos de prueba previos en la base. Se detiene sin escribir nada."
        conn.Close
        Exit Sub
    End If

    conn.BeginTrans
    abierta = True

    Debug.Print "--- Registro valido ---"
    idC = modClientes.RegistrarCliente(conn, "Prueba", "Temporal", "Prueba.Temporal@Example.com", _
        "10000013-k", "+56 9 1234 5678", "Calle Falsa 123", msg)
    Chequear "Registro valido devuelve id > 0", idC > 0, True
    Chequear "Registro valido no deja mensaje", Len(msg) = 0, True
    Debug.Print "     id asignado: " & idC
    Chequear "Existe por RUT (10.000.013-K)", modClientes.ExisteClientePorRut(conn, "10.000.013-K"), True
    Chequear "Existe por RUT en otro formato (10000013-k)", modClientes.ExisteClientePorRut(conn, "10000013-k"), True
    Chequear "Existe por correo en mayusculas", modClientes.ExisteClientePorCorreo(conn, "PRUEBA.TEMPORAL@EXAMPLE.COM"), True
    Chequear "ExisteCliente con solo el RUT repetido", modClientes.ExisteCliente(conn, "10000013-K", "otro@example.com"), True
    Chequear "ExisteCliente con solo el correo repetido", modClientes.ExisteCliente(conn, "10.000.004-0", "prueba.temporal@example.com"), True
    Chequear "ExisteCliente con RUT y correo nuevos", Not modClientes.ExisteCliente(conn, "10.000.004-0", "nuevo@example.com"), True

    Set rs = conn.Execute("SELECT rut, correo, fecha_registro FROM CLIENTES WHERE id = " & idC)
    Chequear "RUT guardado como 10.000.013-K", CStr(rs.Fields(0).Value) = "10.000.013-K", True
    Chequear "Correo guardado en minusculas", CStr(rs.Fields(1).Value) = "prueba.temporal@example.com", True
    Chequear "Fecha de registro es hoy", CDate(rs.Fields(2).Value) = Date, True
    rs.Close

    Debug.Print "--- Duplicados ---"
    idC2 = modClientes.RegistrarCliente(conn, "Otro", "Cliente", "otro.cliente@example.com", "10.000.013-K", "", "", msg)
    Chequear "RUT repetido se rechaza", idC2 = 0, True
    Debug.Print "     motivo: " & msg
    Chequear "RUT repetido explica el motivo", Len(msg) > 0, True
    idC2 = modClientes.RegistrarCliente(conn, "Otro", "Cliente", "prueba.temporal@example.com", "10.000.004-0", "", "", msg)
    Chequear "Correo repetido se rechaza", idC2 = 0, True
    Debug.Print "     motivo: " & msg
    Chequear "Correo repetido explica el motivo", Len(msg) > 0, True

    Debug.Print "--- Datos invalidos ---"
    idC2 = modClientes.RegistrarCliente(conn, "Ana", "Perez", "ana@example.com", "12.345.678-9", "", "", msg)
    Chequear "RUT con digito incorrecto se rechaza", idC2 = 0, True
    Debug.Print "     motivo: " & msg
    idC2 = modClientes.RegistrarCliente(conn, "Ana", "Perez", "sin-arroba", "12.345.678-5", "", "", msg)
    Chequear "Correo invalido se rechaza", idC2 = 0, True
    Debug.Print "     motivo: " & msg
    idC2 = modClientes.RegistrarCliente(conn, "   ", "Perez", "ana@example.com", "12.345.678-5", "", "", msg)
    Chequear "Nombre vacio se rechaza", idC2 = 0, True
    Debug.Print "     motivo: " & msg
    idC2 = modClientes.RegistrarCliente(conn, "Ana", "", "ana@example.com", "12.345.678-5", "", "", msg)
    Chequear "Apellido vacio se rechaza", idC2 = 0, True
    Debug.Print "     motivo: " & msg
    idC2 = modClientes.RegistrarCliente(conn, "Ana", "Perez", "ana@example.com", "12.345.678-5", String(51, "1"), "", msg)
    Chequear "Telefono demasiado largo se rechaza", idC2 = 0, True
    Debug.Print "     motivo: " & msg
    Chequear "Ningun dato invalido llego a la base", Contar(conn, "CLIENTES") = nClientesAntes + 1, True

    Debug.Print "--- Campos opcionales vacios ---"
    idC2 = modClientes.RegistrarCliente(conn, "Sin", "Opcionales", "sin.opcionales@example.com", "10.000.004-0", "", "", msg)
    Chequear "Cliente sin telefono ni direccion se registra", idC2 > 0, True
    Set rs = conn.Execute("SELECT telefono, direccion FROM CLIENTES WHERE id = " & idC2)
    Chequear "Telefono vacio queda NULL", IsNull(rs.Fields(0).Value), True
    Chequear "Direccion vacia queda NULL", IsNull(rs.Fields(1).Value), True
    rs.Close

    Debug.Print "--- Usuarios (modUsuarios) ---"
    idU2 = modUsuarios.RegistrarUsuario(conn, "Usuario Prueba", "Usuario.Prueba@Example.com", "Clave12345", "vendedor", msg)
    Chequear "Registro de usuario valido devuelve id > 0", idU2 > 0, True
    Debug.Print "     id asignado: " & idU2
    ok = modSeguridad.Autenticar(conn, "usuario.prueba@example.com", "Clave12345", rolOut, idU)
    Chequear "Login del usuario recien creado", ok, True
    Chequear "El id del login coincide con el del registro", idU = idU2, True
    Chequear "Rol guardado como Vendedor", rolOut = "Vendedor", True
    ok = modSeguridad.Autenticar(conn, "usuario.prueba@example.com", "ClaveMala99", rolOut, idU)
    Chequear "Login con clave incorrecta falla", ok, False
    idU2 = modUsuarios.RegistrarUsuario(conn, "Otro", "usuario.prueba@example.com", "Clave12345", "Cliente", msg)
    Chequear "Correo de usuario repetido se rechaza", idU2 = 0, True
    Debug.Print "     motivo: " & msg
    idU2 = modUsuarios.RegistrarUsuario(conn, "Otro", "otro.usuario@example.com", "corta", "Cliente", msg)
    Chequear "Clave corta se rechaza", idU2 = 0, True
    Debug.Print "     motivo: " & msg
    idU2 = modUsuarios.RegistrarUsuario(conn, "Otro", "otro.usuario@example.com", "Clave12345", "Superman", msg)
    Chequear "Rol invalido se rechaza", idU2 = 0, True
    Debug.Print "     motivo: " & msg
    idU2 = modUsuarios.RegistrarUsuario(conn, "Otro", "correo-malo", "Clave12345", "Cliente", msg)
    Chequear "Correo de usuario invalido se rechaza", idU2 = 0, True
    Debug.Print "     motivo: " & msg

    Debug.Print "--- Deshacer la transaccion ---"
    conn.RollbackTrans
    abierta = False
    Chequear "CLIENTES vuelve al conteo inicial", Contar(conn, "CLIENTES") = nClientesAntes, True
    Chequear "USUARIOS vuelve al conteo inicial", Contar(conn, "USUARIOS") = nUsuariosAntes, True
    Chequear "El cliente de prueba ya no existe", Not modClientes.ExisteClientePorRut(conn, "10.000.013-K"), True
    Chequear "El usuario de prueba ya no existe", Not modUsuarios.ExisteUsuarioPorCorreo(conn, "usuario.prueba@example.com"), True

    Debug.Print "=== RESULTADO: OK=" & mOk & " | FALLAS=" & mFalla & " ==="
    conn.Close
    Exit Sub
Fallo:
    Debug.Print "ERROR " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If abierta Then conn.RollbackTrans
    conn.Close
End Sub
