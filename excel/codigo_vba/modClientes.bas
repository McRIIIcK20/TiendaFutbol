Attribute VB_Name = "modClientes"
Option Explicit

' ============================================================================
' modClientes - Registro de clientes con validacion y duplicados
' FASE 2B-1. Sin formularios: todo recibe parametros y devuelve un resultado.
'
' Esquema real (verificado con ListarEsquema):
'   CLIENTES: id (autonumerico), nombre(100), apellido(100), correo(255, unico),
'             telefono(50, acepta nulo), direccion(255, acepta nulo),
'             rut(12, unico), fecha_registro (fecha, obligatoria)
' El RUT se guarda con formato 12.345.678-5, igual que los datos demo.
' Los usuarios del sistema se registran en modUsuarios.
' Todas las consultas usan parametros: nunca se concatena texto del usuario.
' ============================================================================

' ---------------------------------------------------------------------------
' Consultas de existencia
' ---------------------------------------------------------------------------

' True si ya hay un cliente con ese RUT (acepta cualquier formato valido).
Public Function ExisteClientePorRut(conn As Object, ByVal rut As String) As Boolean
    On Error GoTo Manejador
    Dim rutFmt As String

    rutFmt = modValidaciones.FormatearRut(rut)
    If Len(rutFmt) = 0 Then Exit Function      ' RUT invalido: no puede existir
    ExisteClientePorRut = ExisteValor(conn, "CLIENTES", "rut", rutFmt, 12)
    Exit Function
Manejador:
    Err.Raise Err.Number, Err.Source, "Error al buscar cliente por RUT: " & Err.Description
End Function

' True si ya hay un cliente con ese correo (sin distinguir mayusculas).
Public Function ExisteClientePorCorreo(conn As Object, ByVal correo As String) As Boolean
    On Error GoTo Manejador
    Dim c As String

    c = LCase$(Trim$(correo))
    If Len(c) = 0 Then Exit Function
    ExisteClientePorCorreo = ExisteValor(conn, "CLIENTES", "correo", c, 255)
    Exit Function
Manejador:
    Err.Raise Err.Number, Err.Source, "Error al buscar cliente por correo: " & Err.Description
End Function

' True si el RUT o el correo ya estan registrados.
Public Function ExisteCliente(conn As Object, ByVal rut As String, ByVal correo As String) As Boolean
    ExisteCliente = ExisteClientePorRut(conn, rut) Or ExisteClientePorCorreo(conn, correo)
End Function

' ---------------------------------------------------------------------------
' Registro de clientes
' ---------------------------------------------------------------------------

' Valida y registra un cliente. Devuelve el id nuevo, o 0 si los datos no son
' validos o ya existe (el motivo queda en "mensaje"). Los errores de base de
' datos se propagan como error de VBA.
Public Function RegistrarCliente(conn As Object, ByVal nombre As String, _
        ByVal apellido As String, ByVal correo As String, ByVal rut As String, _
        ByVal telefono As String, ByVal direccion As String, _
        ByRef mensaje As String) As Long
    On Error GoTo Manejador
    Dim cmd As Object, rs As Object
    Dim rutFmt As String, sql As String
    Dim sqlTel As String, sqlDir As String

    RegistrarCliente = 0
    mensaje = ""

    nombre = Trim$(nombre)
    apellido = Trim$(apellido)
    correo = LCase$(Trim$(correo))
    telefono = Trim$(telefono)
    direccion = Trim$(direccion)

    If Not modValidaciones.ValidarCampoObligatorio(nombre) Then
        mensaje = "El nombre es obligatorio."
        Exit Function
    End If
    If Not modValidaciones.ValidarCampoObligatorio(apellido) Then
        mensaje = "El apellido es obligatorio."
        Exit Function
    End If
    If Not modValidaciones.ValidarCorreo(correo) Then
        mensaje = "El correo no es valido."
        Exit Function
    End If
    If Not modValidaciones.ValidarRut(rut) Then
        mensaje = "El RUT no es valido."
        Exit Function
    End If
    If Len(nombre) > 100 Or Len(apellido) > 100 Then
        mensaje = "El nombre y el apellido admiten hasta 100 caracteres."
        Exit Function
    End If
    If Len(telefono) > 50 Then
        mensaje = "El telefono admite hasta 50 caracteres."
        Exit Function
    End If
    If Len(direccion) > 255 Then
        mensaje = "La direccion admite hasta 255 caracteres."
        Exit Function
    End If

    rutFmt = modValidaciones.FormatearRut(rut)

    If ExisteClientePorRut(conn, rutFmt) Then
        mensaje = "Ya existe un cliente con ese RUT."
        Exit Function
    End If
    If ExisteClientePorCorreo(conn, correo) Then
        mensaje = "Ya existe un cliente con ese correo."
        Exit Function
    End If

    ' Telefono y direccion aceptan nulo: si vienen vacios se guarda NULL.
    ' Los fragmentos "NULL" y "?" son literales internos, no texto del usuario.
    If Len(telefono) = 0 Then sqlTel = "NULL" Else sqlTel = "?"
    If Len(direccion) = 0 Then sqlDir = "NULL" Else sqlDir = "?"

    sql = "INSERT INTO CLIENTES (nombre, apellido, correo, rut, telefono, direccion, fecha_registro) " & _
          "VALUES (?, ?, ?, ?, " & sqlTel & ", " & sqlDir & ", Date())"

    Set cmd = modConexion.NuevoComando(conn, sql)
    modConexion.ParamTexto cmd, "pNombre", nombre, 100
    modConexion.ParamTexto cmd, "pApellido", apellido, 100
    modConexion.ParamTexto cmd, "pCorreo", correo, 255
    modConexion.ParamTexto cmd, "pRut", rutFmt, 12
    If Len(telefono) > 0 Then modConexion.ParamTexto cmd, "pTelefono", telefono, 50
    If Len(direccion) > 0 Then modConexion.ParamTexto cmd, "pDireccion", direccion, 255
    cmd.Execute

    ' @@IDENTITY devuelve el ultimo autonumerico generado en ESTA conexion.
    Set rs = conn.Execute("SELECT @@IDENTITY")
    RegistrarCliente = CLng(rs.Fields(0).Value)
    rs.Close
    Exit Function
Manejador:
    Err.Raise Err.Number, Err.Source, "Error al registrar cliente: " & Err.Description
End Function

' ---------------------------------------------------------------------------
' Auxiliares privadas
' ---------------------------------------------------------------------------

' Cuenta filas con tabla.columna = valor. "tabla" y "columna" son siempre
' literales internos (nunca texto del usuario); el valor va como parametro.
Private Function ExisteValor(conn As Object, ByVal tabla As String, ByVal columna As String, _
        ByVal valor As String, ByVal largo As Long) As Boolean
    On Error GoTo Manejador
    Dim cmd As Object, rs As Object

    Set cmd = modConexion.NuevoComando(conn, _
        "SELECT COUNT(*) FROM " & tabla & " WHERE " & columna & " = ?")
    modConexion.ParamTexto cmd, "pValor", valor, largo

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open cmd, , 3, 1          ' adOpenStatic, adLockReadOnly
    ExisteValor = (CLng(rs.Fields(0).Value) > 0)
    rs.Close
    Exit Function
Manejador:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' Crea la ficha en CLIENTES de un usuario con rol Cliente (para que sus compras queden a su nombre).
' Genera un RUT valido de reemplazo (serie 90.000.000+) que el cliente puede actualizar despues.
Public Function CrearFichaClienteParaUsuario(conn As Object, ByVal idUsuario As Long, _
        ByVal nombreCompleto As String, ByVal correo As String, ByRef mensaje As String) As Long
    Dim nom As String, ape As String, p As Long, cuerpo As String, dv As String
    Dim intentos As Long, idNuevo As Long

    nombreCompleto = Trim$(nombreCompleto)
    p = InStr(nombreCompleto, " ")
    If p > 0 Then
        nom = Left$(nombreCompleto, p - 1)
        ape = Trim$(Mid$(nombreCompleto, p + 1))
    Else
        nom = nombreCompleto
    End If
    If Len(nom) = 0 Then nom = "Cliente"
    If Len(ape) = 0 Then ape = "Tienda"

    For intentos = 0 To 30
        cuerpo = CStr(90000000 + idUsuario + intentos)
        dv = modValidaciones.CalcularDvRut(cuerpo)
        idNuevo = RegistrarCliente(conn, nom, ape, correo, cuerpo & dv, "", "Compra en catalogo", mensaje)
        If idNuevo > 0 Then Exit For
        If InStr(1, mensaje, "RUT", vbTextCompare) = 0 Then Exit For
    Next intentos
    CrearFichaClienteParaUsuario = idNuevo
End Function

