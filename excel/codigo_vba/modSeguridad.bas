Attribute VB_Name = "modSeguridad"
' modSeguridad
' Autenticacion y roles.
' Reglas del encargo:
'  - El hash guardado es SHA-256(sal + clave), UTF-8, hexadecimal en minusculas,
'    identico al generado en Python (hashlib).
'  - TODAS las consultas con ADODB.Command y parametros (nada de concatenar).
'  - Roles: Admin ve todo; Vendedor no ve costos ni usuarios.
Option Explicit

' Devuelve True si la combinacion correo+clave es valida.
' Sale por referencia el rol y el id del usuario.
Public Function Autenticar(conn As Object, correo As String, clave As String, _
                           ByRef rol As String, ByRef idUsuario As Long) As Boolean
    On Error GoTo Manejador
    Dim cmd As Object
    Dim rs As Object
    Dim sal As String, hashEsperado As String, hashCalculado As String

    Autenticar = False
    rol = ""
    idUsuario = 0

    ' Consulta parametrizada: nunca concatenar el correo ingresado
    Set cmd = modConexion.NuevoComando(conn, _
        "SELECT id, hash_contrasena, sal, rol FROM USUARIOS WHERE correo = ?")
    modConexion.ParamTexto cmd, "pCorreo", correo, 255

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open cmd, , 3, 1          ' adOpenStatic, adLockReadOnly

    If rs.EOF Then
        ' Correo inexistente: mismo mensaje que clave incorrecta (no filtrar usuarios)
        rs.Close
        Exit Function
    End If

    sal = CStr(rs.Fields("sal").Value)
    hashEsperado = LCase$(CStr(rs.Fields("hash_contrasena").Value))
    rol = CStr(rs.Fields("rol").Value)
    idUsuario = CLng(rs.Fields("id").Value)
    rs.Close

    ' Mismo calculo que Python: sha256((sal + clave).encode("utf-8")).hexdigest()
    hashCalculado = LCase$(SHA256Hex(sal & clave))
    Autenticar = (hashCalculado = hashEsperado)
    Exit Function
Manejador:
    modConexion.CerrarConexion Nothing
    Err.Raise Err.Number, Err.Source, "Error al autenticar: " & Err.Description
End Function

' SHA-256 de una cadena -> hexadecimal en minusculas (64 chars).
' Implementacion con .NET System.Security.Cryptography (disponible en Windows).
Public Function SHA256Hex(texto As String) As String
    On Error GoTo Manejador
    Dim enc As Object, sha As Object
    Dim bytes() As Byte, digesto() As Byte
    Dim i As Long, salida As String

    Set enc = CreateObject("System.Text.UTF8Encoding")
    Set sha = CreateObject("System.Security.Cryptography.SHA256Managed")

    bytes = enc.GetBytes_4(texto)
    digesto = sha.ComputeHash_2(bytes)
    sha.Clear

    salida = ""
    For i = LBound(digesto) To UBound(digesto)
        salida = salida & Right$("0" & Hex$(digesto(i)), 2)
    Next i
    SHA256Hex = LCase$(salida)    ' minusculas, igual que hexdigest() de Python
    Exit Function
Manejador:
    Err.Raise Err.Number, Err.Source, "Error calculando SHA-256: " & Err.Description
End Function

' Control de roles: puede ver costos?
Public Function PuedeVerCostos(rol As String) As Boolean
    PuedeVerCostos = (StrComp(rol, "Admin", vbTextCompare) = 0)
End Function

' Control de roles: puede administrar usuarios?
Public Function PuedeAdministrarUsuarios(rol As String) As Boolean
    PuedeAdministrarUsuarios = (StrComp(rol, "Admin", vbTextCompare) = 0)
End Function

' Genera una sal aleatoria (32 hex) para nuevos usuarios, igual formato que Python.
' Usa RND con Randomize; no criptografica pero suficiente para datos demo.
Public Function GenerarSal() As String
    On Error GoTo Manejador
    Dim i As Long, s As String
    Const HEX As String = "0123456789abcdef"
    Randomize
    s = ""
    For i = 1 To 32
        s = s & Mid$(HEX, Int(Rnd() * 16) + 1, 1)
    Next i
    GenerarSal = s
    Exit Function
Manejador:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' Crea un usuario con hash+sal (para el sub-paso 2B; aqui centralizado).
Public Sub CrearUsuario(conn As Object, nombre As String, correo As String, _
                        clave As String, rol As String)
    On Error GoTo Manejador
    Dim cmd As Object
    Dim sal As String

    sal = GenerarSal()

    Set cmd = modConexion.NuevoComando(conn, _
        "INSERT INTO USUARIOS (nombre, correo, hash_contrasena, sal, rol) " & _
        "VALUES (?, ?, ?, ?, ?)")
    modConexion.ParamTexto cmd, "pNombre", nombre, 100
    modConexion.ParamTexto cmd, "pCorreo", correo, 255
    modConexion.ParamTexto cmd, "pHash", SHA256Hex(sal & clave), 255
    modConexion.ParamTexto cmd, "pSal", sal, 64
    modConexion.ParamTexto cmd, "pRol", rol, 20
    cmd.Execute
    Exit Sub
Manejador:
    Err.Raise Err.Number, Err.Source, "Error al crear usuario: " & Err.Description
End Sub

