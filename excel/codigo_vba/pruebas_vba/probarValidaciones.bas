Attribute VB_Name = "modPruebaValidaciones"
Option Explicit

' ============================================================================
' modPruebaValidaciones - SOLO PRUEBAS. No va en el libro de produccion.
' Macros (escribe el nombre en la ventana Inmediato, Ctrl+G, y pulsa Enter):
'   ProbarValidaciones : casos fijos de RUT, correo, clave y campo obligatorio
'   ListarEsquema      : columnas, indices y 3 filas de ejemplo (USUARIOS, CLIENTES)
'   ProbarRutsDemo     : valida todos los RUT guardados en CLIENTES
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

Public Sub ProbarValidaciones()
    mOk = 0
    mFalla = 0

    Debug.Print "--- RUT validos ---"
    Chequear "RUT 12.345.678-5", modValidaciones.ValidarRut("12.345.678-5"), True
    Chequear "RUT 12345678-5", modValidaciones.ValidarRut("12345678-5"), True
    Chequear "RUT 123456785 (sin guion)", modValidaciones.ValidarRut("123456785"), True
    Chequear "RUT con espacios", modValidaciones.ValidarRut(" 12.345.678-5 "), True
    Chequear "RUT con K 10.000.013-K", modValidaciones.ValidarRut("10.000.013-K"), True
    Chequear "RUT con k minuscula", modValidaciones.ValidarRut("10000013-k"), True
    Chequear "RUT con DV 0 10.000.004-0", modValidaciones.ValidarRut("10.000.004-0"), True
    Chequear "RUT 11.111.111-1", modValidaciones.ValidarRut("11.111.111-1"), True

    Debug.Print "--- RUT invalidos ---"
    Chequear "DV incorrecto 12.345.678-9", modValidaciones.ValidarRut("12.345.678-9"), False
    Chequear "vacio", modValidaciones.ValidarRut(""), False
    Chequear "texto abc", modValidaciones.ValidarRut("abc"), False
    Chequear "muy corto 1-9", modValidaciones.ValidarRut("1-9"), False
    Chequear "sin DV 12.345.678", modValidaciones.ValidarRut("12.345.678"), False
    Chequear "letra en cuerpo 12.345.67X-5", modValidaciones.ValidarRut("12.345.67X-5"), False
    Chequear "demasiado largo", modValidaciones.ValidarRut("123456789012"), False
    Chequear "K incorrecta 12.345.678-K", modValidaciones.ValidarRut("12.345.678-K"), False

    Debug.Print "--- Digito verificador y formato ---"
    Chequear "DV de 12345678 es 5", (modValidaciones.CalcularDvRut("12345678") = "5"), True
    Chequear "Formato 123456785", (modValidaciones.FormatearRut("123456785") = "12.345.678-5"), True
    Chequear "Formato 1234567-4", (modValidaciones.FormatearRut("1234567-4") = "1.234.567-4"), True
    Chequear "Formato de RUT invalido devuelve vacio", (modValidaciones.FormatearRut("12.345.678-9") = ""), True

    Debug.Print "--- Correos ---"
    Chequear "juan@example.com", modValidaciones.ValidarCorreo("juan@example.com"), True
    Chequear "a.b@sub.example.cl", modValidaciones.ValidarCorreo("a.b@sub.example.cl"), True
    Chequear "vacio", modValidaciones.ValidarCorreo(""), False
    Chequear "sin arroba", modValidaciones.ValidarCorreo("sin-arroba"), False
    Chequear "sin parte local", modValidaciones.ValidarCorreo("@example.com"), False
    Chequear "dominio sin punto a@b", modValidaciones.ValidarCorreo("a@b"), False
    Chequear "con espacio", modValidaciones.ValidarCorreo("a b@example.com"), False
    Chequear "doble arroba", modValidaciones.ValidarCorreo("a@@example.com"), False
    Chequear "puntos seguidos", modValidaciones.ValidarCorreo("a@example..com"), False
    Chequear "dominio empieza con punto", modValidaciones.ValidarCorreo("a@.example.com"), False
    Chequear "termina en punto", modValidaciones.ValidarCorreo("a@example.com."), False
    Chequear "parte local empieza con punto", modValidaciones.ValidarCorreo(".a@example.com"), False

    Debug.Print "--- Claves y campos obligatorios ---"
    Chequear "clave Admin123!", modValidaciones.ValidarClave("Admin123!"), True
    Chequear "clave 12345678", modValidaciones.ValidarClave("12345678"), True
    Chequear "clave corta", modValidaciones.ValidarClave("corta"), False
    Chequear "clave vacia", modValidaciones.ValidarClave(""), False
    Chequear "campo con texto", modValidaciones.ValidarCampoObligatorio("x"), True
    Chequear "campo con espacios", modValidaciones.ValidarCampoObligatorio("   "), False
    Chequear "campo vacio", modValidaciones.ValidarCampoObligatorio(""), False

    Debug.Print "=== RESULTADO: OK=" & mOk & " | FALLAS=" & mFalla & " ==="
End Sub

Public Sub ListarEsquema()
    On Error GoTo Fallo
    Dim conn As Object

    Set conn = CreateObject("ADODB.Connection")
    conn.Open CADENA

    EsquemaTabla conn, "USUARIOS"
    EsquemaTabla conn, "CLIENTES"
    MuestraFilas conn, "CLIENTES", 3

    conn.Close
    Exit Sub
Fallo:
    Debug.Print "ERROR " & Err.Number & ": " & Err.Description
    On Error Resume Next
    conn.Close
End Sub

Public Sub ProbarRutsDemo()
    On Error GoTo Fallo
    Dim conn As Object, rs As Object, colRut As String, v As String
    Dim total As Long, validos As Long, invalidos As Long, mostrados As Long

    Set conn = CreateObject("ADODB.Connection")
    conn.Open CADENA

    colRut = BuscarColumnaRut(conn, "CLIENTES")
    If Len(colRut) = 0 Then
        Debug.Print "No hay ninguna columna con 'rut' en CLIENTES. Ejecuta ListarEsquema y pega el resultado."
        conn.Close
        Exit Sub
    End If
    Debug.Print "Columna de RUT detectada: " & colRut

    ' colRut sale del esquema de la base, no del usuario.
    Set rs = conn.Execute("SELECT [" & colRut & "] FROM CLIENTES")
    Do While Not rs.EOF
        If IsNull(rs.Fields(0).Value) Then v = "" Else v = CStr(rs.Fields(0).Value)
        total = total + 1
        If modValidaciones.ValidarRut(v) Then
            validos = validos + 1
            If validos <= 5 Then Debug.Print "  guardado: " & v & "  ->  " & modValidaciones.FormatearRut(v)
        Else
            invalidos = invalidos + 1
            If mostrados < 10 Then
                Debug.Print "  INVALIDO: '" & v & "'"
                mostrados = mostrados + 1
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close

    Set rs = conn.Execute("SELECT COUNT(*) FROM (SELECT [" & colRut & "] FROM CLIENTES GROUP BY [" & colRut & "] HAVING COUNT(*) > 1) AS t")
    Debug.Print "RUT repetidos en CLIENTES: " & rs.Fields(0).Value
    rs.Close

    Debug.Print "=== RESULTADO: total=" & total & " | validos=" & validos & " | invalidos=" & invalidos & " ==="
    conn.Close
    Exit Sub
Fallo:
    Debug.Print "ERROR " & Err.Number & ": " & Err.Description
    On Error Resume Next
    conn.Close
End Sub

' ---------------------------------------------------------------------------
' Auxiliares privadas
' ---------------------------------------------------------------------------

Private Function BuscarColumnaRut(ByVal conn As Object, ByVal tabla As String) As String
    Dim rs As Object
    Set rs = conn.OpenSchema(4, Array(Empty, Empty, tabla))   ' 4 = adSchemaColumns
    Do While Not rs.EOF
        If InStr(1, CStr(rs.Fields("COLUMN_NAME").Value), "rut", vbTextCompare) > 0 Then
            BuscarColumnaRut = CStr(rs.Fields("COLUMN_NAME").Value)
            rs.Close
            Exit Function
        End If
        rs.MoveNext
    Loop
    rs.Close
End Function

Private Sub EsquemaTabla(ByVal conn As Object, ByVal tabla As String)
    Dim rs As Object, longitud As String

    Debug.Print "=== " & tabla & " : columnas ==="
    Set rs = conn.OpenSchema(4, Array(Empty, Empty, tabla))   ' 4 = adSchemaColumns
    Do While Not rs.EOF
        If IsNull(rs.Fields("CHARACTER_MAXIMUM_LENGTH").Value) Then
            longitud = "-"
        Else
            longitud = CStr(rs.Fields("CHARACTER_MAXIMUM_LENGTH").Value)
        End If
        Debug.Print "  " & rs.Fields("ORDINAL_POSITION").Value & ". " & rs.Fields("COLUMN_NAME").Value & _
            " | tipo=" & NombreTipo(CLng(rs.Fields("DATA_TYPE").Value)) & _
            " | largo=" & longitud & _
            " | acepta nulos=" & rs.Fields("IS_NULLABLE").Value
        rs.MoveNext
    Loop
    rs.Close

    Debug.Print "=== " & tabla & " : indices ==="
    Set rs = conn.OpenSchema(12, Array(Empty, Empty, Empty, Empty, tabla))   ' 12 = adSchemaIndexes
    Do While Not rs.EOF
        Debug.Print "  " & rs.Fields("INDEX_NAME").Value & _
            " | columna=" & rs.Fields("COLUMN_NAME").Value & _
            " | primaria=" & rs.Fields("PRIMARY_KEY").Value & _
            " | unico=" & rs.Fields("UNIQUE").Value
        rs.MoveNext
    Loop
    rs.Close
End Sub

Private Sub MuestraFilas(ByVal conn As Object, ByVal tabla As String, ByVal n As Long)
    Dim rs As Object, i As Long, linea As String

    Debug.Print "=== " & tabla & " : primeras " & n & " filas (datos demo) ==="
    Set rs = conn.Execute("SELECT TOP " & n & " * FROM " & tabla)
    Do While Not rs.EOF
        linea = ""
        For i = 0 To rs.Fields.Count - 1
            linea = linea & rs.Fields(i).Name & "=" & rs.Fields(i).Value & " | "
        Next i
        Debug.Print "  " & linea
        rs.MoveNext
    Loop
    rs.Close
End Sub

Private Function NombreTipo(ByVal codigo As Long) As String
    Select Case codigo
        Case 2:   NombreTipo = "Entero"
        Case 3:   NombreTipo = "Entero largo"
        Case 4:   NombreTipo = "Simple"
        Case 5:   NombreTipo = "Doble"
        Case 6:   NombreTipo = "Moneda"
        Case 7:   NombreTipo = "Fecha"
        Case 11:  NombreTipo = "Si/No"
        Case 17:  NombreTipo = "Byte"
        Case 72:  NombreTipo = "GUID"
        Case 130: NombreTipo = "Texto"
        Case 131: NombreTipo = "Decimal"
        Case 135: NombreTipo = "Fecha/hora"
        Case 202: NombreTipo = "Texto corto"
        Case 203: NombreTipo = "Memo"
        Case Else: NombreTipo = "codigo " & codigo
    End Select
End Function
