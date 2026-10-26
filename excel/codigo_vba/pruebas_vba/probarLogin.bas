Attribute VB_Name = "Módulo1"
Option Explicit

Public Sub ProbarLogin()
    On Error GoTo Fallo
    Dim conn As Object, rs As Object
    Dim correos(0 To 9) As String, roles(0 To 9) As String
    Dim n As Long, i As Long, idU As Long
    Dim rolOut As String, ok As Boolean

    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Proyectos\TiendaFutbol\access\tienda_futbol.accdb;"

    Set rs = conn.Execute("SELECT correo, rol FROM USUARIOS ORDER BY id")
    Do While Not rs.EOF
        correos(n) = CStr(rs.Fields(0).Value)
        roles(n) = CStr(rs.Fields(1).Value)
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close

    Debug.Print "Hash de 'test' (esperado 9f86d081884c...): " & Left$(modSeguridad.SHA256Hex("test"), 12) & "..."

    For i = 0 To n - 1
        ok = modSeguridad.Autenticar(conn, correos(i), roles(i) & "123!", rolOut, idU)
        Debug.Print "Clave correcta   | " & roles(i) & " | acceso=" & ok & " | rol=" & rolOut & " | id=" & idU
    Next i

    ok = modSeguridad.Autenticar(conn, correos(0), "claveMala", rolOut, idU)
    Debug.Print "Clave incorrecta | " & roles(0) & " | acceso=" & ok
    ok = modSeguridad.Autenticar(conn, "noexiste@example.com", "x", rolOut, idU)
    Debug.Print "Correo inexistente | acceso=" & ok

    conn.Close
    Exit Sub
Fallo:
    Debug.Print "ERROR " & Err.Number & ": " & Err.Description
End Sub
