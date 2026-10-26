' modTestLogin
Option Explicit
Public Sub TestLogin()
    On Error GoTo Manejador
    Dim conn As Object
    Dim rol As String, idUsuario As Long
    Dim auth As Boolean
    Dim hoja As Worksheet
    Dim fila As Long
    
    ' Delete the TestResults sheet if it exists to avoid errors
    On Error Resume Next
    ThisWorkbook.Worksheets("TestResults").Delete
    On Error GoTo Manejador
    
    Set hoja = ThisWorkbook.Worksheets.Add
    hoja.Name = "TestResults"
    hoja.Range("A1").Value = "Email"
    hoja.Range("B1").Value = "Password"
    hoja.Range("C1").Value = "Auth"
    hoja.Range("D1").Value = "Role"
    hoja.Range("E1").Value = "ID"
    fila = 2
    
    Set conn = modConexion.AbrirConexion
    
    ' Test admin
    rol = "": idUsuario = 0
    auth = modSeguridad.Autenticar(conn, "admin.demo@example.com", "Admin123!", rol, idUsuario)
    hoja.Range("A" & fila).Value = "admin.demo@example.com"
    hoja.Range("B" & fila).Value = "Admin123!"
    hoja.Range("C" & fila).Value = auth
    hoja.Range("D" & fila).Value = rol
    hoja.Range("E" & fila).Value = idUsuario
    fila = fila + 1
    
    ' Test vendedor
    rol = "": idUsuario = 0
    auth = modSeguridad.Autenticar(conn, "vendedor.demo@example.com", "Vendedor123!", rol, idUsuario)
    hoja.Range("A" & fila).Value = "vendedor.demo@example.com"
    hoja.Range("B" & fila).Value = "Vendedor123!"
    hoja.Range("C" & fila).Value = auth
    hoja.Range("D" & fila).Value = rol
    hoja.Range("E" & fila).Value = idUsuario
    fila = fila + 1
    
    ' Test cliente
    rol = "": idUsuario = 0
    auth = modSeguridad.Autenticar(conn, "cliente.demo@example.com", "Cliente123!", rol, idUsuario)
    hoja.Range("A" & fila).Value = "cliente.demo@example.com"
    hoja.Range("B" & fila).Value = "Cliente123!"
    hoja.Range("C" & fila).Value = auth
    hoja.Range("D" & fila).Value = rol
    hoja.Range("E" & fila).Value = idUsuario
    fila = fila + 1
    
    ' Test bad password
    rol = "": idUsuario = 0
    auth = modSeguridad.Autenticar(conn, "admin.demo@example.com", "ClaveMala", rol, idUsuario)
    hoja.Range("A" & fila).Value = "admin.demo@example.com"
    hoja.Range("B" & fila).Value = "ClaveMala"
    hoja.Range("C" & fila).Value = auth
    hoja.Range("D" & fila).Value = rol
    hoja.Range("E" & fila).Value = idUsuario
    fila = fila + 1
    
    ' Test non-existent email
    rol = "": idUsuario = 0
    auth = modSeguridad.Autenticar(conn, "noexiste@example.com", "X", rol, idUsuario)
    hoja.Range("A" & fila).Value = "noexiste@example.com"
    hoja.Range("B" & fila).Value = "X"
    hoja.Range("C" & fila).Value = auth
    hoja.Range("D" & fila).Value = rol
    hoja.Range("E" & fila).Value = idUsuario
    fila = fila + 1
    
    modConexion.CerrarConexion conn
    hoja.Columns("A:E").AutoFit
    Exit Sub
Manejador:
    modConexion.CerrarConexion Nothing
    ' Write error to the sheet instead of MsgBox
    If Not hoja Is Nothing Then
        hoja.Range("A1").Value = "Error: " & Err.Description
    End If
End Sub
Attribute VB_Name = "modTestLogin"