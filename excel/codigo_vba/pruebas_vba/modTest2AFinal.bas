
Option Explicit
' modTest2AFinal
Public Sub EjecutarPruebas()
    On Error GoTo Manejador
    Dim conn As Object
    Dim rol As String, idUsuario As Long
    Dim auth As Boolean
    Dim hoja As Worksheet
    Dim fila As Long
    Dim salt As String, clave As String, hashVBA As String
    
    ' Delete the TestResults sheet if it exists to avoid errors
    On Error Resume Next
    ThisWorkbook.Worksheets("TestResults").Delete
    On Error GoTo Manejador
    
    Set hoja = ThisWorkbook.Worksheets.Add
    hoja.Name = "TestResults"
    hoja.Range("A1").Value = "Test"
    hoja.Range("B1").Value = "Details"
    hoja.Range("C1").Value = "Result"
    fila = 2
    
    Set conn = modConexion.AbrirConexion
    
    ' Test 1: admin correct
    rol = "": idUsuario = 0
    auth = modSeguridad.Autenticar(conn, "admin.demo@example.com", "Admin123!", rol, idUsuario)
    hoja.Range("A" & fila).Value = "Login admin correcto"
    hoja.Range("B" & fila).Value = "admin.demo@example.com / Admin123!"
    hoja.Range("C" & fila).Value = IIf(auth, "OK", "FAIL")
    fila = fila + 1
    
    ' Test 2: vendedor correct
    rol = "": idUsuario = 0
    auth = modSeguridad.Autenticar(conn, "vendedor.demo@example.com", "Vendedor123!", rol, idUsuario)
    hoja.Range("A" & fila).Value = "Login vendedor correcto"
    hoja.Range("B" & fila).Value = "vendedor.demo@example.com / Vendedor123!"
    hoja.Range("C" & fila).Value = IIf(auth, "OK", "FAIL")
    fila = fila + 1
    
    ' Test 3: cliente correct
    rol = "": idUsuario = 0
    auth = modSeguridad.Autenticar(conn, "cliente.demo@example.com", "Cliente123!", rol, idUsuario)
    hoja.Range("A" & fila).Value = "Login cliente correcto"
    hoja.Range("B" & fila).Value = "cliente.demo@example.com / Cliente123!"
    hoja.Range("C" & fila).Value = IIf(auth, "OK", "FAIL")
    fila = fila + 1
    
    ' Test 4: admin wrong password
    rol = "": idUsuario = 0
    auth = modSeguridad.Autenticar(conn, "admin.demo@example.com", "ClaveMala", rol, idUsuario)
    hoja.Range("A" & fila).Value = "Login admin password incorrecta"
    hoja.Range("B" & fila).Value = "admin.demo@example.com / ClaveMala"
    hoja.Range("C" & fila).Value = IIf(Not auth, "OK (failed as expected)", "FAIL")
    fila = fila + 1
    
    ' Test 5: non-existent email
    rol = "": idUsuario = 0
    auth = modSeguridad.Autenticar(conn, "noexiste@example.com", "X", rol, idUsuario)
    hoja.Range("A" & fila).Value = "Login email inexistente"
    hoja.Range("B" & fila).Value = "noexiste@example.com / X"
    hoja.Range("C" & fila).Value = IIf(Not auth, "OK (failed as expected)", "FAIL")
    fila = fila + 1
    
    ' Hash test with real user salt
    ' Get salt and stored hash for admin user
    Dim cmd As Object, rs As Object
    Set cmd = modConexion.NuevoComando(conn, "SELECT sal, clave FROM USUARIOS WHERE email = ?")
    modConexion.ParamTexto cmd, "email", "admin.demo@example.com"
    Set rs = cmd.Execute
    If Not rs.EOF Then
        salt = rs.Fields("sal").Value
        clave = rs.Fields("clave").Value  ' this is the stored hash (salt+password)
        ' Compute hash of salt + "Admin123!" using VBA
        hashVBA = modSeguridad.SHA256Hex(salt & "Admin123!")
        hoja.Range("A" & fila).Value = "Hash test (salt + password)"
        hoja.Range("B" & fila).Value = "Salt: " & Left(salt, 10) & "... , Password: Admin123!"
        hoja.Range("C" & fila).Value = "VBA SHA256: " & hashVBA
        fila = fila + 1
        hoja.Range("A" & fila).Value = "Stored hash in DB"
        hoja.Range("B" & fila).Value = "Clave: " & clave
        hoja.Range("C" & fila).Value = "Match: " & IIf(hashVBA = clave, "YES", "NO")
        fila = fila + 1
    Else
        hoja.Range("A" & fila).Value = "Error: usuario admin no encontrado"
        hoja.Range("B" & fila).Value = ""
        hoja.Range("C" & fila).Value = ""
        fila = fila + 1
    End If
    rs.Close
    modConexion.CerrarConexion conn
    
    hoja.Columns("A:C").AutoFit
    Exit Sub
Manejador:
    modConexion.CerrarConexion Nothing
    MsgBox "Error en EjecutarPruebas: " & Err.Description, vbCritical
End Sub
Attribute VB_Name = "modTest2AFinal
