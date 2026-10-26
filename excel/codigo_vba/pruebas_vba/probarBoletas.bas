Attribute VB_Name = "modPruebaBoletas"
Option Explicit

' ============================================================================
' modPruebaBoletas - SOLO PRUEBAS (FASE 3A y 3B). No queda en el libro final.
' Macro: ProbarBoletasPDF
'
' Cubre en una sola corrida:
'   - FASE 3A: venta valida (PDF en disco + UPDATE ruta_pdf), regeneracion,
'              carpeta inexistente, venta anulada, venta inexistente.
'   - FASE 3B: redireccion segura de @example.com -> CorreoPrueba, bloqueo si
'              CorreoPrueba esta vacio o es @example.com, auto-generacion de PDF
'              al enviar si faltaba, actualizacion de BOLETAS.enviada_correo = 'S',
'              y rechazo de envio en venta anulada o inexistente.
'   - Limpieza total: elimina el PDF de prueba y hace RollbackTrans.
' ============================================================================

Private Const CADENA As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Proyectos\TiendaFutbol\access\tienda_futbol.accdb;"
Private Const RUTA_SALIDA As String = "D:\Proyectos\TiendaFutbol\docs\salida_pruebas_3a.txt"

Private mOk As Long
Private mFalla As Long
Private mLog As String

Private Sub RegistrarLinea(ByVal texto As String)
    Debug.Print texto
    mLog = mLog & texto & vbCrLf
End Sub

Private Sub Chequear(ByVal etiqueta As String, ByVal obtenido As Boolean, ByVal esperado As Boolean)
    If obtenido = esperado Then
        mOk = mOk + 1
        RegistrarLinea "OK    | " & etiqueta
    Else
        mFalla = mFalla + 1
        RegistrarLinea "FALLA | " & etiqueta & " | esperado=" & esperado & " obtenido=" & obtenido
    End If
End Sub

Private Function Contar(ByVal conn As Object, ByVal tabla As String) As Long
    Dim rs As Object
    Set rs = conn.Execute("SELECT COUNT(*) FROM " & tabla)
    Contar = CLng(rs.Fields(0).Value)
    rs.Close
End Function

Private Sub GuardarArchivoSalida()
    On Error Resume Next
    Dim f As Integer
    f = FreeFile
    Open RUTA_SALIDA For Output As #f
    Print #f, mLog;
    Close #f
    On Error GoTo 0
End Sub

Public Sub ProbarBoletasPDF()
    On Error GoTo Fallo
    Dim conn As Object, rs As Object, cmd As Object, fso As Object
    Dim wsCfg As Worksheet
    Dim abierta As Boolean
    Dim rutaOriginal As String, correoPruebaOriginal As String
    Dim nVentasAntes As Long, nBoletasAntes As Long, nConPdfAntes As Long, nEnviadasAntes As Long
    Dim idAdmin As Long, idVendedor As Long, idCliente As Long, idProv As Long
    Dim sku1 As String, sku2 As String
    Dim skus(0 To 1) As String, cants(0 To 1) As Long
    Dim idVenta As Long, idBoleta As Long, idMov As Long
    Dim corr As String, rutaPdf1 As String, rutaPdf2 As String, rutaFallo As String
    Dim destUsado As String, estadoEnvio As String
    Dim msg As String, ok As Boolean
    Dim rutaGuardada As String

    mOk = 0
    mFalla = 0
    mLog = ""

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set wsCfg = ThisWorkbook.Worksheets(modConexion.HOJA_CONFIG)
    rutaOriginal = Trim$(CStr(wsCfg.Range(modBoletas.NOMBRE_RANGO_BOLETAS).Value))
    On Error Resume Next
    correoPruebaOriginal = Trim$(CStr(wsCfg.Range(modBoletas.NOMBRE_RANGO_CORREO_PRUEBA).Value))
    On Error GoTo Fallo

    Set conn = CreateObject("ADODB.Connection")
    conn.Open CADENA

    nVentasAntes = Contar(conn, "VENTAS")
    nBoletasAntes = Contar(conn, "BOLETAS")
    Set rs = conn.Execute("SELECT COUNT(*) FROM BOLETAS WHERE ruta_pdf IS NOT NULL")
    nConPdfAntes = CLng(rs.Fields(0).Value): rs.Close
    Set rs = conn.Execute("SELECT COUNT(*) FROM BOLETAS WHERE enviada_correo = 'S'")
    nEnviadasAntes = CLng(rs.Fields(0).Value): rs.Close

    RegistrarLinea "Antes: VENTAS=" & nVentasAntes & " | BOLETAS=" & nBoletasAntes & _
                   " | con ruta_pdf=" & nConPdfAntes & " | enviadas='S'=" & nEnviadasAntes
    RegistrarLinea "Carpeta RutaBoletas: " & rutaOriginal

    Set rs = conn.Execute("SELECT TOP 1 id FROM USUARIOS WHERE rol = 'Admin' ORDER BY id")
    idAdmin = CLng(rs.Fields(0).Value): rs.Close
    Set rs = conn.Execute("SELECT TOP 1 id FROM USUARIOS WHERE rol = 'Vendedor' ORDER BY id")
    idVendedor = CLng(rs.Fields(0).Value): rs.Close
    Set rs = conn.Execute("SELECT TOP 1 id FROM CLIENTES ORDER BY id")
    idCliente = CLng(rs.Fields(0).Value): rs.Close
    Set rs = conn.Execute("SELECT TOP 1 id FROM PROVEEDORES ORDER BY id")
    idProv = CLng(rs.Fields(0).Value): rs.Close
    Set rs = conn.Execute("SELECT TOP 2 sku FROM PRODUCTOS ORDER BY sku")
    sku1 = CStr(rs.Fields("sku").Value)
    rs.MoveNext
    sku2 = CStr(rs.Fields("sku").Value)
    rs.Close

    conn.BeginTrans
    abierta = True

    idMov = modInventario.RegistrarMovimientoInventario(conn, sku1, idAdmin, "E", 5, "Stock prueba 3A/3B", idProv, msg, False)
    idMov = modInventario.RegistrarMovimientoInventario(conn, sku2, idAdmin, "E", 5, "Stock prueba 3A/3B", idProv, msg, False)

    skus(0) = sku1: cants(0) = 2
    skus(1) = sku2: cants(1) = 1
    idVenta = modVentas.RegistrarVenta(conn, idVendedor, idCliente, skus, cants, msg, idBoleta, corr, False)
    Chequear "Precondicion: venta de prueba creada (id=" & idVenta & ", " & corr & ")", (idVenta > 0), True

    RegistrarLinea "--- 1. FASE 3A - Venta valida: generacion de PDF y UPDATE en BOLETAS ---"
    ok = modBoletas.GenerarBoletaPDF(conn, idVenta, rutaPdf1, msg)
    Chequear "GenerarBoletaPDF en venta valida devuelve True", ok, True
    Chequear "GenerarBoletaPDF no deja mensaje de error", (Len(msg) = 0), True
    Chequear "Ruta devuelta termina en " & corr & ".pdf", _
        (Right$(UCase$(rutaPdf1), Len(corr) + 4) = UCase$(corr & ".pdf")), True
    Chequear "El archivo PDF existe en disco", fso.FileExists(rutaPdf1), True
    If fso.FileExists(rutaPdf1) Then
        Chequear "El archivo PDF tiene tamano > 0 bytes", (FileLen(rutaPdf1) > 0), True
    Else
        Chequear "El archivo PDF tiene tamano > 0 bytes", False, True
    End If

    Set cmd = modConexion.NuevoComando(conn, "SELECT ruta_pdf FROM BOLETAS WHERE id_venta = ?")
    modConexion.ParamEntero cmd, "pIdVenta", idVenta
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open cmd, , 3, 1
    If Not rs.EOF And Not IsNull(rs.Fields("ruta_pdf").Value) Then
        rutaGuardada = CStr(rs.Fields("ruta_pdf").Value)
    Else
        rutaGuardada = ""
    End If
    rs.Close
    Chequear "BOLETAS.ruta_pdf se actualizo en la BD", (rutaGuardada = rutaPdf1), True
    RegistrarLinea "     PDF generado: " & rutaPdf1

    RegistrarLinea "--- 2. FASE 3A - Regenerar una boleta ya generada ---"
    ok = modBoletas.GenerarBoletaPDF(conn, idVenta, rutaPdf2, msg)
    Chequear "Regenerar boleta existente devuelve True", ok, True
    Chequear "Regenerar mantiene la misma ruta " & corr & ".pdf", (rutaPdf2 = rutaPdf1), True
    Chequear "El PDF regenerado existe y pesa > 0 bytes", _
        (fso.FileExists(rutaPdf2) And FileLen(rutaPdf2) > 0), True

    RegistrarLinea "--- 3. FASE 3A - Carpeta de destino inexistente ---"
    wsCfg.Range(modBoletas.NOMBRE_RANGO_BOLETAS).Value = "D:\Proyectos\TiendaFutbol\carpeta_inexistente_3a_test"
    ok = modBoletas.GenerarBoletaPDF(conn, idVenta, rutaFallo, msg)
    wsCfg.Range(modBoletas.NOMBRE_RANGO_BOLETAS).Value = rutaOriginal
    Chequear "Carpeta inexistente devuelve False", ok, False
    Chequear "Carpeta inexistente deja rutaPdf vacia", (Len(rutaFallo) = 0), True
    Chequear "Carpeta inexistente informa motivo claro", (InStr(1, msg, "carpeta", vbTextCompare) > 0), True
    RegistrarLinea "     motivo: " & msg

    RegistrarLinea "--- 4. FASE 3B - Envio por correo, redireccion @example.com y proteccion anti-spam ---"
    wsCfg.Range(modBoletas.NOMBRE_RANGO_CORREO_PRUEBA).Value = "m.coroseo.ramos@gmail.com"
    ok = modBoletas.EnviarBoletaPorCorreo(conn, idVenta, destUsado, msg, True)
    Chequear "EnviarBoletaPorCorreo con CorreoPrueba valido devuelve True", ok, True
    Chequear "Cliente @example.com se redirige a CorreoPrueba (m.coroseo.ramos@gmail.com)", _
        (destUsado = "m.coroseo.ramos@gmail.com"), True
    RegistrarLinea "     envio: " & msg

    Set cmd = modConexion.NuevoComando(conn, "SELECT enviada_correo FROM BOLETAS WHERE id_venta = ?")
    modConexion.ParamEntero cmd, "pIdVenta", idVenta
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open cmd, , 3, 1
    estadoEnvio = Trim$(CStr(rs.Fields("enviada_correo").Value))
    rs.Close
    Chequear "BOLETAS.enviada_correo se actualizo a 'S'", (estadoEnvio = "S"), True

    ' Bloqueo si CorreoPrueba esta vacio para cliente @example.com
    wsCfg.Range(modBoletas.NOMBRE_RANGO_CORREO_PRUEBA).Value = ""
    ok = modBoletas.EnviarBoletaPorCorreo(conn, idVenta, destUsado, msg, True)
    Chequear "Cliente @example.com sin CorreoPrueba se bloquea (False)", ok, False
    Chequear "Bloqueo deja destinatario vacio", (Len(destUsado) = 0), True
    RegistrarLinea "     motivo bloqueo vacio: " & msg

    ' Bloqueo si CorreoPrueba tambien es @example.com
    wsCfg.Range(modBoletas.NOMBRE_RANGO_CORREO_PRUEBA).Value = "prohibido@example.com"
    ok = modBoletas.EnviarBoletaPorCorreo(conn, idVenta, destUsado, msg, True)
    Chequear "CorreoPrueba con dominio @example.com se bloquea (False)", ok, False
    RegistrarLinea "     motivo bloqueo @example.com: " & msg
    wsCfg.Range(modBoletas.NOMBRE_RANGO_CORREO_PRUEBA).Value = correoPruebaOriginal

    ' Auto-generacion de PDF si se borra antes de enviar
    If fso.FileExists(rutaPdf1) Then Kill rutaPdf1
    conn.Execute "UPDATE BOLETAS SET ruta_pdf = NULL, enviada_correo = 'N' WHERE id_venta = " & idVenta
    wsCfg.Range(modBoletas.NOMBRE_RANGO_CORREO_PRUEBA).Value = "m.coroseo.ramos@gmail.com"
    ok = modBoletas.EnviarBoletaPorCorreo(conn, idVenta, destUsado, msg, True)
    wsCfg.Range(modBoletas.NOMBRE_RANGO_CORREO_PRUEBA).Value = correoPruebaOriginal
    Chequear "EnviarBoletaPorCorreo regenera automaticamente el PDF si faltaba", _
        (ok And fso.FileExists(rutaPdf1)), True

    RegistrarLinea "--- 5. FASE 3A/3B - Venta anulada y venta inexistente ---"
    Chequear "Anular la venta de prueba devuelve True", _
        modVentas.AnularVenta(conn, idVenta, idAdmin, msg, False), True
    ok = modBoletas.GenerarBoletaPDF(conn, idVenta, rutaFallo, msg)
    Chequear "GenerarBoletaPDF en venta anulada devuelve False", ok, False
    Chequear "Venta anulada deja rutaPdf vacia", (Len(rutaFallo) = 0), True
    Chequear "Venta anulada informa que esta anulada", (InStr(1, msg, "anulada", vbTextCompare) > 0), True

    ok = modBoletas.EnviarBoletaPorCorreo(conn, idVenta, destUsado, msg, True)
    Chequear "EnviarBoletaPorCorreo en venta anulada devuelve False", ok, False

    ok = modBoletas.GenerarBoletaPDF(conn, 999999, rutaFallo, msg)
    Chequear "GenerarBoletaPDF en venta inexistente devuelve False", ok, False
    Chequear "Venta inexistente deja rutaPdf vacia", (Len(rutaFallo) = 0), True
    Chequear "Venta inexistente informa que no existe", (InStr(1, msg, "no existe", vbTextCompare) > 0), True

    ok = modBoletas.EnviarBoletaPorCorreo(conn, 999999, destUsado, msg, True)
    Chequear "EnviarBoletaPorCorreo en venta inexistente devuelve False", ok, False

    RegistrarLinea "--- 6. Limpieza de PDF de prueba y RollbackTrans ---"
    If Len(rutaPdf1) > 0 And fso.FileExists(rutaPdf1) Then
        Kill rutaPdf1
    End If
    Chequear "PDF temporal de prueba eliminado del disco", Not fso.FileExists(rutaPdf1), True

    ThisWorkbook.Worksheets(modBoletas.HOJA_PLANTILLA_BOLETA).Cells.Clear
    conn.RollbackTrans
    abierta = False

    Chequear "VENTAS vuelve al conteo inicial (" & nVentasAntes & ")", (Contar(conn, "VENTAS") = nVentasAntes), True
    Chequear "BOLETAS vuelve al conteo inicial (" & nBoletasAntes & ")", (Contar(conn, "BOLETAS") = nBoletasAntes), True
    Set rs = conn.Execute("SELECT COUNT(*) FROM BOLETAS WHERE ruta_pdf IS NOT NULL")
    Chequear "BOLETAS con ruta_pdf vuelve a " & nConPdfAntes, (CLng(rs.Fields(0).Value) = nConPdfAntes), True
    rs.Close
    Set rs = conn.Execute("SELECT COUNT(*) FROM BOLETAS WHERE enviada_correo = 'S'")
    Chequear "BOLETAS con enviada_correo='S' vuelve a " & nEnviadasAntes, (CLng(rs.Fields(0).Value) = nEnviadasAntes), True
    rs.Close

    RegistrarLinea "=== RESULTADO: OK=" & mOk & " | FALLAS=" & mFalla & " ==="
    conn.Close
    GuardarArchivoSalida
    Exit Sub
Fallo:
    RegistrarLinea "ERROR " & Err.Number & ": " & Err.Description
    On Error Resume Next
    wsCfg.Range(modBoletas.NOMBRE_RANGO_BOLETAS).Value = rutaOriginal
    wsCfg.Range(modBoletas.NOMBRE_RANGO_CORREO_PRUEBA).Value = correoPruebaOriginal
    If Len(rutaPdf1) > 0 And fso.FileExists(rutaPdf1) Then Kill rutaPdf1
    If abierta Then conn.RollbackTrans
    conn.Close
    GuardarArchivoSalida
End Sub
