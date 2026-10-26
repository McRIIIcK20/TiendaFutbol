Attribute VB_Name = "modBoletas"
Option Explicit

' ============================================================================
' modBoletas - Generacion de boletas PDF con diseno ejecutivo (FASE 3A)
'              y envio por correo via Gmail SMTP (CDO) / Outlook (FASE 3B)
'
' Rangos en hoja "Configuracion" (modConexion.HOJA_CONFIG):
'   - RutaBoletas    : carpeta donde se guardan los PDF (ej. D:\...\boletas)
'   - CorreoPrueba   : correo real al que se redirigen los clientes @example.com
'   - GmailUsuario   : cuenta Gmail emisora (ej. usuario@gmail.com)
'   - GmailClaveApp  : contrasena de aplicacion de Google (16 letras, nunca en Git)
' ============================================================================

Public Const NOMBRE_RANGO_BOLETAS As String = "RutaBoletas"
Public Const NOMBRE_RANGO_CORREO_PRUEBA As String = "CorreoPrueba"
Public Const NOMBRE_RANGO_GMAIL_USUARIO As String = "GmailUsuario"
Public Const NOMBRE_RANGO_GMAIL_CLAVE As String = "GmailClaveApp"
Public Const HOJA_PLANTILLA_BOLETA As String = "PlantillaBoleta"

Private Const xlTypePDF As Long = 0
Private Const xlQualityStandard As Long = 0
Private Const CDO_SCHEMA As String = "http://schemas.microsoft.com/cdo/configuration/"

' Lee un valor de rango nombrado en la hoja Configuracion sin lanzar error si no existe.
Public Function LeerConfigTexto(ByVal nombreRango As String) As String
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(modConexion.HOJA_CONFIG)
    LeerConfigTexto = Trim$(CStr(ws.Range(nombreRango).Value))
    On Error GoTo 0
End Function

' Devuelve la carpeta configurada en el rango RutaBoletas de la hoja Configuracion.
Public Function ObtenerRutaBoletas() As String
    On Error Resume Next
    Dim carpeta As String, fso As Object, rel As String
    carpeta = LeerConfigTexto(NOMBRE_RANGO_BOLETAS)
    If Right$(carpeta, 1) = "\" Or Right$(carpeta, 1) = "/" Then
        carpeta = Left$(carpeta, Len(carpeta) - 1)
    End If
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(carpeta) > 0 Then
        If fso.FolderExists(carpeta) Then ObtenerRutaBoletas = carpeta: Exit Function
    End If
    ' Si el proyecto se movio de carpeta, usar (y crear si falta) la carpeta boletas junto al libro
    If Len(ThisWorkbook.Path) > 0 Then
        rel = fso.GetAbsolutePathName(ThisWorkbook.Path & Application.PathSeparator & ".." & Application.PathSeparator & "boletas")
        If Not fso.FolderExists(rel) Then fso.CreateFolder rel
        If fso.FolderExists(rel) Then ObtenerRutaBoletas = rel: Exit Function
    End If
    ObtenerRutaBoletas = carpeta
    On Error GoTo 0
End Function

' ============================================================================
' FASE 3A: Generacion de Boleta PDF con diseno visual profesional
' ============================================================================

Public Function GenerarBoletaPDF(ByVal conn As Object, ByVal idVenta As Long, _
        ByRef rutaPdf As String, ByRef mensaje As String) As Boolean
    On Error GoTo Manejador
    Dim cmdCab As Object, rsCab As Object
    Dim cmdDet As Object, rsDet As Object
    Dim cmdUpd As Object
    Dim fso As Object
    Dim ws As Worksheet
    Dim carpeta As String, archivoPdf As String
    Dim estado As String, correlativo As String
    Dim cliNombre As String, cliRut As String, cliCorreo As String, cliDir As String
    Dim vendNombre As String
    Dim fechaVenta As Date
    Dim subtotal As Currency, iva As Currency, total As Currency
    Dim fila As Long, filaInicioItems As Long, filaFinItems As Long
    Dim errExpNum As Long, errExpDesc As String

    GenerarBoletaPDF = False
    rutaPdf = ""
    mensaje = ""

    If idVenta <= 0 Then
        mensaje = "El id de la venta no es valido."
        Exit Function
    End If

    ' 1. Leer cabecera de venta, cliente, vendedor y correlativo
    Set cmdCab = modConexion.NuevoComando(conn, _
        "SELECT v.id, v.fecha, v.subtotal, v.iva, v.total, v.estado, " & _
        "c.nombre AS cli_nombre, c.apellido AS cli_apellido, c.rut AS cli_rut, " & _
        "c.correo AS cli_correo, c.direccion AS cli_direccion, " & _
        "u.nombre AS vend_nombre, b.numero_correlativo " & _
        "FROM (((VENTAS v " & _
        "INNER JOIN CLIENTES c ON v.id_cliente = c.id) " & _
        "INNER JOIN USUARIOS u ON v.id_usuario = u.id) " & _
        "INNER JOIN BOLETAS b ON b.id_venta = v.id) " & _
        "WHERE v.id = ?")
    modConexion.ParamEntero cmdCab, "pIdVenta", idVenta

    Set rsCab = CreateObject("ADODB.Recordset")
    rsCab.Open cmdCab, , 3, 1
    If rsCab.EOF Then
        rsCab.Close
        mensaje = "La venta indicada no existe."
        Exit Function
    End If

    estado = Trim$(CStr(rsCab.Fields("estado").Value))
    If StrComp(estado, "Anulada", vbTextCompare) = 0 Then
        rsCab.Close
        mensaje = "No se puede generar boleta de una venta anulada."
        Exit Function
    End If

    fechaVenta = CDate(rsCab.Fields("fecha").Value)
    subtotal = CCur(rsCab.Fields("subtotal").Value)
    iva = CCur(rsCab.Fields("iva").Value)
    total = CCur(rsCab.Fields("total").Value)
    cliNombre = CStr(rsCab.Fields("cli_nombre").Value) & " " & CStr(rsCab.Fields("cli_apellido").Value)
    cliRut = CStr(rsCab.Fields("cli_rut").Value)
    cliCorreo = CStr(rsCab.Fields("cli_correo").Value)
    If IsNull(rsCab.Fields("cli_direccion").Value) Then
        cliDir = "(Sin direccion registrada)"
    Else
        cliDir = CStr(rsCab.Fields("cli_direccion").Value)
    End If
    vendNombre = CStr(rsCab.Fields("vend_nombre").Value)
    correlativo = Trim$(CStr(rsCab.Fields("numero_correlativo").Value))
    rsCab.Close

    ' 2. Validar carpeta desde rango RutaBoletas
    carpeta = ObtenerRutaBoletas()
    If Len(carpeta) = 0 Then
        mensaje = "La celda '" & NOMBRE_RANGO_BOLETAS & "' en la hoja '" & modConexion.HOJA_CONFIG & "' esta vacia."
        Exit Function
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(carpeta) Then
        mensaje = "La carpeta de destino para boletas no existe: " & carpeta
        Exit Function
    End If

    ' 3. Leer lineas de detalle
    Set cmdDet = modConexion.NuevoComando(conn, _
        "SELECT d.sku_producto, p.nombre AS prod_nombre, d.cantidad, d.precio_unitario, d.subtotal " & _
        "FROM DETALLE_VENTAS d INNER JOIN PRODUCTOS p ON d.sku_producto = p.sku " & _
        "WHERE d.id_venta = ? ORDER BY d.id")
    modConexion.ParamEntero cmdDet, "pIdVenta", idVenta
    Set rsDet = CreateObject("ADODB.Recordset")
    rsDet.Open cmdDet, , 3, 1
    If rsDet.EOF Then
        rsDet.Close
        mensaje = "La venta no posee lineas de detalle."
        Exit Function
    End If

    ' 4. Armar diseno visual en PlantillaBoleta
    Set ws = ThisWorkbook.Worksheets(HOJA_PLANTILLA_BOLETA)
    ws.Visible = -1      ' la hoja esta muy oculta; ExportAsFixedFormat la necesita visible
    ws.Cells.Clear
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 10

    ' Banner superior corporativo (A1:C3) y recuadro tipo timbre electronico (D1:E3)
    ws.Range("A1:C3").Interior.Color = RGB(15, 23, 42)
    ws.Range("A1:C3").Font.Color = RGB(255, 255, 255)
    ws.Range("A1").Value = "TIENDA DE ARTICULOS DE FUTBOL"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 15
    ws.Range("A2").Value = "Equipamiento Deportivo Profesional • Chile"
    ws.Range("A2").Font.Size = 9
    ws.Range("A2").Font.Color = RGB(148, 163, 184)
    ws.Range("A3").Value = "Giro: Venta de Articulos Deportivos"
    ws.Range("A3").Font.Size = 9
    ws.Range("A3").Font.Color = RGB(148, 163, 184)

    ws.Range("D1:E3").Interior.Color = RGB(240, 253, 244)
    ws.Range("D1:E3").Borders.LineStyle = 1
    ws.Range("D1:E3").Borders.Weight = 2
    ws.Range("D1:E3").Borders.Color = RGB(16, 185, 129)
    ws.Range("D1").Value = "R.U.T.: 76.543.210-K"
    ws.Range("D2").Value = "BOLETA ELECTRONICA"
    ws.Range("D3").Value = "N° " & correlativo
    ws.Range("D1:D3").Font.Bold = True
    ws.Range("D1:D3").Font.Color = RGB(6, 95, 70)
    ws.Range("D3").Font.Size = 12

    ' Ficha de datos del cliente y de la venta (A5:E8)
    ws.Range("A5:E8").Interior.Color = RGB(248, 250, 252)
    ws.Range("A5:E8").Borders.LineStyle = 1
    ws.Range("A5:E8").Borders.Color = RGB(203, 213, 225)

    ws.Range("A5").Value = "Cliente:": ws.Range("A5").Font.Bold = True
    ws.Range("B5").Value = cliNombre
    ws.Range("D5").Value = "RUT Cliente:": ws.Range("D5").Font.Bold = True
    ws.Range("E5").Value = cliRut

    ws.Range("A6").Value = "Correo:": ws.Range("A6").Font.Bold = True
    ws.Range("B6").Value = cliCorreo
    ws.Range("D6").Value = "Fecha Emision:": ws.Range("D6").Font.Bold = True
    ws.Range("E6").Value = Format$(fechaVenta, "dd-mm-yyyy")

    ws.Range("A7").Value = "Direccion:": ws.Range("A7").Font.Bold = True
    ws.Range("B7").Value = cliDir
    ws.Range("D7").Value = "ID Venta:": ws.Range("D7").Font.Bold = True
    ws.Range("E7").Value = "#" & idVenta

    ws.Range("A8").Value = "Atendido por:": ws.Range("A8").Font.Bold = True
    ws.Range("B8").Value = vendNombre
    ws.Range("D8").Value = "Estado:": ws.Range("D8").Font.Bold = True
    ws.Range("E8").Value = estado

    ' Encabezado de tabla de productos (Fila 10)
    ws.Range("A10").Value = "SKU"
    ws.Range("B10").Value = "Descripcion del Producto"
    ws.Range("C10").Value = "Cant."
    ws.Range("D10").Value = "P. Unitario"
    ws.Range("E10").Value = "Subtotal"
    ws.Range("A10:E10").Font.Bold = True
    ws.Range("A10:E10").Font.Color = RGB(255, 255, 255)
    ws.Range("A10:E10").Interior.Color = RGB(30, 41, 59)

    filaInicioItems = 11
    fila = filaInicioItems
    Do While Not rsDet.EOF
        If (fila Mod 2) = 0 Then
            ws.Range(ws.Cells(fila, 1), ws.Cells(fila, 5)).Interior.Color = RGB(241, 245, 249)
        End If
        ws.Cells(fila, 1).Value = CStr(rsDet.Fields("sku_producto").Value)
        ws.Cells(fila, 2).Value = CStr(rsDet.Fields("prod_nombre").Value)
        ws.Cells(fila, 3).Value = CLng(rsDet.Fields("cantidad").Value)
        ws.Cells(fila, 4).Value = CCur(rsDet.Fields("precio_unitario").Value)
        ws.Cells(fila, 5).Value = CCur(rsDet.Fields("subtotal").Value)
        ws.Cells(fila, 4).NumberFormat = "$#,##0"
        ws.Cells(fila, 5).NumberFormat = "$#,##0"
        fila = fila + 1
        rsDet.MoveNext
    Loop
    rsDet.Close
    filaFinItems = fila - 1

    ws.Range(ws.Cells(10, 1), ws.Cells(filaFinItems, 5)).Borders.LineStyle = 1
    ws.Range(ws.Cells(10, 1), ws.Cells(filaFinItems, 5)).Borders.Color = RGB(203, 213, 225)

    ' Totales
    fila = fila + 1
    ws.Cells(fila, 4).Value = "Subtotal Neto:": ws.Cells(fila, 4).Font.Bold = True
    ws.Cells(fila, 5).Value = subtotal: ws.Cells(fila, 5).NumberFormat = "$#,##0"

    fila = fila + 1
    ws.Cells(fila, 4).Value = "IVA (19%):": ws.Cells(fila, 4).Font.Bold = True
    ws.Cells(fila, 5).Value = iva: ws.Cells(fila, 5).NumberFormat = "$#,##0"

    fila = fila + 1
    ws.Range(ws.Cells(fila, 4), ws.Cells(fila, 5)).Interior.Color = RGB(209, 250, 229)
    ws.Range(ws.Cells(fila, 4), ws.Cells(fila, 5)).Borders.LineStyle = 1
    ws.Range(ws.Cells(fila, 4), ws.Cells(fila, 5)).Borders.Color = RGB(16, 185, 129)
    ws.Cells(fila, 4).Value = "TOTAL A PAGAR:": ws.Cells(fila, 4).Font.Bold = True
    ws.Cells(fila, 5).Value = total: ws.Cells(fila, 5).Font.Bold = True
    ws.Cells(fila, 5).Font.Size = 11
    ws.Cells(fila, 5).NumberFormat = "$#,##0"

    fila = fila + 2
    ws.Cells(fila, 1).Value = "Gracias por su compra en Tienda de Futbol • Documento generado electronicamente."
    ws.Cells(fila, 1).Font.Italic = True
    ws.Cells(fila, 1).Font.Size = 9
    ws.Cells(fila, 1).Font.Color = RGB(100, 116, 139)

    ws.Columns("A").ColumnWidth = 14
    ws.Columns("B").ColumnWidth = 38
    ws.Columns("C").ColumnWidth = 10
    ws.Columns("D").ColumnWidth = 16
    ws.Columns("E").ColumnWidth = 18

    On Error Resume Next
    ws.PageSetup.Orientation = 1
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = 1
    On Error GoTo Manejador

    ' 5. Exportar a PDF capturando Err.Number ANTES de cambiar el manejador
    archivoPdf = carpeta & Application.PathSeparator & correlativo & ".pdf"

    On Error Resume Next
    If fso.FileExists(archivoPdf) Then Kill archivoPdf
    Err.Clear
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=archivoPdf, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=False
    errExpNum = Err.Number
    errExpDesc = Err.Description
    On Error GoTo Manejador
    ws.Visible = 2       ' volver a ocultarla

    If errExpNum <> 0 Or Not fso.FileExists(archivoPdf) Then
        mensaje = "No se pudo exportar el PDF de la boleta: " & errExpDesc
        Exit Function
    End If

    ' 6. Actualizar BOLETAS.ruta_pdf
    Set cmdUpd = modConexion.NuevoComando(conn, _
        "UPDATE BOLETAS SET ruta_pdf = ? WHERE id_venta = ?")
    modConexion.ParamTexto cmdUpd, "pRuta", archivoPdf, 255
    modConexion.ParamEntero cmdUpd, "pIdVenta", idVenta
    cmdUpd.Execute

    rutaPdf = archivoPdf
    GenerarBoletaPDF = True
    Exit Function
Manejador:
    Dim errN As Long, errS As String, errD As String
    errN = Err.Number: errS = Err.Source: errD = Err.Description
    On Error Resume Next
    ws.Visible = 2
    On Error GoTo 0
    Err.Raise errN, errS, "Error al generar boleta PDF: " & errD
End Function

' ============================================================================
' FASE 3B: Envio de Boleta por Correo (Gmail SMTP via CDO + Fallback Outlook)
' ============================================================================

' Determina el destinatario seguro: si el cliente termina en @example.com,
' redirige al rango CorreoPrueba de la hoja Configuracion (nunca envia a @example.com).
Public Function ResolverDestinatarioCorreo(ByVal correoCliente As String, _
        ByRef destinatarioOut As String, ByRef mensaje As String) As Boolean
    Dim c As String, correoPrueba As String

    ResolverDestinatarioCorreo = False
    destinatarioOut = ""
    mensaje = ""

    c = LCase$(Trim$(correoCliente))
    If Not modValidaciones.ValidarCorreo(c) Then
        mensaje = "El correo del cliente no es valido."
        Exit Function
    End If

    If Right$(c, 12) = "@example.com" Then
        correoPrueba = LCase$(LeerConfigTexto(NOMBRE_RANGO_CORREO_PRUEBA))
        If Len(correoPrueba) = 0 Then
            mensaje = "El cliente usa dominio de prueba (@example.com) y la celda 'CorreoPrueba' en Configuracion esta vacia."
            Exit Function
        End If
        If Not modValidaciones.ValidarCorreo(correoPrueba) Then
            mensaje = "El correo configurado en 'CorreoPrueba' no tiene formato valido."
            Exit Function
        End If
        If Right$(correoPrueba, 12) = "@example.com" Then
            mensaje = "La celda 'CorreoPrueba' no puede usar el dominio reservado @example.com."
            Exit Function
        End If
        destinatarioOut = correoPrueba
    Else
        destinatarioOut = c
    End If

    ResolverDestinatarioCorreo = True
End Function

' Envia la boleta PDF por correo (Gmail CDO o fallback Outlook) y marca BOLETAS.enviada_correo = 'S'.
' Si modoSimulado = True, valida todas las reglas, genera el PDF si falta y actualiza la BD
' sin abrir conexion de red SMTP (ideal para pruebas automatizadas sin spam).
Public Function EnviarBoletaPorCorreo(ByVal conn As Object, ByVal idVenta As Long, _
        ByRef destinatarioUsado As String, ByRef mensaje As String, _
        Optional ByVal modoSimulado As Boolean = False) As Boolean
    On Error GoTo Manejador
    Dim cmd As Object, rs As Object, cmdUpd As Object
    Dim fso As Object
    Dim estado As String, cliNombre As String, cliCorreo As String
    Dim correlativo As String, rutaPdf As String, total As Currency
    Dim asunto As String, cuerpo As String, viaEnvio As String
    Dim errGmail As String, errOutlook As String

    EnviarBoletaPorCorreo = False
    destinatarioUsado = ""
    mensaje = ""

    If idVenta <= 0 Then
        mensaje = "El id de la venta no es valido."
        Exit Function
    End If

    Set cmd = modConexion.NuevoComando(conn, _
        "SELECT v.estado, v.total, c.nombre, c.apellido, c.correo, " & _
        "b.numero_correlativo, b.ruta_pdf " & _
        "FROM (VENTAS v INNER JOIN CLIENTES c ON v.id_cliente = c.id) " & _
        "INNER JOIN BOLETAS b ON b.id_venta = v.id WHERE v.id = ?")
    modConexion.ParamEntero cmd, "pIdVenta", idVenta

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open cmd, , 3, 1
    If rs.EOF Then
        rs.Close
        mensaje = "La venta indicada no existe."
        Exit Function
    End If

    estado = Trim$(CStr(rs.Fields("estado").Value))
    If StrComp(estado, "Anulada", vbTextCompare) = 0 Then
        rs.Close
        mensaje = "No se puede enviar por correo una boleta de una venta anulada."
        Exit Function
    End If

    total = CCur(rs.Fields("total").Value)
    cliNombre = CStr(rs.Fields("nombre").Value) & " " & CStr(rs.Fields("apellido").Value)
    cliCorreo = CStr(rs.Fields("correo").Value)
    correlativo = Trim$(CStr(rs.Fields("numero_correlativo").Value))
    If IsNull(rs.Fields("ruta_pdf").Value) Then
        rutaPdf = ""
    Else
        rutaPdf = Trim$(CStr(rs.Fields("ruta_pdf").Value))
    End If
    rs.Close

    ' Validar y resolver destinatario antes de enviar (bloqueo anti-spam @example.com)
    If Not ResolverDestinatarioCorreo(cliCorreo, destinatarioUsado, mensaje) Then
        Exit Function
    End If

    ' Si el PDF aun no existe en disco, generarlo automaticamente
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(rutaPdf) = 0 Or Not fso.FileExists(rutaPdf) Then
        If Not GenerarBoletaPDF(conn, idVenta, rutaPdf, mensaje) Then
            Exit Function
        End If
    End If

    asunto = "Boleta Electronica " & correlativo & " - Tienda de Futbol"
    cuerpo = "Estimado(a) " & cliNombre & "," & vbCrLf & vbCrLf & _
             "Adjuntamos su Boleta Electronica N° " & correlativo & _
             " correspondiente a su compra en Tienda de Futbol por un total de $" & _
             Format$(total, "#,##0") & " (IVA incluido)." & vbCrLf & vbCrLf & _
             "Atentamente," & vbCrLf & "Tienda de Articulos de Futbol"

    If modoSimulado Then
        viaEnvio = "Simulado"
    Else
        If EnviarViaGmailCDO(destinatarioUsado, asunto, cuerpo, rutaPdf, errGmail) Then
            viaEnvio = "Gmail SMTP (CDO)"
        ElseIf EnviarViaOutlook(destinatarioUsado, asunto, cuerpo, rutaPdf, errOutlook) Then
            viaEnvio = "Outlook"
        Else
            mensaje = "No se pudo enviar el correo. " & errGmail & " | " & errOutlook
            Exit Function
        End If
    End If

    ' Actualizar BOLETAS.enviada_correo = 'S'
    Set cmdUpd = modConexion.NuevoComando(conn, _
        "UPDATE BOLETAS SET enviada_correo = 'S' WHERE id_venta = ?")
    modConexion.ParamEntero cmdUpd, "pIdVenta", idVenta
    cmdUpd.Execute

    mensaje = "Boleta " & correlativo & " enviada a " & destinatarioUsado & " (" & viaEnvio & ")."
    EnviarBoletaPorCorreo = True
    Exit Function
Manejador:
    Err.Raise Err.Number, Err.Source, "Error al enviar boleta por correo: " & Err.Description
End Function

' Envia correo adjuntando PDF mediante CDO.Message autenticado contra smtp.gmail.com:465 (SSL).
Private Function EnviarViaGmailCDO(ByVal destinatario As String, ByVal asunto As String, _
        ByVal cuerpo As String, ByVal rutaAdjunto As String, ByRef errorOut As String) As Boolean
    On Error GoTo Fallo
    Dim usuarioGmail As String, claveApp As String
    Dim msg As Object, cfg As Object

    EnviarViaGmailCDO = False
    errorOut = ""

    usuarioGmail = LeerConfigTexto(NOMBRE_RANGO_GMAIL_USUARIO)
    claveApp = Replace(LeerConfigTexto(NOMBRE_RANGO_GMAIL_CLAVE), " ", "")

    If Len(usuarioGmail) = 0 Or Len(claveApp) = 0 Then
        errorOut = "Gmail CDO: falta 'GmailUsuario' o 'GmailClaveApp' en hoja Configuracion"
        Exit Function
    End If

    Set msg = CreateObject("CDO.Message")
    Set cfg = CreateObject("CDO.Configuration")

    With cfg.Fields
        .Item(CDO_SCHEMA & "sendusing") = 2
        .Item(CDO_SCHEMA & "smtpserver") = "smtp.gmail.com"
        .Item(CDO_SCHEMA & "smtpserverport") = 465
        .Item(CDO_SCHEMA & "smtpauthenticate") = 1
        .Item(CDO_SCHEMA & "smtpusessl") = True
        .Item(CDO_SCHEMA & "sendusername") = usuarioGmail
        .Item(CDO_SCHEMA & "sendpassword") = claveApp
        .Item(CDO_SCHEMA & "smtpconnectiontimeout") = 15
        .Update
    End With

    Set msg.Configuration = cfg
    msg.From = usuarioGmail
    msg.To = destinatario
    msg.Subject = asunto
    msg.TextBody = cuerpo
    msg.AddAttachment rutaAdjunto
    msg.Send

    EnviarViaGmailCDO = True
    Exit Function
Fallo:
    errorOut = "Gmail CDO: " & Err.Description
End Function

' Fallback: envia el correo usando la sesion local de Microsoft Outlook si esta disponible.
Private Function EnviarViaOutlook(ByVal destinatario As String, ByVal asunto As String, _
        ByVal cuerpo As String, ByVal rutaAdjunto As String, ByRef errorOut As String) As Boolean
    On Error GoTo Fallo
    Dim olApp As Object, olMail As Object

    EnviarViaOutlook = False
    errorOut = ""

    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)
    olMail.To = destinatario
    olMail.Subject = asunto
    olMail.Body = cuerpo
    olMail.Attachments.Add rutaAdjunto
    olMail.Send

    EnviarViaOutlook = True
    Exit Function
Fallo:
    errorOut = "Outlook: " & Err.Description
End Function
