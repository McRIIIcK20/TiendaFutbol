Attribute VB_Name = "modVentas"
Option Explicit

' ============================================================================
' modVentas - Motor transaccional de ventas y boletas (con sincronizacion
'             automatica de inventario en BD Access y hojas de Excel)
' ============================================================================

Public Const TASA_IVA As Double = 0.19

' Devuelve el correlativo con formato B-000851, igual que cargar_datos_demo.py.
Public Function FormatearCorrelativoBoleta(ByVal idVenta As Long) As String
    If idVenta < 1000000 Then
        FormatearCorrelativoBoleta = "B-" & Right$("000000" & CStr(idVenta), 6)
    Else
        FormatearCorrelativoBoleta = "B-" & CStr(idVenta)
    End If
End Function

' Registra una venta completa de 1 o mas productos de forma atomica.
' Si usarTransaccionPropia = True, al confirmar la transaccion sincroniza
' automaticamente las hojas "Productos", "Inventario", "Ventas" y "Boletas".
' Si permitirAutocompraCliente = True (usado desde frmCatalogo), permite que un
' usuario con rol "Cliente" compre directamente desde el catalogo descontando stock.
Public Function RegistrarVenta(ByVal conn As Object, ByVal idUsuario As Long, _
        ByVal idCliente As Long, ByRef skus() As String, ByRef cantidades() As Long, _
        ByRef mensaje As String, Optional ByRef idBoletaOut As Long = 0, _
        Optional ByRef correlativoOut As String = "", _
        Optional ByVal usarTransaccionPropia As Boolean = True, _
        Optional ByVal permitirAutocompraCliente As Boolean = False) As Long
    On Error GoTo Manejador
    Dim lb As Long, ub As Long, nItems As Long, i As Long, j As Long
    Dim skuI As String, cantI As Long
    Dim preciosUnit() As Currency, subtotalesLinea() As Currency
    Dim subtotalVenta As Currency, ivaVenta As Currency, totalVenta As Currency
    Dim stockDisp As Long, idVenta As Long, idMov As Long
    Dim rolUsuario As String, msgMov As String
    Dim cmdVenta As Object, cmdDet As Object, cmdBol As Object, rs As Object
    Dim transIniciada As Boolean

    RegistrarVenta = 0
    idBoletaOut = 0
    correlativoOut = ""
    mensaje = ""

    If idUsuario <= 0 Then
        mensaje = "El usuario vendedor es obligatorio."
        Exit Function
    End If
    rolUsuario = ObtenerRolUsuarioPorId(conn, idUsuario)
    If Len(rolUsuario) = 0 Then
        mensaje = "El usuario indicado no existe."
        Exit Function
    End If
    If StrComp(rolUsuario, "Admin", vbTextCompare) <> 0 And _
       StrComp(rolUsuario, "Vendedor", vbTextCompare) <> 0 Then
        If Not (permitirAutocompraCliente And StrComp(rolUsuario, "Cliente", vbTextCompare) = 0) Then
            mensaje = "Su rol (" & rolUsuario & ") no tiene permiso para registrar ventas."
            Exit Function
        End If
    End If

    If idCliente <= 0 Or Not ExisteClientePorId(conn, idCliente) Then
        mensaje = "El cliente indicado no existe."
        Exit Function
    End If

    On Error Resume Next
    lb = LBound(skus)
    ub = UBound(skus)
    If Err.Number <> 0 Or ub < lb Then
        On Error GoTo Manejador
        mensaje = "La venta debe incluir al menos un producto."
        Exit Function
    End If
    If LBound(cantidades) <> lb Or UBound(cantidades) <> ub Then
        On Error GoTo Manejador
        mensaje = "Los arreglos de SKUs y cantidades no tienen el mismo tamano."
        Exit Function
    End If
    On Error GoTo Manejador

    nItems = ub - lb + 1
    ReDim preciosUnit(lb To ub)
    ReDim subtotalesLinea(lb To ub)

    For i = lb To ub
        skuI = UCase$(Trim$(skus(i)))
        cantI = cantidades(i)
        If Len(skuI) = 0 Or Len(skuI) > 50 Then
            mensaje = "Se indico un SKU vacio o invalido en la linea " & (i - lb + 1) & "."
            Exit Function
        End If
        If cantI <= 0 Then
            mensaje = "La cantidad para el SKU " & skuI & " debe ser mayor a 0."
            Exit Function
        End If
        For j = lb To i - 1
            If UCase$(Trim$(skus(j))) = skuI Then
                mensaje = "El SKU " & skuI & " esta repetido en el detalle; agrupe la cantidad en una sola linea."
                Exit Function
            End If
        Next j
    Next i

    If usarTransaccionPropia Then
        On Error Resume Next
        conn.BeginTrans
        transIniciada = (Err.Number = 0)
        On Error GoTo Manejador
    End If

    subtotalVenta = 0
    For i = lb To ub
        skuI = UCase$(Trim$(skus(i)))
        cantI = cantidades(i)

        stockDisp = modInventario.ObtenerStock(conn, skuI)
        If stockDisp < 0 Then
            If transIniciada Then conn.RollbackTrans
            mensaje = "El producto con SKU '" & skuI & "' no existe."
            Exit Function
        End If
        If stockDisp < cantI Then
            If transIniciada Then conn.RollbackTrans
            mensaje = "Stock insuficiente para SKU " & skuI & " (solicitado: " & cantI & ", disponible: " & stockDisp & ")."
            Exit Function
        End If

        preciosUnit(i) = modInventario.ObtenerPrecioVenta(conn, skuI)
        subtotalesLinea(i) = CCur(Round(CDbl(preciosUnit(i)) * CDbl(cantI), 2))
        subtotalVenta = subtotalVenta + subtotalesLinea(i)
    Next i

    ivaVenta = CCur(Round(CDbl(subtotalVenta) * TASA_IVA, 2))
    totalVenta = subtotalVenta + ivaVenta

    Set cmdVenta = modConexion.NuevoComando(conn, _
        "INSERT INTO VENTAS (id_usuario, id_cliente, fecha, subtotal, iva, total, estado) " & _
        "VALUES (?, ?, Date(), ?, ?, ?, 'Completada')")
    modConexion.ParamEntero cmdVenta, "pIdUsuario", idUsuario
    modConexion.ParamEntero cmdVenta, "pIdCliente", idCliente
    modConexion.ParamMoneda cmdVenta, "pSubtotal", subtotalVenta
    modConexion.ParamMoneda cmdVenta, "pIva", ivaVenta
    modConexion.ParamMoneda cmdVenta, "pTotal", totalVenta
    cmdVenta.Execute

    Set rs = conn.Execute("SELECT @@IDENTITY")
    idVenta = CLng(rs.Fields(0).Value)
    rs.Close

    For i = lb To ub
        skuI = UCase$(Trim$(skus(i)))
        cantI = cantidades(i)

        Set cmdDet = modConexion.NuevoComando(conn, _
            "INSERT INTO DETALLE_VENTAS (id_venta, sku_producto, cantidad, precio_unitario, subtotal) " & _
            "VALUES (?, ?, ?, ?, ?)")
        modConexion.ParamEntero cmdDet, "pIdVenta", idVenta
        modConexion.ParamTexto cmdDet, "pSku", skuI, 50
        modConexion.ParamEntero cmdDet, "pCantidad", cantI
        modConexion.ParamMoneda cmdDet, "pPrecioUnit", preciosUnit(i)
        modConexion.ParamMoneda cmdDet, "pSubtotal", subtotalesLinea(i)
        cmdDet.Execute

        ' Descuenta el inventario en PRODUCTOS y registra salida 'S' en MOVIMIENTOS_INVENTARIO
        idMov = modInventario.RegistrarMovimientoInventario(conn, skuI, idUsuario, _
            "S", cantI, "Venta " & CStr(idVenta), 0, msgMov, False)
        If idMov <= 0 Then
            If transIniciada Then conn.RollbackTrans
            mensaje = "No se pudo descontar stock para " & skuI & ": " & msgMov
            Exit Function
        End If
    Next i

    correlativoOut = FormatearCorrelativoBoleta(idVenta)
    Set cmdBol = modConexion.NuevoComando(conn, _
        "INSERT INTO BOLETAS (id_venta, numero_correlativo, ruta_pdf, fecha_emision, enviada_correo) " & _
        "VALUES (?, ?, NULL, Date(), 'N')")
    modConexion.ParamEntero cmdBol, "pIdVenta", idVenta
    modConexion.ParamTexto cmdBol, "pCorrelativo", correlativoOut, 20
    cmdBol.Execute

    Set rs = conn.Execute("SELECT @@IDENTITY")
    idBoletaOut = CLng(rs.Fields(0).Value)
    rs.Close

    If transIniciada Then
        conn.CommitTrans
        ' Sincronizar de inmediato las hojas de Excel para que el descuento de stock
        ' sea visible al instante tanto en "Productos" e "Inventario" como en "Ventas"
        On Error Resume Next
        modInventario.VolcarCatalogoAHoja conn, modSesion.gUsuarioRol
        VolcarVentasAHoja conn
        On Error GoTo Manejador
    End If
    RegistrarVenta = idVenta
    Exit Function
Manejador:
    Dim errNum As Long, errSrc As String, errDesc As String
    errNum = Err.Number: errSrc = Err.Source: errDesc = Err.Description
    If transIniciada Then
        On Error Resume Next
        conn.RollbackTrans
        On Error GoTo 0
    End If
    Err.Raise errNum, errSrc, "Error al registrar venta: " & errDesc
End Function

Public Function RegistrarVentaUnProducto(ByVal conn As Object, ByVal idUsuario As Long, _
        ByVal idCliente As Long, ByVal sku As String, ByVal cantidad As Long, _
        ByRef mensaje As String, Optional ByVal usarTransaccionPropia As Boolean = True) As Long
    Dim skus(0 To 0) As String
    Dim cants(0 To 0) As Long
    skus(0) = sku
    cants(0) = cantidad
    RegistrarVentaUnProducto = RegistrarVenta(conn, idUsuario, idCliente, skus, cants, _
        mensaje, , , usarTransaccionPropia, False)
End Function

Public Function AnularVenta(ByVal conn As Object, ByVal idVenta As Long, _
        ByVal idUsuario As Long, ByRef mensaje As String, _
        Optional ByVal usarTransaccionPropia As Boolean = True) As Boolean
    On Error GoTo Manejador
    Dim cmd As Object, rs As Object, cmdUpd As Object
    Dim estadoActual As String, rolUsuario As String, msgMov As String
    Dim skus() As String, cants() As Long, n As Long, i As Long
    Dim idMov As Long
    Dim transIniciada As Boolean

    AnularVenta = False
    mensaje = ""

    If idVenta <= 0 Then
        mensaje = "El id de la venta no es valido."
        Exit Function
    End If
    rolUsuario = ObtenerRolUsuarioPorId(conn, idUsuario)
    If StrComp(rolUsuario, "Admin", vbTextCompare) <> 0 And _
       StrComp(rolUsuario, "Vendedor", vbTextCompare) <> 0 Then
        mensaje = "Solo Admin o Vendedor pueden anular ventas."
        Exit Function
    End If

    Set cmd = modConexion.NuevoComando(conn, "SELECT estado FROM VENTAS WHERE id = ?")
    modConexion.ParamEntero cmd, "pId", idVenta
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open cmd, , 3, 1
    If rs.EOF Then
        rs.Close
        mensaje = "La venta indicada no existe."
        Exit Function
    End If
    estadoActual = CStr(rs.Fields("estado").Value)
    rs.Close

    If StrComp(estadoActual, "Anulada", vbTextCompare) = 0 Then
        mensaje = "La venta ya se encuentra anulada."
        Exit Function
    End If

    Set cmd = modConexion.NuevoComando(conn, _
        "SELECT sku_producto, cantidad FROM DETALLE_VENTAS WHERE id_venta = ?")
    modConexion.ParamEntero cmd, "pIdVenta", idVenta
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open cmd, , 3, 1
    Do While Not rs.EOF
        ReDim Preserve skus(0 To n)
        ReDim Preserve cants(0 To n)
        skus(n) = CStr(rs.Fields("sku_producto").Value)
        cants(n) = CLng(rs.Fields("cantidad").Value)
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close

    If usarTransaccionPropia Then
        On Error Resume Next
        conn.BeginTrans
        transIniciada = (Err.Number = 0)
        On Error GoTo Manejador
    End If

    Set cmdUpd = modConexion.NuevoComando(conn, "UPDATE VENTAS SET estado = 'Anulada' WHERE id = ?")
    modConexion.ParamEntero cmdUpd, "pId", idVenta
    cmdUpd.Execute

    For i = 0 To n - 1
        idMov = modInventario.RegistrarMovimientoInventario(conn, skus(i), idUsuario, _
            "E", cants(i), "Anulacion de venta " & CStr(idVenta), 0, msgMov, False)
        If idMov <= 0 Then
            If transIniciada Then conn.RollbackTrans
            mensaje = "Error al restituir stock en anulacion: " & msgMov
            Exit Function
        End If
    Next i

    If transIniciada Then
        conn.CommitTrans
        On Error Resume Next
        modInventario.VolcarCatalogoAHoja conn, modSesion.gUsuarioRol
        VolcarVentasAHoja conn
        On Error GoTo Manejador
    End If
    AnularVenta = True
    Exit Function
Manejador:
    Dim errNum As Long, errSrc As String, errDesc As String
    errNum = Err.Number: errSrc = Err.Source: errDesc = Err.Description
    If transIniciada Then
        On Error Resume Next
        conn.RollbackTrans
        On Error GoTo 0
    End If
    Err.Raise errNum, errSrc, "Error al anular venta: " & errDesc
End Function

' Resuelve o crea la ficha en CLIENTES asociada al usuario conectado (para autocompra en Catalogo).
Public Function ObtenerClienteParaSesion(ByVal conn As Object) As Long
    On Error GoTo Fallback
    Dim cmd As Object, rs As Object, correo As String, msg As String, idNuevo As Long

    correo = LCase$(Trim$(modSesion.gUsuarioCorreo))
    If Len(correo) > 0 Then
        Set cmd = modConexion.NuevoComando(conn, "SELECT TOP 1 id FROM CLIENTES WHERE correo = ?")
        modConexion.ParamTexto cmd, "pCorreo", correo, 255
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open cmd, , 3, 1
        If Not rs.EOF Then
            ObtenerClienteParaSesion = CLng(rs.Fields(0).Value)
            rs.Close
            Exit Function
        End If
        rs.Close
    End If

    ' Si es el usuario Cliente Demo y aun no tiene fila en CLIENTES, crearla o usar el primer cliente
    idNuevo = modClientes.CrearFichaClienteParaUsuario(conn, modSesion.gUsuarioId, _
        modSesion.gUsuarioNombre, correo, msg)
    If idNuevo > 0 Then
        ObtenerClienteParaSesion = idNuevo
        Exit Function
    End If
Fallback:
    On Error Resume Next
    Set rs = conn.Execute("SELECT TOP 1 id FROM CLIENTES ORDER BY id")
    If Not rs.EOF Then ObtenerClienteParaSesion = CLng(rs.Fields(0).Value)
    rs.Close
    On Error GoTo 0
End Function

' Vuelca las ultimas 100 ventas y boletas a las hojas "Ventas" y "Boletas".
Public Sub VolcarVentasAHoja(ByVal conn As Object)
    On Error GoTo Manejador
    Dim wsVentas As Worksheet, wsBoletas As Worksheet
    Dim rsV As Object, rsB As Object
    Dim i As Long

    Set wsVentas = ThisWorkbook.Worksheets("Ventas")
    wsVentas.Cells.Clear
    Set rsV = modConexion.Consultar(conn, _
        "SELECT TOP 100 v.id, v.fecha, c.rut AS rut_cliente, u.nombre AS vendedor, " & _
        "v.subtotal, v.iva, v.total, v.estado " & _
        "FROM (VENTAS v INNER JOIN CLIENTES c ON v.id_cliente = c.id) " & _
        "INNER JOIN USUARIOS u ON v.id_usuario = u.id ORDER BY v.id DESC")
    For i = 0 To rsV.Fields.Count - 1
        wsVentas.Cells(1, i + 1).Value = rsV.Fields(i).Name
        wsVentas.Cells(1, i + 1).Font.Bold = True
        wsVentas.Cells(1, i + 1).Interior.Color = RGB(15, 23, 42)
        wsVentas.Cells(1, i + 1).Font.Color = RGB(255, 255, 255)
    Next i
    If Not rsV.EOF Then
        wsVentas.Range("A2").CopyFromRecordset rsV
    End If
    rsV.Close
    wsVentas.Columns.AutoFit

    Set wsBoletas = ThisWorkbook.Worksheets("Boletas")
    wsBoletas.Cells.Clear
    Set rsB = modConexion.Consultar(conn, _
        "SELECT TOP 100 id, id_venta, numero_correlativo, fecha_emision, enviada_correo, ruta_pdf " & _
        "FROM BOLETAS ORDER BY id DESC")
    For i = 0 To rsB.Fields.Count - 1
        wsBoletas.Cells(1, i + 1).Value = rsB.Fields(i).Name
        wsBoletas.Cells(1, i + 1).Font.Bold = True
        wsBoletas.Cells(1, i + 1).Interior.Color = RGB(15, 23, 42)
        wsBoletas.Cells(1, i + 1).Font.Color = RGB(255, 255, 255)
    Next i
    If Not rsB.EOF Then
        wsBoletas.Range("A2").CopyFromRecordset rsB
    End If
    rsB.Close
    wsBoletas.Columns.AutoFit
    Exit Sub
Manejador:
    Err.Raise Err.Number, Err.Source, "Error al volcar ventas a hojas: " & Err.Description
End Sub

Private Function ObtenerRolUsuarioPorId(ByVal conn As Object, ByVal idUsuario As Long) As String
    Dim cmd As Object, rs As Object
    ObtenerRolUsuarioPorId = ""
    Set cmd = modConexion.NuevoComando(conn, "SELECT rol FROM USUARIOS WHERE id = ?")
    modConexion.ParamEntero cmd, "pId", idUsuario
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open cmd, , 3, 1
    If Not rs.EOF Then ObtenerRolUsuarioPorId = CStr(rs.Fields("rol").Value)
    rs.Close
End Function

Private Function ExisteClientePorId(ByVal conn As Object, ByVal idCliente As Long) As Boolean
    Dim cmd As Object, rs As Object
    Set cmd = modConexion.NuevoComando(conn, "SELECT COUNT(*) FROM CLIENTES WHERE id = ?")
    modConexion.ParamEntero cmd, "pId", idCliente
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open cmd, , 3, 1
    ExisteClientePorId = (CLng(rs.Fields(0).Value) > 0)
    rs.Close
End Function
