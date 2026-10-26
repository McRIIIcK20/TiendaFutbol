Attribute VB_Name = "modPruebaVentas"
Option Explicit

' ============================================================================
' modPruebaVentas - SOLO PRUEBAS (FASE 2C-1). No queda en el libro final.
' Macro: ProbarVentasEInventario (ejecutable desde ventana Inmediato o Python)
'
' Todo lo que modifica va dentro de una transaccion que hace RollbackTrans al
' final, dejando intactos los conteos y stocks de tienda_futbol.accdb.
' Ademas de imprimir con Debug.Print, guarda la salida literal en:
'   D:\Proyectos\TiendaFutbol\docs\salida_pruebas_2c1.txt
' ============================================================================

Private Const CADENA As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Proyectos\TiendaFutbol\access\tienda_futbol.accdb;"
Private Const RUTA_SALIDA As String = "D:\Proyectos\TiendaFutbol\docs\salida_pruebas_2c1.txt"

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

Public Sub ProbarVentasEInventario()
    On Error GoTo Fallo
    Dim conn As Object, rs As Object
    Dim abierta As Boolean
    Dim nVentasAntes As Long, nDetAntes As Long, nMovAntes As Long, nBolAntes As Long
    Dim idAdmin As Long, idVendedor As Long, idClienteRol As Long, idClienteCompra As Long, idProv As Long
    Dim sku1 As String, sku2 As String
    Dim stock1Inicial As Long, stock2Inicial As Long
    Dim precio1 As Currency, precio2 As Currency
    Dim nombreProd As String, pVenta As Currency, pCosto As Currency, st As Long, stMin As Long
    Dim rsCatAdmin As Object, rsCatVend As Object
    Dim idMov As Long, idVenta As Long, idVentaFallo As Long, idBoleta As Long
    Dim corr As String, msg As String
    Dim subEsp As Currency, ivaEsp As Currency, totEsp As Currency
    Dim skus(0 To 1) As String, cants(0 To 1) As Long
    Dim skusFallo(0 To 1) As String, cantsFallo(0 To 1) As Long

    mOk = 0
    mFalla = 0
    mLog = ""

    Set conn = CreateObject("ADODB.Connection")
    conn.Open CADENA

    nVentasAntes = Contar(conn, "VENTAS")
    nDetAntes = Contar(conn, "DETALLE_VENTAS")
    nMovAntes = Contar(conn, "MOVIMIENTOS_INVENTARIO")
    nBolAntes = Contar(conn, "BOLETAS")
    RegistrarLinea "Antes: VENTAS=" & nVentasAntes & " | DETALLE_VENTAS=" & nDetAntes & _
                   " | MOVIMIENTOS=" & nMovAntes & " | BOLETAS=" & nBolAntes

    ' Obtener IDs reales de la base demo sin asumir id=1
    Set rs = conn.Execute("SELECT TOP 1 id FROM USUARIOS WHERE rol = 'Admin' ORDER BY id")
    idAdmin = CLng(rs.Fields(0).Value): rs.Close
    Set rs = conn.Execute("SELECT TOP 1 id FROM USUARIOS WHERE rol = 'Vendedor' ORDER BY id")
    idVendedor = CLng(rs.Fields(0).Value): rs.Close
    Set rs = conn.Execute("SELECT TOP 1 id FROM USUARIOS WHERE rol = 'Cliente' ORDER BY id")
    idClienteRol = CLng(rs.Fields(0).Value): rs.Close
    Set rs = conn.Execute("SELECT TOP 1 id FROM CLIENTES ORDER BY id")
    idClienteCompra = CLng(rs.Fields(0).Value): rs.Close
    Set rs = conn.Execute("SELECT TOP 1 id FROM PROVEEDORES ORDER BY id")
    idProv = CLng(rs.Fields(0).Value): rs.Close

    Set rs = conn.Execute("SELECT TOP 2 sku, stock, precio_venta FROM PRODUCTOS ORDER BY sku")
    sku1 = CStr(rs.Fields("sku").Value)
    stock1Inicial = CLng(rs.Fields("stock").Value)
    precio1 = CCur(rs.Fields("precio_venta").Value)
    rs.MoveNext
    sku2 = CStr(rs.Fields("sku").Value)
    stock2Inicial = CLng(rs.Fields("stock").Value)
    precio2 = CCur(rs.Fields("precio_venta").Value)
    rs.Close

    RegistrarLinea "--- 1. Catalogo y control de costos por rol (modInventario) ---"
    Chequear "ExisteProducto con SKU real (" & sku1 & ")", modInventario.ExisteProducto(conn, sku1), True
    Chequear "ExisteProducto en minusculas", modInventario.ExisteProducto(conn, LCase$(sku1)), True
    Chequear "ExisteProducto con SKU inexistente", modInventario.ExisteProducto(conn, "SKU-NO-EXISTE-99"), False
    Chequear "ObtenerStock con SKU inexistente devuelve -1", (modInventario.ObtenerStock(conn, "SKU-NO-EXISTE-99") = -1), True

    Chequear "ObtenerProducto como Admin devuelve True", _
        modInventario.ObtenerProducto(conn, sku1, nombreProd, pVenta, st, stMin, pCosto, "Admin"), True
    Chequear "Admin ve precio_costo > 0", (pCosto > 0), True
    Chequear "precio_venta > precio_costo", (pVenta > pCosto), True

    Chequear "ObtenerProducto como Vendedor devuelve True", _
        modInventario.ObtenerProducto(conn, sku1, nombreProd, pVenta, st, stMin, pCosto, "Vendedor"), True
    Chequear "Vendedor NO ve precio_costo (queda 0)", (pCosto = 0), True

    Chequear "ObtenerProducto como Cliente devuelve True", _
        modInventario.ObtenerProducto(conn, sku1, nombreProd, pVenta, st, stMin, pCosto, "Cliente"), True
    Chequear "Cliente NO ve precio_costo (queda 0)", (pCosto = 0), True

    Set rsCatAdmin = modInventario.ConsultarCatalogoSegunRol(conn, "Admin", False)
    Chequear "Catalogo Admin tiene 60 productos", (rsCatAdmin.RecordCount = 60), True
    Chequear "Catalogo Admin incluye columna precio_costo (9 cols)", (rsCatAdmin.Fields.Count = 9), True
    rsCatAdmin.Close

    Set rsCatVend = modInventario.ConsultarCatalogoSegunRol(conn, "Vendedor", False)
    Chequear "Catalogo Vendedor omite columna precio_costo (8 cols)", (rsCatVend.Fields.Count = 8), True
    rsCatVend.Close

    ' Iniciar transaccion global de prueba (todo se deshace al final)
    conn.BeginTrans
    abierta = True

    RegistrarLinea "--- 2. Movimientos de inventario (E, S, A y rechazo de stock negativo) ---"
    idMov = modInventario.RegistrarMovimientoInventario(conn, sku1, idAdmin, "E", 10, _
        "Reposicion de prueba", idProv, msg, False)
    Chequear "Entrada 'E' (+10) devuelve id > 0", (idMov > 0), True
    Chequear "Stock de " & sku1 & " sube en 10", (modInventario.ObtenerStock(conn, sku1) = stock1Inicial + 10), True

    idMov = modInventario.RegistrarMovimientoInventario(conn, sku2, idAdmin, "E", 5, _
        "Reposicion de prueba SKU2", idProv, msg, False)
    Chequear "Entrada 'E' (+5) en " & sku2 & " devuelve id > 0", (idMov > 0), True
    Chequear "Stock de " & sku2 & " sube en 5", (modInventario.ObtenerStock(conn, sku2) = stock2Inicial + 5), True

    idMov = modInventario.RegistrarMovimientoInventario(conn, sku1, idAdmin, "S", 4, _
        "Salida manual de prueba", 0, msg, False)
    Chequear "Salida 'S' (-4) devuelve id > 0", (idMov > 0), True
    Chequear "Stock de " & sku1 & " queda en inicial + 6", (modInventario.ObtenerStock(conn, sku1) = stock1Inicial + 6), True

    idMov = modInventario.RegistrarMovimientoInventario(conn, sku1, idAdmin, "A", -2, _
        "Ajuste por conteo fisico", 0, msg, False)
    Chequear "Ajuste 'A' (-2) devuelve id > 0", (idMov > 0), True
    Chequear "Stock de " & sku1 & " queda en inicial + 4", (modInventario.ObtenerStock(conn, sku1) = stock1Inicial + 4), True

    idMov = modInventario.RegistrarMovimientoInventario(conn, sku1, idAdmin, "S", stock1Inicial + 500, _
        "Intento de sobregiro", 0, msg, False)
    Chequear "Salida mayor al stock se rechaza (id=0)", (idMov = 0), True
    RegistrarLinea "     motivo: " & msg
    Chequear "Stock de " & sku1 & " no cambio tras rechazo", (modInventario.ObtenerStock(conn, sku1) = stock1Inicial + 4), True

    idMov = modInventario.RegistrarMovimientoInventario(conn, sku1, idAdmin, "E", 0, "Cantidad cero", 0, msg, False)
    Chequear "Movimiento con cantidad 0 se rechaza", (idMov = 0), True
    idMov = modInventario.RegistrarMovimientoInventario(conn, sku1, idAdmin, "X", 2, "Tipo invalido", 0, msg, False)
    Chequear "Movimiento con tipo invalido 'X' se rechaza", (idMov = 0), True

    RegistrarLinea "--- 3. Venta transaccional multi-item (modVentas.RegistrarVenta) ---"
    skus(0) = sku1: cants(0) = 2
    skus(1) = sku2: cants(1) = 3
    idVenta = modVentas.RegistrarVenta(conn, idVendedor, idClienteCompra, skus, cants, msg, idBoleta, corr, False)
    Chequear "RegistrarVenta valida devuelve idVenta > 0", (idVenta > 0), True
    Chequear "RegistrarVenta no deja mensaje de error", (Len(msg) = 0), True
    Chequear "Genera boleta 1:1 con idBoleta > 0", (idBoleta > 0), True
    Chequear "Correlativo tiene formato B-XXXXXX", (corr = modVentas.FormatearCorrelativoBoleta(idVenta)), True
    RegistrarLinea "     idVenta=" & idVenta & " | correlativo=" & corr

    subEsp = CCur(Round(CDbl(precio1) * 2# + CDbl(precio2) * 3#, 2))
    ivaEsp = CCur(Round(CDbl(subEsp) * 0.19, 2))
    totEsp = subEsp + ivaEsp

    Set rs = conn.Execute("SELECT subtotal, iva, total, estado FROM VENTAS WHERE id = " & idVenta)
    Chequear "VENTAS.subtotal cuadra con suma de lineas", (CCur(rs.Fields("subtotal").Value) = subEsp), True
    Chequear "VENTAS.iva cuadra con 19%", (CCur(rs.Fields("iva").Value) = ivaEsp), True
    Chequear "VENTAS.total cuadra con subtotal + iva", (CCur(rs.Fields("total").Value) = totEsp), True
    Chequear "VENTAS.estado es 'Completada'", (CStr(rs.Fields("estado").Value) = "Completada"), True
    rs.Close

    Set rs = conn.Execute("SELECT COUNT(*) FROM DETALLE_VENTAS WHERE id_venta = " & idVenta)
    Chequear "DETALLE_VENTAS tiene 2 lineas para la venta", (CLng(rs.Fields(0).Value) = 2), True
    rs.Close

    Set rs = conn.Execute("SELECT enviada_correo FROM BOLETAS WHERE id_venta = " & idVenta)
    Chequear "BOLETAS registra enviada_correo = 'N'", (CStr(rs.Fields(0).Value) = "N"), True
    rs.Close

    Chequear "Stock de " & sku1 & " descuento 2 unidades", (modInventario.ObtenerStock(conn, sku1) = stock1Inicial + 2), True
    Chequear "Stock de " & sku2 & " descuento 3 unidades", (modInventario.ObtenerStock(conn, sku2) = stock2Inicial + 2), True

    RegistrarLinea "--- 4. Atomicidad ante quiebre de stock en 2da linea del carrito ---"
    skusFallo(0) = sku1: cantsFallo(0) = 1
    skusFallo(1) = sku2: cantsFallo(1) = stock2Inicial + 999
    idVentaFallo = modVentas.RegistrarVenta(conn, idVendedor, idClienteCompra, skusFallo, cantsFallo, msg, , , False)
    Chequear "Venta con quiebre en 2do item se rechaza (id=0)", (idVentaFallo = 0), True
    RegistrarLinea "     motivo: " & msg
    Chequear "Stock del 1er item (" & sku1 & ") NO se desconto", _
        (modInventario.ObtenerStock(conn, sku1) = stock1Inicial + 2), True
    Chequear "No se agregaron ventas tras el rechazo", (Contar(conn, "VENTAS") = nVentasAntes + 1), True

    RegistrarLinea "--- 5. Rechazo por rol Cliente y validaciones de entrada ---"
    idVentaFallo = modVentas.RegistrarVentaUnProducto(conn, idClienteRol, idClienteCompra, sku1, 1, msg, False)
    Chequear "Rol 'Cliente' no puede registrar ventas", (idVentaFallo = 0), True
    RegistrarLinea "     motivo: " & msg
    idVentaFallo = modVentas.RegistrarVentaUnProducto(conn, idVendedor, 999999, sku1, 1, msg, False)
    Chequear "Cliente inexistente se rechaza", (idVentaFallo = 0), True
    idVentaFallo = modVentas.RegistrarVentaUnProducto(conn, idVendedor, idClienteCompra, sku1, 0, msg, False)
    Chequear "Cantidad 0 en venta se rechaza", (idVentaFallo = 0), True

    RegistrarLinea "--- 6. Anulacion de venta y restitucion de stock (modVentas.AnularVenta) ---"
    Chequear "AnularVenta devuelve True", modVentas.AnularVenta(conn, idVenta, idAdmin, msg, False), True
    Set rs = conn.Execute("SELECT estado FROM VENTAS WHERE id = " & idVenta)
    Chequear "Estado de la venta pasa a 'Anulada'", (CStr(rs.Fields("estado").Value) = "Anulada"), True
    rs.Close
    Chequear "Stock de " & sku1 & " restituido tras anulacion", (modInventario.ObtenerStock(conn, sku1) = stock1Inicial + 4), True
    Chequear "Stock de " & sku2 & " restituido tras anulacion", (modInventario.ObtenerStock(conn, sku2) = stock2Inicial + 5), True
    Chequear "Anular una venta ya anulada se rechaza", Not modVentas.AnularVenta(conn, idVenta, idAdmin, msg, False), True

    RegistrarLinea "--- 7. Rollback de la transaccion de pruebas ---"
    conn.RollbackTrans
    abierta = False
    Chequear "VENTAS vuelve al conteo inicial (" & nVentasAntes & ")", (Contar(conn, "VENTAS") = nVentasAntes), True
    Chequear "DETALLE_VENTAS vuelve al conteo inicial (" & nDetAntes & ")", (Contar(conn, "DETALLE_VENTAS") = nDetAntes), True
    Chequear "MOVIMIENTOS vuelve al conteo inicial (" & nMovAntes & ")", (Contar(conn, "MOVIMIENTOS_INVENTARIO") = nMovAntes), True
    Chequear "BOLETAS vuelve al conteo inicial (" & nBolAntes & ")", (Contar(conn, "BOLETAS") = nBolAntes), True
    Chequear "Stock de " & sku1 & " vuelve a " & stock1Inicial, (modInventario.ObtenerStock(conn, sku1) = stock1Inicial), True
    Chequear "Stock de " & sku2 & " vuelve a " & stock2Inicial, (modInventario.ObtenerStock(conn, sku2) = stock2Inicial), True

    RegistrarLinea "=== RESULTADO: OK=" & mOk & " | FALLAS=" & mFalla & " ==="
    conn.Close
    GuardarArchivoSalida
    Exit Sub
Fallo:
    RegistrarLinea "ERROR " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If abierta Then conn.RollbackTrans
    conn.Close
    GuardarArchivoSalida
End Sub
