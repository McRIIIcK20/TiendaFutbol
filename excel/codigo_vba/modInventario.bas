Attribute VB_Name = "modInventario"
Option Explicit

' ============================================================================
' modInventario - Catalogo de productos, imagenes, control de costos y stock
'
' Esquema real verificado:
'   PRODUCTOS: sku (PK, texto 50), nombre (200), id_categoria (Long, FK),
'              talla (50, nulo), marca (100, nulo), precio_costo (Currency >= 0),
'              precio_venta (Currency >= 0), stock (Long >= 0), stock_minimo (Long >= 0)
'   MOVIMIENTOS_INVENTARIO: id (autonumerico), sku_producto (50, FK),
'              id_usuario (Long, FK), id_proveedor (Long, acepta nulo, FK),
'              tipo_movimiento (texto 1: 'E', 'S', 'A'), cantidad (Long <> 0),
'              fecha (Date), motivo (texto 255, acepta nulo)
' ============================================================================

Public Const NOMBRE_RANGO_IMAGENES As String = "RutaImagenes"

' Devuelve la carpeta donde residen las imagenes del catalogo de productos.
Public Function ObtenerRutaImagenes() As String
    On Error Resume Next
    Dim ws As Worksheet, carpeta As String, fso As Object, rutaRel As String
    Set ws = ThisWorkbook.Worksheets(modConexion.HOJA_CONFIG)
    carpeta = Trim$(CStr(ws.Range(NOMBRE_RANGO_IMAGENES).Value))
    On Error GoTo 0

    If Right$(carpeta, 1) = "\" Or Right$(carpeta, 1) = "/" Then
        carpeta = Left$(carpeta, Len(carpeta) - 1)
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(carpeta) > 0 Then
        If fso.FolderExists(carpeta) Then
            ObtenerRutaImagenes = carpeta
            Exit Function
        End If
    End If

    If Len(ThisWorkbook.Path) > 0 Then
        rutaRel = ThisWorkbook.Path & Application.PathSeparator & ".." & _
                  Application.PathSeparator & "imagenes_productos"
        If fso.FolderExists(rutaRel) Then
            ObtenerRutaImagenes = fso.GetAbsolutePathName(rutaRel)
            Exit Function
        End If
    End If
    ObtenerRutaImagenes = carpeta
End Function

' Devuelve la ruta de imagen (.jpg o .bmp) para un SKU o su categoria, o "" si no existe.
Public Function ObtenerRutaImagenProducto(ByVal sku As String, Optional ByVal categoria As String = "") As String
    On Error Resume Next
    Dim carpeta As String, skuLimpio As String, catSlug As String, cand As String
    Dim fso As Object

    ObtenerRutaImagenProducto = ""
    carpeta = ObtenerRutaImagenes()
    If Len(carpeta) = 0 Then Exit Function

    Set fso = CreateObject("Scripting.FileSystemObject")
    skuLimpio = UCase$(Trim$(sku))

    If Len(skuLimpio) > 0 Then
        cand = carpeta & Application.PathSeparator & skuLimpio & ".jpg"
        If fso.FileExists(cand) Then ObtenerRutaImagenProducto = cand: Exit Function
        cand = carpeta & Application.PathSeparator & skuLimpio & ".bmp"
        If fso.FileExists(cand) Then ObtenerRutaImagenProducto = cand: Exit Function
    End If

    catSlug = LCase$(Trim$(categoria))
    If Len(catSlug) = 0 And Len(skuLimpio) >= 3 Then
        Select Case Left$(skuLimpio, 3)
            Case "CAM": catSlug = "camisetas"
            Case "BOT": catSlug = "botines"
            Case "BAL": catSlug = "balones"
            Case "CAN": catSlug = "canilleras"
            Case "GUA": catSlug = "guantes"
            Case "ACC": catSlug = "accesorios"
        End Select
    End If

    If Len(catSlug) > 0 Then
        cand = carpeta & Application.PathSeparator & catSlug & ".bmp"
        If fso.FileExists(cand) Then ObtenerRutaImagenProducto = cand: Exit Function
        cand = carpeta & Application.PathSeparator & catSlug & ".jpg"
        If fso.FileExists(cand) Then ObtenerRutaImagenProducto = cand: Exit Function
    End If
    On Error GoTo 0
End Function

' True si existe un producto con ese SKU.
Public Function ExisteProducto(ByVal conn As Object, ByVal sku As String) As Boolean
    On Error GoTo Manejador
    Dim cmd As Object, rs As Object
    Dim skuLimpio As String

    skuLimpio = UCase$(Trim$(sku))
    If Len(skuLimpio) = 0 Or Len(skuLimpio) > 50 Then Exit Function

    Set cmd = modConexion.NuevoComando(conn, "SELECT COUNT(*) FROM PRODUCTOS WHERE sku = ?")
    modConexion.ParamTexto cmd, "pSku", skuLimpio, 50
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open cmd, , 3, 1
    ExisteProducto = (CLng(rs.Fields(0).Value) > 0)
    rs.Close
    Exit Function
Manejador:
    Err.Raise Err.Number, Err.Source, "Error al verificar SKU: " & Err.Description
End Function

' Devuelve el stock actual de un SKU, o -1 si el SKU no existe.
Public Function ObtenerStock(ByVal conn As Object, ByVal sku As String) As Long
    On Error GoTo Manejador
    Dim cmd As Object, rs As Object
    Dim skuLimpio As String

    ObtenerStock = -1
    skuLimpio = UCase$(Trim$(sku))
    If Len(skuLimpio) = 0 Or Len(skuLimpio) > 50 Then Exit Function

    Set cmd = modConexion.NuevoComando(conn, "SELECT stock FROM PRODUCTOS WHERE sku = ?")
    modConexion.ParamTexto cmd, "pSku", skuLimpio, 50
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open cmd, , 3, 1
    If Not rs.EOF Then
        ObtenerStock = CLng(rs.Fields("stock").Value)
    End If
    rs.Close
    Exit Function
Manejador:
    Err.Raise Err.Number, Err.Source, "Error al consultar stock: " & Err.Description
End Function

' Devuelve el precio de venta vigente de un SKU, o 0 si no existe.
Public Function ObtenerPrecioVenta(ByVal conn As Object, ByVal sku As String) As Currency
    On Error GoTo Manejador
    Dim cmd As Object, rs As Object
    Dim skuLimpio As String

    ObtenerPrecioVenta = 0
    skuLimpio = UCase$(Trim$(sku))
    If Len(skuLimpio) = 0 Or Len(skuLimpio) > 50 Then Exit Function

    Set cmd = modConexion.NuevoComando(conn, "SELECT precio_venta FROM PRODUCTOS WHERE sku = ?")
    modConexion.ParamTexto cmd, "pSku", skuLimpio, 50
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open cmd, , 3, 1
    If Not rs.EOF Then
        ObtenerPrecioVenta = CCur(rs.Fields("precio_venta").Value)
    End If
    rs.Close
    Exit Function
Manejador:
    Err.Raise Err.Number, Err.Source, "Error al consultar precio de venta: " & Err.Description
End Function

' Lee los datos de un producto por su SKU.
' Si rolUsuario no es Admin, precioCostoOut devuelve 0 (oculto por seguridad de rol).
Public Function ObtenerProducto(ByVal conn As Object, ByVal sku As String, _
        ByRef nombreOut As String, ByRef precioVentaOut As Currency, _
        ByRef stockOut As Long, ByRef stockMinimoOut As Long, _
        Optional ByRef precioCostoOut As Currency = 0, _
        Optional ByVal rolUsuario As String = "", _
        Optional ByRef categoriaOut As String = "", _
        Optional ByRef marcaOut As String = "", _
        Optional ByRef tallaOut As String = "") As Boolean
    On Error GoTo Manejador
    Dim cmd As Object, rs As Object
    Dim skuLimpio As String
    Dim rolEfectivo As String

    ObtenerProducto = False
    nombreOut = ""
    precioVentaOut = 0
    stockOut = 0
    stockMinimoOut = 0
    precioCostoOut = 0
    categoriaOut = ""
    marcaOut = ""
    tallaOut = ""

    skuLimpio = UCase$(Trim$(sku))
    If Len(skuLimpio) = 0 Or Len(skuLimpio) > 50 Then Exit Function

    rolEfectivo = Trim$(rolUsuario)
    If Len(rolEfectivo) = 0 Then rolEfectivo = modSesion.gUsuarioRol

    Set cmd = modConexion.NuevoComando(conn, _
        "SELECT p.nombre, c.nombre AS categoria, p.marca, p.talla, p.precio_costo, " & _
        "p.precio_venta, p.stock, p.stock_minimo " & _
        "FROM PRODUCTOS p INNER JOIN CATEGORIAS c ON p.id_categoria = c.id WHERE p.sku = ?")
    modConexion.ParamTexto cmd, "pSku", skuLimpio, 50
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open cmd, , 3, 1
    If rs.EOF Then
        rs.Close
        Exit Function
    End If

    nombreOut = CStr(rs.Fields("nombre").Value)
    categoriaOut = CStr(rs.Fields("categoria").Value)
    If Not IsNull(rs.Fields("marca").Value) Then marcaOut = CStr(rs.Fields("marca").Value)
    If Not IsNull(rs.Fields("talla").Value) Then tallaOut = CStr(rs.Fields("talla").Value) Else tallaOut = "Unica"
    precioVentaOut = CCur(rs.Fields("precio_venta").Value)
    stockOut = CLng(rs.Fields("stock").Value)
    stockMinimoOut = CLng(rs.Fields("stock_minimo").Value)

    If modSeguridad.PuedeVerCostos(rolEfectivo) Then
        precioCostoOut = CCur(rs.Fields("precio_costo").Value)
    Else
        precioCostoOut = 0
    End If

    rs.Close
    ObtenerProducto = True
    Exit Function
Manejador:
    Err.Raise Err.Number, Err.Source, "Error al obtener producto: " & Err.Description
End Function

' Devuelve un Recordset desconectado con el catalogo de productos.
' Si el rol es Admin incluye precio_costo (9 cols); para Vendedor o Cliente omite precio_costo (8 cols).
Public Function ConsultarCatalogoSegunRol(ByVal conn As Object, _
        Optional ByVal rolUsuario As String = "", _
        Optional ByVal soloBajoMinimo As Boolean = False) As Object
    On Error GoTo Manejador
    Dim sql As String
    Dim rolEfectivo As String
    Dim filtro As String

    rolEfectivo = Trim$(rolUsuario)
    If Len(rolEfectivo) = 0 Then rolEfectivo = modSesion.gUsuarioRol

    If soloBajoMinimo Then
        filtro = " WHERE p.stock <= p.stock_minimo"
    Else
        filtro = ""
    End If

    If modSeguridad.PuedeVerCostos(rolEfectivo) Then
        sql = "SELECT p.sku, p.nombre, c.nombre AS categoria, p.marca, p.talla, " & _
              "p.precio_costo, p.precio_venta, p.stock, p.stock_minimo " & _
              "FROM PRODUCTOS p INNER JOIN CATEGORIAS c ON p.id_categoria = c.id" & _
              filtro & " ORDER BY p.sku"
    Else
        sql = "SELECT p.sku, p.nombre, c.nombre AS categoria, p.marca, p.talla, " & _
              "p.precio_venta, p.stock, p.stock_minimo " & _
              "FROM PRODUCTOS p INNER JOIN CATEGORIAS c ON p.id_categoria = c.id" & _
              filtro & " ORDER BY p.sku"
    End If

    Set ConsultarCatalogoSegunRol = modConexion.Consultar(conn, sql)
    Exit Function
Manejador:
    Err.Raise Err.Number, Err.Source, "Error al consultar catalogo: " & Err.Description
End Function

' Cuenta cuantos productos estan en alerta de reposicion (stock <= stock_minimo).
Public Function ContarProductosBajoMinimo(ByVal conn As Object) As Long
    On Error GoTo Manejador
    Dim rs As Object
    Set rs = conn.Execute("SELECT COUNT(*) FROM PRODUCTOS WHERE stock <= stock_minimo")
    ContarProductosBajoMinimo = CLng(rs.Fields(0).Value)
    rs.Close
    Exit Function
Manejador:
    Err.Raise Err.Number, Err.Source, "Error al contar productos bajo minimo: " & Err.Description
End Function

' Registra un movimiento de inventario ('E' entrada, 'S' salida, 'A' ajuste)
' y actualiza PRODUCTOS.stock de forma atomica.
Public Function RegistrarMovimientoInventario(ByVal conn As Object, _
        ByVal sku As String, ByVal idUsuario As Long, _
        ByVal tipoMovimiento As String, ByVal cantidad As Long, _
        ByVal motivo As String, Optional ByVal idProveedor As Long = 0, _
        Optional ByRef mensaje As String = "", _
        Optional ByVal usarTransaccionPropia As Boolean = True) As Long
    On Error GoTo Manejador
    Dim cmdUpd As Object, cmdIns As Object, rs As Object
    Dim skuLimpio As String, tipo As String, motivoLimpio As String
    Dim stockActual As Long, delta As Long, nuevoStock As Long
    Dim sqlProv As String, sqlIns As String
    Dim transIniciada As Boolean

    RegistrarMovimientoInventario = 0
    mensaje = ""
    skuLimpio = UCase$(Trim$(sku))
    tipo = UCase$(Trim$(tipoMovimiento))
    motivoLimpio = Trim$(motivo)

    If Not modValidaciones.ValidarCampoObligatorio(skuLimpio) Or Len(skuLimpio) > 50 Then
        mensaje = "El SKU del producto es obligatorio (maximo 50 caracteres)."
        Exit Function
    End If
    If idUsuario <= 0 Then
        mensaje = "El usuario que registra el movimiento no es valido."
        Exit Function
    End If
    If tipo <> "E" And tipo <> "S" And tipo <> "A" Then
        mensaje = "El tipo de movimiento no es valido. Use E (entrada), S (salida) o A (ajuste)."
        Exit Function
    End If
    If (tipo = "E" Or tipo = "S") And cantidad <= 0 Then
        mensaje = "La cantidad para entradas y salidas debe ser mayor a 0."
        Exit Function
    End If
    If tipo = "A" And cantidad = 0 Then
        mensaje = "La cantidad de un ajuste no puede ser 0."
        Exit Function
    End If
    If Not modValidaciones.ValidarCampoObligatorio(motivoLimpio) Then
        mensaje = "El motivo del movimiento es obligatorio."
        Exit Function
    End If
    If Len(motivoLimpio) > 255 Then
        mensaje = "El motivo admite hasta 255 caracteres."
        Exit Function
    End If

    stockActual = ObtenerStock(conn, skuLimpio)
    If stockActual < 0 Then
        mensaje = "El producto con SKU '" & skuLimpio & "' no existe."
        Exit Function
    End If

    Select Case tipo
        Case "E": delta = cantidad
        Case "S": delta = -cantidad
        Case "A": delta = cantidad
    End Select

    nuevoStock = stockActual + delta
    If nuevoStock < 0 Then
        mensaje = "Stock insuficiente para SKU " & skuLimpio & " (disponible: " & stockActual & ")."
        Exit Function
    End If

    If usarTransaccionPropia Then
        On Error Resume Next
        conn.BeginTrans
        transIniciada = (Err.Number = 0)
        On Error GoTo Manejador
    End If

    ' 1. Actualizar stock en PRODUCTOS
    Set cmdUpd = modConexion.NuevoComando(conn, "UPDATE PRODUCTOS SET stock = ? WHERE sku = ?")
    modConexion.ParamEntero cmdUpd, "pStock", nuevoStock
    modConexion.ParamTexto cmdUpd, "pSku", skuLimpio, 50
    cmdUpd.Execute

    ' 2. Insertar fila en MOVIMIENTOS_INVENTARIO
    If idProveedor > 0 Then sqlProv = "?" Else sqlProv = "NULL"
    sqlIns = "INSERT INTO MOVIMIENTOS_INVENTARIO " & _
             "(sku_producto, id_usuario, id_proveedor, tipo_movimiento, cantidad, fecha, motivo) " & _
             "VALUES (?, ?, " & sqlProv & ", ?, ?, Date(), ?)"

    Set cmdIns = modConexion.NuevoComando(conn, sqlIns)
    modConexion.ParamTexto cmdIns, "pSku", skuLimpio, 50
    modConexion.ParamEntero cmdIns, "pIdUsuario", idUsuario
    If idProveedor > 0 Then modConexion.ParamEntero cmdIns, "pIdProveedor", idProveedor
    modConexion.ParamTexto cmdIns, "pTipo", tipo, 1
    modConexion.ParamEntero cmdIns, "pCantidad", cantidad
    modConexion.ParamTexto cmdIns, "pMotivo", motivoLimpio, 255
    cmdIns.Execute

    Set rs = conn.Execute("SELECT @@IDENTITY")
    RegistrarMovimientoInventario = CLng(rs.Fields(0).Value)
    rs.Close

    If transIniciada Then
        conn.CommitTrans
        ' Sincronizar inmediatamente las hojas Productos e Inventario de Excel
        On Error Resume Next
        VolcarCatalogoAHoja conn, modSesion.gUsuarioRol
        On Error GoTo Manejador
    End If
    Exit Function
Manejador:
    Dim errNum As Long, errSrc As String, errDesc As String
    errNum = Err.Number: errSrc = Err.Source: errDesc = Err.Description
    If transIniciada Then
        On Error Resume Next
        conn.RollbackTrans
        On Error GoTo 0
    End If
    Err.Raise errNum, errSrc, "Error al registrar movimiento de inventario: " & errDesc
End Function

' Vuelca el catalogo a la hoja "Productos" (respetando visibilidad de costos por rol)
' y los ultimos 100 movimientos a la hoja "Inventario", estilizando los encabezados.
Public Sub VolcarCatalogoAHoja(ByVal conn As Object, Optional ByVal rolUsuario As String = "")
    On Error GoTo Manejador
    Dim wsProd As Worksheet, wsInv As Worksheet
    Dim rsProd As Object, rsInv As Object
    Dim i As Long

    Set wsProd = ThisWorkbook.Worksheets("Productos")
    wsProd.Cells.Clear
    Set rsProd = ConsultarCatalogoSegunRol(conn, rolUsuario, False)
    For i = 0 To rsProd.Fields.Count - 1
        wsProd.Cells(1, i + 1).Value = rsProd.Fields(i).Name
        wsProd.Cells(1, i + 1).Font.Bold = True
        wsProd.Cells(1, i + 1).Interior.Color = RGB(15, 23, 42)
        wsProd.Cells(1, i + 1).Font.Color = RGB(255, 255, 255)
    Next i
    If Not rsProd.EOF Then
        wsProd.Range("A2").CopyFromRecordset rsProd
    End If
    rsProd.Close
    wsProd.Columns.AutoFit

    Set wsInv = ThisWorkbook.Worksheets("Inventario")
    wsInv.Cells.Clear
    Set rsInv = modConexion.Consultar(conn, _
        "SELECT TOP 100 id, sku_producto, id_usuario, id_proveedor, tipo_movimiento, cantidad, fecha, motivo " & _
        "FROM MOVIMIENTOS_INVENTARIO ORDER BY id DESC")
    For i = 0 To rsInv.Fields.Count - 1
        wsInv.Cells(1, i + 1).Value = rsInv.Fields(i).Name
        wsInv.Cells(1, i + 1).Font.Bold = True
        wsInv.Cells(1, i + 1).Interior.Color = RGB(15, 23, 42)
        wsInv.Cells(1, i + 1).Font.Color = RGB(255, 255, 255)
    Next i
    If Not rsInv.EOF Then
        wsInv.Range("A2").CopyFromRecordset rsInv
    End If
    rsInv.Close
    wsInv.Columns.AutoFit
    Exit Sub
Manejador:
    Err.Raise Err.Number, Err.Source, "Error al volcar catalogo a hojas: " & Err.Description
End Sub

' Vuelca Clientes, Proveedores y Usuarios (segun permisos del rol) a sus hojas en Excel.
Public Sub VolcarTablasMaestrasAHoja(ByVal conn As Object, Optional ByVal rolUsuario As String = "")
    On Error Resume Next
    Dim wsCli As Worksheet, wsProv As Worksheet, wsUsr As Worksheet
    Dim rs As Object, i As Long, rolEf As String

    rolEf = Trim$(rolUsuario)
    If Len(rolEf) = 0 Then rolEf = modSesion.gUsuarioRol

    Set wsCli = ThisWorkbook.Worksheets("Clientes")
    wsCli.Cells.Clear
    Set rs = modConexion.Consultar(conn, _
        "SELECT TOP 150 id, rut, nombre, apellido, correo, telefono, direccion, fecha_registro FROM CLIENTES ORDER BY id DESC")
    For i = 0 To rs.Fields.Count - 1
        wsCli.Cells(1, i + 1).Value = rs.Fields(i).Name
        wsCli.Cells(1, i + 1).Font.Bold = True
        wsCli.Cells(1, i + 1).Interior.Color = RGB(15, 23, 42)
        wsCli.Cells(1, i + 1).Font.Color = RGB(255, 255, 255)
    Next i
    If Not rs.EOF Then wsCli.Range("A2").CopyFromRecordset rs
    rs.Close
    wsCli.Columns.AutoFit

    Set wsProv = ThisWorkbook.Worksheets("Proveedores")
    wsProv.Cells.Clear
    Set rs = modConexion.Consultar(conn, _
        "SELECT id, nombre, contacto, telefono, correo, direccion FROM PROVEEDORES ORDER BY id")
    For i = 0 To rs.Fields.Count - 1
        wsProv.Cells(1, i + 1).Value = rs.Fields(i).Name
        wsProv.Cells(1, i + 1).Font.Bold = True
        wsProv.Cells(1, i + 1).Interior.Color = RGB(15, 23, 42)
        wsProv.Cells(1, i + 1).Font.Color = RGB(255, 255, 255)
    Next i
    If Not rs.EOF Then wsProv.Range("A2").CopyFromRecordset rs
    rs.Close
    wsProv.Columns.AutoFit

    Set wsUsr = ThisWorkbook.Worksheets("Usuarios")
    wsUsr.Cells.Clear
    If modSeguridad.PuedeAdministrarUsuarios(rolEf) Then
        Set rs = modConexion.Consultar(conn, "SELECT id, nombre, correo, rol FROM USUARIOS ORDER BY id")
        For i = 0 To rs.Fields.Count - 1
            wsUsr.Cells(1, i + 1).Value = rs.Fields(i).Name
            wsUsr.Cells(1, i + 1).Font.Bold = True
            wsUsr.Cells(1, i + 1).Interior.Color = RGB(15, 23, 42)
            wsUsr.Cells(1, i + 1).Font.Color = RGB(255, 255, 255)
        Next i
        If Not rs.EOF Then wsUsr.Range("A2").CopyFromRecordset rs
        rs.Close
        wsUsr.Columns.AutoFit
    Else
        wsUsr.Cells(1, 1).Value = "Vista reservada para rol Admin"
    End If
    On Error GoTo 0
End Sub
