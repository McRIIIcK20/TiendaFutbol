VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVenta 
   Caption         =   "Tienda de Futbol - Punto de Venta (POS)"
   ClientHeight    =   7800
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10400
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.Label lblBanner 
      Caption         =   ""
      Height          =   1000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10400
   End
   Begin MSForms.Label lblAccent 
      Caption         =   ""
      Height          =   80
      Left            =   0
      TabIndex        =   1
      Top             =   1000
      Width           =   10400
   End
   Begin MSForms.Label lblTitulo 
      Caption         =   "PUNTO DE VENTA (POS) • DESCUENTO DE STOCK Y BOLETA PDF"
      Height          =   340
      Left            =   280
      TabIndex        =   2
      Top             =   160
      Width           =   9840
   End
   Begin MSForms.Label lblSubtitulo 
      Caption         =   "Transaccion ACID con rebaja automatica en PRODUCTOS, MOVIMIENTOS_INVENTARIO y hojas de Excel"
      Height          =   280
      Left            =   280
      TabIndex        =   3
      Top             =   540
      Width           =   9840
   End
   Begin MSForms.Label lblCardForm 
      Caption         =   ""
      Height          =   2160
      Left            =   240
      TabIndex        =   4
      Top             =   1240
      Width           =   7200
   End
   Begin MSForms.Label lblCardPreview 
      Caption         =   ""
      Height          =   2160
      Left            =   7600
      TabIndex        =   5
      Top             =   1240
      Width           =   2560
   End
   Begin MSForms.Image imgPreview 
      Height          =   1880
      Left            =   7740
      Top             =   1380
      Width           =   2280
   End
   Begin MSForms.Label lblCliente 
      Caption         =   "CLIENTE:"
      Height          =   280
      Left            =   440
      TabIndex        =   6
      Top             =   1460
      Width           =   1280
   End
   Begin MSForms.ComboBox cboCliente 
      Style           =   2  'fmStyleDropDownList
      Height          =   380
      Left            =   1800
      TabIndex        =   7
      Top             =   1400
      Width           =   5400
   End
   Begin MSForms.Label lblProducto 
      Caption         =   "PRODUCTO:"
      Height          =   280
      Left            =   440
      TabIndex        =   8
      Top             =   2020
      Width           =   1280
   End
   Begin MSForms.ComboBox cboProducto 
      Style           =   2  'fmStyleDropDownList
      Height          =   380
      Left            =   1800
      TabIndex        =   9
      Top             =   1960
      Width           =   5400
   End
   Begin MSForms.Label lblCantidad 
      Caption         =   "CANTIDAD:"
      Height          =   280
      Left            =   440
      TabIndex        =   10
      Top             =   2620
      Width           =   1280
   End
   Begin MSForms.TextBox txtCantidad 
      MaxLength       =   6
      Height          =   400
      Left            =   1800
      TabIndex        =   11
      Top             =   2560
      Width           =   1100
   End
   Begin MSForms.CommandButton cmdAgregarItem 
      Caption         =   "+ Agregar al Carrito"
      Height          =   440
      Left            =   3060
      TabIndex        =   12
      Top             =   2540
      Width           =   2040
   End
   Begin MSForms.CommandButton cmdQuitarItem 
      Caption         =   "Quitar Seleccionado"
      Height          =   440
      Left            =   5220
      TabIndex        =   13
      Top             =   2540
      Width           =   1980
   End
   Begin MSForms.Label lblStockVivo 
      Caption         =   "Stock disponible: --"
      Height          =   280
      Left            =   440
      TabIndex        =   14
      Top             =   3080
      Width           =   6760
   End
   Begin MSForms.Label lblCardCarrito 
      Caption         =   ""
      Height          =   2760
      Left            =   240
      TabIndex        =   15
      Top             =   3540
      Width           =   9920
   End
   Begin MSForms.Label lblCarrito 
      Caption         =   "DETALLE DEL CARRITO DE COMPRA:"
      Height          =   280
      Left            =   440
      TabIndex        =   16
      Top             =   3700
      Width           =   7200
   End
   Begin MSForms.CommandButton cmdLimpiarCarrito 
      Caption         =   "Vaciar Carrito"
      Height          =   360
      Left            =   8200
      TabIndex        =   17
      Top             =   3660
      Width           =   1760
   End
   Begin MSForms.ListBox lstCarrito 
      Height          =   1680
      Left            =   440
      TabIndex        =   18
      Top             =   4080
      Width           =   9520
   End
   Begin MSForms.Label lblTotales 
      Caption         =   "Subtotal Neto: $0   |   IVA (19%): $0   |   TOTAL A PAGAR: $0"
      Height          =   360
      Left            =   440
      TabIndex        =   19
      Top             =   5860
      Width           =   9520
   End
   Begin MSForms.Label lblMensaje 
      Caption         =   ""
      ForeColor       =   &H000000C0&
      WordWrap        =   -1  'True
      Height          =   540
      Left            =   240
      TabIndex        =   20
      Top             =   6400
      Width           =   9920
   End
   Begin MSForms.CommandButton cmdConfirmarVenta 
      Caption         =   "CONFIRMAR COMPRA, DESCONTAR STOCK Y EMITIR PDF"
      Default         =   -1  'True
      Height          =   560
      Left            =   240
      TabIndex        =   21
      Top             =   7040
      Width           =   4800
   End
   Begin MSForms.CommandButton cmdAnularVenta 
      Caption         =   "Anular Venta (Restituir Stock)"
      Height          =   560
      Left            =   5200
      TabIndex        =   22
      Top             =   7040
      Width           =   2600
   End
   Begin MSForms.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Cancel          =   -1  'True
      Height          =   560
      Left            =   7960
      TabIndex        =   23
      Top             =   7040
      Width           =   2200
   End
End
Attribute VB_Name = "frmVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSkus() As String
Private mNombres() As String
Private mCantidades() As Long
Private mPrecios() As Currency
Private mNumItems As Long

Private Sub UserForm_Initialize()
    Me.BackColor = RGB(241, 245, 249)
    lblBanner.BackColor = RGB(15, 23, 42)
    lblAccent.BackColor = RGB(16, 185, 129)

    lblTitulo.BackStyle = 0: lblTitulo.ForeColor = RGB(255, 255, 255): lblTitulo.Font.Bold = True: lblTitulo.Font.Size = 13
    lblSubtitulo.BackStyle = 0: lblSubtitulo.ForeColor = RGB(148, 163, 184): lblSubtitulo.Font.Size = 9

    lblCardForm.BackColor = RGB(255, 255, 255): lblCardForm.BorderStyle = 1: lblCardForm.BorderColor = RGB(203, 213, 225)
    lblCardPreview.BackColor = RGB(255, 255, 255): lblCardPreview.BorderStyle = 1: lblCardPreview.BorderColor = RGB(203, 213, 225)
    lblCardCarrito.BackColor = RGB(255, 255, 255): lblCardCarrito.BorderStyle = 1: lblCardCarrito.BorderColor = RGB(203, 213, 225)

    lblCliente.BackStyle = 0: lblCliente.Font.Bold = True: lblCliente.ForeColor = RGB(51, 65, 85)
    lblProducto.BackStyle = 0: lblProducto.Font.Bold = True: lblProducto.ForeColor = RGB(51, 65, 85)
    lblCantidad.BackStyle = 0: lblCantidad.Font.Bold = True: lblCantidad.ForeColor = RGB(51, 65, 85)
    lblStockVivo.BackStyle = 0: lblStockVivo.Font.Bold = True: lblStockVivo.ForeColor = RGB(5, 150, 105)
    lblCarrito.BackStyle = 0: lblCarrito.Font.Bold = True: lblCarrito.ForeColor = RGB(15, 23, 42)
    lblTotales.BackStyle = 0: lblTotales.Font.Bold = True: lblTotales.Font.Size = 10.5: lblTotales.ForeColor = RGB(15, 23, 42)
    lblMensaje.BackStyle = 0: lblMensaje.Font.Bold = True: lblMensaje.Caption = ""

    On Error Resume Next
    imgPreview.PictureSizeMode = 3  ' fmPictureSizeModeZoom
    imgPreview.BorderStyle = 0
    imgPreview.BackColor = RGB(255, 255, 255)
    On Error GoTo 0

    cmdAgregarItem.BackColor = RGB(30, 41, 59): cmdAgregarItem.ForeColor = RGB(255, 255, 255): cmdAgregarItem.Font.Bold = True
    cmdQuitarItem.BackColor = RGB(100, 116, 139): cmdQuitarItem.ForeColor = RGB(255, 255, 255)
    cmdLimpiarCarrito.BackColor = RGB(100, 116, 139): cmdLimpiarCarrito.ForeColor = RGB(255, 255, 255)
    cmdConfirmarVenta.BackColor = RGB(5, 150, 105): cmdConfirmarVenta.ForeColor = RGB(255, 255, 255): cmdConfirmarVenta.Font.Bold = True
    cmdAnularVenta.BackColor = RGB(180, 83, 9): cmdAnularVenta.ForeColor = RGB(255, 255, 255): cmdAnularVenta.Font.Bold = True
    cmdCerrar.BackColor = RGB(71, 85, 105): cmdCerrar.ForeColor = RGB(255, 255, 255): cmdCerrar.Font.Bold = True

    txtCantidad.Value = "1"
    mNumItems = 0

    CargarClientes
    CargarProductosDisponibles
    ActualizarVistaCarrito
    modEstilo.AplicarTema Me
End Sub

Private Sub CargarClientes()
    On Error GoTo Fallo
    Dim conn As Object, rs As Object
    cboCliente.Clear
    Set conn = modConexion.AbrirConexion()
    Set rs = conn.Execute("SELECT id, rut, nombre, apellido FROM CLIENTES ORDER BY nombre, apellido")
    Do While Not rs.EOF
        cboCliente.AddItem CStr(rs.Fields("id").Value) & " | " & _
            CStr(rs.Fields("rut").Value) & " - " & _
            CStr(rs.Fields("nombre").Value) & " " & CStr(rs.Fields("apellido").Value)
        rs.MoveNext
    Loop
    rs.Close
    modConexion.CerrarConexion conn
    If cboCliente.ListCount > 0 Then cboCliente.ListIndex = 0
    Exit Sub
Fallo:
    modConexion.CerrarConexion conn
    MostrarError "Error al cargar clientes: " & Err.Description
End Sub

Private Sub CargarProductosDisponibles()
    On Error GoTo Fallo
    Dim conn As Object, rs As Object, sku As String, stReal As Long, stDisp As Long
    Dim prevSku As String
    prevSku = ObtenerSkuProductoSeleccionado()

    cboProducto.Clear
    Set conn = modConexion.AbrirConexion()
    Set rs = conn.Execute("SELECT sku, nombre, precio_venta, stock FROM PRODUCTOS WHERE stock > 0 ORDER BY sku")
    Do While Not rs.EOF
        sku = CStr(rs.Fields("sku").Value)
        stReal = CLng(rs.Fields("stock").Value)
        stDisp = stReal - CantidadEnCarrito(sku)
        If stDisp < 0 Then stDisp = 0
        cboProducto.AddItem sku & " | " & _
            CStr(rs.Fields("nombre").Value) & " | $" & _
            Format$(CCur(rs.Fields("precio_venta").Value), "#,##0") & _
            " (Stock disp: " & CStr(stDisp) & " de " & CStr(stReal) & ")"
        rs.MoveNext
    Loop
    rs.Close
    modConexion.CerrarConexion conn

    If cboProducto.ListCount > 0 Then
        Dim i As Long, encontrado As Boolean
        For i = 0 To cboProducto.ListCount - 1
            If Left$(cboProducto.List(i), Len(prevSku) + 3) = prevSku & " | " Then
                cboProducto.ListIndex = i
                encontrado = True
                Exit For
            End If
        Next i
        If Not encontrado Then cboProducto.ListIndex = 0
    End If
    Exit Sub
Fallo:
    modConexion.CerrarConexion conn
    MostrarError "Error al cargar productos: " & Err.Description
End Sub

Private Function CantidadEnCarrito(ByVal sku As String) As Long
    Dim i As Long
    CantidadEnCarrito = 0
    For i = 0 To mNumItems - 1
        If StrComp(mSkus(i), sku, vbTextCompare) = 0 Then
            CantidadEnCarrito = mCantidades(i)
            Exit Function
        End If
    Next i
End Function

Private Sub cboProducto_Click()
    ActualizarPreviewProducto
End Sub

Private Sub ActualizarPreviewProducto()
    On Error Resume Next
    Dim sku As String, conn As Object
    Dim nom As String, pVenta As Currency, stReal As Long, stMin As Long, pCosto As Currency
    Dim cat As String, marca As String, talla As String, stDisp As Long, rutaImg As String

    sku = ObtenerSkuProductoSeleccionado()
    If Len(sku) = 0 Then Exit Sub

    Set conn = modConexion.AbrirConexion()
    If modInventario.ObtenerProducto(conn, sku, nom, pVenta, stReal, stMin, pCosto, modSesion.gUsuarioRol, cat, marca, talla) Then
        stDisp = stReal - CantidadEnCarrito(sku)
        If stDisp < 0 Then stDisp = 0
        lblStockVivo.Caption = "SKU " & sku & " (" & marca & " • Talla " & talla & ")  —  Stock en bodega: " & stReal & _
                               "  |  En carrito: " & CantidadEnCarrito(sku) & "  |  Disponible: " & stDisp
        rutaImg = modInventario.ObtenerRutaImagenProducto(sku, cat)
        If Len(rutaImg) > 0 Then
            Set imgPreview.Picture = LoadPicture(rutaImg)
        Else
            Set imgPreview.Picture = Nothing
        End If
    End If
    modConexion.CerrarConexion conn
    On Error GoTo 0
End Sub

Private Function ObtenerIdClienteSeleccionado() As Long
    Dim txt As String, pos As Long
    ObtenerIdClienteSeleccionado = 0
    If cboCliente.ListIndex < 0 Then Exit Function
    txt = cboCliente.Text
    pos = InStr(txt, " | ")
    If pos > 1 Then ObtenerIdClienteSeleccionado = CLng(Val(Left$(txt, pos - 1)))
End Function

Private Function ObtenerSkuProductoSeleccionado() As String
    Dim txt As String, pos As Long
    ObtenerSkuProductoSeleccionado = ""
    If cboProducto.ListIndex < 0 Then Exit Function
    txt = cboProducto.Text
    pos = InStr(txt, " | ")
    If pos > 1 Then ObtenerSkuProductoSeleccionado = Trim$(Left$(txt, pos - 1))
End Function

Private Sub cmdAgregarItem_Click()
    On Error GoTo Fallo
    Dim sku As String, nombre As String
    Dim cant As Long, stockDisp As Long, i As Long, cantAcumulada As Long
    Dim pVenta As Currency, stMin As Long, pCosto As Currency
    Dim conn As Object

    lblMensaje.Caption = ""
    sku = ObtenerSkuProductoSeleccionado()
    If Len(sku) = 0 Then
        MostrarError "Seleccione un producto con stock disponible."
        Exit Sub
    End If
    If Not IsNumeric(Trim$(txtCantidad.Value)) Then
        MostrarError "Ingrese una cantidad numerica valida."
        Exit Sub
    End If
    cant = CLng(txtCantidad.Value)
    If cant <= 0 Then
        MostrarError "La cantidad debe ser mayor a 0."
        Exit Sub
    End If

    Set conn = modConexion.AbrirConexion()
    If Not modInventario.ObtenerProducto(conn, sku, nombre, pVenta, stockDisp, stMin, pCosto, modSesion.gUsuarioRol) Then
        modConexion.CerrarConexion conn
        MostrarError "El producto seleccionado no existe."
        Exit Sub
    End If
    modConexion.CerrarConexion conn

    For i = 0 To mNumItems - 1
        If mSkus(i) = sku Then
            cantAcumulada = mCantidades(i) + cant
            If cantAcumulada > stockDisp Then
                MostrarError "Stock insuficiente para " & sku & " (en carrito: " & mCantidades(i) & ", disponible: " & stockDisp & ")."
                Exit Sub
            End If
            mCantidades(i) = cantAcumulada
            ActualizarVistaCarrito
            CargarProductosDisponibles
            MostrarOk "Reservadas " & cantAcumulada & " u. de " & sku & " en el carrito (Quedan " & (stockDisp - cantAcumulada) & " disp.)."
            Exit Sub
        End If
    Next i

    If cant > stockDisp Then
        MostrarError "Stock insuficiente para " & sku & " (disponible: " & stockDisp & ")."
        Exit Sub
    End If

    ReDim Preserve mSkus(0 To mNumItems)
    ReDim Preserve mNombres(0 To mNumItems)
    ReDim Preserve mCantidades(0 To mNumItems)
    ReDim Preserve mPrecios(0 To mNumItems)

    mSkus(mNumItems) = sku
    mNombres(mNumItems) = nombre
    mCantidades(mNumItems) = cant
    mPrecios(mNumItems) = pVenta
    mNumItems = mNumItems + 1

    txtCantidad.Value = "1"
    ActualizarVistaCarrito
    CargarProductosDisponibles
    MostrarOk "Agregado " & sku & " x" & cant & " al carrito (Stock disponible ahora: " & (stockDisp - cant) & ")."
    Exit Sub
Fallo:
    modConexion.CerrarConexion conn
    MostrarError "Error al agregar item: " & Err.Description
End Sub

Private Sub cmdQuitarItem_Click()
    Dim idx As Long, i As Long
    idx = lstCarrito.ListIndex
    If idx < 0 Or idx >= mNumItems Then
        MostrarError "Seleccione una linea del carrito para quitarla."
        Exit Sub
    End If
    For i = idx To mNumItems - 2
        mSkus(i) = mSkus(i + 1)
        mNombres(i) = mNombres(i + 1)
        mCantidades(i) = mCantidades(i + 1)
        mPrecios(i) = mPrecios(i + 1)
    Next i
    mNumItems = mNumItems - 1
    If mNumItems <= 0 Then
        cmdLimpiarCarrito_Click
    Else
        ReDim Preserve mSkus(0 To mNumItems - 1)
        ReDim Preserve mNombres(0 To mNumItems - 1)
        ReDim Preserve mCantidades(0 To mNumItems - 1)
        ReDim Preserve mPrecios(0 To mNumItems - 1)
        ActualizarVistaCarrito
        CargarProductosDisponibles
    End If
End Sub

Private Sub cmdLimpiarCarrito_Click()
    mNumItems = 0
    Erase mSkus
    Erase mNombres
    Erase mCantidades
    Erase mPrecios
    lblMensaje.Caption = ""
    ActualizarVistaCarrito
    CargarProductosDisponibles
End Sub

Private Sub ActualizarVistaCarrito()
    Dim i As Long
    Dim subLinea As Currency, subtotal As Currency, iva As Currency, total As Currency

    lstCarrito.Clear
    subtotal = 0
    For i = 0 To mNumItems - 1
        subLinea = CCur(Round(CDbl(mPrecios(i)) * CDbl(mCantidades(i)), 2))
        subtotal = subtotal + subLinea
        lstCarrito.AddItem mSkus(i) & " | " & mNombres(i) & " | " & _
            CStr(mCantidades(i)) & " x $" & Format$(mPrecios(i), "#,##0") & _
            " = $" & Format$(subLinea, "#,##0")
    Next i

    iva = CCur(Round(CDbl(subtotal) * modVentas.TASA_IVA, 2))
    total = subtotal + iva
    lblTotales.Caption = "Subtotal Neto: $" & Format$(subtotal, "#,##0") & _
        "   |   IVA (19%): $" & Format$(iva, "#,##0") & _
        "   |   TOTAL A PAGAR: $" & Format$(total, "#,##0")
End Sub

Private Sub cmdConfirmarVenta_Click()
    Dim conn As Object
    Dim idCliente As Long, idVenta As Long, idBoleta As Long
    Dim corr As String, msg As String, errDesc As String
    Dim rutaPdf As String, msgPdf As String, destCorreo As String, msgCorreo As String
    Dim resumen As String, sku0 As String, stAntes As Long, stDespues As Long

    lblMensaje.Caption = ""
    If Not modSesion.PuedeVender() Then
        MostrarError "Su rol no permite registrar ventas."
        Exit Sub
    End If
    idCliente = ObtenerIdClienteSeleccionado()
    If idCliente <= 0 Then
        MostrarError "Seleccione un cliente."
        Exit Sub
    End If
    If mNumItems <= 0 Then
        MostrarError "Agregue al menos un producto al carrito antes de confirmar."
        Exit Sub
    End If

    On Error GoTo Fallo
    Set conn = modConexion.AbrirConexion()
    sku0 = mSkus(0)
    stAntes = modInventario.ObtenerStock(conn, sku0)

    idVenta = modVentas.RegistrarVenta(conn, modSesion.gUsuarioId, idCliente, _
        mSkus, mCantidades, msg, idBoleta, corr, True, False)

    If idVenta > 0 Then
        stDespues = modInventario.ObtenerStock(conn, sku0)
        resumen = "Venta #" & idVenta & " (" & corr & ") registrada. Stock de " & sku0 & _
                  " rebajado (" & stAntes & " -> " & stDespues & ")."
        If modBoletas.GenerarBoletaPDF(conn, idVenta, rutaPdf, msgPdf) Then
            resumen = resumen & " PDF emitido."
            If Len(modBoletas.LeerConfigTexto(modBoletas.NOMBRE_RANGO_GMAIL_CLAVE)) > 0 Then
                If modBoletas.EnviarBoletaPorCorreo(conn, idVenta, destCorreo, msgCorreo, False) Then
                    resumen = resumen & " Enviada a " & destCorreo & "."
                Else
                    resumen = resumen & " Correo NO enviado: " & msgCorreo
                End If
            End If
        End If
        ' Sincronizar tanto Ventas/Boletas como Productos/Inventario en el libro Excel
        modVentas.VolcarVentasAHoja conn
        modInventario.VolcarCatalogoAHoja conn, modSesion.gUsuarioRol
    End If
    modConexion.CerrarConexion conn

    If idVenta > 0 Then
        cmdLimpiarCarrito_Click
        CargarProductosDisponibles
        MostrarOk resumen
    Else
        MostrarError msg
    End If
    Exit Sub
Fallo:
    errDesc = Err.Description
    modConexion.CerrarConexion conn
    MostrarError "No se pudo completar la venta: " & errDesc
End Sub

Private Sub cmdAnularVenta_Click()
    On Error GoTo Fallo
    Dim sId As String, idVenta As Long, conn As Object, msg As String
    sId = InputBox("Ingrese el ID numerico de la venta que desea anular (el stock se restituira automaticamente):", "Anular Venta y Restituir Stock")
    If Len(Trim$(sId)) = 0 Then Exit Sub
    If Not IsNumeric(Trim$(sId)) Then
        MostrarError "El ID de venta debe ser numerico."
        Exit Sub
    End If
    idVenta = CLng(Trim$(sId))

    Set conn = modConexion.AbrirConexion()
    If modVentas.AnularVenta(conn, idVenta, modSesion.gUsuarioId, msg, True) Then
        modVentas.VolcarVentasAHoja conn
        modInventario.VolcarCatalogoAHoja conn, modSesion.gUsuarioRol
        modConexion.CerrarConexion conn
        CargarProductosDisponibles
        MostrarOk "Venta #" & idVenta & " anulada correctamente y stock restituido en inventario."
    Else
        modConexion.CerrarConexion conn
        MostrarError msg
    End If
    Exit Sub
Fallo:
    modConexion.CerrarConexion conn
    MostrarError "Error al anular venta: " & Err.Description
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub MostrarError(ByVal texto As String)
    lblMensaje.ForeColor = RGB(185, 28, 28)
    lblMensaje.Caption = texto
End Sub

Private Sub MostrarOk(ByVal texto As String)
    lblMensaje.ForeColor = RGB(5, 150, 105)
    lblMensaje.Caption = texto
End Sub
