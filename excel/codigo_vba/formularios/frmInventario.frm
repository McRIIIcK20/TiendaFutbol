VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInventario 
   Caption         =   "Tienda de Futbol - Control de Inventario y Bodega"
   ClientHeight    =   7600
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10200
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.Label lblBanner 
      Caption         =   ""
      Height          =   1000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10200
   End
   Begin MSForms.Label lblAccent 
      Caption         =   ""
      Height          =   80
      Left            =   0
      TabIndex        =   1
      Top             =   1000
      Width           =   10200
   End
   Begin MSForms.Label lblTitulo 
      Caption         =   "CONTROL DE INVENTARIO, KARDEX Y MOVIMIENTOS DE STOCK"
      Height          =   340
      Left            =   280
      TabIndex        =   2
      Top             =   160
      Width           =   9640
   End
   Begin MSForms.Label lblResumen 
      Caption         =   ""
      Height          =   280
      Left            =   280
      TabIndex        =   3
      Top             =   540
      Width           =   9640
   End
   Begin MSForms.Label lblCardConsulta 
      Caption         =   ""
      Height          =   2180
      Left            =   240
      TabIndex        =   4
      Top             =   1240
      Width           =   7000
   End
   Begin MSForms.Label lblCardPreview 
      Caption         =   ""
      Height          =   2180
      Left            =   7400
      TabIndex        =   5
      Top             =   1240
      Width           =   2560
   End
   Begin MSForms.Image imgPreview 
      Height          =   1900
      Left            =   7540
      Top             =   1380
      Width           =   2280
   End
   Begin MSForms.Label lblProducto 
      Caption         =   "PRODUCTO (SKU):"
      Height          =   280
      Left            =   440
      TabIndex        =   6
      Top             =   1480
      Width           =   1560
   End
   Begin MSForms.ComboBox cboProducto 
      Style           =   2  'fmStyleDropDownList
      Height          =   380
      Left            =   2040
      TabIndex        =   7
      Top             =   1420
      Width           =   3360
   End
   Begin MSForms.CommandButton cmdFiltroMinimo 
      Caption         =   "Solo bajo min."
      Height          =   380
      Left            =   5520
      TabIndex        =   8
      Top             =   1420
      Width           =   1520
   End
   Begin MSForms.Label lblDetalleProducto 
      Caption         =   "Seleccione un producto para inspeccionar su stock, costo y fotografia."
      WordWrap        =   -1  'True
      Height          =   1200
      Left            =   440
      TabIndex        =   9
      Top             =   1980
      Width           =   6600
   End
   Begin MSForms.Label lblCardMov 
      Caption         =   ""
      Height          =   2540
      Left            =   240
      TabIndex        =   10
      Top             =   3580
      Width           =   9720
   End
   Begin MSForms.Label lblSeccionMov 
      Caption         =   "REGISTRAR MOVIMIENTO DE BODEGA (ENTRADA / SALIDA / AJUSTE)"
      Height          =   300
      Left            =   440
      TabIndex        =   11
      Top             =   3760
      Width           =   9320
   End
   Begin MSForms.Label lblTipo 
      Caption         =   "TIPO MOV.:"
      Height          =   280
      Left            =   440
      TabIndex        =   12
      Top             =   4220
      Width           =   1400
   End
   Begin MSForms.ComboBox cboTipo 
      Style           =   2  'fmStyleDropDownList
      Height          =   380
      Left            =   2040
      TabIndex        =   13
      Top             =   4160
      Width           =   3400
   End
   Begin MSForms.Label lblCantidad 
      Caption         =   "CANTIDAD:"
      Height          =   280
      Left            =   5800
      TabIndex        =   14
      Top             =   4220
      Width           =   1100
   End
   Begin MSForms.TextBox txtCantidad 
      MaxLength       =   10
      Height          =   380
      Left            =   7000
      TabIndex        =   15
      Top             =   4160
      Width           =   2720
   End
   Begin MSForms.Label lblProveedor 
      Caption         =   "PROVEEDOR:"
      Height          =   280
      Left            =   440
      TabIndex        =   16
      Top             =   4800
      Width           =   1400
   End
   Begin MSForms.ComboBox cboProveedor 
      Style           =   2  'fmStyleDropDownList
      Height          =   380
      Left            =   2040
      TabIndex        =   17
      Top             =   4740
      Width           =   7680
   End
   Begin MSForms.Label lblMotivo 
      Caption         =   "MOTIVO:"
      Height          =   280
      Left            =   440
      TabIndex        =   18
      Top             =   5380
      Width           =   1400
   End
   Begin MSForms.TextBox txtMotivo 
      MaxLength       =   255
      Height          =   380
      Left            =   2040
      TabIndex        =   19
      Top             =   5320
      Width           =   7680
   End
   Begin MSForms.Label lblMensaje 
      Caption         =   ""
      ForeColor       =   &H000000C0&
      WordWrap        =   -1  'True
      Height          =   480
      Left            =   240
      TabIndex        =   20
      Top             =   6240
      Width           =   9720
   End
   Begin MSForms.CommandButton cmdRegistrarMov 
      Caption         =   "GUARDAR MOVIMIENTO Y ACTUALIZAR STOCK"
      Default         =   -1  'True
      Height          =   540
      Left            =   240
      TabIndex        =   21
      Top             =   6840
      Width           =   4000
   End
   Begin MSForms.CommandButton cmdVolcarHoja 
      Caption         =   "Sincronizar Hojas Excel"
      Height          =   540
      Left            =   4440
      TabIndex        =   22
      Top             =   6840
      Width           =   2800
   End
   Begin MSForms.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Cancel          =   -1  'True
      Height          =   540
      Left            =   7440
      TabIndex        =   23
      Top             =   6840
      Width           =   2520
   End
End
Attribute VB_Name = "frmInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSoloBajoMinimo As Boolean

Private Sub UserForm_Initialize()
    Dim puedeEditar As Boolean

    Me.BackColor = RGB(241, 245, 249)
    lblBanner.BackColor = RGB(15, 23, 42)
    lblAccent.BackColor = RGB(16, 185, 129)

    lblTitulo.BackStyle = 0: lblTitulo.ForeColor = RGB(255, 255, 255): lblTitulo.Font.Bold = True: lblTitulo.Font.Size = 13
    lblResumen.BackStyle = 0: lblResumen.ForeColor = RGB(52, 211, 153): lblResumen.Font.Bold = True: lblResumen.Font.Size = 9

    lblCardConsulta.BackColor = RGB(255, 255, 255): lblCardConsulta.BorderStyle = 1: lblCardConsulta.BorderColor = RGB(203, 213, 225)
    lblCardPreview.BackColor = RGB(255, 255, 255): lblCardPreview.BorderStyle = 1: lblCardPreview.BorderColor = RGB(203, 213, 225)
    lblCardMov.BackColor = RGB(255, 255, 255): lblCardMov.BorderStyle = 1: lblCardMov.BorderColor = RGB(203, 213, 225)

    lblProducto.BackStyle = 0: lblProducto.Font.Bold = True: lblProducto.ForeColor = RGB(51, 65, 85)
    lblDetalleProducto.BackStyle = 0: lblDetalleProducto.ForeColor = RGB(15, 23, 42): lblDetalleProducto.Font.Size = 9.5
    lblSeccionMov.BackStyle = 0: lblSeccionMov.Font.Bold = True: lblSeccionMov.ForeColor = RGB(15, 23, 42)
    lblTipo.BackStyle = 0: lblTipo.Font.Bold = True: lblTipo.ForeColor = RGB(51, 65, 85)
    lblCantidad.BackStyle = 0: lblCantidad.Font.Bold = True: lblCantidad.ForeColor = RGB(51, 65, 85)
    lblProveedor.BackStyle = 0: lblProveedor.Font.Bold = True: lblProveedor.ForeColor = RGB(51, 65, 85)
    lblMotivo.BackStyle = 0: lblMotivo.Font.Bold = True: lblMotivo.ForeColor = RGB(51, 65, 85)
    lblMensaje.BackStyle = 0: lblMensaje.Font.Bold = True: lblMensaje.Caption = ""

    On Error Resume Next
    imgPreview.PictureSizeMode = 3  ' fmPictureSizeModeZoom
    imgPreview.BorderStyle = 0
    imgPreview.BackColor = RGB(255, 255, 255)
    On Error GoTo 0

    cmdFiltroMinimo.BackColor = RGB(30, 41, 59): cmdFiltroMinimo.ForeColor = RGB(255, 255, 255): cmdFiltroMinimo.Font.Bold = True
    cmdRegistrarMov.BackColor = RGB(5, 150, 105): cmdRegistrarMov.ForeColor = RGB(255, 255, 255): cmdRegistrarMov.Font.Bold = True
    cmdVolcarHoja.BackColor = RGB(30, 41, 59): cmdVolcarHoja.ForeColor = RGB(255, 255, 255): cmdVolcarHoja.Font.Bold = True
    cmdCerrar.BackColor = RGB(71, 85, 105): cmdCerrar.ForeColor = RGB(255, 255, 255): cmdCerrar.Font.Bold = True

    cboTipo.AddItem "E - Entrada (reposicion)"
    cboTipo.AddItem "S - Salida manual"
    cboTipo.AddItem "A - Ajuste de conteo"
    cboTipo.ListIndex = 0

    puedeEditar = modSesion.PuedeVender()
    cboTipo.Enabled = puedeEditar
    txtCantidad.Enabled = puedeEditar
    cboProveedor.Enabled = puedeEditar
    txtMotivo.Enabled = puedeEditar
    cmdRegistrarMov.Enabled = puedeEditar

    CargarProveedores
    CargarListaProductos
    modEstilo.AplicarTema Me
End Sub

Private Sub CargarProveedores()
    On Error GoTo Fallo
    Dim conn As Object, rs As Object
    cboProveedor.Clear
    cboProveedor.AddItem "0 | (Sin proveedor)"
    Set conn = modConexion.AbrirConexion()
    Set rs = conn.Execute("SELECT id, nombre FROM PROVEEDORES ORDER BY id")
    Do While Not rs.EOF
        cboProveedor.AddItem CStr(rs.Fields("id").Value) & " | " & CStr(rs.Fields("nombre").Value)
        rs.MoveNext
    Loop
    rs.Close
    modConexion.CerrarConexion conn
    cboProveedor.ListIndex = 0
    Exit Sub
Fallo:
    modConexion.CerrarConexion conn
End Sub

Private Sub CargarListaProductos()
    On Error GoTo Fallo
    Dim conn As Object, rs As Object
    Dim nBajo As Long, infoRol As String

    cboProducto.Clear
    Set conn = modConexion.AbrirConexion()
    nBajo = modInventario.ContarProductosBajoMinimo(conn)
    Set rs = modInventario.ConsultarCatalogoSegunRol(conn, modSesion.gUsuarioRol, mSoloBajoMinimo)

    Do While Not rs.EOF
        cboProducto.AddItem CStr(rs.Fields("sku").Value) & " | " & _
            CStr(rs.Fields("nombre").Value) & " (Stock: " & CStr(rs.Fields("stock").Value) & ")"
        rs.MoveNext
    Loop
    rs.Close
    modConexion.CerrarConexion conn

    If modSesion.PuedeVerCostos() Then
        infoRol = "Costos visibles (Admin)"
    Else
        infoRol = "Costos protegidos para rol " & modSesion.gUsuarioRol
    End If
    lblResumen.Caption = "Productos en alerta (stock <= minimo): " & nBajo & "   •   " & infoRol

    If cboProducto.ListCount > 0 Then
        cboProducto.ListIndex = 0
    Else
        lblDetalleProducto.Caption = "No hay productos para el filtro actual."
    End If
    Exit Sub
Fallo:
    modConexion.CerrarConexion conn
    MostrarError "Error al cargar productos: " & Err.Description
End Sub

Private Sub cboProducto_Click()
    MostrarDetalleSeleccionado
End Sub

Private Sub MostrarDetalleSeleccionado()
    On Error GoTo Fallo
    Dim sku As String, conn As Object
    Dim nombre As String, pVenta As Currency, pCosto As Currency
    Dim stock As Long, stockMin As Long, alerta As String, lineaCosto As String
    Dim cat As String, marca As String, talla As String, rutaImg As String

    sku = ObtenerSkuSeleccionado()
    If Len(sku) = 0 Then Exit Sub

    Set conn = modConexion.AbrirConexion()
    If modInventario.ObtenerProducto(conn, sku, nombre, pVenta, stock, stockMin, pCosto, modSesion.gUsuarioRol, cat, marca, talla) Then
        If stock <= 0 Then
            alerta = " [CRITICO: AGOTADO]"
        ElseIf stock <= stockMin Then
            alerta = " [ALERTA: BAJO MINIMO]"
        Else
            alerta = " [STOCK NORMAL]"
        End If
        If modSesion.PuedeVerCostos() Then
            lineaCosto = "   |   Costo Unitario: $" & Format$(pCosto, "#,##0")
        Else
            lineaCosto = ""
        End If
        lblDetalleProducto.Caption = "SKU: " & sku & "  —  " & nombre & vbCrLf & _
            "Categoria: " & cat & "   |   Marca: " & marca & "   |   Talla: " & talla & vbCrLf & _
            "Precio Venta Neto: $" & Format$(pVenta, "#,##0") & lineaCosto & vbCrLf & _
            "Stock Actual en Bodega: " & stock & " u. (Minimo requerido: " & stockMin & " u.)" & alerta

        rutaImg = modInventario.ObtenerRutaImagenProducto(sku, cat)
        On Error Resume Next
        If Len(rutaImg) > 0 Then
            Set imgPreview.Picture = LoadPicture(rutaImg)
        Else
            Set imgPreview.Picture = Nothing
        End If
        On Error GoTo Fallo
    End If
    modConexion.CerrarConexion conn
    Exit Sub
Fallo:
    modConexion.CerrarConexion conn
End Sub

Private Function ObtenerSkuSeleccionado() As String
    Dim txt As String, pos As Long
    ObtenerSkuSeleccionado = ""
    If cboProducto.ListIndex < 0 Then Exit Function
    txt = cboProducto.Text
    pos = InStr(txt, " | ")
    If pos > 1 Then ObtenerSkuSeleccionado = Trim$(Left$(txt, pos - 1))
End Function

Private Function ObtenerIdProveedorSeleccionado() As Long
    Dim txt As String, pos As Long
    ObtenerIdProveedorSeleccionado = 0
    If cboProveedor.ListIndex <= 0 Then Exit Function
    txt = cboProveedor.Text
    pos = InStr(txt, " | ")
    If pos > 1 Then ObtenerIdProveedorSeleccionado = CLng(Val(Left$(txt, pos - 1)))
End Function

Private Sub cmdFiltroMinimo_Click()
    mSoloBajoMinimo = Not mSoloBajoMinimo
    If mSoloBajoMinimo Then
        cmdFiltroMinimo.Caption = "Ver todos"
    Else
        cmdFiltroMinimo.Caption = "Solo bajo min."
    End If
    CargarListaProductos
End Sub

Private Sub cmdRegistrarMov_Click()
    Dim conn As Object
    Dim sku As String, tipo As String, motivo As String, msg As String
    Dim cant As Long, idProv As Long, idMov As Long, stockNuevo As Long
    Dim errDesc As String

    lblMensaje.Caption = ""
    If Not modSesion.PuedeVender() Then
        MostrarError "Su rol no permite modificar el inventario."
        Exit Sub
    End If

    sku = ObtenerSkuSeleccionado()
    If Len(sku) = 0 Then
        MostrarError "Seleccione un producto."
        Exit Sub
    End If
    If Not IsNumeric(Trim$(txtCantidad.Value)) Then
        MostrarError "Ingrese una cantidad numerica entera."
        Exit Sub
    End If
    cant = CLng(txtCantidad.Value)
    tipo = Left$(Trim$(cboTipo.Text), 1)
    idProv = ObtenerIdProveedorSeleccionado()
    motivo = Trim$(CStr(txtMotivo.Value))

    On Error GoTo Fallo
    Set conn = modConexion.AbrirConexion()
    idMov = modInventario.RegistrarMovimientoInventario(conn, sku, modSesion.gUsuarioId, _
        tipo, cant, motivo, idProv, msg, True)
    If idMov > 0 Then
        stockNuevo = modInventario.ObtenerStock(conn, sku)
        modInventario.VolcarCatalogoAHoja conn, modSesion.gUsuarioRol
    End If
    modConexion.CerrarConexion conn

    If idMov > 0 Then
        MostrarOk "Movimiento #" & idMov & " registrado para " & sku & ". Nuevo stock: " & stockNuevo & " u. (Hojas Excel actualizadas)."
        txtCantidad.Value = ""
        txtMotivo.Value = ""
        CargarListaProductos
    Else
        MostrarError msg
    End If
    Exit Sub
Fallo:
    errDesc = Err.Description
    modConexion.CerrarConexion conn
    MostrarError "Error al guardar movimiento: " & errDesc
End Sub

Private Sub cmdVolcarHoja_Click()
    On Error GoTo Fallo
    Dim conn As Object
    Set conn = modConexion.AbrirConexion()
    modInventario.VolcarCatalogoAHoja conn, modSesion.gUsuarioRol
    modInventario.VolcarTablasMaestrasAHoja conn, modSesion.gUsuarioRol
    modConexion.CerrarConexion conn
    MostrarOk "Hojas 'Productos', 'Inventario', 'Clientes' y 'Proveedores' sincronizadas."
    Exit Sub
Fallo:
    modConexion.CerrarConexion conn
    MostrarError "No se pudo actualizar la hoja: " & Err.Description
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
