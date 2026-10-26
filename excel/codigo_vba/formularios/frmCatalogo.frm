VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCatalogo 
   Caption         =   "Tienda de Fútbol - Catálogo de Productos"
   ClientHeight    =   9400
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   15800
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.Label lblBanner 
      Caption         =   ""
      Height          =   960
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15800
   End
   Begin MSForms.Label lblAccent 
      Caption         =   ""
      Height          =   70
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   15800
   End
   Begin MSForms.Image imgLogo 
      Height          =   760
      Left            =   220
      TabIndex        =   2
      Top             =   100
      Width           =   760
   End
   Begin MSForms.Label lblTitulo 
      Caption         =   "CATÁLOGO DE PRODUCTOS"
      Height          =   460
      Left            =   1150
      TabIndex        =   3
      Top             =   150
      Width           =   8000
   End
   Begin MSForms.Label lblSubtitulo 
      Caption         =   "Fotografías de estudio, stock en tiempo real y compra con boleta electrónica"
      Height          =   280
      Left            =   1150
      TabIndex        =   4
      Top             =   620
      Width           =   9000
   End
   Begin MSForms.Label lblUsuario 
      Caption         =   ""
      Height          =   300
      Left            =   9600
      TabIndex        =   5
      Top             =   380
      Width           =   5960
   End
   Begin MSForms.Label lblCategoria 
      Caption         =   "CATEGORÍA"
      Height          =   240
      Left            =   320
      TabIndex        =   6
      Top             =   1180
      Width           =   2600
   End
   Begin MSForms.ComboBox cboCategoria 
      Height          =   360
      Left            =   320
      TabIndex        =   7
      Top             =   1430
      Width           =   2600
   End
   Begin MSForms.Label lblBuscar 
      Caption         =   "BUSCAR (NOMBRE, SKU, MARCA)"
      Height          =   240
      Left            =   3100
      TabIndex        =   8
      Top             =   1180
      Width           =   3000
   End
   Begin MSForms.TextBox txtBuscar 
      MaxLength       =   60
      Height          =   360
      Left            =   3100
      TabIndex        =   9
      Top             =   1430
      Width           =   3000
   End
   Begin MSForms.CommandButton cmdAnterior 
      Caption         =   "< Anterior"
      Height          =   380
      Left            =   6300
      TabIndex        =   10
      Top             =   1410
      Width           =   1300
   End
   Begin MSForms.Label lblPagina 
      Caption         =   "Página 1 de 1"
      Height          =   280
      Left            =   7650
      TabIndex        =   11
      Top             =   1470
      Width           =   1100
   End
   Begin MSForms.CommandButton cmdSiguiente 
      Caption         =   "Siguiente >"
      Height          =   380
      Left            =   8800
      TabIndex        =   12
      Top             =   1410
      Width           =   1020
   End
   Begin MSForms.Label lblCardT1 
      Caption         =   ""
      Height          =   3300
      Left            =   320
      TabIndex        =   13
      Top             =   1850
      Width           =   3100
   End
   Begin MSForms.Image imgP1 
      Height          =   2000
      Left            =   870
      TabIndex        =   14
      Top             =   1930
      Width           =   2000
   End
   Begin MSForms.Label lblNom1 
      Caption         =   ""
      WordWrap        =   -1  'True
      Height          =   400
      Left            =   440
      TabIndex        =   15
      Top             =   3960
      Width           =   2860
   End
   Begin MSForms.Label lblPrecio1 
      Caption         =   ""
      Height          =   290
      Left            =   440
      TabIndex        =   16
      Top             =   4370
      Width           =   1600
   End
   Begin MSForms.Label lblStock1 
      Caption         =   ""
      Height          =   260
      Left            =   2040
      TabIndex        =   17
      Top             =   4395
      Width           =   1260
   End
   Begin MSForms.CommandButton cmdAgregar1 
      Caption         =   "AGREGAR AL CARRITO"
      Height          =   340
      Left            =   440
      TabIndex        =   18
      Top             =   4700
      Width           =   2860
   End
   Begin MSForms.Label lblCardT2 
      Caption         =   ""
      Height          =   3300
      Left            =   3520
      TabIndex        =   19
      Top             =   1850
      Width           =   3100
   End
   Begin MSForms.Image imgP2 
      Height          =   2000
      Left            =   4070
      TabIndex        =   20
      Top             =   1930
      Width           =   2000
   End
   Begin MSForms.Label lblNom2 
      Caption         =   ""
      WordWrap        =   -1  'True
      Height          =   400
      Left            =   3640
      TabIndex        =   21
      Top             =   3960
      Width           =   2860
   End
   Begin MSForms.Label lblPrecio2 
      Caption         =   ""
      Height          =   290
      Left            =   3640
      TabIndex        =   22
      Top             =   4370
      Width           =   1600
   End
   Begin MSForms.Label lblStock2 
      Caption         =   ""
      Height          =   260
      Left            =   5240
      TabIndex        =   23
      Top             =   4395
      Width           =   1260
   End
   Begin MSForms.CommandButton cmdAgregar2 
      Caption         =   "AGREGAR AL CARRITO"
      Height          =   340
      Left            =   3640
      TabIndex        =   24
      Top             =   4700
      Width           =   2860
   End
   Begin MSForms.Label lblCardT3 
      Caption         =   ""
      Height          =   3300
      Left            =   6720
      TabIndex        =   25
      Top             =   1850
      Width           =   3100
   End
   Begin MSForms.Image imgP3 
      Height          =   2000
      Left            =   7270
      TabIndex        =   26
      Top             =   1930
      Width           =   2000
   End
   Begin MSForms.Label lblNom3 
      Caption         =   ""
      WordWrap        =   -1  'True
      Height          =   400
      Left            =   6840
      TabIndex        =   27
      Top             =   3960
      Width           =   2860
   End
   Begin MSForms.Label lblPrecio3 
      Caption         =   ""
      Height          =   290
      Left            =   6840
      TabIndex        =   28
      Top             =   4370
      Width           =   1600
   End
   Begin MSForms.Label lblStock3 
      Caption         =   ""
      Height          =   260
      Left            =   8440
      TabIndex        =   29
      Top             =   4395
      Width           =   1260
   End
   Begin MSForms.CommandButton cmdAgregar3 
      Caption         =   "AGREGAR AL CARRITO"
      Height          =   340
      Left            =   6840
      TabIndex        =   30
      Top             =   4700
      Width           =   2860
   End
   Begin MSForms.Label lblCardT4 
      Caption         =   ""
      Height          =   3300
      Left            =   320
      TabIndex        =   31
      Top             =   5250
      Width           =   3100
   End
   Begin MSForms.Image imgP4 
      Height          =   2000
      Left            =   870
      TabIndex        =   32
      Top             =   5330
      Width           =   2000
   End
   Begin MSForms.Label lblNom4 
      Caption         =   ""
      WordWrap        =   -1  'True
      Height          =   400
      Left            =   440
      TabIndex        =   33
      Top             =   7360
      Width           =   2860
   End
   Begin MSForms.Label lblPrecio4 
      Caption         =   ""
      Height          =   290
      Left            =   440
      TabIndex        =   34
      Top             =   7770
      Width           =   1600
   End
   Begin MSForms.Label lblStock4 
      Caption         =   ""
      Height          =   260
      Left            =   2040
      TabIndex        =   35
      Top             =   7795
      Width           =   1260
   End
   Begin MSForms.CommandButton cmdAgregar4 
      Caption         =   "AGREGAR AL CARRITO"
      Height          =   340
      Left            =   440
      TabIndex        =   36
      Top             =   8100
      Width           =   2860
   End
   Begin MSForms.Label lblCardT5 
      Caption         =   ""
      Height          =   3300
      Left            =   3520
      TabIndex        =   37
      Top             =   5250
      Width           =   3100
   End
   Begin MSForms.Image imgP5 
      Height          =   2000
      Left            =   4070
      TabIndex        =   38
      Top             =   5330
      Width           =   2000
   End
   Begin MSForms.Label lblNom5 
      Caption         =   ""
      WordWrap        =   -1  'True
      Height          =   400
      Left            =   3640
      TabIndex        =   39
      Top             =   7360
      Width           =   2860
   End
   Begin MSForms.Label lblPrecio5 
      Caption         =   ""
      Height          =   290
      Left            =   3640
      TabIndex        =   40
      Top             =   7770
      Width           =   1600
   End
   Begin MSForms.Label lblStock5 
      Caption         =   ""
      Height          =   260
      Left            =   5240
      TabIndex        =   41
      Top             =   7795
      Width           =   1260
   End
   Begin MSForms.CommandButton cmdAgregar5 
      Caption         =   "AGREGAR AL CARRITO"
      Height          =   340
      Left            =   3640
      TabIndex        =   42
      Top             =   8100
      Width           =   2860
   End
   Begin MSForms.Label lblCardT6 
      Caption         =   ""
      Height          =   3300
      Left            =   6720
      TabIndex        =   43
      Top             =   5250
      Width           =   3100
   End
   Begin MSForms.Image imgP6 
      Height          =   2000
      Left            =   7270
      TabIndex        =   44
      Top             =   5330
      Width           =   2000
   End
   Begin MSForms.Label lblNom6 
      Caption         =   ""
      WordWrap        =   -1  'True
      Height          =   400
      Left            =   6840
      TabIndex        =   45
      Top             =   7360
      Width           =   2860
   End
   Begin MSForms.Label lblPrecio6 
      Caption         =   ""
      Height          =   290
      Left            =   6840
      TabIndex        =   46
      Top             =   7770
      Width           =   1600
   End
   Begin MSForms.Label lblStock6 
      Caption         =   ""
      Height          =   260
      Left            =   8440
      TabIndex        =   47
      Top             =   7795
      Width           =   1260
   End
   Begin MSForms.CommandButton cmdAgregar6 
      Caption         =   "AGREGAR AL CARRITO"
      Height          =   340
      Left            =   6840
      TabIndex        =   48
      Top             =   8100
      Width           =   2860
   End
   Begin MSForms.Label lblVacio 
      Caption         =   "No se encontraron productos con ese filtro."
      Height          =   400
      Left            =   2200
      TabIndex        =   49
      Top             =   4200
      Width           =   5500
   End
   Begin MSForms.Label lblCardCarrito 
      Caption         =   ""
      Height          =   7350
      Left            =   10050
      TabIndex        =   50
      Top             =   1200
      Width           =   5550
   End
   Begin MSForms.Label lblCarritoTitulo 
      Caption         =   "MI CARRITO"
      Height          =   340
      Left            =   10250
      TabIndex        =   51
      Top             =   1290
      Width           =   5100
   End
   Begin MSForms.ListBox lstCarrito 
      Height          =   3250
      Left            =   10250
      TabIndex        =   52
      Top             =   1680
      Width           =   5150
   End
   Begin MSForms.CommandButton cmdMenos 
      Caption         =   "-"
      Height          =   340
      Left            =   10250
      TabIndex        =   53
      Top             =   5010
      Width           =   900
   End
   Begin MSForms.CommandButton cmdMas 
      Caption         =   "+"
      Height          =   340
      Left            =   11200
      TabIndex        =   54
      Top             =   5010
      Width           =   900
   End
   Begin MSForms.CommandButton cmdQuitar 
      Caption         =   "Quitar"
      Height          =   340
      Left            =   12150
      TabIndex        =   55
      Top             =   5010
      Width           =   1500
   End
   Begin MSForms.CommandButton cmdVaciar 
      Caption         =   "Vaciar"
      Height          =   340
      Left            =   13700
      TabIndex        =   56
      Top             =   5010
      Width           =   1700
   End
   Begin MSForms.Label lblCliente 
      Caption         =   "CLIENTE DE LA VENTA"
      Height          =   260
      Left            =   10250
      TabIndex        =   57
      Top             =   5480
      Width           =   5100
   End
   Begin MSForms.ComboBox cboCliente 
      Height          =   360
      Left            =   10250
      TabIndex        =   58
      Top             =   5740
      Width           =   5150
   End
   Begin MSForms.Label lblTotales 
      Caption         =   ""
      WordWrap        =   -1  'True
      Height          =   800
      Left            =   10250
      TabIndex        =   59
      Top             =   6250
      Width           =   5150
   End
   Begin MSForms.Label lblMensaje 
      Caption         =   ""
      WordWrap        =   -1  'True
      Height          =   700
      Left            =   10250
      TabIndex        =   60
      Top             =   7100
      Width           =   5150
   End
   Begin MSForms.CommandButton cmdConfirmar 
      Caption         =   "CONFIRMAR COMPRA"
      Default         =   -1  'True
      Height          =   560
      Left            =   10250
      TabIndex        =   61
      Top             =   7850
      Width           =   5150
   End
   Begin MSForms.CommandButton cmdCerrar 
      Caption         =   "Volver al menú"
      Cancel          =   -1  'True
      Height          =   420
      Left            =   320
      TabIndex        =   62
      Top             =   8750
      Width           =   3000
   End
End
Attribute VB_Name = "frmCatalogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const POR_PAGINA As Long = 6

' Catalogo en memoria (nunca se cargan costos)
Private pSku() As String
Private pNom() As String
Private pCat() As String
Private pMarca() As String
Private pTalla() As String
Private pPrecio() As Currency
Private pStock() As Long
Private mNumProd As Long

Private mFiltro() As Long
Private mNumFiltro As Long
Private mPagina As Long
Private mCargando As Boolean

' Carrito
Private mSkus() As String
Private mNombres() As String
Private mCantidades() As Long
Private mPrecios() As Currency
Private mNumItems As Long

Private Sub UserForm_Initialize()
    Dim esStaff As Boolean, t As Long

    mCargando = True
    esStaff = modSesion.PuedeVender()
    mNumItems = 0
    mNumProd = 0

    modEstilo.AplicarTema Me
    modEstilo.CargarLogo imgLogo
    EstiloExtra

    lblUsuario.Caption = "Hola, " & modSesion.gUsuarioNombre & "   |   " & UCase$(modSesion.gUsuarioRol)
    lblMensaje.Caption = ""
    lblVacio.Visible = False

    lstCarrito.ColumnCount = 3
    lstCarrito.ColumnWidths = "140 pt;34 pt;66 pt"

    cboCategoria.Style = 2
    cboCliente.Style = 2
    lblCliente.Visible = esStaff
    cboCliente.Visible = esStaff
    If esStaff Then
        cmdConfirmar.Caption = "REGISTRAR VENTA"
    Else
        cmdConfirmar.Caption = "COMPRAR AHORA"
    End If

    CargarProductos
    CargarCategorias
    If esStaff Then CargarClientes
    ActualizarCarrito

    mCargando = False
    AplicarFiltro
End Sub

Private Sub EstiloExtra()
    Dim t As Long
    On Error Resume Next
    lblUsuario.TextAlign = 3
    lblUsuario.ForeColor = modEstilo.C_SLATE300
    lblUsuario.Font.Size = 10
    lblUsuario.BackStyle = 0

    lblCategoria.BackStyle = 0: lblCategoria.Font.Bold = True: lblCategoria.Font.Size = 8: lblCategoria.ForeColor = modEstilo.C_SLATE600
    lblBuscar.BackStyle = 0: lblBuscar.Font.Bold = True: lblBuscar.Font.Size = 8: lblBuscar.ForeColor = modEstilo.C_SLATE600
    lblCliente.BackStyle = 0: lblCliente.Font.Bold = True: lblCliente.Font.Size = 8: lblCliente.ForeColor = modEstilo.C_SLATE600
    lblPagina.BackStyle = 0: lblPagina.Font.Bold = True: lblPagina.Font.Size = 9: lblPagina.ForeColor = modEstilo.C_SLATE600
    lblPagina.TextAlign = 2
    lblVacio.BackStyle = 0: lblVacio.Font.Size = 12: lblVacio.ForeColor = modEstilo.C_SLATE500: lblVacio.TextAlign = 2

    lblCarritoTitulo.BackStyle = 0
    lblCarritoTitulo.Font.Name = modEstilo.FUENTE_TITULO
    lblCarritoTitulo.Font.Size = 14
    lblCarritoTitulo.Font.Bold = True
    lblCarritoTitulo.ForeColor = modEstilo.C_NAVY

    lblTotales.BackStyle = 0
    lblTotales.Font.Name = modEstilo.FUENTE_TITULO
    lblTotales.Font.Size = 11
    lblTotales.Font.Bold = True
    lblTotales.ForeColor = modEstilo.C_NAVY

    lblMensaje.BackStyle = 0
    lblMensaje.Font.Size = 8.5
    lblMensaje.Font.Bold = True

    lstCarrito.Font.Size = 9
    cmdConfirmar.Font.Size = 12
    cmdMas.Font.Size = 12
    cmdMenos.Font.Size = 12

    For t = 1 To POR_PAGINA
        With Me.Controls("imgP" & t)
            .PictureSizeMode = 3
            .BorderStyle = 0
            .BackColor = modEstilo.C_BLANCO
        End With
        With Me.Controls("lblNom" & t)
            .BackStyle = 0
            .Font.Bold = True
            .Font.Size = 9
            .ForeColor = modEstilo.C_SLATE800
            .TextAlign = 2
        End With
        With Me.Controls("lblPrecio" & t)
            .BackStyle = 0
            .Font.Name = modEstilo.FUENTE_TITULO
            .Font.Bold = True
            .Font.Size = 14
            .ForeColor = modEstilo.C_EMERALD_D
        End With
        With Me.Controls("lblStock" & t)
            .BackStyle = 0
            .Font.Size = 8.5
            .TextAlign = 3
        End With
    Next t
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------
' Datos
' ---------------------------------------------------------------------------
Private Sub CargarProductos()
    On Error GoTo Fallo
    Dim conn As Object, rs As Object, errDesc As String

    mNumProd = 0
    Erase pSku: Erase pNom: Erase pCat: Erase pMarca: Erase pTalla: Erase pPrecio: Erase pStock
    Set conn = modConexion.AbrirConexion()
    Set rs = conn.Execute("SELECT p.sku, p.nombre, c.nombre AS categoria, p.marca, p.talla, " & _
        "p.precio_venta, p.stock FROM PRODUCTOS p INNER JOIN CATEGORIAS c ON p.id_categoria = c.id " & _
        "ORDER BY p.sku")
    Do While Not rs.EOF
        ReDim Preserve pSku(0 To mNumProd)
        ReDim Preserve pNom(0 To mNumProd)
        ReDim Preserve pCat(0 To mNumProd)
        ReDim Preserve pMarca(0 To mNumProd)
        ReDim Preserve pTalla(0 To mNumProd)
        ReDim Preserve pPrecio(0 To mNumProd)
        ReDim Preserve pStock(0 To mNumProd)
        pSku(mNumProd) = CStr(rs.Fields("sku").Value)
        pNom(mNumProd) = CStr(rs.Fields("nombre").Value)
        pCat(mNumProd) = CStr(rs.Fields("categoria").Value)
        If Not IsNull(rs.Fields("marca").Value) Then pMarca(mNumProd) = CStr(rs.Fields("marca").Value)
        If Not IsNull(rs.Fields("talla").Value) Then pTalla(mNumProd) = CStr(rs.Fields("talla").Value)
        pPrecio(mNumProd) = CCur(rs.Fields("precio_venta").Value)
        pStock(mNumProd) = CLng(rs.Fields("stock").Value)
        mNumProd = mNumProd + 1
        rs.MoveNext
    Loop
    rs.Close
    modConexion.CerrarConexion conn
    Exit Sub
Fallo:
    errDesc = Err.Description
    modConexion.CerrarConexion conn
    MostrarMsg "No se pudo cargar el catalogo: " & errDesc, True
End Sub

Private Sub CargarCategorias()
    Dim i As Long, j As Long, existe As Boolean
    cboCategoria.Clear
    cboCategoria.AddItem "Todas"
    For i = 0 To mNumProd - 1
        existe = False
        For j = 1 To cboCategoria.ListCount - 1
            If StrComp(cboCategoria.List(j), pCat(i), vbTextCompare) = 0 Then existe = True: Exit For
        Next j
        If Not existe Then cboCategoria.AddItem pCat(i)
    Next i
    cboCategoria.ListIndex = 0
End Sub

Private Sub CargarClientes()
    On Error GoTo Fallo
    Dim conn As Object, rs As Object, errDesc As String
    cboCliente.Clear
    Set conn = modConexion.AbrirConexion()
    Set rs = conn.Execute("SELECT id, rut, nombre, apellido FROM CLIENTES ORDER BY nombre, apellido")
    Do While Not rs.EOF
        cboCliente.AddItem CStr(rs.Fields("id").Value) & " | " & CStr(rs.Fields("rut").Value) & " - " & _
            CStr(rs.Fields("nombre").Value) & " " & CStr(rs.Fields("apellido").Value)
        rs.MoveNext
    Loop
    rs.Close
    modConexion.CerrarConexion conn
    If cboCliente.ListCount > 0 Then cboCliente.ListIndex = 0
    Exit Sub
Fallo:
    errDesc = Err.Description
    modConexion.CerrarConexion conn
    MostrarMsg "No se pudieron cargar los clientes: " & errDesc, True
End Sub

Private Function ObtenerIdClienteSeleccionado() As Long
    Dim s As String, p As Long
    s = CStr(cboCliente.Value)
    p = InStr(s, " | ")
    If p > 1 Then
        If IsNumeric(Left$(s, p - 1)) Then ObtenerIdClienteSeleccionado = CLng(Left$(s, p - 1))
    End If
End Function

' ---------------------------------------------------------------------------
' Filtro y paginas
' ---------------------------------------------------------------------------
Private Sub AplicarFiltro()
    Dim i As Long, cat As String, txt As String, ok As Boolean
    If mCargando Then Exit Sub

    cat = CStr(cboCategoria.Value)
    txt = LCase$(Trim$(CStr(txtBuscar.Value)))
    mNumFiltro = 0
    ReDim mFiltro(0 To mNumProd)
    For i = 0 To mNumProd - 1
        ok = True
        If Len(cat) > 0 And StrComp(cat, "Todas", vbTextCompare) <> 0 Then
            If StrComp(pCat(i), cat, vbTextCompare) <> 0 Then ok = False
        End If
        If ok And Len(txt) > 0 Then
            If InStr(1, LCase$(pNom(i) & " " & pSku(i) & " " & pMarca(i) & " " & pCat(i)), txt) = 0 Then ok = False
        End If
        If ok Then
            mFiltro(mNumFiltro) = i
            mNumFiltro = mNumFiltro + 1
        End If
    Next i
    mPagina = 0
    MostrarPagina
End Sub

Private Sub MostrarPagina()
    Dim t As Long, idx As Long, totalPag As Long

    totalPag = (mNumFiltro + POR_PAGINA - 1) \ POR_PAGINA
    If totalPag < 1 Then totalPag = 1
    If mPagina > totalPag - 1 Then mPagina = totalPag - 1
    If mPagina < 0 Then mPagina = 0

    For t = 1 To POR_PAGINA
        idx = mPagina * POR_PAGINA + (t - 1)
        If idx < mNumFiltro Then
            MostrarTile t, mFiltro(idx)
        Else
            OcultarTile t
        End If
    Next t

    lblPagina.Caption = "Pagina " & (mPagina + 1) & " de " & totalPag
    cmdAnterior.Enabled = (mPagina > 0)
    cmdSiguiente.Enabled = (mPagina < totalPag - 1)
    lblVacio.Visible = (mNumFiltro = 0)
End Sub

Private Sub MostrarTile(ByVal t As Long, ByVal p As Long)
    Dim disp As Long, ruta As String

    disp = pStock(p) - CantidadEnCarrito(pSku(p))
    If disp < 0 Then disp = 0

    Me.Controls("lblCardT" & t).Visible = True

    With Me.Controls("lblNom" & t)
        .Visible = True
        .Caption = pNom(p)
    End With
    With Me.Controls("lblPrecio" & t)
        .Visible = True
        .Caption = "$" & Format$(pPrecio(p), "#,##0")
    End With
    With Me.Controls("lblStock" & t)
        .Visible = True
        If disp > 0 Then
            .Caption = "Stock: " & disp
            .ForeColor = modEstilo.C_SLATE500
        Else
            .Caption = "AGOTADO"
            .ForeColor = modEstilo.C_ROJO
        End If
    End With

    On Error Resume Next
    ruta = modInventario.ObtenerRutaImagenProducto(pSku(p), pCat(p))
    With Me.Controls("imgP" & t)
        .Visible = True
        If Len(ruta) > 0 Then
            Set .Picture = LoadPicture(ruta)
        Else
            Set .Picture = Nothing
        End If
    End With
    On Error GoTo 0

    With Me.Controls("cmdAgregar" & t)
        .Visible = True
        .Enabled = (disp > 0)
        If disp > 0 Then
            .Caption = "AGREGAR AL CARRITO"
        Else
            .Caption = "SIN STOCK"
        End If
    End With
    If disp > 0 Then
        modEstilo.EstiloBoton Me.Controls("cmdAgregar" & t), "P"
    Else
        modEstilo.EstiloBoton Me.Controls("cmdAgregar" & t), "N"
    End If
End Sub

Private Sub OcultarTile(ByVal t As Long)
    Me.Controls("lblCardT" & t).Visible = False
    Me.Controls("imgP" & t).Visible = False
    Me.Controls("lblNom" & t).Visible = False
    Me.Controls("lblPrecio" & t).Visible = False
    Me.Controls("lblStock" & t).Visible = False
    Me.Controls("cmdAgregar" & t).Visible = False
End Sub

Private Sub cboCategoria_Change()
    AplicarFiltro
End Sub

Private Sub txtBuscar_Change()
    AplicarFiltro
End Sub

Private Sub cmdAnterior_Click()
    If mPagina > 0 Then
        mPagina = mPagina - 1
        MostrarPagina
    End If
End Sub

Private Sub cmdSiguiente_Click()
    mPagina = mPagina + 1
    MostrarPagina
End Sub

' ---------------------------------------------------------------------------
' Carrito
' ---------------------------------------------------------------------------
Private Function IndiceCarrito(ByVal sku As String) As Long
    Dim i As Long
    IndiceCarrito = -1
    For i = 0 To mNumItems - 1
        If StrComp(mSkus(i), sku, vbTextCompare) = 0 Then
            IndiceCarrito = i
            Exit Function
        End If
    Next i
End Function

Private Function CantidadEnCarrito(ByVal sku As String) As Long
    Dim i As Long
    i = IndiceCarrito(sku)
    If i >= 0 Then CantidadEnCarrito = mCantidades(i)
End Function

Private Function IndiceProducto(ByVal sku As String) As Long
    Dim i As Long
    IndiceProducto = -1
    For i = 0 To mNumProd - 1
        If StrComp(pSku(i), sku, vbTextCompare) = 0 Then
            IndiceProducto = i
            Exit Function
        End If
    Next i
End Function

Private Sub AgregarAlCarrito(ByVal t As Long)
    Dim idxF As Long, p As Long, i As Long

    idxF = mPagina * POR_PAGINA + (t - 1)
    If idxF >= mNumFiltro Then Exit Sub
    p = mFiltro(idxF)

    If CantidadEnCarrito(pSku(p)) + 1 > pStock(p) Then
        MostrarMsg "No queda mas stock de " & pSku(p) & ".", True
        Exit Sub
    End If

    i = IndiceCarrito(pSku(p))
    If i >= 0 Then
        mCantidades(i) = mCantidades(i) + 1
    Else
        ReDim Preserve mSkus(0 To mNumItems)
        ReDim Preserve mNombres(0 To mNumItems)
        ReDim Preserve mCantidades(0 To mNumItems)
        ReDim Preserve mPrecios(0 To mNumItems)
        mSkus(mNumItems) = pSku(p)
        mNombres(mNumItems) = pNom(p)
        mCantidades(mNumItems) = 1
        mPrecios(mNumItems) = pPrecio(p)
        i = mNumItems
        mNumItems = mNumItems + 1
    End If

    ActualizarCarrito i
    MostrarPagina
    MostrarMsg "Agregado: " & pNom(p), False
End Sub

Private Sub cmdAgregar1_Click()
    AgregarAlCarrito 1
End Sub

Private Sub cmdAgregar2_Click()
    AgregarAlCarrito 2
End Sub

Private Sub cmdAgregar3_Click()
    AgregarAlCarrito 3
End Sub

Private Sub cmdAgregar4_Click()
    AgregarAlCarrito 4
End Sub

Private Sub cmdAgregar5_Click()
    AgregarAlCarrito 5
End Sub

Private Sub cmdAgregar6_Click()
    AgregarAlCarrito 6
End Sub


Private Sub ActualizarCarrito(Optional ByVal seleccion As Long = -1)
    Dim i As Long
    Dim subLinea As Currency, subtotal As Currency, iva As Currency, total As Currency

    lstCarrito.Clear
    For i = 0 To mNumItems - 1
        subLinea = CCur(Round(CDbl(mPrecios(i)) * CDbl(mCantidades(i)), 2))
        subtotal = subtotal + subLinea
        lstCarrito.AddItem mNombres(i)
        lstCarrito.List(i, 1) = CStr(mCantidades(i))
        lstCarrito.List(i, 2) = "$" & Format$(subLinea, "#,##0")
    Next i

    iva = CCur(Round(CDbl(subtotal) * modVentas.TASA_IVA, 2))
    total = subtotal + iva
    lblTotales.Caption = "Subtotal neto:  $" & Format$(subtotal, "#,##0") & vbCrLf & _
                         "IVA (19%):  $" & Format$(iva, "#,##0") & vbCrLf & _
                         "TOTAL A PAGAR:  $" & Format$(total, "#,##0")
    If seleccion >= 0 And seleccion < lstCarrito.ListCount Then lstCarrito.ListIndex = seleccion
    cmdConfirmar.Enabled = (mNumItems > 0)
    cmdMas.Enabled = (mNumItems > 0)
    cmdMenos.Enabled = (mNumItems > 0)
    cmdQuitar.Enabled = (mNumItems > 0)
    cmdVaciar.Enabled = (mNumItems > 0)
End Sub

Private Sub QuitarItem(ByVal i As Long)
    Dim k As Long
    For k = i To mNumItems - 2
        mSkus(k) = mSkus(k + 1)
        mNombres(k) = mNombres(k + 1)
        mCantidades(k) = mCantidades(k + 1)
        mPrecios(k) = mPrecios(k + 1)
    Next k
    mNumItems = mNumItems - 1
    If mNumItems > 0 Then
        ReDim Preserve mSkus(0 To mNumItems - 1)
        ReDim Preserve mNombres(0 To mNumItems - 1)
        ReDim Preserve mCantidades(0 To mNumItems - 1)
        ReDim Preserve mPrecios(0 To mNumItems - 1)
    Else
        Erase mSkus: Erase mNombres: Erase mCantidades: Erase mPrecios
    End If
End Sub

Private Sub cmdMas_Click()
    Dim i As Long, p As Long
    i = lstCarrito.ListIndex
    If i < 0 Then
        MostrarMsg "Seleccione un producto del carrito.", True
        Exit Sub
    End If
    p = IndiceProducto(mSkus(i))
    If p >= 0 Then
        If mCantidades(i) + 1 > pStock(p) Then
            MostrarMsg "No queda mas stock de " & mSkus(i) & ".", True
            Exit Sub
        End If
    End If
    mCantidades(i) = mCantidades(i) + 1
    ActualizarCarrito i
    MostrarPagina
End Sub

Private Sub cmdMenos_Click()
    Dim i As Long
    i = lstCarrito.ListIndex
    If i < 0 Then
        MostrarMsg "Seleccione un producto del carrito.", True
        Exit Sub
    End If
    If mCantidades(i) > 1 Then
        mCantidades(i) = mCantidades(i) - 1
        ActualizarCarrito i
    Else
        QuitarItem i
        ActualizarCarrito
    End If
    MostrarPagina
End Sub

Private Sub cmdQuitar_Click()
    Dim i As Long
    i = lstCarrito.ListIndex
    If i < 0 Then
        MostrarMsg "Seleccione un producto del carrito.", True
        Exit Sub
    End If
    QuitarItem i
    ActualizarCarrito
    MostrarPagina
End Sub

Private Sub cmdVaciar_Click()
    mNumItems = 0
    Erase mSkus: Erase mNombres: Erase mCantidades: Erase mPrecios
    ActualizarCarrito
    MostrarPagina
    MostrarMsg "", False
End Sub

' ---------------------------------------------------------------------------
' Confirmar compra / venta
' ---------------------------------------------------------------------------
Private Sub cmdConfirmar_Click()
    On Error GoTo Fallo
    Dim conn As Object
    Dim idCliente As Long, idVenta As Long, idBoleta As Long
    Dim corr As String, msg As String, errDesc As String
    Dim rutaPdf As String, msgPdf As String, destCorreo As String, msgCorreo As String
    Dim resumen As String, okPdf As Boolean, okMail As Boolean, avisoCorreo As String

    If Not modSesion.PuedeComprarEnCatalogo() Then
        MostrarMsg "Debe iniciar sesion para comprar.", True
        Exit Sub
    End If
    If mNumItems <= 0 Then
        MostrarMsg "Agregue al menos un producto al carrito.", True
        Exit Sub
    End If

    MostrarMsg "Procesando... (puede tardar si se envia la boleta por correo)", False
    DoEvents

    Set conn = modConexion.AbrirConexion()
    If StrComp(modSesion.gUsuarioRol, "Cliente", vbTextCompare) = 0 Then
        idCliente = modVentas.ObtenerClienteParaSesion(conn)
    Else
        idCliente = ObtenerIdClienteSeleccionado()
    End If
    If idCliente <= 0 Then
        modConexion.CerrarConexion conn
        MostrarMsg "Seleccione el cliente de la venta.", True
        Exit Sub
    End If

    idVenta = modVentas.RegistrarVenta(conn, modSesion.gUsuarioId, idCliente, _
        mSkus, mCantidades, msg, idBoleta, corr, True, _
        (StrComp(modSesion.gUsuarioRol, "Cliente", vbTextCompare) = 0))

    If idVenta > 0 Then
        resumen = "Compra registrada: boleta " & corr & "."

        On Error Resume Next
        okPdf = modBoletas.GenerarBoletaPDF(conn, idVenta, rutaPdf, msgPdf)
        If Err.Number <> 0 Then msgPdf = Err.Description: okPdf = False: Err.Clear
        On Error GoTo Fallo

        If okPdf Then
            resumen = resumen & " PDF generado."
            If Len(modBoletas.LeerConfigTexto(modBoletas.NOMBRE_RANGO_GMAIL_CLAVE)) > 0 Then
                On Error Resume Next
                okMail = modBoletas.EnviarBoletaPorCorreo(conn, idVenta, destCorreo, msgCorreo, False)
                If Err.Number <> 0 Then msgCorreo = Err.Description: okMail = False: Err.Clear
                On Error GoTo Fallo
                If okMail Then
                    resumen = resumen & " Enviada a " & destCorreo & "."
                Else
                    avisoCorreo = "La compra quedo registrada, pero el correo NO se pudo enviar:" & vbCrLf & msgCorreo
                    resumen = resumen & " (Correo no enviado)"
                End If
            Else
                resumen = resumen & " (Correo no configurado)"
            End If
        Else
            resumen = resumen & " No se pudo generar el PDF: " & msgPdf
        End If

        On Error Resume Next
        modVentas.VolcarVentasAHoja conn
        modInventario.VolcarCatalogoAHoja conn, modSesion.gUsuarioRol
        On Error GoTo Fallo
    End If
    modConexion.CerrarConexion conn

    If idVenta > 0 Then
        mNumItems = 0
        Erase mSkus: Erase mNombres: Erase mCantidades: Erase mPrecios
        ActualizarCarrito
        CargarProductos
        MostrarPagina
        MostrarMsg resumen, False
        If Len(avisoCorreo) > 0 Then MsgBox avisoCorreo, vbExclamation, "Tienda de Futbol"
        If Len(rutaPdf) > 0 Then
            If MsgBox("Compra registrada (" & corr & ")." & vbCrLf & "Desea abrir la boleta PDF?", _
                      vbYesNo + vbQuestion, "Tienda de Futbol") = vbYes Then
                On Error Resume Next
                CreateObject("WScript.Shell").Run """" & rutaPdf & """"
                On Error GoTo Fallo
            End If
        End If
    Else
        MostrarMsg msg, True
    End If
    Exit Sub
Fallo:
    errDesc = Err.Description
    modConexion.CerrarConexion conn
    MostrarMsg "No se pudo completar la compra: " & errDesc, True
End Sub

Private Sub cmdCerrar_Click()
    If mNumItems > 0 Then
        If MsgBox("Tiene productos en el carrito. Desea salir y descartarlos?", _
                  vbYesNo + vbQuestion, "Tienda de Futbol") = vbNo Then Exit Sub
    End If
    Unload Me
End Sub

Private Sub MostrarMsg(ByVal texto As String, ByVal esError As Boolean)
    modEstilo.Mensaje lblMensaje, texto, esError
End Sub
