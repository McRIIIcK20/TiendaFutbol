VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMenu 
   Caption         =   "Tienda de Fútbol - Panel Principal"
   ClientHeight    =   7500
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9000
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.Label lblBanner 
      Caption         =   ""
      Height          =   1160
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9000
   End
   Begin MSForms.Label lblAccent 
      Caption         =   ""
      Height          =   80
      Left            =   0
      TabIndex        =   1
      Top             =   1160
      Width           =   9000
   End
   Begin MSForms.Image imgLogo 
      Height          =   920
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   920
   End
   Begin MSForms.Label lblTitulo 
      Caption         =   "PANEL PRINCIPAL"
      Height          =   440
      Left            =   1400
      TabIndex        =   3
      Top             =   200
      Width           =   7400
   End
   Begin MSForms.Label lblSubtitulo 
      Caption         =   "Tienda de Fútbol · Ventas, inventario y facturación"
      Height          =   300
      Left            =   1400
      TabIndex        =   4
      Top             =   700
      Width           =   7400
   End
   Begin MSForms.Label lblCardUsuario 
      Caption         =   ""
      Height          =   860
      Left            =   320
      TabIndex        =   5
      Top             =   1440
      Width           =   8360
   End
   Begin MSForms.Label lblUsuario 
      Caption         =   ""
      WordWrap        =   -1  'True
      Height          =   700
      Left            =   520
      TabIndex        =   6
      Top             =   1520
      Width           =   7960
   End
   Begin MSForms.CommandButton cmdCatalogo 
      Caption         =   "CATÁLOGO Y COMPRAS"
      Height          =   1100
      Left            =   320
      TabIndex        =   7
      Top             =   2480
      Width           =   4100
   End
   Begin MSForms.CommandButton cmdVenta 
      Caption         =   "PUNTO DE VENTA"
      Height          =   1100
      Left            =   4580
      TabIndex        =   8
      Top             =   2480
      Width           =   4100
   End
   Begin MSForms.CommandButton cmdInventario 
      Caption         =   "INVENTARIO Y STOCK"
      Height          =   1100
      Left            =   320
      TabIndex        =   9
      Top             =   3740
      Width           =   4100
   End
   Begin MSForms.CommandButton cmdRegistrarCliente 
      Caption         =   "REGISTRAR CLIENTE"
      Height          =   1100
      Left            =   4580
      TabIndex        =   10
      Top             =   3740
      Width           =   4100
   End
   Begin MSForms.CommandButton cmdRegistrarUsuario 
      Caption         =   "GESTIÓN DE USUARIOS"
      Height          =   1100
      Left            =   320
      TabIndex        =   11
      Top             =   5000
      Width           =   4100
   End
   Begin MSForms.CommandButton cmdConfiguracion 
      Caption         =   "CONFIGURACIÓN DE CORREO"
      Height          =   1100
      Left            =   4580
      TabIndex        =   12
      Top             =   5000
      Width           =   4100
   End
   Begin MSForms.Label lblNotaRol 
      Caption         =   ""
      WordWrap        =   -1  'True
      Height          =   300
      Left            =   320
      TabIndex        =   13
      Top             =   6260
      Width           =   8360
   End
   Begin MSForms.CommandButton cmdCerrarSesion 
      Caption         =   "Cambiar de usuario"
      Height          =   520
      Left            =   320
      TabIndex        =   14
      Top             =   6760
      Width           =   4100
   End
   Begin MSForms.CommandButton cmdSalir 
      Caption         =   "Volver al inicio"
      Cancel          =   -1  'True
      Height          =   520
      Left            =   4580
      TabIndex        =   15
      Top             =   6760
      Width           =   4100
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim texto As String, conn As Object, nBajo As Long

    modEstilo.AplicarTema Me
    modEstilo.CargarLogo imgLogo

    nBajo = 0
    If modSesion.PuedeVerInventario() Then
        On Error Resume Next
        Set conn = modConexion.AbrirConexion()
        If Not conn Is Nothing Then
            nBajo = modInventario.ContarProductosBajoMinimo(conn)
            modConexion.CerrarConexion conn
        End If
        On Error GoTo 0
    End If

    texto = "Sesion activa: " & modSesion.gUsuarioNombre & "   |   Rol: " & modSesion.gUsuarioRol
    Select Case modSesion.gUsuarioRol
        Case "Cliente"
            texto = texto & vbCrLf & "Explore el catalogo con fotos y compre; recibira su boleta en PDF."
        Case "Vendedor"
            texto = texto & vbCrLf & "Alertas de stock: " & nBajo & " productos bajo el minimo."
        Case "Admin"
            texto = texto & vbCrLf & "Acceso total. Alertas de stock: " & nBajo & " productos bajo el minimo."
    End Select

    On Error Resume Next
    lblUsuario.BackStyle = 0
    lblUsuario.ForeColor = modEstilo.C_SLATE800
    lblUsuario.Font.Bold = True
    lblUsuario.Font.Size = 10
    lblUsuario.Caption = texto

    lblNotaRol.BackStyle = 0
    lblNotaRol.ForeColor = modEstilo.C_SLATE500
    lblNotaRol.Font.Size = 8.5
    lblNotaRol.TextAlign = 2

    modEstilo.EstiloBoton cmdCatalogo, "P"
    modEstilo.EstiloBoton cmdVenta, "S"
    modEstilo.EstiloBoton cmdInventario, "S"
    modEstilo.EstiloBoton cmdRegistrarCliente, "S"
    modEstilo.EstiloBoton cmdRegistrarUsuario, "S"
    modEstilo.EstiloBoton cmdConfiguracion, "S"
    cmdCatalogo.Font.Size = 12: cmdVenta.Font.Size = 12: cmdInventario.Font.Size = 12
    cmdRegistrarCliente.Font.Size = 12: cmdRegistrarUsuario.Font.Size = 12: cmdConfiguracion.Font.Size = 12
    cmdCatalogo.Font.Name = modEstilo.FUENTE_TITULO: cmdVenta.Font.Name = modEstilo.FUENTE_TITULO
    cmdInventario.Font.Name = modEstilo.FUENTE_TITULO: cmdRegistrarCliente.Font.Name = modEstilo.FUENTE_TITULO
    cmdRegistrarUsuario.Font.Name = modEstilo.FUENTE_TITULO: cmdConfiguracion.Font.Name = modEstilo.FUENTE_TITULO
    On Error GoTo 0

    If StrComp(modSesion.gUsuarioRol, "Cliente", vbTextCompare) = 0 Then
        cmdCatalogo.Caption = "VER CATALOGO Y COMPRAR"
        lblNotaRol.Caption = "Su cuenta de cliente permite ver el catalogo y realizar compras."
    ElseIf StrComp(modSesion.gUsuarioRol, "Vendedor", vbTextCompare) = 0 Then
        lblNotaRol.Caption = "Su cuenta de vendedor permite vender, consultar inventario y registrar clientes."
    Else
        lblNotaRol.Caption = "Su cuenta de administrador tiene acceso a todas las funciones."
    End If

    cmdVenta.Visible = modSesion.PuedeVender()
    cmdInventario.Visible = modSesion.PuedeVerInventario()
    cmdRegistrarCliente.Visible = modSesion.PuedeRegistrarClientes()
    cmdRegistrarUsuario.Visible = modSesion.PuedeAdministrarUsuarios()
    cmdConfiguracion.Visible = modSesion.PuedeConfigurar()
    cmdCatalogo.Visible = modSesion.PuedeComprarEnCatalogo()
    Reacomodar
End Sub

' Coloca los botones visibles en una cuadricula de 2 columnas, sin huecos.
' (En codigo las medidas van en puntos: twips / 20)
Private Sub Reacomodar()
    Dim nombres As Variant, i As Long, k As Long, b As Object
    nombres = Array("cmdCatalogo", "cmdVenta", "cmdInventario", "cmdRegistrarCliente", "cmdRegistrarUsuario", "cmdConfiguracion")
    k = 0
    For i = LBound(nombres) To UBound(nombres)
        Set b = Me.Controls(CStr(nombres(i)))
        If b.Visible Then
            b.Left = IIf(k Mod 2 = 0, 16, 229)
            b.Top = 124 + (k \ 2) * 63
            k = k + 1
        End If
    Next i
    ' Un boton impar queda a todo el ancho
    If k Mod 2 = 1 Then
        For i = UBound(nombres) To LBound(nombres) Step -1
            Set b = Me.Controls(CStr(nombres(i)))
            If b.Visible Then
                b.Left = 16
                b.Width = 418
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub cmdCatalogo_Click()
    frmCatalogo.Show
End Sub

Private Sub cmdRegistrarCliente_Click()
    frmRegistroCliente.Show
End Sub

Private Sub cmdRegistrarUsuario_Click()
    frmUsuarios.Show
End Sub

Private Sub cmdConfiguracion_Click()
    frmConfiguracion.Show
End Sub

Private Sub cmdVenta_Click()
    frmVenta.Show
End Sub

Private Sub cmdInventario_Click()
    frmInventario.Show
End Sub

Private Sub cmdCerrarSesion_Click()
    modSesion.gReiniciar = True
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    modSesion.gReiniciar = False
    Unload Me
End Sub
