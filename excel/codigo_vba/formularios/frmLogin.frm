VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin 
   Caption         =   "Tienda de Futbol - Acceso Ejecutivo"
   ClientHeight    =   4800
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6800
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.Label lblBanner 
      Caption         =   ""
      Height          =   1040
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6800
   End
   Begin MSForms.Label lblAccent 
      Caption         =   ""
      Height          =   80
      Left            =   0
      TabIndex        =   1
      Top             =   1040
      Width           =   6800
   End
   Begin MSForms.Image imgLogo 
      Height          =   800
      Left            =   240
      TabIndex        =   13
      Top             =   120
      Width           =   800
   End
   Begin MSForms.Label lblTitulo 
      Caption         =   "ACCESO AL SISTEMA"
      Height          =   360
      Left            =   1250
      TabIndex        =   2
      Top             =   200
      Width           =   5300
   End
   Begin MSForms.Label lblSubtitulo 
      Caption         =   "Tienda de Fútbol · Ingrese con su cuenta"
      Height          =   280
      Left            =   1250
      TabIndex        =   3
      Top             =   600
      Width           =   5300
   End
   Begin MSForms.Label lblCard 
      Caption         =   ""
      Height          =   2240
      Left            =   320
      TabIndex        =   4
      Top             =   1320
      Width           =   6160
   End
   Begin MSForms.Label lblCorreo 
      Caption         =   "CORREO ELECTRONICO"
      Height          =   260
      Left            =   640
      TabIndex        =   5
      Top             =   1540
      Width           =   5520
   End
   Begin MSForms.TextBox txtCorreo 
      MaxLength       =   255
      Height          =   420
      Left            =   640
      TabIndex        =   6
      Top             =   1820
      Width           =   5520
   End
   Begin MSForms.Label lblClave 
      Caption         =   "CONTRASENA DE ACCESO"
      Height          =   260
      Left            =   640
      TabIndex        =   7
      Top             =   2400
      Width           =   5520
   End
   Begin MSForms.TextBox txtClave 
      MaxLength       =   64
      PasswordChar    =   42
      Height          =   420
      Left            =   640
      TabIndex        =   8
      Top             =   2680
      Width           =   5520
   End
   Begin MSForms.Label lblDemoHint 
      Caption         =   "Demo: admin.demo@example.com (Admin123!) | vendedor.demo@example.com (Vendedor123!) | cliente.demo@example.com (Cliente123!)"
      WordWrap        =   -1  'True
      Height          =   340
      Left            =   640
      TabIndex        =   9
      Top             =   3180
      Width           =   5520
   End
   Begin MSForms.Label lblMensaje 
      Caption         =   ""
      ForeColor       =   &H000000C0&
      WordWrap        =   -1  'True
      Height          =   400
      Left            =   320
      TabIndex        =   10
      Top             =   3640
      Width           =   6160
   End
   Begin MSForms.CommandButton cmdIngresar 
      Caption         =   "INGRESAR AL SISTEMA"
      Default         =   -1  'True
      Height          =   540
      Left            =   320
      TabIndex        =   11
      Top             =   4080
      Width           =   3960
   End
   Begin MSForms.CommandButton cmdSalir 
      Caption         =   "Salir"
      Cancel          =   -1  'True
      Height          =   540
      Left            =   4480
      TabIndex        =   12
      Top             =   4080
      Width           =   2000
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_INTENTOS As Long = 3
Private mIntentos As Long

Private Sub UserForm_Initialize()
    Me.BackColor = RGB(241, 245, 249)

    lblBanner.BackColor = RGB(15, 23, 42)
    lblAccent.BackColor = RGB(16, 185, 129)

    lblTitulo.BackStyle = 0
    lblTitulo.ForeColor = RGB(255, 255, 255)
    lblTitulo.Font.Name = "Segoe UI"
    lblTitulo.Font.Size = 13
    lblTitulo.Font.Bold = True

    lblSubtitulo.BackStyle = 0
    lblSubtitulo.ForeColor = RGB(148, 163, 184)
    lblSubtitulo.Font.Name = "Segoe UI"
    lblSubtitulo.Font.Size = 9

    lblCard.BackColor = RGB(255, 255, 255)
    lblCard.BorderStyle = 1
    lblCard.BorderColor = RGB(203, 213, 225)

    lblCorreo.BackStyle = 0
    lblCorreo.ForeColor = RGB(51, 65, 85)
    lblCorreo.Font.Bold = True
    lblCorreo.Font.Size = 8.5

    lblClave.BackStyle = 0
    lblClave.ForeColor = RGB(51, 65, 85)
    lblClave.Font.Bold = True
    lblClave.Font.Size = 8.5

    txtCorreo.SpecialEffect = 0
    txtCorreo.BorderStyle = 1
    txtCorreo.BorderColor = RGB(148, 163, 184)
    txtCorreo.BackColor = RGB(248, 250, 252)
    txtCorreo.Font.Size = 10

    txtClave.SpecialEffect = 0
    txtClave.BorderStyle = 1
    txtClave.BorderColor = RGB(148, 163, 184)
    txtClave.BackColor = RGB(248, 250, 252)
    txtClave.Font.Size = 10

    lblDemoHint.BackStyle = 0
    lblDemoHint.ForeColor = RGB(100, 116, 139)
    lblDemoHint.Font.Size = 7.5

    lblMensaje.BackStyle = 0
    lblMensaje.Font.Bold = True
    lblMensaje.Caption = ""

    cmdIngresar.BackColor = RGB(5, 150, 105)
    cmdIngresar.ForeColor = RGB(255, 255, 255)
    cmdIngresar.Font.Bold = True
    cmdIngresar.Font.Size = 10

    cmdSalir.BackColor = RGB(100, 116, 139)
    cmdSalir.ForeColor = RGB(255, 255, 255)
    cmdSalir.Font.Bold = True
    cmdSalir.Font.Size = 9.5

    mIntentos = 0
    modEstilo.AplicarTema Me
    modEstilo.CargarLogo imgLogo
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdIngresar_Click()
    Dim conn As Object
    Dim correo As String, clave As String, rol As String
    Dim idU As Long
    Dim ok As Boolean
    Dim errDesc As String

    lblMensaje.Caption = ""
    correo = LCase$(Trim$(CStr(txtCorreo.Value)))
    clave = CStr(txtClave.Value)

    If Len(correo) = 0 Or Len(clave) = 0 Then
        lblMensaje.ForeColor = RGB(185, 28, 28)
        lblMensaje.Caption = "Ingrese su correo electronico y contrasena."
        Exit Sub
    End If

    On Error GoTo Fallo
    Set conn = modConexion.AbrirConexion()
    ok = modSeguridad.Autenticar(conn, correo, clave, rol, idU)

    If ok Then
        modSesion.IniciarSesion conn, idU, correo, rol
        modConexion.CerrarConexion conn
        txtClave.Value = ""
        Unload Me
        Exit Sub
    End If
    modConexion.CerrarConexion conn

    mIntentos = mIntentos + 1
    txtClave.Value = ""
    If mIntentos >= MAX_INTENTOS Then
        MsgBox "Demasiados intentos fallidos.", vbExclamation, "Acceso denegado"
        Unload Me
        Exit Sub
    End If
    lblMensaje.ForeColor = RGB(185, 28, 28)
    lblMensaje.Caption = "Correo o clave incorrectos. Intentos restantes: " & (MAX_INTENTOS - mIntentos)
    txtClave.SetFocus
    Exit Sub
Fallo:
    errDesc = Err.Description
    modConexion.CerrarConexion conn
    txtClave.Value = ""
    lblMensaje.ForeColor = RGB(185, 28, 28)
    lblMensaje.Caption = "No se pudo verificar el acceso: " & errDesc
End Sub
