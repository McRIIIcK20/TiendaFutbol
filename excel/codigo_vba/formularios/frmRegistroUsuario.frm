VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRegistroUsuario 
   Caption         =   "Tienda de Futbol - Administracion de Usuarios"
   ClientHeight    =   6000
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7200
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.Label lblBanner 
      Caption         =   ""
      Height          =   1000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7200
   End
   Begin MSForms.Label lblAccent 
      Caption         =   ""
      Height          =   80
      Left            =   0
      TabIndex        =   1
      Top             =   1000
      Width           =   7200
   End
   Begin MSForms.Label lblTitulo 
      Caption         =   "ADMINISTRACION DE USUARIOS Y ROLES (SHA-256 + SAL)"
      Height          =   340
      Left            =   280
      TabIndex        =   2
      Top             =   160
      Width           =   6640
   End
   Begin MSForms.Label lblSubtitulo 
      Caption         =   "Alta de cuentas Admin, Vendedor y Cliente con cifrado criptografico SHA-256"
      Height          =   280
      Left            =   280
      TabIndex        =   3
      Top             =   540
      Width           =   6640
   End
   Begin MSForms.Label lblCard 
      Caption         =   ""
      Height          =   3200
      Left            =   280
      TabIndex        =   4
      Top             =   1280
      Width           =   6640
   End
   Begin MSForms.Label lblNombre 
      Caption         =   "NOMBRE COMPLETO *"
      Height          =   280
      Left            =   560
      TabIndex        =   5
      Top             =   1540
      Width           =   1800
   End
   Begin MSForms.TextBox txtNombre 
      MaxLength       =   100
      Height          =   400
      Left            =   2440
      TabIndex        =   6
      Top             =   1480
      Width           =   4160
   End
   Begin MSForms.Label lblCorreo 
      Caption         =   "CORREO *"
      Height          =   280
      Left            =   560
      TabIndex        =   7
      Top             =   2120
      Width           =   1800
   End
   Begin MSForms.TextBox txtCorreo 
      MaxLength       =   255
      Height          =   400
      Left            =   2440
      TabIndex        =   8
      Top             =   2060
      Width           =   4160
   End
   Begin MSForms.Label lblClave 
      Caption         =   "CLAVE (MIN. 8) *"
      Height          =   280
      Left            =   560
      TabIndex        =   9
      Top             =   2700
      Width           =   1800
   End
   Begin MSForms.TextBox txtClave 
      MaxLength       =   64
      PasswordChar    =   42
      Height          =   400
      Left            =   2440
      TabIndex        =   10
      Top             =   2640
      Width           =   4160
   End
   Begin MSForms.Label lblClave2 
      Caption         =   "REPETIR CLAVE *"
      Height          =   280
      Left            =   560
      TabIndex        =   11
      Top             =   3280
      Width           =   1800
   End
   Begin MSForms.TextBox txtClave2 
      MaxLength       =   64
      PasswordChar    =   42
      Height          =   400
      Left            =   2440
      TabIndex        =   12
      Top             =   3220
      Width           =   4160
   End
   Begin MSForms.Label lblRol 
      Caption         =   "ROL EN SISTEMA *"
      Height          =   280
      Left            =   560
      TabIndex        =   13
      Top             =   3860
      Width           =   1800
   End
   Begin MSForms.ComboBox cboRol 
      Style           =   2  'fmStyleDropDownList
      Height          =   400
      Left            =   2440
      TabIndex        =   14
      Top             =   3800
      Width           =   4160
   End
   Begin MSForms.Label lblMensaje 
      Caption         =   ""
      ForeColor       =   &H000000C0&
      WordWrap        =   -1  'True
      Height          =   520
      Left            =   280
      TabIndex        =   15
      Top             =   4600
      Width           =   6640
   End
   Begin MSForms.CommandButton cmdGuardar 
      Caption         =   "REGISTRAR USUARIO"
      Default         =   -1  'True
      Height          =   540
      Left            =   280
      TabIndex        =   16
      Top             =   5240
      Width           =   4120
   End
   Begin MSForms.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Cancel          =   -1  'True
      Height          =   540
      Left            =   4600
      TabIndex        =   17
      Top             =   5240
      Width           =   2320
   End
End
Attribute VB_Name = "frmRegistroUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Me.BackColor = RGB(241, 245, 249)
    lblBanner.BackColor = RGB(15, 23, 42)
    lblAccent.BackColor = RGB(16, 185, 129)

    lblTitulo.BackStyle = 0: lblTitulo.ForeColor = RGB(255, 255, 255): lblTitulo.Font.Bold = True: lblTitulo.Font.Size = 13
    lblSubtitulo.BackStyle = 0: lblSubtitulo.ForeColor = RGB(148, 163, 184): lblSubtitulo.Font.Size = 9

    lblCard.BackColor = RGB(255, 255, 255)
    lblCard.BorderStyle = 1
    lblCard.BorderColor = RGB(203, 213, 225)

    lblNombre.BackStyle = 0: lblNombre.Font.Bold = True: lblNombre.ForeColor = RGB(51, 65, 85)
    lblCorreo.BackStyle = 0: lblCorreo.Font.Bold = True: lblCorreo.ForeColor = RGB(51, 65, 85)
    lblClave.BackStyle = 0: lblClave.Font.Bold = True: lblClave.ForeColor = RGB(51, 65, 85)
    lblClave2.BackStyle = 0: lblClave2.Font.Bold = True: lblClave2.ForeColor = RGB(51, 65, 85)
    lblRol.BackStyle = 0: lblRol.Font.Bold = True: lblRol.ForeColor = RGB(51, 65, 85)

    lblMensaje.BackStyle = 0: lblMensaje.Font.Bold = True: lblMensaje.Caption = ""

    cmdGuardar.BackColor = RGB(5, 150, 105): cmdGuardar.ForeColor = RGB(255, 255, 255): cmdGuardar.Font.Bold = True
    cmdCerrar.BackColor = RGB(71, 85, 105): cmdCerrar.ForeColor = RGB(255, 255, 255): cmdCerrar.Font.Bold = True

    cboRol.AddItem "Admin"
    cboRol.AddItem "Vendedor"
    cboRol.AddItem "Cliente"
    cboRol.ListIndex = -1
    modEstilo.AplicarTema Me
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    Dim conn As Object
    Dim nombre As String, correo As String, clave As String, clave2 As String
    Dim rol As String, msg As String
    Dim idNuevo As Long
    Dim errDesc As String

    lblMensaje.Caption = ""

    If Not modSesion.PuedeAdministrarUsuarios() Then
        MostrarError "Solo un Admin puede registrar usuarios."
        Exit Sub
    End If

    nombre = CStr(txtNombre.Value)
    correo = CStr(txtCorreo.Value)
    clave = CStr(txtClave.Value)
    clave2 = CStr(txtClave2.Value)
    rol = CStr(cboRol.Text)

    If clave <> clave2 Then
        MostrarError "Las claves no coinciden."
        Exit Sub
    End If

    On Error GoTo Fallo
    Set conn = modConexion.AbrirConexion()
    idNuevo = modUsuarios.RegistrarUsuario(conn, nombre, correo, clave, rol, msg)
    If idNuevo > 0 Then
        modInventario.VolcarTablasMaestrasAHoja conn, modSesion.gUsuarioRol
    End If
    modConexion.CerrarConexion conn

    If idNuevo > 0 Then
        MostrarOk "Usuario registrado correctamente (ID #" & idNuevo & ") y hoja Usuarios actualizada."
        LimpiarCampos
        txtNombre.SetFocus
    Else
        MostrarError msg
    End If
    Exit Sub
Fallo:
    errDesc = Err.Description
    modConexion.CerrarConexion conn
    MostrarError "No se pudo registrar: " & errDesc
End Sub

Private Sub LimpiarCampos()
    txtNombre.Value = ""
    txtCorreo.Value = ""
    txtClave.Value = ""
    txtClave2.Value = ""
    cboRol.ListIndex = -1
End Sub

Private Sub MostrarError(ByVal texto As String)
    lblMensaje.ForeColor = RGB(185, 28, 28)
    lblMensaje.Caption = texto
End Sub

Private Sub MostrarOk(ByVal texto As String)
    lblMensaje.ForeColor = RGB(5, 150, 105)
    lblMensaje.Caption = texto
End Sub
