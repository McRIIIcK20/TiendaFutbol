VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRegistroCliente 
   Caption         =   "Tienda de Futbol - Registro de Cliente"
   ClientHeight    =   6600
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
      Caption         =   "REGISTRO DE CLIENTE • FICHA COMERCIAL"
      Height          =   340
      Left            =   280
      TabIndex        =   2
      Top             =   160
      Width           =   6640
   End
   Begin MSForms.Label lblSubtitulo 
      Caption         =   "Validacion automatica de RUT Modulo 11, correo unico y sincronizacion con hoja Clientes"
      Height          =   280
      Left            =   280
      TabIndex        =   3
      Top             =   540
      Width           =   6640
   End
   Begin MSForms.Label lblCard 
      Caption         =   ""
      Height          =   3760
      Left            =   280
      TabIndex        =   4
      Top             =   1280
      Width           =   6640
   End
   Begin MSForms.Label lblNombre 
      Caption         =   "NOMBRE *"
      Height          =   280
      Left            =   560
      TabIndex        =   5
      Top             =   1540
      Width           =   1760
   End
   Begin MSForms.TextBox txtNombre 
      MaxLength       =   100
      Height          =   400
      Left            =   2400
      TabIndex        =   6
      Top             =   1480
      Width           =   4200
   End
   Begin MSForms.Label lblApellido 
      Caption         =   "APELLIDO *"
      Height          =   280
      Left            =   560
      TabIndex        =   7
      Top             =   2120
      Width           =   1760
   End
   Begin MSForms.TextBox txtApellido 
      MaxLength       =   100
      Height          =   400
      Left            =   2400
      TabIndex        =   8
      Top             =   2060
      Width           =   4200
   End
   Begin MSForms.Label lblRut 
      Caption         =   "RUT CHILENO *"
      Height          =   280
      Left            =   560
      TabIndex        =   9
      Top             =   2700
      Width           =   1760
   End
   Begin MSForms.TextBox txtRut 
      MaxLength       =   12
      Height          =   400
      Left            =   2400
      TabIndex        =   10
      Top             =   2640
      Width           =   4200
   End
   Begin MSForms.Label lblCorreo 
      Caption         =   "CORREO *"
      Height          =   280
      Left            =   560
      TabIndex        =   11
      Top             =   3280
      Width           =   1760
   End
   Begin MSForms.TextBox txtCorreo 
      MaxLength       =   255
      Height          =   400
      Left            =   2400
      TabIndex        =   12
      Top             =   3220
      Width           =   4200
   End
   Begin MSForms.Label lblTelefono 
      Caption         =   "TELEFONO (OPC.)"
      Height          =   280
      Left            =   560
      TabIndex        =   13
      Top             =   3860
      Width           =   1760
   End
   Begin MSForms.TextBox txtTelefono 
      MaxLength       =   50
      Height          =   400
      Left            =   2400
      TabIndex        =   14
      Top             =   3800
      Width           =   4200
   End
   Begin MSForms.Label lblDireccion 
      Caption         =   "DIRECCION (OPC.)"
      Height          =   280
      Left            =   560
      TabIndex        =   15
      Top             =   4440
      Width           =   1760
   End
   Begin MSForms.TextBox txtDireccion 
      MaxLength       =   255
      Height          =   400
      Left            =   2400
      TabIndex        =   16
      Top             =   4380
      Width           =   4200
   End
   Begin MSForms.Label lblMensaje 
      Caption         =   ""
      ForeColor       =   &H000000C0&
      WordWrap        =   -1  'True
      Height          =   540
      Left            =   280
      TabIndex        =   17
      Top             =   5160
      Width           =   6640
   End
   Begin MSForms.CommandButton cmdGuardar 
      Caption         =   "GUARDAR CLIENTE"
      Default         =   -1  'True
      Height          =   540
      Left            =   280
      TabIndex        =   18
      Top             =   5820
      Width           =   4120
   End
   Begin MSForms.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Cancel          =   -1  'True
      Height          =   540
      Left            =   4600
      TabIndex        =   19
      Top             =   5820
      Width           =   2320
   End
End
Attribute VB_Name = "frmRegistroCliente"
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
    lblApellido.BackStyle = 0: lblApellido.Font.Bold = True: lblApellido.ForeColor = RGB(51, 65, 85)
    lblRut.BackStyle = 0: lblRut.Font.Bold = True: lblRut.ForeColor = RGB(51, 65, 85)
    lblCorreo.BackStyle = 0: lblCorreo.Font.Bold = True: lblCorreo.ForeColor = RGB(51, 65, 85)
    lblTelefono.BackStyle = 0: lblTelefono.Font.Bold = True: lblTelefono.ForeColor = RGB(51, 65, 85)
    lblDireccion.BackStyle = 0: lblDireccion.Font.Bold = True: lblDireccion.ForeColor = RGB(51, 65, 85)

    lblMensaje.BackStyle = 0: lblMensaje.Font.Bold = True: lblMensaje.Caption = ""

    cmdGuardar.BackColor = RGB(5, 150, 105): cmdGuardar.ForeColor = RGB(255, 255, 255): cmdGuardar.Font.Bold = True
    cmdCerrar.BackColor = RGB(71, 85, 105): cmdCerrar.ForeColor = RGB(255, 255, 255): cmdCerrar.Font.Bold = True
    modEstilo.AplicarTema Me
End Sub

Private Sub txtRut_AfterUpdate()
    Dim r As String
    r = CStr(txtRut.Value)
    If modValidaciones.ValidarRut(r) Then txtRut.Value = modValidaciones.FormatearRut(r)
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    Dim conn As Object
    Dim nombre As String, apellido As String, rut As String, correo As String
    Dim telefono As String, direccion As String, msg As String
    Dim idNuevo As Long
    Dim errDesc As String

    lblMensaje.Caption = ""

    If Not modSesion.PuedeRegistrarClientes() Then
        MostrarError "Su rol no permite registrar clientes."
        Exit Sub
    End If

    nombre = CStr(txtNombre.Value)
    apellido = CStr(txtApellido.Value)
    rut = CStr(txtRut.Value)
    correo = CStr(txtCorreo.Value)
    telefono = CStr(txtTelefono.Value)
    direccion = CStr(txtDireccion.Value)

    On Error GoTo Fallo
    Set conn = modConexion.AbrirConexion()
    idNuevo = modClientes.RegistrarCliente(conn, nombre, apellido, correo, rut, telefono, direccion, msg)
    If idNuevo > 0 Then
        modInventario.VolcarTablasMaestrasAHoja conn, modSesion.gUsuarioRol
    End If
    modConexion.CerrarConexion conn

    If idNuevo > 0 Then
        MostrarOk "Cliente registrado correctamente (ID #" & idNuevo & ") y hoja Clientes actualizada."
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
    txtApellido.Value = ""
    txtRut.Value = ""
    txtCorreo.Value = ""
    txtTelefono.Value = ""
    txtDireccion.Value = ""
End Sub

Private Sub MostrarError(ByVal texto As String)
    lblMensaje.ForeColor = RGB(185, 28, 28)
    lblMensaje.Caption = texto
End Sub

Private Sub MostrarOk(ByVal texto As String)
    lblMensaje.ForeColor = RGB(5, 150, 105)
    lblMensaje.Caption = texto
End Sub
