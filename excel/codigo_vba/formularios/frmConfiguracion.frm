VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConfiguracion 
   Caption         =   "Tienda de Fútbol - Configuración"
   ClientHeight    =   6300
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8000
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.Label lblBanner 
      Caption         =   ""
      Height          =   1000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8000
   End
   Begin MSForms.Label lblAccent 
      Caption         =   ""
      Height          =   70
      Left            =   0
      TabIndex        =   1
      Top             =   1000
      Width           =   8000
   End
   Begin MSForms.Image imgLogo 
      Height          =   760
      Left            =   220
      TabIndex        =   2
      Top             =   110
      Width           =   760
   End
   Begin MSForms.Label lblTitulo 
      Caption         =   "CONFIGURACIÓN DE CORREO"
      Height          =   440
      Left            =   1150
      TabIndex        =   3
      Top             =   160
      Width           =   6600
   End
   Begin MSForms.Label lblSubtitulo 
      Caption         =   "Envío automático de boletas por Gmail (solo Admin)"
      Height          =   280
      Left            =   1150
      TabIndex        =   4
      Top             =   640
      Width           =   6600
   End
   Begin MSForms.Label lblCardForm 
      Caption         =   ""
      Height          =   3700
      Left            =   320
      TabIndex        =   5
      Top             =   1300
      Width           =   7360
   End
   Begin MSForms.Label lblEmisor 
      Caption         =   "CUENTA GMAIL EMISORA"
      Height          =   240
      Left            =   520
      TabIndex        =   6
      Top             =   1400
      Width           =   6900
   End
   Begin MSForms.TextBox txtEmisor 
      MaxLength       =   255
      Height          =   360
      Left            =   520
      TabIndex        =   7
      Top             =   1650
      Width           =   6900
   End
   Begin MSForms.Label lblApp 
      Caption         =   "CONTRASEÑA DE APLICACIÓN (16 LETRAS, SE GUARDA EN EL LIBRO)"
      Height          =   240
      Left            =   520
      TabIndex        =   8
      Top             =   2150
      Width           =   6900
   End
   Begin MSForms.TextBox txtApp 
      MaxLength       =   64
      PasswordChar    =   42
      Height          =   360
      Left            =   520
      TabIndex        =   9
      Top             =   2400
      Width           =   6900
   End
   Begin MSForms.Label lblPrueba 
      Caption         =   "CORREO DE PRUEBA (RECIBE LAS BOLETAS DE CLIENTES @EXAMPLE.COM)"
      Height          =   240
      Left            =   520
      TabIndex        =   10
      Top             =   2900
      Width           =   6900
   End
   Begin MSForms.TextBox txtPrueba 
      MaxLength       =   255
      Height          =   360
      Left            =   520
      TabIndex        =   11
      Top             =   3150
      Width           =   6900
   End
   Begin MSForms.Label lblAyuda 
      Caption         =   "Cree la contraseña en myaccount.google.com/apppasswords (requiere verificación en 2 pasos). Péguela sin espacios. Si se deja vacía, la boleta solo se genera en PDF y no se envía por correo."
      WordWrap        =   -1  'True
      Height          =   1250
      Left            =   520
      TabIndex        =   12
      Top             =   3620
      Width           =   6900
   End
   Begin MSForms.Label lblMensaje 
      Caption         =   ""
      WordWrap        =   -1  'True
      Height          =   380
      Left            =   320
      TabIndex        =   13
      Top             =   5130
      Width           =   7360
   End
   Begin MSForms.CommandButton cmdGuardar 
      Caption         =   "GUARDAR CONFIGURACIÓN"
      Default         =   -1  'True
      Height          =   520
      Left            =   320
      TabIndex        =   14
      Top             =   5600
      Width           =   3600
   End
   Begin MSForms.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Cancel          =   -1  'True
      Height          =   520
      Left            =   4080
      TabIndex        =   15
      Top             =   5600
      Width           =   3600
   End
End
Attribute VB_Name = "frmConfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    modEstilo.AplicarTema Me
    modEstilo.CargarLogo imgLogo
    On Error Resume Next
    lblEmisor.BackStyle = 0: lblEmisor.Font.Bold = True: lblEmisor.Font.Size = 8: lblEmisor.ForeColor = modEstilo.C_SLATE600
    lblApp.BackStyle = 0: lblApp.Font.Bold = True: lblApp.Font.Size = 8: lblApp.ForeColor = modEstilo.C_SLATE600
    lblPrueba.BackStyle = 0: lblPrueba.Font.Bold = True: lblPrueba.Font.Size = 8: lblPrueba.ForeColor = modEstilo.C_SLATE600
    lblAyuda.BackStyle = 0: lblAyuda.Font.Size = 8.5: lblAyuda.ForeColor = modEstilo.C_SLATE500
    On Error GoTo 0
    If Not modSesion.PuedeConfigurar() Then
        modEstilo.Mensaje lblMensaje, "Solo un Admin puede cambiar la configuracion.", True
        cmdGuardar.Enabled = False
        Exit Sub
    End If
    txtEmisor.Value = modBoletas.LeerConfigTexto(modBoletas.NOMBRE_RANGO_GMAIL_USUARIO)
    txtApp.Value = modBoletas.LeerConfigTexto(modBoletas.NOMBRE_RANGO_GMAIL_CLAVE)
    txtPrueba.Value = modBoletas.LeerConfigTexto("CorreoPrueba")
End Sub

Private Sub GuardarCelda(ByVal nombreRango As String, ByVal valor As String)
    ThisWorkbook.Worksheets(modConexion.HOJA_CONFIG).Range(nombreRango).Value = valor
End Sub

Private Sub cmdGuardar_Click()
    On Error GoTo Fallo
    Dim emisor As String, clave As String, prueba As String
    emisor = Trim$(CStr(txtEmisor.Value))
    clave = Replace(Trim$(CStr(txtApp.Value)), " ", "")
    prueba = Trim$(CStr(txtPrueba.Value))
    If Len(emisor) > 0 And Not modValidaciones.ValidarCorreo(emisor) Then
        modEstilo.Mensaje lblMensaje, "El correo emisor no es valido.", True
        Exit Sub
    End If
    If Len(prueba) > 0 And Not modValidaciones.ValidarCorreo(prueba) Then
        modEstilo.Mensaje lblMensaje, "El correo de prueba no es valido.", True
        Exit Sub
    End If
    GuardarCelda modBoletas.NOMBRE_RANGO_GMAIL_USUARIO, emisor
    GuardarCelda modBoletas.NOMBRE_RANGO_GMAIL_CLAVE, clave
    GuardarCelda "CorreoPrueba", prueba
    txtApp.Value = clave
    ThisWorkbook.Save
    modEstilo.Mensaje lblMensaje, "Configuracion guardada en el libro.", False
    Exit Sub
Fallo:
    modEstilo.Mensaje lblMensaje, "No se pudo guardar: " & Err.Description, True
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub
