VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUsuarios 
   Caption         =   "Tienda de Fútbol - Gestión de Usuarios"
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
      Height          =   70
      Left            =   0
      TabIndex        =   1
      Top             =   1000
      Width           =   10400
   End
   Begin MSForms.Image imgLogo 
      Height          =   760
      Left            =   220
      TabIndex        =   2
      Top             =   110
      Width           =   760
   End
   Begin MSForms.Label lblTitulo 
      Caption         =   "GESTIÓN DE USUARIOS"
      Height          =   440
      Left            =   1150
      TabIndex        =   3
      Top             =   160
      Width           =   8000
   End
   Begin MSForms.Label lblSubtitulo 
      Caption         =   "Registrar, cambiar rol, restablecer clave, bloquear o eliminar cuentas"
      Height          =   280
      Left            =   1150
      TabIndex        =   4
      Top             =   640
      Width           =   8800
   End
   Begin MSForms.Label lblCardLista 
      Caption         =   ""
      Height          =   5200
      Left            =   320
      TabIndex        =   5
      Top             =   1300
      Width           =   6600
   End
   Begin MSForms.Label lblListaTitulo 
      Caption         =   "USUARIOS DEL SISTEMA"
      Height          =   300
      Left            =   480
      TabIndex        =   6
      Top             =   1380
      Width           =   6200
   End
   Begin MSForms.ListBox lstUsuarios 
      Height          =   4600
      Left            =   480
      TabIndex        =   7
      Top             =   1720
      Width           =   6280
   End
   Begin MSForms.Label lblCardAcciones 
      Caption         =   ""
      Height          =   5200
      Left            =   7100
      TabIndex        =   8
      Top             =   1300
      Width           =   2980
   End
   Begin MSForms.CommandButton cmdNuevo 
      Caption         =   "NUEVO USUARIO"
      Height          =   520
      Left            =   7260
      TabIndex        =   9
      Top             =   1460
      Width           =   2660
   End
   Begin MSForms.Label lblRol 
      Caption         =   "CAMBIAR ROL A"
      Height          =   240
      Left            =   7260
      TabIndex        =   10
      Top             =   2120
      Width           =   2660
   End
   Begin MSForms.ComboBox cboRol 
      Height          =   340
      Left            =   7260
      TabIndex        =   11
      Top             =   2370
      Width           =   2660
   End
   Begin MSForms.CommandButton cmdCambiarRol 
      Caption         =   "Aplicar rol"
      Height          =   440
      Left            =   7260
      TabIndex        =   12
      Top             =   2780
      Width           =   2660
   End
   Begin MSForms.Label lblClave 
      Caption         =   "NUEVA CLAVE (MIN. 8)"
      Height          =   240
      Left            =   7260
      TabIndex        =   13
      Top             =   3400
      Width           =   2660
   End
   Begin MSForms.TextBox txtClave 
      MaxLength       =   64
      PasswordChar    =   42
      Height          =   340
      Left            =   7260
      TabIndex        =   14
      Top             =   3650
      Width           =   2660
   End
   Begin MSForms.CommandButton cmdResetClave 
      Caption         =   "Restablecer clave"
      Height          =   440
      Left            =   7260
      TabIndex        =   15
      Top             =   4060
      Width           =   2660
   End
   Begin MSForms.CommandButton cmdBloquear 
      Caption         =   "Bloquear acceso"
      Height          =   440
      Left            =   7260
      TabIndex        =   16
      Top             =   4800
      Width           =   2660
   End
   Begin MSForms.CommandButton cmdEliminar 
      Caption         =   "Eliminar usuario"
      Height          =   440
      Left            =   7260
      TabIndex        =   17
      Top             =   5380
      Width           =   2660
   End
   Begin MSForms.Label lblMensaje 
      Caption         =   ""
      WordWrap        =   -1  'True
      Height          =   420
      Left            =   320
      TabIndex        =   18
      Top             =   6620
      Width           =   9760
   End
   Begin MSForms.CommandButton cmdCerrar 
      Caption         =   "Volver al menú"
      Cancel          =   -1  'True
      Height          =   500
      Left            =   320
      TabIndex        =   19
      Top             =   7150
      Width           =   3000
   End
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mIds() As Long
Private mN As Long

Private Sub UserForm_Initialize()
    modEstilo.AplicarTema Me
    modEstilo.CargarLogo imgLogo
    On Error Resume Next
    lblListaTitulo.BackStyle = 0: lblListaTitulo.Font.Bold = True: lblListaTitulo.Font.Size = 9: lblListaTitulo.ForeColor = modEstilo.C_SLATE600
    lblRol.BackStyle = 0: lblRol.Font.Bold = True: lblRol.Font.Size = 8: lblRol.ForeColor = modEstilo.C_SLATE600
    lblClave.BackStyle = 0: lblClave.Font.Bold = True: lblClave.Font.Size = 8: lblClave.ForeColor = modEstilo.C_SLATE600
    modEstilo.EstiloBoton cmdNuevo, "P"
    modEstilo.EstiloBoton cmdBloquear, "A"
    modEstilo.EstiloBoton cmdEliminar, "D"
    lstUsuarios.ColumnCount = 4
    lstUsuarios.ColumnWidths = "34 pt;110 pt;170 pt;80 pt"
    lstUsuarios.Font.Size = 9
    cboRol.Style = 2
    cboRol.AddItem "Admin"
    cboRol.AddItem "Vendedor"
    cboRol.AddItem "Cliente"
    On Error GoTo 0
    If Not modSesion.PuedeAdministrarUsuarios() Then
        MostrarMsg "Solo un Admin puede administrar usuarios.", True
        cmdNuevo.Enabled = False: cmdCambiarRol.Enabled = False: cmdResetClave.Enabled = False
        cmdBloquear.Enabled = False: cmdEliminar.Enabled = False
        Exit Sub
    End If
    CargarLista
End Sub

Private Sub CargarLista()
    On Error GoTo Fallo
    Dim conn As Object, rs As Object, i As Long, estado As String, errDesc As String
    lstUsuarios.Clear
    mN = 0
    Erase mIds
    Set conn = modConexion.AbrirConexion()
    Set rs = modUsuarios.ListarUsuarios(conn)
    Do While Not rs.EOF
        ReDim Preserve mIds(0 To mN)
        mIds(mN) = CLng(rs.Fields("id").Value)
        If CStr(rs.Fields("hash_contrasena").Value) = modUsuarios.CLAVE_BLOQUEADA Then
            estado = "Bloqueado"
        Else
            estado = "Activo"
        End If
        lstUsuarios.AddItem CStr(rs.Fields("id").Value)
        lstUsuarios.List(mN, 1) = CStr(rs.Fields("nombre").Value)
        lstUsuarios.List(mN, 2) = CStr(rs.Fields("correo").Value) & "  (" & CStr(rs.Fields("rol").Value) & ")"
        lstUsuarios.List(mN, 3) = estado
        mN = mN + 1
        rs.MoveNext
    Loop
    rs.Close
    modConexion.CerrarConexion conn
    Exit Sub
Fallo:
    errDesc = Err.Description
    modConexion.CerrarConexion conn
    MostrarMsg "No se pudo cargar la lista: " & errDesc, True
End Sub

Private Function IdSeleccionado() As Long
    If lstUsuarios.ListIndex >= 0 And lstUsuarios.ListIndex < mN Then IdSeleccionado = mIds(lstUsuarios.ListIndex)
End Function

Private Sub Sincronizar(ByVal conn As Object)
    On Error Resume Next
    modInventario.VolcarTablasMaestrasAHoja conn, modSesion.gUsuarioRol
    On Error GoTo 0
End Sub

Private Sub cmdNuevo_Click()
    frmRegistroUsuario.Show
    CargarLista
End Sub

Private Sub cmdCambiarRol_Click()
    Dim conn As Object, msg As String, ok As Boolean, errDesc As String
    On Error GoTo Fallo
    If Len(cboRol.Text) = 0 Then
        MostrarMsg "Elija el rol nuevo.", True
        Exit Sub
    End If
    Set conn = modConexion.AbrirConexion()
    ok = modUsuarios.CambiarRolUsuario(conn, IdSeleccionado(), modSesion.gUsuarioId, cboRol.Text, msg)
    If ok Then Sincronizar conn
    modConexion.CerrarConexion conn
    If ok Then
        CargarLista
        MostrarMsg "Rol actualizado.", False
    Else
        MostrarMsg msg, True
    End If
    Exit Sub
Fallo:
    errDesc = Err.Description
    modConexion.CerrarConexion conn
    MostrarMsg "No se pudo cambiar el rol: " & errDesc, True
End Sub

Private Sub cmdResetClave_Click()
    Dim conn As Object, msg As String, ok As Boolean, errDesc As String
    On Error GoTo Fallo
    Set conn = modConexion.AbrirConexion()
    ok = modUsuarios.RestablecerClave(conn, IdSeleccionado(), CStr(txtClave.Value), msg)
    modConexion.CerrarConexion conn
    If ok Then
        txtClave.Value = ""
        CargarLista
        MostrarMsg "Clave restablecida (si el usuario estaba bloqueado, queda activo).", False
    Else
        MostrarMsg msg, True
    End If
    Exit Sub
Fallo:
    errDesc = Err.Description
    modConexion.CerrarConexion conn
    MostrarMsg "No se pudo restablecer la clave: " & errDesc, True
End Sub

Private Sub cmdBloquear_Click()
    Dim conn As Object, msg As String, ok As Boolean, errDesc As String
    On Error GoTo Fallo
    If IdSeleccionado() <= 0 Then
        MostrarMsg "Seleccione un usuario de la lista.", True
        Exit Sub
    End If
    If MsgBox("El usuario no podra ingresar hasta que se le restablezca la clave. Continuar?", _
              vbYesNo + vbQuestion, "Bloquear acceso") = vbNo Then Exit Sub
    Set conn = modConexion.AbrirConexion()
    ok = modUsuarios.BloquearUsuario(conn, IdSeleccionado(), modSesion.gUsuarioId, msg)
    modConexion.CerrarConexion conn
    If ok Then
        CargarLista
        MostrarMsg "Acceso bloqueado.", False
    Else
        MostrarMsg msg, True
    End If
    Exit Sub
Fallo:
    errDesc = Err.Description
    modConexion.CerrarConexion conn
    MostrarMsg "No se pudo bloquear: " & errDesc, True
End Sub

Private Sub cmdEliminar_Click()
    Dim conn As Object, msg As String, ok As Boolean, errDesc As String, r As VbMsgBoxResult
    On Error GoTo Fallo
    If IdSeleccionado() <= 0 Then
        MostrarMsg "Seleccione un usuario de la lista.", True
        Exit Sub
    End If
    If MsgBox("Eliminar definitivamente al usuario seleccionado?", vbYesNo + vbExclamation, "Eliminar usuario") = vbNo Then Exit Sub
    Set conn = modConexion.AbrirConexion()
    ok = modUsuarios.EliminarUsuario(conn, IdSeleccionado(), modSesion.gUsuarioId, msg)
    If ok Then Sincronizar conn
    modConexion.CerrarConexion conn
    If ok Then
        CargarLista
        MostrarMsg "Usuario eliminado.", False
    Else
        MostrarMsg msg & " Use 'Bloquear acceso' para conservar el historial.", True
    End If
    Exit Sub
Fallo:
    errDesc = Err.Description
    modConexion.CerrarConexion conn
    MostrarMsg "No se pudo eliminar: " & errDesc, True
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub MostrarMsg(ByVal texto As String, ByVal esError As Boolean)
    modEstilo.Mensaje lblMensaje, texto, esError
End Sub
