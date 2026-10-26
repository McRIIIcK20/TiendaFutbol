Attribute VB_Name = "modInicio"
Option Explicit

' ============================================================================
' modInicio - Portada, arranque de la aplicacion y "modo tienda"
'   * Al abrir el libro solo se ve la hoja "Inicio" (logo + botones).
'     Todas las demas hojas quedan MUY ocultas (no aparecen en "Mostrar").
'   * Boton "ABRIR TIENDA" -> login -> menu segun rol (Admin / Vendedor / Cliente).
'   * Boton "SALIR" -> restaura Excel y cierra el libro.
'   * Macro de emergencia para el administrador de Excel: MostrarHojasAdmin
'     (Alt+F8). Pide credenciales de un usuario con rol Admin.
' ============================================================================

Public Const HOJA_INICIO As String = "Inicio"
Private Const xlSheetVisible As Long = -1
Private Const xlSheetVeryHidden As Long = 2

' ---------------------------------------------------------------------------
' Portada
' ---------------------------------------------------------------------------

' Deja el libro en estado "tienda": portada visible, resto de hojas ocultas.
Public Sub PrepararPortada()
    On Error GoTo Fallo
    Dim ws As Worksheet, h As Worksheet

    Set ws = ThisWorkbook.Worksheets(HOJA_INICIO)
    ConstruirPortada ws
    ws.Visible = xlSheetVisible
    ws.Activate
    For Each h In ThisWorkbook.Worksheets
        If h.Name <> HOJA_INICIO Then h.Visible = xlSheetVeryHidden
    Next h
    ModoTienda True
    Exit Sub
Fallo:
    AvisarError "No se pudo preparar la portada: " & Err.Description
End Sub

' Dibuja (o redibuja) el logo y los botones en la hoja Inicio.
Public Sub ConstruirPortada(ByVal ws As Worksheet)
    On Error GoTo Fallo
    Dim shp As Object, ruta As String
    Dim anchoImg As Single, altoImg As Single, izq As Single, tope As Single

    On Error Resume Next
    ws.Unprotect
    On Error GoTo Fallo

    ws.Cells.Clear
    ws.Cells.Interior.Color = modEstilo.C_NAVY
    ws.Cells.Font.Name = modEstilo.FUENTE_TEXTO
    ws.Cells.Font.Color = modEstilo.C_SLATE400
    ws.Columns("A:Z").ColumnWidth = 9

    For Each shp In ws.Shapes
        shp.Delete
    Next shp

    anchoImg = 780
    altoImg = anchoImg * 420 / 1200
    izq = 30
    tope = 24

    ruta = modEstilo.RutaAsset("portada.jpg")
    If Len(ruta) > 0 Then
        Set shp = ws.Shapes.AddPicture(ruta, False, True, izq, tope, anchoImg, altoImg)
        shp.Name = "imgPortada"
        shp.Placement = 3
    Else
        Set shp = ws.Shapes.AddTextbox(1, izq, tope, anchoImg, altoImg)
        shp.Name = "imgPortada"
        shp.TextFrame2.TextRange.Text = "TIENDA DE FUTBOL"
        shp.TextFrame2.TextRange.Font.Size = 40
        shp.TextFrame2.TextRange.Font.Name = modEstilo.FUENTE_TITULO
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modEstilo.C_BLANCO
        shp.Fill.ForeColor.RGB = modEstilo.C_NAVY
        shp.Line.Visible = False
    End If

    CrearBoton ws, "btnAbrirTienda", "ABRIR TIENDA", izq, tope + altoImg + 26, 500, 58, modEstilo.C_EMERALD_D, "modInicio.AbrirTienda"
    CrearBoton ws, "btnSalir", "SALIR", izq + 520, tope + altoImg + 26, 260, 58, modEstilo.C_ROJO, "modInicio.SalirAplicacion"

    Set shp = ws.Shapes.AddTextbox(1, izq, tope + altoImg + 100, anchoImg, 40)
    shp.Name = "txtPie"
    shp.Line.Visible = False
    shp.Fill.Visible = False
    shp.TextFrame2.TextRange.Text = "Haga clic en ABRIR TIENDA para ingresar con su correo y contrasena."
    shp.TextFrame2.TextRange.Font.Size = 11
    shp.TextFrame2.TextRange.Font.Name = modEstilo.FUENTE_TEXTO
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modEstilo.C_SLATE400
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = 2

    ws.Range("A1").Select
    ws.Protect DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True
    ws.EnableSelection = -4142      ' xlNoSelection
    Exit Sub
Fallo:
    AvisarError "No se pudo dibujar la portada: " & Err.Description
End Sub

' Con Excel visible muestra un aviso; si corre oculto (construir_xlsm.py) lanza el error al llamador.
Private Sub AvisarError(ByVal texto As String)
    If Application.Visible Then
        MsgBox texto, vbExclamation, "Tienda de Futbol"
    Else
        Err.Raise vbObjectError + 513, "modInicio", texto
    End If
End Sub

Private Sub CrearBoton(ByVal ws As Worksheet, ByVal nombre As String, ByVal texto As String, _
        ByVal izq As Single, ByVal tope As Single, ByVal ancho As Single, ByVal alto As Single, _
        ByVal color As Long, ByVal macro As String)
    Dim b As Object
    Set b = ws.Shapes.AddShape(5, izq, tope, ancho, alto)       ' 5 = rectangulo redondeado
    b.Name = nombre
    b.Fill.ForeColor.RGB = color
    b.Line.Visible = False
    b.Shadow.Visible = False
    b.OnAction = macro
    With b.TextFrame2.TextRange
        .Text = texto
        .Font.Name = modEstilo.FUENTE_TITULO
        .Font.Size = 20
        .Font.Bold = True
        .Font.Fill.ForeColor.RGB = modEstilo.C_BLANCO
        .ParagraphFormat.Alignment = 2
    End With
    b.TextFrame2.VerticalAnchor = 3     ' centrado
End Sub

' ---------------------------------------------------------------------------
' Modo tienda: oculta la interfaz de Excel mientras el libro esta activo
' ---------------------------------------------------------------------------
Public Sub ModoTienda(ByVal activar As Boolean)
    On Error Resume Next
    If activar Then
        If Not ActiveWorkbook Is ThisWorkbook Then Exit Sub
        ActiveWindow.WindowState = -4137        ' xlMaximized
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayHeadings = False
        ActiveWindow.DisplayWorkbookTabs = False
        ActiveWindow.DisplayHorizontalScrollBar = False
        ActiveWindow.DisplayVerticalScrollBar = False
        ActiveWindow.Zoom = 100
        Application.DisplayFormulaBar = False
        Application.DisplayStatusBar = False
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
        Application.Caption = "Tienda de Futbol"
    Else
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
        Application.DisplayFormulaBar = True
        Application.DisplayStatusBar = True
        Application.Caption = Empty
        If Not ActiveWindow Is Nothing Then
            ActiveWindow.DisplayGridlines = True
            ActiveWindow.DisplayHeadings = True
            ActiveWindow.DisplayWorkbookTabs = True
            ActiveWindow.DisplayHorizontalScrollBar = True
            ActiveWindow.DisplayVerticalScrollBar = True
        End If
    End If
    On Error GoTo 0
End Sub

Public Sub RestaurarEntorno()
    ModoTienda False
End Sub

' ---------------------------------------------------------------------------
' Botones de la portada
' ---------------------------------------------------------------------------
Public Sub AbrirTienda()
    Iniciar
    On Error Resume Next
    ThisWorkbook.Worksheets(HOJA_INICIO).Activate
    On Error GoTo 0
End Sub

' Cierra la aplicacion: restaura Excel y cierra este libro sin guardar cambios.
Public Sub SalirAplicacion()
    On Error Resume Next
    modSesion.CerrarSesion
    RestaurarEntorno
    Application.DisplayAlerts = False
    If Application.Workbooks.Count <= 1 Then
        ThisWorkbook.Saved = True
        Application.Quit
    Else
        ThisWorkbook.Close SaveChanges:=False
    End If
End Sub

' ---------------------------------------------------------------------------
' Login -> sincronizacion -> menu (se repite si el usuario pulsa "Cambiar de usuario")
' ---------------------------------------------------------------------------
Public Sub Iniciar()
    On Error GoTo Fallo
    Dim conn As Object
    Do
        modSesion.CerrarSesion
        frmLogin.Show
        If Not modSesion.gSesionIniciada Then Exit Do

        ' Sincronizar las hojas internas con la base Access (no se muestran al usuario)
        Application.ScreenUpdating = False
        On Error Resume Next
        Set conn = modConexion.AbrirConexion()
        If Not conn Is Nothing Then
            modInventario.VolcarCatalogoAHoja conn, modSesion.gUsuarioRol
            modVentas.VolcarVentasAHoja conn
            modInventario.VolcarTablasMaestrasAHoja conn, modSesion.gUsuarioRol
            modConexion.CerrarConexion conn
        End If
        On Error GoTo Fallo
        Application.ScreenUpdating = True

        modSesion.gReiniciar = False
        frmMenu.Show
    Loop While modSesion.gReiniciar
    modSesion.CerrarSesion
    Application.ScreenUpdating = True
    Exit Sub
Fallo:
    Application.ScreenUpdating = True
    MsgBox "Error al iniciar la aplicacion: " & Err.Description, vbCritical, "Tienda de Futbol"
End Sub

' Acceso directo al catalogo (si ya hay sesion) o al login.
Public Sub AbrirCatalogoVisual()
    On Error GoTo Fallo
    If Not modSesion.gSesionIniciada Then
        Iniciar
    Else
        frmCatalogo.Show
    End If
    Exit Sub
Fallo:
    MsgBox "No se pudo abrir el catalogo: " & Err.Description, vbExclamation, "Tienda de Futbol"
End Sub

' ---------------------------------------------------------------------------
' Emergencia: muestra todas las hojas (solo si se autentica un usuario Admin)
' ---------------------------------------------------------------------------
Public Sub MostrarHojasAdmin()
    Dim h As Worksheet
    modSesion.CerrarSesion
    frmLogin.Show
    If modSesion.gSesionIniciada And StrComp(modSesion.gUsuarioRol, "Admin", vbTextCompare) = 0 Then
        ModoTienda False
        On Error Resume Next
        ThisWorkbook.Worksheets(HOJA_INICIO).Unprotect
        For Each h In ThisWorkbook.Worksheets
            h.Visible = xlSheetVisible
        Next h
        On Error GoTo 0
        MsgBox "Hojas visibles. Para volver al modo tienda ejecute la macro PrepararPortada.", vbInformation, "Tienda de Futbol"
    Else
        MsgBox "Se requiere una cuenta con rol Admin.", vbExclamation, "Tienda de Futbol"
    End If
    modSesion.CerrarSesion
End Sub
