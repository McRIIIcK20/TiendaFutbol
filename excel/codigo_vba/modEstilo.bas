Attribute VB_Name = "modEstilo"
Option Explicit

' ============================================================================
' modEstilo - Tema visual unico de todos los formularios
'   * Paleta: azul marino (#0F172A), esmeralda (#10B981), grises pizarra.
'   * Tipografias: Georgia (titulos, elegante) y Segoe UI (resto).
'   * AplicarTema(frm) recorre los controles del formulario y los estiliza
'     segun su tipo y su nombre (lblBanner, lblTitulo, cmdSalir, txtXxx, ...).
' ============================================================================

Public Const C_NAVY As Long = &H2A170F
Public Const C_SLATE800 As Long = &H3B291E
Public Const C_SLATE700 As Long = &H554133
Public Const C_SLATE600 As Long = &H695547
Public Const C_SLATE500 As Long = &H8B7464
Public Const C_SLATE400 As Long = &HB8A394
Public Const C_SLATE300 As Long = &HE1D5CB
Public Const C_SLATE200 As Long = &HF0E8E2
Public Const C_FONDO As Long = &HF9F5F1
Public Const C_EMERALD As Long = &H81B910
Public Const C_EMERALD_D As Long = &H699605
Public Const C_ROJO As Long = &H1C1CB9
Public Const C_NARANJA As Long = &H953B4
Public Const C_BLANCO As Long = &HFFFFFF
Public Const C_CAMPO As Long = &HFCFAF8
Public Const C_VERDE_CL As Long = &HE5FAD1

Public Const FUENTE_TITULO As String = "Georgia"
Public Const FUENTE_TEXTO As String = "Segoe UI"

' Devuelve la ruta de un archivo de la carpeta assets (logo, portada) o "" si no existe.
Public Function RutaAsset(ByVal nombre As String) As String
    On Error Resume Next
    Dim fso As Object, cand As String, base As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    base = ThisWorkbook.Path
    If Len(base) = 0 Then Exit Function
    cand = base & Application.PathSeparator & ".." & Application.PathSeparator & "assets" & Application.PathSeparator & nombre
    If fso.FileExists(cand) Then RutaAsset = fso.GetAbsolutePathName(cand): Exit Function
    cand = base & Application.PathSeparator & "assets" & Application.PathSeparator & nombre
    If fso.FileExists(cand) Then RutaAsset = fso.GetAbsolutePathName(cand): Exit Function
    On Error GoTo 0
End Function

' Carga el logo (jpg) en un control Image. Si no hay archivo, lo deja vacio.
Public Sub CargarLogo(ByVal img As Object)
    On Error Resume Next
    Dim ruta As String
    ruta = RutaAsset("logo_form.jpg")
    If Len(ruta) > 0 Then
        Set img.Picture = LoadPicture(ruta)
        img.PictureSizeMode = 3
        img.BorderStyle = 0
        img.BackStyle = 0
    End If
    On Error GoTo 0
End Sub

' Estiliza un boton segun su "tipo": P=principal (esmeralda), S=secundario (marino),
' N=neutro (gris), D=peligro (rojo), A=advertencia (naranja).
Public Sub EstiloBoton(ByVal b As Object, ByVal tipo As String)
    On Error Resume Next
    Select Case UCase$(tipo)
        Case "P": b.BackColor = C_EMERALD_D
        Case "S": b.BackColor = C_SLATE800
        Case "N": b.BackColor = C_SLATE600
        Case "D": b.BackColor = C_ROJO
        Case "A": b.BackColor = C_NARANJA
        Case Else: b.BackColor = C_SLATE800
    End Select
    b.ForeColor = C_BLANCO
    b.Font.Name = FUENTE_TEXTO
    b.Font.Bold = True
    b.TakeFocusOnClick = False
    On Error GoTo 0
End Sub

' Deduce el tipo de boton por palabras clave de su nombre.
Private Function TipoPorNombre(ByVal nombre As String) As String
    Dim n As String
    n = LCase$(nombre)
    If InStr(n, "salir") > 0 Or InStr(n, "eliminar") > 0 Or InStr(n, "anular") > 0 Then
        TipoPorNombre = "D"
    ElseIf InStr(n, "cerrar") > 0 Or InStr(n, "cancelar") > 0 Or InStr(n, "quitar") > 0 _
        Or InStr(n, "limpiar") > 0 Or InStr(n, "vaciar") > 0 Or InStr(n, "sesion") > 0 _
        Or InStr(n, "anterior") > 0 Or InStr(n, "siguiente") > 0 Then
        TipoPorNombre = "N"
    ElseIf InStr(n, "confirmar") > 0 Or InStr(n, "ingresar") > 0 Or InStr(n, "guardar") > 0 _
        Or InStr(n, "catalogo") > 0 Or InStr(n, "registrarmov") > 0 Or InStr(n, "agregar") > 0 _
        Or InStr(n, "comprar") > 0 Then
        TipoPorNombre = "P"
    Else
        TipoPorNombre = "S"
    End If
End Function

' Aplica el tema a todos los controles del formulario. Llamar al FINAL de UserForm_Initialize.
Public Sub AplicarTema(ByVal frm As Object)
    On Error Resume Next
    Dim ctl As Object, tn As String, nom As String

    frm.BackColor = C_FONDO

    For Each ctl In frm.Controls
        tn = TypeName(ctl)
        nom = ctl.Name
        Select Case tn
            Case "Label"
                ctl.Font.Name = FUENTE_TEXTO
                If nom = "lblBanner" Then
                    ctl.BackColor = C_NAVY
                ElseIf nom = "lblAccent" Then
                    ctl.BackColor = C_EMERALD
                ElseIf nom = "lblTitulo" Then
                    ctl.Font.Name = FUENTE_TITULO
                    ctl.Font.Size = 17
                    ctl.Font.Bold = True
                    ctl.ForeColor = C_BLANCO
                    ctl.BackStyle = 0
                ElseIf nom = "lblSubtitulo" Then
                    ctl.Font.Size = 9.5
                    ctl.ForeColor = C_SLATE400
                    ctl.BackStyle = 0
                ElseIf Left$(nom, 7) = "lblCard" Then
                    ctl.BackColor = C_BLANCO
                    ctl.BorderStyle = 1
                    ctl.BorderColor = C_SLATE300
                End If
            Case "CommandButton"
                EstiloBoton ctl, TipoPorNombre(nom)
                If ctl.Font.Size < 9.5 Then ctl.Font.Size = 9.5
            Case "TextBox", "ComboBox"
                ctl.Font.Name = FUENTE_TEXTO
                ctl.BorderStyle = 1
                ctl.BorderColor = C_SLATE400
                ctl.BackColor = C_CAMPO
                ctl.SpecialEffect = 0
                If ctl.Font.Size < 10 Then ctl.Font.Size = 10
            Case "ListBox"
                ctl.Font.Name = FUENTE_TEXTO
                ctl.BorderStyle = 1
                ctl.BorderColor = C_SLATE400
                ctl.BackColor = C_BLANCO
                ctl.SpecialEffect = 0
        End Select
    Next ctl
    On Error GoTo 0
End Sub

' Muestra un mensaje de estado (verde si ok, rojo si error) en un Label.
Public Sub Mensaje(ByVal lbl As Object, ByVal texto As String, ByVal esError As Boolean)
    lbl.BackStyle = 0
    lbl.Font.Bold = True
    If esError Then lbl.ForeColor = C_ROJO Else lbl.ForeColor = C_EMERALD_D
    lbl.Caption = texto
End Sub
