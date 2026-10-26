Attribute VB_Name = "modConexion"
' modConexion
' Modulo de conexion a la base de datos Access via ADODB (ACE OLEDB).
' La ruta de la base se lee de la hoja "Configuracion" (con fallback relativo
' a la carpeta del libro si el proyecto se movio de unidad/carpeta).
Option Explicit

' Constantes de ADO (evita referencia a Microsoft ActiveX Data Objects)
Private Const adCmdText As Long = 1
Private Const adParamInput As Long = 1
Private Const adVarWChar As Long = 202
Private Const adLongVarWChar As Long = 203
Private Const adInteger As Long = 3
Private Const adCurrency As Long = 6
Private Const adDate As Long = 7
Private Const adBoolean As Long = 11

' Nombre de la hoja de configuracion (unico lugar con el nombre)
Public Const HOJA_CONFIG As String = "Configuracion"
' Celda donde vive la ruta de la base (con nombre definido RutaBaseDatos)
Public Const NOMBRE_RANGO_RUTA As String = "RutaBaseDatos"

' Devuelve la ruta de la base leida desde la hoja Configuracion (o ruta relativa al libro).
Public Function ObtenerRutaBase() As String
    On Error GoTo Manejador
    Dim ws As Worksheet
    Dim ruta As String, rutaRelativa As String

    Set ws = ThisWorkbook.Worksheets(HOJA_CONFIG)
    On Error Resume Next
    ruta = Trim$(CStr(ws.Range(NOMBRE_RANGO_RUTA).Value))
    On Error GoTo Manejador

    If Len(ruta) = 0 Then
        Err.Raise vbObjectError + 513, "modConexion", _
            "La celda '" & NOMBRE_RANGO_RUTA & "' de la hoja '" & HOJA_CONFIG & _
            "' esta vacia. Indique ahi la ruta de la base de datos."
    End If
    If Dir(ruta) = "" Then
        ' Fallback relativo si el usuario descomprimio el proyecto en otra ruta
        If Len(ThisWorkbook.Path) > 0 Then
            rutaRelativa = ThisWorkbook.Path & Application.PathSeparator & ".." & _
                           Application.PathSeparator & "access" & Application.PathSeparator & "tienda_futbol.accdb"
            If Dir(rutaRelativa) <> "" Then
                ObtenerRutaBase = rutaRelativa
                Exit Function
            End If
        End If
        Err.Raise vbObjectError + 514, "modConexion", _
            "La base de datos no existe en: " & ruta
    End If
    ObtenerRutaBase = ruta
    Exit Function
Manejador:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' Abre y devuelve una conexion ADODB a la base (ACE OLEDB 12.0).
' El llamador SIEMPRE debe cerrarla (usar CerrarConexion).
Public Function AbrirConexion() As Object ' ADODB.Connection
    On Error GoTo Manejador
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ObtenerRutaBase() & ";"
    Set AbrirConexion = conn
    Exit Function
Manejador:
    Err.Raise Err.Number, Err.Source, "No se pudo abrir la base: " & Err.Description
End Function

' Cierra y libera una conexion (seguro si ya esta cerrada o es Nothing).
Public Sub CerrarConexion(ByRef conn As Object)
    On Error Resume Next
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close   ' 1 = adStateOpen
        Set conn = Nothing
    End If
    On Error GoTo 0
End Sub

' Crea un ADODB.Command ligado a una conexion, con texto y tipo.
Public Function NuevoComando(conn As Object, sql As String) As Object
    On Error GoTo Manejador
    Dim cmd As Object
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandText = sql
    cmd.CommandType = adCmdText
    Set NuevoComando = cmd
    Exit Function
Manejador:
    Err.Raise Err.Number, Err.Source, "No se pudo crear el comando: " & Err.Description
End Function

' Agrega un parametro de texto (VarWChar) a un comando.
Public Sub ParamTexto(cmd As Object, nombre As String, valor As Variant, Optional tamano As Long = 255)
    On Error GoTo Manejador
    cmd.Parameters.Append cmd.CreateParameter(nombre, adVarWChar, adParamInput, tamano, valor)
    Exit Sub
Manejador:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' Agrega un parametro entero (Long) a un comando.
Public Sub ParamEntero(cmd As Object, nombre As String, valor As Long)
    On Error GoTo Manejador
    cmd.Parameters.Append cmd.CreateParameter(nombre, adInteger, adParamInput, , valor)
    Exit Sub
Manejador:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' Agrega un parametro moneda (Currency) a un comando.
Public Sub ParamMoneda(cmd As Object, nombre As String, valor As Variant)
    On Error GoTo Manejador
    cmd.Parameters.Append cmd.CreateParameter(nombre, adCurrency, adParamInput, , valor)
    Exit Sub
Manejador:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' Agrega un parametro fecha (Date) a un comando.
Public Sub ParamFecha(cmd As Object, nombre As String, valor As Variant)
    On Error GoTo Manejador
    cmd.Parameters.Append cmd.CreateParameter(nombre, adDate, adParamInput, , valor)
    Exit Sub
Manejador:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' Ejecuta un SELECT y devuelve un recordset desconectado (el llamador lo cierra).
Public Function Consultar(conn As Object, sql As String) As Object ' ADODB.Recordset
    On Error GoTo Manejador
    Dim cmd As Object
    Dim rs As Object
    Set cmd = NuevoComando(conn, sql)
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3           ' 3 = adUseClient (permita desconectar)
    rs.Open cmd, , 3, 1             ' 3 = adOpenStatic, 1 = adLockReadOnly
    Set rs.ActiveConnection = Nothing   ' desconectar
    Set Consultar = rs
    Exit Function
Manejador:
    Err.Raise Err.Number, Err.Source, "Error en consulta: " & Err.Description
End Function

' Ejecuta un INSERT/UPDATE/DELETE con parametros y devuelve filas afectadas.
Public Function EjecutarAccion(conn As Object, cmd As Object) As Long
    On Error GoTo Manejador
    cmd.Execute EjecutarAccion
    Exit Function
Manejador:
    Err.Raise Err.Number, Err.Source, "Error al ejecutar accion: " & Err.Description
End Function
