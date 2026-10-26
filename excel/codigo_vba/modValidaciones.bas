Attribute VB_Name = "modValidaciones"
Option Explicit

' ============================================================================
' modValidaciones - Validaciones de datos de entrada (sin acceso a la base)
' FASE 2B-1. Todas las funciones son puras: reciben texto y devuelven un valor.
' ============================================================================

' Quita puntos, guiones y espacios de un RUT y lo pasa a mayusculas.
Public Function LimpiarRut(ByVal rut As String) As String
    Dim s As String
    s = UCase$(Trim$(rut))
    s = Replace(s, ".", "")
    s = Replace(s, "-", "")
    s = Replace(s, " ", "")
    LimpiarRut = s
End Function

' Calcula el digito verificador (modulo 11) de un cuerpo numerico de RUT.
' Devuelve "0".."9" o "K". Si el cuerpo no es numerico devuelve "".
Public Function CalcularDvRut(ByVal cuerpo As String) As String
    On Error GoTo Fallo
    Dim i As Long, suma As Long, factor As Long, r As Long

    factor = 2
    For i = Len(cuerpo) To 1 Step -1
        suma = suma + CLng(Mid$(cuerpo, i, 1)) * factor
        factor = factor + 1
        If factor > 7 Then factor = 2
    Next i

    r = 11 - (suma Mod 11)
    Select Case r
        Case 11
            CalcularDvRut = "0"
        Case 10
            CalcularDvRut = "K"
        Case Else
            CalcularDvRut = CStr(r)
    End Select
    Exit Function
Fallo:
    CalcularDvRut = ""
End Function

' Valida un RUT chileno. Acepta puntos, guion, espacios y K minuscula.
' Exige cuerpo de 7 u 8 digitos mas el digito verificador.
Public Function ValidarRut(ByVal rut As String) As Boolean
    Dim s As String, cuerpo As String, dv As String
    Dim i As Long, c As String

    s = LimpiarRut(rut)
    If Len(s) < 8 Or Len(s) > 9 Then Exit Function

    cuerpo = Left$(s, Len(s) - 1)
    dv = Right$(s, 1)

    For i = 1 To Len(cuerpo)
        c = Mid$(cuerpo, i, 1)
        If c < "0" Or c > "9" Then Exit Function
    Next i

    ValidarRut = (CalcularDvRut(cuerpo) = dv)
End Function

' Devuelve el RUT con formato 12.345.678-5, o "" si el RUT no es valido.
Public Function FormatearRut(ByVal rut As String) As String
    Dim s As String, cuerpo As String, dv As String, resultado As String
    Dim i As Long, n As Long

    If Not ValidarRut(rut) Then Exit Function

    s = LimpiarRut(rut)
    cuerpo = Left$(s, Len(s) - 1)
    dv = Right$(s, 1)

    For i = Len(cuerpo) To 1 Step -1
        resultado = Mid$(cuerpo, i, 1) & resultado
        n = n + 1
        If n Mod 3 = 0 And i > 1 Then resultado = "." & resultado
    Next i

    FormatearRut = resultado & "-" & dv
End Function

' Validacion basica de correo: una sola @, texto antes y despues,
' dominio con punto, sin espacios ni puntos pegados o en los extremos.
Public Function ValidarCorreo(ByVal correo As String) As Boolean
    Dim c As String, arroba As Long
    Dim parteLocal As String, dominio As String

    c = Trim$(correo)
    If Len(c) = 0 Or Len(c) > 255 Then Exit Function
    If InStr(c, " ") > 0 Then Exit Function
    If InStr(c, "..") > 0 Then Exit Function

    arroba = InStr(c, "@")
    If arroba < 2 Then Exit Function
    If InStr(arroba + 1, c, "@") > 0 Then Exit Function

    parteLocal = Left$(c, arroba - 1)
    dominio = Mid$(c, arroba + 1)

    If Left$(parteLocal, 1) = "." Or Right$(parteLocal, 1) = "." Then Exit Function
    If InStr(dominio, ".") = 0 Then Exit Function
    If Left$(dominio, 1) = "." Or Right$(dominio, 1) = "." Then Exit Function

    ValidarCorreo = True
End Function

' Clave valida: al menos 8 caracteres.
Public Function ValidarClave(ByVal clave As String) As Boolean
    ValidarClave = (Len(clave) >= 8)
End Function

' Campo obligatorio: no vacio ni solo espacios.
Public Function ValidarCampoObligatorio(ByVal valor As String) As Boolean
    ValidarCampoObligatorio = (Len(Trim$(valor)) > 0)
End Function
