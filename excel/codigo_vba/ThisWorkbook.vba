Option Explicit

' Codigo del objeto ThisWorkbook (lo inyecta construir_xlsm.py).
' Al abrir: solo se ve la portada. Al salir/cambiar de libro: se restaura la interfaz de Excel.

Private Sub Workbook_Open()
    modInicio.PrepararPortada
End Sub

Private Sub Workbook_Activate()
    modInicio.ModoTienda True
End Sub

Private Sub Workbook_Deactivate()
    modInicio.ModoTienda False
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    modInicio.RestaurarEntorno
End Sub
