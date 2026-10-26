' modTestConn
Option Explicit
Public Function TestConnectionState(conn As Object) As Long
    On Error GoTo ErrHandler
    If conn Is Nothing Then
        TestConnectionState = -1
        Exit Function
    End If
    TestConnectionState = conn.State
    Exit Function
ErrHandler:
    TestConnectionState = -2
End Function
Attribute VB_Name = "modTestConn"