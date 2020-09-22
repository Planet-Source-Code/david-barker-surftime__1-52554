Attribute VB_Name = "RtnData"
Public Sub SendData(TransBytes As String)
On Error GoTo ErrHandler
    If (frmSurfMonitor.SurfTimeSocket.State = 7) Then
        frmSurfMonitor.SurfTimeSocket.SendData TransBytes
        Exit Sub
    End If
ErrHandler:
    MsgBox "Data Transmission Error" & Chr(58) & Chr(32) & Err.Description, vbOKOnly
End Sub
