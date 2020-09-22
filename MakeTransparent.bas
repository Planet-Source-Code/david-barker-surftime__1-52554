Attribute VB_Name = "MakeFormTransparent"
Public Function MakeTransparent(ByVal hWnd As Long, Perc As Integer) As Long
Dim Msg As Long
Dim rtn As Long
Dim rtn2 As Long
On Error Resume Next
'This routine applies to Windows NT 2000 only.
    If Perc < 0 Or Perc > 255 Then
    'MakeTransparent = 1
    Else
        Msg = GetWindowLong(hWnd, GWL_EXSTYLE)
        Msg = Msg Or WS_EX_LAYERED
        rtn = SetWindowLong(hWnd, GWL_EXSTYLE, Msg)
        rtn2 = SetLayeredWindowAttributes(hWnd, 0, Perc, LWA_ALPHA)
    'MakeTransparent = 0
    End If

If Err Then
    'MakeTransparent = 2
End If
End Function
