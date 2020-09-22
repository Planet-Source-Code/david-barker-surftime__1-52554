Attribute VB_Name = "WndChild"
Public Function WndEnumChildProc(ByVal hWnd As Long, ByVal lParam As Long) As Long   'ByVal hwnd As Long
    Dim bRet As Long
    Dim myStr As String * 50
    Dim ltxtLn As Long
    Dim tmyStr  As String
    ltxtLn = GetWindowTextLength(hWnd)
    tmyStr = GetWindowText(hWnd, myStr, ltxtLn)
    bRet = GetClassName(hWnd, myStr, 50)
    'If (InStr(myStr, "MsoCommandBar") = 0) Then CloseBrowser (hwnd)
    WndEnumChildProc = 1
End Function

