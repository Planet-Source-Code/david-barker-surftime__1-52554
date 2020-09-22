Attribute VB_Name = "Close_Browser"

Public Sub CloseBrowser(hWnd As Long)
    ActWn = PostMessage(hWnd, WM_CLOSE, 0&, 0&)
End Sub


