Attribute VB_Name = "MakeFormTransparent2"
Public Sub MakeTransparent2(TransForm As Form)
Dim ErrorTest As Double
Dim Regn As Long
Dim TmpRegn As Long
Dim TmpControl As Control
'In case there's an error, ignore it
On Error Resume Next
    
    TransForm.ScaleMode = 3
    'makes everything invisible
    Regn = CreateRectRgn(0, 0, 0, 0)
    TmpRegn = CreateRectRgn(TransForm.Left, TransForm.Top, TransForm.Width, TransForm.Height)
    'Checks to make sure that the control has a width
    'or else you'll get some weird results
    ErrorTest = 0
    'Set TmpControl = Control.Width
    'ErrorTest = TmpControl
    'If ErrorTest <> 0 Or TypeOf TmpControl Is Line Then 'Combines the regions
        CombineRgn Regn, Regn, TmpRegn, RGN_XOR
    'End If
    'Make the regions
    SetWindowRgn TransForm.hWnd, Regn, True
End Sub
