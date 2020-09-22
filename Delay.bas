Attribute VB_Name = "Delay"
Public Function TmDelay(Tm As Variant)
Dim i As Long
Const Period = 1000000
    Tm = CDec(Tm * Period)
    For i = 0 To Tm Step 1
        DoEvents
        If (Reset = True) Then Exit For
    Next
    'If (Reset = True) Then Unload frmTmr
End Function

