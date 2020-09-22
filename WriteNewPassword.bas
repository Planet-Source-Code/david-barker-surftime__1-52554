Attribute VB_Name = "WriteNewPassword"
Public Sub CreateNewPssWrd(ByVal NewPssWrd3 As String)
Dim NewPssWrd2  As String
Dim NewPssWrd As String
Dim Frfl As Long
    Frfl = FreeFile
    If (Dir(App.Path & "\Security.pwl") <> Empty) Then Kill App.Path & "\Security.pwl"
    If (ChGAdPss = True) Then Globals.RdFlPssWrd = NewPssWrd3
    Open (App.Path & "\Security.pwl") For Append As #Frfl
    'Encrypt all the text that is written to the file.
        If (ChGAdPss = True) Then
            NewPssWrd = Trim(Crypt("[PASSWORD]=" & NewPssWrd3 & Chr(166)))
            NewPssWrd2 = Trim(Crypt("[PASSWORD]=" & Globals.RdFlPssWrd2 & Chr(166)))
        Else
            NewPssWrd = Trim(Crypt("[PASSWORD]=" & Globals.RdFlPssWrd & Chr(166)))
            NewPssWrd2 = Trim(Crypt("[PASSWORD]=" & NewPssWrd3 & Chr(166)))
        End If
        'Admin Password.
        Print #Frfl, NewPssWrd
        'Master Password.
        Print #Frfl, NewPssWrd2
    Close #Frfl
    
    If (frmSurfMonitor.SurfTimeSocket.State = 7) Then
        frmSurfMonitor.SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_PASSWORD_UPDATED & Chr(58) & ST_END
    End If
    
    If (ChGAdPss = False) Then RdFlPssWrd2 = NewPssWrd3
    ChGAdPss = False
End Sub
