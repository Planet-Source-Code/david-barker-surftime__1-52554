Attribute VB_Name = "Commit"
Public Function WriteToFile(NwUsrNm As String, CrntUser As String, NewPssWrd As String) As String
Dim Frfl, FndFl  As String
Dim successOrnot As Boolean
    Frfl = FreeFile
    successOrnot = True
    
    If (Dir(App.Path & "\Security.pwl") <> Empty) Then
        Kill App.Path & "\Security.pwl"
    End If
    
    Open (App.Path & "\Security.pwl") For Output As #Frfl
    
        'Encrypt all the text that is written to the file.
        NewPssWrd = Crypt("[PASSWORD]=" & NewPssWrd & Chr(166))
    
        'Check to see if a new user is writing to the file or a current one.
        If (NwUsrNm = Empty) Then
            CrntUser = Crypt("[USERNAME]=" & CrntUser & Chr(166))
            Print #Frfl, CrntUser
        Else
            NwUsrNm = Crypt("[USERNAME]=" & NwUsrNm & Chr(166))
            Print #Frfl, Trim(NwUsrNm)
        End If
    
        Print #Frfl, NewPssWrd
    Close #Frfl
    
    'Return whether the file writing was successful or not.
    WriteToFile = successOrnot
End Function
