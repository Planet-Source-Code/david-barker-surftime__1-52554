Attribute VB_Name = "Start"
Public Sub StartProG()
Dim LnchOrNo As Long
Dim FndFl As String
Dim Frfl As Long
Dim FlCnts As String
Dim TxtSz As Long
Dim count As Byte
Dim lFcs  As Long
    WndHndl = Array()
    FrstLd = False
    PsWrd1 = False
    PsWrd2 = True
    ChGAdPss = False
    DestrState = True
    Hours = 0
    MnuClicked = False
    OvrRide = False
    NwTm = False
    MsWord = False
    MsExcel = False
    MsPublisher = False
    MsAccess = False
    MsPowerPoint = False
    FindAccess = False
    PwrPntAddTm = False
    ExcelAddTm = False
    AccessAddTm = False
    WrdAddTm = False
    StpPrG = False
    BrwsrSt = False
    TmOutHr = -1
    'Obtain the next available file.
    Frfl = FreeFile
    FlCnts = 0
    count = 0
    Globals.RdFlPssWrd = ""
    Globals.RdFlPssWrd2 = ""
    'If someone is trying to use SurfTime in Windows 3.1x
    #If Win16 Then
        MsgBox ("You cannot run SurfTime in Windows 3.1x!")
        End
    #End If
    'This function makes sure that should anyone trie to get this software working then it will not work until it is renamed
    If (Dir(App.Path & "\ST4301.exe") <> Empty) Then End
    'Make sure that Windows Explorer is exists before running the application
    If (Dir(NetScpPth) = Empty And Dir(NetScp6Pth) = Empty) Then frmSurfMonitor.mnuNetScape.Enabled = False
    If (App.PrevInstance = True) Then End
    If (Dir(ExplrPth) <> Empty Or Dir(ExplrPth2) <> Empty Or NetScpPth <> Empty) Then
        FndFl = Dir(App.Path & "\Security.pwl")
        If (FndFl <> Empty) Then
            Open (App.Path & "\Security.pwl") For Binary Access Read As #Frfl
            'Reset the string to the size of the length of the file.
                Do While Not EOF(1)
                'Read each character from the file into the array.
                    Get #Frfl, , FlCnts
                    'Now decrypt the file so that the password can be validated.
                    FlCnts = Crypt(FlCnts)
                    If (FlCnts = Chr(61)) Then
                        Do While FlCnts <> Chr(124) And FlCnts <> Chr(144)
                            Get #Frfl, , FlCnts
                            FlCnts = Crypt(FlCnts)
                            If (FlCnts > Chr(47) And FlCnts < Chr(58) _
                                Or FlCnts > Chr(64) And FlCnts < Chr(91) _
                                    Or FlCnts > Chr(96) And FlCnts < Chr(123) Or FlCnts = Chr(95) Xor FlCnts = Chr(32)) Then
                                        Select Case count <> -1
                                            Case count = 0
                                                Globals.RdFlPssWrd = Globals.RdFlPssWrd & FlCnts
                                            Case count = 1
                                                Globals.RdFlPssWrd2 = Globals.RdFlPssWrd2 & FlCnts
                                        End Select
                            End If
                            'Validate the read character to make sure that it is legitimate.
                            If (InStr("01234567890 abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_", FlCnts) = 0) Then
                                count = count + 1
                                Exit Do
                            End If
                        Loop
                    End If
                Loop
                Globals.RdFlPssWrd = Trim(Globals.RdFlPssWrd)
                Globals.RdFlPssWrd2 = Trim(Globals.RdFlPssWrd2)
                If (Globals.RdFlPssWrd = Empty Or Globals.RdFlPssWrd = Chr(32)) Or _
                    (Globals.RdFlPssWrd2 = Empty Or Globals.RdFlPssWrd2 = Chr(32)) Then
                    Close Frfl
                    Kill App.Path & Chr(92) & "Security.PWL"
                    If (NTPlatform = True) Then
                        If (Dir("C:\WINNT\ST4302.stb") <> Empty) Then Kill "C:\WINNT\ST4302.stb"
                    Else: If (Dir("C:\Windows\ST4302.stb") <> Empty) Then Kill "C:\Windows\ST4302.stb"
                    End If
                End If
            Close Frfl
                'This procedure is to write the application up into the system registry.
                'Find out which directives are for Windows NT and Windows 95
            #If Win32 Then
                LdHstAddrss
                'SocketsInitialize
                'STIP_AddressCl = STIP_Address
                'SendData (STIP_AddressCl)
                Disable_Ctrl_Alt_Del
                MicrosoftOffice
            #End If
        Else
            Msg$ = MsgBox("SurfTime Cannot run because its password file is missing! Please re-intstall its password file", vbOKCancel)
            End
        End If
    '*************************************************************************************
    'If Company wants to incorporate the ability to change the password at a later stage.
    'Load frmChngPwsWrd
    'frmChngPwsWrd.Visible = True
    '*************************************************************************************
        Load frmCorpLg
        frmCorpLg.Visible = True
    Else
        Msg$ = MsgBox("SurfTime Cannot run because Explorer is missing! Please re-intstall Explorer", vbOKCancel)
        End
    End If
End Sub

Private Sub LdHstAddrss()
Dim Frfl As Long
Dim Rdln As String

Frfl = FreeFile()
If (Dir(App.Path & Chr(92) & "Host Server.ini") <> Empty) Then
    Open (App.Path & Chr(92) & "Host Server.ini") For Input As #Frfl
        Line Input #Frfl, Rdln
        Rdln = Crypt(Rdln)
    Close Frfl
    HostServerIP = Rdln
End If
End Sub
