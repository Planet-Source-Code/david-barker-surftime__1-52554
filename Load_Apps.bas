Attribute VB_Name = "Load_Apps"
Public Sub SoftwareLoad()
Dim PrvWnd As Boolean
    If (Application = True) Then
        If (frmSurfMonitor.mnuAccess.Checked = True) Then
            PrvWnd = HwndExst("OMain", "Microsoft Access")
            If (PrvWnd = True) Then Exit Sub
            If (MsAccessEst = True) Then
                lFcs = Shell(MsAccsPth, vbMaximizedFocus)
                TransmitData (STIP_AddressCl & Chr(58) & ST_APP_ACTIVATE & Chr(58) & ST_END)
            End If
        ElseIf (frmSurfMonitor.mnuExcel.Checked = True) Then
            PrvWnd = HwndExst("XLMAIN", "Microsoft Excel")
            If (PrvWnd = True) Then Exit Sub
            If (MsExcelEst = True) Then
                lFcs = Shell(MsExlPth, vbMaximizedFocus)
                TransmitData (STIP_AddressCl & Chr(58) & ST_APP_ACTIVATE & Chr(58) & ST_END)
            End If
        ElseIf (frmSurfMonitor.mnuPowerPoint.Checked = True) Then
            PrvWnd = HwndExst("PP9FrameClass", "Microsoft PowerPoint")
            If (PrvWnd = True) Then Exit Sub
            If (MsPowrPntEst = True) Then
                lFcs = Shell(MsPwrPntPth, vbMaximizedFocus)
                TransmitData (STIP_AddressCl & Chr(58) & ST_APP_ACTIVATE & Chr(58) & ST_END)
            End If
        ElseIf (frmSurfMonitor.mnuWord.Checked = True) Then
            PrvWnd = HwndExst("OpusApp", "Microsoft Word")
            If (PrvWnd = True) Then Exit Sub
            If (MsWordEst = True) Then
                lFcs = Shell(MsWrdPth, vbMaximizedFocus)
                TransmitData (STIP_AddressCl & Chr(58) & ST_APP_ACTIVATE & Chr(58) & ST_END)
            End If
        End If
        Application = False
        Exit Sub
    End If

    If (Application = False) Then
        If (frmSurfMonitor.mnuIE.Checked = True) Then
            If (Dir(ExplrPth) <> Empty) Then
                PrvWnd = HwndExst("IEFrame", "")
                If (PrvWnd = True) Then Exit Sub
                lFcs = Shell(ExplrPth, vbMaximizedFocus)
                TransmitData ((STIP_AddressCl & Chr(58) & ST_APP_ACTIVATE & Chr(58) & ST_END))
                IExplorer = True
            Else
                If (lFcs = 0) Then
                    PrvWnd = HwndExst("IEFrame", "")
                    If (PrvWnd = True) Then Exit Sub
                    lFcs = Shell(ExplrPth2, vbMaximizedFocus)
                    TransmitData ((STIP_AddressCl & Chr(58) & ST_APP_ACTIVATE & Chr(58) & ST_END))
                    IExplorer = True
                End If
            End If
        Else
            If (Dir(NetScpPth) <> Empty) Then
                'If (Netscape = True) Then Exit Sub
                lFcs = Shell(NetScpPth, vbMaximizedFocus)
                TransmitData ((STIP_AddressCl & Chr(58) & ST_APP_ACTIVATE & Chr(58) & ST_END))
                Netscape = True
            Else
                If (lFcs = 0) Then
                    'If (Netscape = True) Then Exit Sub
                    lFcs = Shell(NetScp6Pth, vbMaximizedFocus)
                    TransmitData ((STIP_AddressCl & Chr(58) & ST_APP_ACTIVATE & Chr(58) & ST_END))
                    Netscape = True
                End If
            End If
        End If
    End If
End Sub

Private Sub TransmitData(InstrC As String)
If (frmSurfMonitor.SurfTimeSocket.State = 7) Then
    frmSurfMonitor.SurfTimeSocket.SendData InstrC
End If
End Sub

Private Function HwndExst(Wnd As String, WndNm As String) As Boolean
Dim lParent As Long
Dim lChild As Long

lParent = FindWindow(Wnd, WndNm)
lChild = FindWindowEx(lParent, 0, Wnd, vbNullString)
If (lChild > 0 Or lParent > 0) Then HwndExst = True
End Function

Public Sub ResetUnload()
        frmTmr.Timer1.Enabled = False
        Unload frmTmr
        If (Reset) = False Then Load frmTimeOut
        frmTimeOut.Visible = True
        TmOutHr = -1
        OvrRide = False
        ScndsCnt = 0
        Hours = -1
        DestrState = True
        frmProwler.SystemTray1.Action = sys_Delete
        Unload frmProwler
        Unload FrmMsG
        Unload frmLogin
        Unload frmApplicationsRun
        'Reset the application variables back to false.
        MsWord = False
        MsAccess = False
        MsPowerPoint = False
        MsExcel = False
        MsPublisher = False
        WrdTm = 0
        XclTm = 0
        AcssTm = 0
        PwrPntTm = 0
        '************ Internet Applications ************
        IExplorer = False
        Netscape = False
        '***********************************************
        NwTm = False
        'Reset = False
        Rst = False
        '***********************************************
        If (frmSurfMonitor.SurfTimeSocket.State = 7) Then
            frmSurfMonitor.SurfTimeSocket.SendData (STIP_AddressCl & Chr(58) & ST_APP_ClOSE & Chr(58) & ST_END)
        End If
End Sub
