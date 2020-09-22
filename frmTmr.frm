VERSION 5.00
Begin VB.Form frmTmr 
   ClientHeight    =   450
   ClientLeft      =   12300
   ClientTop       =   11175
   ClientWidth     =   1110
   ControlBox      =   0   'False
   Icon            =   "frmTmr.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   450
   ScaleWidth      =   1110
   WindowState     =   1  'Minimized
   Begin VB.Timer ExplrClssAggrsr 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   600
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   120
      Top             =   0
   End
End
Attribute VB_Name = "frmTmr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TmMin As Long
Public EndTime As Long
Public CountDwnHr
Public CountDwnMn
Public StartTime  As Long
Public Pass As Boolean
Public CntTr  As Byte
Public CrntHr As Long

Private Sub Form_Load()
Dim rtn As Long
Dim lIsTrns As Long
    Pass = False
    StartTime = Hour(Time)
    CountDwnHr = StartTime + Hours
    TmMin = Minute(Time)
    TmOutHr = (Hours * TmHour)
    EndTime = TmHour
    Load frmProwler
    frmProwler.Visible = True

    MakeTransparent2 frmTmr
    n$ = TmDelay(2)
    If (TmOutHr = -1 Or TmOutHr = 0) Then Exit Sub
    
    Load frmCorpLg
    Load frmApplicationsRun
    frmApplicationsRun.Visible = True
    frmApplicationsRun.lblWrdTmCp.Caption = "Time Out:" & Chr(32) & "00" & Chr(58) & "00" & Chr(32)
    frmApplicationsRun.lblXclTmCp.Caption = "Time Out:" & Chr(32) & "00" & Chr(58) & "00" & Chr(32)
    frmApplicationsRun.lblAcssTmCp.Caption = "Time Out:" & Chr(32) & "00" & Chr(58) & "00" & Chr(32)
    frmApplicationsRun.lblPwrPntTmCp.Caption = "Time Out:" & Chr(32) & "00" & Chr(58) & "00" & Chr(32)
    frmApplicationsRun.lblInternetTm.Caption = "Time Out:" & Chr(32) & "00" & Chr(58) & "00" & Chr(32)

    frmCorpLg.Visible = True
    Me.Caption = "SurfTime Running..."
End Sub

Private Sub Timer1_Timer()
Dim Strg As String
Dim NwStrG As String
Dim lEmn  As Long
Dim StrgLn As Long
Dim lStrg As Long
Dim lStrgRtn  As Long
Dim HrLf As Byte
Dim dblDigit  As String
Dim dblDigitmn As String
Dim idx As Integer
Dim NmApp As Byte
Dim TmOutMn As Long
Dim RnGApps As String
Static ScndsCnt As Integer

If (Pass = False) Then
    If (Hours = Unlimited) Then TmOutHr = (Hours * TmHour)
End If

If (Pass = False) Then TmOutHr = TmOutHr - 1
If (MsWord = True) Then WrdTm = App_Timer("Word", WrdTm)
If (MsExcel = True) Then XclTm = App_Timer("Excel", XclTm)
If (MsAccess = True) Then AcssTm = App_Timer("Access", AcssTm)
If (MsPowerPoint = True) Then PwrPntTm = App_Timer("PowerPoint", PwrPntTm)

If (frmProwler.Visible = True And TmOutHr = 59) Then frmProwler.Visible = False
'This line of code is for safety reasons. If for whatever reason the timer variable above has been decremented to a lower value than permissable then reset it from here.
If (TmOutHr < -1) Then
    TmOutHr = -1
    Pass = True
End If
'************************************************************************************************************************
HrLf = Fix(TmOutHr / 60)
Hours = HrLf
TmOutMn = TmOutHr Mod 60
'This makes sure that the user is warned that there is only limited time left.
If (TmOutHr = 8 And Hours = 0 And Pass = False) Then FrmMsG.lblMsG.Caption = "Please start to save any work before your time runs out!": Load FrmMsG: FrmMsG.Visible = True
If (TmOutHr = 5 And Hours = 0 And Pass = False) Then FrmMsG.lblMsG.Caption = "You have only 5 minutes of Web Surfing left!": Load FrmMsG: FrmMsG.Visible = True
If (TmOutHr = 3 And Hours = 0 And Pass = False) Then Beep: FrmMsG.lblMsG.Caption = "You have only 3 minutes of Web Surfing left!": Load FrmMsG: FrmMsG.Visible = True
If (TmOutHr = 1 And Hours = 0 And Pass = False) Then Beep: FrmMsG.lblMsG.Caption = "You have only 1 minute of Web Surfing left!": Load FrmMsG: FrmMsG.Visible = True
If (Pass = False) Then
    If (Hours < 10) Then dblDigit = Chr(48) & Hours Else dblDigit = Hours
    If (TmOutMn < 10) Then dblDigitmn = Chr(48) & TmOutMn Else dblDigitmn = TmOutMn
        frmTmr.Caption = "Time Out:" & Chr(32) & dblDigit & Chr(58) & dblDigitmn & Chr(32) & TmConst
        frmApplicationsRun.lblInternetTm.Caption = "Time Out:" & Chr(32) & dblDigit & Chr(58) & dblDigitmn
    'Determine the correct information to update/inform SurfTime Manager about what applications are currently running.
    If (MsWord = True) Then
        If (InStr(RnGApps, ST_MS_WORD) > 0) Then
            RnGApps = RnGApps & ST_MS_WORD
        End If
    End If
    
    If (MsAccess = True) Then
        If (InStr(RnGApps, ST_MS_ACCESS) > 0) Then
            RnGApps = RnGApps & ST_MS_ACCESS
        End If
    End If
    
    If (MsPowerPoint = True) Then
        If (InStr(RnGApps, ST_MS_POWERPOINT) > 0) Then
            RnGApps = RnGApps & ST_MS_POWERPOINT
        End If
    End If
    
    If (MsExcel = True) Then
        If (InStr(RnGApps, ST_MS_EXCEL) > 0) Then
            RnGApps = RnGApps & ST_MS_EXCEL
        End If
    End If
    
    If (MsPublisher = True) Then
        If (InStr(RnGApps, ST_MS_PUBLISHER) > 0) Then
            RnGApps = RnGApps & ST_MS_PUBLISHER
        End If
    End If
        
    If (Netscape = True) Then
        If (InStr(RnGApps, ST_MS_NETSCAPE) > 0) Then
            RnGApps = RnGApps & ST_MS_NETSCAPE
        End If
    End If
    
    If (IExplorer = True) Then
        If (InStr(RnGApps, ST_MS_IEXPLORER) > 0) Then
            RnGApps = RnGApps & ST_MS_IEXPLORER
        End If
    End If
    '********************************************************************************************************
    If (frmSurfMonitor.SurfTimeSocket.State = 7) Then
        frmSurfMonitor.SurfTimeSocket.SendData (STIP_AddressCl & Chr(58) & ST_TM_REM & Chr(58) & dblDigit & Chr(45) & dblDigitmn & Chr(58) & RnGApps & Chr(58) & ST_END)
    End If
End If

If (TmOutHr = 1) Then
    If (Pass = False) Then
        Timer1.Interval = 500   'Reset the timer so that one can flash the windows
        lEmn = EnumWindows(AddressOf WndEnumFnd, 0&)
        CntTr = 59
        Pass = True
    End If
End If

If (Pass = True) Then
    idx = 0
    ScndsCnt = ScndsCnt + 1
    Select Case ScndsCnt <> -1
        Case ScndsCnt Mod 2 = 1
            While idx < UBound(WndHndl, 1) + 1
                lParent = WndHndl(idx)(0)
                lWnFlsh = FlashWindow(lParent, True)
                idx = idx + 1
            Wend
        Case ScndsCnt Mod 2 = 0
            idx = 0
            NwStrG = "SurfTime Timeout in " & CntTr & Chr(32) & "seconds"
            While idx < UBound(WndHndl, 1) + 1
                lParent = WndHndl(idx)(0)
                lWnFlsh = FlashWindow(lParent, False)
                lStrgRtn = SetWindowText(lParent, NwStrG)
                idx = idx + 1
            Wend
            CntTr = CntTr - 1
    End Select
    'If the countdown is complete then close all browsers down and re-initialise the Prowler and the Task Manager disabler.
    If (CntTr = 0) Then
      ResetUnload
    End If
End If
End Sub

Private Function App_Timer(ByVal App As String, ByVal Tm As Long) As Integer
Dim HrLf As Byte
Dim dblDigit  As String
Dim dblDigitmn As String
Dim WrdHours As Long
Dim TmOutMn As Long

Tm = Tm - 1
HrLf = Fix(Tm / 60)
Hours = HrLf
TmOutMn = Tm Mod 60

Select Case App
    Case "Word"
        If (Tm = -1) Then MsWord = False: Exit Function
        If (Hours < 10) Then dblDigit = Chr(48) & Hours Else dblDigit = Hours
        If (TmOutMn < 10) Then dblDigitmn = Chr(48) & TmOutMn Else dblDigitmn = TmOutMn
        frmApplicationsRun.lblWrdTmCp.Caption = "Time Out:" & Chr(32) & dblDigit & Chr(58) & dblDigitmn & Chr(32)
    Case "Excel"
        If (Tm = -1) Then MsExcel = False: Exit Function
        If (Hours < 10) Then dblDigit = Chr(48) & Hours Else dblDigit = Hours
        If (TmOutMn < 10) Then dblDigitmn = Chr(48) & TmOutMn Else dblDigitmn = TmOutMn
        frmApplicationsRun.lblXclTmCp.Caption = "Time Out:" & Chr(32) & dblDigit & Chr(58) & dblDigitmn & Chr(32)
    Case "Access"
        If (Tm = -1) Then MsAccess = False: Exit Function
        If (Hours < 10) Then dblDigit = Chr(48) & Hours Else dblDigit = Hours
        If (TmOutMn < 10) Then dblDigitmn = Chr(48) & TmOutMn Else dblDigitmn = TmOutMn
        frmApplicationsRun.lblAcssTmCp.Caption = "Time Out:" & Chr(32) & dblDigit & Chr(58) & dblDigitmn & Chr(32)
    Case "PowerPoint"
        If (Tm = -1) Then MsPowerPoint = False: Exit Function
        If (Hours < 10) Then dblDigit = Chr(48) & Hours Else dblDigit = Hours
        If (TmOutMn < 10) Then dblDigitmn = Chr(48) & TmOutMn Else dblDigitmn = TmOutMn
        frmApplicationsRun.lblPwrPntTmCp.Caption = "Time Out:" & Chr(32) & dblDigit & Chr(58) & dblDigitmn & Chr(32)
End Select

App_Timer = Tm
End Function
