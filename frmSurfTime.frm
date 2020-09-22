VERSION 5.00
Object = "{33155A3D-0CE0-11D1-A6B4-444553540000}#1.0#0"; "SYSTRAY.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSurfMonitor 
   BorderStyle     =   0  'None
   ClientHeight    =   480
   ClientLeft      =   12015
   ClientTop       =   10020
   ClientWidth     =   3225
   ControlBox      =   0   'False
   Icon            =   "frmSurfTime.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   32
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   215
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin MSComctlLib.ImageList CorpIcon 
      Left            =   1080
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   64
      ImageHeight     =   64
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSurfTime.frx":1CCA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer ChGTmr 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2280
      Top             =   0
   End
   Begin SysTray.SystemTray SystemTray1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      SysTrayText     =   ""
      IconFile        =   0
   End
   Begin VB.Timer Detectr 
      Interval        =   1
      Left            =   1800
      Top             =   15
   End
   Begin MSComctlLib.ImageList AggrsrMode 
      Left            =   480
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSurfTime.frx":4D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSurfTime.frx":5280
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSurfTime.frx":6088
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSurfTime.frx":6DE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSurfTime.frx":7C98
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSurfTime.frx":ACEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSurfTime.frx":BA44
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSurfTime.frx":C0A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSurfTime.frx":D15C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSurfTime.frx":E368
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock SurfTimeSocket 
      Left            =   2760
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "SurfTime Settings"
      Visible         =   0   'False
      Begin VB.Menu mnuChgDflt 
         Caption         =   "Set Browser Time"
      End
      Begin VB.Menu mnuSp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOvrRide 
         Caption         =   "SurfTime Override"
      End
      Begin VB.Menu mnuReset 
         Caption         =   "Reset SurfTime"
      End
      Begin VB.Menu mnuSp9 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuChange 
         Caption         =   "Change Master Password"
      End
      Begin VB.Menu mnuChangeAd 
         Caption         =   "Change Admin Password"
      End
      Begin VB.Menu mnuSp5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLaunch 
         Caption         =   "Launch"
         Begin VB.Menu mnuIE 
            Caption         =   "Internet Explorer"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuSp10 
            Caption         =   "-"
         End
         Begin VB.Menu mnuNetScape 
            Caption         =   "Netscape"
         End
      End
      Begin VB.Menu mnuMsOffice 
         Caption         =   "Microsoft Office"
         Begin VB.Menu mnuWord 
            Caption         =   "Microsoft Word..."
         End
         Begin VB.Menu mnuSp6 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAccess 
            Caption         =   "Microsoft Access..."
         End
         Begin VB.Menu mnuSp4 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPowerPoint 
            Caption         =   "Microsoft PowerPoint..."
         End
         Begin VB.Menu mnuSp11 
            Caption         =   "-"
         End
         Begin VB.Menu mnuExcel 
            Caption         =   "Microsoft Excel..."
         End
         Begin VB.Menu mnuSp8 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPublisher 
            Caption         =   "Microsoft Publisher..."
         End
      End
      Begin VB.Menu mnuSp3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDisable 
         Caption         =   "Disable Task Manager"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEnTsk 
         Caption         =   "Enable Task Manager"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSp7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit SurfTime"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About SurfTime"
      End
      Begin VB.Menu mnuVersion 
         Caption         =   "Version No"
      End
      Begin VB.Menu mnu12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuServerManager 
         Caption         =   "Connection Manager"
      End
      Begin VB.Menu mnu13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Help and Assistance"
      End
      Begin VB.Menu mnu14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIP 
         Caption         =   ""
      End
   End
End
Attribute VB_Name = "frmSurfMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const vbLeftButton = 1     'Left button is pressed
Private Const vbRightButton = 2    'Right button is pressed
Private Const vbMiddleButton = 4   'Middle button is pressed
Public ClckPsX As Long
Public ClckPsY As Long
Public idx As Byte

Private Sub ChGTmr_Timer()
    
    SystemTray1.Action = sys_Modify
    Me.Icon = AggrsrMode.ListImages(idx).ExtractIcon
    SystemTray1.Icon = Val(Me.Icon)
    
    idx = idx + 1
    If (idx = 11) Then idx = 1

End Sub

Private Sub Form_Load()
Dim lFcs As Long
Dim lIsTrns As Long
On Error Resume Next
    frmSurfMonitor.Visible = False
    Me.Icon = AggrsrMode.ListImages(1).ExtractIcon
    idx = 1
    Me.Icon = CorpIcon.ListImages(1).ExtractIcon
    
    SystemTray1.Action = sys_Add
    SystemTray1.Icon = Val(Me.Icon)
    
    If (SystemTray1.IsIconLoaded = True) Then
        SystemTray1.SysTrayText = "SurfTime Copyright" & Chr(169) & Chr(32) & "2000"
        ChGTmr.Enabled = True
    Else
        End
    End If
    'Display the current version of SurfTime for resolving any technical difficuties.
    Me.mnuVersion.Caption = Me.mnuVersion.Caption & Chr(58) & Chr(32) & App.Major & Chr(46) & App.Minor & Chr(46) & App.Revision
    '********************************************************************************************************
    'I will have to consider a boolean to determine whether a connection was established or not. If not, then
    'prevent unlawful calls to the socket which could generate errors.
    '********************************************************************************************************
    If (Globals.HostServerIP <> Empty) Then
        With SurfTimeSocket
            .RemotePort = 600
            .Connect Globals.HostServerIP
        End With
        'This is where I send to the connected socket the ID/IP address of SurfTime.
        TmDelay (0.25)
        If (SurfTimeSocket.State = 7) Then
            'Send message.
            'MsgBox ("SurfTime was connected!")
            SurfTimeSocket.SendData (STIP_AddressCl & Chr(58) & ST_APP_ACTIVATE & Chr(58) & ST_END)
        End If
    End If
    
    STIP_AddressCl = SurfTimeSocket.LocalIP
    'Make sure that the IP address of the workstation is visible on the menubar.
    If (STIP_AddressCl <> Empty) Then
        Me.mnuIP.Caption = "IP Address" & Chr(58) & Chr(32) & Chr(32) & Chr(91) & STIP_AddressCl & Chr(93)
    Else: Me.mnuIP.Caption = "IP Address" & Chr(32) & Chr(32) & "No Workstation IP."
    End If
    
End Sub
'Exe Files - ExplrSurftTime, NetScpSurfTime
Private Sub Detectr_Timer()
Dim ActWn As Long
Dim WnStrG As String * 100
Dim WnTxt As String
Dim BufLen As Long
Dim lmyLong As Long
    'Detect whether Windows Explorer browser is running or not then close it down.
    If (DestrState = True) Then lmyLong = EnumWindows(AddressOf WndEnumProc, 0&)
    If (DestrState = False) Then
        lmyLong = EnumWindows(AddressOf WndEnumMsOffice, 0&)
    End If
'Check to make sure that its own exe file exists. If not then replace it from a windows copy
'Decide if Windows NT Platform or not then search through the appropriate path
    SelfCheck
End Sub

Private Sub Form_Resize()
    Me.WindowState = vbMinimized
End Sub

Private Sub Form_Terminate()
    Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SystemTray1.Action = sys_Delete
    If (frmSurfMonitor.SurfTimeSocket.State = 7) Then
        SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_APP_EXIT & Chr(58) & ST_END
    End If
End Sub

Private Sub mnuAbout_Click()
    Load frmAbout
    frmAbout.Visible = True
    Reset = False
End Sub

Private Sub mnuAccess_Click()
    Globals.MsAccess = True
    Globals.AccessAddTm = True
    Application = True
    Reset = False
    MnuClicked = True
    'Test to show thet the path exists, if not then disable this control.
    If (Me.mnuAccess.Checked = False) Then Me.mnuAccess.Checked = True
    If (Me.mnuPublisher.Checked = True) Then Me.mnuPublisher.Checked = False
    If (Me.mnuExcel.Checked = True) Then Me.mnuExcel.Checked = False
    If (Me.mnuPowerPoint.Checked = True) Then Me.mnuPowerPoint.Checked = False
    If (Me.mnuWord.Checked = True) Then Me.mnuWord.Checked = False
    
    AdChgDflt
    
End Sub

Private Sub mnuChange_Click()
'Routine to change the master password.
    PsWrd1 = True
    PsWrd2 = False
    Reset = False
    Load frmLogin
    frmLogin.Visible = True
End Sub

Private Sub mnuChangeAd_Click()
'Routine to change the administrators password.
    ChGAdPss = True
    PsWrd1 = False
    PsWrd2 = True
    Reset = False
    OvrRide = True
    Load frmLogin
    frmLogin.Visible = True
End Sub

Private Sub mnuChgDflt_Click()
'Routine to change the default hour to the maximum of 10 hours.
    PsWrd1 = False
    PsWrd2 = True
    Reset = False
    OvrRide = False
    BrwsrSt = True
    Load frmLogin
    frmLogin.Visible = True
End Sub

Private Sub mnuDisable_Click()
'Routine to make the Ctrl-Alt-Del function available
    If (NTPlatform = True) Then
    Else: App.TaskVisible = False
    End If
End Sub

Private Sub mnuEnTsk_Click()
'Routine to enable the Task Manager, and the Ctrl-Alt-Del function available
    EnblTskMn = True
    PsWrd1 = False
    PsWrd2 = True
    Reset = False
    OvrRide = False
    Load frmLogin
    frmLogin.Visible = True

End Sub

Private Sub mnuExcel_Click()
    Globals.MsExcel = True
    Globals.ExcelAddTm = True
    Application = True
    MnuClicked = True
    Reset = False
'Test to show thet the path exists, if not then disable this control.
    If (Me.mnuExcel.Checked = False) Then Me.mnuExcel.Checked = True
    If (Me.mnuAccess.Checked = True) Then Me.mnuAccess.Checked = False
    If (Me.mnuPublisher.Checked = True) Then Me.mnuPublisher.Checked = False
    If (Me.mnuPowerPoint.Checked = True) Then Me.mnuPowerPoint.Checked = False
    If (Me.mnuWord.Checked = True) Then Me.mnuWord.Checked = False
    
    AdChgDflt
    
End Sub

Private Sub mnuExit_Click()
    ExtPrG = True
'These variables have been set as if the Master password is going to be changed.
    OvrRide = True
    PsWrd2 = True
    Reset = False
    EnblTskMn = False
    ChGAdPss = True
    Load frmLogin: frmLogin.Visible = True
End Sub

Private Sub mnuHelp_Click()
    Load frmBrowser
    frmBrowser.Visible = True
    Reset = False
End Sub

Private Sub mnuIE_Click()
    Me.mnuIE.Checked = True
    Me.mnuNetScape.Checked = False
    IExplorer = True
    Netscape = False
    Reset = False
End Sub

Private Sub mnuNetScape_Click()
    Me.mnuIE.Checked = False
    Me.mnuNetScape.Checked = True
    Netscape = True
    IExplorer = False
    Reset = False
End Sub

Private Sub mnuOvrRide_Click()
'This routine changes the limit of SurfTime to 100 hours.
Dim ChGTm As Boolean
    OvrRide = True
    PsWrd2 = True
    PsWrd1 = False
    Reset = False
    Application = False
    frmRstTime.btnOvrRd.Enabled = True
    Load frmLogin: frmLogin.Visible = True
End Sub

Private Sub mnuPowerPoint_Click()
    Globals.MsPowerPoint = True
    Globals.PwrPntAddTm = True
    Application = True
    MnuClicked = True
    Reset = False
'Test to show that the path exists, if not then disable this control.
    If (Me.mnuPowerPoint.Checked = True) Then Me.mnuPowerPoint.Checked = False Else Me.mnuPowerPoint.Checked = True
    If (Me.mnuAccess.Checked = True) Then Me.mnuAccess.Checked = False
    If (Me.mnuExcel.Checked = True) Then Me.mnuExcel.Checked = False
    If (Me.mnuPublisher.Checked = True) Then Me.mnuPublisher.Checked = False
    If (Me.mnuWord.Checked = True) Then Me.mnuWord.Checked = False
    
    AdChgDflt

End Sub

Private Sub mnuPublisher_Click()
    Globals.MsPublisher = True
    Application = True
    MnuClicked = True
    Reset = False
'Test to show that the path exists, if not then disable this control.
    If (Me.mnuPublisher.Checked = False) Then Me.mnuPublisher.Checked = True
    If (Me.mnuAccess.Checked = True) Then Me.mnuAccess.Checked = False
    If (Me.mnuExcel.Checked = True) Then Me.mnuExcel.Checked = False
    If (Me.mnuPowerPoint.Checked = True) Then Me.mnuPowerPoint.Checked = False
    If (Me.mnuWord.Checked = True) Then Me.mnuWord.Checked = False
    
    AdChgDflt

End Sub

Private Sub mnuReset_Click()
'This routine resets SurfTime to the default of 1 hour.
    OvrRide = True
    PsWrd2 = True
    Rst = True
    Reset = True
    frmRstTime.btnOvrRd.Enabled = True
    Load frmLogin: frmLogin.Visible = True
End Sub

Private Sub mnuServerManager_Click()
Load frmConnectionManager
frmConnectionManager.Visible = True
Reset = False
End Sub

Private Sub mnuWord_Click()
    Globals.MsWord = True
    Globals.WrdAddTm = True
    Application = True
    MnuClicked = True
    Reset = False
    'Test to show thet the path exists, if not then disable this control.
    If (Me.mnuWord.Checked = False) Then Me.mnuWord.Checked = True
    If (Me.mnuAccess.Checked = True) Then Me.mnuAccess.Checked = False
    If (Me.mnuExcel.Checked = True) Then Me.mnuExcel.Checked = False
    If (Me.mnuPowerPoint.Checked = True) Then Me.mnuPowerPoint.Checked = False
    If (Me.mnuPublisher.Checked = True) Then Me.mnuPublisher.Checked = False
    
    AdChgDflt
    
End Sub

Private Sub SurfTimeSocket_Close()
If (frmSurfMonitor.SurfTimeSocket.State = 7) Then
    SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_APP_EXIT & Chr(58) & ST_END
End If
End Sub

Private Sub SurfTimeSocket_DataArrival(ByVal bytesTotal As Long)
Dim ReceivedBytes As String
Dim Pos As Long
Dim pos2 As Long
Dim IPInstr  As String
Dim IPCall As String
On Error GoTo ErrHandler
    SurfTimeSocket.GetData ReceivedBytes

    Pos = InStr(ReceivedBytes, Chr(58))
    
    If (Pos = 0) Then Exit Sub
    
    IPCall = Mid(ReceivedBytes, 1, Pos - 1)
    ReceivedBytes = Mid(ReceivedBytes, Pos + 1)
    pos2 = InStr(ReceivedBytes, Chr(58))
    
    If (pos2 > 0) Then
        IPInstr = Mid(ReceivedBytes, 1, pos2 - 1)
        ReceivedBytes = Mid(ReceivedBytes, pos2 + 1)
    Else: Exit Sub
    End If
    
    If (InStr(ReceivedBytes, Chr(58)) > 0) Then
        Pos = InStr(ReceivedBytes, Chr(58))
        IPMSG = Mid(ReceivedBytes, 1, Pos - 1)
    End If
    'If the IPCall is the the correct application, or if there is a network wide broadcast.
    If (StrComp(IPCall, STIP_AddressCl) = 0 Or IPInstr = CStr(ST_ALL) Or IPInstr = CStr(ST_BROADCAST_IP_RETURN)) Then
        Select Case IPInstr
            Case ST_TIMEOUT_01HRS
                Hours = 1
                    SurfTimeMonitor.BrowserLaunch
                        SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_TIMEOUT_01HRS & Chr(58) & ST_END
            Case ST_TIMEOUT_02HRS
                Hours = 2
                    SurfTimeMonitor.BrowserLaunch
                        SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_TIMEOUT_02HRS & Chr(58) & ST_END
            Case ST_TIMEOUT_03HRS
                Hours = 3
                    SurfTimeMonitor.BrowserLaunch
                        SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_TIMEOUT_03HRS & Chr(58) & ST_END
            Case ST_TIMEOUT_04HRS
                Hours = 4
                    SurfTimeMonitor.BrowserLaunch
                        SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_TIMEOUT_04HRS & Chr(58) & ST_END
            Case ST_TIMEOUT_05HRS
                Hours = 5
                    SurfTimeMonitor.BrowserLaunch
                        SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_TIMEOUT_05HRS & Chr(58) & ST_END
            Case ST_TIMEOUT_06HRS
                Hours = 6
                    SurfTimeMonitor.BrowserLaunch
                        SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_TIMEOUT_06HRS & Chr(58) & ST_END
            Case ST_TIMEOUT_07HRS
                Hours = 7
                    SurfTimeMonitor.BrowserLaunch
                        SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_TIMEOUT_07HRS & Chr(58) & ST_END
            Case ST_TIMEOUT_08HRS
                Hours = 8
                    SurfTimeMonitor.BrowserLaunch
                        SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_TIMEOUT_08HRS & Chr(58) & ST_END
            Case ST_TIMEOUT_09HRS
                Hours = 9
                    SurfTimeMonitor.BrowserLaunch
                        SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_TIMEOUT_09HRS & Chr(58) & ST_END
            Case ST_TIMEOUT_10HRS
                Hours = 10
                    SurfTimeMonitor.BrowserLaunch
                        SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_TIMEOUT_10HRS & Chr(58) & ST_END
            Case ST_TIMEOUT_UNLMTD
                Hours = Unlimited
                    SurfTimeMonitor.BrowserLaunch
                        SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_TIMEOUT_UNLMTD & Chr(58) & ST_END
            Case STLOAD_MS_WORD
                If (MsWordEst = True) Then
                    Hours = 1
                    MsWord = True
                    DestrState = False
                    Load frmTmr: frmTmr.Visible = True
                    lFcs = Shell(MsWrdPth, vbMaximizedFocus)
                    SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_APP_ACTIVATE & Chr(58) & ST_END
                Else
                    SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_NOT_INSTALLED & Chr(58) & ST_END
                End If
            Case STLOAD_MS_ACCESS
                If (MsAccessEst = True) Then
                    Hours = 1
                    MsAccess = True
                    DestrState = False
                    Load frmTmr
                    frmTmr.Visible = True
                    lFcs = Shell(MsAccsPth, vbMaximizedFocus)
                    SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_APP_ACTIVATE & Chr(58) & ST_END
                Else
                    SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_NOT_INSTALLED & Chr(58) & ST_END
                End If
            Case STLOAD_MS_EXCEL
                If (MsExcelEst = True) Then
                    Hours = 1
                    MsExcel = True
                    DestrState = False
                    Load frmTmr
                    frmTmr.Visible = True
                    lFcs = Shell(MsExlPth, vbMaximizedFocus)
                    SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_APP_ACTIVATE & Chr(58) & ST_END
                Else: SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_NOT_INSTALLED & Chr(58) & ST_END
                End If
            Case STLOAD_MS_POWERPOINT
                If (MsPowrPntEst = True) Then
                    Hours = 1
                    MsPowerPoint = True
                    DestrState = False
                    Load frmTmr
                    frmTmr.Visible = True
                    lFcs = Shell(MsPwrPntPth, vbMaximizedFocus)
                    SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_APP_ACTIVATE & Chr(58) & ST_END
                Else
                    SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_NOT_INSTALLED & Chr(58) & ST_END
                End If
            Case STLOAD_MS_PUBLISHER
                MsPublisher = True
                    SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_APP_ACTIVATE
            Case STLOAD_MS_IEXPLORER
                If (Dir(ExplrPth) <> Empty) Then
                    Hours = 1
                    DestrState = False
                    Load frmTmr: frmTmr.Visible = True
                    lFcs = Shell(ExplrPth, vbMaximizedFocus)
                    SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_APP_ACTIVATE
                Else
                    If (lFcs = 0) Then
                        DestrState = False
                        Load frmTmr
                        frmTmr.Visible = True
                        lFcs = Shell(ExplrPth2, vbMaximizedFocus)
                        SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_APP_ACTIVATE & Chr(58) & ST_END
                    Else
                        SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_NOT_INSTALLED & Chr(58) & ST_END
                    End If
                End If
            Case STLOAD_MS_NETSCAPE
                If (Dir(NetScpPth) <> Empty) Then
                    Hours = 1
                    lFcs = Shell(NetScpPth, vbMaximizedFocus)
                    SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_APP_ACTIVATE & Chr(58) & ST_END
                Else
                    If (lFcs = 0) Then
                        Hours = 1
                        DestrState = False
                        Load frmTmr
                        frmTmr.Visible = True
                        lFcs = Shell(NetScp6Pth, vbMaximizedFocus)
                        SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_APP_ACTIVATE & Chr(58) & ST_END
                    Else
                        SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_NOT_INSTALLED & Chr(58) & ST_END
                    End If
                End If
            Case ST_TIME_RESET
                Hours = Chr(48)
                    DestrState = True
                    'Reset all the MsOffice application variables to false.
                    Globals.MsAccess = False
                    Globals.MsWord = False
                    Globals.MsPowerPoint = False
                    Globals.MsExcel = False
                    '************************************
                        TmOutHr = -1
                            Unload frmTmr
                                frmProwler.SystemTray1.Action = sys_Delete
                                    Unload frmProwler
                                        SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_APP_EXIT & Chr(58) & ST_END
            Case ST_TIME_OVERRIDE
                    Hours = Unlimited
                        DestrState = False
                            frmProwler.Caption = "Unlimited surf time..."
                                Load frmTmr: frmTmr.Visible = True
                                    frmTmr.Timer1.Enabled = True
                                        SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_APP_ACTIVATE & Chr(58) & ST_END
            Case ST_TASkMANAGER_ACTIVE
                SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_APP_ACTIVATE & Chr(58) & ST_END
            Case ST_TASkMANAGER_INACTIV
                SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_APP_ACTIVATE & Chr(58) & ST_END
            Case ST_CHGPSSWRD_ADMIN
                ChGAdPss = True
                    CreateNewPssWrd (IPMSG)
                        SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_PASSWORD_UPDATED & Chr(58) & ST_END
            Case ST_CHGPSSWRD_SUPER
                ChGAdPss = False
                    CreateNewPssWrd (IPMSG)
                        SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_PASSWORD_UPDATED & Chr(58) & ST_END
            Case ST_APP_EXIT
                SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_APP_EXIT & Chr(58) & ST_END
                    frmSurfMonitor.SystemTray1.Action = sys_Delete
                    End
            'Return the message
            Case ST_SURFTIME_ACTIVE
                SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_APP_ACTIVATE & Chr(58) & ST_END
            Case ST_BROADCAST_IP_RETURN
                SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_IPADDRESS & Chr(58) & ST_END
            Case ST_ALL
                SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_IPADDRESS & Chr(58) & ST_END
        End Select
    End If
ErrHandler:
'Resume Next
End Sub

Private Sub SystemTray1_MouseDown(ByVal Button As Integer)
Dim MsPos As POINTAPI
Dim MsX As Long
Dim MsY As Long
Dim CrntPos As Long
Dim dl As Long
    'Place the popup menu where the cursor position is.
    CrntPos = GetCursorPos(MsPos)
    Me.Left = MsPos.x * Screen.TwipsPerPixelX
    Me.Top = MsPos.y * Screen.TwipsPerPixelY
    MnuPoslft = Me.Left
    MnuPostp = Me.Top
    'When the user presses on the icon in the system tray, show the application menu
    If (Button = vbRightButton) Then PopupMenu mnuSettings, 0&, 0&, mnuSettings
End Sub

Private Sub SelfCheck()
Static cntr As Byte
    cntr = cntr + 1
    'This routine checks the default files to make sure that they exist. If they have been removed then Surftime replaces the missing files.
    '****************************************************************************************
    'This makes sure that if the software file is missing, then it is automatically replaced.
    If (cntr = 1) Then
        If (NTPlatform = False) Then
            If (Dir("C:\Windows\ST4301.exe") = Empty) Then If (Dir(App.Path & Chr(92) & "SurfTime.exe") <> Empty) Then FileCopy App.Path & Chr(92) & "SurfTime.exe", "C:\Windows\ST4301.exe"
            If (Dir(App.Path & Chr(92) & "SurfTime.exe") = Empty) Then If (Dir("C:\Windows\ST4301.exe") <> Empty) Then FileCopy "C:\Windows\ST4301.exe", App.Path & Chr(92) & "SurfTime.exe"
        Else
            If (Dir("C:\WINNT\ST4301.exe") = Empty) Then FileCopy App.Path & Chr(92) & "SurfTime.exe", "C:\WINNT\ST4301.exe"
            If (Dir(App.Path & Chr(92) & "SurfTime.exe") = Empty) Then FileCopy "C:\WINNT\ST4301.exe", App.Path & Chr(92) & "SurfTime.exe"
        End If
    'This makes sure that if the security file is missing, then it is automatically replaced.
    If (NTPlatform = False) Then
            If (Dir("C:\Windows\ST4302.stb") = Empty) Then If (Dir(App.Path & Chr(92) & "Security.pwl") <> Empty) Then FileCopy App.Path & Chr(92) & "Security.pwl", "C:\Windows\ST4302.stb"
            If (Dir(App.Path & Chr(92) & "Security.pwl") = Empty) Then If (Dir("C:\Windows\ST4302.stb") <> Empty) Then FileCopy "C:\Windows\ST4302.stb", App.Path & Chr(92) & "Security.pwl"
        Else
            If (Dir("C:\WINNT\ST4302.stb") = Empty) Then If (Dir(App.Path & Chr(92) & "Security.pwl") <> Empty) Then FileCopy App.Path & Chr(92) & "Security.pwl", "C:\WINNT\ST4302.stb"
            If (Dir(App.Path & Chr(92) & "Security.pwl") = Empty) Then If (Dir("C:\WINNT\ST4302.stb") <> Empty) Then FileCopy "C:\WINNT\ST4302.stb", App.Path & Chr(92) & "Security.pwl"
        End If
    End If
    If (cntr = 20) Then cntr = 0
End Sub

Public Sub BrowserLaunch()
    DestrState = False
    Load frmTmr
    frmTmr.Visible = True
    frmTmr.Timer1.Enabled = True
End Sub

Private Sub AdChgDflt()
    PsWrd1 = False
    PsWrd2 = True
    Reset = False
    OvrRide = False
    Load frmLogin
    frmLogin.Visible = True
End Sub
