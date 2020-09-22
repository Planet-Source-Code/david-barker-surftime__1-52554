VERSION 5.00
Object = "{33155A3D-0CE0-11D1-A6B4-444553540000}#1.0#0"; "SYSTRAY.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProwler 
   ClientHeight    =   555
   ClientLeft      =   12300
   ClientTop       =   10725
   ClientWidth     =   1545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   555
   ScaleWidth      =   1545
   Visible         =   0   'False
   Begin VB.Timer ChGIcon 
      Interval        =   1000
      Left            =   1080
      Top             =   0
   End
   Begin SysTray.SystemTray SystemTray1 
      Left            =   600
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      SysTrayText     =   ""
      IconFile        =   0
   End
   Begin MSComctlLib.ImageList SafeMode 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProwler.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProwler.frx":0352
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProwler.frx":06A4
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmProwler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const sys_Add = 0       'Specifies that an icon is being add
Const sys_Modify = 1    'Specifies that an icon is being modified
Const sys_Delete = 2    'Specifies that an icon is being deleted
Public idx As Byte
Private Sub ChGIcon_Timer()
    SystemTray1.Action = sys_Modify
    Me.Icon = SafeMode.ListImages(idx).ExtractIcon
    SystemTray1.Icon = Val(Me.Icon)
    idx = idx + 1
    If (idx = 4) Then idx = 1
End Sub

Private Sub Form_Load()
Dim lIsTrns As Long
    MakeTransparent2 frmProwler
    Me.Icon = SafeMode.ListImages(1).ExtractIcon
    Me.Caption = "Running in Safe Mode."
    
    idx = 1
    SystemTray1.Action = sys_Add
    SystemTray1.Icon = Val(Me.Icon)
    
    If (SystemTray1.IsIconLoaded = True) Then
        SystemTray1.SysTrayText = "SurfTime Prowler Copyright" & Chr(169) & Chr(32) & "2000"
    End If

End Sub

Private Sub SystemTray1_MouseDown(ByVal Button As Integer)
Dim MsPos As POINTAPI
Dim CrntPos As Long
    CrntPos = GetCursorPos(MsPos)
    Me.Left = MsPos.x * Screen.TwipsPerPixelX
    Me.Top = MsPos.y * Screen.TwipsPerPixelY - MsPos.y * Screen.TwipsPerPixelY + Me.Top
    frmApplicationsRun.Visible = True
End Sub
