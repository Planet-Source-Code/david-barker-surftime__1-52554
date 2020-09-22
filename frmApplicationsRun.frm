VERSION 5.00
Begin VB.Form frmApplicationsRun 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Timer Control"
   ClientHeight    =   5235
   ClientLeft      =   8745
   ClientTop       =   3435
   ClientWidth     =   2715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   2715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Running Applications"
      Height          =   5175
      Left            =   840
      TabIndex        =   0
      Top             =   0
      Width           =   1815
      Begin VB.CommandButton btnClose 
         Caption         =   "Close"
         Height          =   375
         Left            =   480
         TabIndex        =   9
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Browser Time"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label lblInternetTm 
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Label lblPwrPntTmCp 
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label lblAcssTmCp 
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label lblXclTmCp 
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label lblWrdTmCp 
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Microsoft PowerPoint"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Microsoft Access"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label label3 
         Caption         =   "Microsof Excel"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Microsoft Word"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Image Image1 
      Height          =   7650
      Left            =   0
      Picture         =   "frmApplicationsRun.frx":0000
      Top             =   0
      Width           =   750
   End
End
Attribute VB_Name = "frmApplicationsRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnClose_Click()
Me.Hide
End Sub

Private Sub Form_Load()
Dim MsPos As POINTAPI
Dim CrntPos As Long
Dim lrtn As Long
'lrtn = SetWindowPos(Me.hWnd, -1, 0&, 0&, 0&, 0&, 3)
'CrntPos = GetCursorPos(MsPos)
'Me.Left = MsPos.x * Screen.TwipsPerPixelX
'Me.Top = MsPos.y * Screen.TwipsPerPixelY

End Sub

Private Sub Form_Terminate()
Cancel = 1
End Sub

