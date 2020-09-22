VERSION 5.00
Begin VB.Form FrmMsG 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7680
   ClipControls    =   0   'False
   Icon            =   "FrmMsG.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   240
      TabIndex        =   3
      Top             =   3720
      Width           =   7215
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      Top             =   4065
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   0
      Top             =   480
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   -3480
      Picture         =   "FrmMsG.frx":1CCA
      Top             =   0
      Width           =   11250
   End
   Begin VB.Label lblMsG 
      Height          =   1455
      Left            =   1560
      TabIndex        =   0
      Top             =   1680
      Width           =   4455
   End
End
Attribute VB_Name = "FrmMsG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOk_Click()
    Unload Me
    If (EnblTskMn = True) Then EnblTskMn = False
End Sub

Private Sub Form_Load()
Dim lNwParent As Long
Dim rtn  As Long
'Set this form above all other windows
    rtn = SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub
