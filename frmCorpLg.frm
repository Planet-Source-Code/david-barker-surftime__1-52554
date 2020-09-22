VERSION 5.00
Begin VB.Form frmCorpLg 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   Picture         =   "frmCorpLg.frx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4320
      Top             =   120
   End
End
Attribute VB_Name = "frmCorpLg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim rtn As Long
    rtn = SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    If (DestrState = True) Then
        Load frmSurfMonitor
        frmSurfMonitor.Visible = False
    End If
    Unload Me
End Sub
