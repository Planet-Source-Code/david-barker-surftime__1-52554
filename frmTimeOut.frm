VERSION 5.00
Begin VB.Form frmTimeOut 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   Picture         =   "frmTimeOut.frx":0000
   ScaleHeight     =   1395
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   720
   End
End
Attribute VB_Name = "frmTimeOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Timer1.Enabled = True
rtn = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)
End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub
