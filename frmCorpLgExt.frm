VERSION 5.00
Begin VB.Form frmCorpLgExt 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   Picture         =   "frmCorpLgExt.frx":0000
   ScaleHeight     =   1410
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   960
   End
End
Attribute VB_Name = "frmCorpLgExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim rtn As Long
    rtn = SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)
End Sub

Private Sub Timer1_Timer()
    Unload Me
End Sub
