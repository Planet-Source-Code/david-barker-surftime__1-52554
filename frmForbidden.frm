VERSION 5.00
Begin VB.Form frmForbidden 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   Picture         =   "frmForbidden.frx":0000
   ScaleHeight     =   2190
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5520
      Top             =   0
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   0
      Picture         =   "frmForbidden.frx":5E47
      Top             =   1440
      Width           =   6000
   End
End
Attribute VB_Name = "frmForbidden"
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
