VERSION 5.00
Begin VB.Form frmTaskManager 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2700
   LinkTopic       =   "Form1"
   ScaleHeight     =   525
   ScaleWidth      =   2700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer TimeOut 
      Interval        =   1000
      Left            =   240
      Top             =   0
   End
End
Attribute VB_Name = "frmTaskManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Countdown As Long
Public PassOrFalse As Long

Private Sub Form_Load()
If (TmOutHr = -1) Then
    Countdown = 30
    PassOrFalse = True
    Load FrmMsG
    FrmMsG.Visible = True
Else
    Countdown = TmOutHr * TmHour
    PassOrFalse = False
End If

End Sub

Private Sub TimeOut_Timer()
Countdown = Countdown - 1
If (PassOrFalse = True) Then FrmMsG.lblMsG.Caption = "You have " & Countdown & Chr(32) & "seconds left before Task Manager is restarted."
    
If (Countdown = 0) Then
    EnblTskMn = False
    Unload Me
    Unload FrmMsG
End If
End Sub
