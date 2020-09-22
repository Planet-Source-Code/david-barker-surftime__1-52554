VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "About SurfTime"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5370
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmAbout.frx":1CCA
   ScaleHeight     =   8175
   ScaleWidth      =   5370
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_Resize()
    Me.Height = 8582
    Me.Width = 5492
End Sub
