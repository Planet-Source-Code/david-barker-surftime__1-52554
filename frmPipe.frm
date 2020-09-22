VERSION 5.00
Begin VB.Form frmPipe 
   ClientHeight    =   495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1560
   Icon            =   "frmPipe.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   495
   ScaleWidth      =   1560
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Pipe 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmPipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub RecievedPipes()
Dim NamedPipe As Long
Dim ReadBytes As Long
Dim AvailBytes As Long
Dim LeftBytes As Long
Dim rtn As Long
'A routine for receiving pipes from over a network server
If (NTPlatform = True) Then
        
End If
End Sub
