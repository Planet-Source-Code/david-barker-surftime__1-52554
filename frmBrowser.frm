VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmBrowser 
   ClientHeight    =   5130
   ClientLeft      =   3060
   ClientTop       =   3345
   ClientWidth     =   10665
   Icon            =   "frmBrowser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   10665
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   9240
      TabIndex        =   2
      Top             =   4680
      Width           =   1215
   End
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   10440
      ExtentX         =   18415
      ExtentY         =   6588
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.PictureBox picAddress 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   0
      Picture         =   "frmBrowser.frx":1CCA
      ScaleHeight     =   675
      ScaleWidth      =   10665
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   10665
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnOk_Click()
Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.Show
    Form_Resize
    brwWebBrowser.Navigate (App.Path & "\Help and Support\Information_Page_Frame.htm")
End Sub

Private Sub Form_Resize()
    If (Me.ScaleWidth <> 0) Then brwWebBrowser.Width = Me.ScaleWidth - 270
    If (Me.ScaleHeight <> 0) Then
        brwWebBrowser.Height = Me.ScaleHeight - (picAddress.Top + picAddress.Height) - 700
        Me.btnOk.Top = Me.ScaleHeight - Me.btnOk.Height - 100
        Me.btnOk.Left = Me.ScaleWidth - Me.btnOk.Width - 200
    End If
End Sub



