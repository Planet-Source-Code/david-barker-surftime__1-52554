VERSION 5.00
Begin VB.Form frmNewPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SurfTime - Change Password"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   Icon            =   "frmNewPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   120
      TabIndex        =   7
      Top             =   2655
      Width           =   6015
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Change Password"
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   5535
      Begin VB.TextBox txtNewPsWrd 
         Height          =   255
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txtNewPsWrd2 
         Height          =   255
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   "Retype Password"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "New Password"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   -240
      Picture         =   "frmNewPassword.frx":1CCA
      Top             =   0
      Width           =   7020
   End
End
Attribute VB_Name = "frmNewPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOk_Click()
    If (StrComp(Trim(txtNewPsWrd.Text), Trim(txtNewPsWrd2.Text)) = 0) Then
        CreateNewPssWrd (Trim(txtNewPsWrd.Text))
        Unload Me
    Else
        Msg$ = MsgBox("The Password does not match, please retype your password.", vbOKOnly, "SurfTime - New Password")
    End If
End Sub

