VERSION 5.00
Begin VB.Form frmConnectionManager 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   Icon            =   "frmConnectionManager.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtip 
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Text            =   " "
      Top             =   1560
      Width           =   3855
   End
   Begin VB.CommandButton cmdTstConnection 
      Caption         =   "Test Connection"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6360
      TabIndex        =   3
      Top             =   4065
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   240
      TabIndex        =   4
      Top             =   3720
      Width           =   7215
   End
   Begin VB.Label lblCrntIP 
      Caption         =   "Current Host Server Surftime Manager IP Address"
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      Top             =   2040
      Width           =   3615
   End
   Begin VB.Label lblIPAddress 
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   $"frmConnectionManager.frx":000C
      Height          =   1215
      Left            =   600
      TabIndex        =   7
      Top             =   2520
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "SurfTime Manager IP Address"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the IP Address of SurfTime Manager below so that SurfTime Professional can connect to it."
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   6975
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   0
      Picture         =   "frmConnectionManager.frx":0196
      Top             =   0
      Width           =   7800
   End
End
Attribute VB_Name = "frmConnectionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ValidIP As Boolean
Private Sub Image2_Click()

End Sub

Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub btnOk_Click()
Dim Frfl As Long
Dim HostAddress As String
If (validateIP = True) Then
    HostAddress = Trim(txtip.Text)
    lblIPAddress.Caption = HostAddress
    HostAddress = Crypt(HostAddress)
    If Dir(App.Path & Chr(92) & "Host Server.ini") <> Empty Then Kill App.Path & Chr(92) & "Host Server.ini"
    Frfl = FreeFile()
    Open (App.Path & Chr(92) & "Host Server.ini") For Output As #Frfl
        Print #Frfl, HostAddress
        lblIPAddress.Caption = HostAddress
    Close #Frfl
End If
Unload Me

End Sub

Private Sub cmdTstConnection_Click()
On Error GoTo failed
If (validateIP = True) Then
    frmSurfMonitor.SurfTimeSocket.LocalPort = 600
    frmSurfMonitor.SurfTimeSocket.Connect Trim(txtip)
    If (frmSurfMonitor.SurfTimeSocket.State = 7) Then
        MsgBox "Connection to SurfTime Manager successful."
    Else: MsgBox "Connection to SurfTime Manager failed!"
    End If
Else: MsgBox "You must enter a valid IP."
End If
failed:
MsgBox "Connection to SurfTime Manager failed!"
Exit Sub
End Sub

Private Sub Form_Load()
If (HostServerIP <> Empty) Then
    lblIPAddress.Caption = HostServerIP
End If
End Sub

Private Function validateIP() As Boolean
Dim idx As Byte
Dim Pos As Long
Pos = 1
Do While InStr(Pos, txtip.Text, Chr(46)) > 0
    idx = idx + 1
    Pos = InStr(Pos, txtip.Text, Chr(46)) + 1
Loop

If (idx < 3) Then
    MsgBox "You have not entered the host IP Address correctly."
    txtip.Text = Empty
Else
    validateIP = True
End If

End Function
