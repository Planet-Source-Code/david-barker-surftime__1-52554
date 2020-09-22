VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6300
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   120
      TabIndex        =   5
      Top             =   2055
      Width           =   6015
   End
   Begin VB.TextBox txtPassword 
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1440
      Width           =   3375
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   3930
      TabIndex        =   1
      Top             =   2295
      Width           =   1095
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5115
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   750
      Left            =   -720
      Picture         =   "frmLogin.frx":1CCA
      Top             =   0
      Width           =   7020
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   270
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   840
   End
   Begin VB.Label lblCaps 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
Private Sub btnCancel_Click()
Dim NwVar As Integer
    OvrRide = False
    PsWrd2 = False
    Unload Me
End Sub

Private Sub btnOk_Click()
Dim LoadOrNot As Boolean
'Confirm the user nama and password.
    LoadOrNot = chkpsswrd
    If (LoadOrNot = True) Then
'    Msg$ = MsgBox("Would you like to Change your password?", vbYesNo, "SurfTime - Change Password")
'    If (Msg$ = "7") Then
        'Load frmRstTime
        'frmRstTime.Visible = True
        Unload Me
'    Else
'        Load frmChngPwsWrd
'        frmChngPwsWrd.Visible = True
'        Me.Visible = False
'    End If
    End If
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
Static CpsOnOff As Boolean
If (KeyCode = 20 And CpsOnOff = False) Then lblCaps.Caption = "CAPS": CpsOnOff = True: Exit Sub
If (KeyCode = 20 And CpsOnOff = True) Then lblCaps.Caption = Empty: CpsOnOff = False
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
Dim RtnVl As Boolean
'If the user presses the Enter key then check the password.
If (KeyAscii = 13) Then btnOk_Click
'If RtnVl = True Then
    'Msg$ = MsgBox("Would you like to Change your password?", vbYesNo, "SurfTime- Change Password")
    'If (Msg$ = "7") Then
'    End If
'End If
End Sub

Private Function chkpsswrd() As Boolean
Dim YeaOrNea As Boolean
If (StpPrG = True) Then
    If (StrComp(Me.txtPassword.Text, frmSetup.PssWrdSupr) <> 0) Then
        rtn$ = MsgBox("You have not entered the correct password!", vbOKOnly, "SurfTime")
        txtPassword.Text = Empty
        Exit Function
    Else
        Unload Me
        frmSetup.Visible = True
    End If
End If
 'This function validates the user's password against the one that is stored on the file.
If (PsWrd1 = True) Then
        If (PsWrd2 = False And Reset = False) Then
        If (StrComp(Me.txtPassword.Text, Globals.RdFlPssWrd2) <> 0) Then
            rtn$ = MsgBox("You have not entered the correct password!", vbOKOnly, "SurfTime")
            txtPassword.Text = Empty
            Exit Function
        Else
            Unload Me
            Load frmNewPassword
            frmNewPassword.Visible = True
            YeaOrNea = True
        End If
    Else
        If (StrComp(Me.txtPassword.Text, Globals.RdFlPssWrd) <> 0) Then
            rtn$ = MsgBox("You have not entered the correct password!", vbOKOnly, "SurfTime")
            txtPassword.Text = Empty
            Exit Function
        End If
    End If
End If

If (PsWrd2 = True) Then
    'This is the Supervisors password routine.
    If (OvrRide = True And PsWrd1 = False And PsWrd2 = True And Rst = False) Then
        If (StrComp(txtPassword.Text, Globals.RdFlPssWrd2) <> 0) Then
            rtn$ = MsgBox("You have not entered the correct password!", vbOKOnly, "SurfTime")
            txtPassword.Text = Empty
            YeaOrNea = False
            Exit Function
        End If
    Else
        If (StrComp(txtPassword.Text, Globals.RdFlPssWrd) <> 0) Then
            rtn$ = MsgBox("You have not entered the correct password!", vbOKOnly, "SurfTime")
            txtPassword.Text = Empty
            YeaOrNea = False
            Exit Function
        End If
    End If
    
    If (OvrRide <> True And PsWrd2 = True And Rst = False And EnblTskMn = False) Then
        'Change default hour selected or Windows applications.
            YeaOrNea = True
            If (BrwsrSt = True) Then
                iniBrowser = True
                BrwsrSt = False
            End If
            Unload Me
            Load frmRstTime
            frmRstTime.Visible = True
    ElseIf (OvrRide <> False And PsWrd2 = False And Rst = False And EnblTskMn = False And ChGAdPss = False) Then
        'Change Administrator password selected.
            Unload Me
            PsWrd1 = False
            Load frmNewPassword
            frmNewPassword.Visible = True
    ElseIf (OvrRide <> False And PsWrd2 = True And Rst = False And EnblTskMn = False And ChGAdPss = True) Then
        'Change Master password selected or Exit program.
            If (ExtPrG = False) Then
                Unload Me
                PsWrd2 = False
                Load frmNewPassword
                frmNewPassword.Visible = True
            Else
                frmSurfMonitor.SystemTray1.Action = sys_Delete
                frmProwler.SystemTray1.Action = sys_Delete
                Unload Me
                Msg$ = MsgBox("SurfTime has been closed down", vbOKOnly, "SurfTime")
                If (frmSurfMonitor.SurfTimeSocket.State = 7) Then
                    frmSurfMonitor.SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_APP_EXIT & Chr(58) & ST_END
                    frmSurfMonitor.SurfTimeSocket.Close
                End If
                App.TaskVisible = True
                End
            End If
    ElseIf (OvrRide = True And PsWrd2 = True And Rst = False And EnblTskMn = False) Then
        'Override selected.
            Hours = Unlimited
            DestrState = False
            Msg$ = MsgBox("You now have unlimited surf time.", vbOKOnly, "SurtTime Unlimited")
            frmProwler.Caption = "Unlimited surf time..."
            Unload Me
            SoftwareLoad
            Load frmTmr
            frmTmr.Visible = True
            frmTmr.Timer1.Enabled = True
            If (frmSurfMonitor.SurfTimeSocket.State = 7) Then
                frmSurfMonitor.SurfTimeSocket.SendData STIP_AddressCl & Chr(58) & ST_APP_ACTIVATE & Chr(58) & ST_END
            End If
    ElseIf (OvrRide = True And PsWrd2 = True And Rst = True And EnblTskMn = False) Then
        'Reset selected.
            ResetUnload
    Else: YeaOrNea = True
    End If
End If
MnuClicked = False
chkpsswrd = YeaOrNea
End Function
