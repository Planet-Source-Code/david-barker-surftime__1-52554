VERSION 5.00
Begin VB.Form frmRstTime 
   Caption         =   "Default Settings Change"
   ClientHeight    =   1095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3450
   Icon            =   "frmRstTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   3450
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnOvrRd 
      Caption         =   "Override"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.VScrollBar VScroll1 
      Enabled         =   0   'False
      Height          =   495
      Left            =   1680
      Max             =   1
      Min             =   10
      TabIndex        =   3
      Top             =   120
      Value           =   2
      Width           =   255
   End
   Begin VB.TextBox txtHour 
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   240
      Width           =   495
   End
   Begin VB.CheckBox chkDflt 
      Caption         =   "Default setting 1 hour"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Time in Hours"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmRstTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TmSet As Boolean
Private Sub Check1_Click()
    If (chkDflt.Value = Chr(49)) Then
        chkDflt.Value = Chr(48)
        NwTm = True
    Else
        chkDflt.Value = Chr(49)
    End If

End Sub

Private Sub btnOk_Click()
Dim TmpTm As Long
On Error Resume Next
    If (chkDflt.Value <> Chr(49)) Then
        If (VScroll1.Value > 1) Then
            Hours = CLng(txtHour.Text)
            NwTm = True
            If (iniBrowser = False) Then SetAppTm
            If (TmSet = False) Then
                TmpTm = (Hours * TmHour)
                If (TmpTm < TmOutHr) Then TmOutHr = TmpTm Else TmOutHr = TmOutHr + (Hours * TmHour)
            End If
            TmSet = False
        Else
            Hours = CLng(txtHour.Text)
            NwTm = True
            If (iniBrowser = False) Then SetAppTm
            If (TmSet = False) Then
                TmOutHr = TmOutHr + 60
            End If
            TmSet = False
        End If
    Else
        Hours = Chr(49)
        NwTm = True
        If (iniBrowser = False) Then SetAppTm
        If (TmSet = False) Then
            TmOutHr = TmOutHr + 60
        End If
        TmSet = False
    End If
    
    SoftwareLoad
    DestrState = False
    Unload Me
    iniBrowser = False
    Load frmTmr: frmTmr.Visible = True: frmTmr.Timer1.Enabled = True
End Sub

Private Sub btnOvrRd_Click()
    If (OvrRide = True) Then Hours = Unlimited
End Sub

Private Sub chkDflt_Click()
    If (chkDflt.Value <> Chr(49)) Then
        Hours = Chr(49)
        VScroll1.Enabled = True
        txtHour.Text = Chr(49)
        txtHour.Enabled = True
        txtHour.BackColor = vbWhite
    Else    'If the check default button is enabled then disable any settings.
        'VScroll1.Value = Chr(49)
        VScroll1.Enabled = False
        txtHour.Text = Chr(49)
        txtHour.Enabled = False
        txtHour.BackColor = vbButtonFace
    End If
End Sub

Private Sub Form_Load()
    If (NwTm = False) Then
        txtHour.Text = Chr(49)
        chkDflt.Value = Chr(49)
    Else
        VScroll1.Value = Chr(49)
        txtHour.Text = Chr(49)
        chkDflt.Value = Chr(48)
    End If
    
If (chkDflt.Value = 0) Then
    txtHour.Enabled = True
    txtHour.BackColor = vbWhite
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    PsWrd2 = False
End Sub

Private Sub VScroll1_Change()
If (VScroll1.Value = 1) Then VScroll1.Value = 2
    txtHour.Text = VScroll1.Value - 1
End Sub

Public Sub SetAppTm()
Dim AppTm(1 To 4) As Long
Dim Tmp As Long
Dim cntr As Byte
Dim idx As Byte

    If (WrdAddTm = True) Then WrdTm = WrdTm + (Hours * TmHour): WrdAddTm = False: TmSet = True
    If (ExcelAddTm = True) Then XclTm = XclTm + (Hours * TmHour): ExcelAddTm = False: TmSet = True
    If (AccessAddTm = True) Then AcssTm = AcssTm + (Hours * TmHour): AccessAddTm = False: TmSet = True
    If (PwrPntAddTm = True) Then PwrPntTm = PwrPntTm + (Hours * TmHour): PwrPntAddTm = False: TmSet = True
    
    For cntr = 0 To 3
        Select Case cntr
            Case Is = 0
                AppTm(1) = WrdTm
            Case Is = 1
                AppTm(2) = XclTm
            Case Is = 2
                AppTm(3) = AcssTm
            Case Is = 3
                AppTm(4) = PwrPntTm
        End Select
    Next
    
    cntr = 1
    idx = 1
    While cntr < UBound(AppTm)
    
        While idx < UBound(AppTm)
            If (AppTm(idx) < AppTm(idx + 1)) Then
                Tmp = AppTm(idx)
                AppTm(idx) = AppTm(idx + 1)
                AppTm(idx + 1) = Tmp
            End If
            idx = idx + 1
        Wend
        idx = 1
        cntr = cntr + 1
    Wend
    
    If (TmOutHr < AppTm(1)) Then TmOutHr = AppTm(1)
    
End Sub
