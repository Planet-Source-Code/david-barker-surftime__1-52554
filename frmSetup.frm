VERSION 5.00
Begin VB.Form frmSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SurfTime - Setup"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   120
      TabIndex        =   10
      Top             =   3735
      Width           =   7215
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6240
      TabIndex        =   5
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox txtAdminConfirm 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Text            =   " "
      Top             =   1920
      Width           =   3735
   End
   Begin VB.TextBox txtSuperConfirm 
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Text            =   " "
      Top             =   3120
      Width           =   3735
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox txtSuper 
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Text            =   " "
      Top             =   2760
      Width           =   3735
   End
   Begin VB.TextBox txtAdmin 
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Text            =   " "
      Top             =   1560
      Width           =   3735
   End
   Begin VB.Label Label5 
      Caption         =   "SurfTime will run automatically every time you start your computer."
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   4815
   End
   Begin VB.Image Image2 
      Height          =   750
      Left            =   -360
      Top             =   0
      Width           =   7800
   End
   Begin VB.Label Label4 
      Caption         =   "Enter password again"
      Height          =   255
      Left            =   4440
      TabIndex        =   9
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Enter password again"
      Height          =   255
      Left            =   4440
      TabIndex        =   8
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Supervisiors Password"
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Administrators Password"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   1320
      Width           =   2055
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PssWrdAdmin As String
Public PssWrdSupr As String
Public Baamba As Boolean
Public Enum eAutoRunTypes
    eNever
    eOnce
    ealways
End Enum
Private Sub btnCancel_Click()
End
End Sub

Private Sub btnOk_Click()
'Validate all the text boxes before writing to disk.
If (txtAdmin.Text = Empty Or txtAdminConfirm.Text = Empty) Then
    MsgBox "Please enter your administrators password."
    Exit Sub
End If

If (txtSuper.Text = Empty Or txtSuperConfirm.Text = Empty) Then
    MsgBox "Please enter your supervisors password."
    Exit Sub
End If

If (StrComp(Trim(txtAdmin.Text), Trim(txtAdminConfirm.Text)) = 0 And _
    StrComp(Trim(txtSuper.Text), Trim(txtSuper.Text)) = 0) Then
        If (txtAdmin.Text <> Chr(32)) And (txtAdminConfirm.Text <> Chr(32)) And _
            (txtSuper.Text <> Chr(32)) And (txtSuperConfirm.Text <> Chr(32)) Then
            ChGAdPss = True
            CreateNewPssWrd (Trim(txtAdmin.Text))
            ChGAdPss = False
            CreateNewPssWrd (Trim(txtSuper.Text))
            Unload Me
            AutoRun = ealways
            StartProG
        Else: ClearAll: MsgBox "Please try re-entering your missing passwords again!"
        End If
Else: ClearAll: MsgBox "Password inaccurate, please retype again!"
End If


End Sub

Private Sub Form_Load()
Dim PrvInst As Boolean
    If (Dir(App.Path & Chr(92) & "Security.pwl") <> Empty) Then
        Unload Me
        StartProG
    Else
        'Check if the software was previously installed.
        PrvInst = Detctpth
        If (PrvInst = True) Then
            'Read from the found file the previous passwords then ask the user to input them.
            StpPrG = True
            Me.Hide
            MsgBox "SurfTime was previously installed on this computer."
            MsgBox "Please enter the master password, before installation will take place."
            Load frmLogin
            frmLogin.Visible = True
        End If
    End If

End Sub

Private Function ClearAll()
txtAdmin.Text = Empty
txtAdminConfirm.Text = Empty
txtSuper.Text = Empty
txtSuperConfirm.Text = Empty
End Function

Private Sub Form_Resize()
Me.Height = 4995
Me.Width = 7530
End Sub
Public Property Let AutoRun(ByVal eType As eAutoRunTypes)
Dim sExe As String

    sExe = App.Path
    If (Right$(sExe, 1) <> "\") Then sExe = sExe & "\"
    sExe = sExe & App.EXEName
    
    Dim cR As New cRegistry
    cR.ClassKey = HKEY_LOCAL_MACHINE
    If (eType = eNever) Then
        ' Remove entry from always Run if it is there:
        cR.SectionKey = "Software\Microsoft\Windows\CurrentVersion\Run"
        cR.ValueKey = App.EXEName
        On Error Resume Next
        cR.DeleteValue
        Err.Clear
        ' Remove entry from RunOnce if it is there:
        cR.SectionKey = "Software\Microsoft\Windows\CurrentVersion\RunOnce"
        On Error Resume Next
        cR.DeleteValue
        Err.Clear
    ElseIf eType = eOnce Then
        ' Remove entry from always Run if it is there:
        cR.SectionKey = "Software\Microsoft\Windows\CurrentVersion\Run"
        cR.ValueKey = App.EXEName
        On Error Resume Next
        cR.DeleteValue
        Err.Clear
        ' Add an entry to RunOnce (or just ensure the exe name and path
        ' is correct if it is already there):
        cR.SectionKey = "Software\Microsoft\Windows\CurrentVersion\RunOnce"
        cR.ValueKey = App.EXEName
        cR.ValueType = REG_SZ
        cR.Value = sExe
    Else
        ' Remove entry from RunOnce if it is there:
        cR.SectionKey = "Software\Microsoft\Windows\CurrentVersion\RunOnce"
        cR.ValueKey = App.EXEName
        On Error Resume Next
        cR.DeleteValue
        Err.Clear
        ' Add an entry to RunOnce (or just ensure the exe name and path
        ' is correct if it is already there):
        cR.SectionKey = "Software\Microsoft\Windows\CurrentVersion\Run"
        cR.ValueKey = App.EXEName
        cR.ValueType = REG_SZ
        cR.Value = sExe
    End If
        
End Property
Public Property Get AutoRun() As eAutoRunTypes
    Dim cR As New cRegistry
    cR.ClassKey = HKEY_LOCAL_MACHINE
    cR.SectionKey = "Software\Microsoft\Windows\CurrentVersion\Run"
    cR.ValueKey = App.EXEName
    cR.Default = "?"
    cR.ValueType = REG_SZ
    If (cR.Value = "?") Then
        cR.SectionKey = "Software\Microsoft\Windows\CurrentVersion\RunOnce"
        If (cR.Value = "?") Then
            AutoRun = eNever
        Else
            AutoRun = eOnce
        End If
    Else
        AutoRun = ealways
    End If
End Property

Private Function Detctpth() As Boolean
Dim Wndws As Boolean
Dim NT As Boolean
    'This makes sure that if the security file is found then SurfTime was installed before.
    If (Dir("C:\Windows\ST4302.stb") <> Empty) Then
        Wndws = True
        OpenPrvFl ("C:\Windows\ST4302.stb")
    End If

    If (Dir("C:\WINNT\ST4302.stb") <> Empty) Then
        NT = True
        OpenPrvFl ("C:\WINNT\ST4302.stb")
    End If

    If (Wndws = True Or NT = True) Then Detctpth = True Else Detctpth = False

End Function

Private Sub OpenPrvFl(Pth As String)
Dim EncrptStG As String
Dim idx As Byte
Dim PswStrG As String
Dim NxtPos As Long
Dim PosEq As Long

    Open (Pth) For Input As #1
        For idx = 1 To 2
            Line Input #1, EncrptStG
            If (idx = 1) Then
                PssWrdAdmin = Crypt(EncrptStG)
                PswStrG = PssWrdAdmin
            Else: PssWrdSupr = Crypt(EncrptStG)
                PswStrG = PssWrdSupr
            End If
            
            PosEq = InStr(PswStrG, Chr(61)) + 1
            PswStrG = Switch(idx, Left$(Mid(PswStrG, PosEq), Len(PswStrG) - PosEq))
            If (idx = 1) Then PssWrdAdmin = PswStrG
            If (idx = 2) Then PssWrdSupr = PswStrG
        Next
    Close #1
    
End Sub

Private Sub txtSuperConfirm_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then btnOk_Click
End Sub
