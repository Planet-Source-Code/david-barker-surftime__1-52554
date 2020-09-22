Attribute VB_Name = "WndParent"
Public Function WndEnumProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    Dim wtext As String * 512
    Dim bRet As Long, WLen As Long
    Dim wclass As String * 50
    Dim Pos As Long
    Dim idx As Integer
    Dim lmyChild As Long
    Dim lmyParent As Long
    Dim HwndExst As Boolean
    WLen = GetWindowTextLength(hWnd)
    bRet = GetWindowText(hWnd, wtext, WLen + 1)
    GetClassName hWnd, wclass, 50
    'This piece of code is for Microsoft Publisher
    '*******************************
    ' Or InStr(WClass, "MSWinPub") > 0 And InStr(WClass, "#32770") > 0
    '*******************************
    If (FindAccess = True And MsAccess = False) Then
        If (InStr(wclass, "OMain") > 0) Then
            CloseBrowser (hWnd)
            FindAccess = False
        End If
    End If
    
    If (NTPlatform = True) Then TaskManager wclass, wtext, hWnd
    If (InStr(wclass, "IEFrame") > 0) Then
        'This is the correct Window below, for using Windows API for searching through the text on the Web Page for illicit words.
        'lChild& = FindWindowEx(hWnd, 0, "Shell DocObject View", vbNullString)
        'CloseBrowser (lChild&)
    End If
    
    'These are Internet sex programs which run on the pc.
    If (InStr(wtext, "Dialers") > 0 And InStr(wclass, "CabinetWClass") > 0 _
        Or InStr(wtext, "UNCUTBIGBROTHER_COM") > 0 And InStr(wclass, "DlgClass") > 0 _
            Or InStr(wtext, "Go in Direct") > 0 And InStr(wclass, "#32770") > 0 _
                Or InStr(wtext, "Erotica") > 0 And InStr(wclass, "TApplication") _
                    Or InStr(wtext, "Cuca-fresca") > 0 And InStr(wclass, "TCuca_fresca") > 0) Then
                        CloseBrowser (hWnd)
                        Exit Function
    End If
    
    If (InStr(wclass, "ExploreWClass") > 0) Then
        CloseBrowser (hWnd)
        Exit Function
    End If
    
    If (InStr(wclass, "MozillaWindowClass") > 0 Or InStr(wclass, "Afx:400000:0") > 0 And InStr(wtext, "Preview") = 0 And InStr(wtext, "Paint Shop Pro") = 0 _
          Or InStr(wclass, "IEFrame") > 0 And DestrState = True Or InStr(wclass, "MSN6 Window") > 0 And DestrState Or InStr(wclass, "AOL Frame25") > 0 And DestrState Or InStr(wtext, "Catalog") > 0 Or InStr(wclass, "CabinetWClass") > 0 _
            Or InStr(wtext, "CompuServe") > 0 _
                Or InStr(wtext, "Microsoft Word") > 0 And MsWord = False _
                    Or InStr(wtext, "Microsoft PowerPoint") > 0 And MsPowerPoint = False _
                        Or InStr(wclass, "XLMAIN") > 0 And MsExcel = False _
                            Or InStr(wtext, "Paltalk Logon") > 0 And InStr(wclass, "#32770") _
                                Or InStr(wtext, "ICQ Welcome") > 0 And InStr(wclass, "#32770") _
                                    Or InStr(wtext, "Excite Virtual Places Chat") > 0 And InStr(wclass, "VP Frame") _
                                        Or InStr(wtext, "NJStar Internet Search Center") _
                                            Or InStr(wtext, "MSN Messenger Service") > 0 And InStr(wclass, "MSBLClass") _
                                                Or InStr(wtext, "Sign On") > 0 And InStr(wclass, "AIM_CSignOnWnd") _
                                                    Or InStr(wtext, "AOL InstantMessenger") And InStr(wclass, "#32770") _
                                                        Or InStr(wclass, "_Oscar_StatusNotify") Or InStr(wclass, "Oscar_Balloon") _
                                                            Or InStr(wtext, "Buddy List Window") And InStr(wclass, "_Oscar_BuddyListWin") _
                                                                Or InStr(wtext, "Yahoo! Messenger") > 0 And InStr(wclass, "YahooBuddyMain") _
                                                                    Or InStr(wtext, "RocketPipe") > 0 And InStr(wclass, "RocketPipe Category") _
                                                                        Or InStr(wtext, "Tencent Explorer") And InStr(wclass, "#32770") _
                                                                            Or InStr(wtext, "Tencent Explorer") And InStr(wclass, "Afx:4000000:0") _
                                                                                Or InStr(wtext, "Microsoft Access") > 0 And MsAccess = False) _
                                                                                    Or InStr(wtext, "Microsoft Internet Explorer") > 0 And InStr(wclass, "CabinetWClass") > 0 Then
                                                                                        If (InStr(wtext, "IE3_Class") = 0 And InStr(wtext, "Microsoft Spy") = 0 And InStr(wclass, "Afx:400000:8:10011:0:0") = 0) Then
                                                                                            If (FindWindow("IE3_Class", vbNullString) = 0) Then
                                                                                                If (InStr(wtext, "Effect Palette") = 0 And InStr(wclass, "Afx:400000:8") = 0) Then
                                                                                                    If (InStr(wtext, "Site - UD4 Tutorial - ASP") = 0 And InStr(wclass, "Afx:400000:b:") = 0) Then
                                                                                                    'If the window is an Access window then lookto see if there are any child windows first to close down.
                                                                                                        If (InStr(wclass, "#32770") > 0) Then FindAccess = True
                                                                                                        If (InStr(wtext, "Microsoft Word") > 0 And FindWindow("bosa_sdm_Microsoft Word 9.0", vbNullString) <> 0) Then
                                                                                                            lmyParent = FindWindow("bosa_sdm_Microsoft Word 9.0", vbNullString)
                                                                                                            CloseBrowser (lmyParent)
                                                                                                        ElseIf (InStr(wtext, "Microsoft PowerPoint") > 0 And InStr(wclass, "#32770") > 0) Then
                                                                                                            CloseBrowser (hWnd)
                                                                                                        End If
                                                                                                        CloseBrowser (hWnd)
                                                                                                        'Load frmCorpLgExt: frmCorpLgExt.Visible = True
                                                                                                    End If
                                                                                                End If
                                                                                            End If
                                                                                        End If
                                                                                    End If
    WndEnumProc = 1
End Function

Public Function WndEnumFnd(ByVal hWnd As Long, ByVal lParam As Long) As Long
Dim wtext As String * 512
Dim bRet As Long, WLen As Long
Dim wclass As String * 50
Dim Pos As Long
Dim idx As Integer
Dim ArrySz As Long
Dim HwndExst As Boolean
    ArrySz = UBound(WndHndl, 1)
    WLen = GetWindowTextLength(hWnd)
    bRet = GetWindowText(hWnd, wtext, WLen + 1)
    GetClassName hWnd, wclass, 50
    If (InStr(wclass, "MozillaWindowClass") > 0 Or InStr(wclass, "Afx:400000:") > 0 _
        Or InStr(wclass, "IEFrame") > 0 Or InStr(wclass, "MSN6 Window") > 0 Or InStr(wclass, "AOL Frame") > 0) Then
            If (InStr(wtext, "IE3_Class") = 0 And InStr(wtext, "Microsoft Spy") = 0 And InStr(wclass, "Afx:400000:8:10011:0:0") = 0) Then
                ReDim Preserve WndHndl(ArrySz + 1)
                WndHndl(ArrySz + 1) = Array(hWnd)
            End If
    End If
    WndEnumFnd = 1
End Function

Public Function WndEnumMsOffice(ByVal hWnd As Long, ByVal lParam As Long) As Long
Dim wtext As String * 512
Dim bRet As Long, WLen As Long
Dim wclass As String * 50
Dim Pos As Long
Dim idx As Integer
Dim ArrySz As Long
Dim HwndExst As Boolean
Dim Language_Definition As Boolean

    WLen = GetWindowTextLength(hWnd)
    bRet = GetWindowText(hWnd, wtext, WLen + 1)
    GetClassName hWnd, wclass, 50

    If (FindAccess = True And MsAccess = False) Then
        If (InStr(wclass, "OMain") > 0) Then
            CloseBrowser (hWnd)
            FindAccess = False
        End If
    End If

    If (NTPlatform = True) Then TaskManager wclass, wtext, hWnd
    'If (DestrState = False And InStr(wclass, "IEFrame") > 0) Then
    '    Language_Definition = Stage3_Content_Filter(wtext)
    '    If (Language_Definition = True) Then
    '        Load frmForbidden
    '        frmForbidden.Visible = True
    '        CloseBrowser (hWnd)
    '    End If
    'End If
    
    If (InStr(wclass, "IEFrame") > 0) Then
        'This is the correct Window below, for using Windows API for searching through the text on the Web Page for illicit words.
        lChild& = FindWindowEx(hWnd, 0, "Shell DocObject View", vbNullString)
        ReadWBControlFormat (hWnd)
        'CloseBrowser (lChild&)
    End If
    
    If (InStr(wtext, "Microsoft Word") > 0 And MsWord = False _
        Or InStr(wtext, "Microsoft PowerPoint") > 0 And MsPowerPoint = False _
            Or InStr(wclass, "XLMAIN") > 0 And MsExcel = False _
                Or InStr(wtext, "Microsoft Access") > 0 And InStr(wclass, "#32770") > 0 And MsAccess = False) Then
                    If (InStr(wclass, "#32770") > 0) Then
                        FindAccess = True
                   End If
                    CloseBrowser (hWnd)
                    Load frmCorpLgExt: frmCorpLgExt.Visible = True
    End If
    WndEnumMsOffice = 1
End Function

Public Sub TaskManager(wclass As String, wtext As String, hWnd As Long)
Dim lParent As Long
Dim lChild As Long
Dim lChild2 As Long

'This routine makes sure that if the task manager is enabled then no-one will have access to it to stop SurfTime from running.
    If (InStr(wtext, "Windows NT Task Manager") > 0 Or InStr(wtext, "Windows Task Manager") > 0 And InStr(wclass, "#32770") > 0) Then
        lParent = FindWindow("#32770", vbNullString)
        lChild = FindWindowEx(lParent, 0, "#32770", vbNullString)
        lChild2 = FindWindowEx(lChild, 0, "SysListView32", vbNullString)
        CloseBrowser (lChild2)
    End If
    
End Sub

Private Sub ReadWBControlFormat(ByRef handle As Long)
Dim Nwvl As String
'*******************************************************************************************
'This routine is for validating all the text on the web page making sure that all the _
contents of the page fall within the acceptable criteria for what is considered a _
reasonable and good site. Any words which suggest that the site is a lewd one _
will automatically warrant instant closure.
'*******************************************************************************************


End Sub
