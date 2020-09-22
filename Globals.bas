Attribute VB_Name = "Globals"
'********************************************************************************
'                SurfTime Copyright of the Stuffed Dog Company 2000.
'                       Programming and design by David Barker.
            'No part of this code may be reproduced, copied or distributed,
'                   without the written permission from the author.
'               This software is the property of The Stuffed Dog Company
'
'               This software was developed for Ashosh. Only Ashosh has permission
'                    and license to copy, distribute and backup this software.
'                   This permission does not apply to the source code!

'                          Contact: David Barker - 0958 280269
'                      E-mail address: davidbarker38@hotmail.com
Option Explicit
'********************************************************************************
Public Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function UpdateLayeredWindow Lib "user32" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
'****************************************************************************************************************************
Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Public Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        offset As Long
        OffsetHigh As Long
        hEvent As Long
End Type
'****************************************************************************************************************************
Public Declare Function gethostname Lib "WSOCK32.DLL" (ByVal hostname$, ByVal HostLen As Long) As Long
Public Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal hostname$) As Long
Public Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)
Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long

Public Const WS_VERSION_REQD = &H101
Public Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD = 1
Public Const SOCKET_ERROR = -1
Public Const WSADescription_Len = 256
Public Const WSASYS_Status_Len = 128

Public Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired&, lpWSAData As WSADATA) As Long
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

Public Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type

Public Type WSADATA
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADescription_Len) As Byte
    szSystemStatus(0 To WSASYS_Status_Len) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpszVendorInfo As Long
End Type

Global Const BROADCASTPORT = 1055
'****************************************************************************************************************************

'****************************************************************************************************************************
Const REG_SZ = 1 ' Unicode nul terminated String
Const REG_DWORD = 4 ' 32-bit number
'****************************************************************************************************************************
Global Const RGN_XOR = 3
Public Const RGN_OR = 2
Global Const SWP_NOSIZE = &H1
Global Const HWND_BOTTOM = 1
Global Const SWP_NOMOVE = &H2
Global Const SW_SHOWMAXIMIZED = 3
Global Const EW_REBOOT = &H43
Global Const EW_RESTART = &H42
Global Const EW_EXIT = 0
Global Const WM_CLOSE = &H10
Global Const GWL_EXSTYLE = (-20)
Global Const GWL_HINSTANCE = (-6)
Global Const GWL_HWNDPARENT = (-8)
Global Const LWA_COLORKEY = &H1
Global Const LWA_ALPHA = &H2
Global Const ULW_COLORKEY = &H1
Global Const ULW_ALPHA = &H2
Global Const ULW_OPAQUE = &H4
Global Const WS_EX_LAYERED = &H80000
Global Const RSP_SIMPLE_SERVICE = 1
Global Const RSP_UNREGISTER_SERVICE = 0
Public Const SPI_SCREENSAVERRUNNING = 97
Public Const MF_BYPOSITION = &H400&
'****************************************************************************************************************************
Public RdFlPssWrd As String
Public RdFlPssWrd2 As String
Public EnblTskMn As Boolean
'These variables determine what password is entered or not.
Public PsWrd2 As Boolean
Public PsWrd1 As Boolean
'****************************************************************************************************************************
'Window Message Pipe Constants
'Setting Time Limits
Public Const ST_TIMEOUT_01HRS = &H186A1
Public Const ST_TIMEOUT_02HRS = &H186A2
Public Const ST_TIMEOUT_03HRS = &H186A3
Public Const ST_TIMEOUT_04HRS = &H186A4
Public Const ST_TIMEOUT_05HRS = &H186A5
Public Const ST_TIMEOUT_06HRS = &H186A6
Public Const ST_TIMEOUT_07HRS = &H186A7
Public Const ST_TIMEOUT_08HRS = &H186A8
Public Const ST_TIMEOUT_09HRS = &H186A9
Public Const ST_TIMEOUT_10HRS = &H186AA
Public Const ST_TIMEOUT_UNLMTD = &H186AB
'Starting Applications
Public Const STLOAD_MS_WORD = &HC350
Public Const STLOAD_MS_ACCESS = &HC351
Public Const STLOAD_MS_EXCEL = &HC352
Public Const STLOAD_MS_POWERPOINT = &HC353
Public Const STLOAD_MS_PUBLISHER = &HC354
Public Const STLOAD_MS_IEXPLORER = &HEA60
Public Const STLOAD_MS_NETSCAPE = &HEA61
'Running Applications
Public Const ST_MS_WORD = &HC360
Public Const ST_MS_ACCESS = &HC361
Public Const ST_MS_EXCEL = &HC362
Public Const ST_MS_POWERPOINT = &HC363
Public Const ST_MS_PUBLISHER = &HC364
Public Const ST_MS_IEXPLORER = &HEA70
Public Const ST_MS_NETSCAPE = &HEA71
'Resetting SurfTime
Public Const ST_TIME_RESET = &H1046A
'Overriding SurfTime
Public Const ST_TIME_OVERRIDE = &H1046B
'Task Manager commends
Public Const ST_TASkMANAGER_ACTIVE = &HB7001
Public Const ST_TASkMANAGER_INACTIVE = &HB7002
'Network Wide Broadcast to all SurfTime Applications
'Changing Password
Public Const ST_CHGPSSWRD_ADMIN = &H46B71
Public Const ST_CHGPSSWRD_SUPER = &H46B72
Public Const ST_BROADCAST_IP_RETURN = &H17DE8
'Testing if SurfTime Exists
Public Const ST_SURFTIME_ACTIVE = &H46B73
'SurfTime Professional Message Return
Public Const ST_APP_ACTIVATE = &H2D6DE
Public Const ST_APP_EXIT = &H2D6DF
Public Const ST_APP_ClOSE = &H2D7DF
Public Const ST_ALL = &H5FFF5
Public Const ST_TM_REM = &H2D6E0
Public Const ST_ACTIVE_RUNNING = &H2D6E1
Public Const ST_NOT_INSTALLED = &H2D6E2
Public Const ST_PASSWORD_UPDATED = &H4FFF2
Public Const ST_SHUTDOWN = &H4FFF5
Public Const ST_IPADDRESS = &H2D6E3
'End Message
Public Const ST_END = &H1FBD1
'****************************************************************************
Public Hours As Long
Public TmOutHr As Long
Public WrdTm As Long
Public XclTm As Long
Public AcssTm As Long
Public PwrPntTm As Long
Public PwrPntAddTm As Boolean
Public ExcelAddTm As Boolean
Public AccessAddTm As Boolean
Public WrdAddTm As Boolean
Public OvrRide As Boolean
Public NwTm As Boolean
Public lParent As Long
Public NTPlatform As Boolean
Public FindAccess As Boolean
Public ChGAdPss As Boolean
Public DestrState As Boolean
Public Reset As Boolean
Public Rst As Boolean
Public WndHndl As Variant
Public MsWord As Boolean
Public MsAccess As Boolean
Public MsPowerPoint As Boolean
Public MsExcel As Boolean
Public MsPublisher As Boolean
Public IExplorer As Boolean
Public Netscape As Boolean
Public MsWordEst As Boolean
Public MsAccessEst As Boolean
Public MsPowrPntEst As Boolean
Public PublisherEst As Boolean
Public MsExcelEst As Boolean
Public MnuClicked As Boolean
Public Application As Boolean
Public iniBrowser As Boolean
Public BrwsrSt As Boolean
Public ExtPrG As Boolean
Public MnuPoslft As Long
Public MnuPostp As Long
Public STIP_AddressCl As String
Public StpPrG As Boolean
Public Const sys_Add = 0       'Specifies that an icon is being add
Public Const sys_Modify = 1    'Specifies that an icon is being modified
Public Const sys_Delete = 2    'Specifies that an icon is being deleted
Public HostServerIP As String

Global Const TmHour = 60
Global Const DftTm = 1
Global Const Unlimited = 100
Global Const TmConst = "hh:mm"
Global Const ExplrPth = "C:\Program Files\Internet Explorer\IEXPLORE.exe"
Global Const ExplrPth2 = "C:\Program Files\Plus!\Microsoft Internet\IEXPLORE.exe"
Global Const NetScpPth = "c:\Program Files\Netscape\Communicator\Program\Netscape.exe"
Global Const NetScp6Pth = "c:\Program Files\Netscape\Netscape6\Netscp6.exe"
'Check these paths to make sure that they are common to all platforms.
Global Const MsWrdPth = "C:\Program Files\Microsoft Office\Office\Winword.exe"
Global Const MsExlPth = "C:\Program Files\Microsoft Office\Office\Excel.exe"
Global Const MsAccsPth = "C:\Program Files\Microsoft Office\Office\MsAccess.exe"
Global Const MsPwrPntPth = "C:\Program Files\Microsoft Office\Office\PowerPnt.exe"
Global Const MsPubPth = "C:\Program Files\Microsoft Office\Office\MSPub.exe"

Public Type POINTAPI
        x As Long
        y As Long
End Type

'*******************************************************
'this code makes the window stay on top.
'rtn = SetWindowPos(Me.hwnd, -2, 0, 0, 0, 0, 3)
'this code makes the window stay on bottom.
'rtn = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)
'*******************************************************
