Attribute VB_Name = "Disable"
Public Sub Disable_Ctrl_Alt_Del()
Dim AyW As Integer
Dim TurFls As Boolean
    AwY = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, TurFls, 0)
    If (AwY = 0) Then
        NTPlatform = True
        frmSurfMonitor.mnuDisable.Enabled = False
        frmSurfMonitor.mnuEnTsk.Enabled = False
    Else
        'AwY = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, TurFls, 0)
        NTPlatform = False
        App.TaskVisible = False
    End If
End Sub
