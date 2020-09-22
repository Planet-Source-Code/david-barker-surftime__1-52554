Attribute VB_Name = "Enable"
Public Sub Enable_Ctrl_Alt_Del()
    'Enables the Crtl+Alt+Del
    Dim AwY As Integer
    Dim TurFls As Boolean
    AwY = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, TurFls, 0)

End Sub
