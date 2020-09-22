Attribute VB_Name = "MsOffice"
Public Sub MicrosoftOffice()
        If (Dir(MsWrdPth) = Empty) Then
            frmSurfMonitor.mnuWord.Enabled = False
            'Set the variable which tells the computer that the application does not exist.
            MsWordEst = False
        Else: MsWordEst = True
        End If
    
        If (Dir(MsExlPth) = Empty) Then
            frmSurfMonitor.mnuExcel.Enabled = False
            MsExcelEst = False
        Else: MsExcelEst = True
        End If
    
        If (Dir(MsAccsPth) = Empty) Then
            frmSurfMonitor.mnuAccess.Enabled = False
            MsAccessEst = False
        Else: MsAccessEst = True
        End If
    
        If (Dir(MsPwrPntPth) = Empty) Then
            frmSurfMonitor.mnuPowerPoint.Enabled = False
            MsPowrPntEst = False
        Else: MsPowrPntEst = True
        End If
    
        If (Dir(MsPubPth) = Empty) Then
            frmSurfMonitor.mnuPublisher.Enabled = False
            PublisherEst = False
        Else: PublisherEst = True
        End If
End Sub
