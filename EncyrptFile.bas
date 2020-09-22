Attribute VB_Name = "EncyrptFile"

Public Function Crypt(Text As String) As String
Dim strTempChar As String
    'Get the length of the passed data to decide how many iterations need to be made.
    For i = 1 To Len(Text)
        'If the string or character is ASCII, then convert it into extended character code.
        If Asc(Mid$(Text, i, 1)) < 128 Then
            strTempChar = Asc(Mid$(Text, i, 1)) + 128
        ElseIf Asc(Mid$(Text, i, 1)) > 128 Then
            'If the characters are extended ASCII, then convert them back into ASCII.
            strTempChar = Asc(Mid$(Text, i, 1)) - 128
        End If
        'Replace the passed string with the new character code.
        Mid$(Text, i, 1) = Chr(strTempChar)
    Next i
    
    'Return the sent data back to the procedure which called it.
    Crypt = Text

End Function

