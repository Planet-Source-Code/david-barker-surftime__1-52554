Attribute VB_Name = "Content_Filter"
Public Function Stage3_Content_Filter(ByVal Wordcheck As String) As Boolean
Dim OffnWrds As String
Dim illGlWrd As String
Dim Pos As Long
Dim pos2 As Long
Dim pos3 As Long
Dim nxtSpc As Long
Dim frstSpc As Long
Dim numSpaces As Long
Dim idx As Integer
Dim icwhk As Long
Dim nLdwrds As Byte
Dim SubWrd As String
Dim StrGSpcs As Boolean
Dim FoundOffnsWrd As Boolean
    OffnWrds = "Porn,Hard,Soft,Bisexual,Homosexual,Transsexual,Transvestite,Penis,Cock,Dick,Boob,Tit,Pussy,Cunt,Pork,Sex,Nude,Naked,Babe,Slut,Pump,Filthy,Read,Faggot,Well Endowed,"
        OffnWrds = OffnWrds & "Intercourse,Glamour,Models,Anal,Bust,Spunk,Sperm,Watersports,Escorts,Prostitution,Pimp,Whore,Hor,Contact,Casual,Red,Falacio,Oral,Drinking,"
            OffnWrds = OffnWrds & "Sucking,Adult,Couples,Assertive,Unassertive,Masterbation,Wanking,Vibrator,Toys,Rubber,Dolls,Mature,Fit,Fetish,Orgy,Orgies,Group,Sado Masochist,"
                OffnWrds = OffnWrds & "Reader,Wife,Bitch,Husband,Fuck,Fucking,Fisting,Pump,Bird,Girl,Boy,Men,Male,Women,Female,Fun,Play,Ball,Chick,Swapping,TV,Blow,Job,Jobs,Playmates,"
                    OffnWrds = OffnWrds & "Granny,Fatty,Mature,Ass,Butt,Nipples,Hot,Fantasy,Horny,Lesbian,Whitehouse.com,Hentai,Depraved,Paedo,Core,Erotic,Gangbang,Gangbangs,Animals,"
                        OffnWrds = OffnWrds & "Pornographic,Pornography,Porno,Fallacio,She-male,She-males,Transvestis,Grannies,Hot,Girls,Teen,Teens,Peep,Sale,Masturbate,Stiff,Wood,Orgasm,"
                            OffnWrds = OffnWrds & "Bagnate-sesso Lecco,Pornstar,Hardcore,Action,Amateur,Amateurs,Explicit,Gang Bang,Gang Bangs,Film,Films,Photo,Photos,Dildo,video,Videos,Stud,Studs,"
                                OffnWrds = OffnWrds & "Tart,Tarts,Slag,Slags,Backdoor Access,Horniest,Youngest,Voyeur,Cam,Cams,Gulp,Gulping,"
                                    OffnWrds = OffnWrds & "CyberPorn,Voluptuos,Dirty,Palace,Voyeur,Pervert,Naughty,Disgusting,XXX,Slut,Cum,Filthy,Beastiality,Animals,Peeping,Flasher,Star"
    'This line decides what the measure of illicit content is allowed before the protocols take action - between 1 and 3.
    Randomize
    nLmt = Int(2 - 0 + Rnd * 2)
    'This routine makes sure that the text within the browser is readable.
    If (Asc(Mid(Wordcheck, 1, 1)) = Chr(48) And Asc(Mid(Wordcheck, 2, 1)) = Chr(48)) Or Mid(Wordcheck, 1, 1) = vbNullString Then Exit Function
    If (InStr(Wordcheck, "- Microsoft Internet Explorer") > 0) Then
        'if so then shorten the string to make it more manageable.
        Pos = InStr(Wordcheck, "- Microsoft Internet Explorer")
        Wordcheck = Mid(Wordcheck, 1, Pos - 2)
    End If
    Pos = 0
    Wordcheck = UCase(Wordcheck)
        
    Do While StrGSpcs = False
        pos3 = InStr(pos3 + 1, Wordcheck, Chr(32))
        If (pos3 > 0) Then
            numSpaces = numSpaces + 1
        Else: Exit Do
        End If
    Loop
    'Now go through the passed string and validate all the words within the context."
    SubWrd = Wordcheck
        For idx = 0 To numSpaces + 1
            'Now check to see if the word is recognised. If so then close the browser down.
            nxtSpc = InStr(nxtSpc + 1, SubWrd, Chr(32))
            If (nxtSpc = 0 And nLdwrds < 1) Then
                StringCheck (Wordcheck)
                Exit Function
            End If
            
            If (nxtSpc > 0) Then
                iwchk = Mid(SubWrd, frstSpc + 1, nxtSpc - 1)
                SubWrd = Mid(SubWrd, nxtSpc + 1)
            Else: iwchk = SubWrd
            End If
            
            Do While True
                pos2 = InStr(Pos + 1, OffnWrds, Chr(44), vbTextCompare)
                If (pos2 = 0) Then
                    pos2 = InStr(Pos + 1, OffnWrds)
                    illGlWrd = UCase(Mid(OffnWrds, Pos + 1))
                Else
                    illGlWrd = UCase(Mid(OffnWrds, Pos + 1, (pos2 - 1) - Pos))
                   
                End If
                'This line of code finds out how many spaces there are within the text.
                pos3 = 0
                If (iwchk <> Empty) Then
                'This gives SurfTime the authority to close the browser down.
                    If (nLdwrds = nLmt) Then
                        FoundOffnsWrd = True
                        Exit Do
                    End If
            
                    If (StrComp(iwchk, illGlWrd) = 0) Then
                        nLdwrds = nLdwrds + 1
                        Pos = 0
                        pos2 = 0
                        Exit Do
                    End If
                    If (illGlWrd = "STAR") Then
                        Pos = 0
                        pos2 = 0
                        Exit Do
                    End If
                End If
                Pos = pos2
                DoEvents
            Loop
            frstSpc = 0
            nxtSpc = 0
        Next
        Pos = pos2
        'If (illGlWrd = "STAR") And idx < 2 Then Exit Do
    Stage3_Content_Filter = FoundOffnsWrd

End Function
Private Function StringCheck(ByVal Wordcheck As String)

End Function

Private Function Stage1(ByVal Wordcheck As String)

End Function

Private Function Stage2(ByVal Wordcheck As String)

End Function

Private Function Stage3(ByVal Wordcheck As String)

End Function

