' Get Mail Adress
' from EMF Spool
'

' rewrite of Spoolfile-Checker script
'
'
'
'

Watch.Log "***********************************************************************", 4
Watch.Log "****            START   Get Mail from Spool                        ****", 4
Watch.Log "***********************************************************************", 4

Dim JobFileName 'String JobFile Name (Spoolfile)
        JobFileName = ""
Dim sJobFile 'String JobFile Data
        sJobFile = ""
Dim sParsingColor 'Special Color
        sParsingColor   = "1 0.996 0.992 1 scol" 'Special EMF Command for Set Color to 255, 254, 253
Dim iParseStartLine ' Parser StartLine
        iParseStartLine = 0
Dim iParseStopLine ' Parser StopLine
        iParseStopLine = 0
Dim sTempMailGarbage 'Garbage .. pre Parsing
        sTempMailGarbage = ""
Dim sMailAdress 'when "PRINT" then no Mail Adress found
        sMailAdress = "PRINT"
Dim sMailStartTag
        sMailStartTag = "#email#"
Dim sMailEndTag
        sMailEndTag = "#/email#"
Dim iEMFRowNum
        iEMFRowNum = 0


' FileSystemObject-Objekt erzeugen
Set oFSO = CreateObject("Scripting.FileSystemObject")
' Dateiname des Spoolfiles auslesen
JobFileName      =  Watch.GetJobFileName
' Datei öffnen (1, zum Lesen)
Set oFileStream  = oFSO.OpenTextFile(JobFileName, 1)
' Inhalt als String komplett einlesen
sJobFile         = oFileStream.ReadAll
' Datei können wir jetzt schließen
oFileStream.Close
' Handles zerstören
Set oFileStream  = Nothing
Set oFSO         = Nothing

' Split Files zu Array Lines
arrJobFile       = Split(sJobFile, chr(10)) 'Split by LineFeed
sJobFile         = "" 'CleanUp

' nach sParsingColor suchen (spezielle Farbe für Parsing)
For ItemIndex = 0 To UBound(arrJobFile)
        If InStr(arrJobFile(ItemIndex), sParsingColor) > 0 Then
                Watch.Log ItemIndex & " .. found ParserColor .."  , 4
                'Watch.Log ItemIndex & " Line: " & arrJobFile(ItemIndex)  , 4
                iParseStartLine = ItemIndex
                Exit For
        End If
Next

' nach ColorChange suchen.. = Ende Parsing
For ItemIndex = iParseStartLine + 1 To UBound(arrJobFile)
        If InStr(arrJobFile(ItemIndex), "scol") > 0 Then
                Watch.Log ItemIndex & " .. found scol, stop parsing there .."  , 4
                'Watch.Log ItemIndex & " Line: " & arrJobFile(ItemIndex)  , 4
                iParseStopLine = ItemIndex - 1
                Exit For
        End If
Next

'Parse Text in Klammern
If iParseStartLine > 0 Then
        For ItemIndex = iParseStartLine To iParseStopLine
                If InStr(arrJobFile(ItemIndex), "(") > 0 Then
                        iStartParse = InStr(arrJobFile(ItemIndex), "(") + 1
                        iStopParse = InStr(arrJobFile(ItemIndex), ")")
                        iLengthParse = iStopParse - iStartParse
                        sTempParse = Mid(arrJobFile(ItemIndex), iStartParse, iLengthParse)
                        sTempMailGarbage = sTempMailGarbage & sTempParse
                        'Watch.Log ItemIndex & " garbage: " & sTempParse , 4
                End If
        Next
End If

Watch.Log "Garbage: " & sTempMailGarbage , 4

' Auslesen der Mail Tags aus dem Garbage
If InStr(sTempMailGarbage, "#email#") > 0 Then
        If InStr(sTempMailGarbage, "#/email#") > 0 Then
                iStartPos = InStr(sTempMailGarbage, "#email#") + 7
                iEndPos = InStr(sTempMailGarbage, "#/email#")
                iLenghtPos = iEndPos - iStartPos
                tempMail = Mid(sTempMailGarbage, iStartPos, iLenghtPos)
                tempMail = Trim(tempMail)
                sMailAdress = tempMail
        End If
End If

Watch.SetVariable "mailAdress", sMailAdress
Watch.Log "Parsed: " & sMailAdress , 4

' 1 Sekunde warten
Watch.Sleep 1000

Watch.Log "***********************************************************************", 4
Watch.Log "****             ENDE    Get Mail from Spool                       ****", 4
Watch.Log "***********************************************************************", 4