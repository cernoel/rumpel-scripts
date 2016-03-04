' prüft ob im EMF Spool eine spezielle MailAdresse vor kommt
' und speichert diese in eine Variable

Watch.Log "***********************************************************************", 4
Watch.Log "****            START   EMF Mail Adress Parser                     ****", 4
Watch.Log "***********************************************************************", 4


' ============================================================================
'           Variablen festlegen
' ============================================================================
Dim lstrfnInput                ' Dateiname der Input-Datei
Dim loFSO                      ' FileSystem-Objekt für Dateizugriffe
Dim loDataFile                 ' Handle zu einer Datei
Dim lstrfcInputFile            ' Kompletter Inhalt der Datei
Dim lstrArrayInputFile         ' Array der einzelnen Zeilen des Files
Dim MailParser                 ' True while Parsing
Dim MailAdress                 ' Final Mail Adress or "print" if not

' ============================================================================
'      Original-Dateiname+Pfad einlesen
' ============================================================================
lstrfnInput =  watch.getjobfilename


' ============================================================================
'          Datei öffnen und in ein Array einlesen
' ============================================================================
' FileSystemObject-Objekt erzeugen
Set loFSO = CreateObject("Scripting.FileSystemObject")
' Datei öffnen (zum Lesen)
Set loDataFile = loFSO.OpenTextFile(lstrfnInput, 1)
' Inhalt komplett einlesen
lstrfcInputFile = loDataFile.ReadAll
' Datei können wir jetzt schließen
loDataFile.Close

Set loFSO = Nothing
Set loDataFile = Nothing

lstrArrayInputFile = Split(lstrfcInputFile, vbLf)
Set lstrfcInputFile = Nothing


' ============================================================================
'                       Im File nach der Mail suchen
' ============================================================================
' Default Wert:
MailAdress = "print"
tempMailAdress = ""
MailParser = false
gotMail = false

For Each Zeile in lstrArrayInputFile
        dim strZeile
        strZeile = ""
        strZeile = Zeile

        ' Search for Text with color 254, 254, 254
        If(InStr(strZeile,"0.996 0.996 0.996 1 scol")) then
                'Start Parsing
                Watch.Log "   .. Mail Line Found.. parsing.. '" ,4
                MailParser = true
        End If

        If(MailParser) then
                Dim intStart
                intStart = 0
                Dim intEnd
                intEnd = 0
                Dim tempString
                tempString = ""

                ' Watch.log "line: " & strZeile, 4
                If (InStr(strZeile, "(")) then
                        intStart = InStr(strZeile, "(") + 1
                        ' Watch.log "intStart: " & intStart, 4
                        intCount = InStr(strZeile, ")") - intStart
                        ' Watch.log "intCount: " & intCount, 4

                        tempString = Mid(strZeile, intStart, intCount)
                        ' Watch.log "found : " & tempString, 4
                        tempMailAdress = tempMailAdress & tempString
                End If

                If(InStr(strZeile,"0 0 0 1 scol")) then
                        MailParser = false 'end parsing
                        gotMail = true
                        Watch.Log "   .. finished parsing..  '" ,4
                        Exit For
                End If
        End If
Next

If (gotMail) then
        MailAdress = tempMailAdress
End If
Call watch.setvariable("mailAdress", MailAdress)
Watch.Log "   .. result: " & MailAdress ,4
Watch.Log "***********************************************************************", 4
Watch.Log "****             ENDE   EMF Mail Adress Parser                     ****", 4
Watch.Log "***********************************************************************", 4