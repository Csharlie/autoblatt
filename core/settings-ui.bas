Attribute VB_Name = "settingsUI"
Option Explicit

' ============================================================
' AUTOBLATT - BEALLITASOK PANEL
' ============================================================
' InputBox-alapu menu a felhasznalo-szintu beallitasok
' szerkesztesere. Az ertekek a settings/system-config.json
' fajlba mentodnek el; a config.bas konstansai szolgalnak
' alapertelmezeseknek.
'
' MEGJEGYZES: jovoben ez lecserelheto natív UserForm-ra
' (frmSettings.frm), addig ez biztositja a teljes funkcionalítast.
' ============================================================

' --- Publikus belepesi pont ---
Public Sub OpenSettingsDialog()
    Dim choice As String
    Do
        choice = ShowMainMenu()
        Select Case choice
            Case "1": Call EditPathsSection
            Case "2": Call EditCellsSection
            Case "3": Call EditReplaceRulesSection
            Case "4": Call EditEmailSection
            Case "R": Call ResetToDefaults
            Case "X", "": Exit Do
        End Select
    Loop
End Sub

Private Function ShowMainMenu() As String
    Dim msg As String
    msg = APP_NAME & " v" & APP_VERSION & " - Beallitasok" & vbCrLf & vbCrLf & _
          "1 - Utvonalak (data.xlsm, OneDrive)" & vbCrLf & _
          "2 - Cellanevek (Fotoblatt)" & vbCrLf & _
          "3 - Csere-szabalyok (bpl/sw/li)" & vbCrLf & _
          "4 - Email beallitasok (alairas mode)" & vbCrLf & _
          "R - Visszaallitas alapertelmezesre" & vbCrLf & _
          "X - Kilepes" & vbCrLf & vbCrLf & _
          "Valassz egy opciot:"
    ShowMainMenu = UCase(Trim(InputBox(msg, APP_NAME & " Beallitasok", "X")))
End Function

' --- 1. Utvonalak ---
Private Sub EditPathsSection()
    Dim current As String, newVal As String

    current = GetConfigValue("dataXlsmPath", "")
    newVal = InputBox("data.xlsm utvonala (uresen hagyva: app/data/data.xlsm)" & vbCrLf & _
                      "Jelenlegi: " & IIf(Len(current) = 0, "(alapertelmezes)", current), _
                      APP_NAME, current)
    If newVal <> current Then SaveConfigValue "dataXlsmPath", newVal
End Sub

' --- 2. Cellanevek ---
Private Sub EditCellsSection()
    EditCellSetting "cellUsername", "Felhasznalonev cella", CELL_USERNAME
    EditCellSetting "cellFullname", "Teljes nev cella", CELL_FULLNAME
    EditCellSetting "cellShift", "Muszak cella", CELL_SHIFT
    EditCellSetting "cellDate", "Datum cella", CELL_DATE
    EditCellSetting "cellSubject1", "Email tárgy 1 cella", CELL_SUBJECT_1
    EditCellSetting "cellSubject2", "Email tárgy 2 cella", CELL_SUBJECT_2
    EditCellSetting "cellSubject3", "Email tárgy 3 cella", CELL_SUBJECT_3
    EditCellSetting "cellSubject4", "Email tárgy 4 cella", CELL_SUBJECT_4
End Sub

Private Sub EditCellSetting(ByVal key As String, ByVal label As String, ByVal defaultVal As String)
    Dim current As String, newVal As String
    current = GetConfigValue(key, defaultVal)
    newVal = InputBox(label & " (alapert.: " & defaultVal & ")", APP_NAME, current)
    If newVal <> current Then SaveConfigValue key, newVal
End Sub

' --- 3. Csere-szabalyok ---
Private Sub EditReplaceRulesSection()
    Dim current As String, newVal As String
    current = GetConfigValue("replaceRules", REPLACE_RULES)
    newVal = InputBox("Csere-szabalyok (formatum: kereses=csere;kereses=csere;...)" & vbCrLf & _
                      "Jelenlegi: " & current & vbCrLf & _
                      "Alapertelmezes: " & REPLACE_RULES, _
                      APP_NAME, current)
    If newVal <> current Then SaveConfigValue "replaceRules", newVal
End Sub

' --- 4. Email beallitasok ---
Private Sub EditEmailSection()
    Dim current As String, newVal As String
    current = GetConfigValue("emailSignatureMode", EMAIL_SIGNATURE_MODE)
    newVal = InputBox("Email alairas mode:" & vbCrLf & _
                      "  outlook - Outlook alapertelmezett alairas" & vbCrLf & _
                      "  custom  - sajat alairas" & vbCrLf & _
                      "  none    - nincs alairas" & vbCrLf & vbCrLf & _
                      "Jelenlegi: " & current, _
                      APP_NAME, current)
    If newVal <> current Then SaveConfigValue "emailSignatureMode", LCase(newVal)
End Sub

' --- Reset ---
Private Sub ResetToDefaults()
    If MsgBox("Biztosan visszaallitod a beallitasokat alapertelmezesre?" & vbCrLf & _
              "(A settings/system-config.json fajl ki lesz uritve.)", _
              vbYesNo + vbQuestion, APP_NAME) = vbYes Then
        WriteConfigText "{}"
        MsgBox "Beallitasok visszaallitva.", vbInformation, APP_NAME
    End If
End Sub

' ------------------------------------------------------------
' KONFIG ROGZITES (egyszeru JSON felulirassal)
' ------------------------------------------------------------
' MEGJEGYZES: ez a primitiv jsonkezeles eleg az autoblatt szerinyi
' konfigjahoz (lapos kulcs-ertek parok). Bonyolultabb strukturahoz
' valodi JSON parser kellene.
Private Sub SaveConfigValue(ByVal key As String, ByVal value As String)
    Dim cfg As String
    cfg = ReadConfigText()
    If Len(cfg) = 0 Then cfg = "{}"

    Dim updated As String
    updated = SetJsonValue(cfg, key, value)
    WriteConfigText updated
End Sub

' Egy adott kulcs erteket allitja be vagy hozzaadja.
Private Function SetJsonValue(ByVal jsonText As String, ByVal key As String, ByVal value As String) As String
    Dim escapedValue As String
    escapedValue = Replace(value, "\", "\\")
    escapedValue = Replace(escapedValue, """", "\""")

    Dim newPair As String
    newPair = """" & key & """: """ & escapedValue & """"

    ' Ha letezik a kulcs, csereljuk
    Dim pattern As String
    pattern = """" & key & """"
    Dim startPos As Long
    startPos = InStr(1, jsonText, pattern, vbBinaryCompare)
    If startPos > 0 Then
        ' Megkeressuk az ertek vegét (zaro idezojel)
        Dim valStart As Long, valEnd As Long
        valStart = InStr(startPos + Len(pattern), jsonText, """")
        If valStart > 0 Then
            valEnd = valStart + 1
            Do While valEnd <= Len(jsonText)
                If Mid(jsonText, valEnd, 1) = """" And Mid(jsonText, valEnd - 1, 1) <> "\" Then Exit Do
                valEnd = valEnd + 1
            Loop
            SetJsonValue = Left(jsonText, startPos - 1) & newPair & Mid(jsonText, valEnd + 1)
            Exit Function
        End If
    End If

    ' Uj kulcs hozzaadasa
    Dim trimmed As String
    trimmed = Trim(jsonText)
    If trimmed = "{}" Or Len(trimmed) = 0 Then
        SetJsonValue = "{" & newPair & "}"
    Else
        ' Kulcs hozzaadasa a zaro } ele
        Dim lastBrace As Long
        lastBrace = InStrRev(trimmed, "}")
        If lastBrace > 0 Then
            Dim before As String
            before = Left(trimmed, lastBrace - 1)
            ' Ha mar van benne kulcs, vesszovel valasztunk
            If InStr(before, """") > 0 Then
                SetJsonValue = before & ", " & newPair & "}"
            Else
                SetJsonValue = before & newPair & "}"
            End If
        Else
            SetJsonValue = "{" & newPair & "}"
        End If
    End If
End Function

Private Sub WriteConfigText(ByVal jsonText As String)
    Dim path As String
    path = GetConfigPath()

    ' Mappa biztositasa
    Dim parent As String
    parent = Left(path, InStrRev(path, "\") - 1)
    EnsureFolder parent

    Dim fnum As Integer
    fnum = FreeFile
    Open path For Output As #fnum
    Print #fnum, jsonText
    Close #fnum
End Sub
