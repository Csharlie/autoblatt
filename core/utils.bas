Attribute VB_Name = "utils"
Option Explicit

' ============================================================
' AUTOBLATT - KOZOS SEGEDFUGGVENYEK
' ============================================================
' String-tisztitas, JSON, log, mappa-kezeles, Excel utility-k.
' Minden modul ezeken keresztul fer hozza a kozos funkcionalitashoz.
' ============================================================

' ------------------------------------------------------------
' STRING TISZTITAS
' ------------------------------------------------------------

' Eltavolitja az ekezeteket es a fajlnevben tilos karaktereket.
' Hasznalat: filenev / mappanev epitese cellaerterkbol.
Public Function SanitizeForFilename(ByVal text As String) As String
    Dim i As Long
    Dim c As String
    Dim result As String
    Dim code As Long

    For i = 1 To Len(text)
        c = Mid(text, i, 1)
        code = AscW(c)

        Select Case code
            Case 225, 193:  result = result & IIf(code = 225, "a", "A")
            Case 233, 201:  result = result & IIf(code = 233, "e", "E")
            Case 237, 205:  result = result & IIf(code = 237, "i", "I")
            Case 243, 211:  result = result & IIf(code = 243, "o", "O")
            Case 246, 214:  result = result & IIf(code = 246, "o", "O")
            Case 337, 336:  result = result & IIf(code = 337, "o", "O")
            Case 250, 218:  result = result & IIf(code = 250, "u", "U")
            Case 252, 220:  result = result & IIf(code = 252, "u", "U")
            Case 369, 368:  result = result & IIf(code = 369, "u", "U")
            Case 92, 47, 58, 42, 63, 34, 60, 62, 124
                ' Tiltott fajlnevkarakterek: \ / : * ? " < > | -- atugorva
            Case Else
                If c Like "[A-Za-z0-9 _-]" Then result = result & c
        End Select
    Next i

    SanitizeForFilename = result
End Function

' URL-dekodolas (%XX -> karakter). SharePoint URL-ekben hasznos.
Public Function UrlDecode(ByVal encodedText As String) As String
    Dim result As String
    Dim i As Long
    Dim ch As String
    Dim hexCode As String

    result = encodedText
    ' Plusz jelek -> szokoz
    result = Replace(result, "+", " ")

    ' %XX hexadecimalis karakterek dekodolasa
    i = 1
    Dim sb As String
    sb = ""
    Do While i <= Len(result)
        ch = Mid(result, i, 1)
        If ch = "%" And i + 2 <= Len(result) Then
            hexCode = Mid(result, i + 1, 2)
            If hexCode Like "[0-9A-Fa-f][0-9A-Fa-f]" Then
                sb = sb & Chr(CLng("&H" & hexCode))
                i = i + 3
            Else
                sb = sb & ch
                i = i + 1
            End If
        Else
            sb = sb & ch
            i = i + 1
        End If
    Loop

    UrlDecode = sb
End Function

' Igaz, ha a string Null, Empty, vagy csak whitespace.
Public Function IsStringEmpty(ByVal s As Variant) As Boolean
    If IsNull(s) Or IsEmpty(s) Then
        IsStringEmpty = True
    ElseIf VarType(s) = vbString Then
        IsStringEmpty = (Len(Trim$(CStr(s))) = 0)
    Else
        IsStringEmpty = False
    End If
End Function

' ------------------------------------------------------------
' UTVONALAK ES MAPPAK
' ------------------------------------------------------------

' OneDrive (Mercedes) gyokerelerese az Environ-bol.
' Visszaadja az ures stringet, ha nem talalhato.
Public Function GetOneDrivePath() As String
    Dim p As String
    p = Environ("OneDriveCommercial")
    If Len(p) = 0 Then p = Environ("OneDrive")
    If Len(p) = 0 Then
        ' Fallback: feltetelezett ut a felhasznalonev alapjan
        p = "C:\Users\" & Environ("USERNAME") & "\OneDrive - Mercedes-Benz (corpdir.onmicrosoft.com)"
    End If
    GetOneDrivePath = p
End Function

' A futtato workbook fizikai mappaja (Documents\autoblatt vagy OneDrive\Autoblatt\app).
Public Function GetWorkbookPath() As String
    GetWorkbookPath = ThisWorkbook.Path
End Function

' Az autoblatt projekt forras-mappaja (ahonnan a .bas modulok jonnek).
' Sorrend:
'   1. config "installPath" mezoje (ha letezik es ervenyes)
'   2. ThisWorkbook.Path (fejlesztoi mod, ahol az xlsm van)
Public Function GetAppRootPath() As String
    Dim configured As String
    configured = ReadInstallPathFromConfig()
    If Len(configured) > 0 Then
        ' Ervenyesseg ellenorzes: letezik-e a core mappa?
        If Dir(configured & "\core", vbDirectory) <> "" Then
            GetAppRootPath = configured
            Exit Function
        End If
    End If
    GetAppRootPath = ThisWorkbook.Path
End Function

' Beolvassa az installPath-ot a felhasznalo-szintu config fajlbol.
' Ez nem fugg az ertekek elhelyezkedesetol - mindig a %LOCALAPPDATA%-bol.
Private Function ReadInstallPathFromConfig() As String
    Dim path As String
    path = GetConfigPath()
    If Dir(path) = "" Then
        ReadInstallPathFromConfig = ""
        Exit Function
    End If

    Dim fnum As Integer
    Dim content As String
    fnum = FreeFile
    Open path For Input As #fnum
    Do While Not EOF(fnum)
        Dim line As String
        Line Input #fnum, line
        content = content & line & vbLf
    Loop
    Close #fnum

    Dim val As String
    val = GetJsonString(content, "installPath")
    ' Az escape-elt backslash visszaalakitasa (\\\\  ->  \)
    val = Replace(val, "\\", "\")
    ' Trailing backslash levagas
    If Right(val, 1) = "\" Then val = Left(val, Len(val) - 1)
    ReadInstallPathFromConfig = val
End Function

' Letrehoz egy mappat es minden hianyzo szuloket. Idempotens.
Public Sub EnsureFolder(ByVal folderPath As String)
    If Len(folderPath) = 0 Then Exit Sub
    If Dir(folderPath, vbDirectory) <> "" Then Exit Sub

    Dim parent As String
    parent = Left(folderPath, InStrRev(folderPath, "\") - 1)
    If Len(parent) > 0 And parent <> folderPath Then EnsureFolder parent

    On Error Resume Next
    MkDir folderPath
    On Error GoTo 0
End Sub

' system-config.json utvonala. Felhasznalo-szintu hely (per Windows-user):
'   %LOCALAPPDATA%\Autoblatt\system-config.json
' Itt elnevezesi konflikus nem lehetseges, es a workbook helye nem fugg ettol.
Public Function GetConfigPath() As String
    GetConfigPath = Environ("LOCALAPPDATA") & "\" & APP_NAME & "\system-config.json"
End Function

' ------------------------------------------------------------
' EXCEL UTILITY-K
' ------------------------------------------------------------

' Idempotens lap-lekero: visszaadja a lapot, vagy letrehozza ha nincs.
Public Function GetOrCreateWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = sheetName
    End If
    Set GetOrCreateWorksheet = ws
End Function

' Inline (vesszovel elvalasztott) adaterveny. lista hozzaadasa egy cellahoz.
Public Sub ApplyDropdown(ByVal targetCell As Range, ByVal listValues As String)
    With targetCell.Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:=listValues
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = False
    End With
End Sub

' ------------------------------------------------------------
' JSON - MINIMALIS KULCS-KERESO
' ------------------------------------------------------------
' Nem teljes ertekeu JSON parser, de robusztusabb a string Replace-nel.
' Egy adott kulcs string ertekenek visszaadasa egy JSON szovegbol.
' Pelda: GetJsonString("{""installPath"":""C:\\foo""}", "installPath")
Public Function GetJsonString(ByVal jsonText As String, ByVal key As String) As String
    Dim pattern As String
    Dim startPos As Long, endPos As Long, valStart As Long
    Dim result As String

    pattern = """" & key & """"
    startPos = InStr(1, jsonText, pattern, vbBinaryCompare)
    If startPos = 0 Then
        GetJsonString = ""
        Exit Function
    End If

    ' Kulcs utan keressuk az elso "-t (ertekkezdet)
    valStart = InStr(startPos + Len(pattern), jsonText, """")
    If valStart = 0 Then
        GetJsonString = ""
        Exit Function
    End If

    ' Az ertek vege a kovetkezo nem-escapelt "
    Dim i As Long
    i = valStart + 1
    Do While i <= Len(jsonText)
        If Mid(jsonText, i, 1) = """" And Mid(jsonText, i - 1, 1) <> "\" Then
            endPos = i
            Exit Do
        End If
        i = i + 1
    Loop

    If endPos > valStart Then
        result = Mid(jsonText, valStart + 1, endPos - valStart - 1)
        ' Egyszeru escape feloldas
        result = Replace(result, "\\", "\")
        result = Replace(result, "\""", """")
        result = Replace(result, "\n", vbLf)
        result = Replace(result, "\t", vbTab)
        GetJsonString = result
    Else
        GetJsonString = ""
    End If
End Function

' Beolvassa a teljes JSON config fajlt egy stringbe. Ures stringet ad ha nem letezik.
Public Function ReadConfigText() As String
    Dim path As String
    path = GetConfigPath()
    If Dir(path) = "" Then
        ReadConfigText = ""
        Exit Function
    End If

    Dim fnum As Integer
    Dim content As String
    fnum = FreeFile
    Open path For Input As #fnum
    Do While Not EOF(fnum)
        Dim line As String
        Line Input #fnum, line
        content = content & line & vbLf
    Loop
    Close #fnum
    ReadConfigText = content
End Function

' Konfigbol kiolvasott string ertek, vagy fallback ha nincs.
Public Function GetConfigValue(ByVal key As String, ByVal fallback As String) As String
    Dim cfg As String
    cfg = ReadConfigText()
    If Len(cfg) = 0 Then
        GetConfigValue = fallback
        Exit Function
    End If
    Dim val As String
    val = GetJsonString(cfg, key)
    If Len(val) = 0 Then
        GetConfigValue = fallback
    Else
        GetConfigValue = val
    End If
End Function

' ------------------------------------------------------------
' CSERE-SZABALYOK FELDOLGOZASA
' ------------------------------------------------------------
' "bpl=BPL;sw=SW;li=Li" -> alkalmazva egy szovegre case-insensitive.
Public Function ApplyReplaceRules(ByVal text As String, ByVal rulesString As String) As String
    If Len(rulesString) = 0 Then
        ApplyReplaceRules = text
        Exit Function
    End If

    Dim rules() As String
    Dim rule As Variant
    Dim parts() As String
    Dim result As String

    result = text
    rules = Split(rulesString, ";")
    For Each rule In rules
        If Len(rule) > 0 And InStr(rule, "=") > 0 Then
            parts = Split(rule, "=")
            If UBound(parts) >= 1 Then
                If InStr(1, result, parts(0), vbTextCompare) > 0 Then
                    result = Replace(result, parts(0), parts(1), , , vbTextCompare)
                End If
            End If
        End If
    Next rule
    ApplyReplaceRules = result
End Function

' ------------------------------------------------------------
' LOGOLAS
' ------------------------------------------------------------
' UTF-8 logfajl rotacioval. Cel: %LOCALAPPDATA%\Autoblatt\autoblatt.log
Public Sub WriteLog(ByVal message As String)
    On Error Resume Next
    Dim logDir As String
    Dim logPath As String

    logDir = Environ("LOCALAPPDATA") & "\" & APP_NAME
    EnsureFolder logDir
    logPath = logDir & "\" & LCase(APP_NAME) & ".log"

    ' Rotacio: ha tul nagy, atnevezzuk .1.log-ra
    On Error Resume Next
    If Dir(logPath) <> "" Then
        If FileLen(logPath) > LOG_MAX_BYTES Then
            If Dir(logPath & ".1") <> "" Then Kill logPath & ".1"
            Name logPath As logPath & ".1"
        End If
    End If
    On Error GoTo 0

    Dim fnum As Integer
    fnum = FreeFile
    Open logPath For Append As #fnum
    Print #fnum, Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & message
    Close #fnum
End Sub
