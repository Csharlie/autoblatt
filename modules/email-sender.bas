Attribute VB_Name = "emailSender"
Option Explicit

' ============================================================
' AUTOBLATT - EMAIL DRAFT KESZITES
' ============================================================
' Letrehoz egy Outlook draft emailt az aktiv munkafuzettel
' mellekletkent. A targy az A6 + C6 + E6 + G8 cellak ertekebol
' all ossze, az alairast az EMAIL_SIGNATURE_MODE config dontes
' alapjan kezeli (Outlook alap / sajat / nincs).
' ============================================================

' --- Publikus belepesi pont ---
Public Sub RunEmailSender()
    Call CreateEmailDraft
End Sub

Private Sub CreateEmailDraft()
    On Error GoTo Fail

    Dim subject As String
    subject = BuildEmailSubject()

    If Len(subject) = 0 Then
        MsgBox "Nem sikerult osszealitani az email targyat!" & vbCrLf & _
               "Ellenorizd a cellakat: " & CELL_SUBJECT_1 & ", " & CELL_SUBJECT_2 & ", " & _
               CELL_SUBJECT_3 & ", " & CELL_SUBJECT_4, vbExclamation, APP_NAME
        Exit Sub
    End If

    Call OpenOutlookDraft(subject)
    Exit Sub

Fail:
    MsgBox "Hiba az email letrehozasa soran: " & Err.Description, vbCritical, APP_NAME
    WriteLog "ERROR CreateEmailDraft: " & Err.Description
End Sub

' --- Targy: a 4 fixed cella szovege szokozzel elvalasztva ---
Private Function BuildEmailSubject() As String
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim parts(0 To 3) As String
    parts(0) = NormalizeSubjectPart(ws.Range(CELL_SUBJECT_1).Value)
    parts(1) = NormalizeSubjectPart(ws.Range(CELL_SUBJECT_2).Value)
    parts(2) = NormalizeSubjectPart(ws.Range(CELL_SUBJECT_3).Value)
    parts(3) = NormalizeSubjectPart(ws.Range(CELL_SUBJECT_4).Value)

    Dim result As String
    Dim i As Long
    For i = 0 To 3
        If Len(parts(i)) > 0 Then
            If Len(result) > 0 Then result = result & " "
            result = result & parts(i)
        End If
    Next i

    BuildEmailSubject = Trim(result)
End Function

Private Function NormalizeSubjectPart(ByVal val As Variant) As String
    If IsStringEmpty(val) Then
        NormalizeSubjectPart = ""
        Exit Function
    End If
    Dim s As String
    s = Trim(CStr(val))
    If LCase(s) = "null" Then s = ""
    NormalizeSubjectPart = s
End Function

' --- Outlook draft letrehozasa es megjelenitese ---
Private Sub OpenOutlookDraft(ByVal subject As String)
    On Error GoTo Fail

    Dim outlookApp As Object
    Dim mail As Object
    Set outlookApp = CreateObject("Outlook.Application")
    Set mail = outlookApp.CreateItem(0) ' 0 = olMailItem

    Dim attachmentPath As String
    attachmentPath = ResolveAttachmentPath()

    With mail
        .Subject = subject
        Call ApplyEmailSignature(mail)
        .Attachments.Add attachmentPath
        .Display
    End With

    WriteLog "INFO Email draft created: " & subject
    Exit Sub

Fail:
    MsgBox "Hiba az Outlook email letrehozasa soran: " & Err.Description & vbCrLf & vbCrLf & _
           "Lehetseges okok:" & vbCrLf & _
           "- Outlook nincs telepitve" & vbCrLf & _
           "- Outlook nincs beallitva" & vbCrLf & _
           "- A munkafuzet nincs mentve (mellekletkent szukseges)", vbCritical, APP_NAME
    WriteLog "ERROR OpenOutlookDraft: " & Err.Description
End Sub

' --- Alairas kezelese a config alapjan ---
' "outlook" - .Display() betolti az Outlook alap alairast (nem irunk a body-ba)
' "custom"  - sajat alairas-szoveg (jelenleg ures, kesobb bovitheto)
' "none"    - ures body
Private Sub ApplyEmailSignature(ByVal mail As Object)
    Select Case LCase(EMAIL_SIGNATURE_MODE)
        Case "outlook"
            ' Nem allitjuk a body-t, igy az Outlook hozzaadja az alap alairast a Display-nel
        Case "custom"
            mail.Body = BuildCustomSignature()
        Case "none"
            mail.Body = ""
        Case Else
            ' Ismeretlen mode - alapertelmezetten az Outlook alap alairas
    End Select
End Sub

' --- Sajat alairas-szoveg (TODO: a user igenye szerint testre szabhato) ---
Private Function BuildCustomSignature() As String
    BuildCustomSignature = vbCrLf & vbCrLf & _
                           "--" & vbCrLf & _
                           Environ("USERNAME") & vbCrLf & _
                           "Mercedes-Benz Manufacturing Hungary"
End Function

' --- Csatolando fajl utvonala (lokalis vagy temp masolat SharePoint URL eseten) ---
Private Function ResolveAttachmentPath() As String
    Dim wbFullName As String
    wbFullName = ActiveWorkbook.FullName

    If InStr(wbFullName, "http") > 0 Then
        ' SharePoint/OneDrive URL - temp masolat
        ResolveAttachmentPath = CreateTempCopyForAttachment()
    Else
        ResolveAttachmentPath = UrlDecode(wbFullName)
    End If
End Function

Private Function CreateTempCopyForAttachment() As String
    On Error GoTo Fail
    Dim tempPath As String
    Dim cleanName As String
    tempPath = Environ("TEMP")
    cleanName = UrlDecode(ActiveWorkbook.Name)
    Dim tempFile As String
    tempFile = tempPath & "\" & cleanName

    ActiveWorkbook.SaveCopyAs tempFile
    CreateTempCopyForAttachment = tempFile
    Exit Function

Fail:
    WriteLog "ERROR CreateTempCopyForAttachment: " & Err.Description
    CreateTempCopyForAttachment = ActiveWorkbook.FullName
End Function
