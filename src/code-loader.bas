Attribute VB_Name = "modKodFrissites"
' Keszitette es karbantartja : Sardy Peter (PESARDY)
' Mercedes-Benz Manufacturing Hungary, Kecskemet - PW
' Kodfrissites a fejlesztoi forras mappabol (fejlesztoi eszkoz)
' Verzio 1.0.2 - 2026.09.29.
'   1.0.2: uzenetek es megjegyzesek egyszerusitese
'   1.0.1: mappa/fajl letezes GetAttr-ral (a Dir nem latta a rejtett
'          AppData mappat -> "Path/File access error" a mentesnel)
'   1.0: KodFrissitese - a munkafuzet moduljait es lapkodjat lecsereli az
'        src mappaban levo valtozatra. Elotte a regi kodot elmenti.
'
' Hasznalat: Alt+F8 -> KodFrissitese, utana Debug -> Compile VBAProject,
'            majd mentes.
' Feltetel : File -> Options -> Trust Center -> Trust Center Settings ->
'            Macro Settings -> "Trust access to the VBA project object model".
'
' Ugyanez a modul van a fotoblatt.xlsm-ben es az lbl-product-data.xlsm-ben.
' A munkafuzetet a tartalma alapjan ismeri fel:
'   Fotoblatt + Parts lap -> fotoblatt-form-logic.bas, fotoblatt-setup.bas,
'                            fotoblatt-parts-import.bas, es a Fotoblatt lap
'                            kodja: fotoblatt-sheet-events.vba
'   tblTermek tabla       -> lbl-product-data-import.bas
' Sajat magat nem frissiti: ha a code-loader.bas valtozik, kezzel kell
' ujraimportalni (VBE -> modKodFrissites Remove, File -> Import File).
'
' A regi kod mentese: %LOCALAPPDATA%\autoblatt\kodmentes\<munkafuzet>\<ido>\
' (nem a munkafuzet mellett, mert OneDrive-on a Workbook.Path URL is lehet).
'
' A kod Sardy Peter munkaja. Modositas elott egyeztess vele.

Option Explicit

Private Const KF_CIM         As String = "Kodfrissites"
Private Const KF_NEV_SRC     As String = "KodForrasUtvonal"      ' rejtett nev: a valasztott src mappa
Private Const KF_SRC_ALAP    As String = "\Projects\autoblatt\src" ' %USERPROFILE% alatt
Private Const KF_JELZO_FAJL  As String = "code-loader.bas"         ' ebbol ismeri fel a forras mappat
Private Const KF_MENTES      As String = "\autoblatt\kodmentes"    ' %LOCALAPPDATA% alatt

Private Const KF_MODUL       As String = "M"    ' standard modul (.bas)
Private Const KF_LAPKOD      As String = "L"    ' lap kodja (.vba, beillesztendo szoveg)

Private Const KF_CT_STD      As Long = 1        ' vbext_ct_StdModule
Private Const KF_CT_DOC      As Long = 100      ' vbext_ct_Document

'==========================================================================
' Belepesi pont
'==========================================================================
Public Sub KodFrissitese()

    Dim wb As Workbook
    Dim fajlok() As String, nevek() As String, tipusok() As String
    Dim n As Long, i As Long
    Dim srcUt As String, mentesUt As String
    Dim hiba As String, lista As String, kesz As String, lepes As String

    Set wb = ThisWorkbook

    lepes = "VBA projekt elerese"
    If Not KfVbProjektElerheto(wb, hiba) Then
        MsgBox hiba, vbExclamation, KF_CIM
        Exit Sub
    End If

    lepes = "munkafuzet felismerese"
    n = KfListaOsszeallitas(wb, fajlok, nevek, tipusok, hiba)
    If n = 0 Then
        MsgBox hiba, vbExclamation, KF_CIM
        Exit Sub
    End If

    lepes = "src mappa"
    srcUt = KfSrcMappa()
    If Len(srcUt) = 0 Then Exit Sub

    ' 1. Minden forrasfajl ellenorzese, mielott barmihez nyulnank
    lepes = "forrasfajlok ellenorzese"
    For i = 1 To n
        hiba = KfFajlEllenorzes(srcUt & "\" & fajlok(i), nevek(i), tipusok(i))
        If Len(hiba) > 0 Then
            MsgBox "Nem csereltem semmit." & vbLf & vbLf & hiba, vbExclamation, KF_CIM
            Exit Sub
        End If
        lista = lista & "  " & fajlok(i) & "  ->  " & nevek(i) & _
                "  (" & KfVerzioSor(srcUt & "\" & fajlok(i)) & ")" & vbLf
    Next i

    If MsgBox("Munkafuzet: " & wb.Name & vbLf & _
              "Forras:     " & srcUt & vbLf & vbLf & _
              "Lecserelendo kod:" & vbLf & lista & vbLf & _
              "A regi kodot elotte elmentem. Mehet?", _
              vbQuestion + vbYesNo + vbDefaultButton2, KF_CIM) <> vbYes Then Exit Sub

    On Error GoTo Hiba

    ' 2. A regi kod mentese
    lepes = "regi kod mentese"
    mentesUt = KfMentes(wb, nevek, tipusok, n)

    ' 3. Csere
    For i = 1 To n
        lepes = "csere: " & nevek(i)
        If tipusok(i) = KF_MODUL Then
            KfModulCsere wb, nevek(i), srcUt & "\" & fajlok(i)
        Else
            KfLapkodCsere wb, nevek(i), srcUt & "\" & fajlok(i)
        End If
        kesz = kesz & "  " & fajlok(i) & vbLf
    Next i

    On Error GoTo 0
    MsgBox "Kesz. Lecserelve:" & vbLf & kesz & vbLf & _
           "A regi kod mentese:" & vbLf & "  " & mentesUt & vbLf & vbLf & _
           "Kovetkezo lepes: Debug -> Compile VBAProject, majd mentes.", _
           vbInformation, KF_CIM
    Exit Sub

Hiba:
    MsgBox "A frissites megszakadt [" & lepes & "]:" & vbLf & Err.Description & vbLf & vbLf & _
           "Mar lecserelve:" & vbLf & IIf(Len(kesz) > 0, kesz, "  (semmi)" & vbLf) & vbLf & _
           "A regi kod mentese:" & vbLf & "  " & IIf(Len(mentesUt) > 0, mentesUt, "(nem keszult)"), _
           vbCritical, KF_CIM

End Sub

'==========================================================================
' Melyik fajl melyik komponensbe megy
'==========================================================================
Private Function KfListaOsszeallitas(ByVal wb As Workbook, ByRef fajlok() As String, _
                                     ByRef nevek() As String, ByRef tipusok() As String, _
                                     ByRef hiba As String) As Long

    Dim lapKomp As String

    If KfLapLetezik(wb, "Fotoblatt") And KfLapLetezik(wb, "Parts") Then

        lapKomp = KfLapKomponens(wb, "Fotoblatt")
        If Len(lapKomp) = 0 Then
            hiba = "A Fotoblatt lap kodmodulja nem talalhato."
            Exit Function
        End If

        ReDim fajlok(1 To 4): ReDim nevek(1 To 4): ReDim tipusok(1 To 4)
        fajlok(1) = "fotoblatt-form-logic.bas":   nevek(1) = "modFotolap":      tipusok(1) = KF_MODUL
        fajlok(2) = "fotoblatt-setup.bas":        nevek(2) = "modBeallitas":    tipusok(2) = KF_MODUL
        fajlok(3) = "fotoblatt-parts-import.bas": nevek(3) = "modTermekadatok": tipusok(3) = KF_MODUL
        fajlok(4) = "fotoblatt-sheet-events.vba": nevek(4) = lapKomp:           tipusok(4) = KF_LAPKOD
        KfListaOsszeallitas = 4

    ElseIf KfTablaLetezik(wb, "tblTermek") Then

        ReDim fajlok(1 To 1): ReDim nevek(1 To 1): ReDim tipusok(1 To 1)
        fajlok(1) = "lbl-product-data-import.bas": nevek(1) = "modTermekadatok": tipusok(1) = KF_MODUL
        KfListaOsszeallitas = 1

    Else
        hiba = "Ismeretlen munkafuzet: " & wb.Name & vbLf & _
               "Nem Fotoblatt (Fotoblatt + Parts lap) es nem LBL Product Data (tblTermek tabla)."
    End If

End Function

'==========================================================================
' Ellenorzesek
'==========================================================================
Private Function KfVbProjektElerheto(ByVal wb As Workbook, ByRef hiba As String) As Boolean

    Dim db As Long

    On Error Resume Next
    db = wb.VBProject.VBComponents.Count
    If Err.Number <> 0 Then
        hiba = "A VBA projekt nem erheto el." & vbLf & vbLf & _
               "Engedelyezd egyszer: File -> Options -> Trust Center -> Trust Center Settings ->" & vbLf & _
               "Macro Settings -> 'Trust access to the VBA project object model'."
        Exit Function
    End If
    On Error GoTo 0

    If wb.VBProject.Protection = 1 Then
        hiba = "A VBA projekt jelszoval vedett. Elobb old fel (VBE -> Tools -> VBAProject Properties -> Protection)."
        Exit Function
    End If

    KfVbProjektElerheto = True

End Function

' Ures string = rendben, egyebkent a hiba szovege.
Private Function KfFajlEllenorzes(ByVal ut As String, ByVal vartNev As String, _
                                  ByVal tipus As String) As String

    Dim b() As Byte, i As Long, txt As String

    If Not KfFajlLetezik(ut) Then
        KfFajlEllenorzes = "Hianyzik: " & ut
        Exit Function
    End If
    If FileLen(ut) = 0 Then
        KfFajlEllenorzes = "Ures fajl: " & ut
        Exit Function
    End If

    b = KfBajtok(ut)
    For i = LBound(b) To UBound(b)
        If b(i) > 127 Then
            KfFajlEllenorzes = "Nem ASCII karakter a(z) " & (i + 1) & ". bajtnal:" & vbLf & ut & vbLf & _
                               "(Szabaly: minden .bas/.vba csak ASCII, ekezet ChrW()-vel.)"
            Exit Function
        End If
    Next i

    txt = StrConv(b, vbUnicode)
    If tipus = KF_MODUL Then
        If InStr(1, txt, "Attribute VB_Name = """ & vartNev & """", vbBinaryCompare) = 0 Then
            KfFajlEllenorzes = "A fajl VB-modulneve nem " & vartNev & ":" & vbLf & ut
        End If
    Else
        If InStr(1, txt, "Attribute VB_", vbBinaryCompare) > 0 Then
            KfFajlEllenorzes = "A lapkod-fajlban nem lehet 'Attribute' sor:" & vbLf & ut
        End If
    End If

End Function

' A fejlec "' Verzio ..." sora, pl. "Verzio 9.2 - 2026.09.22."
Private Function KfVerzioSor(ByVal ut As String) As String

    Dim txt As String, p As Long, q As Long

    txt = StrConv(KfBajtok(ut), vbUnicode)
    p = InStr(1, txt, "' Verzio ", vbBinaryCompare)
    If p = 0 Then
        KfVerzioSor = "verzio ?"
        Exit Function
    End If
    q = InStr(p, txt, vbCr)
    If q = 0 Then q = InStr(p, txt, vbLf)
    If q = 0 Then q = Len(txt) + 1
    KfVerzioSor = Trim$(Mid$(txt, p + 2, q - p - 2))

End Function

'==========================================================================
' Mentes es csere
'==========================================================================
Private Function KfMentes(ByVal wb As Workbook, ByRef nevek() As String, _
                          ByRef tipusok() As String, ByVal n As Long) As String

    Dim ut As String, i As Long, komp As Object, cm As Object, f As Integer

    ut = Environ$("LOCALAPPDATA") & KF_MENTES & "\" & KfFajlnevBiztos(wb.Name) & "\" & _
         Format$(Now, "yyyymmdd-hhnnss")
    KfMappaLetrehozas ut

    For i = 1 To n
        Set komp = Nothing
        On Error Resume Next
        Set komp = wb.VBProject.VBComponents(nevek(i))
        On Error GoTo 0
        If Not komp Is Nothing Then
            If tipusok(i) = KF_MODUL Then
                komp.Export ut & "\" & nevek(i) & ".bas"
            Else
                Set cm = komp.CodeModule
                f = FreeFile
                Open ut & "\" & nevek(i) & ".vba" For Output As #f
                If cm.CountOfLines > 0 Then Print #f, cm.Lines(1, cm.CountOfLines)
                Close #f
            End If
        End If
    Next i

    KfMentes = ut

End Function

' A regi modult elobb atnevezi, csak utana torli: a VBE a torlest a makro
' vegeig kesleltetheti, es akkor az import "modFotolap1" nevet kapna.
Private Sub KfModulCsere(ByVal wb As Workbook, ByVal nev As String, ByVal ut As String)

    Dim vbc As Object, regi As Object, uj As Object

    Set vbc = wb.VBProject.VBComponents

    On Error Resume Next
    Set regi = vbc(nev)
    On Error GoTo 0

    If Not regi Is Nothing Then
        If regi.Type <> KF_CT_STD Then
            Err.Raise vbObjectError + 513, , nev & " nem standard modul, nem cserelem."
        End If
        regi.Name = KfSzabadNev(vbc, nev & "_old")
        vbc.Remove regi
    End If

    Set uj = vbc.Import(ut)
    If uj.Name <> nev Then
        Err.Raise vbObjectError + 514, , "Az importalt modul neve " & uj.Name & " lett " & nev & " helyett."
    End If

End Sub

Private Sub KfLapkodCsere(ByVal wb As Workbook, ByVal nev As String, ByVal ut As String)

    Dim cm As Object

    Set cm = wb.VBProject.VBComponents(nev).CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString StrConv(KfBajtok(ut), vbUnicode)

End Sub

'==========================================================================
' src mappa: rejtett nevben megjegyezve, alapbol %USERPROFILE%\Projects\autoblatt\src
'==========================================================================
Private Function KfSrcMappa() As String

    Dim ut As String

    ut = KfNevOlvasas(KF_NEV_SRC)
    If Not KfSrcRendben(ut) Then ut = Environ$("USERPROFILE") & KF_SRC_ALAP
    If KfSrcRendben(ut) Then
        KfSrcMappa = ut
        Exit Function
    End If

    With Application.FileDialog(4)          ' msoFileDialogFolderPicker
        .Title = "A forras mappa kivalasztasa"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Function
        ut = .SelectedItems(1)
    End With

    If Not KfSrcRendben(ut) Then
        MsgBox "Ebben a mappaban nincs " & KF_JELZO_FAJL & ", ez nem a forras mappa:" & vbLf & ut, _
               vbExclamation, KF_CIM
        Exit Function
    End If

    KfNevMentes KF_NEV_SRC, ut
    KfSrcMappa = ut

End Function

Private Function KfSrcRendben(ByVal ut As String) As Boolean
    If Len(ut) = 0 Then Exit Function
    If Right$(ut, 1) = "\" Then ut = Left$(ut, Len(ut) - 1)
    KfSrcRendben = KfFajlLetezik(ut & "\" & KF_JELZO_FAJL)
End Function

Private Sub KfNevMentes(ByVal nev As String, ByVal ertek As String)
    On Error Resume Next
    ThisWorkbook.Names(nev).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nev, _
        RefersTo:="=""" & Replace(ertek, """", """""") & """", Visible:=False
End Sub

Private Function KfNevOlvasas(ByVal nev As String) As String
    Dim s As String
    On Error Resume Next
    s = ThisWorkbook.Names(nev).RefersTo            ' ="C:\...\src"
    On Error GoTo 0
    If Len(s) > 3 Then
        s = Mid$(s, 3, Len(s) - 3)
        KfNevOlvasas = Replace(s, """""", """")
    End If
End Function

'==========================================================================
' Segedfuggvenyek
'==========================================================================
Private Function KfLapLetezik(ByVal wb As Workbook, ByVal nev As String) As Boolean
    Dim ws As Object
    On Error Resume Next
    Set ws = wb.Worksheets(nev)
    On Error GoTo 0
    KfLapLetezik = Not ws Is Nothing
End Function

Private Function KfTablaLetezik(ByVal wb As Workbook, ByVal nev As String) As Boolean
    Dim ws As Worksheet, lo As ListObject
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, nev, vbTextCompare) = 0 Then
                KfTablaLetezik = True
                Exit Function
            End If
        Next lo
    Next ws
End Function

' A lap (fulnev) kodmodul-komponensenek neve, pl. "Fotoblatt" -> "shFotoblatt".
Private Function KfLapKomponens(ByVal wb As Workbook, ByVal lapNev As String) As String
    Dim komp As Object, nev As String
    For Each komp In wb.VBProject.VBComponents
        If komp.Type = KF_CT_DOC Then
            nev = ""
            On Error Resume Next
            nev = komp.Properties("Name").Value
            On Error GoTo 0
            If StrComp(nev, lapNev, vbTextCompare) = 0 Then
                KfLapKomponens = komp.Name
                Exit Function
            End If
        End If
    Next komp
End Function

Private Function KfSzabadNev(ByVal vbc As Object, ByVal alap As String) As String
    Dim jelolt As String, i As Long, komp As Object
    jelolt = alap
    Do
        Set komp = Nothing
        On Error Resume Next
        Set komp = vbc(jelolt)
        On Error GoTo 0
        If komp Is Nothing Then Exit Do
        i = i + 1
        jelolt = alap & i
    Loop
    KfSzabadNev = jelolt
End Function

Private Function KfBajtok(ByVal ut As String) As Byte()
    Dim f As Integer, b() As Byte
    f = FreeFile
    Open ut For Binary Access Read As #f
    If LOF(f) > 0 Then
        ReDim b(0 To LOF(f) - 1)
        Get #f, , b
    End If
    Close #f
    KfBajtok = b
End Function

' GetAttr-ral, nem Dir-rel: a Dir a rejtett mappakat (pl. AppData) nem latja.
Private Function KfFajlLetezik(ByVal ut As String) As Boolean
    Dim a As Long
    On Error Resume Next
    a = GetAttr(ut)
    If Err.Number = 0 Then KfFajlLetezik = ((a And vbDirectory) = 0)
End Function

Private Function KfMappaLetezik(ByVal ut As String) As Boolean
    Dim a As Long
    On Error Resume Next
    a = GetAttr(ut)
    If Err.Number = 0 Then KfMappaLetezik = ((a And vbDirectory) <> 0)
End Function

Private Sub KfMappaLetrehozas(ByVal ut As String)
    Dim reszek() As String, i As Long, akt As String
    reszek = Split(ut, "\")
    akt = reszek(0)
    For i = 1 To UBound(reszek)
        akt = akt & "\" & reszek(i)
        If Not KfMappaLetezik(akt) Then MkDir akt
    Next i
End Sub

Private Function KfFajlnevBiztos(ByVal s As String) As String
    Dim c As Variant
    For Each c In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, c, "_")
    Next c
    KfFajlnevBiztos = s
End Function
