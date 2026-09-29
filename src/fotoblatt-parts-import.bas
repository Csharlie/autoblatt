Attribute VB_Name = "modTermekadatok"
' Keszitette es karbantartja : Sardy Peter (PESARDY)
' Mercedes-Benz Manufacturing Hungary, Kecskemet - PW
' Verzio 9.13 - 2026.09.24.
'   9.13: R-cikkszam kivalasztasakor az E4 fejlec "Alapanyag/Material"-ra
'        valt (eredetileg "Preseles/Abpressung"), es az E6 cellaba
'        "Platine" kerul. Minden mas esetben (pl. A-cikkszam) az E4
'        visszaall az eredeti "Preseles/Abpressung" feliratra, es ha az
'        E6-ban meg a korabban automatikusan beirt "Platine" all, azt
'        toroljuk (a felhasznalo sajat E6 bejegyzeset nem bantjuk).
'   9.12: ha a kivalasztott/beirt cikkszam A-val kezdodik, a gyarto
'        berendezes (H6) mindig 7.41-re all be - fuggetlenul attol, hogy
'        az Article tablaban van-e sajat Berendezes ertek. (R-szamnal
'        tovabbra is a tetel sajat Berendezes erteke szamit, l. 9.5.)
'   9.11: az A6 melletti legordulo ismet EGYETLEN, tiszta lista (mint
'        v9.9-ben) - a v9.7-9.10-es "kereso-segedsor" trukkok (technikai
'        elotagok) VEGLEGESEN visszavonva. A6-ba KOZVETLENUL gepelve a
'        cikkszam elejetol szur (szokoz/nagybetu nem szamit - lasd 9.6).
'        Barmiben (szamban VAGY szoban) valo keresesre a legordulo NYIL-ra
'        kattintva megjeleno, Excel sajat beepitett keresomezoje hasznalando
'        - ez ugyanaz a legordulo, csak a nyillal nyitva meg egy keresosav
'        is van benne, nincs kulon ablak/popup.
'        Uj, ujrafelhasznalhato fuggveny: CikkszamBeallitasa(ws, kulcs) -
'        a lenyeg innentol ugyanaz, mint korabban volt.
'   9.6: a Lista oszlop cikkszam resze SZOKOZ NELKUL kerul be (SUBSTITUTE
'        a keplettben), hogy az Excel legordulo gepeles-kozbeni szurese
'        akkor is talaljon, ha a felhasznalo szokoz nelkul (es barmilyen
'        betumerettel) gepeli be a cikkszamot, pl. "a24463701" -> talalja
'        "A244 637 0100"-t. Kivalasztas utan A6-ba mindig a valodi PartNum
'        kerul (a Parts!Article-ben tarolt formaban, szokozokkel egyutt).
'   9.5: az A6 listaban MINDEN cikkszam (A- es R-szam, aktiv es kifutott
'        is) MINDIG valaszthato, fuggetlenul H6-tol ES G6-tol - nincs se
'        Berendezes-, se Tipus-alapu szures. A cikkszam kivalasztasakor/
'        beirasakor a H6 (gyarto berendezes) automatikusan beallitodik a
'        tetel sajat Berendezes ertekere, ha az elter a jelenlegitol
'        (CikkszamValasztas). Igy pl. R-szam valasztasakor H6 automatikusan
'        7.50-re ugrik. G6 (tipus) tovabbra is kitoltodik a tetelbol, de
'        mar nem szukiti az A6 listat.
'   9.3: a forras MOST MAR a sajat LBL_Termekadatok_sajat.xlsm tblTermek
'        tablaja (nev szerint, ha nincs, fejlec alapjan) - nem az eredeti
'        TERM_KADATOK nyers fajl tobbe. Ez a dokumentalt A->B interfeszt
'        koveti (lasd HANDOFF). A Berendezes erteket is a sajat tablabol
'        vesszuk at (oszlop hianyaban 7.50-et feltetelezunk).
'
' A kod Sardy Peter munkaja. Modositas elott egyeztess vele.

Option Explicit
Option Compare Text

'==========================================================================
' Termekadatok (LBL 7.50) -> Parts!Article, es az A6 cikkszam-legordulo
'--------------------------------------------------------------------------
' Parts!Article oszlopai (a nev szamit, a sorrend nem):
'   PartNum    - cikkszam (A/N: szokoz nelkul, R: ahogy a sajat fajlban all)
'   PartName   - megnevezes
'   Type       - tipus (R-szamnal a sajat tblTermek 'Tipus' oszlopabol)
'   Forras     - A = alkatresz-torzs, R = sajat termekadatok (a frissites kezeli)
'   Statusz    - Aktiv / Kifutott (ures = aktiv)
'   Berendezes - R-szamnal a sajat tblTermek 'Berendezes' oszlopabol (7.50/7.41/7.22),
'                hianyaban 7.50
'   Ellenorzes - tisztitasi megjegyzes (utkozo megnevezes, hianyzo adat)
'   Lista      - keplet: "Cikkszam . Megnevezes . Tipus (kifutott)"
'
' Forras: a sajat LBL_Termekadatok_sajat.xlsm tblTermek tablaja.
'   - Termekszam -> PartNum (ahogy ott all, nem formazzuk ujra)
'   - Termeknev  -> PartName
'   - Tipus      -> Type (uj kodot automatikusan felvesz a Type tablaba)
'   - Berendezes -> Berendezes (oszlop hianyaban 7.50, figyelmeztetessel)
'   - Statusz    -> "Kifutott" = jelolve; "...eltunt a forrasbol..." = kimarad;
'                   minden mas (pl. "Aktiv", "Aktiv es kifutott - ellenorizendo") = aktiv
' Ha nincs 'tblTermek' nevu tabla a fajlban, fejlec alapjan keresunk
' (Termekszam/Cikkszam, Termeknev/Megnevezes, Tipus, Berendezes, Statusz).
'
' Az A6 legordulo forrasa az A6Lista nev (Parts!J oszlop), ezt az
' A6ListaFrissites epiti ujra H6 vagy G6 valtozasakor:
'   MINDIG -> az osszes cikkszam (R aktiv, R kifutott, A), fuggetlenul
'             H6-tol ES G6-tol is (nincs szures egyik alapjan sem)
'
' FONTOS: ez a fajl szandekosan CSAK ASCII karaktereket tartalmaz. Az
' ekezetes szovegeket ChrW() allitja elo.
'==========================================================================

'--- Tabla- es oszlopnevek -------------------------------------------------
Public Const LO_ARTICLE  As String = "Article"
Public Const LO_TYPE     As String = "Type"

Public Const C_NUM   As String = "PartNum"
Public Const C_NAME  As String = "PartName"
Public Const C_TYPE  As String = "Type"
Public Const C_SRC   As String = "Forras"
Public Const C_STAT  As String = "Statusz"
Public Const C_PROD  As String = "Berendezes"
Public Const C_NOTE  As String = "Ellenorzes"
Public Const C_LIST  As String = "Lista"

'--- Legordulo ------------------------------------------------------------
Public Const NEV_A6LISTA    As String = "A6Lista"
Public Const BERENDEZES_LBL As String = "7.50"
Private Const LISTA_OSZLOP  As Long = 10          ' Parts!J
Private Const STAT_KIFUTOTT As String = "Kifutott"

'--- Forrasfajl -----------------------------------------------------------
Private Const TERMEKADATOK_MAPPA As String = _
    "P:\_Departments\PW\2. PW_Fertigung\4_LB-XL_7.50_Laser_Blanking_line\07_TEAM\"
Private Const NEV_UTVONAL As String = "TermekadatokUtvonal"   ' rejtett nev
Private Const T_TERMEK As String = "tblTermek"

'==========================================================================
' 1. Termekadatok frissitese (Alt+F8 -> TermekadatokFrissitese)
'--------------------------------------------------------------------------
' A forras a sajat LBL_Termekadatok_sajat.xlsm tblTermek tablaja.
'==========================================================================
Public Sub TermekadatokFrissitese()

    Dim fajl As String
    Dim src As Workbook, sajatNyitas As Boolean
    Dim dNum As Object, sorrend As Collection
    Dim t As Variant, idx As Variant
    Dim duplak As String, ujTipusok As String
    Dim nAktiv As Long, nKifutott As Long, nHianyzoBer As Long
    Dim lepes As String

    On Error GoTo Hiba

    lepes = "fajl kivalasztasa"
    fajl = ForrasFajlValasztas()
    If Len(fajl) = 0 Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    lepes = "Article oszlopok ellenorzese"
    ArticleOszlopokBiztositasa

    '--- Forras megnyitasa (ha mar nyitva van, azt hasznaljuk) ------------
    lepes = "forrasfajl megnyitasa"
    Set src = NyitottMunkafuzet(fajl)
    If src Is Nothing Then
        Set src = Workbooks.Open(Filename:=fajl, ReadOnly:=True, _
                                 UpdateLinks:=0, AddToMru:=False)
        sajatNyitas = True
    End If

    lepes = "tblTermek beolvasasa"
    Set dNum = CreateObject("Scripting.Dictionary")
    dNum.CompareMode = vbTextCompare
    Set sorrend = New Collection
    SajatTablaBeolvasasa src, dNum, sorrend, duplak, nHianyzoBer

    If sajatNyitas Then src.Close SaveChanges:=False
    Set src = Nothing

    If dNum.Count = 0 Then
        Err.Raise vbObjectError + 521, , _
            "A fajlban nem talalhato 'tblTermek' tabla (sem fejlec alapjan 'Termekszam' oszlop), vagy nincs benne ervenyes tetel."
    End If

    '--- Article ujrairasa: A-sorok maradnak, R-sorok cserelodnek ---------
    lepes = "Article ujrairasa"
    ArticleRSorokCsereje dNum, sorrend, nAktiv, nKifutott

    '--- Uj tipuskodok a Type tablaba -------------------------------------
    lepes = "Type tabla bovitese"
    For Each idx In sorrend
        t = dNum(idx)
        If Len(t(2)) > 0 Then
            If TipusBiztositas(CStr(t(2))) Then
                If InStr(1, ujTipusok, "  " & t(2) & vbLf) = 0 Then
                    ujTipusok = ujTipusok & "  " & t(2) & vbLf
                End If
            End If
        End If
    Next idx

    '--- Utvonal megjegyzese, lista ujraepitese ---------------------------
    lepes = "utvonal mentese"
    UtvonalMentes fajl

    lepes = "A6 lista ujraepitese"
    A6ListaFrissites

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    '--- Osszegzes --------------------------------------------------------
    Dim uzenet As String
    uzenet = "Forras: " & Mid$(fajl, InStrRev(fajl, "\") + 1) & vbLf & vbLf & _
             "R-szamok az Article tablaban: " & (nAktiv + nKifutott) & vbLf & _
             "  aktiv: " & nAktiv & vbLf & _
             "  kifutott: " & nKifutott & vbLf
    If nHianyzoBer > 0 Then uzenet = uzenet & vbLf & _
        "Berendezes oszlop nelkuli/ures tetel (7.50-et feltetelezve): " & nHianyzoBer & vbLf
    If Len(duplak) > 0 Then uzenet = uzenet & vbLf & _
        "Ismetlodo termekszam a sajat tablaban (az elso maradt):" & vbLf & duplak
    If Len(ujTipusok) > 0 Then uzenet = uzenet & vbLf & _
        "Uj tipuskod a Type tablaban:" & vbLf & ujTipusok

    MsgBox uzenet, vbInformation, "Termekadatok frissitese"
    Exit Sub

Hiba:
    Dim hSzam As Long, hSzoveg As String
    hSzam = Err.Number: hSzoveg = Err.Description
    On Error Resume Next
    If sajatNyitas Then
        If Not src Is Nothing Then src.Close SaveChanges:=False
    End If
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "A frissites megszakadt [" & lepes & "]:" & vbLf & hSzoveg & _
           " (" & hSzam & ")", vbExclamation, "Termekadatok frissitese"
End Sub

'--------------------------------------------------------------------------
' A sajat termekadat-fajl tblTermek tablajanak beolvasasa.
' dNum(kulcs) = Array(Termekszam-ahogy-all, Termeknev, Tipus, aktivBoolean, Berendezes)
'--------------------------------------------------------------------------
Private Sub SajatTablaBeolvasasa(src As Workbook, dNum As Object, sorrend As Collection, _
                                 ByRef duplak As String, ByRef nHianyzoBer As Long)

    Dim lo As ListObject
    Dim cNum As Long, cNev As Long, cTip As Long, cBer As Long, cStat As Long
    Dim r As Long, k As String, nev As String, tip As String, ber As String, stat As String
    Dim aktiv As Boolean, kif As Boolean

    Set lo = TablaBarmelyLapon(src, T_TERMEK)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then
        FejlecAlapuBeolvasas src, dNum, sorrend, duplak, nHianyzoBer
        Exit Sub
    End If

    cNum = OszlopFejlecAlapjan(lo, Array("TERMEKSZAM", "CIKKSZAM"))
    cNev = OszlopFejlecAlapjan(lo, Array("TERMEKNEV", "MEGNEVEZES"))
    cTip = OszlopFejlecAlapjan(lo, Array("TIPUS"))
    cBer = OszlopFejlecAlapjan(lo, Array("BERENDEZES"))
    cStat = OszlopFejlecAlapjan(lo, Array("STATUSZ"))
    If cNum = 0 Then Exit Sub

    For r = 1 To lo.ListRows.Count
        k = Kulcs(Szoveg(lo.DataBodyRange.Cells(r, cNum).Value))
        If Len(k) = 0 Then GoTo KovSor

        nev = "": If cNev > 0 Then nev = Szoveg(lo.DataBodyRange.Cells(r, cNev).Value)
        tip = "": If cTip > 0 Then tip = Szoveg(lo.DataBodyRange.Cells(r, cTip).Value)
        ber = ""
        If cBer > 0 Then ber = Trim$(Replace(Szoveg(lo.DataBodyRange.Cells(r, cBer).Value), ",", "."))
        If Len(ber) = 0 Then
            ber = BERENDEZES_LBL
            nHianyzoBer = nHianyzoBer + 1
        End If
        stat = "": If cStat > 0 Then stat = Szoveg(lo.DataBodyRange.Cells(r, cStat).Value)

        ' "...eltunt a forrasbol..." -> kimarad
        If InStr(1, NormB(stat), "ELTUNT", vbTextCompare) > 0 Then GoTo KovSor
        kif = (StrComp(Trim$(stat), "Kifutott", vbTextCompare) = 0)
        aktiv = Not kif

        If dNum.Exists(k) Then
            duplak = duplak & "  " & Szoveg(lo.DataBodyRange.Cells(r, cNum).Value) & vbLf
            GoTo KovSor
        End If

        dNum.Add k, Array(Szoveg(lo.DataBodyRange.Cells(r, cNum).Value), nev, tip, aktiv, ber)
        sorrend.Add k
KovSor:
    Next r

End Sub

' Tartalek: ha nincs nev szerinti tblTermek tabla, fejlec alapjan keresunk
' (barmelyik lapon, elso 10 soron belul egy "Termekszam"/"Cikkszam" cella).
Private Sub FejlecAlapuBeolvasas(src As Workbook, dNum As Object, sorrend As Collection, _
                                 ByRef duplak As String, ByRef nHianyzoBer As Long)

    Dim ws As Worksheet
    Dim hdrSor As Long, r As Long, c As Long, utSor As Long, utOszlop As Long
    Dim cNum As Long, cNev As Long, cTip As Long, cBer As Long, cStat As Long
    Dim h As String, k As String, nev As String, tip As String, ber As String, stat As String
    Dim aktiv As Boolean, kif As Boolean

    For Each ws In src.Worksheets
        hdrSor = 0
        For r = 1 To 10
            For c = 1 To 30
                h = NormB(Szoveg(ws.Cells(r, c).Value))
                If h = "TERMEKSZAM" Or h = "CIKKSZAM" Then hdrSor = r: Exit For
            Next c
            If hdrSor > 0 Then Exit For
        Next r
        If hdrSor = 0 Then GoTo KovLap

        utOszlop = ws.Cells(hdrSor, ws.Columns.Count).End(xlToLeft).Column
        cNum = 0: cNev = 0: cTip = 0: cBer = 0: cStat = 0
        For c = 1 To utOszlop
            h = NormB(Szoveg(ws.Cells(hdrSor, c).Value))
            Select Case h
                Case "TERMEKSZAM", "CIKKSZAM": If cNum = 0 Then cNum = c
                Case "TERMEKNEV", "MEGNEVEZES": If cNev = 0 Then cNev = c
                Case "TIPUS": If cTip = 0 Then cTip = c
                Case "BERENDEZES": If cBer = 0 Then cBer = c
                Case "STATUSZ": If cStat = 0 Then cStat = c
            End Select
        Next c
        If cNum = 0 Then GoTo KovLap

        utSor = ws.Cells(ws.Rows.Count, cNum).End(xlUp).Row
        For r = hdrSor + 1 To utSor
            k = Kulcs(Szoveg(ws.Cells(r, cNum).Value))
            If Len(k) = 0 Then GoTo KovSor

            nev = "": If cNev > 0 Then nev = Szoveg(ws.Cells(r, cNev).Value)
            tip = "": If cTip > 0 Then tip = Szoveg(ws.Cells(r, cTip).Value)
            ber = ""
            If cBer > 0 Then ber = Trim$(Replace(Szoveg(ws.Cells(r, cBer).Value), ",", "."))
            If Len(ber) = 0 Then
                ber = BERENDEZES_LBL
                nHianyzoBer = nHianyzoBer + 1
            End If
            stat = "": If cStat > 0 Then stat = Szoveg(ws.Cells(r, cStat).Value)

            If InStr(1, NormB(stat), "ELTUNT", vbTextCompare) > 0 Then GoTo KovSor
            kif = (StrComp(Trim$(stat), "Kifutott", vbTextCompare) = 0)
            aktiv = Not kif

            If dNum.Exists(k) Then
                duplak = duplak & "  " & Szoveg(ws.Cells(r, cNum).Value) & vbLf
                GoTo KovSor
            End If

            dNum.Add k, Array(Szoveg(ws.Cells(r, cNum).Value), nev, tip, aktiv, ber)
            sorrend.Add k
KovSor:
        Next r
        Exit Sub        ' az elso talalt lap utan mar nem keresunk tovabb
KovLap:
    Next ws

End Sub

Private Function TablaBarmelyLapon(wb As Workbook, ByVal nev As String) As ListObject
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In wb.Worksheets
        Set TablaBarmelyLapon = ws.ListObjects(nev)
        If Not TablaBarmelyLapon Is Nothing Then Exit Function
    Next ws
End Function

' Oszlopindex keresese egy ListObject-ben, a fejlec normalizalt szovege alapjan.
Private Function OszlopFejlecAlapjan(lo As ListObject, ByVal jeloltek As Variant) As Long
    Dim j As Long, h As String, p As Variant
    For j = 1 To lo.ListColumns.Count
        h = NormB(lo.ListColumns(j).Name)
        For Each p In jeloltek
            If h = CStr(p) Then OszlopFejlecAlapjan = j: Exit Function
        Next p
    Next j
End Function

' Normalizalas: ekezet nelkul, nagybetu, csak A-Z es 0-9 (a helyi Ekezetlen()-re epul)
Private Function NormB(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    s = UCase$(Ekezetlen(s))
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then out = out & ch
    Next i
    NormB = out
End Function

'--------------------------------------------------------------------------
' Fajlvalaszto: az utoljara hasznalt fajlrol, vagy a 7.50 TEAM mappabol indul
'--------------------------------------------------------------------------
Private Function ForrasFajlValasztas() As String

    Dim kezdo As String

    kezdo = UtvonalOlvasas()
    If Len(kezdo) > 0 Then
        If Not UtvonalLetezik(kezdo, vbNormal) Then kezdo = ""
    End If
    If Len(kezdo) = 0 Then
        If UtvonalLetezik(TERMEKADATOK_MAPPA, vbDirectory) Then kezdo = TERMEKADATOK_MAPPA
    End If

    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "TERMEKADATOK fajl kivalasztasa"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel", "*.xlsx;*.xlsm;*.xls"
        If Len(kezdo) > 0 Then .InitialFileName = kezdo
        If .Show = -1 Then ForrasFajlValasztas = .SelectedItems(1)
    End With

End Function

' Nem csatolt P: meghajtonal a Dir hibat dob, ezert vedett
Private Function UtvonalLetezik(ByVal ut As String, ByVal attr As VbFileAttribute) As Boolean
    On Error Resume Next
    UtvonalLetezik = (Len(Dir$(ut, attr)) > 0)
    On Error GoTo 0
End Function

Private Function NyitottMunkafuzet(ByVal fajl As String) As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fajl, vbTextCompare) = 0 Then
            Set NyitottMunkafuzet = wb
            Exit Function
        End If
    Next wb
End Function

Private Sub UtvonalMentes(ByVal fajl As String)
    On Error Resume Next
    ThisWorkbook.Names(NEV_UTVONAL).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=NEV_UTVONAL, _
        RefersTo:="=""" & Replace(fajl, """", """""") & """", Visible:=False
End Sub

Private Function UtvonalOlvasas() As String
    Dim s As String
    On Error Resume Next
    s = ThisWorkbook.Names(NEV_UTVONAL).RefersTo      ' ="P:\...\x.xlsx"
    On Error GoTo 0
    If Len(s) > 3 Then
        s = Mid$(s, 3, Len(s) - 3)
        UtvonalOlvasas = Replace(s, """""", """")
    End If
End Function

'==========================================================================
' 2. Article tabla kezelese
'==========================================================================

' Hianyzo oszlopok felvetele + Forras kitoltese + Lista keplet.
Public Sub ArticleOszlopokBiztositasa()

    Dim lo As ListObject
    Dim nevek As Variant, i As Long
    Dim vNum As Variant, vSrc As Variant

    Set lo = ThisWorkbook.Worksheets(SH_PARTS).ListObjects(LO_ARTICLE)

    nevek = Array(C_TYPE, C_SRC, C_STAT, C_PROD, C_NOTE, C_LIST)
    For i = LBound(nevek) To UBound(nevek)
        If Not OszlopLetezik(lo, CStr(nevek(i))) Then
            lo.ListColumns.Add.Name = CStr(nevek(i))
        End If
    Next i

    If lo.DataBodyRange Is Nothing Then Exit Sub

    ' Forras: ahol ures, a cikkszambol
    vNum = OszlopTomb(lo, C_NUM)
    vSrc = OszlopTomb(lo, C_SRC)
    For i = 1 To UBound(vNum, 1)
        If Len(Szoveg(vSrc(i, 1))) = 0 Then
            If Kulcs(Szoveg(vNum(i, 1))) Like "R#*" Then vSrc(i, 1) = "R" Else vSrc(i, 1) = "A"
        End If
    Next i
    lo.ListColumns(C_SRC).DataBodyRange.NumberFormat = "@"
    lo.ListColumns(C_SRC).DataBodyRange.Value = vSrc

    ListaKepletBeallitas lo

End Sub

Private Sub ListaKepletBeallitas(lo As ListObject)
    Dim sep As String, q As String, f As String
    q = Chr(34)                                             ' egy literal idezojel
    sep = """ " & ChrW(183) & " """                        ' " . "

    ' A cikkszam resz SZOKOZ NELKUL kerul a Lista-ba (SUBSTITUTE), hogy az
    ' Excel legordulo sajat "gepeles kozbeni" szurese (ami elolrol, karakterrol
    ' karakterre hasonlit) mukodjon akkor is, ha a felhasznalo szokoz nelkul
    ' gepeli be a cikkszamot (pl. "a24463701" -> illeszkedik "A2446370100"-ra).
    ' A nagy-/kisbetu kulonbseg mar eleve nem szamit az Excel legordulo
    ' beepitett szuresenel. Kivalasztas utan a CikkszamValasztas mindig a
    ' valodi (esetleg szokozos) PartNum erteket irja vissza A6-ba, tehat a
    ' vegleges megjelenites nem valtozik.
    f = "=SUBSTITUTE([@" & C_NUM & "]," & q & " " & q & "," & q & q & ")" & _
        "&IF([@" & C_NAME & "]=" & q & q & "," & q & q & "," & sep & "&[@" & C_NAME & "])" & _
        "&IF([@" & C_TYPE & "]=" & q & q & "," & q & q & "," & sep & "&[@" & C_TYPE & "])" & _
        "&IF([@" & C_STAT & "]=" & q & STAT_KIFUTOTT & q & "," & q & " (kifutott)" & q & "," & q & q & ")"

    With lo.ListColumns(C_LIST).DataBodyRange
        .NumberFormat = "General"
        .Formula = f
    End With
End Sub

' Az R-sorokat kicsereli az uj beolvasasra, az A-sorok valtozatlanok.
Private Sub ArticleRSorokCsereje(dNum As Object, sorrend As Collection, _
                                 ByRef nAktiv As Long, ByRef nKifutott As Long)

    Dim lo As ListObject
    Dim oszlopok As Variant, regi() As Variant, uj() As Variant
    Dim nCol As Long, nRegi As Long, nUj As Long, n As Long
    Dim i As Long, j As Long
    Dim vSrc As Variant, t As Variant, idx As Variant

    Set lo = ThisWorkbook.Worksheets(SH_PARTS).ListObjects(LO_ARTICLE)
    oszlopok = Array(C_NUM, C_NAME, C_TYPE, C_SRC, C_STAT, C_PROD, C_NOTE)
    nCol = UBound(oszlopok) + 1

    ' Regi adatok oszloponkent (a sorrend a lapon mindegy)
    nRegi = 0
    If Not lo.DataBodyRange Is Nothing Then nRegi = lo.DataBodyRange.Rows.Count
    ReDim regi(1 To nCol)
    For j = 1 To nCol
        If nRegi > 0 Then regi(j) = OszlopTomb(lo, CStr(oszlopok(j - 1)))
    Next j
    vSrc = regi(4)

    ' Uj tomb: megtartott A-sorok + uj R-sorok
    ReDim uj(1 To nRegi + sorrend.Count, 1 To nCol)
    n = 0
    For i = 1 To nRegi
        If StrComp(Szoveg(vSrc(i, 1)), "R", vbTextCompare) <> 0 Then
            n = n + 1
            For j = 1 To nCol
                uj(n, j) = regi(j)(i, 1)
            Next j
        End If
    Next i
    For Each idx In sorrend
        t = dNum(idx)
        n = n + 1
        uj(n, 1) = t(0)
        uj(n, 2) = t(1)
        uj(n, 3) = t(2)
        uj(n, 4) = "R"
        If t(3) Then
            uj(n, 5) = StatAktiv(): nAktiv = nAktiv + 1
        Else
            uj(n, 5) = STAT_KIFUTOTT: nKifutott = nKifutott + 1
        End If
        uj(n, 6) = t(4)          ' Berendezes: tetelenkent, a sajat tablabol
        uj(n, 7) = ""
    Next idx
    nUj = n

    ArticleIras lo, oszlopok, uj, nUj, nRegi

End Sub

' Tomb visszairasa az Article tablaba, a tabla atmeretezesevel.
Private Sub ArticleIras(lo As ListObject, oszlopok As Variant, adat() As Variant, _
                        ByVal n As Long, ByVal nRegi As Long)

    Dim j As Long, i As Long
    Dim oszlop() As Variant
    Dim maradek As Range

    If n < 1 Then Err.Raise vbObjectError + 530, , "Az Article tabla ures lenne."

    ' Ha kisebb lesz, a felesleges sorokat az atmeretezes utan uritjuk
    If nRegi > n Then
        Set maradek = lo.Range.Rows(n + 2).Resize(nRegi - n)
    End If
    lo.Resize lo.Range.Cells(1, 1).Resize(n + 1, lo.Range.Columns.Count)
    If Not maradek Is Nothing Then maradek.ClearContents

    For j = LBound(oszlopok) To UBound(oszlopok)
        ReDim oszlop(1 To n, 1 To 1)
        For i = 1 To n
            oszlop(i, 1) = adat(i, j - LBound(oszlopok) + 1)
        Next i
        With lo.ListColumns(CStr(oszlopok(j))).DataBodyRange
            .NumberFormat = "@"          ' a "7.50" ne legyen 7,5
            .Value = oszlop
        End With
    Next j

    ListaKepletBeallitas lo

End Sub

'==========================================================================
' 3. Article tisztitasa (egyszeri, Alt+F8 -> ArticleTisztitas)
'--------------------------------------------------------------------------
' - A/N cikkszam szokoz nelkuli alakra ("A 177 611 13 00" -> A1776111300)
' - duplikatum: az elso marad; eltero megnevezesnel a tobbi az Ellenorzes
'   oszlopba kerul, hogy kezzel el lehessen donteni
' - hianyzo megnevezes jelolese az Ellenorzes oszlopban
' Az R-sorokhoz nem nyul (azokat a frissites kezeli).
'==========================================================================
Public Sub ArticleTisztitas()

    Dim lo As ListObject
    Dim oszlopok As Variant, regi() As Variant, uj() As Variant
    Dim nCol As Long, nRegi As Long, n As Long, i As Long, j As Long
    Dim d As Object
    Dim k As String, nev As String, hely As Long
    Dim nForma As Long, nAzonos As Long, nUtkozo As Long, nUres As Long
    Dim uzenet As String

    On Error GoTo Hiba

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ArticleOszlopokBiztositasa

    Set lo = ThisWorkbook.Worksheets(SH_PARTS).ListObjects(LO_ARTICLE)
    If lo.DataBodyRange Is Nothing Then GoTo Vege
    nRegi = lo.DataBodyRange.Rows.Count

    oszlopok = Array(C_NUM, C_NAME, C_TYPE, C_SRC, C_STAT, C_PROD, C_NOTE)
    nCol = UBound(oszlopok) + 1
    ReDim regi(1 To nCol)
    For j = 1 To nCol
        regi(j) = OszlopTomb(lo, CStr(oszlopok(j - 1)))
    Next j

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    ReDim uj(1 To nRegi, 1 To nCol)
    n = 0

    For i = 1 To nRegi
        If StrComp(Szoveg(regi(4)(i, 1)), "R", vbTextCompare) = 0 Then
            k = "R|" & Kulcs(Szoveg(regi(1)(i, 1)))       ' R-sor: nem bantjuk
        Else
            k = Kulcs(Szoveg(regi(1)(i, 1)))
            If k <> Szoveg(regi(1)(i, 1)) Then nForma = nForma + 1
        End If
        nev = Szoveg(regi(2)(i, 1))

        If d.Exists(k) Then
            hely = d(k)
            If Len(Szoveg(uj(hely, 2))) = 0 And Len(nev) > 0 Then
                uj(hely, 2) = nev                   ' az elsonek nem volt neve
                nAzonos = nAzonos + 1
            ElseIf StrComp(nev, Szoveg(uj(hely, 2)), vbTextCompare) = 0 Or Len(nev) = 0 Then
                nAzonos = nAzonos + 1
            Else
                nUtkozo = nUtkozo + 1
                If InStr(1, Szoveg(uj(hely, 7)), nev, vbTextCompare) = 0 Then
                    uj(hely, 7) = Szoveg(uj(hely, 7)) & _
                        IIf(Len(Szoveg(uj(hely, 7))) > 0, "; ", "Mas megnevezes is: ") & nev
                End If
            End If
        Else
            n = n + 1
            d.Add k, n
            For j = 1 To nCol
                uj(n, j) = regi(j)(i, 1)
            Next j
            If Left$(k, 2) <> "R|" Then uj(n, 1) = k
        End If
    Next i

    For i = 1 To n
        If Len(Szoveg(uj(i, 2))) = 0 And Len(Szoveg(uj(i, 7))) = 0 Then
            uj(i, 7) = "Nincs megnevezes": nUres = nUres + 1
        End If
    Next i

    Application.ScreenUpdating = True
    uzenet = "Article tisztitas - elonezet" & vbLf & vbLf & _
             "Sorok most: " & nRegi & vbLf & _
             "Sorok utana: " & n & vbLf & vbLf & _
             "Szokoz nelkuli alakra hozva: " & nForma & vbLf & _
             "Torolt duplikatum (azonos megnevezes): " & nAzonos & vbLf & _
             "Utkozo megnevezes (Ellenorzes oszlopba irva): " & nUtkozo & vbLf & _
             "Hianyzo megnevezes (jelolve): " & nUres & vbLf & vbLf & _
             "Vegrehajtsam? (Elotte erdemes elmenteni egy masolatot.)"
    If MsgBox(uzenet, vbQuestion + vbYesNo + vbDefaultButton2, "Article tisztitas") <> vbYes Then GoTo Vege
    Application.ScreenUpdating = False

    ArticleIras lo, oszlopok, uj, n, nRegi
    A6ListaFrissites

Vege:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

Hiba:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "A tisztitas megszakadt: " & Err.Description, vbExclamation, "Article tisztitas"
End Sub

'==========================================================================
' 4. A6 legordulo lista (Parts!J -> A6Lista nev)
'==========================================================================
' Parameter nelkuli "csomagolo", hogy Alt+F8-bol is lathato/futtathato legyen
' (a parameteres Sub-ok nem jelennek meg a Makro listaban).
Public Sub A6ListaUjraepites()
    A6ListaFrissites
    MsgBox "Az A6 lista ujraepult.", vbInformation, "A6 lista"
End Sub

Public Sub A6ListaFrissites(Optional ws As Worksheet = Nothing)

    Dim wsParts As Worksheet, lo As ListObject
    Dim vList As Variant, vSrc As Variant, vStat As Variant
    Dim ki() As Variant, n As Long, i As Long, menet As Long, nSor As Long
    Dim rSor As Boolean, kifutott As Boolean, bele As Boolean
    Dim regiUtolso As Long

    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets(SH_FOTO)
    Set wsParts = ThisWorkbook.Worksheets(SH_PARTS)
    Set lo = wsParts.ListObjects(LO_ARTICLE)
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Not OszlopLetezik(lo, C_LIST) Then Exit Sub     ' meg nincs beallitva

    ' v9.5: minden cikkszam (A- es R-szam, aktiv es kifutott is) MINDIG
    ' valaszthato, fuggetlenul H6-tol ES G6-tol (nincs se Berendezes-,
    ' se Tipus-alapu szures). H6 csak kimenet: R-szam kivalasztasakor a
    ' CikkszamValasztas allitja be a tetel sajat Berendezes ertekere.
    '
    ' v9.11: EGYETLEN, tiszta lista (nincs kereso-elotag, nincs kulon ablak).
    ' A "barhol keres" (pl. "637" vagy "sitenwand" begepelese) az A6 melletti
    ' legordulo NYIL-ra kattintva megjeleno, Excel sajat beepitett kereso-
    ' mezojevel mukodik - ugyanaz a legordulo, csak a nyillal nyitva meg
    ' egy keresosav is van benne. Kozvetlen begepeleskor (nyil nelkul) az
    ' egyezes a cikkszam ELEJETOL szur (szokoz/nagybetu nem szamit, l. 9.6).

    vList = OszlopTomb(lo, C_LIST)
    vSrc = OszlopTomb(lo, C_SRC)
    vStat = OszlopTomb(lo, C_STAT)
    nSor = UBound(vList, 1)

    ReDim ki(1 To nSor, 1 To 1)
    ' 1. menet: aktiv R, 2. menet: kifutott R, 3. menet: A-sorok
    For menet = 1 To 3
        For i = 1 To nSor
            rSor = (StrComp(Szoveg(vSrc(i, 1)), "R", vbTextCompare) = 0)
            kifutott = (StrComp(Szoveg(vStat(i, 1)), STAT_KIFUTOTT, vbTextCompare) = 0)
            Select Case menet
                Case 1: bele = rSor And Not kifutott
                Case 2: bele = rSor And kifutott
                Case 3: bele = Not rSor
            End Select
            If bele Then
                If Not IsError(vList(i, 1)) Then
                    n = n + 1
                    ki(n, 1) = vList(i, 1)
                End If
            End If
        Next i
    Next menet

    ' Kiiras a J oszlopba (a regi lista helyere)
    regiUtolso = wsParts.Cells(wsParts.Rows.Count, LISTA_OSZLOP).End(xlUp).Row
    If regiUtolso >= 2 Then wsParts.Range(wsParts.Cells(2, LISTA_OSZLOP), _
                                          wsParts.Cells(regiUtolso, LISTA_OSZLOP)).ClearContents
    wsParts.Cells(1, LISTA_OSZLOP).Value = "A6 lista (automatikus, ne szerkeszd)"
    If n > 0 Then
        With wsParts.Cells(2, LISTA_OSZLOP).Resize(n, 1)
            .NumberFormat = "@"
            .Value = ki          ' a tomb nagyobb, csak az elso n sor kerul ki
        End With
    Else
        n = 1                    ' ures lista: egy ures cellara mutatunk
    End If

    On Error Resume Next
    ThisWorkbook.Names(NEV_A6LISTA).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=NEV_A6LISTA, _
        RefersTo:="='" & SH_PARTS & "'!$J$2:$J$" & (n + 1)

End Sub

'==========================================================================
' 5. A6 valasztas -> A6 = cikkszam, C6 = megnevezes, G6 = tipus
'==========================================================================
Public Sub CikkszamValasztas(ws As Worksheet)

    Dim s As String, szam As String, sep As String
    Dim p As Long

    s = Szoveg(ws.Range("A6").Value)
    If Len(s) = 0 Then
        ws.Range("C6").Value = ""
        Exit Sub
    End If

    ' "R 244 637 0100 . BPL SEITENWAND LI . X244/248 (kifutott)" -> szam
    sep = " " & ChrW(183) & " "
    p = InStr(1, s, sep)
    If p > 0 Then szam = Left$(s, p - 1) Else szam = s
    If Right$(szam, 11) = " (kifutott)" Then szam = Left$(szam, Len(szam) - 11)
    szam = Trim$(szam)

    If Not CikkszamBeallitasa(ws, Kulcs(szam)) Then
        ' Nincs a torzsben: a szam marad, megnevezes ures
        ws.Range("A6").Value = szam
        ws.Range("C6").Value = ""
    End If

End Sub

' A tenyleges "A6/C6/G6/H6 kitoltese egy tudott Kulcs() alapjan" logika
' kulon Public Function-ba kiemelve. True, ha talalt es kitoltott egy tetelt.
Public Function CikkszamBeallitasa(ws As Worksheet, ByVal k As String) As Boolean

    Dim lo As ListObject
    Dim i As Long
    Dim vNum As Variant, vName As Variant, vType As Variant, vBer As Variant
    Dim vanTipus As Boolean, vanBer As Boolean
    Dim ujBer As String, regiBer As String

    Set lo = ThisWorkbook.Worksheets(SH_PARTS).ListObjects(LO_ARTICLE)
    If lo.DataBodyRange Is Nothing Then Exit Function
    vNum = OszlopTomb(lo, C_NUM)
    vName = OszlopTomb(lo, C_NAME)
    vanTipus = OszlopLetezik(lo, C_TYPE)
    If vanTipus Then vType = OszlopTomb(lo, C_TYPE)
    vanBer = OszlopLetezik(lo, C_PROD)
    If vanBer Then vBer = OszlopTomb(lo, C_PROD)

    For i = 1 To UBound(vNum, 1)
        If Kulcs(Szoveg(vNum(i, 1))) = k Then
            ws.Range("A6").Value = Szoveg(vNum(i, 1))
            ws.Range("C6").Value = Szoveg(vName(i, 1))
            If vanTipus Then
                If Len(Szoveg(vType(i, 1))) > 0 Then ws.Range("G6").Value = Szoveg(vType(i, 1))
            End If

            ' v9.3: a gyarto berendezes (H6) automatikus at-/kitoltese a tetel
            ' sajat Berendezes ertekebol, ha az elter a jelenlegi H6-tol
            ' v9.12: ha a cikkszam A-val kezdodik, a berendezes mindig 7.41
            ' (fuggetlenul attol, hogy az Article tablaban mi szerepel)
            If Left$(Trim$(Szoveg(vNum(i, 1))), 1) = "A" Then
                ujBer = "7.41"
            ElseIf vanBer Then
                ujBer = Trim$(Replace(Szoveg(vBer(i, 1)), ",", "."))
            Else
                ujBer = ""
            End If

            If Len(ujBer) > 0 Then
                regiBer = Trim$(Replace(Szoveg(ws.Range("H6").Value), ",", "."))
                If StrComp(regiBer, ujBer, vbTextCompare) <> 0 Then
                    ws.Range("H6").Value = ujBer
                    HibafeltarasKitoltes ws
                End If
            End If

            ' v9.13: R-cikkszamnal E4/E6 = Alapanyag/Material - Platine,
            ' minden mas esetben az E4 eredeti felirata all vissza
            If Left$(Trim$(Szoveg(vNum(i, 1))), 1) = "R" Then
                ws.Range("E4").Value = "Alapanyag/Material"
                ws.Range("E6").Value = "Platine"
            Else
                ws.Range("E4").Value = PreselesFeliratEredeti()
                If StrComp(Szoveg(ws.Range("E6").Value), "Platine", vbTextCompare) = 0 Then
                    ws.Range("E6").Value = ""
                End If
            End If

            A6ListaFrissites ws
            CikkszamBeallitasa = True
            Exit Function
        End If
    Next i

End Function

'==========================================================================
' 6. Tipuskodok (Data!Type)
'==========================================================================

' Felveszi a kodot, ha meg nincs. True = most lett felveve.
Public Function TipusBiztositas(ByVal kod As String) As Boolean
    Dim lo As ListObject, c As Range
    Set lo = ThisWorkbook.Worksheets(SH_DATA).ListObjects(LO_TYPE)
    If Not lo.DataBodyRange Is Nothing Then
        For Each c In lo.ListColumns(1).DataBodyRange.Cells
            If StrComp(Szoveg(c.Value), kod, vbTextCompare) = 0 Then Exit Function
        Next c
    End If
    lo.ListRows.Add.Range.Cells(1, 1).Value = kod
    TipusBiztositas = True
End Function

' v9.2 tipuskod-dontesek: H248 uj kod, az E-osztaly kodja V530 (X530 helyett).
Public Function TipusKodok92() As String
    Dim lo As ListObject, c As Range, vanV530 As Boolean
    Dim ki As String
    Set lo = ThisWorkbook.Worksheets(SH_DATA).ListObjects(LO_TYPE)

    For Each c In lo.ListColumns(1).DataBodyRange.Cells
        If StrComp(Szoveg(c.Value), "V530", vbTextCompare) = 0 Then vanV530 = True
    Next c
    For Each c In lo.ListColumns(1).DataBodyRange.Cells
        If StrComp(Szoveg(c.Value), "X530", vbTextCompare) = 0 Then
            If vanV530 Then
                lo.ListRows(c.Row - lo.DataBodyRange.Row + 1).Delete
            Else
                c.Value = "V530"
            End If
            ki = ki & "- X530 -> V530" & vbLf
            Exit For
        End If
    Next c
    With ThisWorkbook.Worksheets(SH_FOTO).Range("G6")
        If StrComp(Szoveg(.Value), "X530", vbTextCompare) = 0 Then .Value = "V530"
    End With

    If TipusBiztositas("H248") Then ki = ki & "- H248 felveve" & vbLf
    TipusKodok92 = ki
End Function

'==========================================================================
' 7. Segedek
'==========================================================================
Public Function Kulcs(ByVal s As String) As String
    s = Replace(s, " ", "")
    s = Replace(s, ChrW(160), "")
    s = Replace(s, vbTab, "")
    Kulcs = UCase$(s)
End Function

Private Function RSzamFormatum(ByVal s As String) As String
    Dim k As String
    k = Kulcs(s)
    If k Like "R##########" Then
        RSzamFormatum = "R " & Mid$(k, 2, 3) & " " & Mid$(k, 5, 3) & " " & Mid$(k, 8, 4)
    Else
        RSzamFormatum = Trim$(s)
    End If
End Function

Private Function StatAktiv() As String
    StatAktiv = "Akt" & ChrW(237) & "v"
End Function

' Az E4 eredeti felirata ("Preseles/Abpressung"), ekezetekkel
Private Function PreselesFeliratEredeti() As String
    PreselesFeliratEredeti = "Pr" & ChrW(233) & "sel" & ChrW(233) & "s/Abpressung"
End Function

' Hibaertek- es Empty-biztos szovegge alakitas, levagott szokozokkel
Public Function Szoveg(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    Szoveg = Trim$(Replace(CStr(v), ChrW(160), " "))
End Function

Private Function OszlopLetezik(lo As ListObject, ByVal nev As String) As Boolean
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(lc.Name, nev, vbTextCompare) = 0 Then
            OszlopLetezik = True
            Exit Function
        End If
    Next lc
End Function

' Egy oszlop adatai mindig 2D tombkent (1 soros tablanal is)
Private Function OszlopTomb(lo As ListObject, ByVal nev As String) As Variant
    Dim rng As Range, t() As Variant
    Set rng = lo.ListColumns(nev).DataBodyRange
    If rng.Rows.Count = 1 Then
        ReDim t(1 To 1, 1 To 1)
        t(1, 1) = rng.Value
        OszlopTomb = t
    Else
        OszlopTomb = rng.Value
    End If
End Function
