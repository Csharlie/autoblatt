Attribute VB_Name = "modTermekadatok"
' Keszitette es karbantartja : Sardy Peter (PESARDY)
' Mercedes-Benz Manufacturing Hungary, Kecskemet - PW
' Verzio 9.18 - 2026.10.14.
'   9.18: verzio-igazitas a modFotolap 9.18-hoz, ebben a modulban nincs
'        kodvaltozas.
'   9.17: rendrakas, uj funkcio nelkul. Ismeretlen vagy torolt cikkszamnal
'        a C6, G6 es H6 is kiurul (TetelAdatokUritese; a H6 kezelese egy
'        helyen: BerendezesBeallitas). Az E4/E6 logika egy helyen:
'        AnyagFeliratBeallitas (a lapkod is ezt hivja). A cikkszam-valasztas
'        nem epiti ujra az A6 listat; lapra lepeskor csak akkor epul ujra, ha
'        az Article valtozott (A6ListaAlairas / A6ListaElavult).
'        Uj: hibakep-legordulo a Hibafeltaras cimsorban (a "Hibakep /
'        Fehlerbild" felirat melletti cella), forrasa a Data lap tblHibakep
'        tablaja (HibakepLista nev). A HibakepekFrissitese a Hibamegnevezesek
'        fajl magyar neveit (Bezeichnung HU) veszi at, csak hozzaad, nem torol.
'        Forras megnyitasa egy helyen (ForrasMegnyitasa): ha egy azonos nevu
'        masik fajl mar nyitva van, erthetoen megall (korabban felrevezeto
'        "nem talalhato tabla" hibat adott).
'   9.16: uj forras az A- es N-szamokhoz: Autoblatt_Cikkszamtorzs.xlsx /
'        tblCikkszam, uj makro: CikkszamtorzsFrissitese (elonezettel; az
'        Article nem R sorait lecsereli, az R-sorok valtozatlanok). A
'        cikkszam kivalasztasakor a G6 es a H6 mindig a tetel szerint all
'        be: ures tipusnal a G6 kiurul, ures berendezesnel a H6 a
'        H6_URES_HA_NINCS_BERENDEZES szerint kiurul vagy marad. A 9.12-es
'        "A-szam = mindig 7.41" szabaly KIVEZETVE: a berendezes minden
'        cikkszamnal a tetel sajat Berendezes erteke.
'   9.15: egyseges cikkszam-formatum: betu + 3-3-4, pl. "A177 637 0100" /
'        "R177 637 0100" (az N-szam valtozatlan). Uj kozos fuggvenyek:
'        CikkszamFormatum, CikkszamErvenyes (a nem hasznalt RSzamFormatum
'        torolve). A Lista oszlop a tarolt format mutatja (a 9.6-os
'        SUBSTITUTE kivezetve), az A6-ba mindig az egyseges forma kerul,
'        talalat nelkul is. Import: az R-szam is egyseges formara kerul, a
'        nem szabvanyos szam jelolve. ArticleTisztitas: minden sor (R is)
'        egyseges formara, a nem szabvanyos cikkszam jelolve, az elonezet
'        listazza az utkozo megnevezesu cikkszamokat.
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
' Termekadatok (LBL 7.50) es cikkszamtorzs -> Parts!Article, es az A6
' cikkszam-legordulo
'--------------------------------------------------------------------------
' Parts!Article oszlopai (a nev szamit, a sorrend nem):
'   PartNum    - cikkszam, egyseges forma: A177 637 0100 / R177 637 0100 (N-szam valtozatlan)
'   PartName   - megnevezes
'   Type       - tipus (R: tblTermek 'Tipus', A/N: tblCikkszam 'Tipus'), lehet ures
'   Forras     - R = LBL sajat termekadatok, A / N = cikkszamtorzs (ld. lent)
'   Statusz    - Aktiv / Kifutott (ures = aktiv)
'   Berendezes - a tetel sajat berendezese (7.50/7.41/7.22), lehet ures;
'                R-szamnal oszlop hianyaban 7.50
'   Ellenorzes - megjegyzes (tisztitas, nyitott tetel a cikkszamtorzsben)
'   Lista      - keplet: "Cikkszam . Megnevezes . Tipus (kifutott)"
'
' Ket forras (v9.16), mindketto csak a sajat sorait csereli:
'
' A) A- es N-szamok: Autoblatt_Cikkszamtorzs.xlsx / tblCikkszam, a
'    CikkszamtorzsFrissitese olvassa (csak a Csoport <> "R" sorokat):
'   - Cikkszam -> PartNum, Megnevezes -> PartName, Tipus -> Type,
'     Csoport -> Forras, Berendezes -> Berendezes (szovegkent, "7.41")
'   - Statusz  -> "Kifutott" = jelolve; "...eltunt a forrasbol..." = kimarad;
'                 minden mas = aktiv
'   - Ellenorzes (db) > 0 -> Ellenorzes: "Ellenorizendo (cikkszamtorzs)"
'   A fejlecet PONTOS egyezessel keressuk (NormB), igy a "Tipus (forras)",
'   "Tipus (kezi)" stb. oszlop nem talal, csak a szamolt "Tipus".
'
' B) R-szamok: LBL_Termekadatok_sajat.xlsm / tblTermek, a
'    TermekadatokFrissitese olvassa:
'   - Termekszam -> PartNum (egyseges formara hozva, l. CikkszamFormatum)
'   - Termeknev  -> PartName
'   - Tipus      -> Type (uj kodot automatikusan felvesz a Type tablaba)
'   - Berendezes -> Berendezes (oszlop hianyaban 7.50, figyelmeztetessel)
'   - Statusz    -> "Kifutott" = jelolve; "...eltunt a forrasbol..." = kimarad;
'                   minden mas (pl. "Aktiv", "Aktiv es kifutott - ellenorizendo") = aktiv
' Ha nincs 'tblTermek' nevu tabla a fajlban, fejlec alapjan keresunk
' (Termekszam/Cikkszam, Termeknev/Megnevezes, Tipus, Berendezes, Statusz).
'
' Az A6 legordulo forrasa az A6Lista nev (Parts!J oszlop), ezt az
' A6ListaFrissites epiti ujra: importkor, tisztitaskor, beallitaskor, kezzel
' (A6ListaUjraepites), es lapra lepeskor, ha az Article valtozott (v9.17).
' A tartalma nem fugg H6-tol es G6-tol (v9.5):
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

'--- Viselkedes (v9.16) ----------------------------------------------------
' True  = ha a kivalasztott tetelnek nincs Berendezes erteke, a H6 kiurul
'         (a kerdesblokk elrejtodik, a berendezest kezzel kell valasztani)
' False = ilyenkor a H6 marad, ahogy volt
Public Const H6_URES_HA_NINCS_BERENDEZES As Boolean = True

'--- Forrasfajl -----------------------------------------------------------
Private Const TERMEKADATOK_MAPPA As String = _
    "P:\_Departments\PW\2. PW_Fertigung\4_LB-XL_7.50_Laser_Blanking_line\07_TEAM\"
Private Const NEV_UTVONAL As String = "TermekadatokUtvonal"   ' rejtett nev
Private Const T_TERMEK As String = "tblTermek"
Private Const NEV_TORZS_UTVONAL As String = "CikkszamtorzsUtvonal"   ' rejtett nev
Private Const T_CIKKSZAM As String = "tblCikkszam"

'--- Hibakepek (v9.17) ----------------------------------------------------
Public Const NEV_HIBAKEPLISTA As String = "HibakepLista"          ' legordulo forrasa
Private Const T_HIBAKEP As String = "tblHibakep"                  ' Data lap
Private Const C_HIBAKEP As String = "Hibakep"
Private Const NEV_HIBAKEP_UTVONAL As String = "HibakepUtvonal"    ' rejtett nev

'--- Allapot (v9.17) ------------------------------------------------------
' Az Article alairasa az utolso A6-lista epitesekor (l. A6ListaAlairas).
' VBA-reset utan ures, ilyenkor a kovetkezo lapra lepes egyszer ujraepit.
Private mA6ListaAlairas As String

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
    Set src = ForrasMegnyitasa(fajl, sajatNyitas)       ' v9.17

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

' v9.16: a rejtett nev parameter (alapertelmezes: TermekadatokUtvonal)
'--------------------------------------------------------------------------
' v9.17: forrasfajl megnyitasa CSAK OLVASASRA, mindharom import szamara.
'   - ha pontosan ez a fajl mar nyitva van, azt hasznaljuk (sajatNyitas = False,
'     a vegen nem zarjuk be)
'   - ha egy AZONOS NEVU, de mas helyen levo fajl van nyitva (pl. egy masolat a
'     Letoltesek mappaban, vagy a OneDrive-rol URL-kent megnyitott peldany),
'     megallunk: az Excel ket azonos nevu munkafuzetet nem kezel, es az Open
'     ilyenkor hiba nelkul Nothing-ot adhat (ebbol lett korabban a felrevezeto
'     "nem talalhato tabla" uzenet)
'   - ha az Open megsem ad munkafuzetet, szinten erthetoen megallunk
'--------------------------------------------------------------------------
Private Function ForrasMegnyitasa(ByVal fajl As String, ByRef sajatNyitas As Boolean) As Workbook

    Dim wb As Workbook, nev As String

    sajatNyitas = False
    Set wb = NyitottMunkafuzet(fajl)
    If wb Is Nothing Then
        nev = Mid$(fajl, InStrRev(fajl, "\") + 1)
        For Each wb In Application.Workbooks
            If StrComp(wb.Name, nev, vbTextCompare) = 0 Then
                Err.Raise vbObjectError + 560, , _
                    "Mar nyitva van egy azonos nevu fajl:" & vbLf & wb.FullName & vbLf & _
                    "Zard be, es inditsd ujra a frissitest."
            End If
        Next wb

        Set wb = Workbooks.Open(Filename:=fajl, ReadOnly:=True, _
                                UpdateLinks:=0, AddToMru:=False)
        If wb Is Nothing Then
            Err.Raise vbObjectError + 561, , "A fajl nem nyithato meg: " & fajl
        End If
        sajatNyitas = True
    End If
    Set ForrasMegnyitasa = wb

End Function

Private Sub UtvonalMentes(ByVal fajl As String, Optional ByVal nev As String = NEV_UTVONAL)
    On Error Resume Next
    ThisWorkbook.Names(nev).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nev, _
        RefersTo:="=""" & Replace(fajl, """", """""") & """", Visible:=False
End Sub

Private Function UtvonalOlvasas(Optional ByVal nev As String = NEV_UTVONAL) As String
    Dim s As String
    On Error Resume Next
    s = ThisWorkbook.Names(nev).RefersTo              ' ="P:\...\x.xlsx"
    On Error GoTo 0
    If Len(s) > 3 Then
        s = Mid$(s, 3, Len(s) - 3)
        UtvonalOlvasas = Replace(s, """""", """")
    End If
End Function

'==========================================================================
' 1b. Cikkszamtorzs frissitese (Alt+F8 -> CikkszamtorzsFrissitese), v9.16
'--------------------------------------------------------------------------
' A forras az Autoblatt_Cikkszamtorzs.xlsx tblCikkszam tablaja. Csak az A-
' es N-szamokat (Csoport <> "R") vesszuk at: az Article nem R sorait teljesen
' lecsereljuk, az R-sorokhoz (LBL, TermekadatokFrissitese) nem nyulunk.
' A fajlt csak olvasasra nyitjuk meg, es mentes nelkul zarjuk be.
'==========================================================================
Public Sub CikkszamtorzsFrissitese()

    Const MAX_UJ_PELDA As Long = 10
    Const MAX_ELTUNO_PELDA As Long = 15

    Dim fajl As String
    Dim src As Workbook, sajatNyitas As Boolean
    Dim lo As ListObject
    Dim torzs() As Variant, nTorzs As Long
    Dim dTorzs As Object, dRegi As Object, dR As Object, dTipus As Object
    Dim vNum As Variant, vSrc As Variant, idx As Variant
    Dim i As Long, k As String, tip As String
    Dim nR As Long, nRegiNemR As Long, nUjCikk As Long, nEltunik As Long
    Dim nTipus As Long, nBer As Long, nKihagyott As Long, nRUtkozes As Long, nDupla As Long
    Dim ujPeldak As String, eltunoPeldak As String, ujTipusok As String
    Dim uzenet As String, lepes As String

    On Error GoTo Hiba

    lepes = "fajl kivalasztasa"
    fajl = TorzsFajlValasztas()
    If Len(fajl) = 0 Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    lepes = "Article oszlopok ellenorzese"
    ArticleOszlopokBiztositasa

    '--- Jelenlegi Article: R-kulcsok es a cserelendo A/N-kulcsok ---------
    lepes = "Article beolvasasa"
    Set lo = ThisWorkbook.Worksheets(SH_PARTS).ListObjects(LO_ARTICLE)
    Set dRegi = CreateObject("Scripting.Dictionary")
    dRegi.CompareMode = vbTextCompare
    Set dR = CreateObject("Scripting.Dictionary")
    dR.CompareMode = vbTextCompare
    If Not lo.DataBodyRange Is Nothing Then
        vNum = OszlopTomb(lo, C_NUM)
        vSrc = OszlopTomb(lo, C_SRC)
        For i = 1 To UBound(vNum, 1)
            k = Kulcs(Szoveg(vNum(i, 1)))
            If StrComp(Szoveg(vSrc(i, 1)), "R", vbTextCompare) = 0 Then
                nR = nR + 1
                If Len(k) > 0 Then dR(k) = True
            Else
                nRegiNemR = nRegiNemR + 1
                If Len(k) > 0 Then
                    If Not dRegi.Exists(k) Then dRegi.Add k, CikkszamFormatum(Szoveg(vNum(i, 1)))
                End If
            End If
        Next i
    End If

    '--- Forras megnyitasa (ha mar nyitva van, azt hasznaljuk) ------------
    lepes = "forrasfajl megnyitasa"
    Set src = ForrasMegnyitasa(fajl, sajatNyitas)       ' v9.17

    lepes = "tblCikkszam beolvasasa"
    Set dTorzs = CreateObject("Scripting.Dictionary")
    dTorzs.CompareMode = vbTextCompare
    TorzsTablaBeolvasasa src, dR, dTorzs, torzs, nTorzs, nKihagyott, nRUtkozes, nDupla

    If sajatNyitas Then src.Close SaveChanges:=False
    Set src = Nothing

    If nTorzs = 0 Then
        Err.Raise vbObjectError + 542, , "A tblCikkszam tablaban nincs atveheto A/N-cikkszam."
    End If

    '--- Elonezet ---------------------------------------------------------
    lepes = "elonezet"
    For i = 1 To nTorzs
        If Len(Szoveg(torzs(i, 3))) > 0 Then nTipus = nTipus + 1
        If Len(Szoveg(torzs(i, 6))) > 0 Then nBer = nBer + 1
        If Not dRegi.Exists(Kulcs(Szoveg(torzs(i, 1)))) Then
            nUjCikk = nUjCikk + 1
            If nUjCikk <= MAX_UJ_PELDA Then ujPeldak = ujPeldak & "  " & torzs(i, 1) & vbLf
        End If
    Next i
    If nUjCikk > MAX_UJ_PELDA Then ujPeldak = ujPeldak & TovabbiSzoveg(nUjCikk - MAX_UJ_PELDA)

    For Each idx In dRegi.Keys
        If Not dTorzs.Exists(idx) Then
            nEltunik = nEltunik + 1
            If nEltunik <= MAX_ELTUNO_PELDA Then eltunoPeldak = eltunoPeldak & "  " & dRegi(idx) & vbLf
        End If
    Next idx
    If nEltunik > MAX_ELTUNO_PELDA Then eltunoPeldak = eltunoPeldak & TovabbiSzoveg(nEltunik - MAX_ELTUNO_PELDA)

    uzenet = "Cikkszamtorzs - elonezet" & vbLf & _
             "Forras: " & Mid$(fajl, InStrRev(fajl, "\") + 1) & vbLf & vbLf & _
             "A/N-sorok most: " & nRegiNemR & vbLf & _
             "A/N-sorok utana: " & nTorzs & vbLf & _
             "R-sorok (valtozatlanok): " & nR & vbLf & vbLf & _
             "Tipust kap: " & nTipus & vbLf & _
             "Berendezest kap: " & nBer & vbLf
    If nKihagyott > 0 Then uzenet = uzenet & _
        "Kihagyva (eltunt a forrasbol): " & nKihagyott & vbLf
    If nRUtkozes > 0 Then uzenet = uzenet & _
        "Kihagyva (R-sorkent mar megvan): " & nRUtkozes & vbLf
    If nDupla > 0 Then uzenet = uzenet & _
        "Ismetlodo cikkszam a torzsben (az elso maradt): " & nDupla & vbLf
    uzenet = uzenet & vbLf & "Uj cikkszam: " & nUjCikk & vbLf & ujPeldak & _
             vbLf & "Eltuno A/N-cikkszam (nincs a torzsben): " & nEltunik & vbLf & eltunoPeldak & _
             vbLf & "Vegrehajtsam? (Elotte erdemes elmenteni egy masolatot.)"

    Application.ScreenUpdating = True
    If MsgBox(uzenet, vbQuestion + vbYesNo + vbDefaultButton2, "Cikkszamtorzs frissitese") <> vbYes Then GoTo Vege
    Application.ScreenUpdating = False

    '--- Article ujrairasa: R-sorok maradnak, A/N-sorok cserelodnek -------
    lepes = "Article ujrairasa"
    ArticleNemRSorokCsereje torzs, nTorzs

    '--- Uj tipuskodok a Type tablaba -------------------------------------
    lepes = "Type tabla bovitese"
    Set dTipus = CreateObject("Scripting.Dictionary")
    dTipus.CompareMode = vbTextCompare
    For i = 1 To nTorzs
        tip = Szoveg(torzs(i, 3))
        If Len(tip) > 0 Then
            If Not dTipus.Exists(tip) Then
                dTipus.Add tip, True
                If TipusBiztositas(tip) Then ujTipusok = ujTipusok & "  " & tip & vbLf
            End If
        End If
    Next i

    '--- Utvonal megjegyzese, lista ujraepitese ---------------------------
    lepes = "utvonal mentese"
    UtvonalMentes fajl, NEV_TORZS_UTVONAL

    lepes = "A6 lista ujraepitese"
    A6ListaFrissites

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    '--- Osszegzes --------------------------------------------------------
    uzenet = "Forras: " & Mid$(fajl, InStrRev(fajl, "\") + 1) & vbLf & vbLf & _
             "A/N-cikkszamok az Article tablaban: " & nTorzs & vbLf & _
             "R-sorok (valtozatlanok): " & nR & vbLf
    If Len(ujTipusok) > 0 Then uzenet = uzenet & vbLf & _
        "Uj tipuskod a Type tablaban:" & vbLf & ujTipusok

    MsgBox uzenet, vbInformation, "Cikkszamtorzs frissitese"
    Exit Sub

Vege:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
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
           " (" & hSzam & ")", vbExclamation, "Cikkszamtorzs frissitese"
End Sub

'--------------------------------------------------------------------------
' tblCikkszam -> torzs(1..n, 1..7), az Article oszlopsorrendjeben:
' PartNum, PartName, Type, Forras, Statusz, Berendezes, Ellenorzes.
' Kimarad: Csoport = "R" (LBL), "eltunt a forrasbol" statusz, az Article-ben
' mar R-sorkent meglevo kulcs (nRUtkozes) es az ismetlodo kulcs (nDupla).
' dTorzs(kulcs) = sorindex a torzs tombben.
'--------------------------------------------------------------------------
Private Sub TorzsTablaBeolvasasa(src As Workbook, dR As Object, dTorzs As Object, _
                                 torzs() As Variant, ByRef nTorzs As Long, _
                                 ByRef nKihagyott As Long, ByRef nRUtkozes As Long, _
                                 ByRef nDupla As Long)

    Dim lo As ListObject, v As Variant, db As Variant
    Dim cNum As Long, cNev As Long, cTip As Long, cStat As Long
    Dim cBer As Long, cCsop As Long, cEll As Long
    Dim r As Long, k As String, csop As String, stat As String, hiany As String

    Set lo = TablaBarmelyLapon(src, T_CIKKSZAM)
    If lo Is Nothing Then
        Err.Raise vbObjectError + 540, , _
            "A fajlban (" & src.Name & ") nem talalhato 'tblCikkszam' tabla. " & _
            "Az Autoblatt_Cikkszamtorzs fajlt valasztottad?"
    End If
    If lo.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 541, , "A 'tblCikkszam' tabla ures."
    End If

    ' PONTOS fejlec-egyezes (NormB): a "Tipus (forras)", "Tipus (kezi)",
    ' "Statusz (forras)" stb. normalizalt neve mas, igy azok nem talalnak.
    cNum = OszlopFejlecAlapjan(lo, Array("CIKKSZAM"))
    cNev = OszlopFejlecAlapjan(lo, Array("MEGNEVEZES"))
    cTip = OszlopFejlecAlapjan(lo, Array("TIPUS"))
    cStat = OszlopFejlecAlapjan(lo, Array("STATUSZ"))
    cBer = OszlopFejlecAlapjan(lo, Array("BERENDEZES"))
    cCsop = OszlopFejlecAlapjan(lo, Array("CSOPORT"))
    cEll = OszlopFejlecAlapjan(lo, Array("ELLENORZESDB"))
    If cNum = 0 Then hiany = hiany & " Cikkszam,"
    If cNev = 0 Then hiany = hiany & " Megnevezes,"
    If cTip = 0 Then hiany = hiany & " Tipus,"
    If cStat = 0 Then hiany = hiany & " Statusz,"
    If cBer = 0 Then hiany = hiany & " Berendezes,"
    If cCsop = 0 Then hiany = hiany & " Csoport,"
    If cEll = 0 Then hiany = hiany & " Ellenorzes (db),"
    If Len(hiany) > 0 Then
        Err.Raise vbObjectError + 543, , _
            "A 'tblCikkszam' tablabol hianyzo oszlop(ok):" & Left$(hiany, Len(hiany) - 1)
    End If

    ' Egyszerre olvassuk be (a kepletes oszlopok tarolt erteke, .Value)
    v = lo.DataBodyRange.Value
    ReDim torzs(1 To UBound(v, 1), 1 To 7)
    nTorzs = 0

    For r = 1 To UBound(v, 1)
        k = Kulcs(Szoveg(v(r, cNum)))
        If Len(k) = 0 Then GoTo KovSor

        csop = UCase$(Szoveg(v(r, cCsop)))
        If Len(csop) = 0 Then csop = Left$(k, 1)        ' tartalek: a cikkszam betuje
        If csop = "R" Then GoTo KovSor                  ' R-szam: az LBL-bol jon

        stat = Szoveg(v(r, cStat))
        If InStr(1, NormB(stat), "ELTUNT", vbTextCompare) > 0 Then
            nKihagyott = nKihagyott + 1
            GoTo KovSor
        End If
        If dR.Exists(k) Then
            nRUtkozes = nRUtkozes + 1
            GoTo KovSor
        End If
        If dTorzs.Exists(k) Then
            nDupla = nDupla + 1
            GoTo KovSor
        End If

        nTorzs = nTorzs + 1
        dTorzs.Add k, nTorzs
        torzs(nTorzs, 1) = CikkszamFormatum(Szoveg(v(r, cNum)))
        torzs(nTorzs, 2) = Szoveg(v(r, cNev))
        torzs(nTorzs, 3) = Szoveg(v(r, cTip))
        torzs(nTorzs, 4) = csop
        If StrComp(stat, STAT_KIFUTOTT, vbTextCompare) = 0 Then
            torzs(nTorzs, 5) = STAT_KIFUTOTT
        Else
            torzs(nTorzs, 5) = StatAktiv()
        End If
        torzs(nTorzs, 6) = BerendezesSzoveg(v(r, cBer))
        torzs(nTorzs, 7) = ""
        db = v(r, cEll)
        If Not IsError(db) Then
            If IsNumeric(db) Then
                If CDbl(db) > 0 Then torzs(nTorzs, 7) = TorzsEllenorizendoSzoveg()
            End If
        End If
KovSor:
    Next r

End Sub

' Fajlvalaszto a cikkszamtorzshoz: az utoljara hasznalt fajlrol, ha nincs,
' az LBL fajl mappajabol, vegul a 7.50 TEAM mappabol indul.
Private Function TorzsFajlValasztas() As String

    Dim kezdo As String, mappa As String

    kezdo = UtvonalOlvasas(NEV_TORZS_UTVONAL)
    If Len(kezdo) > 0 Then
        If Not UtvonalLetezik(kezdo, vbNormal) Then kezdo = ""
    End If
    If Len(kezdo) = 0 Then
        mappa = UtvonalOlvasas(NEV_UTVONAL)
        If InStrRev(mappa, "\") > 0 Then
            mappa = Left$(mappa, InStrRev(mappa, "\"))
            If UtvonalLetezik(mappa, vbDirectory) Then kezdo = mappa
        End If
    End If
    If Len(kezdo) = 0 Then
        If UtvonalLetezik(TERMEKADATOK_MAPPA, vbDirectory) Then kezdo = TERMEKADATOK_MAPPA
    End If

    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Cikkszamtorzs (Autoblatt_Cikkszamtorzs) kivalasztasa"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel", "*.xlsx;*.xlsm"
        If Len(kezdo) > 0 Then .InitialFileName = kezdo
        If .Show = -1 Then TorzsFajlValasztas = .SelectedItems(1)
    End With

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

    ' v9.15: a Lista a tarolt, egyseges format mutatja ("A177 637 0100"), a
    ' 9.6-os SUBSTITUTE kivezetve. A gepeles kozbeni szures (M365) a szovegben
    ' barhol keres, de a szokozon nem ugrik at: "637" es "637 0100" talal,
    ' "6370100" nem. A nagy-/kisbetu nem szamit. Kivalasztas utan a
    ' CikkszamValasztas az egyseges format irja A6-ba.
    f = "=[@" & C_NUM & "]" & _
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
        uj(n, 1) = CikkszamFormatum(t(0))
        uj(n, 2) = t(1)
        uj(n, 3) = t(2)
        uj(n, 4) = "R"
        If t(3) Then
            uj(n, 5) = StatAktiv(): nAktiv = nAktiv + 1
        Else
            uj(n, 5) = STAT_KIFUTOTT: nKifutott = nKifutott + 1
        End If
        uj(n, 6) = t(4)          ' Berendezes: tetelenkent, a sajat tablabol
        If CikkszamErvenyes(t(0)) Then uj(n, 7) = "" Else uj(n, 7) = NemSzabvanyosSzoveg()
    Next idx
    nUj = n

    ArticleIras lo, oszlopok, uj, nUj, nRegi

End Sub

' v9.16: az ArticleRSorokCsereje tukorkepe: az R-sorok valtozatlanok, a
' nem R (A/N) sorok helyere a cikkszamtorzs sorai kerulnek. A torzs tomb
' oszlopsorrendje azonos az oszlopok tombevel.
Private Sub ArticleNemRSorokCsereje(torzs() As Variant, ByVal nTorzs As Long)

    Dim lo As ListObject
    Dim oszlopok As Variant, regi() As Variant, uj() As Variant
    Dim nCol As Long, nRegi As Long, n As Long
    Dim i As Long, j As Long
    Dim vSrc As Variant

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

    ' Uj tomb: megtartott R-sorok + a torzs A/N-sorai
    ReDim uj(1 To nRegi + nTorzs, 1 To nCol)
    n = 0
    For i = 1 To nRegi
        If StrComp(Szoveg(vSrc(i, 1)), "R", vbTextCompare) = 0 Then
            n = n + 1
            For j = 1 To nCol
                uj(n, j) = regi(j)(i, 1)
            Next j
        End If
    Next i
    For i = 1 To nTorzs
        n = n + 1
        For j = 1 To nCol
            uj(n, j) = torzs(i, j)
        Next j
    Next i

    ArticleIras lo, oszlopok, uj, n, nRegi

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
' - minden cikkszam (R is) egyseges formara: betu + 3-3-4
'   ("A 177 611 13 00" -> "A177 611 1300"), az N-szam valtozatlan
' - duplikatum (Kulcs() szerint): az elso marad; eltero megnevezesnel a
'   tobbi az Ellenorzes oszlopba kerul, hogy kezzel el lehessen donteni
' - hianyzo megnevezes es nem szabvanyos cikkszam jelolese az Ellenorzes
'   oszlopban (az N-szam nem kap jelolest)
'==========================================================================
Public Sub ArticleTisztitas()

    Const MAX_UTKOZO_LISTA As Long = 15

    Dim lo As ListObject
    Dim oszlopok As Variant, regi() As Variant, uj() As Variant
    Dim nCol As Long, nRegi As Long, n As Long, i As Long, j As Long
    Dim d As Object, dUtkozo As Object, kUtkozo As Variant
    Dim k As String, nev As String, eredeti As String, jegyzet As String, hely As Long
    Dim nForma As Long, nAzonos As Long, nUtkozo As Long, nUres As Long
    Dim nNemSzabvanyos As Long, nListazott As Long
    Dim uzenet As String, lepes As String

    On Error GoTo Hiba

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    lepes = "Article oszlopok ellenorzese"
    ArticleOszlopokBiztositasa

    lepes = "Article beolvasasa"
    Set lo = ThisWorkbook.Worksheets(SH_PARTS).ListObjects(LO_ARTICLE)
    If lo.DataBodyRange Is Nothing Then GoTo Vege
    nRegi = lo.DataBodyRange.Rows.Count

    oszlopok = Array(C_NUM, C_NAME, C_TYPE, C_SRC, C_STAT, C_PROD, C_NOTE)
    nCol = UBound(oszlopok) + 1
    ReDim regi(1 To nCol)
    For j = 1 To nCol
        regi(j) = OszlopTomb(lo, CStr(oszlopok(j - 1)))
    Next j

    lepes = "egysegesites es duplikatumok"
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set dUtkozo = CreateObject("Scripting.Dictionary")
    dUtkozo.CompareMode = vbTextCompare
    ReDim uj(1 To nRegi, 1 To nCol)
    n = 0

    For i = 1 To nRegi
        ' v9.15: minden sor (R is) egyseges formara, a kulcs mindenhol Kulcs()
        eredeti = Szoveg(regi(1)(i, 1))
        k = Kulcs(eredeti)
        If StrComp(CikkszamFormatum(eredeti), eredeti, vbBinaryCompare) <> 0 Then nForma = nForma + 1
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
                If Not dUtkozo.Exists(k) Then dUtkozo.Add k, hely
                If InStr(1, Szoveg(uj(hely, 7)), nev, vbTextCompare) = 0 Then
                    uj(hely, 7) = Szoveg(uj(hely, 7)) & _
                        IIf(Len(Szoveg(uj(hely, 7))) > 0, "; ", MasMegnevezesSzoveg()) & nev
                End If
            End If
        Else
            n = n + 1
            d.Add k, n
            For j = 1 To nCol
                uj(n, j) = regi(j)(i, 1)
            Next j
            uj(n, 1) = CikkszamFormatum(eredeti)
        End If
    Next i

    lepes = "jelolesek"
    For i = 1 To n
        If Len(Szoveg(uj(i, 2))) = 0 And Len(Szoveg(uj(i, 7))) = 0 Then
            uj(i, 7) = "Nincs megnevezes": nUres = nUres + 1
        End If
        ' v9.15: nem szabvanyos cikkszam jelolese (az N-szam nem kap jelolest)
        If Not CikkszamErvenyes(Szoveg(uj(i, 1))) And Not (Kulcs(Szoveg(uj(i, 1))) Like "N#*") Then
            nNemSzabvanyos = nNemSzabvanyos + 1
            jegyzet = Szoveg(uj(i, 7))
            If InStr(1, Ekezetlen(jegyzet), "Nem szabvanyos cikkszam", vbTextCompare) = 0 Then
                uj(i, 7) = jegyzet & IIf(Len(jegyzet) > 0, "; ", "") & NemSzabvanyosSzoveg()
            End If
        End If
    Next i

    lepes = "elonezet"
    Application.ScreenUpdating = True
    uzenet = "Article tisztitas - elonezet" & vbLf & vbLf & _
             "Sorok most: " & nRegi & vbLf & _
             "Sorok utana: " & n & vbLf & vbLf & _
             "Egyseges formara hozva (A177 637 0100): " & nForma & vbLf & _
             "Torolt duplikatum (azonos megnevezes): " & nAzonos & vbLf & _
             "Utkozo megnevezes (Ellenorzes oszlopba irva): " & nUtkozo & vbLf & _
             "Hianyzo megnevezes (jelolve): " & nUres & vbLf & _
             "Nem szabvanyos cikkszam (jelolve): " & nNemSzabvanyos & vbLf
    ' Csak a cikkszamok (a MsgBox kb. 1024 karakter utan levag)
    If dUtkozo.Count > 0 Then
        uzenet = uzenet & vbLf & "Utkozo megnevezesu cikkszamok:" & vbLf
        For Each kUtkozo In dUtkozo.Keys
            If nListazott >= MAX_UTKOZO_LISTA Then
                uzenet = uzenet & "  " & ChrW(8230) & " " & ChrW(233) & "s tov" & ChrW(225) & "bbi " & _
                         (dUtkozo.Count - MAX_UTKOZO_LISTA) & vbLf
                Exit For
            End If
            uzenet = uzenet & "  " & Szoveg(uj(dUtkozo(kUtkozo), 1)) & vbLf
            nListazott = nListazott + 1
        Next kUtkozo
    End If
    uzenet = uzenet & vbLf & "Vegrehajtsam? (Elotte erdemes elmenteni egy masolatot.)"
    If MsgBox(uzenet, vbQuestion + vbYesNo + vbDefaultButton2, "Article tisztitas") <> vbYes Then GoTo Vege
    Application.ScreenUpdating = False

    lepes = "Article ujrairasa"
    ArticleIras lo, oszlopok, uj, n, nRegi
    lepes = "A6 lista ujraepitese"
    A6ListaFrissites

Vege:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

Hiba:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "A tisztitas megszakadt [" & lepes & "]: " & Err.Description & _
           " (" & Err.Number & ")", vbExclamation, "Article tisztitas"
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
    ' egy keresosav is van benne. A lista a tarolt, egyseges cikkszam-format
    ' mutatja (v9.15), a szuresnel a szokoz szamit (l. ListaKepletBeallitas).

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

    ' v9.17: ehhez az allapothoz kepest nezi a lapra lepes, kell-e ujraepites
    mA6ListaAlairas = A6ListaAlairas()

End Sub

'--------------------------------------------------------------------------
' v9.17: az Article "alairasa": sorszam | a Lista oszlop osszhossza.
' Nativ Excel-szamitas (gyors), a Parts lap Evaluate-jevel, hogy ne fuggjon
' az aktiv munkafuzettol. Hiba (pl. ures tabla, hibaertek a Lista-ban)
' eseten ures sztring - ilyenkor a hivo inkabb ujraepit.
' Korlat: az azonos hosszusagu atirast (pl. "LI" -> "RE") nem eszleli.
'--------------------------------------------------------------------------
Public Function A6ListaAlairas() As String
    Dim v As Variant
    On Error GoTo Hiba
    v = ThisWorkbook.Worksheets(SH_PARTS).Evaluate( _
            "ROWS(" & LO_ARTICLE & ")&""|""&SUMPRODUCT(LEN(" & LO_ARTICLE & "[" & C_LIST & "]))")
    If IsError(v) Then Exit Function
    A6ListaAlairas = CStr(v)
    Exit Function
Hiba:
    A6ListaAlairas = ""
End Function

' v9.17: True, ha az A6 lista ujraepitese szukseges (az Article valtozott
' az utolso epites ota, az alairas nem kepezheto, vagy VBA-reset volt).
Public Function A6ListaElavult() As Boolean
    Dim alairas As String
    alairas = A6ListaAlairas()
    A6ListaElavult = (Len(alairas) = 0) Or (alairas <> mA6ListaAlairas)
End Function

'==========================================================================
' 5. A6 valasztas -> A6 = cikkszam, C6 = megnevezes, G6 = tipus
'==========================================================================
Public Sub CikkszamValasztas(ws As Worksheet)

    Dim s As String, szam As String, sep As String
    Dim p As Long

    s = Szoveg(ws.Range("A6").Value)
    If Len(s) = 0 Then
        TetelAdatokUritese ws               ' v9.17: C6, G6 es H6 is
        Exit Sub
    End If

    ' "R244 637 0100 . BPL SEITENWAND LI . X244/248 (kifutott)" -> szam
    sep = " " & ChrW(183) & " "
    p = InStr(1, s, sep)
    If p > 0 Then szam = Left$(s, p - 1) Else szam = s
    If Right$(szam, 11) = " (kifutott)" Then szam = Left$(szam, Len(szam) - 11)
    szam = Trim$(szam)

    If Not CikkszamBeallitasa(ws, Kulcs(szam)) Then
        ' Nincs a torzsben: a szam marad (egyseges formaban), a tetel
        ' adatai (C6, G6, H6) kiurulnek - v9.17
        ws.Range("A6").Value = CikkszamFormatum(szam)
        TetelAdatokUritese ws
    End If

End Sub

' A tenyleges "A6/C6/G6/H6 kitoltese egy tudott Kulcs() alapjan" logika
' kulon Public Function-ba kiemelve. True, ha talalt es kitoltott egy tetelt.
Public Function CikkszamBeallitasa(ws As Worksheet, ByVal k As String) As Boolean

    Dim lo As ListObject
    Dim i As Long
    Dim vNum As Variant, vName As Variant, vType As Variant, vBer As Variant
    Dim vanTipus As Boolean, vanBer As Boolean
    Dim ujBer As String

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
            ws.Range("A6").Value = CikkszamFormatum(Szoveg(vNum(i, 1)))
            ws.Range("C6").Value = Szoveg(vName(i, 1))
            ' v9.16: G6 mindig a tetel tipusa; ha az ures, a G6 is kiurul
            ' (kulonben egy korabbi tipus, pl. W177 maradna a lapon)
            If vanTipus Then ws.Range("G6").Value = Szoveg(vType(i, 1))

            ' v9.3: a gyarto berendezes (H6) automatikus at-/kitoltese a tetel
            ' sajat Berendezes ertekebol, ha az elter a jelenlegi H6-tol
            ' v9.16: a 9.12-es "A-szam = mindig 7.41" szabaly KIVEZETVE: minden
            ' cikkszamnal a tetel sajat Berendezes erteke szamit. Ha ures, a
            ' H6_URES_HA_NINCS_BERENDEZES dont: True = H6 kiurul, False = marad.
            ' v9.17: a H6 kezelese egy helyen, a BerendezesBeallitas-ban.
            If vanBer Then
                ujBer = Szoveg(vBer(i, 1))
            Else
                ujBer = ""
            End If
            BerendezesBeallitas ws, ujBer

            ' v9.13: R-cikkszamnal E4/E6 = Alapanyag/Material - Platine;
            ' v9.17: a logika egy helyen, az AnyagFeliratBeallitas-ban
            AnyagFeliratBeallitas ws, Szoveg(vNum(i, 1))

            ' v9.17: nincs A6ListaFrissites - a valasztas nem valtoztatja a listat
            CikkszamBeallitasa = True
            Exit Function
        End If
    Next i

End Function

' v9.17: ismeretlen vagy torolt cikkszam - a korabbi tetel adatai ne
' maradjanak a lapon: C6 es G6 kiurul, a H6 a BerendezesBeallitas szerint
' (H6_URES_HA_NINCS_BERENDEZES = True: kiurul, a kerdesblokk elrejtodik).
Private Sub TetelAdatokUritese(ws As Worksheet)
    ws.Range("C6").Value = ""
    ws.Range("G6").Value = ""
    BerendezesBeallitas ws, ""
End Sub

' v9.17: a H6 (gyarto berendezes) beallitasa, egy helyen a
' CikkszamBeallitasa es a TetelAdatokUritese szamara.
'   ujBer nem ures: ha elter a H6-tol, beirjuk, es a kerdesblokk ujratoltodik
'   ujBer ures:     H6_URES_HA_NINCS_BERENDEZES = True -> a H6 kiurul (ha nem
'                   volt ures), a kerdesblokk elrejtodik; False -> a H6 marad
Private Sub BerendezesBeallitas(ws As Worksheet, ByVal ujBer As String)

    Dim regiBer As String

    ujBer = Trim$(Replace(ujBer, ",", "."))
    regiBer = Trim$(Replace(Szoveg(ws.Range("H6").Value), ",", "."))

    If Len(ujBer) > 0 Then
        If StrComp(regiBer, ujBer, vbTextCompare) <> 0 Then
            ws.Range("H6").Value = ujBer
            HibafeltarasKitoltes ws
        End If
    ElseIf H6_URES_HA_NINCS_BERENDEZES Then
        If Len(regiBer) > 0 Then
            ws.Range("H6").Value = ""
            HibafeltarasKitoltes ws          ' a kerdesblokk elrejtodik
        End If
    End If

End Sub

'--------------------------------------------------------------------------
' v9.17: E4/E6 a cikkszam elso betuje alapjan. Ez az EGYETLEN hely, ahol ez
' a logika el: a lapkod (AnyagFelirat) es a CikkszamBeallitasa is ezt hivja.
'   R-szam:     E4 = "Alapanyag/Material", E6 = "Platine"
'   minden mas: E4 = eredeti felirat ("Preseles/Abpressung", ekezetekkel),
'               az E6-bol az automatikus "Platine" torlodik (a felhasznalo
'               sajat E6 bejegyzeset nem bantjuk)
'--------------------------------------------------------------------------
Public Sub AnyagFeliratBeallitas(ws As Worksheet, ByVal cikkszam As String)
    If UCase$(Left$(Trim$(cikkszam), 1)) = "R" Then
        ws.Range("E4").Value = "Alapanyag/Material"
        ws.Range("E6").Value = "Platine"
    Else
        ws.Range("E4").Value = PreselesFeliratEredeti()
        If StrComp(Szoveg(ws.Range("E6").Value), "Platine", vbTextCompare) = 0 Then
            ws.Range("E6").Value = ""
        End If
    End If
End Sub

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
' 6b. Hibakepek (v9.17): Data!tblHibakep -> HibakepLista -> cimsor-legordulo
'--------------------------------------------------------------------------
' A Hibafeltaras cimsorban a "Hibakep / Fehlerbild" felirat melletti cella
' legordulobol valaszthato, de szabadon felulirhato (ShowError = False).
' A lista a Data lap tblHibakep tablaja; a HibakepekFrissitese a
' Hibamegnevezesek fajl "Bezeichnung HU" oszlopabol boviti (csak hozzaad,
' a kezzel felvett sorok megmaradnak), ABC-sorrendbe rendezve.
'==========================================================================

' Alt+F8 -> HibakepekFrissitese
Public Sub HibakepekFrissitese()

    Dim fajl As String
    Dim src As Workbook, sajatNyitas As Boolean
    Dim nevek As Collection, idx As Variant
    Dim lo As ListObject, dTabla As Object, r As Long
    Dim nUj As Long, nMar As Long, ujLista As String
    Dim uzenet As String, lepes As String

    Const MAX_UJ_LISTA As Long = 15

    On Error GoTo Hiba

    lepes = "fajl kivalasztasa"
    fajl = HibakepFajlValasztas()
    If Len(fajl) = 0 Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    '--- Forras megnyitasa (ha mar nyitva van, azt hasznaljuk) ------------
    lepes = "forrasfajl megnyitasa"
    Set src = ForrasMegnyitasa(fajl, sajatNyitas)       ' v9.17

    lepes = "magyar nevek beolvasasa"
    Set nevek = HibakepNevekBeolvasasa(src)

    If sajatNyitas Then src.Close SaveChanges:=False
    Set src = Nothing

    If nevek.Count = 0 Then
        Err.Raise vbObjectError + 550, , _
            "A fajlban nem talalhato 'Bezeichnung HU' oszlop, vagy nincs alatta magyar nev."
    End If

    '--- tblHibakep bovitese (csak hozzaad) --------------------------------
    lepes = "tblHibakep bovitese"
    Set lo = HibakepTablaBiztositasa()
    Set dTabla = CreateObject("Scripting.Dictionary")
    dTabla.CompareMode = vbTextCompare
    If Not lo.DataBodyRange Is Nothing Then
        For r = 1 To lo.ListRows.Count
            If Len(Szoveg(lo.DataBodyRange.Cells(r, 1).Value)) > 0 Then
                dTabla(Szoveg(lo.DataBodyRange.Cells(r, 1).Value)) = True
            End If
        Next r
    End If

    For Each idx In nevek
        If dTabla.Exists(idx) Then
            nMar = nMar + 1
        Else
            HibakepSorHozzaadasa lo, CStr(idx)
            dTabla.Add idx, True
            nUj = nUj + 1
            If nUj <= MAX_UJ_LISTA Then ujLista = ujLista & "  " & idx & vbLf
        End If
    Next idx
    If nUj > MAX_UJ_LISTA Then ujLista = ujLista & TovabbiSzoveg(nUj - MAX_UJ_LISTA)

    lepes = "rendezes"
    If lo.ListRows.Count > 1 Then
        With lo.Sort
            .SortFields.Clear
            .SortFields.Add Key:=lo.ListColumns(1).DataBodyRange, _
                            SortOn:=xlSortOnValues, Order:=xlAscending
            .Header = xlYes
            .Apply
        End With
    End If

    lepes = "legordulo beallitasa"
    HibakepLegorduloBeallitas ThisWorkbook.Worksheets(SH_FOTO)

    lepes = "utvonal mentese"
    UtvonalMentes fajl, NEV_HIBAKEP_UTVONAL

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    uzenet = "Forras: " & Mid$(fajl, InStrRev(fajl, "\") + 1) & vbLf & vbLf & _
             "Magyar nev a forrasban: " & nevek.Count & vbLf & _
             "Mar szerepelt: " & nMar & vbLf & _
             "Uj hibakep: " & nUj & vbLf & ujLista & vbLf & _
             "Hibakepek a listaban osszesen: " & lo.ListRows.Count
    MsgBox uzenet, vbInformation, "Hibakepek frissitese"
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
           " (" & hSzam & ")", vbExclamation, "Hibakepek frissitese"
End Sub

' A forrasfajl magyar nevei, a forras sorrendjeben, ismetles nelkul.
' Fejlec: az elso 10 sor / 20 oszlop "Bezeichnung HU..." kezdetu cellaja
' (NormB). A lista az elso teljesen ures sorig tart (utana mar csak
' megjegyzes jon, pl. "Anmerkung: ..."); a magyar nev nelkuli sor kimarad.
Private Function HibakepNevekBeolvasasa(src As Workbook) As Collection

    Dim ws As Worksheet, nevek As Collection, dVan As Object
    Dim hdrSor As Long, cHu As Long, r As Long, c As Long, utSor As Long
    Dim nev As String

    Set nevek = New Collection
    Set dVan = CreateObject("Scripting.Dictionary")
    dVan.CompareMode = vbTextCompare

    For Each ws In src.Worksheets
        hdrSor = 0: cHu = 0
        For r = 1 To 10
            For c = 1 To 20
                If Left$(NormB(Szoveg(ws.Cells(r, c).Value)), 13) = "BEZEICHNUNGHU" Then
                    hdrSor = r: cHu = c
                    Exit For
                End If
            Next c
            If cHu > 0 Then Exit For
        Next r

        If cHu > 0 Then
            utSor = ws.Cells(ws.Rows.Count, cHu).End(xlUp).Row
            For r = hdrSor + 1 To utSor
                If Application.WorksheetFunction.CountA(ws.Rows(r)) = 0 Then Exit For
                nev = Szoveg(ws.Cells(r, cHu).Value)
                If Len(nev) > 0 And Left$(NormB(nev), 9) <> "ANMERKUNG" Then
                    If Not dVan.Exists(nev) Then
                        dVan.Add nev, True
                        nevek.Add nev
                    End If
                End If
            Next r
            Exit For            ' az elso talalt lap utan nem keresunk tovabb
        End If
    Next ws

    Set HibakepNevekBeolvasasa = nevek

End Function

' Egy uj sor a tblHibakep vegere (egy uj, ures tabla elso sorat hasznalja).
Private Sub HibakepSorHozzaadasa(lo As ListObject, ByVal nev As String)
    If Not lo.DataBodyRange Is Nothing Then
        If lo.ListRows.Count = 1 And Len(Szoveg(lo.DataBodyRange.Cells(1, 1).Value)) = 0 Then
            lo.DataBodyRange.Cells(1, 1).Value = nev
            Exit Sub
        End If
    End If
    lo.ListRows.Add.Range.Cells(1, 1).Value = nev
End Sub

' A tblHibakep tabla (ha nincs, a Data lapon letrehozza, a hasznalt terulet
' utan egy ures oszloppal) es a HibakepLista nev biztositasa.
Public Function HibakepTablaBiztositasa() As ListObject

    Dim ws As Worksheet, lo As ListObject, oszlop As Long

    Set lo = TablaBarmelyLapon(ThisWorkbook, T_HIBAKEP)
    If lo Is Nothing Then
        Set ws = ThisWorkbook.Worksheets(SH_DATA)
        With ws.UsedRange
            oszlop = .Column + .Columns.Count + 1
        End With
        ws.Cells(1, oszlop).Value = C_HIBAKEP
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Cells(1, oszlop).Resize(2, 1), , xlYes)
        lo.Name = T_HIBAKEP
    End If

    On Error Resume Next
    ThisWorkbook.Names(NEV_HIBAKEPLISTA).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=NEV_HIBAKEPLISTA, _
        RefersTo:="=" & lo.Name & "[" & lo.ListColumns(1).Name & "]"

    Set HibakepTablaBiztositasa = lo

End Function

' A Hibafeltaras cimsorban a "Hibakep" kezdetu felirat (most E8:F8) utani
' egyesitett cellara (most G8:I8) legordulot tesz: HibakepLista, de a
' listan kivuli ertek is beirhato. True, ha beallitotta.
Public Function HibakepLegorduloBeallitas(ws As Worksheet) As Boolean

    Dim cimSor As Long, c As Long
    Dim terulet As Range, cel As Range

    cimSor = SorKereses(ws, H_CIM, 2, 200)
    If cimSor = 0 Then Exit Function

    c = 1
    Do While c <= 9
        Set terulet = ws.Cells(cimSor, c).MergeArea
        If UCase$(Left$(Ekezetlen(Szoveg(terulet.Cells(1, 1).Value)), 7)) = "HIBAKEP" Then
            c = terulet.Column + terulet.Columns.Count
            If c > 9 Then Exit Function
            Set cel = ws.Cells(cimSor, c).MergeArea

            On Error Resume Next
            cel.Validation.Delete
            On Error GoTo 0
            With cel.Validation
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                     Operator:=xlBetween, Formula1:="=" & NEV_HIBAKEPLISTA
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowError = False              ' listan kivuli ertek is beirhato
            End With
            HibakepLegorduloBeallitas = True
            Exit Function
        End If
        c = terulet.Column + terulet.Columns.Count
    Loop

End Function

' Fajlvalaszto: az utoljara hasznalt fajlrol, ha nincs, a Fotoblatt
' mappajabol indul.
Private Function HibakepFajlValasztas() As String

    Dim kezdo As String

    kezdo = UtvonalOlvasas(NEV_HIBAKEP_UTVONAL)
    If Len(kezdo) > 0 Then
        If Not UtvonalLetezik(kezdo, vbNormal) Then kezdo = ""
    End If
    If Len(kezdo) = 0 Then
        If UtvonalLetezik(ThisWorkbook.Path & "\", vbDirectory) Then kezdo = ThisWorkbook.Path & "\"
    End If

    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Hibamegnevezesek fajl kivalasztasa"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel", "*.xls;*.xlsx;*.xlsm"
        If Len(kezdo) > 0 Then .InitialFileName = kezdo
        If .Show = -1 Then HibakepFajlValasztas = .SelectedItems(1)
    End With

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

'--------------------------------------------------------------------------
' v9.15: egyseges cikkszam-formatum: betu + 3-3-4, pl. "A177 637 0100".
' Barmilyen bevitelt elfogad (szokoz, NBSP, tab, kisbetu nem szamit).
' Ami nem illik a mintara (pl. N-szam), valtozatlanul (Trim) marad.
'--------------------------------------------------------------------------
Public Function CikkszamFormatum(ByVal s As String) As String
    Dim k As String
    k = Kulcs(s)
    If k Like "[AR]##########" Then
        CikkszamFormatum = Left$(k, 4) & " " & Mid$(k, 5, 3) & " " & Mid$(k, 8, 4)
    Else
        CikkszamFormatum = Trim$(s)
    End If
End Function

' True, ha a cikkszam a szabvanyos A/R mintara illik
Public Function CikkszamErvenyes(ByVal s As String) As Boolean
    CikkszamErvenyes = (Kulcs(s) Like "[AR]##########")
End Function

Private Function StatAktiv() As String
    StatAktiv = "Akt" & ChrW(237) & "v"
End Function

' "Nem szabvanyos cikkszam", ekezetekkel (Ellenorzes oszlop)
Private Function NemSzabvanyosSzoveg() As String
    NemSzabvanyosSzoveg = "Nem szabv" & ChrW(225) & "nyos cikksz" & ChrW(225) & "m"
End Function

' "Mas megnevezes is: ", ekezetekkel (Ellenorzes oszlop)
Private Function MasMegnevezesSzoveg() As String
    MasMegnevezesSzoveg = "M" & ChrW(225) & "s megnevez" & ChrW(233) & "s is: "
End Function

' "Ellenorizendo (cikkszamtorzs)", ekezetekkel (Ellenorzes oszlop), v9.16
Private Function TorzsEllenorizendoSzoveg() As String
    TorzsEllenorizendoSzoveg = "Ellen" & ChrW(337) & "rizend" & ChrW(337) & _
                               " (cikksz" & ChrW(225) & "mt" & ChrW(246) & "rzs)"
End Function

' "  ... es tovabbi N" sor az elonezeti listak vegere, v9.16
Private Function TovabbiSzoveg(ByVal db As Long) As String
    TovabbiSzoveg = "  " & ChrW(8230) & " " & ChrW(233) & "s tov" & ChrW(225) & "bbi " & db & vbLf
End Function

' Berendezes szovegkent, tizedesponttal, v9.16: a szam 7.5 -> "7.50",
' a szoveg "7,41" -> "7.41" (a helyi tizedesvesszotol fuggetlenul)
Private Function BerendezesSzoveg(ByVal v As Variant) As String
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbCurrency, vbDecimal, vbInteger, vbLong
            BerendezesSzoveg = Replace(Format$(v, "0.00"), ",", ".")
        Case Else
            BerendezesSzoveg = Trim$(Replace(Szoveg(v), ",", "."))
    End Select
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
