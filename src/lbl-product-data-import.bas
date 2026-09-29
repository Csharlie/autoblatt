Attribute VB_Name = "modTermekadatok"
' Keszitette es karbantartja : Sardy Peter (PESARDY)
' Mercedes-Benz Manufacturing Hungary, Kecskemet - PW
' LBL termekadatok - sajat valtozat, frissites a forrasfajlbol
' Verzio 1.2 - 2026.09.23.
'   1.2: 'Forras kiemeles' oszlop (a forras sarga/piros jelolesei szovegkent);
'        forras-szinek atvetele csak ha Info!nmForrasSzinek = Igen; gomb helye E1
'   1.1: Sorozat = sajat oszlop (csak ures cellat tolt), uj 'Sorozat (forras)' oszlop;
'        kis/nagybetu es szelso szokoz elteres nem szamit valtozasnak
'
' FONTOS: ez a fajl szandekosan CSAK ASCII karaktereket tartalmaz (a VBE a
' .bas fajlt a rendszer kodlapjaval olvassa be). Az ekezetes szovegek a U()
' fuggvennyel keszulnek: U("Akt\u00edv") = "Aktiv" ekezettel.
'
' A FORRASFAJL SOHA NEM VALTOZIK: csak olvasasra nyitjuk meg, csak olvasunk
' belole, es mentes nelkul zarjuk be. Ha mar nyitva volt, nyitva marad,
' es nem nyulunk hozza.
'
' Mit ir a makro a sajat tablaban (tblTermek)?
'   - a Lekepezes lap tblOszlop tablajaban felsorolt forras-oszlopokat
'   - Termekszam (uj tetelnel), Statusz, Sorozat (forras), Adatjelzes,
'     Elso forras, Utolso forras, Utolso valtozas
' Mit NEM ir soha: a sajat oszlopokat (Tipus (kezi), db/utem, Rahagyas,
' Sajat megjegyzes) es a kepletes oszlopokat. A Sorozat oszlopot csak akkor
' tolti ki, ha ures (a Lekepezes lap tblSorozat parositasa alapjan).
'
' Belepesi pontok:
'   TermekadatokFrissitese    - forrasfajl kivalasztasa, elonezet, frissites
'   FrissitoGombLetrehozasa   - gomb a Termekadatok lapra (egyszer kell)

Option Explicit
Option Compare Binary

'--- Tabla- es nevneveket egy helyen ------------------------------------
Private Const T_TERMEK   As String = "tblTermek"
Private Const T_OSZLOP   As String = "tblOszlop"
Private Const T_SOROZAT  As String = "tblSorozat"
Private Const T_NAPLO    As String = "tblNaplo"
Private Const T_FIGY     As String = "tblFigy"
Private Const SH_MENTES  As String = "Elozo_allapot"
Private Const BTN_NEV    As String = "btnFrissites"

'--- Meta oszlopok a sajat tablaban (Norm alakban) -----------------------
Private Const M_TERMEKSZAM As String = "TERMEKSZAM"
Private Const M_STATUSZ    As String = "STATUSZ"
Private Const M_SOROZAT    As String = "SOROZAT"
Private Const M_SORFORRAS  As String = "SOROZATFORRAS"
Private Const M_KIEMELES   As String = "FORRASKIEMELES"
Private Const M_JELZES     As String = "ADATJELZES"
Private Const M_ELSO       As String = "ELSOFORRAS"
Private Const M_UTOLSO     As String = "UTOLSOFORRAS"
Private Const M_VALTOZAS   As String = "UTOLSOVALTOZAS"

Private Const MAX_KERES As Long = 40        ' fejleckereses sor/oszlop hatar
Private Const NINCS_SZIN As Long = -1

'--- Mezoleiras (tblOszlop) -----------------------------------------------
Private nF As Long
Private fCel() As String       ' cel oszlop neve a sajat tablaban
Private fPat() As Variant      ' Norm-alt mintak tombje
Private fNum() As Boolean      ' szam mezo?
Private fCol() As Long         ' oszlopindex a tblTermek-ben
Private keyF As Long           ' Termekszam mezo indexe
Private hlF As Long            ' "Fejlec nelkuli adat" mezo indexe (0 = nincs)

'--- Beolvasott forras ----------------------------------------------------
Private nS As Long
Private sKey() As String
Private sRaw() As Variant
Private sAkt() As Boolean
Private sKif() As Boolean
Private sSor() As String
Private sVal() As Variant      ' (mezo, tetel)
Private sFill() As Long        ' (mezo, tetel)
Private sFont() As Long        ' (mezo, tetel)
Private dictS As Object
Private figy As Collection
Private naplo As Collection
Private forrasFejlec As String

'--- Szamlalok ------------------------------------------------------------
Private cUj As Long, cMod As Long, cMezo As Long, cStat As Long, cElt As Long

'==========================================================================
' Belepesi pont: frissites a forrasfajlbol
'==========================================================================
Public Sub TermekadatokFrissitese()
    Dim f As Variant, fajlNev As String, cimke As String
    Dim wbS As Workbook, wb As Workbook, mar As Boolean
    Dim secRegi As Long, calcRegi As Long, valasz As VbMsgBoxResult
    Dim lo As ListObject, msg As String

    On Error GoTo Hiba
    calcRegi = Application.Calculation

    Set lo = Tabla(T_TERMEK)
    If lo Is Nothing Then Err.Raise vbObjectError + 1, , U("Nem tal\u00e1lhat\u00f3 a tblTermek t\u00e1bla.")
    If Not MezokBetoltese(lo) Then Exit Sub

    f = Application.GetOpenFilename("Excel (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", , _
        U("V\u00e1laszd ki a forr\u00e1s Term\u00e9kadatok f\u00e1jlt (csak olvas\u00e1sra ny\u00edlik meg)"))
    If VarType(f) = vbBoolean Then Exit Sub
    If StrComp(CStr(f), ThisWorkbook.FullName, vbTextCompare) = 0 Then
        MsgBox U("Ez a saj\u00e1t f\u00e1jl, nem a forr\u00e1s."), vbExclamation
        Exit Sub
    End If
    fajlNev = Mid$(CStr(f), InStrRev(CStr(f), Application.PathSeparator) + 1)

    ' Nyitva van mar? (ugyanaz a fajl: hasznaljuk; azonos nevu masik: nem lehet megnyitni)
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, CStr(f), vbTextCompare) = 0 Then
            Set wbS = wb: mar = True: Exit For
        ElseIf StrComp(wb.Name, fajlNev, vbTextCompare) = 0 Then
            MsgBox U("M\u00e1r nyitva van egy azonos nev\u0171 f\u00e1jl (") & wb.Name & _
                U("). Z\u00e1rd be, \u00e9s ind\u00edtsd \u00fajra a friss\u00edt\u00e9st."), vbExclamation
            Exit Sub
        End If
    Next wb

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    If Not mar Then
        secRegi = Application.AutomationSecurity
        Application.AutomationSecurity = 3          ' msoAutomationSecurityForceDisable
        Set wbS = Workbooks.Open(Filename:=CStr(f), UpdateLinks:=0, ReadOnly:=True, _
                                 AddToMru:=False, Notify:=False)
        Application.AutomationSecurity = secRegi
    End If

    cimke = wbS.Name
    ForrasBeolvasasa wbS

    ' A forrast azonnal bezarjuk, mentes nelkul (ha mi nyitottuk meg)
    If Not mar Then wbS.Close SaveChanges:=False
    Set wbS = Nothing
    ThisWorkbook.Activate

    If nS = 0 Then
        MsgBox U("A kiv\u00e1lasztott f\u00e1jlban nem tal\u00e1lhat\u00f3 'Term\u00e9ksz\u00e1m' fejl\u00e9c egyik lapon sem."), vbExclamation
        GoTo Vege
    End If

    ' 1. kor: csak szamolunk (elonezet)
    Set naplo = New Collection
    Osszevetes lo, cimke, False
    Application.ScreenUpdating = True

    msg = U("Forr\u00e1s: ") & cimke & vbLf & U("Fejl\u00e9c: ") & forrasFejlec & vbLf & vbLf & _
          U("T\u00e9telek a forr\u00e1sban: ") & nS & vbLf & _
          U("\u00daj t\u00e9tel: ") & cUj & vbLf & _
          U("M\u00f3dosult t\u00e9tel: ") & cMod & " (" & cMezo & U(" mez\u0151)") & vbLf & _
          U("St\u00e1tuszv\u00e1ltoz\u00e1s: ") & cStat & vbLf & _
          U("Elt\u0171nt a forr\u00e1sb\u00f3l: ") & cElt & vbLf & _
          U("Figyelmeztet\u00e9s: ") & figy.Count & IIf(figy.Count > 0, U(" (r\u00e9szletek az Info lapon)"), "") & vbLf & vbLf & _
          U("Friss\u00edted a saj\u00e1t t\u00e1bl\u00e1t? A forr\u00e1sf\u00e1jl nem v\u00e1ltozik.")
    FigyelmeztetesekKiirasa (False)
    valasz = MsgBox(msg, vbYesNo + vbQuestion, U("Term\u00e9kadatok friss\u00edt\u00e9se \u2013 el\u0151n\u00e9zet"))
    If valasz <> vbYes Then GoTo Vege

    ' 2. kor: alkalmazas
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    AllapotMentese lo
    lo.Parent.Activate
    Set naplo = New Collection
    Osszevetes lo, cimke, True
    KepletekPotlasa lo
    NaploKiirasa cimke
    FigyelmeztetesekKiirasa True
    InfoKiirasa cimke, CStr(f)
    Application.Calculation = calcRegi

    Application.ScreenUpdating = True
    MsgBox U("K\u00e9sz. \u00daj: ") & cUj & U(", m\u00f3dosult: ") & cMod & _
           U(", elt\u0171nt: ") & cElt & U(". R\u00e9szletek a V\u00e1ltoz\u00e1snapl\u00f3 lapon."), vbInformation

Vege:
    If secRegi <> 0 Then Application.AutomationSecurity = secRegi
    Application.Calculation = calcRegi
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

Hiba:
    Dim hibaSzoveg As String
    hibaSzoveg = Err.Description
    On Error Resume Next
    If Not wbS Is Nothing And Not mar Then wbS.Close SaveChanges:=False
    On Error GoTo 0
    MsgBox U("Hiba a friss\u00edt\u00e9s k\u00f6zben: ") & hibaSzoveg & vbLf & _
           U("A forr\u00e1sf\u00e1jl nem v\u00e1ltozott."), vbExclamation
    Resume Vege
End Sub

'==========================================================================
' Mezoleiras betoltese a Lekepezes lap tblOszlop tablajabol
'==========================================================================
Private Function MezokBetoltese(lo As ListObject) As Boolean
    Dim tOsz As ListObject, i As Long, j As Long, p As Variant, n As String
    Dim hiany As String, db As Long
    Dim meta As Variant

    Set tOsz = Tabla(T_OSZLOP)
    If tOsz Is Nothing Then
        MsgBox U("Nem tal\u00e1lhat\u00f3 a tblOszlop t\u00e1bla (Lek\u00e9pez\u00e9s lap)."), vbExclamation
        Exit Function
    End If
    nF = tOsz.ListRows.Count
    ReDim fCel(1 To nF): ReDim fPat(1 To nF): ReDim fNum(1 To nF): ReDim fCol(1 To nF)
    keyF = 0: hlF = 0
    For i = 1 To nF
        fCel(i) = Trim$(SafeStr(tOsz.DataBodyRange.Cells(i, 1).Value))
        n = Norm(fCel(i))
        fNum(i) = (Norm(SafeStr(tOsz.DataBodyRange.Cells(i, 3).Value)) = "IGEN")
        ' mintak: pontosvesszovel elvalasztva, Norm alakban
        Dim darabok() As String, lista() As String
        darabok = Split(SafeStr(tOsz.DataBodyRange.Cells(i, 2).Value), ";")
        ReDim lista(0 To 0): db = 0
        For j = LBound(darabok) To UBound(darabok)
            If Len(Norm(darabok(j))) > 0 And Left$(Trim$(darabok(j)), 1) <> "(" Then
                ReDim Preserve lista(0 To db): lista(db) = Norm(darabok(j)): db = db + 1
            End If
        Next j
        If db = 0 Then fPat(i) = Array() Else fPat(i) = lista
        If n = M_TERMEKSZAM Then keyF = i
        If Left$(n, 13) = "FEJLECNELKULI" Then hlF = i
        fCol(i) = OszlopIndexPontos(lo, fCel(i))
        If fCol(i) = 0 Then hiany = hiany & vbLf & "  - " & fCel(i)
    Next i
    If keyF = 0 Then hiany = hiany & vbLf & U("  - tblOszlop: hi\u00e1nyzik a 'Term\u00e9ksz\u00e1m' sor")
    meta = Array(M_STATUSZ, M_SOROZAT, M_SORFORRAS, M_KIEMELES, M_JELZES, M_ELSO, M_UTOLSO, M_VALTOZAS)
    For Each p In meta
        If OszlopIndexNorm(lo, CStr(p)) = 0 Then hiany = hiany & vbLf & "  - " & CStr(p)
    Next p
    If Len(hiany) > 0 Then
        MsgBox U("A saj\u00e1t t\u00e1bl\u00e1b\u00f3l hi\u00e1nyoznak oszlopok, a friss\u00edt\u00e9s nem indul el:") & hiany, vbExclamation
        Exit Function
    End If
    MezokBetoltese = True
End Function

'==========================================================================
' Forras beolvasasa: minden lap, ahol van "Termekszam" fejlec
'==========================================================================
Private Sub ForrasBeolvasasa(wbS As Workbook)
    Dim ws As Worksheet, adat As Variant
    Dim hdr As Long, kc As Long, r As Long, c As Long, i As Long, k As Long, p As Variant
    Dim utSor As Long, utOsz As Long, kif As Boolean, felso As String
    Dim colMap() As Long, hl As Collection, hlHasznalt As String
    Dim n As String, h As String, legjobb As Long, pont As Long, s As Long
    Dim sor As String, a As Variant, kv As Variant, v As Variant, extra As String
    Dim idx As Long, tarol As Boolean

    nS = 0
    ReDim sKey(1 To 256): ReDim sRaw(1 To 256): ReDim sAkt(1 To 256): ReDim sKif(1 To 256)
    ReDim sSor(1 To 256): ReDim sVal(1 To nF, 1 To 256)
    ReDim sFill(1 To nF, 1 To 256): ReDim sFont(1 To nF, 1 To 256)
    Set dictS = CreateObject("Scripting.Dictionary")
    Set figy = New Collection
    forrasFejlec = ""

    For Each ws In wbS.Worksheets
        '--- horgony: Termekszam fejlec ------------------------------------
        hdr = 0: kc = 0
        For r = 1 To MAX_KERES
            For c = 1 To MAX_KERES
                n = Norm(SafeStr(ws.Cells(r, c).Value))
                If Len(n) > 0 Then
                    For Each p In fPat(keyF)
                        If n = CStr(p) Then hdr = r: kc = c: Exit For
                    Next p
                End If
                If hdr > 0 Then Exit For
            Next c
            If hdr > 0 Then Exit For
        Next r
        If hdr = 0 Then GoTo KovLap

        If Len(forrasFejlec) = 0 Then forrasFejlec = SafeStr(ws.Cells(1, 1).Value)
        felso = ""
        For r = 1 To 3
            For c = 1 To 2
                felso = felso & " " & Norm(SafeStr(ws.Cells(r, c).Value))
            Next c
        Next r
        kif = (InStr(Norm(ws.Name), "KIFUT") > 0) Or (InStr(felso, "ARCHIV") > 0)

        With ws.UsedRange
            utSor = .Row + .Rows.Count - 1
            utOsz = .Column + .Columns.Count - 1
        End With
        If utOsz < kc Then utOsz = kc
        adat = ws.Range(ws.Cells(1, 1), ws.Cells(utSor, utOsz)).Value

        '--- oszlopok parositasa --------------------------------------------
        ReDim colMap(1 To nF)
        Set hl = New Collection
        For c = kc To utOsz
            h = Norm(SafeStr(adat(hdr, c)))
            If Len(h) = 0 Then
                hl.Add c
            Else
                legjobb = 0: pont = 0
                For i = 1 To nF
                    For Each p In fPat(i)
                        s = 0
                        If h = CStr(p) Then
                            s = 1000 + Len(CStr(p))
                        ElseIf Left$(h, Len(CStr(p))) = CStr(p) Then
                            s = Len(CStr(p))
                        End If
                        If s > pont Then pont = s: legjobb = i
                    Next p
                Next i
                If legjobb = 0 Then
                    figy.Add ws.Name & U(": ismeretlen oszlop '") & EgySor(SafeStr(adat(hdr, c))) & "' (" & OszlopBetu(c) & ")"
                ElseIf colMap(legjobb) > 0 Then
                    figy.Add ws.Name & U(": k\u00e9tszer szerepl\u0151 mez\u0151 '") & fCel(legjobb) & "' (" & OszlopBetu(c) & ")"
                Else
                    colMap(legjobb) = c
                End If
            End If
        Next c

        '--- sorok ----------------------------------------------------------
        sor = "": hlHasznalt = ""
        For r = hdr + 1 To utSor
            kv = adat(r, kc)
            n = Norm(SafeStr(kv))
            If ErvenyesKulcs(n) Then
                ' ertekek
                Dim ertek() As Variant
                ReDim ertek(1 To nF)
                For i = 1 To nF
                    If i <> keyF And i <> hlF And colMap(i) > 0 Then
                        v = adat(r, colMap(i))
                        If IsError(v) Then v = "#HIBA"
                        ertek(i) = v
                    End If
                Next i
                If hlF > 0 Then
                    extra = ""
                    For Each p In hl
                        v = adat(r, CLng(p))
                        If Not IsError(v) Then
                            If Len(SafeStr(v)) > 0 Then
                                extra = extra & IIf(Len(extra) > 0, " | ", "") & SafeStr(v)
                                If InStr(hlHasznalt, OszlopBetu(CLng(p)) & ",") = 0 Then hlHasznalt = hlHasznalt & OszlopBetu(CLng(p)) & ","
                            End If
                        End If
                    Next p
                    If Len(extra) > 0 Then ertek(hlF) = extra
                End If

                ' rekord
                tarol = True
                If dictS.Exists(n) Then
                    idx = dictS(n)
                    If (sKif(idx) And kif) Or (sAkt(idx) And Not kif) Then
                        figy.Add ws.Name & U(": t\u00f6bbsz\u00f6r szerepl\u0151 term\u00e9ksz\u00e1m ") & KulcsFormazas(n) & U(" (az els\u0151 marad)")
                        GoTo KovSor
                    End If
                    tarol = (Not kif) Or (Not sAkt(idx))
                Else
                    nS = nS + 1
                    If nS > UBound(sKey) Then Novel
                    idx = nS
                    dictS.Add n, idx
                    sKey(idx) = n
                    For i = 1 To nF
                        sFill(i, idx) = NINCS_SZIN: sFont(i, idx) = NINCS_SZIN
                    Next i
                End If
                If tarol Then
                    sRaw(idx) = kv
                    For i = 1 To nF
                        sVal(i, idx) = ertek(i)
                        sFill(i, idx) = NINCS_SZIN: sFont(i, idx) = NINCS_SZIN
                        If colMap(i) > 0 Then
                            sFill(i, idx) = KitoltesSzin(ws.Cells(r, colMap(i)))
                            sFont(i, idx) = BetuSzin(ws.Cells(r, colMap(i)))
                        End If
                    Next i
                    If Len(sor) > 0 Then sSor(idx) = sor
                End If
                If kif Then sKif(idx) = True Else sAkt(idx) = True

            ElseIf Left$(n, 10) = "KESZITETTE" Then
                Exit For
            ElseIf Len(n) = 0 Then
                ' csoportcim (pl. "A-osztaly (W177)") az A oszlopban
                a = adat(r, 1)
                If VarType(a) = vbString Then
                    If Len(Trim$(a)) > 0 And VanBetu(CStr(a)) Then sor = Trim$(CStr(a))
                End If
            Else
                figy.Add ws.Name & "!" & OszlopBetu(kc) & r & U(": \u00e9rv\u00e9nytelen term\u00e9ksz\u00e1m '") & EgySor(SafeStr(kv)) & "'"
            End If
KovSor:
        Next r
        If Len(hlHasznalt) > 0 Then
            figy.Add ws.Name & U(": fejl\u00e9c n\u00e9lk\u00fcli oszlop(ok) (") & Left$(hlHasznalt, Len(hlHasznalt) - 1) & _
                     U(") adatai a 'Fejl\u00e9c n\u00e9lk\u00fcli adat' oszlopba ker\u00fcltek")
        End If
KovLap:
    Next ws
End Sub

Private Sub Novel()
    Dim m As Long
    m = UBound(sKey) + 256
    ReDim Preserve sKey(1 To m): ReDim Preserve sRaw(1 To m)
    ReDim Preserve sAkt(1 To m): ReDim Preserve sKif(1 To m): ReDim Preserve sSor(1 To m)
    ReDim Preserve sVal(1 To nF, 1 To m)
    ReDim Preserve sFill(1 To nF, 1 To m): ReDim Preserve sFont(1 To nF, 1 To m)
End Sub

'==========================================================================
' Osszevetes a sajat tablaval. alkalmaz=False: csak szamol (elonezet)
'==========================================================================
Private Sub Osszevetes(lo As ListObject, ByVal cimke As String, ByVal alkalmaz As Boolean)
    Dim dictT As Object, i As Long, r As Long, f As Long, n As String
    Dim cKi As Long, szinek As Boolean
    Dim cSt As Long, cSor As Long, cSF As Long, cJel As Long, cEl As Long, cUt As Long, cVal As Long, cKey As Long
    Dim st As String, regiSt As String, regi As Variant, uj As Variant
    Dim valt As Boolean, cellaObj As Range, k As Variant, uj_sor As Boolean
    Dim ujSorozatok As Object

    cUj = 0: cMod = 0: cMezo = 0: cStat = 0: cElt = 0
    cKey = fCol(keyF)
    cSt = OszlopIndexNorm(lo, M_STATUSZ)
    cSor = OszlopIndexNorm(lo, M_SOROZAT)
    cSF = OszlopIndexNorm(lo, M_SORFORRAS)
    cKi = OszlopIndexNorm(lo, M_KIEMELES)
    szinek = SzinekAtvetele()
    cJel = OszlopIndexNorm(lo, M_JELZES)
    cEl = OszlopIndexNorm(lo, M_ELSO)
    cUt = OszlopIndexNorm(lo, M_UTOLSO)
    cVal = OszlopIndexNorm(lo, M_VALTOZAS)
    Set ujSorozatok = CreateObject("Scripting.Dictionary")

    ' sajat tabla kulcsai
    Set dictT = CreateObject("Scripting.Dictionary")
    If Not lo.DataBodyRange Is Nothing Then
        For r = 1 To lo.ListRows.Count
            n = Norm(SafeStr(lo.DataBodyRange.Cells(r, cKey).Value))
            If Len(n) > 0 And Not dictT.Exists(n) Then dictT.Add n, r
        Next r
    End If

    For i = 1 To nS
        n = sKey(i)
        st = StatuszSzoveg(i)
        uj_sor = Not dictT.Exists(n)
        If uj_sor Then
            cUj = cUj + 1
            If alkalmaz Then
                lo.ListRows.Add AlwaysInsert:=True
                r = lo.ListRows.Count
                dictT.Add n, r
                CellaIr lo.DataBodyRange.Cells(r, cKey), KulcsFormazas(n)
                For f = 1 To nF
                    If f <> keyF Then CellaIr lo.DataBodyRange.Cells(r, fCol(f)), sVal(f, i)
                Next f
                CellaIr lo.DataBodyRange.Cells(r, cSt), st
                CellaIr lo.DataBodyRange.Cells(r, cEl), cimke
                CellaIr lo.DataBodyRange.Cells(r, cVal), cimke
                NaploSor KulcsFormazas(n), "", Empty, sVal(NevMezo(), i), U("\u00daj t\u00e9tel"), cimke
            End If
        Else
            r = dictT(n)
            valt = False
            For f = 1 To nF
                If f <> keyF Then
                    Set cellaObj = lo.DataBodyRange.Cells(r, fCol(f))
                    regi = cellaObj.Value: uj = sVal(f, i)
                    If Not Azonos(regi, uj) Then
                        cMezo = cMezo + 1: valt = True
                        If alkalmaz Then
                            NaploSor KulcsFormazas(n), fCel(f), regi, uj, U("M\u00f3dos\u00edt\u00e1s"), cimke
                            CellaIr cellaObj, uj
                        End If
                    End If
                End If
            Next f
            regiSt = SafeStr(lo.DataBodyRange.Cells(r, cSt).Value)
            If regiSt <> st Then
                cStat = cStat + 1: valt = True
                If alkalmaz Then
                    NaploSor KulcsFormazas(n), U("St\u00e1tusz"), regiSt, st, U("St\u00e1tusz"), cimke
                    CellaIr lo.DataBodyRange.Cells(r, cSt), st
                End If
            End If
            If valt Then
                cMod = cMod + 1
                If alkalmaz Then CellaIr lo.DataBodyRange.Cells(r, cVal), cimke
            End If
        End If

        If alkalmaz Then
            ' sorozat: csak ha a forras ad erteket
            If Len(sSor(i)) > 0 Then
                ' a forras blokkcime a "Sorozat (forras)" oszlopba
                regi = SafeStr(lo.DataBodyRange.Cells(r, cSF).Value)
                If CStr(regi) <> sSor(i) Then
                    If Len(CStr(regi)) > 0 Then NaploSor KulcsFormazas(n), U("Sorozat (forr\u00e1s)"), regi, sSor(i), U("M\u00f3dos\u00edt\u00e1s"), cimke
                    CellaIr lo.DataBodyRange.Cells(r, cSF), sSor(i)
                End If
                ' a sajat Sorozat oszlopot csak akkor toltjuk, ha ures
                Dim sajatNev As String, ismert As Boolean
                sajatNev = SorozatNev(sSor(i), ismert)
                If Not ismert Then
                    If Not ujSorozatok.Exists(sSor(i)) Then ujSorozatok.Add sSor(i), True
                End If
                If Len(SafeStr(lo.DataBodyRange.Cells(r, cSor).Value)) = 0 And Len(sajatNev) > 0 Then
                    NaploSor KulcsFormazas(n), "Sorozat", Empty, sajatNev, U("Kit\u00f6lt\u00e9s (\u00fcres volt)"), cimke
                    CellaIr lo.DataBodyRange.Cells(r, cSor), sajatNev
                End If
            End If
            ' a forras sarga/piros jelolesei szovegkent; a szinek csak ha be van kapcsolva
            CellaIr lo.DataBodyRange.Cells(r, cKi), Kiemeles(i)
            If szinek Then
                For f = 1 To nF
                    SzinBeallitas lo.DataBodyRange.Cells(r, fCol(f)), sFill(f, i), sFont(f, i)
                Next f
            End If
            CellaIr lo.DataBodyRange.Cells(r, cJel), Jelzes(i, st)
            CellaIr lo.DataBodyRange.Cells(r, cUt), cimke
        End If
    Next i

    ' a forrasbol eltunt tetelek: nem torlunk, csak jelolunk
    For Each k In dictT.Keys
        If Not dictS.Exists(k) Then
            r = dictT(k)
            regiSt = SafeStr(lo.DataBodyRange.Cells(r, cSt).Value)
            If regiSt <> StElt() Then
                cElt = cElt + 1
                If alkalmaz Then
                    NaploSor KulcsFormazas(CStr(k)), U("St\u00e1tusz"), regiSt, StElt(), StElt(), cimke
                    CellaIr lo.DataBodyRange.Cells(r, cSt), StElt()
                    CellaIr lo.DataBodyRange.Cells(r, cVal), cimke
                End If
            End If
        End If
    Next k

    ' uj forras blokkok felvetele a tblSorozat tablaba (sorozatnev es tipus nelkul)
    If alkalmaz And ujSorozatok.Count > 0 Then
        Dim tS As ListObject, lr As ListRow, cMegj As Long, cBlokk As Long
        Set tS = Tabla(T_SOROZAT)
        If Not tS Is Nothing Then
            cBlokk = OszlopIndexNorm(tS, "FORRASBLOKK"): If cBlokk = 0 Then cBlokk = 1
            cMegj = OszlopIndexNorm(tS, "MEGJEGYZES")
            For Each k In ujSorozatok.Keys
                Set lr = tS.ListRows.Add
                lr.Range.Cells(1, cBlokk).Value = CStr(k)
                If cMegj > 0 Then lr.Range.Cells(1, cMegj).Value = U("\u00daj blokk a forr\u00e1sban \u2013 sorozatn\u00e9v \u00e9s t\u00edpusk\u00f3d kell")
                figy.Add U("\u00daj blokk a forr\u00e1sban: '") & CStr(k) & U("' \u2013 add meg a sorozatnev\u00e9t \u00e9s t\u00edpusk\u00f3dj\u00e1t a Lek\u00e9pez\u00e9s lapon")
            Next k
        End If
    End If
End Sub

Private Function StatuszSzoveg(ByVal i As Long) As String
    If sAkt(i) And sKif(i) Then
        StatuszSzoveg = U("Akt\u00edv \u00e9s kifutott \u2013 ellen\u0151rizend\u0151")
    ElseIf sAkt(i) Then
        StatuszSzoveg = U("Akt\u00edv")
    Else
        StatuszSzoveg = "Kifutott"
    End If
End Function

Private Function StElt() As String
    StElt = U("Elt\u0171nt a forr\u00e1sb\u00f3l")
End Function

Private Function NevMezo() As Long
    Dim i As Long
    For i = 1 To nF
        If Left$(Norm(fCel(i)), 10) = "TERMEKNEV" Then NevMezo = i: Exit Function
    Next i
    NevMezo = keyF
End Function

' Forras blokk -> sajat sorozatnev (tblSorozat). ismert = szerepel-e a blokk a tablaban.
Private Function SorozatNev(ByVal blokk As String, ByRef ismert As Boolean) As String
    Dim tS As ListObject, r As Long, cB As Long, cS As Long
    ismert = False
    Set tS = Tabla(T_SOROZAT)
    If tS Is Nothing Then ismert = True: Exit Function
    If tS.DataBodyRange Is Nothing Then Exit Function
    cB = OszlopIndexNorm(tS, "FORRASBLOKK"): If cB = 0 Then cB = 1
    cS = OszlopIndexNorm(tS, M_SOROZAT)
    For r = 1 To tS.ListRows.Count
        If SafeStr(tS.DataBodyRange.Cells(r, cB).Value) = blokk Then
            ismert = True
            If cS > 0 Then SorozatNev = Trim$(SafeStr(tS.DataBodyRange.Cells(r, cS).Value))
            Exit Function
        End If
    Next r
End Function

' Info lap: "Forras szineinek atvetele" (nmForrasSzinek) = Igen?
Private Function SzinekAtvetele() As Boolean
    On Error Resume Next
    SzinekAtvetele = (Norm(SafeStr(ThisWorkbook.Names("nmForrasSzinek").RefersToRange.Value)) = "IGEN")
End Function

'--- Forras kiemeles: sarga (kitoltott cellan) es piros jelolesek -------------
Private Function Kiemeles(ByVal i As Long) As String
    Dim f As Long, n As String, sarga As String, piros As String, out As String
    For f = 1 To nF
        n = Norm(fCel(f))
        If f <> keyF And f <> hlF And n <> "ANYAG" And n <> "TOOLING" Then
            If sFill(f, i) = RGB(255, 255, 0) And Len(SafeStr(sVal(f, i))) > 0 Then
                sarga = sarga & IIf(Len(sarga) > 0, ", ", "") & Rovid(fCel(f))
            End If
            If sFill(f, i) = RGB(255, 0, 0) Or sFont(f, i) = RGB(255, 0, 0) Then
                piros = piros & IIf(Len(piros) > 0, ", ", "") & Rovid(fCel(f))
            End If
        End If
    Next f
    If Len(sarga) > 0 Then out = U("s\u00e1rga: ") & sarga
    If Len(piros) > 0 Then out = out & IIf(Len(out) > 0, U(" \u00b7 "), "") & "piros: " & piros
    Kiemeles = out
End Function

'--- Adatjelzes: formatumhibak a forrasban --------------------------------
Private Function Jelzes(ByVal i As Long, ByVal st As String) As String
    Dim out As String, f As Long, v As Variant, raw As String
    raw = SafeStr(sRaw(i))
    If Not (raw Like "[A-Z] ### ### ####") Then out = U("forr\u00e1s term\u00e9ksz\u00e1m: '") & raw & "'"
    For f = 1 To nF
        v = sVal(f, i)
        If VarType(v) = vbString And f <> keyF Then
            If fNum(f) And Len(Trim$(v)) > 0 Then
                out = out & IIf(Len(out) > 0, "; ", "") & Rovid(fCel(f)) & U(": sz\u00f6veg '") & v & "'"
            End If
            If CStr(v) <> Trim$(CStr(v)) Then
                out = out & IIf(Len(out) > 0, "; ", "") & Rovid(fCel(f)) & U(": sz\u00f3k\u00f6z a sz\u00e9l\u00e9n")
            End If
        End If
    Next f
    If sAkt(i) And sKif(i) Then out = out & IIf(Len(out) > 0, "; ", "") & U("akt\u00edv \u00e9s kifutott lapon is")
    Jelzes = out
End Function

Private Function Rovid(ByVal s As String) As String
    Dim p As Long
    p = InStr(s, " (")
    If p > 0 Then Rovid = Left$(s, p - 1) Else Rovid = s
End Function

'==========================================================================
' Kiiras: naplo, figyelmeztetesek, info, allapotmentes
'==========================================================================
Private Sub NaploSor(ByVal kulcs As String, ByVal mezo As String, ByVal regi As Variant, ByVal uj As Variant, _
                     ByVal esemeny As String, ByVal cimke As String)
    naplo.Add Array(Now, cimke, kulcs, mezo, regi, uj, esemeny)
End Sub

Private Sub NaploKiirasa(ByVal cimke As String)
    Dim lo As ListObject, db As Long, i As Long, j As Long, t() As Variant, e As Variant
    Set lo = Tabla(T_NAPLO)
    If lo Is Nothing Or naplo.Count = 0 Then Exit Sub
    db = naplo.Count
    ReDim t(1 To db, 1 To 7)
    For i = 1 To db
        e = naplo(i)
        For j = 0 To 6
            If IsEmpty(e(j)) Then t(i, j + 1) = "" Else t(i, j + 1) = e(j)
        Next j
    Next i
    ' uj bejegyzesek a tabla vegere (idorendben)
    Dim alap As Long
    If lo.DataBodyRange Is Nothing Then
        lo.ListRows.Add
        alap = 0
    Else
        alap = lo.ListRows.Count
    End If
    lo.Resize lo.HeaderRowRange.Resize(1 + alap + db)
    With lo.DataBodyRange.Rows(alap + 1).Resize(db)
        .Columns(1).NumberFormat = "yyyy.mm.dd hh:mm"
        .Columns(2).Resize(, 6).NumberFormat = "@"
        .Value = t
    End With
End Sub

Private Sub FigyelmeztetesekKiirasa(ByVal alkalmazva As Boolean)
    Dim lo As ListObject, i As Long, lr As ListRow
    Set lo = Tabla(T_FIGY)
    If lo Is Nothing Then Exit Sub
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
    If Not alkalmazva Then
        Set lr = lo.ListRows.Add
        lr.Range.Cells(1, 1).Value = 0
        lr.Range.Cells(1, 2).Value = U("El\u0151n\u00e9zet ") & Format$(Now, "yyyy.mm.dd hh:mm") & _
                                     U(" \u2013 a lista az utolj\u00e1ra beolvasott f\u00e1jlra vonatkozik")
    End If
    For i = 1 To figy.Count
        Set lr = lo.ListRows.Add
        lr.Range.Cells(1, 1).Value = i
        lr.Range.Cells(1, 2).Value = figy(i)
    Next i
End Sub

Private Sub InfoKiirasa(ByVal cimke As String, ByVal utvonal As String)
    On Error Resume Next
    ThisWorkbook.Names("nmForrasFajl").RefersToRange.Value = cimke
    ThisWorkbook.Names("nmForrasUtvonal").RefersToRange.Value = utvonal
    ThisWorkbook.Names("nmForrasFejlec").RefersToRange.Value = forrasFejlec
    ThisWorkbook.Names("nmImportIdo").RefersToRange.Value = Now
    ThisWorkbook.Names("nmTetelSzam").RefersToRange.Value = nS
    On Error GoTo 0
End Sub

Private Sub AllapotMentese(lo As ListObject)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_MENTES)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_MENTES
    End If
    ws.Visible = xlSheetHidden
    ws.Cells.Clear
    ws.Range("A1").Value = U("A tblTermek \u00e1llapota a ") & Format$(Now, "yyyy.mm.dd hh:mm") & U(" friss\u00edt\u00e9s el\u0151tt")
    ws.Range("A2").Resize(lo.Range.Rows.Count, lo.Range.Columns.Count).Value = lo.Range.Value
End Sub

' Uj soroknal a kepletes oszlopokat biztonsagbol potoljuk (Excel altalaban maga kiterjeszti)
Private Sub KepletekPotlasa(lo As ListObject)
    Dim j As Long, r As Long, elso As Range
    If lo.DataBodyRange Is Nothing Then Exit Sub
    For j = 1 To lo.ListColumns.Count
        Set elso = lo.DataBodyRange.Cells(1, j)
        If elso.HasFormula Then
            For r = 2 To lo.ListRows.Count
                With lo.DataBodyRange.Cells(r, j)
                    If Not .HasFormula And IsEmpty(.Value) Then .Formula = elso.Formula
                End With
            Next r
        End If
    Next j
End Sub

'==========================================================================
' Gomb a Termekadatok lapra
'==========================================================================
Public Sub FrissitoGombLetrehozasa()
    Dim lo As ListObject, ws As Worksheet, b As Object, hely As Range
    Set lo = Tabla(T_TERMEK)
    If lo Is Nothing Then MsgBox U("Nem tal\u00e1lhat\u00f3 a tblTermek t\u00e1bla."), vbExclamation: Exit Sub
    Set ws = lo.Parent
    On Error Resume Next
    ws.Shapes(BTN_NEV).Delete
    On Error GoTo 0
    Set hely = ws.Range("E1")
    Set b = ws.Buttons.Add(hely.Left, hely.Top + 3, 170, ws.Rows(1).Height - 6)
    b.Name = BTN_NEV
    b.Caption = U("Friss\u00edt\u00e9s forr\u00e1sb\u00f3l\u2026")
    b.OnAction = "TermekadatokFrissitese"
    b.Placement = xlFreeFloating
End Sub

'==========================================================================
' Segedfuggvenyek
'==========================================================================
Private Function Tabla(ByVal nev As String) As ListObject
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        Set Tabla = ws.ListObjects(nev)
        If Not Tabla Is Nothing Then Exit Function
    Next ws
End Function

Private Function OszlopIndexPontos(lo As ListObject, ByVal nev As String) As Long
    Dim j As Long
    For j = 1 To lo.ListColumns.Count
        If StrComp(Trim$(lo.ListColumns(j).Name), nev, vbBinaryCompare) = 0 Then OszlopIndexPontos = j: Exit Function
    Next j
End Function

Private Function OszlopIndexNorm(lo As ListObject, ByVal normNev As String) As Long
    Dim j As Long
    For j = 1 To lo.ListColumns.Count
        If Norm(lo.ListColumns(j).Name) = normNev Then OszlopIndexNorm = j: Exit Function
    Next j
End Function

Private Sub CellaIr(cel As Range, ByVal v As Variant)
    If IsEmpty(v) Then
        If Not IsEmpty(cel.Value) Then cel.ClearContents
    ElseIf VarType(v) = vbString Then
        If Len(v) = 0 Then
            If Not IsEmpty(cel.Value) Then cel.ClearContents
        Else
            ' betu nelkuli szoveg (pl. "1430+50", "1231/1532") szovegkent maradjon
            If Not VanBetu(CStr(v)) Or Left$(CStr(v), 1) = "=" Then
                cel.NumberFormat = "@"
            ElseIf cel.NumberFormat = "@" Then
                cel.NumberFormat = "General"
            End If
            cel.Value = v
        End If
    Else
        If cel.NumberFormat = "@" Then cel.NumberFormat = "General"
        cel.Value = v
    End If
End Sub

Private Sub SzinBeallitas(cel As Range, ByVal kitoltes As Long, ByVal betu As Long)
    If kitoltes = NINCS_SZIN Then
        If cel.Interior.Pattern <> xlNone Then cel.Interior.Pattern = xlNone
    Else
        If cel.Interior.Color <> kitoltes Or cel.Interior.Pattern = xlNone Then cel.Interior.Color = kitoltes
    End If
    If betu = NINCS_SZIN Then
        cel.Font.ColorIndex = xlAutomatic
    Else
        cel.Font.Color = betu
    End If
End Sub

Private Function KitoltesSzin(cel As Range) As Long
    Dim c As Variant
    KitoltesSzin = NINCS_SZIN
    If cel.Interior.Pattern = xlNone Then Exit Function
    c = cel.Interior.Color
    If IsNull(c) Then Exit Function
    If CLng(c) = RGB(255, 255, 255) Then Exit Function
    KitoltesSzin = CLng(c)
End Function

Private Function BetuSzin(cel As Range) As Long
    Dim c As Variant
    BetuSzin = NINCS_SZIN
    If IsNull(cel.Font.ColorIndex) Then Exit Function
    If cel.Font.ColorIndex = xlAutomatic Then Exit Function
    c = cel.Font.Color
    If IsNull(c) Then Exit Function
    If CLng(c) = 0 Then Exit Function
    BetuSzin = CLng(c)
End Function

Private Function Azonos(ByVal a As Variant, ByVal b As Variant) As Boolean
    Dim ures1 As Boolean, ures2 As Boolean
    ures1 = IsEmpty(a) Or (VarType(a) = vbString And Len(SafeStr(a)) = 0)
    ures2 = IsEmpty(b) Or (VarType(b) = vbString And Len(SafeStr(b)) = 0)
    If ures1 And ures2 Then Azonos = True: Exit Function
    If ures1 Or ures2 Then Exit Function
    If SzamTipus(a) And SzamTipus(b) Then
        Azonos = (Abs(CDbl(a) - CDbl(b)) < 0.000000001)
    Else
        ' kis/nagybetu es szelso szokoz elteres nem valtozas (a kezi javitas megmarad)
        Azonos = (LCase$(Trim$(SafeStr(a))) = LCase$(Trim$(SafeStr(b))))
    End If
End Function

Private Function SzamTipus(ByVal v As Variant) As Boolean
    Select Case VarType(v)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
            SzamTipus = True
    End Select
End Function

Private Function SafeStr(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    SafeStr = CStr(v)
End Function

Private Function EgySor(ByVal s As String) As String
    EgySor = Replace(Replace(s, vbCr, " "), vbLf, " ")
End Function

Private Function ErvenyesKulcs(ByVal n As String) As Boolean
    ErvenyesKulcs = (Len(n) = 11) And (n Like "[A-Z]##########")
End Function

' R2446370100 -> "R244 637 0100"
Private Function KulcsFormazas(ByVal n As String) As String
    KulcsFormazas = Left$(n, 4) & " " & Mid$(n, 5, 3) & " " & Mid$(n, 8, 4)
End Function

Private Function OszlopBetu(ByVal c As Long) As String
    OszlopBetu = Split(Cells(1, c).Address(True, False), "$")(0)
End Function

Private Function VanBetu(ByVal s As String) As Boolean
    VanBetu = (Norm(s) Like "*[A-Z]*")
End Function

' Normalizalas: ekezet nelkul, nagybetu, csak A-Z es 0-9
Public Function Norm(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    s = UCase$(Ekezetlen(s))
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then out = out & ch
    Next i
    Norm = out
End Function

Public Function Ekezetlen(ByVal s As String) As String
    Dim ekezetes As Variant, sima As Variant, i As Long
    ' a a e i o o o u u u (kis) + ugyanez nagyban, plusz ss
    ekezetes = Array(225, 228, 233, 237, 243, 246, 337, 250, 252, 369, _
                     193, 196, 201, 205, 211, 214, 336, 218, 220, 368)
    sima = Array("a", "a", "e", "i", "o", "o", "o", "u", "u", "u", _
                 "A", "A", "E", "I", "O", "O", "O", "U", "U", "U")
    For i = LBound(ekezetes) To UBound(ekezetes)
        s = Replace(s, ChrW(ekezetes(i)), sima(i))
    Next i
    s = Replace(s, ChrW(223), "ss")
    Ekezetlen = s
End Function

' \uXXXX escape-ek feloldasa, hogy a .bas fajl ASCII maradhasson
Private Function U(ByVal s As String) As String
    Dim p As Long
    p = InStr(1, s, "\u")
    Do While p > 0
        s = Left$(s, p - 1) & ChrW(CLng(Val("&H" & Mid$(s, p + 2, 4) & "&"))) & Mid$(s, p + 6)
        p = InStr(p + 1, s, "\u")
    Loop
    U = s
End Function
