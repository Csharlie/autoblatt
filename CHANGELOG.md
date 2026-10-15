# Changelog

## Fotoblatt

| Verzió | Dátum | Változás |
|---|---|---|
| 9.21 | 2026.10.15. | „Mentés Drive-ra” gomb (fejlesztés alatt, csak a fejlesztőnek látható): Mentés másként ide: `OneDrive\Autoblatt\<Év>\<Hónap>\<esetmappa>\<dátum> <típus> <megnevezés> <préselés> <hibakép> (<id>-foto).xlsm`, ékezet nélkül; az id a Data!User táblából, kisbetűvel; ütközéskor v2, v3… Az e-mail tárgya: Cikkszám \| Típus \| Megnevezés \| Préselés \| Hibakép \| Gyártó berendezés \| Fotoblatt (ékezetekkel). A fejlesztői gombok láthatósága egy helyen; az e-mail-másolatból mindkét gomb kimarad. |
| 9.20 | 2026.10.15. | Új mappaszerkezet: a bemenő fájlok a munkafüzet melletti `data\` mappában. A három fájlválasztó (LBL, cikkszámtörzs, hibaképek) a `data\` mappában nyílik, ha a megjegyzett hely már nem létezik (OneDrive-on az URL-útvonalat a helyi mappára képezi le). Az élő munkafüzet az `autoblatt.xlsm`. |
| 9.19 | 2026.10.15. | E-mail gomb a Fotoblatt lapon (fejlesztés alatt, csak a fejlesztőnek látható): új Outlook-levél, tárgy: Cikkszám \| Megnevezés \| Típus \| Préselés \| Hibakép \| Gyártó berendezés, csatolmány: a munkafüzet másolata. `HibakepCella`. |
| 9.18 | 2026.10.14. | Hibafeltárás válaszok: alapérték és üres = szürke kitöltés mindhárom kérdésblokkban (ServoTPS, LBL, Utómunka); kulcsszavas igen/nem minták, új igen/nem legördülők. |
| 9.17 | 2026.10.08. | Rendrakás: ismeretlen/törölt cikkszámnál C6/G6/H6 is kiürül; E4/E6 egy helyen; kevesebb A6-lista-újraépítés; `FOTOBLATT_VERZIO`, lábléc-verzió; horgony: „Hibaleírás / Fehlerbeschreibung”; hibakép-legördülő (Data!tblHibakep, `HibakepekFrissitese`); `ForrasMegnyitasa` (azonos nevű nyitott fájl kezelése). |
| 9.16 | 2026.10.08. | Új forrás az A/N-számokhoz: Autoblatt_Cikkszamtorzs.xlsx (`CikkszamtorzsFrissitese`); G6/H6 mindig a tétel szerint, a 9.12-es „A-szám = 7.41” szabály kivezetve; `H6_URES_HA_NINCS_BERENDEZES`. |
| 9.15 | 2026.10.08. | Egységes cikkszám-formátum: betű + 3-3-4 (`A177 637 0100`), `CikkszamFormatum`, `CikkszamErvenyes`; a Lista a tárolt formát mutatja. |
| 9.14 | 2026.09.29. | Lapkód: az E4/E6 közvetlenül az A6 alapján vált, minden más lépés előtt (üres A6-nál is visszaáll). A törlés nem ír be dátumot és felhasználót. Csak ASCII, a hibás ékezetek javítva. |
| 9.13 | 2026.09.24. | R-számnál E4 = „Alapanyag/Material”, E6 = „Platine”; egyébként E4 = „Préselés/Abpressung”. |
| 9.12 | 2026.09.24. | A-számnál H6 mindig 7.41. |
| 9.11 | – | Az A6 ismét egyetlen, tiszta lista; kereső-előtagok és felugró ablak véglegesen elvetve. Új: `CikkszamBeallitasa`. |
| 9.7–9.10 | – | Kereső-segédsorok, böngésző/kereső mód, `frmA6Kereso` – elvetett kísérletek. |
| 9.6 | – | A Lista oszlopban a cikkszám szóköz nélkül áll → szóközfüggetlen gépelés közbeni szűrés. |
| 9.5 | – | Az A6 lista szűrés nélküli (H6/G6 nem szűr); a választás beállítja a H6-ot. |
| 9.3 | – | Forrás: a saját lbl-product-data tblTermek táblája (nem a nyers TERM_KADATOK). |
| 9.2 | 2026.09.22. | Parts!Article új oszlopai, A6 „Cikkszám · Megnevezés · Típus” lista, `TermekadatokFrissitese`, `ArticleTisztitas`, V530/H248 típuskódok. |
| 9.1 | – | Kiinduló változat. |

## LBL Product Data

| Verzió | Dátum | Változás |
|---|---|---|
| 1.2 | 2026.09.23. | „Forrás kiemelés” oszlop (a forrás sárga/piros jelölései szövegként); színátvétel csak ha `nmForrasSzinek` = Igen; a gomb helye E1. |
| 1.1 | – | A Sorozat saját oszlop (csak üreset tölt), új „Sorozat (forrás)” oszlop; a kis- és nagybetű, illetve a szélső szóköz eltérése nem változás. |
| 1.0 | – | Első változat: horgonyszöveges beolvasás, előnézet, változásnapló, állapotmentés. |

## Fejlesztői eszközök

| Eszköz | Verzió | Dátum | Változás |
|---|---|---|---|
| code-loader (`modKodFrissites`) | 1.0.2 | 2026.09.29. | Üzenetek és megjegyzések egyszerűsítése („forrás mappa”). |
| | 1.0.1 | 2026.09.29. | A mappa- és fájlvizsgálat `GetAttr`-ral: a `Dir` nem látta a rejtett `AppData` mappát, ezért a mentés „Path/File access error”-ral leállt. |
| | 1.0 | 2026.09.29. | `KodFrissitese`: a munkafüzet moduljait és a Fotoblatt lapkódját lecseréli az `src/` változatra; előtte ellenőriz (ASCII, modulnév) és a régi kódot elmenti. |
