# Handoff – autoblatt (Fotoblatt + LBL Product Data)

Állapot: 2026.09.29. · Karbantartó: Sárdy Péter (PESARDY), MBMH Kecskemét – PW
Ez a dokumentum az `autoblatt` repó kiindulópontja. A korábbi autoblatt-bővítmény (az AutoSheets utódja) az `autoblatt-archive` repóba került, l. 5. pont.

---

## 1. Mi ez, és hogyan folynak az adatok

Két Excel-munkafüzet, egy irányban összekötve:

```
Eredeti TERM_KADATOK (a tulajdonos kiadja: V1–V3, jön a V4)   ← SOHA nem módosítjuk
        │  csak olvasás · makró: lbl-product-data-import.bas → TermekadatokFrissitese
        ▼
lbl-product-data.xlsm  (saját, optimalizált termékadatok, tábla: tblTermek)
        │  csak olvasás · makró: fotoblatt-parts-import.bas → TermekadatokFrissitese
        ▼
fotoblatt.xlsm  – Parts!Article → A6 legördülő → C6 név, G6 típus, H6 berendezés, E4/E6
```

Mindkét lépésnél a forrást a makró csak olvasásra nyitja meg, és mentés nélkül zárja be.

### Kötelező szabályok (Péter döntése, nem változtatható egyeztetés nélkül)
1. A termékszám a saját változatban mindig `RXXX XXX XXXX`; ha A betűvel kezdődik, akkor `AXXX XXX XXXX`.
2. A forrásfájl semmilyen esetben nem változik. A saját fájl frissítésekor az eredeti oszlopok tartalma nem módosul.
3. A V4-et (és minden új verziót) Péter választja ki. A kifutott tételek is jönnek, a színkódok is (az utóbbi kapcsolható: Info!`nmForrasSzinek`).
4. Minden `.bas` / `.vba` fájl **csak ASCII** karaktert tartalmaz, CRLF sorvégekkel. Az ékezetes szöveg `ChrW()`-vel vagy `U("\uXXXX")`-szel készül (a VBE a rendszer kódlapjával olvas, nem UTF-8-cal).

---

## 2. Repó szerkezete és fájlnevek

Fájlnév-konvenció: kisbetű, kötőjel (mint a régi autoblatt `core/`/`modules/` fájljainál). A verziószám a fájlnévbe **nem** kerül. A verziót a kódfejléc, a `CHANGELOG.md` és a git tag (`v9.14`, `lbl-v1.2`) tartja nyilván.

```
autoblatt/
├── README.md
├── HANDOFF.md                        ← ez a fájl
├── CHANGELOG.md
├── .gitignore
├── .gitattributes                    ← CRLF a VBA-fájlokra
├── workbooks/                        ← gitignore, nem kerül fel (a OneDrive-on futnak)
│   ├── fotoblatt.xlsm
│   └── lbl-product-data.xlsm
├── src/
│   ├── fotoblatt-form-logic.bas      VB modul: modFotolap
│   ├── fotoblatt-setup.bas           VB modul: modBeallitas
│   ├── fotoblatt-parts-import.bas    VB modul: modTermekadatok (Fotoblatt oldal)
│   ├── fotoblatt-sheet-events.vba    a Fotoblatt lap kódja (beillesztendő, nem importálható)
│   ├── lbl-product-data-import.bas   VB modul: modTermekadatok (LBL oldal)
│   └── code-loader.bas               VB modul: modKodFrissites (mindkét munkafüzetben; fejlesztői eszköz)
├── docs/
│   ├── fotoblatt-install.md
│   ├── lbl-product-data-install.md
│   └── development.md                ← fejlesztés: kódfrissítés src-ből (KodFrissitese)
└── test-data/                        ← gitignore, nem kerül fel
    └── lbl-source-data-v3-sample.xlsx
```

### Régi → új név

| Régi fájlnév | Új fájlnév | Megjegyzés |
|---|---|---|
| Fotoblatt_auto_v9.1.xlsm | `workbooks/fotoblatt.xlsm` | a név ellenére 9.12-es modul + 9.14-es lapkód van benne |
| Fotoblatt_auto_v9.1_backup.xlsm | – | **byte-ra azonos** a fentivel, törölhető |
| Fotoblatt_auto_v9.1_bazis.xlsm, Fotoblatt_auto.xlsm, autoblatt.xlsm, …bazis-jav.xlsm | – | elavult, nem kerül a repóba |
| LBL_Termekadatok_sajat.xlsm | `workbooks/lbl-product-data.xlsm` | |
| modFotolap_v9.2.bas | `src/fotoblatt-form-logic.bas` | |
| modBeallitas_v9.2.bas | `src/fotoblatt-setup.bas` | |
| modTermekadatok_B_v9.2__16_.bas | `src/fotoblatt-parts-import.bas` | 9.13 |
| shFotoblatt_lapkod_v9.14.txt | `src/fotoblatt-sheet-events.vba` | 9.14 |
| modTermekadatok_A_v1.2.bas | `src/lbl-product-data-import.bas` | 1.2 |
| TELEPITES_v9.14.txt | `docs/fotoblatt-install.md` | frissítve az új nevekre |
| TERM_KADATOK_V3_forras_teszthez.xlsx | `test-data/lbl-source-data-v3-sample.xlsx` | nem kerül gitbe |
| struktura.txt, HANDOFF_LBL_Fotoblatt.md | – | helyettük: README.md, HANDOFF.md |
| a többi modTermekadatok_B_v9.2__N_.bas, *-jav.bas | – | elavult köztes verziók, törölhetők |

**Átnevezés után egyszer:** a Fotoblatt `TermekadatokUtvonal` neve még a régi `LBL_Termekadatok_sajat.xlsm`-re mutat. A következő `TermekadatokFrissitese` futásakor a fájlválasztóban egyszer ki kell választani az új `lbl-product-data.xlsm`-et, utána már azt jegyzi meg.

---

## 3. Ellenőrzött állapot (2026.09.29.)

A feltöltött munkafüzetekből kinyert VBA-kódot soronként összevetettem a forrásfájlokkal.

| Elem | Munkafüzetben | `src/`-ben | Egyezik? |
|---|---|---|---|
| Fotoblatt lapkód | 9.14 | 9.14 | ✅ |
| modFotolap | 9.2 | 9.2 | ✅ (csak betűméret-normalizálás, l. 5.6) |
| modBeallitas | 9.2 | 9.2 | ✅ |
| modTermekadatok (Fotoblatt) | **9.12** | **9.13** | ⚠️ eltér |
| frmA6Kereso (üres űrlap) | benne van | – | ⚠️ felesleges |
| modTermekadatok (LBL) | 1.2 | 1.2 | ✅ (csak betűméret-normalizálás) |

**Adatok:**
- `lbl-product-data.xlsm` – tblTermek: 117 tétel, mind helyes formátumú (`R### ### ####`). Státusz: Aktív 77 · Kifutott 26 · Eltűnt a forrásból 10 · Aktív és kifutott – ellenőrizendő 4. Az utolsó import Excelben futott (2026.09.23., a V3 tesztmásolatából), a makró tehát élesben működik.
- `fotoblatt.xlsm` – Parts!Article: 2781 A-szám + 107 R-szám (81 aktív = 77 + 4 „ellenőrizendő”, 26 kifutott; a 10 eltűnt tétel helyesen kimaradt). Mind a 107 R-szám Berendezés = 7.50. Az A-számoknak nincs típusa.
- A két fájl közötti lánc tehát rendben lefutott.

---

## 4. Talált problémák és teendők (prioritás szerint)

### Hibák
1. **G6 beragad.** R-számról A-számra váltva a G6-ban benne marad az előző típus, mert a kód csak akkor ír a G6-ba, ha a tételnek van típusa, az A-számoknak pedig nincs. A mentett példányban is ez látszik: A0009908209 SCHWEISSSCHRAUBE mellett G6 = W177. *Javítás: ha a tételnek nincs típusa, a G6 ürüljön (`CikkszamBeallitasa`).*
2. **A Fotoblatt modulja nem a legfrissebb.** A munkafüzetben 9.12 van, a repóban 9.13. A 9.14-es lapkód miatt a működés most azonos, de a munkafüzet és a repó ne térjen el. *Teendő: `fotoblatt-parts-import.bas` újraimportálása.*
3. **Dupla E4/E6-logika.** A feliratváltás a lapkódban (9.14) és a modulban (9.13) is benne van. *Javaslat: a modulból kivenni, maradjon egy helyen, a lapkódban.*

### Takarítás
4. Az üres `frmA6Kereso` űrlap törlése a Fotoblattból (VBE → jobb klikk → Remove).
5. Az `A6ListaMod` név törlése (Formulas → Name Manager). Semmi nem használja.
6. A Fotoblatt és a Parts lap használt tartománya túl nagy (2862, illetve 11526 sor, üres formázás). Érdemes levágni, kisebb és gyorsabb lesz a fájl.
7. A `fotoblatt.xlsm` jelenleg kitöltött példány (A6, C6, H6 stb.). A repóba üres sablon kerüljön: a kitöltendő mezőket törölni, majd menteni.
8. A helyi mappa (`struktura.txt`) még a régi fájlokat mutatja (bazis.xlsm, v9.2-es txt-k). Ezt az új szerkezetre kell cserélni.
9. A modulok fejlécében még a régi `LBL_Termekadatok_sajat.xlsm` név szerepel. A következő kódmódosításkor frissíteni kell.

### Nyitott döntések (Péteré)
10. **Az A-számok formátuma.** A Parts listán szóköz nélkül állnak (`A0009908209`, az `ArticleTisztitas` is így hozza őket), a saját szabály szerint viszont `AXXX XXX XXXX`. A keresés mindkettővel működik (a `Kulcs()` szóközfüggetlen), de a megjelenítést egységesíteni kell.
11. **„Aktív és kifutott – ellenőrizendő”** (4 tétel): a Fotoblatt aktívként kezeli. Maradjon így, vagy a forrásban kell tisztázni?
12. **Berendezés oszlop** a tblTermek-ben (LBL v1.3): most nincs, ezért minden R-szám 7.50-et kap, és a frissítés 107-szer figyelmeztet. Ha az LBL mindig 7.50, az oszlop elhagyható, és a figyelmeztetést ki lehet kapcsolni.
13. **Parts!Article tisztítás:** 242 ismétlődő cikkszám van, ebből 8 eltérő névvel (1 valódi hiba: A0009908415 KONSOLE… vs KOMBIMUTTER). Az `ArticleTisztitas` még nem futott.
14. **„Üres lap” makró:** egy gombbal kiürítené az összes kitöltendő mezőt. Ehhez rögzíteni kell, mely mezők tartoznak ide (fejléc, Hibafeltárás-válaszok, WZG IH, Intézkedések).
15. **VB-modulnevek angolra** (opcionális): `modFormLogic`, `modSetup`, `modPartsImport`, `modProductImport`. Ez megszüntetné a két `modTermekadatok` névütközését. Modulnévvel minősített hívás nincs a kódban, így az átnevezés biztonságos. Az LBL gomb `OnAction`-je eljárásnévre hivatkozik, nem modulnévre.

---

## 5. Git és GitHub

### 5.1 Adatvédelem – FONTOS
A munkafüzetek céges belső adatot tartalmaznak: cikkszámok, beszállító (pl. Salzgitter), programazonosítók, munkatársak neve és azonosítója (Data!User), OneDrive- és P:-meghajtó-útvonalak. Ezért:
- **csak privát repóba** kerülhet. 2026.09.29-én az `autoblatt` és az `autosheets` is **nyilvános** volt, benne `pesardy`-útvonalakkal, az autosheetsben `install.log`-gal és `system-config.json`-nal. Mindkettőt privátra kell állítani (Settings → Danger Zone → Change visibility). Érdemes egyeztetni, hogy céges kód maradhat-e személyes GitHub-fiókban;
- a `test-data/` és az eredeti TERM_KADATOK fájlok a `.gitignore` miatt nem kerülnek fel;
- a munkafüzetek (`workbooks/*.xlsm`) a `.gitignore` miatt szintén nem kerülnek fel, amíg nincs céges jóváhagyás. A kód (`src/`) önmagában is visszaállítható.

### 5.2 Munkahelyek

*Frissítve 2026.10.15.: új mappaszerkezet (l. README).*

| Hely | Mi van ott | Mire való |
|---|---|---|
| `C:\Users\pesardy\Projects\autoblatt\` | git-repó: `src/`, `docs/`, HANDOFF, CHANGELOG, `test-data/` | fejlesztés, verziókezelés, push. Innen makró nem fut, nem is kell. Munkafüzet nem kerül ide. |
| `OneDrive\Projects\autoblatt\` | csak az `autoblatt.xlsm` (a Fotoblatt; a korábbi `fotoblatt.xlsm` archiválva) | futtatás és Excel-munka (céges szabály: makró csak OneDrive-ról fut) |
| `OneDrive\Projects\autoblatt\data\` | `lbl-product-data.xlsm`, `Autoblatt_Cikkszamtorzs.xlsx`, `hibamegnevezések.xls`, `Fehlerarten_Hibamegnevezések.xls`, `cikkszamok\`, `part-details.xlsm`, `id-sfm-cikkszam-*`, `Pres_Termekadatok.xlsx`, `Hianyzo_A_szamok_ellenorzes.xlsx`, `v9_Utómunka formanyomtatványok a Présüzemben.xlsx` | minden, amit az autoblatt beolvas. A fájlválasztók itt nyílnak, ha a megjegyzett hely már nem létezik (9.20). |
| `OneDrive\Autoblatt\<Év>\<Hónap>\`, `OneDrive\Autoblatt\Status\` | fotólapok év/hónap szerint (pl. `2026\10`), státuszjelentések alkatrészenként | munkatermékek; a céges példány a P: meghajtón. A régi `OneDrive\Sheets`, `OneDrive\AutoSheets` és `OneDrive\Autoblatt\2026` ide rendezve (2026.10.15.). |
| `C:\Users\pesardy\Projects\autoblatt-archive\` | régi generációk és bővítmények (AutoSheets, Autoblatt v1.0, Fotoblatt_auto), elavult munkafüzetek, git-bundle-ök; tartalomjegyzék: `ARCHIV.md` | csak olvasásra. Helyi mappa, nincs automatikus mentése. |

A régi bővítmények (AutoSheets, Autoblatt v1.0) 2026.10.15-én nyugdíjazva: az Excel indításkor nem tölt be régi bővítményt, a kódjuk és telepítésük az archívumban van.

### 5.3 Egyszeri átállás a GitHubon

- **Régi autoblatt → `autoblatt-archive`:** a meglévő `Csharlie/autoblatt` repó átnevezése (Settings → Rename), privátra állítása, `addin-v1-final` tag, majd Archive. Az előzmény teljes egészében megmarad.
  - A régi bővítmény helyi clone-jaiban a remote-ot át kell állítani: `git remote set-url origin https://github.com/Csharlie/autoblatt-archive.git`. Enélkül a GitHub átirányítása addig működik, amíg létre nem jön az új `autoblatt` repó, onnantól viszont az új repóra mutatnának.
  - Ha a telepített régi bővítmény egy `OneDrive\Projects\autoblatt\` mappából tölt, a mappát előbb át kell nevezni (`autoblatt-archive`), és a `%LOCALAPPDATA%\Autoblatt\system-config.json` `installPath`-ját is át kell írni, különben a két projekt ugyanabba a mappába kerül.
- **Új `Csharlie/autoblatt`:** új, **privát** repó, tiszta előzménnyel. Az első commit ez a csomag (`feat: autoblatt – Fotoblatt 9.14 + LBL product data 1.2`), tagek: `v9.14`, `lbl-v1.2`. Mivel a régi előzmény nem jön át, a korábban nyilvános céges útvonalak sem kerülnek bele.
- **autosheets:** a helyi fájlok kivétele a követésből, `autosheets-final` és `autosheets-2.0-draft` (`origin/dev`) tag, a `dev`/`dev2` törlése, a README-be „Superseded by autoblatt”, privát, Archive. (Egységesség kedvéért átnevezhető `autosheets-archive`-ra is.)

A régi kód később is elérhető, például ha egy funkciót (e-mail, képimport, Drive-mentés) át kell venni: `autoblatt-archive` → `modules/…`.

### 5.4 Munkamenet
- A `main` ág mindig működő állapot. Fejlesztés rövid ágon (`feat/g6-fix`, `fix/…`), kész után merge és az ág törlése. Nincsenek hosszan élő `dev` ágak.
- Commitüzenet: `feat:` / `fix:` / `docs:` / `chore:`.
- A kódot a `C:\Users\pesardy\Projects\autoblatt\src\` mappában szerkeszted (pl. VS Code). A OneDrive-on lévő munkafüzetben Alt+F8 → `KodFrissitese`: a makró lecseréli a munkafüzet moduljait és a Fotoblatt lapkódját az `src/` változatra (előtte elmenti a régit), utána Debug → Compile VBAProject, próba, mentés. Részletek: [`docs/development.md`](docs/development.md).
- A VBE-ben ne javíts: a következő `KodFrissitese` felülírja. Ha mégis ott javítottál, a változást kézzel vidd át az `src/` fájlba.
- A fejléc `Verzio` sorának és a `CHANGELOG.md`-nek a frissítése, majd commit és tag (`v9.15`, `lbl-v1.3`).
- **Tilos:** a régi bővítmény `AB_SaveToDrive` makróját az új `fotoblatt.xlsm`-en futtatni. Kiüríti a `Data` lapot (User, Type, OP stb. táblák), és a teljes munkafüzetben szöveget cserél.

### 5.5 Sorvég és kódolás
A `.gitattributes` a `.bas`/`.cls`/`.frm`/`.vba` fájlokat CRLF-fel tartja. Enélkül a git LF-re alakíthatja őket, és a VBE import elrontja a sorokat. Minden fájl maradjon ASCII.

### 5.6 Zajos diffek elkerülése
A VBE az azonos nevű azonosítókat egységes betűméretre írja át. Ezért a kódban nem lehet két név, amely csak kis- és nagybetűben tér el. Jelenleg ilyen:
- `Szoveg` (függvény) és `szoveg` (paraméter) a Fotoblattban;
- `cEl` (változó) és `cel` (paraméter) az LBL-ben.

Exportkor ezek „maguktól” megváltoznak. Javítás: a paraméterek átnevezése (pl. `txt`, `celCella`).

---

## 6. Működés röviden

### Fotoblatt (9.14)
- **A2** felhasználó → B3 név, F3 műszak (Data!User).
- **I3** dátum és **A2** automatikusan csak valódi bevitelre. Törlés (Delete) nem tölt ki semmit, így a lap teljesen üresre törölhető.
- **A6** legördülő: minden tétel szűrés nélkül (aktív R → kifutott R → A), „Cikkszám · Megnevezés · Típus” alakban. Gépelve a cikkszám elejétől szűr, szóköz- és kisbetűfüggetlenül; a legördülő nyilával bárhol kereshető.
- Választáskor: C6 név, G6 típus, H6 berendezés (A-szám → 7.41, R-szám → a tétel Berendezés értéke). A H6 váltása újraépíti a Hibafeltárás-blokkot.
- **E4/E6:** R-szám → „Alapanyag/Material” + „Platine”; A-szám vagy üres A6 → „Préselés/Abpressung”, és a gép által beírt „Platine” törlődik.
- **H6:** 7.41/7.22 → ServoTPS-kérdések, 7.50 → LBL-kérdések, Utómunka → NA-kérdések, más → a blokk rejtve. A képterület egy A4-es oldalra igazodik.

### LBL Product Data (1.2)
- Forrás kiválasztása → előnézet → „Igen” után ír. Előtte mentés a rejtett `Elozo_allapot` lapra, minden változás a Változásnaplóba kerül.
- Kulcs: a normalizált termékszám. Az oszlopokat fejlécszöveg alapján párosítja (`tblOszlop`), a blokkcímekből `Sorozat (forrás)` lesz.
- A saját oszlopokat soha nem írja. A `Sorozat` oszlopot csak akkor tölti ki, ha üres. A kis- és nagybetű, illetve a szélső szóköz eltérése nem számít változásnak.
- A forrásból eltűnt tételt nem törli, csak „Eltűnt a forrásból” jelölést kap.

---

## 7. Következő lépések (javasolt sorrend)
1. Mindkét régi repó privátra állítása, `autoblatt` → `autoblatt-archive`, az új `autoblatt` repó létrehozása, az autosheets archiválása (5.3).
2. Takarítás: 4–8. pont.
3. **v9.15:** G6-javítás, E4/E6 egy helyre, paraméter-átnevezések (5.6), fejléc-hivatkozások az új fájlnevekre; opcionálisan modulnevek angolra (15. pont).
4. Döntések: 10–13. pont, utána `ArticleTisztitas` futtatása.
5. A V4 megjelenésekor: LBL frissítés → Fotoblatt frissítés, a `docs/` leírások szerint.
