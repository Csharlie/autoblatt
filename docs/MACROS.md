# Makrók — részletes leírás

Az összes makró az `AB_` prefix-szel kezdődik és Alt+F8 menüből futtatható.

---

## Fő munka-makrók

### `AB_SaveToDrive`

**Cél:** Adatok lekérése, formázás, verziózott mentés a OneDrive-ra.

**Folyamat:**
1. Megnyitja a `data.xlsm` fájlt (config `dataXlsmPath` vagy `app\data\data.xlsm`)
2. A `Personal` lap `tablePersonal` táblájában keresi a Windows usernevet
3. Beírja az aktív lap (Fotoblatt) `A2`/`B3`/`F3`/`I3` cellájába: username / név / műszak / dátum
4. Másolja a teljes `tablePersonal`-t a `Data` lapra
5. Formázza a Fotoblatt cellákat (font, méret, szín, igazítás)
6. Globális csere minden lapon: `bpl→BPL`, `sw→SW`, `li→Li` (config `replaceRules`)
7. Aktiválja a Fotoblatt lapot, Page Break Preview, 100% zoom
8. Verziózott mentés:
   ```
   OneDrive\Sheets\<év>\<A6 E6 G8>\<dátum>_<C6>_<E6>_<G8>_(PS-foto|PS-info).xlsm
   ```
   (suffix: `foto` ha B1 = „fotoblatt", egyébként `info`)
9. Verzió-konfliktus esetén: kérdez (overwrite vagy v2/v3/...)
10. Megnyitja a mappát Explorer-ben

**Ami megváltozott az autosheets-hez képest:**
- Fill Panel és Helpre Panel **NEM** futnak le automatikusan (külön makrók)
- Szétbontva 6 al-procedúrára (Fetch / Populate / Format / Normalize / Activate / Save)
- A 9 hardkódolt cella a `config.bas`-ból jön

---

### `AB_ImportImages`

**Cél:** Képek (.jpg, .jpeg, .png) elhelyezése 2 oszlopos rácsban az aktív lapon.

**Folyamat:**
1. Töröl minden shape-t a lapról kivéve a `SHAPE_KEEP` ("Grafik 5") nevűt
2. Felderíti a kép forrásmappát:
   - Helyi munkafüzet → `ActiveWorkbook.Path`
   - SharePoint URL → `OneDrive\Sheets\<év>\<dekódolt mappanév>`
3. Beolvas minden .jpg/.jpeg/.png fájlt
4. Mindegyiket elhelyezi a rácsban:
   - Pozicionálás a `IMAGE_START_CELL` (A17) `Top`/`Left`-jéből számolva — **felbontásfüggetlen!**
   - Tájolás: WIA EXIF Orientation (1-4: eredeti, 5-8: 90°-os forgatás)
   - Méret: 9.5 cm szélesség (fekvő) vagy 4.75 cm magasság (álló)
   - Név: `Import_<filenév>`

**Diagnosztika:** `AB_ImportImagesDebugPosition` — megmutatja az `A17` cella aktuális pixel/cm pozícióját (más felbontáson változó).

**Karbantartás:** `AB_CleanImportImages` — csak az `Import_*` képeket törli, a többi shape (pl. Grafik 5) marad.

---

### `AB_SendEmails`

**Cél:** Outlook draft email létrehozása az aktív munkafüzettel mellékletként.

**Folyamat:**
1. Tárgy építés: `A6 + " " + C6 + " " + E6 + " " + G8` (üres cellákat kihagyja)
2. Outlook COM elérés → új MailItem
3. Aláírás kezelés (config `emailSignatureMode`):
   - `outlook` (alapértelmezett) — Outlook saját aláírást vár (Display() betölti)
   - `custom` — saját aláírás (`BuildCustomSignature()`-ből)
   - `none` — üres body
4. Melléklet:
   - Helyi fájl → `ActiveWorkbook.FullName` (URL dekódolva)
   - SharePoint URL → `%TEMP%`-be SaveCopyAs, onnan csatolva
5. Megjeleníti a draft-ot (.Display)

**Konfigurálás:** A `BuildCustomSignature()` a `modules/email-sender.bas`-ban szerkeszthető, ha `custom` mode-ot használsz.

---

### `AB_FillPanel` *(új különálló makró)*

**Cél:** „Fill Panel" lap létrehozása vagy frissítése egy 6 mezős kitöltőpanellel.

**Tartalom:**
- Bal oldali űrlap (B oszloptól, 2. sortól):
  - Mikor készült? (dropdown: Szériajóváhagyó / Gyártás elején / ...)
  - Ki vette észre? (RK1/RK2/NA)
  - Zárolás történt? (Igen/Nem)
  - Mennyiség (szabad szöveg)
  - Egység (db/LT)
- Jobb oldali összefoglaló blokk (5×6 cella merged):
  - Auto-generált mondat IF formulával az inputok alapján
- Named range: `FillPanel`

**Az autosheets-ben:** automatikusan futott a `SaveToDrive`-val. **Most:** csak ha a felhasználó kifejezetten hívja.

---

### `AB_HelprePanel` *(új különálló makró)*

**Cél:** „Helpre Panel" lap létrehozása — fotólap kitöltési segédlet.

**Tartalom (7 szekció):**
1. Főcím
2. Rövid leírás
3. Kitöltési sorrend táblázat (6 szempont × 3 oszlop: Szempont / Mire figyelj? / Példa)
4. Jó/rossz fogalmazás példák
5. Ajánlott mondatváz
6. Saját vázlat blokk (felhasználói terület, érintetlen marad újrafutáskor)
7. Ellenőrzőlista

**Megjegyzés:** A felhasználói beviteli cellák (B28:D31, B34:D38) nem törlődnek újrafutáskor.

---

## UI-makrók

### `AB_InstallButtons`

**Cél:** Worksheet-embedded gombos UI a custom ribbon helyett.

Létrehoz egy „AB Indító" lapot az aktív munkafüzetben, gombokkal:
- Drive Mentés → `AB_SaveToDrive`
- Kép Import → `AB_ImportImages`
- Email Küldés → `AB_SendEmails`
- Fill Panel → `AB_FillPanel`
- Helpre Panel → `AB_HelprePanel`
- Beállítások → `AB_Settings`
- Súgó → `AB_Help`
- Modulok listázása → `AB_ListAllModules`

**Mikor használd:** egyszer, az új munkafüzet előkészítésekor. Utána a gombok benne maradnak.

---

### `AB_Settings`

**Cél:** Felhasználó-szintű beállítások szerkesztése (`%LOCALAPPDATA%\Autoblatt\system-config.json`).

InputBox-alapú menü:
1. Útvonalak (data.xlsm)
2. Cellanevek (mind a 9)
3. Csere-szabályok (bpl/sw/li szerkeszthető)
4. Email beállítások (signature mode)
5. R — Visszaállítás alapértelmezésre
6. X — Kilépés

**Megjegyzés:** Jövőben natív UserForm-ra cserélhető (a `.frm`/`.frx` páros nem importálható szöveges szerkesztőből, manuálisan kell hozzáadni a VBA editor-ban).

---

## Segéd-makrók

### `AB_SystemTest`

Lefuttatja az összes munka-modult egymás után, és mutat egy összesítést (sikeres / hibás darabszám). Hasznos regressziós tesztre.

---

### `AB_CleanupSystem`

Eltávolítja a fejlesztői módban a dinamikusan importált modulok klónjait (`driveManager`, `imageImporter`, ...). Hasznos ha a VBA editor-ban felhalmozódnak.

---

### `AB_ListAllModules`

Listázza a jelenleg betöltött VBA komponenseket (név + típus). Diagnosztika.

---

### `AB_Help`

Statikus súgó dialog az összes makró rövid leírásával.

---

## Makrók adatbiztonsága

| Makró | Olvas | Ír | Kockázat |
|-------|-------|----|----------|
| `AB_SaveToDrive` | data.xlsm + aktív wb | aktív wb + új fájl OneDrive-on | Globális csere (bpl/sw/li) — minden lap minden cellája |
| `AB_ImportImages` | helyi mappa képek | aktív lap shape-jei | Töröl minden shape-t kivéve `Grafik 5`-öt |
| `AB_SendEmails` | aktív wb | Outlook draft | Csak draft-ot készít, nem küld el |
| `AB_FillPanel` | – | „Fill Panel" lap | Új lap létrehozás/frissítés |
| `AB_HelprePanel` | – | „Helpre Panel" lap | Felhasználói cellákat NEM írja felül |
| `AB_ImportImages` | – | shape-ek | – |
| `AB_InstallButtons` | – | „AB Indító" lap + gombok | Új lap |
| `AB_Settings` | config.json | config.json | Csak a config-ot módosítja |
