# Hibaelhárítás

Eddig megismert hibák és javítások.

---

## Telepítés / induláskor

### „Compile error: Syntax error" valamelyik `.bas` modulban

**Ok:** VBA fenntartott szót használtál változónévként.

**Példák fenntartott szavakra:**
- `Local`, `Type`, `New`, `Date`, `Loop`, `Step`, `Next`, `End`, `Sub`, `Function`, `Public`, `Private`, `Static`, `Const`, `Enum`, `Class`, `Implements`, `Friend`, `Property`, `Get`, `Let`, `Set`, `Object`, `Variant`, `Boolean`, `Integer`, `Long`, `Double`, `Single`, `String`, `Currency`, `Decimal`, `Byte`, `LongLong`, `LongPtr`

**Megoldás:** átnevezés (pl. `local` → `localPath`, `type` → `compType`).

**Eddig javított eset:** `modules/drive-manager.bas:139` `Dim local As String` → `Dim localPath As String` (commit `8e5a376`).

---

### „Modul fájl nem található: Documents\autoblatt\modules\drive-manager.bas"

**Ok:** A loader a `ThisWorkbook.Path`-ot használja, ami a `Documents\autoblatt\` (ahol az xlsb fut), de a `.bas` modulok az `OneDrive\Autoblatt\app\modules\`-ban vannak.

**Megoldás:** A loader (`utils.bas:GetAppRootPath`) a felhasználó-szintű configből (`%LOCALAPPDATA%\Autoblatt\system-config.json`) olvassa az `installPath`-ot, és az alapján keresi a modulokat. Ha a config helyes, megtalálja.

**Ellenőrzés:**
```cmd
type "%LOCALAPPDATA%\Autoblatt\system-config.json"
```

Az `installPath` a OneDrive-os mappára kell mutasson. Ha nem, futtasd újra: `setup-launcher.bat` → `1`.

**Javított commit:** `8e03241`.

---

### „HIBA: Cel mappa letrehozas sikertelen: Bad file name or number"

**Ok:** A JSON parser hibásan kezelte a `\\` (escape-elt backslash) + `"` (string-vég) sorrendet — az érték `Documents\autoblatt\` helyett `Documents\autoblatt\",\n  ` lett.

**Megoldás:** Páros/páratlan backslash számlálás a JSON parserben (`utils.bas:GetJsonString`, `startup-manager.bas:ExtractJsonValue`).

**Javított commit:** `04e95ed`.

---

### „Programmatic access to Visual Basic Project is not trusted"

**Ok:** Excel Trust Center beállítás hiányzik — a dinamikus modul-importáláshoz kell.

**Megoldás:** File → Options → Trust Center → Trust Center Settings → Macro Settings → ☑ *Trust access to the VBA project object model*.

---

### „Excel nem elérhető" a telepítéskor

**Ok:** Az Excel COM nem indul.

**Megoldás:**
1. Bezár minden Excel-t: `setup-launcher.bat` → `3` vagy `taskkill /f /im EXCEL.EXE`
2. Próbáld újra a telepítést

---

## Futáskor

### A `AB_SaveToDrive` „A 'data.xlsm' nem található" hibát ad

**Ellenőrizd:**
1. `%LOCALAPPDATA%\Autoblatt\system-config.json` — `dataXlsmPath` mező
2. Ha üres, fall back: `app\data\data.xlsm`
3. Ha az sincs: `OneDrive\data.xlsm`

**Megoldás:**
- Másold be a `data.xlsm`-et az `app\data\` mappába
- Vagy állítsd be: `AB_Settings` → `1` → adj meg egy abszolút útvonalat

---

### A `AB_SaveToDrive` „A 'Personal' munkalap nem található" hibát ad

**Ok:** A `data.xlsm` nem tartalmaz `Personal` nevű lapot, vagy nincs benne `tablePersonal` ListObject.

**Megoldás:** Ellenőrizd a `data.xlsm` szerkezetét:
- Lapnév: `Personal`
- Excel táblázat (Insert → Table) név: `tablePersonal`
- Oszlopok: `ID`, `Name`, `Shift`

A nevek a `core/config.bas`-ban módosíthatók (`SHEET_PERSONAL`, `TABLE_PERSONAL`, `COL_PERSONAL_*`).

---

### A `AB_ImportImages` „Nem található helyi mappa" hibát ad

**Ok:** A munkafüzet SharePoint URL-en él (https://...), és a SaveCopyAs / Path felderítés nem találja a helyi mappát.

**Megoldás:**
1. A munkafüzetnek mentve kell lennie OneDrive-on, ami szinkronizálva van a helyi gépre
2. A `ResolveLocalImageFolder()` az URL utolsó szegmensét veszi, és ezt keresi a `OneDrive\Sheets\<év>\` alatt
3. Ha máshol vannak a képek, állítsd át a logikát a `image-importer.bas`-ban

---

### A képek rossz pozícióban jelennek meg

**Ok:** Felbontás-specifikus pozicionálási probléma.

**Diagnosztika:** Alt+F8 → `AB_ImportImagesDebugPosition` — megmutatja az `A17` cella aktuális Top/Left értékeit.

**Megoldás:** A `image-importer.bas` konstansait hangold:
- `COL_LEFT_OFFSET_CM` (jelenleg `0.05`)
- `COL_RIGHT_OFFSET_CM` (jelenleg `9.6`)
- `ROW_STEP_CM` (jelenleg `5.38`)

Ezek a startCell-hez képesti eltolások cm-ben.

---

### A `AB_SendEmails` „Outlook nincs telepítve" hibát ad

**Ok:** Outlook COM nem elérhető.

**Megoldás:**
1. Outlook telepített és legalább egyszer elindítva
2. Outlook profil beállítva (legalább egy email fiók)
3. Office bit-megfelelőség (32 vs 64) — Excel és Outlook ugyanaz kell

---

### A globális csere (bpl/sw/li) felülír valódi felhasználói adatot

**Ok:** A `NormalizeWorkbookText` minden lap minden szöveges cellájában fut, a részsztring-ekben is cserél (pl. „Lika" → „Lika" — itt nem, de „Lima" → „Lima" — itt sem, viszont „Likör" lehet probléma).

**Megoldás:** Állítsd át a csere-szabályokat:
- `AB_Settings` → `3` → szerkesztheted a `replaceRules`-t
- Vagy commit-old a `core/config.bas` `REPLACE_RULES` konstansát

A csere case-insensitive és minden lapon fut. Ha túl agresszív, módosítsd a `utils.bas:ApplyReplaceRules` logikáját whole-word matchre.

---

## VBA Editor problémák

### A modulok felhalmozódnak (`driveManager1`, `driveManager2`, ...)

**Ok:** A loader `Remove` művelete néha sikertelen (zárolt komponens, futás közben).

**Megoldás:** Alt+F8 → `AB_CleanupSystem` — eltávolítja a klónokat.

### A módosított `core/utils.bas` nem érvényesül futás közben

**Ok:** A `core/` modulok az xlsb-be be vannak égetve, csak telepítéskor frissülnek (a `installer/startup-manager.bas:ImportCoreModules` listázza őket).

**Megoldás:**
- Vagy: `setup-launcher.bat` → `1` (újratelepítés)
- Vagy: a VBA editor-ban manuálisan re-importálod (Alt+F11 → File → Import File → válaszd ki a `core/utils.bas`-t)

A `modules/` és `ui/` modulok minden hívásnál frissülnek a loaderen keresztül.

---

## Encoding problémák

### „Nem tudom olvasni a magyar karaktereket"

**Ok:** A `.bas` fájl UTF-8-ban van, de a VBA editor cp1250-et vár.

**Megoldás:**
```powershell
PowerShell -ExecutionPolicy Bypass -File installer\utf8-1250.ps1
```

Aztán futtasd újra a telepítést. A jelenlegi kódban natív magyar ékezetek nincsenek (ASCII + `Chr()` escape-ek), így ez most nem szükséges.

---

## Git problémák

### CRLF figyelmeztetések commit-kor

```
warning: in the working copy of 'core/utils.bas', LF will be replaced by CRLF the next time Git touches it
```

**Ok:** A repó vegyes vagy LF/CRLF.

**Megoldás:** Ártalmatlan figyelmeztetés, Windows alatt CRLF a default. Ha zavar, beállíthatod:
```bash
git config --local core.autocrlf true
```

---

## Diagnosztikai parancsok

```cmd
:: Telepítési log
type "%LOCALAPPDATA%\Autoblatt\install-%USERNAME%.log"

:: Runtime log
type "%LOCALAPPDATA%\Autoblatt\autoblatt.log"

:: Aktuális config
type "%LOCALAPPDATA%\Autoblatt\system-config.json"

:: Telepített xlsb
dir "%USERPROFILE%\Documents\autoblatt\"
```

VBA Immediate Window (Ctrl+G) tesztek:
```vba
?GetAppRootPath()             ' modulok forrása
?GetWorkbookPath()            ' xlsb fizikai helye
?GetConfigPath()              ' config útvonala
?GetConfigValue("installPath", "")   ' aktuális installPath
?Environ("USERNAME")          ' Windows usernév
?Environ("LOCALAPPDATA")      ' AppData mappa
?Environ("OneDriveCommercial") ' OneDrive gyökér
```
