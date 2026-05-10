# Fejlesztői workflow

## Előkészület (egyszer, új gépen)

### 1. Klónozás
```powershell
cd "C:\Users\$env:USERNAME\OneDrive - Mercedes-Benz (corpdir.onmicrosoft.com)\Autoblatt"
git clone https://github.com/Csharlie/autoblatt.git app
```

### 2. data.xlsm bemásolás
A repó nem tartalmazza (érzékeny adat, `.gitignore`). Másold be:
```
app\data\data.xlsm
```

### 3. Excel Trust Center engedélyezés
File → Options → Trust Center → Trust Center Settings → Macro Settings:
- ☑ *Trust access to the VBA project object model*

Ez a dinamikus modul-importáláshoz kell.

### 4. Telepítés
```cmd
setup-launcher.bat
```
→ válaszd `1` (lokális telepítés). Ez:
- Létrehozza `Documents\autoblatt\autoblatt.xlsb`-t
- Importálja az összes core modult
- Létrehozza `%LOCALAPPDATA%\Autoblatt\system-config.json`-t (a runtime ezt olvassa)

---

## Mindennapi fejlesztés

### Modul módosítása

1. Szerkeszd a `.bas` fájlt szövegszerkesztőben (VS Code, Notepad++, ...)
2. Mentés UTF-8-ban
3. Excelben: Alt+F8 → futtasd a kapcsolódó `AB_*` makrót

A `module-loader` minden hívásnál újraimportálja a friss `.bas` fájlt — **nem kell újratelepíteni**.

### Encoding (csak ha natív magyar ékezetes karaktert írsz)

A VBA editor cp1250-et (Windows-1250) vár. Ha natív ékezetes karaktert használsz egy `.bas` fájlban (pl. „árvíztűrő"), futtasd:

```powershell
PowerShell -ExecutionPolicy Bypass -File installer\utf8-1250.ps1
```

Ez bejárja a projektet és a `.bas`/`.frm`/`.cls` fájlokat UTF-8 → cp1250 konvertálja.

A jelenlegi kódban natív ékezetek nincsenek (ASCII karakterekkel és `Chr()`-ekkel írtam), tehát ez most nem szükséges.

### Új modul hozzáadása

1. Készíts egy `modules/<nev>.bas` fájlt:
   ```vba
   Attribute VB_Name = "<modulNev>"
   Option Explicit

   Public Sub Run<ModulNev>()
       ' ... logika
   End Sub
   ```

2. Add hozzá a kulcs konstansot a `core/config.bas`-hoz:
   ```vba
   Public Const MODULE_KEY_<NEV>  As String = "<nev>"
   Public Const VBMODULE_<NEV>    As String = "<modulNev>"
   ```

3. Regisztráld a `core/module-loader.bas` `GetModuleSpec`-jébe:
   ```vba
   Case MODULE_KEY_<NEV>
       spec.relativePath = "modules\<nev>.bas"
       spec.componentName = VBMODULE_<NEV>
       spec.mainFunction = "Run<ModulNev>"
   ```

4. Add hozzá a `GetAvailableModules()`-hoz (hogy `AB_SystemTest` lefuttassa).

5. Készíts egy `AB_<Nev>` makrót a `core/simple-macros.bas`-ba:
   ```vba
   Public Sub AB_<Nev>()
       On Error GoTo Fail
       Call moduleLoader.LoadAndRunModule(MODULE_KEY_<NEV>)
       Exit Sub
   Fail:
       MsgBox "Hiba: " & Err.Description, vbCritical, APP_NAME
   End Sub
   ```

6. Add hozzá az `installer/startup-manager.bas` `ImportCoreModules` listájához, ha **core** szinten van.

7. Frissítsd a `installer/build-xlsb.ps1` `$modules` listáját.

---

## Standalone build (kiosztáshoz)

```powershell
PowerShell -ExecutionPolicy Bypass -File installer\build-xlsb.ps1
```

Ez:
1. Bezár minden Excel-t
2. Létrehoz egy üres workbook-ot
3. Importálja az összes `core/`, `modules/`, `ui/` `.bas` fájlt
4. Lementi `dist\autoblatt.xlsb`-be (xlExcel12 = 50 formátum)

A generált xlsb standalone — más felhasználóknak ezt kell elküldeni. Nincs benne loader, nem függ a forrás-mappa elérhetőségétől.

---

## Git workflow

```bash
# Változtatás után:
git add -A
git status                  # ellenőrzés
git commit -m "leírás"
git push

# data/data.xlsm soha nem fog commit-olódni (gitignore védi)
# settings/system-config.json sem (per-user érték)
```

### Commit üzenet konvenciók

A repó `feat:`, `fix:`, `docs:`, `refactor:`, `chore:` prefixeket használ.

Példák:
- `feat: add AB_ExportPDF macro`
- `fix: handle SharePoint URL in image import`
- `docs: update troubleshooting guide`
- `refactor: extract common dropdown logic to utils`

---

## Debug tippek

### Nézd a logot
```cmd
type "%LOCALAPPDATA%\Autoblatt\autoblatt.log"
type "%LOCALAPPDATA%\Autoblatt\install-%USERNAME%.log"
```

### Excel tesztelés VBA editor-ban
1. Alt+F11 → VBA editor
2. Immediate Window (Ctrl+G):
   ```vba
   ?GetAppRootPath()
   ?GetConfigPath()
   ?GetConfigValue("installPath", "")
   ?ResolveDataXlsmPath()  ' csak driveManager-ben elérhető
   ```

### Modul-klón takarítás
Ha a VBA editor-ban `driveManager1`, `driveManager2` stb. felhalmozódik:
- Alt+F8 → `AB_CleanupSystem`

### Compile errors gyakori okai
- VBA fenntartott szó változónévként: `Local`, `Type`, `New`, `Date`, `Loop`, `Step`, ... — átnevezés (pl. `localPath`)
- `Option Explicit` mellett deklarálatlan változó — `Dim` hozzáadása
- Hiányzó `Call` vagy nem `Public Sub` — hozzáadás

### Ha az xlsb futás közben hibázik
A loader minden hívásnál friss kódot olvas — **nem kell újratelepíteni**, csak javítsd a `.bas` fájlt és futtasd újra. Kivétel: ha `core/` modult módosítasz (config.bas, utils.bas, módule-loader.bas), azokat csak a startup-manager importálja telepítéskor.

→ ezeket telepítéssel frissíteni: `setup-launcher.bat` → `1`.

---

## Munkafüzet beállítás (új Fotoblatt sablon)

Ha új Excel sablont készítesz amin az autoblatt makróknak futniuk kell:

1. A munkafüzetnek legyen egy `Fotoblatt` lapja (vagy más, de a `B1` cellába írd hogy „fotoblatt" / „infolap" — ez dönti el a fájlnév suffix-ét: `PS-foto` / `PS-info`).
2. A makrók ezeket a cellákat olvassák/írják:
   - `A2` — username (auto)
   - `B3` — név (auto, data.xlsm-ből)
   - `F3` — műszak (auto)
   - `I3` — dátum (auto)
   - `A6, C6, E6, G8` — felhasználó tölti, ezekből épül a fájlnév és email tárgy
   - `H6` — kiemelés (formázás)
   - `C15` — törölt (a SaveToDrive törli)

Ha más cellákat akarsz használni, állítsd át a `core/config.bas` `CELL_*` konstansokat.

3. (Opcionális) Indítsd el a worksheet-button UI-t:
   ```
   Alt+F8 → AB_InstallButtons
   ```
   Ez létrehoz egy „AB Indító" lapot kattintható gombokkal.
