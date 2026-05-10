# Autoblatt — Handoff (otthoni folytatáshoz)

**Aktuális commit:** lásd `git log --oneline | head -5`
**Repo:** https://github.com/Csharlie/autoblatt
**Munkamappa:** `OneDrive\Autoblatt\app\`

---

## Hol tartunk

A teljes újraírás befejeződött az `autosheets` projektből, és minden modul felkerült a GitHub-ra. A telepítő működik, az `autoblatt.xlsb` sikeresen létrejön a `Documents\autoblatt\` mappában.

**Aktuális tesztelési pont:** `AB_SaveToDrive` futtatása. A legutóbbi hiba (`Compile error: Dim local As String` — VBA fenntartott szó) ki lett javítva (lásd commit `8e5a376`). A loader minden futáskor újraimportálja a `.bas` fájlokat a fájlrendszerből, így a fix automatikusan érvényesül.

---

## Otthoni indulás (3 lépés)

1. **Klónozd a repót a OneDrive-ra:**
   ```powershell
   cd "C:\Users\<USER>\OneDrive - Mercedes-Benz (corpdir.onmicrosoft.com)\Autoblatt\"
   git clone https://github.com/Csharlie/autoblatt.git app
   ```

2. **Tedd be a `data.xlsm`-et:** A repó nem tartalmazza (érzékeny személyzeti adat, `.gitignore` védi). Másold be a saját példányodat:
   ```
   C:\Users\<USER>\OneDrive - Mercedes-Benz (corpdir.onmicrosoft.com)\Autoblatt\app\data\data.xlsm
   ```

3. **Excel Trust Center engedély** (egyszer kell):
   File → Options → Trust Center → Trust Center Settings → Macro Settings → pipa: *„Trust access to the VBA project object model"*

4. **Telepítés:** Futtasd `setup-launcher.bat` → `1` (lokális telepítés). Ez:
   - Létrehozza `Documents\autoblatt\autoblatt.xlsb`-t
   - Importálja a `core/` és `ui/` modulokat
   - Létrehozza `%LOCALAPPDATA%\Autoblatt\system-config.json`-t

5. **Indítás:** `setup-launcher.bat` → `0` *(vagy `launch-autoblatt.bat`)*. Ezután Alt+F8 → `AB_*` makrók.

---

## Folytatandó tesztek

| Makró | Tesztelés |
|-------|-----------|
| `AB_SaveToDrive` | Várja a működés ellenőrzését — `local`→`localPath` fix után |
| `AB_ImportImages` | Felbontásfüggetlen pozicionálás teszt; ellenőrizd hogy a 2 oszlopos rács jó pozícióban van-e |
| `AB_SendEmails` | Outlook signature kiderítés (jelenleg `outlook` mode = Outlook saját aláírás) |
| `AB_FillPanel` | Új különálló makró — korábban a SaveToDrive-val futott, most külön |
| `AB_HelprePanel` | Ugyanaz mint fent |
| `AB_InstallButtons` | Új: létrehoz egy „AB Indító" lapot kattintható gombokkal |
| `AB_Settings` | InputBox-alapú konfig szerkesztő |

---

## Ami nyitva van

1. **Email aláírás** — jelenleg `EMAIL_SIGNATURE_MODE = "outlook"` (Outlook saját aláírást vár). Ha sajátot kell, állítsd `"custom"`-ra a configban, és írd át a `BuildCustomSignature()`-t a `modules/email-sender.bas`-ban.

2. **`AB Indító` lap** — egyszer futtasd `AB_InstallButtons`-ot egy munkafüzetben, és lesz egy lap kattintható gombokkal. Ribbon helyettesítő (céges policy blokkolja a custom ribbont).

3. **Settings UI** — most InputBox menü. Ha tényleg natív UserForm kell, manuálisan kell létrehozni a VBA editor-ban (a `.frm`/`.frx` páros nem importálható tisztán szöveges szerkesztőből).

4. **`dist/autoblatt.xlsb` build** — ha kiosztásra kész verziót akarsz, futtasd:
   ```powershell
   PowerShell -ExecutionPolicy Bypass -File installer\build-xlsb.ps1
   ```

5. **Image-importer felülvizsgálat** — a `TestPosition` és `CleanAllImportImages` viszonyát végigtesztelni: különböző felbontáson, hogy a `startCell.Top/Left`-alapú pozicionálás stabilan dolgozik-e.

---

## Részletes dokumentáció

- [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) — modulok, adatfolyam, két verzió architektúra
- [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md) — fejlesztői workflow, encoding, új modul hozzáadása
- [docs/MACROS.md](docs/MACROS.md) — minden `AB_*` makró részletes leírása
- [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md) — eddig megismert hibák és javítások
- [README.md](README.md) — magas szintű projekt áttekintés

---

## Git workflow

```bash
# Változtatás után:
git add -A
git commit -m "rövid leírás"
git push

# Új modul / fix után az xlsb-ben automatikusan a friss kód fut
# (a loader minden AB_* hívásnál újraimportál)
```

Az `installer/utf8-1250.ps1` futtatása csak akkor kell, ha natív magyar ékezetes karaktert írsz egy `.bas` fájlba (a VBA cp1250-et vár).
