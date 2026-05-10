# Architektúra

## Tervezési alapelvek

1. **Két verzió** — Fejlesztői `.xlsm` (loader, git) + kiosztott standalone `.xlsb` (`dist/`, beégetett kód).
2. **Egy hely a hardkódolt értékeknek** — `core/config.bas`. Cellanevek, lapnevek, shape-nevek, csere-szabályok, útvonalak.
3. **Loose coupling a modulok között** — minden munka-modul önálló, csak a `core/utils.bas` közös segédeket használja.
4. **Vékony orchestrátorok** — a `SaveToDrive` 5 al-procedúrából áll, nem 356 soros monolit.

---

## Mappastruktúra

```
autoblatt/app/
├── core/                       # Mindig betöltött modulok (XLSB-ben élnek)
│   ├── config.bas              # Konstansok
│   ├── utils.bas               # Közös függvények
│   ├── module-loader.bas       # Dinamikus modul betöltő (fejlesztői mód)
│   ├── ribbon-callbacks.bas    # Ribbon gomb handlerek
│   ├── ribbon-config.xml       # Custom ribbon (céges policy miatt nem aktív)
│   ├── simple-macros.bas       # AB_ Alt+F8 belépési pontok
│   ├── cleanup-helper.bas      # Modul-klón takarítás
│   └── settings-ui.bas         # InputBox-alapú beállítások
├── modules/                    # Dinamikusan betöltött munka-modulok
│   ├── drive-manager.bas       # Adat lekérés + mentés
│   ├── fill-panel.bas          # Fotólap kitöltőpanel
│   ├── helpre-panel.bas        # Kitöltési segédlet
│   ├── image-importer.bas      # Képek 2 oszlopos rácsban
│   └── email-sender.bas        # Outlook draft email
├── ui/
│   └── buttons-installer.bas   # 'AB Indító' lap (ribbon helyettesítő)
├── data/
│   └── data.xlsm               # Személyzeti adat (gitignore!)
├── dist/
│   └── autoblatt.xlsb          # Generált standalone build
├── installer/
├── settings/
├── docs/
└── update/                     # Egyelőre üres
```

---

## Két verzió architektúra

### Fejlesztői `.xlsm`

- **Helye:** `OneDrive\Autoblatt\app\` (a projekten belül)
- **Indítás:** Excel-ben megnyitva
- **Modul betöltés:** A `core/module-loader.bas` minden `AB_*` futtatáskor importálja a megfelelő `modules/*.bas` fájlt a fájlrendszerről, futás után törli (vagy újraimportálja).
- **Előnyök:** A `.bas` fájlok git-ben verziókezeltek, szövegszerkesztőben szerkeszthetők. A módosítások azonnal érvényesülnek a következő `AB_*` hívásnál.

### Kiosztott `dist/autoblatt.xlsb`

- **Helye:** `Documents\autoblatt\` (telepítés után)
- **Indítás:** `launch-autoblatt.bat` vagy közvetlen kattintás
- **Modul betöltés:** Nincs! A kód be van égetve az xlsb-be a build során.
- **Előnyök:** Standalone, nem függ a forrás-mappa elérhetőségétől. Más felhasználóknak ezt kell kiosztani.
- **Generálás:** `installer\build-xlsb.ps1`

---

## Modulbetöltési flow (fejlesztői mód)

```
1. Felhasználó: Alt+F8 → AB_SaveToDrive
       │
       ▼
2. simple-macros.AB_SaveToDrive
       │
       ▼ Call moduleLoader.LoadAndRunModule(MODULE_KEY_DRIVE)
3. module-loader.LoadAndRunModule("drive")
       │
       ▼ GetAppRootPath() → config.installPath  (pl. OneDrive\Autoblatt\app)
       ▼ Eltávolítja az előző driveManager komponenst (ha van)
       ▼ ThisWorkbook.VBProject.VBComponents.Import("modules\drive-manager.bas")
       ▼
4. driveManager.RunDriveManager
       │
       ▼ Call SaveToDrive
5. SaveToDrive (orchestrátor)
       ├─ FetchPersonalData      ← data.xlsm olvasás
       ├─ PopulateDataSheet
       ├─ ApplyFotoblattFormatting
       ├─ NormalizeWorkbookText  ← bpl/sw/li csere
       ├─ ActivateFotoblatt
       └─ SaveVersionedWorkbook  ← OneDrive\Sheets\<év>\... mentés
```

A loader trükkje, hogy minden hívásnál friss `.bas` fájlt olvas — ezért a kódjavítás azonnal érvénybe lép, nem kell újratelepíteni.

---

## Konfiguráció: két szint

### `app/settings/system-config.json` (projekt-szintű)
- Helye: a projekt mappájában
- Cél: alapértelmezett értékek a fejlesztéshez
- Git-ben: a `.gitignore` kizárja (per-felhasználó konfig)

### `%LOCALAPPDATA%\Autoblatt\system-config.json` (felhasználó-szintű) — **ezt olvassa a runtime**
- Helye: `C:\Users\<USER>\AppData\Local\Autoblatt\system-config.json`
- Cél: a futási idejű kód innen olvas, mert a workbook bárhol futhat (`Documents\autoblatt\`), de a forrás fájloknak más a helye (`OneDrive\Autoblatt\app`).
- Generálja: `installer/local-installer.vbs` vagy `network-installer.vbs`

### Séma

```json
{
  "appName": "Autoblatt",
  "appVersion": "1.0.0",
  "installType": "local",
  "installPath": "C:\\Users\\...\\Autoblatt\\app",
  "installDate": "...",
  "userName": "...",
  "targetWorkbookPath": "C:\\Users\\...\\Documents\\autoblatt\\",
  "targetWorkbookName": "autoblatt.xlsb",
  "dataXlsmPath": "C:\\Users\\...\\Autoblatt\\app\\data\\data.xlsm",
  "saveRootFolder": "Sheets",
  "replaceRules": "bpl=BPL;sw=SW;li=Li",
  "emailSignatureMode": "outlook"
}
```

---

## Adatfolyam

### `AB_SaveToDrive`

```
data.xlsm (Personal lap, tablePersonal)
    │ Find: Windows username
    ▼
Aktív munkafüzet Fotoblatt lap:
    A2 ← username (nagybetűs)
    B3 ← Name
    F3 ← Shift
    I3 ← Date

Aktív munkafüzet Data lap:
    ← teljes tablePersonal másolat

Formázás (A2, B3, F3, I3, A6, C6, E6, G8, H6, C15, A9, G8 nagybetű)

Csere minden lapon:
    bpl→BPL, sw→SW, li→Li (REPLACE_RULES)

Mentés:
    OneDrive\Sheets\<év>\<A6 E6 G8>\<dátum>_<C6>_<E6>_<G8>_(PS-foto|PS-info).xlsm
    Verziózás (v2, v3, ...)
    Explorer megnyitása
```

### `AB_ImportImages`

```
DeleteAllShapesExcept(SHAPE_KEEP)  ← "Grafik 5" kivétel
    │
ResolveLocalImageFolder()
    ├─ Helyi munkafüzet: ActiveWorkbook.Path
    └─ SharePoint URL: OneDrive\Sheets\<év>\<dekódolt mappanév>
    │
GetImageFiles(folder)              ← .jpg/.jpeg/.png
    │
For each:
    PlaceImageInGrid(file, IMAGE_START_CELL, index)
        ├─ Pozicionálás: startCell.Top/Left + offset (felbontásfüggetlen!)
        ├─ Tájolás: WIA EXIF Orientation
        ├─ Méret: 9.5 cm szélesség (fekvő) vagy 4.75 cm magasság (álló)
        └─ Név: "Import_<filename>"
```

### `AB_SendEmails`

```
BuildEmailSubject = A6 + " " + C6 + " " + E6 + " " + G8

ResolveAttachmentPath:
    ├─ Helyi: ActiveWorkbook.FullName (UrlDecode)
    └─ SharePoint: SaveCopyAs %TEMP%\<dekódolt név>

Outlook.CreateItem(0)
    .Subject = subject
    ApplyEmailSignature(mail)   ← config alapján: "outlook" / "custom" / "none"
    .Attachments.Add attachment
    .Display
```

---

## Modul függőségek

```
core/config.bas      (önálló, csak konstansok)
       │
       ▼
core/utils.bas       (config.bas-t használja)
       │
       ▼
core/module-loader   (config.bas + utils.bas)
       │
       ▼
core/simple-macros   ─→ core/cleanup-helper, core/settings-ui
       │
       ▼ (Hívja a moduleLoader-en keresztül)
modules/drive-manager
modules/fill-panel
modules/helpre-panel
modules/image-importer
modules/email-sender

ui/buttons-installer  (önálló, közvetlenül hívható)
```

A munka-modulok egymást **nem** hívják. Mindegyik kizárólag a `config.bas` és `utils.bas` segédjeit használja.
