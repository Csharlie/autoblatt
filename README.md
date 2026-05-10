# Autoblatt

VBA-alapú Excel automatizáció a Mercedes-Benz Manufacturing Hungary Kft. fotó- és infólap dokumentációhoz.

Az `autoblatt` az `autosheets` projekt teljes újraírása: tisztább architektúra, AB_ makró-prefix, szétbontott modulok, központi config, dokumentált felépítés.

---

## Licensing and Usage Notice / Lizenz- und Nutzungsbedingungen / Licenc és használati figyelmeztetés

**ENGLISH**
This repository is intended strictly for internal use by authorized employees of Mercedes-Benz Manufacturing Hungary Kft.
All contents, including Excel files and VBA macros, are the intellectual property of Sárdy Péter.
Unauthorized use, distribution, or modification is strictly prohibited.

**DEUTSCH**
Dieses Repository ist ausschließlich für den internen Gebrauch durch autorisierte Mitarbeiter von Mercedes-Benz Manufacturing Hungary Kft. bestimmt.
Alle Inhalte, einschließlich Excel-Dateien und VBA-Makros, sind geistiges Eigentum von Sárdy Péter.
Jegliche unbefugte Nutzung, Verbreitung oder Änderung ist strengstens untersagt.

**MAGYAR**
Ez a repozitórium kizárólag a Mercedes-Benz Manufacturing Hungary Kft. engedéllyel rendelkező munkatársainak belső használatára készült.
A benne található Excel fájlok és VBA makrók Sárdy Péter szellemi tulajdonát képezik.
A jogosulatlan felhasználás, terjesztés vagy módosítás tilos.

---

## Mappastruktúra

```
autoblatt/
├── core/                       # Mindig betöltött modulok
│   ├── config.bas              # Központi konstansok (cellák, lapok, útvonalak)
│   ├── utils.bas               # Közös segédfüggvények (string, JSON, log)
│   ├── module-loader.bas       # Dinamikus modul betöltő (fejlesztői mód)
│   ├── ribbon-callbacks.bas    # Szalag gomb handlerek
│   ├── ribbon-config.xml       # Ribbon definíció (céges policy miatt nem aktív)
│   ├── simple-macros.bas       # Alt+F8 belépési pontok (AB_ prefix)
│   ├── cleanup-helper.bas      # Klón takarítás (fejlesztői segéd)
│   └── settings-ui.bas         # Beállítások panel (InputBox menü)
├── modules/                    # Dinamikusan betöltött munka-modulok
│   ├── drive-manager.bas       # Adatok lekérése + mentés
│   ├── fill-panel.bas          # Fotólap kitöltőpanel
│   ├── helpre-panel.bas        # Kitöltési segédlet
│   ├── image-importer.bas      # Képek 2 oszlopos rácsban
│   └── email-sender.bas        # Outlook draft email
├── ui/                         # Worksheet-embedded UI
│   └── buttons-installer.bas   # 'AB Indító' lap gombokkal
├── data/
│   └── data.xlsm               # Személyzeti adatok (gitignore!)
├── dist/                       # Kiosztott standalone xlsb
├── installer/
│   ├── startup-manager.bas     # XLSB inicializáció
│   ├── local-installer.vbs
│   ├── network-installer.vbs
│   ├── create_autoblatt_shortcut.vbs
│   ├── utf8-1250.ps1           # Encoding konverter
│   └── build-xlsb.ps1          # Standalone xlsb builder
├── settings/
│   ├── system-config.json      # Felhasználó-specifikus konfig
│   ├── close-all-excel.vbs
│   └── delete-autoblatt-xlsb.vbs
├── docs/
├── launch-autoblatt.bat        # Kiosztott xlsb indítás
└── setup-launcher.bat          # Telepítő menü
```

---

## Két verzió

- **Fejlesztői `.xlsm`** — Itt él a projekt, a VBA modulokat a `core/` és `modules/` mappákból tölti be dinamikusan a loader. Git-ben verziókezelt.
- **Kiosztott `dist/autoblatt.xlsb`** — Standalone, beégetett VBA, nincs loader. Más felhasználók ezt kapják kézbe.

A `dist/autoblatt.xlsb` a `installer/build-xlsb.ps1`-vel generálható.

---

## Felhasználói makrók (Alt+F8)

| Makró | Funkció |
|-------|---------|
| `AB_SaveToDrive` | Személyzeti adat lekérés, formázás, OneDrive verziózott mentés |
| `AB_ImportImages` | Képek 2 oszlopos rácsban elhelyezve az aktív lapra |
| `AB_SendEmails` | Outlook draft email a munkafüzettel mellékletként |
| `AB_FillPanel` | Fotólap kitöltőpanel létrehozása |
| `AB_HelprePanel` | Kitöltési segédlet létrehozása |
| `AB_InstallButtons` | „AB Indító" lap létrehozása kattintható gombokkal |
| `AB_Settings` | Beállítások panel (utak, cellák, csere-szabályok) |
| `AB_Help` | Súgó |
| `AB_SystemTest` | Minden modul ellenőrzése |
| `AB_CleanupSystem` | Modul-klón takarítás (fejlesztői mód) |
| `AB_ListAllModules` | VBA komponensek listázása |

---

## UI

A céges policy nem engedi a custom ribbon betöltődését. Az elsődleges UI a **munkalapra ágyazott gombok** — futtasd egyszer az `AB_InstallButtons`-t, és egy „AB Indító" lap készül a megnyitott munkafüzetben kattintható gombokkal.

A ribbon XML (`core/ribbon-config.xml`) megmarad arra az esetre, ha egyszer engedélyezi a policy.

---

## Fejlesztői workflow

1. Klónozd a repót: `OneDrive\Autoblatt\app\`
2. Szerkeszd a `.bas` fájlokat egy szövegszerkesztőben (UTF-8)
3. Futtasd: `installer\utf8-1250.ps1` — átkonvertálja a fájlokat Windows-1250-re (ezt várja a VBA)
4. Nyisd meg a fejlesztői `autoblatt.xlsm`-et (vagy a `setup-launcher.bat` opció 0)
5. Tesztelj az Alt+F8 menüből

Új standalone build:
```
PowerShell -ExecutionPolicy Bypass -File installer\build-xlsb.ps1
```

---

## Telepítés (más felhasználók)

A `setup-launcher.bat` futtatása interaktív menüt ad:
- **0** — Autoblatt indítása (a meglévő `dist/autoblatt.xlsb`)
- **1** — Lokális telepítés (új `autoblatt.xlsb` létrehozása `Documents\autoblatt\`-ba)
- **2** — Hálózati telepítés (a `P:\` megosztásról)
- **3** — Excel bezárása
- **4** — Telepített `autoblatt.xlsb` törlése
