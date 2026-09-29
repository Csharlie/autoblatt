# Fejlesztés – kódfrissítés src-ből (code-loader 1.0.2)

A kód a git-repóban él, a munkafüzetek a OneDrive-on futnak. A kettőt a `KodFrissitese` makró köti össze: a munkafüzetben futtatva lecseréli annak kódját az `src/` mappában lévő változatra.

| Hely | Mi van ott |
|---|---|
| `C:\Users\pesardy\Projects\autoblatt\src\` | a kód (git) – itt szerkeszted |
| `OneDrive\Projects\autoblatt\fotoblatt.xlsm`, `lbl-product-data.xlsm` | a munkafüzetek – itt futtatod és próbálod |

## Egyszeri beállítás

1. Excel: File → Options → Trust Center → Trust Center Settings → Macro Settings → pipa: **Trust access to the VBA project object model**.
2. Mindkét munkafüzetbe egyszer kézzel: Alt+F11 → File → Import File → `src/code-loader.bas` (VB modul: `modKodFrissites`), majd mentés.

## Egy fejlesztési kör

1. Rövid ág a repóban: `git switch -c fix/…`
2. Szerkesztés az `src/` fájlban. Csak ASCII, ékezet `ChrW()`-vel (l. HANDOFF 1. pont); a fejléc `Verzio` sorát is frissítsd.
3. A munkafüzetben Alt+F8 → **`KodFrissitese`**:
   - felismeri a munkafüzetet (Fotoblatt + Parts lap → Fotoblatt; `tblTermek` tábla → LBL);
   - ellenőrzi a forrásfájlokat (létezik, csak ASCII, helyes VB-modulnév) – ha bármelyik hibás, semmit nem cserél;
   - kiírja, mit cserélne (fájl → modul, `Verzio` sor), és rákérdez;
   - elmenti a régi kódot: `%LOCALAPPDATA%\autoblatt\kodmentes\<munkafüzet>\<időbélyeg>\`;
   - lecseréli a modulokat, illetve a Fotoblatt lapkódját.
4. Debug → Compile VBAProject, próba, mentés.
5. `CHANGELOG.md`, commit, merge a `main`-be, tag (`v9.15`, `lbl-v1.3`), push.

| Munkafüzet | Fájl | Hova |
|---|---|---|
| fotoblatt.xlsm | `fotoblatt-form-logic.bas` | modul `modFotolap` |
| | `fotoblatt-setup.bas` | modul `modBeallitas` |
| | `fotoblatt-parts-import.bas` | modul `modTermekadatok` |
| | `fotoblatt-sheet-events.vba` | a Fotoblatt lap kódja (beillesztve) |
| lbl-product-data.xlsm | `lbl-product-data-import.bas` | modul `modTermekadatok` |

## Tudnivalók

- Az `src` mappát alapból a `%USERPROFILE%\Projects\autoblatt\src` helyen keresi (a `code-loader.bas` jelenléte alapján). Ha nincs ott, mappaválasztót nyit, és a választást a munkafüzet rejtett `KodForrasUtvonal` nevében megjegyzi.
- A `KodFrissitese` saját magát nem frissíti. Ha a `code-loader.bas` változik, kézzel kell újraimportálni (VBE → `modKodFrissites` → Remove, majd File → Import File).
- A VBE-ben végzett javítást a következő frissítés felülírja – a javítás helye az `src/`.
- A két `modTermekadatok` modul nem keveredhet: a fájl–munkafüzet összerendelés rögzített.
- Visszaállítás: a mentett `.bas` fájlok VBE → File → Import File-lal visszatölthetők, a lapkód (`.vba`) beillesztéssel.
