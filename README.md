# autoblatt

Excel/VBA eszközök a MBMH Kecskemét PW Fotólap (Fotoblatt) kitöltéséhez és az LBL (7.50 Laser Blanking Line) termékadatok kezeléséhez.

| Munkafüzet | Mire való | Kód |
|---|---|---|
| `OneDrive\Projects\autoblatt\autoblatt.xlsm` | Fotólap: cikkszám-legördülő, automatikus név/típus/berendezés, Hibafeltárás-blokk, hibakép-lista, A4 igazítás | `src/fotoblatt-*` |
| `OneDrive\Projects\autoblatt\data\lbl-product-data.xlsm` | Saját, optimalizált LBL termékadatok, frissítés a tulajdonos TERM_KADATOK fájljából | `src/lbl-product-data-import.bas` |

## Mappaszerkezet

| Hely | Mi van ott |
|---|---|
| `~\Projects\autoblatt\` | a kód (ez a git-repó); a `KodFrissitese` innen tölt (`src\`) |
| `OneDrive\Projects\autoblatt\` | a futtatott munkafüzet: csak az `autoblatt.xlsm` |
| `OneDrive\Projects\autoblatt\data\` | minden, amit az autoblatt beolvas: `lbl-product-data.xlsm`, `Autoblatt_Cikkszamtorzs.xlsx`, `hibamegnevezések.xls`, `Fehlerarten_Hibamegnevezések.xls`, `cikkszamok\`, `part-details.xlsm`, `id-sfm-cikkszam-*` listák, `Pres_Termekadatok.xlsx`, `Hianyzo_A_szamok_ellenorzes.xlsx`, `v9_Utómunka formanyomtatványok a Présüzemben.xlsx` |
| `OneDrive\Autoblatt\<Év>\<Hónap>\` | a fotólapok (pl. `2026\10`); a céges példány a P: meghajtón (PW QS Reklamation, Informationsblätter) |
| `OneDrive\Autoblatt\Status\` | alkatrészenkénti, folyamatos státuszjelentések |
| `~\Projects\autoblatt-archive\` | minden archív anyag (régi generációk és bővítmények, elavult munkafüzetek, git-mentések), tartalomjegyzék: `ARCHIV.md` |

Munkafüzet nem kerül a repóba (`.gitignore`).

- Telepítés: [`docs/fotoblatt-install.md`](docs/fotoblatt-install.md), [`docs/lbl-product-data-install.md`](docs/lbl-product-data-install.md)
- Állapot, szabályok, teendők: [`HANDOFF.md`](HANDOFF.md)
- Verziók: [`CHANGELOG.md`](CHANGELOG.md)

Előzmény: a korábbi autoblatt bővítmény (az AutoSheets utódja) az `autoblatt-archive` repóban található.

Aktuális verzió: Fotoblatt **9.20** · LBL Product Data **1.2**

Karbantartó: Sárdy Péter (PESARDY). A kód módosítása előtt egyeztess vele.

> Belső céges adat: kizárólag privát, céges jóváhagyású repóban tárolható.
