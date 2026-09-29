# autoblatt

Excel/VBA eszközök a MBMH Kecskemét PW Fotólap (Fotoblatt) kitöltéséhez és az LBL (7.50 Laser Blanking Line) termékadatok kezeléséhez.

| Munkafüzet | Mire való | Kód |
|---|---|---|
| `workbooks/fotoblatt.xlsm` | Fotólap: cikkszám-legördülő, automatikus név/típus/berendezés, Hibafeltárás-blokk, A4 igazítás | `src/fotoblatt-*` |
| `workbooks/lbl-product-data.xlsm` | Saját, optimalizált LBL termékadatok, frissítés a tulajdonos TERM_KADATOK fájljából | `src/lbl-product-data-import.bas` |

- Telepítés: [`docs/fotoblatt-install.md`](docs/fotoblatt-install.md), [`docs/lbl-product-data-install.md`](docs/lbl-product-data-install.md)
- Állapot, szabályok, teendők: [`HANDOFF.md`](HANDOFF.md)
- Verziók: [`CHANGELOG.md`](CHANGELOG.md)

Előzmény: a korábbi autoblatt bővítmény (az AutoSheets utódja) az `autoblatt-archive` repóban található.

Aktuális verzió: Fotoblatt **9.14** · LBL Product Data **1.2**

Karbantartó: Sárdy Péter (PESARDY). A kód módosítása előtt egyeztess vele.

> Belső céges adat: kizárólag privát, céges jóváhagyású repóban tárolható.
