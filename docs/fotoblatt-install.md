# Fotoblatt – telepítés és működés (9.14)

## Fájlok

| Fájl | Mi ez | Hogyan kerül a munkafüzetbe |
|---|---|---|
| `workbooks/fotoblatt.xlsm` | a munkafüzet | – |
| `src/fotoblatt-sheet-events.vba` | a Fotoblatt lap kódja (eseménykezelők) | **beillesztés** a shFotoblatt modulba |
| `src/fotoblatt-form-logic.bas` | VB modul: `modFotolap` | Import |
| `src/fotoblatt-setup.bas` | VB modul: `modBeallitas` | Import |
| `src/fotoblatt-parts-import.bas` | VB modul: `modTermekadatok` | Import |

> ⚠️ A `src/lbl-product-data-import.bas` NEM ide tartozik. Ugyanaz a modulneve (`modTermekadatok`), ezért ha ide importálod, felülírja a Fotoblatt modulját.

## A lapkód cseréje (ha csak az változik)
0. Készíts egy másolatot a `fotoblatt.xlsm`-ről.
1. Alt+F11 → a bal oldali listában dupla klikk a **shFotoblatt (Fotoblatt)** elemre.
2. A jobb oldali ablakban Ctrl+A, majd Delete.
3. Nyisd meg a `fotoblatt-sheet-events.vba`-t Jegyzettömbben, Ctrl+A, Ctrl+C, és illeszd be a VBE-be.
4. Debug → Compile VBAProject. Ha hibát jelez, küldd el a sort.
5. Mentsd el xlsm-ként.

A lapkód nem makró: nem kell futtatni, a lap módosításakor magától lefut.

## Modulok frissítése / teljes újratelepítés
1. Alt+F11 → `modFotolap`, `modBeallitas`, `modTermekadatok`: jobb klikk → Remove (No az exportra).
2. Ha van `frmA6Kereso` űrlap, azt is töröld (elavult).
3. File → Import File: `fotoblatt-form-logic.bas`, `fotoblatt-setup.bas`, `fotoblatt-parts-import.bas`.
4. A lapkódot a fenti 1–3. lépés szerint illeszd be.
5. Debug → Compile VBAProject, majd mentés.
6. Csak új vagy sérült munkafüzetnél: Alt+F8 → `SetupFotolap`. Ez létrehozza a neveket és a legördülőket, rendbe teszi az Article oszlopait és a típuskódokat. A végén rákérdez a termékadatok beolvasására: Igen → `lbl-product-data.xlsm`.

## Új termékadatok (pl. megjelenik a V4)
1. Előbb az `lbl-product-data.xlsm`-et frissítsd (l. `lbl-product-data-install.md`).
2. Fotoblatt: Alt+F8 → `TermekadatokFrissitese` → válaszd az `lbl-product-data.xlsm`-et. A fájlválasztó legközelebb ugyanitt nyílik (`TermekadatokUtvonal` név).
   - A tblTermek táblát név szerint keresi, ha nincs, fejléc alapján.
   - A Parts!Article R-sorait teljesen újraírja, az A-sorokhoz nem nyúl.
   - Az „Eltűnt a forrásból” státuszú tétel kimarad, a „Kifutott” jelölést kap, minden más aktív.
   - Az ismeretlen típuskódot felveszi a Data!Type táblába.
   - A végén összegzést ad: darabszámok, Berendezés nélküli tételek, ismétlődések, új típuskódok.

## Hogyan működik a lap

| Cella | Működés |
|---|---|
| A2 | Felhasználó (legördülő). Belőle töltődik a B3 (név) és az F3 (műszak) a Data!User táblából. |
| I3 | Dátum. Valódi bevitelnél beíródik a mai dátum, az A2-be pedig a gép felhasználója. **A törlés nem számít bevitelnek**, így a lap teljesen üresre törölhető. |
| A6 | Cikkszám-legördülő, minden tétellel, szűrés nélkül: aktív R → kifutott R „(kifutott)” jelöléssel → A. Egy sor alakja: „Cikkszám · Megnevezés · Típus”. Gépelve a cikkszám elejétől szűr (szóköz- és kisbetűfüggetlenül); a legördülő nyilával bárhol kereshető. |
| C6 | Megnevezés (a választásból). |
| G6 | Típus (a választásból, ha a tételnek van). Nem szűri az A6 listát. |
| H6 | Gyártóberendezés. A-számnál 7.41, R-számnál a tétel saját értéke (hiányában 7.50). 7.41/7.22 → ServoTPS-, 7.50 → LBL-, Utómunka → NA-kérdések; más vagy üres → a Hibafeltárás-blokk rejtve. |
| E4/E6 | R-szám → „Alapanyag/Material” + „Platine”. A-szám vagy üres A6 → „Préselés/Abpressung”, és a gép által beírt „Platine” törlődik (a kézi E6-hoz nem nyúl). |

## Egyéb makrók (Alt+F8)

| Makró | Mire való |
|---|---|
| `ArticleTisztitas` | Egyszeri, opcionális. Előnézetet mutat, csak Igen után ír. Az A/N számokat szóköz nélküli alakra hozza, törli az azonos duplikátumokat, eltérő névnél az első marad (a többi az Ellenorzes oszlopba kerül), és jelöli a hiányzó megnevezést. |
| `A6ListaUjraepites` | Kézzel újraépíti az A6 listát (lapra lépéskor amúgy is frissül). |
| `FotolapUjraigazitas` | Újraigazítja a Hibafeltárás-blokkot és az A4 méretet. |
| `HorgonyokEllenorzese` | Kiírja, melyik sorban találja a szerkezeti horgonyokat. |
| `HibafeltarasMegmutatasa` | Megmutatja a Hibafeltárás-blokkot, ha rejtve ragadt. |

## Ne szerkeszd kézzel
- Parts!J oszlop (`A6Lista`): automatikus segédlista.
- Parts!Article `Lista` oszlop: képlet.
- R-sorok a Parts!Article-ben: a következő frissítés felülírja őket, a javítást az `lbl-product-data.xlsm`-ben végezd.
