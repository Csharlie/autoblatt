# LBL Product Data – telepítés és frissítés (1.2)

## Fájlok

| Fájl | Mi ez |
|---|---|
| `workbooks/lbl-product-data.xlsm` | a saját, optimalizált termékadatok (tábla: `tblTermek`) |
| `src/lbl-product-data-import.bas` | VB modul: `modTermekadatok` (LBL oldal) |

> ⚠️ Ez a modul csak ebbe a munkafüzetbe való. A Fotoblatt ugyanilyen nevű moduljával nem szabad összekeverni.

## Lapok
Áttekintés · Termékkártya · **Termékadatok** (`tblTermek`, a fejléc a 4. sorban) · Leképezés (`tblOszlop`, `tblSorozat`, `tblSuruseg`) · Változásnapló (`tblNaplo`) · Info (utolsó import, `nmForrasSzinek`, `tblFigy`) · Leírás · `Elozo_allapot` (rejtett, a frissítés előtti állapot).

## Modul (újra)telepítése
1. Alt+F11 → `modTermekadatok`: jobb klikk → Remove.
2. File → Import File: `lbl-product-data-import.bas`.
3. Debug → Compile VBAProject, majd mentés.
4. Csak ha nincs gomb: Alt+F8 → `FrissitoGombLetrehozasa` (a gomb az E1 cellához kerül).

## Frissítés új forrásverzióból (pl. V4)
1. Termékadatok lap → **Frissítés forrásból…** gomb, vagy Alt+F8 → `TermekadatokFrissitese`.
2. Válaszd ki a tulajdonos TERM_KADATOK fájlját. Csak olvasásra nyílik meg, és mentés nélkül zárul be.
3. Előnézet: új, módosult, státuszváltozás, eltűnt, figyelmeztetések. Csak Igen után ír.
4. Ellenőrizd: Változásnapló lap, Info lap figyelmeztetései, Adatjelzés oszlop.
5. Ha a forrásban új blokk (sorozat) jelenik meg, a Leképezés lap `tblSorozat` táblájában add meg a sorozatnevet és a típuskódot.
6. Utána frissítsd a Fotoblattot (`fotoblatt-install.md`).

## Mit ír és mit nem a makró

| Ír | Soha nem ír |
|---|---|
| a `tblOszlop`-ban felsorolt forrásoszlopokat | Típus (kézi), db/ütem, Ráhagyás, Saját megjegyzés |
| Termékszám (új tételnél), Státusz, Sorozat (forrás), Forrás kiemelés, Adatjelzés, Első/Utolsó forrás, Utolsó változás | képletes oszlopok (Típus, Súly számolt, Súly eltérés, Szalagsebesség számolt, Hiányzó adatok) |
| Sorozat – **csak ha üres** | az eredeti forrásfájlt |

**Státuszok:** Aktív · Kifutott · Aktív és kifutott – ellenőrizendő · Eltűnt a forrásból (a tétel nem törlődik).

**Színek:** a forrás sárga/piros jelölései mindig szövegként kerülnek a „Forrás kiemelés” oszlopba. A cellaszínek átvétele csak akkor történik meg, ha Info!`nmForrasSzinek` = Igen.

**Ellenőrzés:** az utolsó ismert import újrafuttatásakor az előnézetnek mindenhol 0-t kell mutatnia.
