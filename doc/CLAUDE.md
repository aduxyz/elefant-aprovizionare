# CLAUDE.md — Proiect aprovizionare Elefant.ro

Documentație completă pentru continuarea lucrului pe alt cont Claude sau cu alt developer.

---

## Reguli de lucru

- **Nu actualiza fișiere de documentație** (CLAUDE.md, project-steps.md etc.) decât la cerere explicită.
- **Nu face git push** decât la cerere explicită.
- **Nu face git commit** decât la cerere explicită.

---

## Ce este proiectul

Sistem de automatizare a aprovizionării cu carte pentru Elefant.ro. Sursa de date este un export din ERP-ul EBS (fișier Excel ~60k rânduri, sheet MAIN). Un script Google Apps Script procesează MAIN-ul și generează automat sheet-urile de lucru: aprovizionare, comenzi, reeditări, dashboard furnizori, listă comenzi.

Miezul proiectului este formula din coloana K (Cantitate) — câte exemplare să comande dintr-un titlu. Formula a fost dezvoltată iterativ prin backtesting pe date istorice SQLite.

---

## Structura repo

```
elefant-aprovizionare/
├── doc/
│   ├── CLAUDE.md           ← acest fișier
│   └── project-steps.md    ← evoluția proiectului în termeni business
│
├── data/                   ← gitignore: *.xlsx, export-erp/
│   ├── erp-export-DOI.xlsx ← snapshot EBS BRUT la 2026-03-22 (nu în git)
│   ├── export-erp/         ← XLS-uri ERP din GDrive (nu în git)
│   │   ├── 01_achizitii/   ← 011_Achizitii_DOI_*.xlsx, 012_Achizitii_RETUR_DOI_*.xlsx
│   │   └── 02_vanzari/     ← 021_Vanzari_DOI_*.xlsx
│   └── rebuild_db.py       ← reconstruiește elefant-erp.db din zero
│
├── etc/                    ← documente, transcrieri, misc
│
├── scripts/
│   ├── aprovizionare.gs    ← FIȘIERUL PRINCIPAL (Google Apps Script)
│   ├── build_stock_history.py
│   ├── simulate_procurement.py
│   ├── backtest_v1.py
│   ├── backtest_v2.py
│   └── procurement/        ← modul Python (config, backtest, generate_orders, import_ebs, normalize_db)
│
└── .gitignore
```

`elefant-erp.db` (SQLite, ~2 GB) nu e în repo — se reconstruiește cu `python data/rebuild_db.py`.

---

## Scriptul principal: scripts/aprovizionare.gs

### Setup Google Apps Script

Scriptul funcționează ca **librărie GAS** cu identificatorul `Aprov`. Există un singur proiect de librărie și 3 spreadsheet-uri host (test + 2 useri).

**Librărie** (script standalone):
- URL: `https://script.google.com/home/projects/18ofdqiBdy2DhmdOX8rrgA8su68kKI8eb52RIyDmi3DeCv0UKcszYXEcI/edit`
- Conținut: `scripts/aprovizionare.gs` (tot codul)

**Host script** (atașat fiecărui spreadsheet, minim):
```javascript
function onOpen()                        { Aprov.onOpen(); }
function generateSheets()                { Aprov.generateSheets(); }
function exportOrdersToNewSpreadsheet_() { Aprov.exportOrdersToNewSpreadsheet(); }
function exportAllOrders_()              { Aprov.exportAllOrders(); }
function onEdit(e)                       { Aprov.onEdit(e); }
```

**Reguli librărie GAS:**
- Funcțiile cu `_` suffix sunt private în librărie — nu pot fi apelate din host. De aceea există wrapper-ele publice `exportOrdersToNewSpreadsheet()` și `exportAllOrders()`.
- `onOpen` și `onEdit` sunt simple triggers — trebuie să fie în host script, nu în librărie.
- Menu item callbacks se caută în host script, nu în librărie.

**Versioning:**
- Spreadsheet-ul de test folosește **HEAD** (development mode) — vede codul imediat la salvare.
- Spreadsheet-urile userilor folosesc o **versiune numerotată** (ex. `2`) — nu se actualizează până când tu nu creezi explicit o versiune nouă în Deploy → Manage deployments → edit → New version.

### Convenții fișier .gs

Linia 2 și ultima linie sunt sincronizate:
```javascript
// ============================================================
const VERSION = 'v2.44'; // 2026-04-09 12:48:20
...
// v2.44 (2026-04-09 12:48:20)
```

La fiecare modificare: bump VERSION pe linia 2 + actualizează ultima linie.

### Cum funcționează generateSheets()

1. **Citește MAIN** (export EBS, 33 coloane, ~12k rânduri)
2. **Grupuri reeditări**: same Titlu+Autor+Furnizor, EAN-uri diferite → sumează vânzările grupului
3. **Filtrare** (în ordine):
   - Excludere: Epuizat=1 sau Indisponibil=1
   - Inclus dacă: SalesLY_efectiv > 0 AND (SalesLM − Stoc) >= 0
   - Inclus dacă: SalesLY=0 AND DataCreare în ultimele 30 de zile (RECENT_TITLES const)
4. **Deduplicare reeditări**: per grup, rămâne doar ediția cu PublishDate cel mai recent — și doar dacă aceasta trece filtrul
5. **Sortare**: Furnizor desc după total Necesar → Necesar desc → PublishDate desc
6. **Scrie aprovizionare** (46 coloane) via `_proc_helper`
7. **Generează**: reeditări, dashboard furnizori, comenzi, listă comenzi

### Coloanele din aprovizionare (46 total)

```
A-I   : ElefantSKU, EAN, Cod Articol, Articol, Autor, Furnizor, RRP, Reducere, Stoc Online
J     : Necesar (=O-I) — formulă în sheet
K     : Cantitate — formula v4_adaptive (verde+bold) — formulă în sheet
L     : DOS (=IFERROR((K+I)/S,0)) — formulă în sheet
M     : Disp. Rec.
N-R   : SalesL1W, SalesLM, SalesL2M, SalesLS, SalesLY
        → pentru reeditări: suma tuturor edițiilor din grup
S     : VZ medie/zi (=N/7*0.6+MAX(0,(P-O)/30)*0.4) — formulă în sheet
T-AK  : restul coloanelor din MAIN (PublishDate, Categorie, etc.)
AL-AT : REEDITARE, Nr_Editii, Alte_EAN, Chilipir, Spike_vz, Suprastoc, Ruptură, Zombie, Noutăți
```

### Formula Cantitate K (v4_adaptive, MOQ=2)

```
=LET(q,
  MIN(
    MAX(0, ROUND(
      MIN(SalesLM, MAX(SalesL2M-SalesLM, SalesL2M*0.55))
      * IF(AND(SalesLM>=10, SalesL1W*4<SalesLM),
          POWER(SalesL1W*4/SalesLM, 0.7)*0.8+0.2, 1)
      * 1.15 - StocOnline,
    0)),
    MAX(5, SalesL1W*4)   ← hard cap
  ),
  IF(q>0, MAX(2,q), 0)   ← MOQ=2
)
```

Implementare JS identică în `evalQuantity_()` (folosită pentru sortare înainte de scriere).

### Scriere aprovizionare (v2.45+)

Toate cele 46 de coloane sunt construite în JS din `sortedRows[i]` (date MAIN) + `sortedSums[i]` (sume reeditări) + `sortedExtra[i]` (flag-uri), apoi scrise printr-un singur `setValues`. Apoi `setFormulas` suprascrie J, K, L, S cu formulele live.

Nu mai există `_proc_helper` sau ARRAYFORMULA — elimina un bug de sincronizare în care ARRAYFORMULA evalua asincron față de setValues și producea mix de date din rânduri diferite în același rând din aprovizionare.

### _log sheet

Loghează toate operațiile cu timing. Format:
- Col A: Timestamp (ora reală a operației)
- Col B: Operație
- Col C: Delta (ms)
- Col D: Total (ms)

Separatorul de run arată: `── run start v2.44 (total: 79 s) ──`

Cele mai recente run-uri apar primele (insert la top).

### Parsing date

`new Date("26.10.2020")` returnează NaN în JS (month 26 invalid). Toate datele trec prin `parseDate_()` care handle-uiește formatul `dd.mm.yyyy`:
```javascript
function parseDate_(d) {
  if (!d) return 0;
  if (d instanceof Date) return isNaN(d.getTime()) ? 0 : d.getTime();
  const m = String(d).match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (m) return new Date(Number(m[3]), Number(m[2])-1, Number(m[1])).getTime();
  const t = new Date(d).getTime();
  return isNaN(t) ? 0 : t;
}
```

### setSharing

```javascript
DriveApp.getFileById(targetSS.getId())
  .setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);
```
`ANYONE_WITH_LINK` e blocat de Workspace policy — folosești `DOMAIN_WITH_LINK`.

---

## Schema DB: elefant-erp.db

### Reconstruire

```bash
python data/rebuild_db.py            # reconstruiește din zero (~36s)
python data/rebuild_db.py --dry-run  # verifică că fișierele sursă există
```

Date necesare în `data/export-erp/` (download din GDrive: https://drive.google.com/drive/folders/16mkZPAFdALgSZVD9dErGy0JV_5QfKLoz):
- `01_achizitii/011_Achizitii_DOI_*.xlsx`
- `01_achizitii/012_Achizitii_RETUR_DOI_*.xlsx`
- `02_vanzari/021_Vanzari_DOI_*.xlsx`

`data/erp-export-DOI.xlsx` = snapshot EBS BRUT la 2026-03-22 (necesar pentru `stock_history`).

### Tabele

| Tabel | Conținut | Rânduri |
|-------|----------|---------|
| `purchases` | Recepții (NIR): date, article_code, qty | ~194k |
| `purchase_returns` | Retururi furnizori (NAC) | ~11k |
| `sales` | Vânzări: order_date, article_code, order_quantity, order_line_status | ~1M |
| `procurement_erp` | Snapshot EBS BRUT la 2026-03-22 (33 col, ~60k rânduri) | ~60k |
| `stock_history` | Stoc zilnic reconstruit per articol ⭐ | ~750k |
| `daily_sales` | Agregate pre-calculate | — |
| `daily_purchases` | Agregate pre-calculate | — |
| `daily_purchase_returns` | Agregate pre-calculate | — |

**Atenție**: `sales` folosește coduri PF*, `purchases` folosește coduri AM*/CM*. Nu se pot join-ui direct pe article_code.

Interval date: 2025-01-03 → 2026-03-25.

### stock_history — logică reconstrucție

```
stock[D] = ref_stock + cum_sum[D] − ref_cum
```
- `ref_stock` = `procurement_erp.stock_online` (snapshot 2026-03-22)
- `cum_sum[D]` = suma cumulată NIR − NAC − vânzări + retururi client până la D
- `ref_cum` = cum_sum la ultima zi ≤ 2026-03-22

Query stoc la dată arbitrară:
```sql
SELECT stock_online FROM stock_history
WHERE article_code='PF0002568485' AND stock_date<='2026-03-15'
ORDER BY stock_date DESC LIMIT 1;
```

---

## Backtesting

`scripts/backtest_v2.py` — pentru fiecare dată de simulare din `SIMULATION_DATES` (iulie 2025 → feb 2026):
1. Reconstruiește stocul la data T din `stock_history`
2. Calculează SalesL1W/LM/L2M/LS/LY la T din `sales`
3. Aplică formula de cantitate
4. Compară (Stoc_T + Cantitate) cu vânzările reale din [T, T+30]
5. Raportează abaterile ponderate (weight = min(SalesLM, 50))

Metrica principală: % din greutatea totală cu abatere > 30%.

**Notă**: date sparse înainte de aprilie 2025 — pentru backtesting folosește iulie 2025 încoace.

---

## Decizii de design

| Decizie | Motivul |
|---------|---------|
| Filtru SalesLY > 0 | Exclude articole moarte |
| Filtru Necesar >= 0 | Nu aproviziona suprastoc |
| Titluri noi (SalesLY=0, DataCreare < 30 zile) | Include titluri fără istoric de vânzări |
| Reeditări: doar ediția activă | Edițiile vechi nu mai pot fi comandate |
| Reeditări: suma vânzărilor | Cererea pentru un titlu = suma tuturor edițiilor |
| Hard cap MAX(5, S1W×4) | Previne supracomanda când spike-ul s-a terminat |
| MOQ = 2 | Cantitate minimă de comandat dacă e nevoie de reaprovizionare |
| Sortare după Necesar (nu Cantitate) | Necesar reflectă cererea reală; Cantitate e plafonată |
| DOS = (Cantitate + Stoc) / VZ/zi | Include cantitatea comandată în calculul zilelor de stoc |
| _proc_helper persistent | Evită overhead de 11s la creare/ștergere la fiecare run |
| DOMAIN_WITH_LINK (nu ANYONE_WITH_LINK) | Workspace policy blochează sharing extern |
| parseDate_ custom | new Date("dd.mm.yyyy") returnează NaN în JS |
