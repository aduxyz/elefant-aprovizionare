# Evoluția proiectului — aprovizionare Elefant.ro



## Punctul de plecare

Elefant.ro aproviziona cu carte manual: cineva se uita la un fișier Excel exportat din ERP și decidea câte exemplare să comande din fiecare titlu. Procesul era lent, inconsistent și depindea mult de experiența persoanei.

ERP-ul (EBS) poate exporta un snapshot al catalogului: ~60.000 de titluri cu stocul curent, vânzările recente (ultima săptămână, ultima lună, ultimele 2 luni, ultimul sezon, ultimul an) și alte date despre produs.

**Obiectivul**: un script care citește acest export și propune automat cantitățile de comandat.


## Pasul 1 — Înțelegerea datelor

Primul lucru a fost să înțelegem structura datelor disponibile:
- Exportul ERP are 33 de coloane per titlu: cod, EAN, furnizor, prețuri, stocuri, vânzări istorice pe mai multe ferestre de timp
- Vânzările din ultimul an (`SalesLY`) și ultima lună (`SalesLM`) sunt cele mai relevante
- Stocul online (`StocOnline`) arată câte exemplare sunt disponibile acum

S-a stabilit că o carte are nevoie de reaprovizionare dacă: vinde cel puțin o bucată pe an (`SalesLY > 0`) **și** stocul e mai mic decât ce s-a vândut luna trecută (`SalesLM > StocOnline`, adică `Necesar >= 0`).


## Pasul 2 — Prima formulă de cantitate

Formula inițială: comandă cât să acoperi 30 de zile de vânzări medii.

Problema: nu ținea cont de spike-uri (titluri care brusc s-au vândut mult luna trecută dar săptămâna asta se vând normal) și nici de suprastoc (titluri cu mulți ani de stoc).


## Pasul 3 — Backtesting pe date istorice

Pentru a valida și îmbunătăți formula, s-a construit o bază de date SQLite cu istoricul complet al tranzacțiilor pe 14 luni (ian 2025 – mar 2026):
- ~1 milion de înregistrări de vânzări
- ~200k recepții
- Stoc zilnic reconstruit per articol (din diferența dintre intrări și ieșiri)

Logica de backtesting: alegi o dată din trecut (ex. 1 septembrie 2025), simulezi ce ai fi comandat cu formula, și compari cu ce s-a vândut real în următoarele 30 de zile. Dacă ai comandat prea puțin sau prea mult față de cererea reală, formula are probleme.

S-au rulat simulări pe 8 date din intervalul iulie 2025 – februarie 2026.


## Pasul 4 — Formula v4_adaptive

Prin iterații de backtesting s-a ajuns la formula actuală cu trei mecanisme:

**1. Plafon auto-corector** (`MIN(SalesLM, MAX(SalesL2M - SalesLM, SalesL2M * 55%))`)
Limitează cantitatea la un nivel rezonabil față de vânzările din luna precedentă. Dacă luna trecută a fost un spike, corectează în jos.

**2. Detecție spike via vânzări recente** (`IF AND SalesLM>=10 AND SalesL1W*4 < SalesLM`)
Dacă ultima săptămână arată că vânzările au scăzut față de luna trecută, reduce proportional cantitatea (cu o funcție de putere pentru tranziție lină).

**3. Hard cap** (`MAX(5, SalesL1W*4)`)
Nu comanda niciodată mai mult decât ~4 săptămâni de vânzări recente. Previne comenzi masive când un spike s-a terminat.

**MOQ = 2**: cantitate minimă de comandat dacă e nevoie de reaprovizionare (nu merită o comandă pentru un singur exemplar).


## Pasul 5 — Reeditări

Problema: același titlu poate apărea de mai multe ori în ERP cu EAN-uri diferite (ediții diferite ale aceleiași cărți). Dacă ediția veche are vânzări istorice dar ediția nouă nu, filtrul `SalesLY > 0` ar elimina ediția nouă.

Soluția: grupăm titlurile cu același Titlu+Autor+Furnizor și EAN-uri diferite ca reeditări. Suma vânzărilor grupului se folosește pentru decizia de aprovizionare. Din grup apare în lista de comandă **doar ediția cea mai recentă** (cea cu data publicării cea mai mare) — pentru că edițiile vechi nu mai pot fi comandate de la furnizor.


## Pasul 6 — Titluri noi

Titlurile lansate recent (`DataCreare < 30 zile`) au `SalesLY = 0` prin definiție (nu au un an de istoric). Filtrul inițial le excludea complet.

Soluția: dacă `SalesLY = 0` **și** titlul e lansat în ultimele 30 de zile, îl includem cu o cantitate calculată pe baza vânzărilor disponibile. Aceste titluri sunt marcate cu flag `Noutăți=1` și colorate verde deschis în sheet.


## Pasul 7 — Google Apps Script

Formula finală a fost implementată ca script Google Apps Script atașat unui Google Sheet. Fluxul de lucru:

1. Utilizatorul importă exportul ERP în sheet-ul **MAIN**
2. Rulează „Pregătire aprovizionare" din meniu
3. Scriptul generează automat:
   - **aprovizionare** — lista cu toate titlurile de comandat, sortată pe furnizor
   - **comenzi** — filtrat per furnizor, gata de trimis
   - **reeditări** — tabel cu toate grupurile de reeditări detectate
   - **dashboard furnizori** — KPI-uri per furnizor (SKU-uri, stoc, vânzări, rupturi)
   - **listă comenzi** — log al exporturilor generate


