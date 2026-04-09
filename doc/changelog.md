# Changelog — aprovizionare Elefant.ro

---

## v2.45 (2026-04-09)

**Funcționalități majore**
- Eliminat `_proc_helper` și ARRAYFORMULA; toate cele 46 de coloane sunt acum construite direct în JS și scrise printr-un singur `setValues`

**Alte detalii**
- Fix bug critic: ARRAYFORMULA evalua asincron față de setValues, producând mix de date din rânduri diferite în același rând din aprovizionare
- Adăugat `changelog.md`; `project-steps.md` trunchiat la primele 7 pași

---

## v2.44 (2026-04-09)

**Funcționalități majore**
- `_log` redesenat: run-uri ordonate descrescător (cel mai recent primul), timestamp per operație, total embedded în separator (`── run start v2.44 (total: 79 s) ──`)

**Alte detalii**
- Eliminată coloana Versiune din `_log`; eliminat blocul de sumar F:G
- Restructurare repo: directoare `doc/`, `data/`, `scripts/`, `etc/`; git fresh cu SSH key dedicat

---

## v2.40

**Funcționalități majore**
- Librărie GAS partajată: un singur cod sursă, spreadsheet-urile de producție rămân pe versiune numerotată
- Dashboard furnizori cu checkbox „Activ" și export selectiv per furnizor sau pentru toți furnizorii bifați

**Alte detalii**
- Formule live pentru J/K/L/S (`setFormulas`) în loc de valori statice
- Sheet `_log` pentru diagnosticare timing

---

## v2.0

**Funcționalități majore**
- Formula v4_adaptive: plafon auto-corector + detecție spike S1W + hard cap; validată prin backtesting pe 8 date (iul 2025 – feb 2026)
- Reeditări: grupare Titlu+Autor+Furnizor, sume agregate, doar ediția cea mai recentă în lista de comandă
- Titluri noi: incluse dacă `DataCreare < 30 zile`, marcate cu flag `Noutăți`

**Alte detalii**
- MOQ=2; DOS calculat cu cantitatea inclusă
- Bază de date SQLite cu ~14 luni de tranzacții pentru backtesting
