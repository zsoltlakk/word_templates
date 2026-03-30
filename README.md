# Word Templates

Magyar banki **Folyószámla összesítő** Word (.docx) sablon generátor.

## Fájlok

- `generate_folyoszamla_template.py` – Python script, amely legenerálja a Word sablont
- `requirements.txt` – Python függőségek
- `folyoszamla_osszesito_template.docx` – A legenerált Word sablon (placeholderekkel)

## Használat

```bash
pip install -r requirements.txt
python generate_folyoszamla_template.py
```

A script létrehozza a `folyoszamla_osszesito_template.docx` fájlt.

## Sablon struktúra

- **Fekvő (landscape) A4** formátum, 1 cm margókkal
- **Fejléc tábla** (2 sor × 3 oszlop, szegély nélkül): logó, cím, dátum, oldalszám
- **Fő adattábla** (Table Grid stílusú):
  - Csoportfejlécek: Jóváírások, Terhelések, Jogi költségek
  - 11 oszlop (Banki dátum → Záróegyenleg)
  - Előző havi záróegyenleg sor
  - 20 placeholder adatsor
  - Megjegyzés sor (összevont cella)

## Placeholderek

| Placeholder | Leírás |
|---|---|
| `{{LOGO}}` | Cég logó helye |
| `{{DATUM}}` | Dokumentum dátuma |
| `{{EV}}`, `{{HONAP}}` | Tárgyév és hónap |
| `{{OLDAL}}` | Oldalszám |
| `{{EV_ELOZO}}`, `{{HONAP_ELOZO}}` | Előző év/hónap |
| `{{ELOZO_HAVI_ZAROEGYENLEG}}` | Előző havi záróegyenleg összege |
| `{{BANKI_DATUM_N}}` | N-edik sor banki dátuma (N=1..20) |
| `{{SZAMLA_BEVET_N}}` | Számla bevét |
| `{{ATTUTORA_BEVET_N}}` | Áttutóra bevét |
| `{{EGYEB_JOVAIRAS_N}}` | Egyéb jóváírás |
| `{{SZAMLAROL_TULFIZ_N}}` | Számláról túlfizetés visszautalás |
| `{{ATTUTOR_VISSZAUT_N}}` | Áttutóról visszautalás |
| `{{HM_VGH_JOGI_N}}` | HM VGH jogi költség |
| `{{HM_EI_JOGI_N}}` | HM EI jogi költség |
| `{{BANKI_KEZ_KOLTSEG_N}}` | Banki kezelési költség |
| `{{EGYENLEG_LEEMELES_N}}` | Egyenleg leemelés |
| `{{ZAROEGYENLEG_N}}` | Záróegyenleg |
| `{{MEGJEGYZES}}` | Megjegyzés sor |
