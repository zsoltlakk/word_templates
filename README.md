# Word Templates

## Folyószámla összesítő sablon generátor

Ez a repository egy Python scriptet tartalmaz, amely egy magyar banki **"Folyószámla összesítő"** Word (`.docx`) sablont generál placeholderekkel.

## Használat

### 1. Függőségek telepítése

```bash
pip install -r requirements.txt
```

### 2. Sablon generálása

```bash
python generate_folyoszamla_template.py
```

A script létrehozza a **`folyoszamla_osszesito_template.docx`** fájlt az aktuális könyvtárban.

## Fájlstruktúra

| Fájl | Leírás |
|------|--------|
| `generate_folyoszamla_template.py` | Fő generátor script |
| `requirements.txt` | Python függőségek (`python-docx`) |
| `folyoszamla_osszesito_template.docx` | Generált Word sablon (script futtatása után) |

## Sablon jellemzői

- **Fekvő (landscape) A4** formátum
- **Margók:** 1 cm mindenhol
- **Betűtípus:** Arial, 8pt alapértelmezett
- **Fejléc:** logó, dokumentum dátuma, tárgyév/hónap, oldalszám
- **Fő adattábla:** 11 oszlop, csoportfejlécekkel, 20 adatsor placeholderekkel

## Placeholder összefoglaló

| Placeholder | Leírás |
|-------------|--------|
| `{{LOGO}}` | Cég logó helye |
| `{{DATUM}}` | Dokumentum dátuma (pl. 2017. 06. 21.) |
| `{{EV}}` | Tárgyév (pl. 2016) |
| `{{HONAP}}` | Tárgyhónap (pl. május) |
| `{{OLDAL}}` | Oldalszám |
| `{{EV_ELOZO}}` | Előző hónap éve |
| `{{HONAP_ELOZO}}` | Előző hónap száma |
| `{{ELOZO_HAVI_ZAROEGYENLEG}}` | Előző havi záróegyenleg összege |
| `{{BANKI_DATUM_N}}` | N-edik sor banki dátuma (N = 1..20) |
| `{{SZAMLA_BEVET_N}}` | N-edik sor számla bevéte |
| `{{ATTUTORA_BEVET_N}}` | N-edik sor átutalásra bevét |
| `{{EGYEB_JOVAIRAS_N}}` | N-edik sor egyéb jóváírás |
| `{{SZAMLAROL_TULFIZ_N}}` | N-edik sor számláról túlfizetés visszautalás |
| `{{ATTUTOR_VISSZAUT_N}}` | N-edik sor átutalóról visszautalás |
| `{{HM_VGH_JOGI_N}}` | N-edik sor HM VGH jogi költség |
| `{{HM_EI_JOGI_N}}` | N-edik sor HM EI jogi költség |
| `{{BANKI_KEZ_KOLTSEG_N}}` | N-edik sor banki kezelési költség |
| `{{EGYENLEG_LEEMELES_N}}` | N-edik sor egyenleg leemelés |
| `{{ZAROEGYENLEG_N}}` | N-edik sor záróegyenleg |
| `{{MEGJEGYZES}}` | Megjegyzés szöveg |