# R-rdimensjonering

Arbeidsmiljø for å lage, lese og bearbeide Excel-filer (`.xlsx`) med Python.
Repositoriet inneholder foreløpig oppsett og en funksjonskontroll, ikke en ferdig dimensjoneringsmodell.

## Verktøy

- Python 3.12 anbefales.
- `openpyxl` leser og skriver arbeidsbøker, formler og formatering.
- `pandas` leser regneark som tabeller for videre analyse.
- LibreOffice Calc beregner formler uten et grafisk grensesnitt.

Microsoft Excel trenger ikke å være installert. Kompatibilitet med VBA, eksterne datakilder og alle Excel-funksjoner er ikke verifisert.

## Lokal installasjon

Kjør fra roten av repositoriet:

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
```

Installer også LibreOffice Calc med operativsystemets pakkebehandler. På Debian/Ubuntu:

```bash
sudo apt-get update
sudo apt-get install --no-install-recommends libreoffice-calc
```

Programmet `soffice` må være tilgjengelig i `PATH`.

## Kontroll av oppsettet

```bash
.venv/bin/python scripts/check_excel.py
```

Kontrollen oppretter en midlertidig arbeidsbok, skriver norske tegn og formatering,
beregner numeriske formler og `IF` med LibreOffice, og leser resultatene med både
`openpyxl` og `pandas`. Feil gir en avslutningskode som er ulik null. Testdataene er
kun til kontroll av verktøyene. De skal ikke brukes som en dimensjoneringsmodell.
Midlertidige filer fjernes etter kjøring; eksisterende arbeidsbøker endres ikke.

## Bruk i Codex-skymiljøet

Python-avhengighetene er installert i `/workspace/.cloud-onboarding/excel/venv`.
LibreOffice er tilgjengelig som `soffice` i miljøets `PATH`.

```bash
cd /workspace/R-rdimensjonering
/workspace/.cloud-onboarding/excel/venv/bin/python scripts/check_excel.py
```

Bruk samme Python-program når du kjører egne skript. Miljøet trenger ingen
bakgrunnstjeneste eller ekstra hemmeligheter for denne arbeidsflyten.

`openpyxl` beregner ikke formler. Etter formelendringer må arbeidsboken åpnes i Excel
eller beregnes med LibreOffice før lagrede resultater leses med `data_only=True`.
Beregn alltid en kopi i en separat utdatakatalog, og kontroller resultatene før
originalen eventuelt erstattes. LibreOffice kan endre formatering og funksjoner som
er spesifikke for Excel.

Neste prosjektsteg er å legge til arbeidsboken og beskrive inndata, formler,
enheter og forventede kontrollresultater for selve dimensjoneringen.
