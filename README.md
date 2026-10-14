# R-rdimensjonering

Excel-ark for hydraulisk trykkberegning av industrirør med væsker.

Åpne **[workbooks/Trykkberegning_ror.xlsx](workbooks/Trykkberegning_ror.xlsx)** i Excel
eller LibreOffice Calc. Arbeidsboken er makrofri og leveres med beregnede
eksempelverdier. Endre de blå feltene; les kontrollfeltet og fanen **Veiledning**.
Excel anbefales for valglister: LibreOffice kan fjerne dimensjonens avhengige
valgliste ved lagring. Generatoren gjenoppretter denne listen etter beregning.

Arket inneholder:

- Rørvalg fra et vanlig DN/OD-utvalg for EN 10220 / ISO 4200 og ISO 1127, eller
  metrisk OD. Faktisk veggtykkelse og egen innvendig diameter kan oppgis.
- Karbonstål, rustfritt og duplex med valg av EN-produktstandard og egen ruhet.
- Vann ved flere temperaturer, en nominell oljeviskositet og egne verdier for
  tetthet og dynamisk viskositet. Volumstrøm angis i m³/h.
- Rette rørmeter, høydemeter opp/ned og valgfritt innløpstrykk.
- 25 rader for bend med ulike vinkler/radier, Y/T-stykker og egne K-verdier.
- Antall flensede rørkoblinger med K per kobling; full lysåpning har K=0 som
  idealisert startverdi.
- 25 ventilrader og 25 øvrige komponentrader med antall og trykkfall per enhet.
- Samlet trykkdifferanse i bar, trykkhøyde, hastighet, Reynolds-tall og en
  grafisk oversikt over trykkbidragene.

Y/T-stykker krever dokumentert K for den valgte strømningsveien. Ventil- og
komponenttrykkfall må gjelde den aktuelle volumstrømmen og oppdateres hvis denne
endres. Én komponent skal telles én gang.

Modellen bruker Darcy–Weisbach og Colebrook for turbulent rørfriksjon, 64/Re for
laminær strømning og Rennels for rundede bend. Overgangsområdet Re=2300–4000
blokkeres. Ved lave Reynolds-tall kreves egne dokumenterte K-verdier for bend.
Høydebidraget har fortegn; netto fall kan gi negativ nødvendig trykkdifferanse.

Beregningsområdet er én rørstrekning med konstant diameter og volumstrøm, med
Newtonsk, inkompressibel væske og lik hastighet ved endene. Arket beregner
hydraulisk trykktap, ikke tillatt arbeidstrykk, godstykkelse eller et helt
rørnettverk. Referansetabellene er begrensede utvalg/startanslag og må kontrolleres
mot faktisk rør og leverandørdata. Detaljer og kilder står i arbeidsboken.

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

## Regenerering og beregningskontroller

Fra roten av repositoriet, med Python-miljøet over:

```bash
python scripts/build_pressure_workbook.py --output /tmp/ny_trykkberegning.xlsx
python tests/test_pressure_workbook.py
```

Bruk `.venv/bin/python` lokalt eller
`/workspace/.cloud-onboarding/excel/venv/bin/python` i skymiljøet som `python` i
kommandoene over. LibreOffice beregner lagrede formelresultater ved generering.
Generatoren nekter å overskrive en eksisterende fil uten `--overwrite`. Denne
opsjonen skal bare brukes til regenerering av malen, ikke på utfylte brukerfiler.

Regresjonskontrollen kjører 16 gyldige beregningstilfeller og 22 tilfeller med
ugyldige/manglende inndata gjennom LibreOffice. Turbulente resultater sammenlignes
med en uavhengig Colebrook-løsning med biseksjon, laminære med Poiseuille, og
bendmodellen med et publisert Rennels-eksempel. I tillegg kontrolleres fortegn,
antall, enheter, flenser, valglister, formler og at leveringsfilen har komplette
beregnede resultater uten Excel-feilkoder.
