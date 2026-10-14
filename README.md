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
- Konsentriske reduksjoner/ekspansjoner i en sammenhengende diameterekjede,
  med lokalt K-tap og netto bidrag fra endret hastighet ved endene.
- Rørstrekk-tabell for friksjon i rette lengder med ulike innvendige diametre.
  Bend kan også få sin lokale innvendige diameter etter en reduksjon.
- 25 ventilrader med modellvalg, dokumentert trykkfall eller Kv ved aktuell DN,
  åpning og væske, eller eget K, samt 25 øvrige komponentrader. Generelle
  fullåpne ventiler beregner trykktap automatisk ved endret volumstrøm og ID.
- Separat fane **Rørstyrke** med Lamé/von Mises for indre overtrykk, korrosjon,
  negativ godstoleranse, dokumentert flytegrense ved temperatur, ønsket
  sikkerhetsfaktor, beregnet margin mot flyt og modellens trykkgrense.
- Samlet trykkdifferanse i bar, trykkhøyde, hastighet, Reynolds-tall og en
  grafisk oversikt over trykkbidragene. Navn og verdier følger stolpeenden:
  høyre ved positivt bidrag, venstre ved negativt bidrag.

Y/T-stykker krever dokumentert K for den valgte strømningsveien. Ventil- og
komponenttrykkfall som oppgis direkte må gjelde den aktuelle volumstrømmen og
oppdateres hvis denne endres. Generelle K-modeller og egne Kv/K beregnes på nytt
ved endret volumstrøm. Eget Kv skal dokumenteres ved aktuell DN, åpning og væske; det gjøres ingen
automatisk korreksjon for viskositet, kavitasjon eller choking. Én komponent
skal telles én gang.

Modellen bruker Darcy–Weisbach og Colebrook for turbulent rørfriksjon, 64/Re for
laminær strømning og Rennels for rundede bend. Overgangsområdet Re=2300–4000
blokkeres. Ved lave Reynolds-tall kreves egne dokumenterte K-verdier for bend.
Høydebidraget har fortegn; netto fall kan gi negativ nødvendig trykkdifferanse.

Beregningsområdet er én seriell strømningsvei med samme volumstrøm og Newtonsk,
inkompressibel væske. Reduksjonenes diametre følger hverandre fra hovedfanens
innløps-ID til siste utløps-ID. Med **Hoveddimensjon** legges alle rette meter
før første reduksjon; med **Rørstrekk-tabell** summeres de faktiske lengdene ved
hver diameter. Netto endehastighet regnes én gang, ikke som et dissipativt tap.
Høydeforskjellen har fortegn og skal ikke legges til friksjonslengden.

**Rørstyrke** er en idealisert kontroll av en rett, homogen sylinder med lukkede
ender under indre overtrykk. Dimensjonerende overtrykk er en egen inndata og
skal ikke erstattes av hydraulisk trykktap. Materialets flytegrense må oppgis
med kilde ved riktig temperatur og bekreftes før resultater vises. Nominelt
metallgods brukes til styrke, uavhengig av en overstyrt hydraulisk ID/belegg.
Dette er ikke full verifikasjon etter EN 13480: reduksjoner, flenser, termiske
laster, bøying, utmatting, kryp og utvendig trykk/stabilitet inngår ikke i
styrkekontrollen. «Modellens minste gods» er ikke et universelt normkrav.
Referansetabellene må kontrolleres mot faktisk rør og leverandørdata.

## Generelle ventiler: automatisk K, Kv og trykktap

Velg en av de tre generelle ventilene i **Ventiler**, og oppgi antall:

- **Kuleventil, full åpning:** DN15–DN300, K=3 fT. Modellen forutsetter full
  lysåpning; redusert ventilboring inngår ikke.
- **Skyvespjeld, sluseanslag:** DN50–DN300, K=8 fT. Dette er et foreløpig anslag
  fra Crane-modellen for sluseventiler, ikke dokumenterte skyvespjeldsdata.
  Det forutsetter fri passasje, full åpning og klar væske.
- **Sentrisk dreiespjeld:** DN50–DN200, K=45 fT; DN250–DN300, K=35 fT.
  Eksentriske ventiler og delvis åpning inngår ikke.

**Dette er modellestimater, ikke et verifisert snitt fra flere leverandører.**
Fagerberg, Bray, EBRO og Belimo var blokkert av nettverksproxyen ved innhenting.
Grunnlaget er Crane TP-410 (2009), kontrollert via den tilgjengelige gjengivelsen
i [fluids.fittings](https://github.com/CalebBell/fluids/blob/master/fluids/fittings.py).
Å oppgi konstruerte leverandørgjennomsnitt ville gi et uriktig datagrunnlag.
Modellene og kilde-URL-ene ligger i `data/generic_valve_models.json`.

**Ventildata** viser alle arket sine DN-valg, gyldighetsområdet, referanse-K og
ekvivalent Kv. Referanse-ID i denne tabellen settes til DN i mm for sammenligning;
det er ikke et mål på faktisk ventilboring. Utenfor modellområdet er verdiene
tomme. DN følger hovedfanen i ventilradene. Ved metrisk rørvalg må egnet DN velges
manuelt. Etter en reduksjon velges ventilens DN og faktisk lokal ID separat;
lokal ID skal finnes i reduksjonskjeden. DN ved redusert ID kan ikke utledes uten
å kjenne rørserie og godstykkelse der.

**K brukt** i ventilfanen beregnes med faktisk lokal ID. Den særskilte
Crane-referansefaktoren fT følger fluids sin tilpasning til tabulerte Crane-verdier:
Colebrook med Re*=7,5×10⁶ D og ε/D=3,4126825352925×10⁻⁵ D^(−1,0112), D i meter.
Dette er ikke rørets friksjonsfaktor ved aktuell væske og ruhet. Tilpasningen
gjengis i [fluids.friction](https://github.com/CalebBell/fluids/blob/master/fluids/friction.py).

Trykktapet beregnes som Δp[kPa]=Kρv²/2000 med hastighet ved lokal ID. Ekvivalent
Kv=3600 A √(200/K), A i m², er en omregning av K; det er ikke et produktdatablad.
Generelle modeller krever 100 % åpning og Re≥4000, eller null volumstrøm. Arket
beregner ikke en åpningskurve eller en automatisk viskositetskorreksjon.

Prioriteten per rad er **eget Δp → eget dokumentert Kv → eget K → generell
modell**. Eget K trenger en kilde og skal gjelde aktuell åpning, væske og lokal
ID; det kan brukes for andre strømningsregimer når dette er dokumentert.
Eget Kv har fortsatt felt for kilde, DN, åpning og bekreftelse.

## Fagerberg Norge: produktkatalogstatus

**Hele Fagerberg-katalogen er foreløpig ikke innhentet.** Nettverksproxyen avviser
tilgang til `fagerberg.no` og `www.fagerberg.no` med HTTP 403. Forslag om å
tillate domenene er lagret i miljøkonfigurasjonen, men tilgang må aktiveres der
før katalogen kan hentes og komplettheten kontrolleres.

Ventilfanen har valgliste fra **Ventilkatalog** og støtter generelle modellestimater,
dokumentert Kv, eget K eller eget trykkfall. Ingen Fagerberg-modeller eller
produkt-Kv er oppdiktet. Importgrunnlaget er `data/fagerberg_valves.csv`, med
kolonnene `name,category,manufacturer_model,url,retrieved`. Ved regenerering
brukes alle importerte rader, og datakilder og dato vises i katalogfanen.

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
python tests/test_pressure_extensions.py
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

Utvidelseskontrollen bruker 42 tilfeller for styrke/sikkerhetsfaktor, toleranser,
korrosjon, konsentriske overganger, fortegn, hastighetsbidrag, rørstrekk med ulike
diametre, lokal benddiameter, Kv og grafetiketter. Styrken kontrolleres mot
sylinderlikevekt/Lamé, og skarp reduksjon mot et publisert Crane-eksempel.
De 18 ventiltilfellene kontrollerer automatisk K/Kv/trykkfall, DN15 og DN300,
metrisk rørvalg, lokal diameter etter reduksjon, endret volumstrøm, null strømning,
egne K-verdier og avvisning av ugyldige DN, åpninger, regimer og manglende kilder.
