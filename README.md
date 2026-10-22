# Feedbackgesprekken VT1

Planning van de wekelijkse feedbackgesprekken (12 november tot 17 december 2026) voor VT1, bij Stijn en Joris.
Studenten zoeken hun naam en zien wanneer en bij wie ze verwacht worden. Docenten kunnen afvinken wie geweest is.

## Bestanden

- `index.html`: de pagina, met de volledige planning erin.
- `config.js`: hier komt het adres van de Google Sheet. Blijft het leeg, dan toont de pagina alleen de planning, zonder vinkjes.
- `apps-script.gs`: het script dat je in de Google Sheet plakt.

## Vinkjes aanzetten (eenmalig)

1. Maak een nieuwe, lege Google Sheet, bijvoorbeeld "Feedback VT1".
2. Kies **Extensies > Apps Script**. Wis wat er staat en plak de inhoud van `apps-script.gs`.
3. Vervang `kies-hier-een-code` bovenaan door een eigen docentencode en bewaar.
4. Kies bovenaan de functie **installeer** en klik op **Uitvoeren**. Geef de toestemming die Google vraagt.
   De Sheet krijgt nu een tabblad "Planning" met alle gesprekken en een vinkjeskolom.
5. Kies **Implementeren > Nieuwe implementatie**. Kies als type **Web-app**.
   Zet "Uitvoeren als" op jezelf en "Wie heeft toegang" op **Iedereen**.
6. Kopieer de URL die eindigt op `/exec` en plak ze in `config.js` tussen de aanhalingstekens.
7. Zet de wijziging op GitHub. Na een minuut werkt het afvinken.

## Afvinken

- Op de pagina: klik onderaan op "Docent? Meld je aan om af te vinken" en geef de docentencode. De browser onthoudt dat.
- Of rechtstreeks in de Google Sheet, in de kolom "geweest".

Studenten zien de vinkjes binnen een halve minuut.

## Planning aanpassen

Een wijziging in de planning moet in `index.html` (de pagina) en in het tabblad "Planning" van de Sheet gebeuren.
Pas je in `apps-script.gs` iets aan, kies dan **Implementeren > Implementaties beheren**, klik op het potlood en kies "Nieuwe versie". De URL blijft dan dezelfde.
