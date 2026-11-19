# Optimera fakturabehandling

Streamlit-appen leser én eller flere Optimera-fakturaer fra PDF og kobler
fakturalinjene mot kategoriseringsdata fra en Excel-fil. Resultatet kan lastes
ned som en Excel-arbeidsbok.

## Kjør appen

Installer avhengighetene:

```powershell
python -m pip install -r "requirements (1).txt"
```

Start appen:

```powershell
streamlit run "OptimeraFaktura_Streamlit (1) (1).py"
```

## Inndata og behandling

- Last opp én eller flere PDF-fakturaer.
- Last opp en `.xlsx`- eller `.xlsm`-kategoriseringsfil med arket `Linjer`.
- Varelinjer matches mot kategoriseringsfilen med NOBB-nummer eller artikkelnummer.
- Etter at fakturalinjene er vist i appen, velger brukeren én eller flere rader
  og tildeler kode `0`, `01`, `05` eller `12`.
- Kode `01` gir fag `RD`, `05` gir `BE`, og `12` gir `TØ`. Kode `0` lar fag stå
  tomt, men varelinjens beløp blir fortsatt med i summeringen.
- Kategoriene `Spesifikasjon` og `Frikode` hentes fra arket `Linjer`.
- Summeringsarket grupperer etter fakturanummer, Spesifikasjon, Fag og Frikode.
- Appen søker ikke etter varer med variasjon. Resultatfilen inneholder likevel
  arket `Varer med variasjon`, som alltid er tomt.
- Opplastede filer lagres i en midlertidig arbeidsmappe under behandlingen.

Excel-resultatet inneholder arkene `Kategorisert`, `Vareoversikt`,
`Ukjente varer` og `Varer med variasjon`.
