# HR Årsrapport

En statisk, lokal rapportarbeidsflate for HR-årsrapporten. Regneark behandles i nettleseren og blir ikke sendt til eller lagret på en server.

## Lokal bruk

```bash
cd nettside
npm install
npm run dev -- --host 127.0.0.1
```

Åpne `http://localhost:3773/`, velg «Last opp flere filer» og last opp de fire XLSX-rådatafilene. Du kan også erstatte én fil om gangen i de fire kortene.

## Kontroller

```bash
cd nettside
npm test
npm run build
npm audit --audit-level=moderate
```

Produksjonsbygget skrives til `nettside/build`. Datamappen, e-postgrunnlaget og lokale standardfiler er eksplisitt utelatt fra versjonskontroll og bygg.

## Før publisering

- Kontroller at alle fire rådataområder er markert som klare.
- Bruk bare et fastlønnsuttrekk med eksakt dato 01.05 eller 31.12; løsningen bruker ikke 31.12 som erstatning for 01.05.
- Avklar alle varsler om umatchede ansatte, ukjente stillingskoder og deltidstall før rapportteksten publiseres.
- Manuelle sykefraværs- og tekstfelt lagres lokalt i nettleseren. Opplastede data lagres ikke mellom økter.

Repoet har Gitea som primær `origin`. GitHub oppdateres bare gjennom det konfigurerte enveis speilet.
