# HR Årsrapport

En statisk rapportarbeidsflate for fastlønn, overtid, vakttillegg og foreldrepermisjon. XLSX-filer leses lokalt i nettleseren. Appen trenger ingen database eller tjeneste som mottar HR-grunnlaget.

## Start lokalt

Krever Node.js 22.12 eller nyere og npm.

```bash
cd nettside
npm ci
npm run dev -- --host 127.0.0.1
```

Åpne [appen lokalt](http://localhost:3773/). Bruk «Last opp alle fire samlet», eller velg en fil i hvert kort. Velg rapportår og lønnsdato etter opplasting.

## Grunnlag og arbeidsflyt

- Bruk rådata i XLSX-format. Bare det første regnearket leses; appen varsler dersom filen har flere ark. Bearbeidede fasitark skal ikke lastes opp som rådata.
- Fastlønnsfilen må ha eksakt uttrekksdato **01.05.ÅÅÅÅ** eller **31.12.ÅÅÅÅ** i filnavnet. Uttrekkene kan lastes inn samtidig og velges hver for seg. En manglende dato erstattes ikke med et annet uttrekk.
- Overtid og vakttillegg skal være komplette uttrekk for perioden, normalt hele kalenderåret. Året bestemmes av utbetalingsdatoen. Permisjon avgrenses til kalenderåret.
- Det brukes én kilde per område og år. En ny fil erstatter en overlappende kilde, slik at samme beløp ikke telles flere ganger. Slå sammen oppdelte årsuttrekk før opplasting. Filer som overlapper i samme opplastingsrunde avvises.
- Maksimalt 15 MB per fil, 16 filer og 60 MB per opplasting, 25 000 datarader og 250 kolonner per ark. Utpakket regneark begrenses til 100 MB. Importen kjører i en separat nettlesertråd med tidsavbrudd.
- Avklar ukjente stillingsgrupper og underdirektører uten oppgitt personalansvar i Datagrunnlag. `Personalansvar` med ja/nei eller en eksplisitt `Rapportgruppe` i kilden kan bestemme gruppen. Uavklarte personer inngår i totalen og vises separat. Personer kan også utelates uttrykkelig; avgrensningen varsles i rapporten.
- Gruppevalg gjelder bare det valgte uttrekket i den åpne økten. De nullstilles når fastlønnsfilen erstattes eller fjernes.
- Kontroller dataavvik, fyll inn manuelle tall under Årsrapport og last ned et redigerbart Word-utkast. Tomme tall vises som utfyllingsfelt; ugyldige tall stopper eksport. Faglige vurderinger, historikk og forklaringer ferdigstilles i Word.

## Beregningsprinsipper

Fastlønn viser gjennomsnittlig oppgitt årslønn uten vekting med stillingsprosent. Manglende lønn eller sammenligningsgrunnlag vises med tankestrek. Deltid bygger på registrert stillingsprosent; frivillig og ufrivillig deltid kan ikke skilles automatisk.

Overtid og vakttillegg summeres per mottaker og år. Gjennomsnitt og deltakelsesgrad bruker mottakere med positivt samlet årsbeløp som kan kobles til valgt fastlønnsuttrekk. Personer uten utbetaling inngår ikke i gjennomsnittet. Kildetotal, koblede og ukoblede beløp avstemmes separat. Nullbeløp og negative nettokorreksjoner opplyses særskilt. Sammenfallende navn med ulike personidentifikatorer kobles ikke vilkårlig.

Foreldrepermisjon beregnes som unike norske arbeidsdager per person, med offentlige helligdager trukket fra og fem dager per uke. Overlappende intervaller telles én gang. Uttaket vektes ikke med permisjonsgrad. Gjennomsnitt beregnes blant personer med minst én tellende dag.

## Personvern og lokale filer

Opplastede filer, personrader og gruppevalg ligger bare i minnet i den åpne appen. Nettleseroppdatering fjerner dem. Manuelle rapportfelt og menyinnstilling lagres lokalt i nettleseren; lagrede rapportfelt kan slettes per år. Appen fungerer også dersom lokal lagring er blokkert. Nedlastede rapporter må håndteres som interne dokumenter.

HR-grunnlag, e-post og lokale fasiter skal aldri ligge i `static/`, kildekoden eller versjonskontroll. Lokalt grunnlag ligger i den ignorerte `data/`-mappen på prosjektnivå. Bygget inneholder ingen eksempeldata fra virkelige personer. Produksjonsappens Content Security Policy blokkerer nettverksforbindelser fra appen; kode, skrifter og diagrammer leveres lokalt fra nettstedet.

## Kontroller

```bash
cd nettside
npm ci
npx playwright install chromium
npm run verify
npm audit --audit-level=low
```

`verify` kjører Svelte-kontroll, beregnings- og regresjonstester, produksjonsbygg, publiseringskontroll og nettlesertester. En allerede installert Chromium kan brukes med `PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH=/sti/til/chromium npm run verify`.

Nettlesertestene bruker utelukkende oppdiktede data og kontrollerer blant annet lokal behandling, Word-nedlasting, gruppevalg, år/dato, filbytte, tastatur og mobilbredde. Spor, video og skjermbilder er avslått i testoppsettet.

Fasittester bruker lokale rådata og bearbeidede regneark i `data/`, samt den lokale, ignorerte `data/local-regression-config.json`. Konfigurasjonen angir `snapshotKey` og `assignments` med kontrollerte gruppevalg per person. Den inneholder interne opplysninger og følger ikke med prosjektet. Testene sammenligner kjønnsfordeling, lønn, permisjon og utbetalinger mot fasitcellene. Kjente avvik mellom historisk rapporttekst og rådata testes uttrykkelig. Mangler kildene eller konfigurasjonen, markeres disse testene tydelig som hoppet over.

`npm run build` kontrollerer publiseringsfiler og mellomliggende bygg for uventede filtyper, interne mapper og symbolske lenker. Når lokale fastlønnsrådata finnes, kontrolleres også fullstendige personnavn og ellevesifrede identifikatorer. Dette supplerer regelen om aldri å legge HR-data i kildekoden; det er ikke en generell anonymiseringstjeneste.

## Publisering på GitHub Pages

Appen er tilpasset [GitHub Pages-adressen](https://ae-nkom.github.io/HRAArsrapport/). Arbeidsflyten i `.github/workflows/pages.yml` kjører ved oppdatering av `main` og kan også startes manuelt. GitHub-repoets Pages-kilde skal være **GitHub Actions**.

Arbeidsflyten installerer avhengigheter, kjører kodekontroll og tester, bygger appen med Pages-stien og kjører nettlesertester under samme sti. Bare innholdet i `nettside/build/` lastes opp og publiseres. Fasittester som trenger interne filer, hoppes tydelig over i GitHub; disse skal fortsatt kjøres lokalt før endringer sendes inn.

Slik kontrolleres Pages-bygget lokalt:

```bash
cd nettside
BASE_PATH=/HRAArsrapport npm run build
BASE_PATH=/HRAArsrapport npm run test:browser
BASE_PATH=/HRAArsrapport npm run preview -- --host 127.0.0.1
```

Forhåndsvisningen åpnes på `http://127.0.0.1:3773/HRAArsrapport/`. Vanlig lokal utvikling bruker fortsatt rotstien og krever ikke `BASE_PATH`.

GitHub Pages bruker ikke filen `_headers`. Appens CSP og referrer-policy er derfor også angitt i HTML og gjelder på Pages. Ekstra HTTP-hoder for blant annet innramming og Permissions-Policy krever en vert som støtter dem. Ingen persondata følger med nettstedet; filopplasting og beregning foregår fortsatt lokalt i nettleseren.

Ved bruk av en annen statisk HTTPS-vert kan appen bygges uten `BASE_PATH` for domenets rotsti. Publiser **bare innholdet i `nettside/build/`**. Ikke publiser prosjektroten, `data/`, `node_modules/` eller `.svelte-kit/`. `build/_headers` kan brukes av verter som støtter dette formatet.

Oppsettet følger [SvelteKits veiledning for GitHub Pages](https://svelte.dev/docs/kit/adapter-static#GitHub-Pages) og [GitHubs dokumentasjon for egendefinerte publiseringsarbeidsflyter](https://docs.github.com/en/pages/getting-started-with-github-pages/using-custom-workflows-with-github-pages).

## Rapportens omfang

Appen er gjennomgått mot e-posten og den vedlagte rapportmalen. Rapporttabellene vises først, med lesbare diagrammer av de samme målingene. Automatiske rangeringer og ekstragrafer er fjernet fra hovedflyten. Se [avgrensning mot oppdraget](docs/rapportomfang.md). Historikktabellen i Word omfatter rapportåret og fem foregående år.

## Kodestruktur

Arbeidsflaten ligger i `nettside/src/components/`. Import, beregninger, datoer, diagramgrunnlag, rapporttekst og Word-eksport har separate moduler i `nettside/src/lib/`. Endringer i beregningsmetode skal følges av relevante regresjonstester.

Git brukes for versjonskontroll. [GitHub-repoet i ae-nkom](https://github.com/ae-nkom/HRAArsrapport) er `origin`. Oppdateringer av `main` starter automatisk kontroll og publisering på GitHub Pages.
