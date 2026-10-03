# Prosjektinstruksjoner

- Behandle alle HR-filer og personopplysninger som interne. De skal ikke legges i nettstedets statiske katalog, byggartefakter eller versjonskontroll.
- Nettstedet skal motta HR-grunnlag gjennom lokal filopplasting. Ingen eksempeldata med virkelige personer skal følge med bygget.
- Bruk fasittbaserte regresjonstester mot lokale, ignorerte filer i `data/` når de finnes. Testene skal hoppe over kontrollen tydelig dersom filene ikke er tilgjengelige.

<!-- git-github:start -->
## Git / GitHub

- Bruk Git for lokal versjonskontroll.
- `origin` er GitHub-repoet `https://github.com/ae-nkom/HRAArsrapport` (`ae-nkom/HRAArsrapport`). Fetch og push går mot `origin`.
- Standardgren er `main`.
- GitHub Pages publiseres automatisk fra `main` til `https://ae-nkom.github.io/HRAArsrapport/` gjennom GitHub Actions. Bare `nettside/build/` inngår i nettstedet.
- Bruk `gh` for GitHub-operasjoner som issues, pull requests og releases.
<!-- git-github:end -->
