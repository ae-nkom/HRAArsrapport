# Prosjektinstruksjoner

- Behandle alle HR-filer og personopplysninger som interne. De skal ikke legges i nettstedets statiske katalog, byggartefakter eller versjonskontroll.
- Nettstedet skal motta HR-grunnlag gjennom lokal filopplasting. Ingen eksempeldata med virkelige personer skal følge med bygget.
- Bruk fasittbaserte regresjonstester mot lokale, ignorerte filer i `data/` når de finnes. Testene skal hoppe over kontrollen tydelig dersom filene ikke er tilgjengelige.

<!-- setup-jj-gitea-mirror:start -->
## JJ / Gitea / GitHub

- Bruk Jujutsu (`jj`) for alle lokale versjonskontrolloperasjoner. Ikke bruk `git`-CLI.
- Lokal Gitea er primær forge og `origin`: `http://127.0.0.1:3000/admin/HRAArsrapport` (`admin/HRAArsrapport`).
- GitHub er et enveis push-speil: `https://github.com/ae-nkom/HRAArsrapport` (`ae-nkom/HRAArsrapport`). Ikke push direkte til `github`.
- Standard bookmark er `main`. Fetch og push skal gå mot `origin`; Gitea speiler videre til GitHub ved commit.
- Bruk Gitea MCP for støttede Gitea-operasjoner som repository-metadata, issues, pull requests, releases, labels, milestones, brukere, organisasjoner og hooks. Bruk Gitea REST API bare når MCP-en mangler nødvendig operasjon, særlig administrasjon av push-speil.
- Bruk `gh` kun for GitHub-side forgeadministrasjon som ikke er en lokal versjonskontrolloperasjon.
<!-- setup-jj-gitea-mirror:end -->
