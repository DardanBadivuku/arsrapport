# Årsrapportgenerator

Last opp Word-mal + Excel → få én PDF per kunde som ZIP.

## Deploy til Railway (gratis, 5 minutter)

1. Gå til **https://github.com/new** og lag et nytt repository kalt `arsrapport`
2. Last opp alle filene i denne mappen til repoet
3. Gå til **https://railway.app** og logg inn med GitHub
4. Klikk **"New Project"** → **"Deploy from GitHub repo"**
5. Velg `arsrapport`-repoet ditt
6. Railway deployer automatisk — du får en URL som `arsrapport.up.railway.app`

**Det er alt. Ingen terminal, ingen konfig.**

## Felter i Word-malen

| Plassholder i Word | Erstattes med |
|---|---|
| `Kunde` | Firmanavn |
| `Adresse….` | Gateadresse |
| `Post/sted` | Postnr + poststed |
| `V/ ` | V/ Kontaktperson |

## Kolonner i Excel

Programmet finner automatisk kolonnene: `Navn`, `Adresse`, `Postnr`, `Poststed`, `Kontaktperson`
