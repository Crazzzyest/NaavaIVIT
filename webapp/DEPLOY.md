# Naava Takst Webapp — Sliplane Deploy Guide

## Forutsetninger

- Sliplane-konto med tilgang til samme team som IVIT-scraperen (`naavaivit.sliplane.app`)
- Google Cloud Console-tilgang til OAuth-klienten (`819512973260-...`)
- Eksisterende refresh-token i `webapp/.env` (verifisert lokalt)

## Steg 1 — Push koden til GitHub

```powershell
cd "C:\Users\edson\Claude\ivit-webhook\.claude\worktrees\agitated-hopper-686df5"
git push origin claude/agitated-hopper-686df5
```

Lag deretter pull request fra `claude/agitated-hopper-686df5` til `main`, eller merge direkte hvis det er kun deg på repoen.

## Steg 2 — Opprett Sliplane-service

1. Sliplane-dashboard → **New Service**
2. Type: **Service from Git**
3. Repo: samme som IVIT-scraperen
4. **Branch**: `main` (etter merge) eller `claude/agitated-hopper-686df5` (for test)
5. **Root directory**: `webapp/` (slik at Sliplane bygger fra `webapp/Dockerfile`)
6. **Port**: la stå tom, eller `3001` — appen lytter på `process.env.PORT`

## Steg 3 — Sett env-vars i Sliplane

Kopier disse fra din lokale `webapp/.env`, men **endre verdier merket med ⚠️**:

```bash
# ── Modes (start trygt) ──
NODE_ENV=production
DEMO_MODE=false                                 # Real data
TEST_MODE=true                                  # ⚠️ Hold TRUE inntil verifisert
WEBAPP_CRON_ENABLED=false                       # ⚠️ Hold FALSE — Apps Script kjører cron

# ── Login OAuth (end-user sign-in) ──
LOGIN_OAUTH_CLIENT_ID=<eksisterende fra .env>
LOGIN_OAUTH_CLIENT_SECRET=<eksisterende>
LOGIN_OAUTH_CALLBACK_URL=https://<din-app>.sliplane.app/auth/google/callback   # ⚠️ Bytt til faktisk Sliplane-URL

# ── Session ──
SESSION_SECRET=<generer ny: openssl rand -hex 32>   # ⚠️ IKKE bruk lokal verdi

# ── Allowed users ──
ALLOWED_EMAILS=jacob@naava.no,afki@naava.no,edsongreistad99@gmail.com

# ── Google service-account (Sheets/Gmail/Drive) ──
GOOGLE_CLIENT_ID=<eksisterende>
GOOGLE_CLIENT_SECRET=<eksisterende>
GOOGLE_REFRESH_TOKEN=<eksisterende>

# ── OpenAI ──
OPENAI_API_KEY=<eksisterende>
OPENAI_MODEL=gpt-5.2

# ── IVIT webhook ──
IVIT_WEBHOOK_URL=https://naavaivit.sliplane.app/webhook
IVIT_WEBHOOK_SECRET=<eksisterende>

# ── Server ──
PORT=3001                                       # Sliplane setter ofte selv
```

## Steg 4 — Legg til Sliplane-URL som OAuth redirect URI

Når du vet Sliplane-URL-en (f.eks. `naava-webapp.sliplane.app`):

1. Gå til https://console.cloud.google.com/apis/credentials
2. Klikk OAuth-klienten (`819512973260-fs7p9n4qarjnqufquajlb5onfn70qi8n`)
3. Under **Authorized redirect URIs**, legg til:
   ```
   https://<din-app>.sliplane.app/auth/google/callback
   ```
4. **Save**

## Steg 5 — Deploy

Sliplane bygger Docker-image fra `webapp/Dockerfile` og starter automatisk. Følg loggene — du bør se:

```
Naava Takst webapp listening on port 3001
Test mode: true
Demo mode: false
Allowed emails: 3 configured
Cron jobs: DISABLED (WEBAPP_CRON_ENABLED=false). Apps Script handles automation.
```

## Steg 6 — Verifiser

Åpne `https://<din-app>.sliplane.app` i nettleseren:

1. **Login**: Klikk "Logg inn med Google" → bruk whitelistet konto → kommer inn på oppdragsliste
2. **Data**: Listen viser ekte oppdrag fra prod-sheet (114 rader)
3. **Stats**: Totalt, omsetning, snittpris stemmer med Dashboard-fanen i sheet
4. **Detalj**: Klikk et kort → 8 seksjoner med felter
5. **Mobil**: Test på mobiltelefon — kort-layout, dashboard og admin skal være lesbart
6. **Healthcheck**: `https://<din-app>.sliplane.app/health` returnerer `{ status: "ok", testMode: true, demoMode: false }`

## Sikkerhetsbrytere — anbefalt rekkefølge

### Uke 1 — Webapp som lesevindu
```
TEST_MODE=true
WEBAPP_CRON_ENABLED=false
```
Team blar gjennom oppdrag i webappen. Apps Script gjør all automatisering. Ingen risiko.

### Uke 2 — Edit aktivert, fortsatt test-mode
Beholdt samme env-vars. Team begynner å redigere felter via webappen i stedet for sheetet. PATCHes skriver til prod-sheet (det er det vi vil), men e-poster filtreres til jacob@naava.no + regnskap@naava.no fordi TEST_MODE.

### Uke 3 — Slå av test-mode
```
TEST_MODE=false
WEBAPP_CRON_ENABLED=false                       # fortsatt off
```
Faktura-e-poster og befaringsbekreftelser går nå til ekte mottakere. Sjekk én e-post manuelt før du stoler på det.

### Uke 4–8 — Flytt cron-jobber én og én
Hver uke: skru av én trigger i Apps Script-editoren + slå på tilsvarende cron i webappens admin-panel:

```
WEBAPP_CRON_ENABLED=true     # da starter alle de aktive cron-jobbene
```

Men siden hver jobb har egen cron-uttrykk i `Innstillinger`-sheetet, kan du sette uttrykket til en ugyldig verdi (`# off`) for jobbene som ikke skal kjøre ennå. Eller bare gjør dem manuelt fra admin-panelet.

Rekkefølge med lavest risiko:
1. IVIT-henting (uke 4)
2. Dashboard-oppdatering (uke 5)
3. Påminnelser (uke 6)
4. E-post-skanning (uke 7) — mest komplekst
5. Ukerapport (uke 8)

## Rollback

Hvis noe går galt:

**Tilbake til Apps Script-only:**
```
WEBAPP_CRON_ENABLED=false
DEMO_MODE=true                                 # webappen viser demo-data
```
Restart i Sliplane. Eventuelle manuelle edits stoppes. Apps Script tar fortsatt all automatisering.

**Skru av hele webappen midlertidig:**
Sliplane → Service → Stop

Apps Script-en kjører helt uavhengig, så ingenting forsvinner.

## Troubleshooting

| Symptom | Sjekk |
|---|---|
| `redirect_uri_mismatch` ved login | Steg 4 — er Sliplane-URL lagt til i Google OAuth-klient? |
| `unauthenticated` på /api-kall | SESSION_SECRET endret nylig? Logg ut og inn igjen |
| Tom oppdragsliste | GOOGLE_REFRESH_TOKEN gyldig? Sjekk Sliplane-loggen for "invalid_grant" |
| "Tilgang nektet" | E-post i ALLOWED_EMAILS? Restart trengs etter endring |
| 502 / 503 | Server crashet — sjekk Sliplane-logger |
| Faktura-e-post går til feil sted | TEST_MODE-flagg — kommer til jacob+regnskap kun mens TRUE |

## Vedlikehold

- Refresh-tokenet utløper ikke automatisk så lenge OAuth-klienten er "In production"
- Innstillinger-fanen i sheetet er sannhetskilden for alle bryterene i admin-panelet
- Aktivitetsloggen er in-memory — forsvinner ved redeploy, ikke noe permanent oppbevaring
