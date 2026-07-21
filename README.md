# Illuminate - WordCount Tracker

A Microsoft Word task-pane add-in that tracks:

- Time spent during the active session
- Number of words added
- Number of words deleted

## Prerequisites

- Node.js 20 to 24
- npm 9 to 11
- Microsoft Word desktop (Microsoft 365)

## Install

```bash
npm install
```

## Run In Word

```bash
npm start
```

This command builds the add-in, starts the local dev server, and sideloads `manifest.word.xml` into Word.

## Stop

```bash
npm stop
```

## Build

```bash
npm run build
```

## Production Manifest (Go Live)

Use [manifest.word.prod.template.xml](manifest.word.prod.template.xml) as the production manifest.

1. Replace every `YOUR_DOMAIN_HERE` with your real HTTPS domain.
2. Host build output so these URLs are publicly reachable:
	- `/taskpane.html`
	- `/commands.html`
	- `/assets/icon-16.png`
	- `/assets/icon-32.png`
	- `/assets/icon-80.png`
3. Validate the manifest:

```bash
npm run validate -- manifest.word.prod.template.xml
```

4. Share the finalized manifest for sideload testing or deploy centrally via Microsoft 365 admin center.

### Placeholder Domain Note

A placeholder domain will not work for live installs. Office clients must be able to load real HTTPS URLs from your manifest.

### Go-Live Checklist

- Real public HTTPS domain configured
- Hosted web assets available at that domain
- Manifest URLs updated from localhost to production domain
- Manifest validation passes
- Test sideload on a second machine/account

## Azure Static Web Apps

This project is prepared to deploy to Azure Static Web Apps from GitHub.

### What You Need

1. A GitHub repository containing this project.
2. An Azure Static Web App resource.
3. A repository variable named `PROD_URL` with the full site URL, for example `https://polite-wave-123456.azurestaticapps.net`.
4. A repository secret named `AZURE_STATIC_WEB_APPS_API_TOKEN` from the Azure Static Web Apps deployment page.

### How It Works

1. GitHub Actions runs `npm ci` and `npm run build`.
2. The build uses `PROD_URL` to replace localhost manifest URLs.
3. The `dist` folder is deployed to Azure Static Web Apps.

### First-Time Setup

1. Push this project to GitHub.
2. In Azure, create a Static Web App connected to that repository.
3. In GitHub repository settings:
	- Add variable `PROD_URL`
	- Add secret `AZURE_STATIC_WEB_APPS_API_TOKEN`
4. Push to `main` to trigger deployment.
5. After deploy, use the hosted domain in [manifest.word.prod.template.xml](manifest.word.prod.template.xml) or the generated `dist/manifest.word.prod.template.xml`.

### Local Production Build Example

PowerShell:

```powershell
$env:PROD_URL = "https://your-app-name.azurestaticapps.net"; npm run build
```

This generates production manifest files in `dist` with your Azure domain substituted for localhost.

## How Tracking Works

- Tracking starts automatically when the task pane opens.
- Click **Start Session** only if you manually stopped tracking.
- The add-in snapshots document text every 2 seconds.
- It compares word-frequency maps between snapshots and accumulates:
	- Words Added
	- Words Deleted
- Click **Stop Session** to pause.
- Click **Show Productivity Stats** for daily totals and weekly summary.
- Click **Show Lifetime Stats** for all-time totals, weekly rollups, and streaks.
- Click **Reset** to clear only current session counters (confirmation required).
- Click **Reset Lifetime** in Lifetime Stats to clear all persisted document totals (confirmation required).

## Notes

- Session counters reset when you click **Reset** or reload the add-in pane.
- Lifetime stats are stored in the current Word document's add-in settings.
- The add-in counts word-level changes using normalized tokens (letters, digits, apostrophes).

## Custom Logo

Put your custom logo at [assets/logo-source.png](assets/logo-source.png) and I can wire it into all icon sizes used by the manifest.
