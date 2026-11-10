# PG Creative Request Form (GitHub Pages, no login)

A form that writes into the "Creative Tracker/Requirement - 2026" tab of the PG Creative Playbook sheet
(blue columns only) and shows the tracker. GitHub Pages hosts the page, Apps Script is the bridge to the sheet.

## Set up

1. Open the sheet, Extensions > Apps Script. Paste `apps-script/Code.gs` into `Code.gs`.
2. Deploy > New deployment > Web app. Execute as: **Me**. Who has access: **Anyone**. Deploy, authorize, copy the URL ending in `/exec`.
3. Put that URL in `config.js` (`PG_API_URL`) and commit.
4. Repo Settings > Pages > Source: **GitHub Actions**. The `pages` workflow publishes this folder.
5. Open `https://<username>.github.io/<repo>/`.

After any change to `Code.gs`: Deploy > Manage deployments > Edit > New version > Deploy. The URL stays the same.

## Note

Access "Anyone" means anyone who has the link can add rows and change statuses, without logging in.
Share the link only with the team. The form can only write the blue columns.
