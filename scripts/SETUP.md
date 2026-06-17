# PM Hub daily task generation — GitHub Actions setup

This replaces the Power Automate 6 AM flow. A scheduled GitHub Action runs
[`generate-tasks.mjs`](generate-tasks.mjs) once a day and writes recurring
tasks into the `PM_Activity` SharePoint list — app-only, no user sign-in,
$0 on GitHub-hosted runners.

Until the three secrets below are added, the scheduled run is a harmless
no-op (it logs "Secrets not set" and exits cleanly).

## 1. Create an app registration (app-only Graph access)

> The PM Hub browser app uses a *delegated* SPA registration. The scheduled
> job runs with nobody signed in, so it needs **application** permissions +
> a secret. Use a new registration (or add a secret to an existing app-only one).

1. Entra admin center → **App registrations** → **New registration**.
   - Name: `PM Hub Task Generator`
   - Supported account types: single tenant. Register.
2. **API permissions** → Add → Microsoft Graph → **Application permissions** →
   add **`Sites.ReadWrite.All`** → **Grant admin consent**.
   *(Application, not Delegated. The green check next to the permission confirms consent.)*
3. **Certificates & secrets** → **New client secret** → copy the **Value**
   immediately (shown once).
4. From **Overview**, copy the **Application (client) ID** and **Directory (tenant) ID**.

## 2. Add the secrets to this repo

GitHub → repo **Settings** → **Secrets and variables** → **Actions** → **New repository secret**:

| Secret name           | Value                                  |
|-----------------------|----------------------------------------|
| `PMHUB_TENANT_ID`     | Directory (tenant) ID                  |
| `PMHUB_CLIENT_ID`     | Application (client) ID                |
| `PMHUB_CLIENT_SECRET` | the client secret **Value** from step 1.3 |

(`PMHUB_SITE_ID` is optional — the script defaults to the PM Hub site.)

> Secrets are encrypted and are **not** exposed in logs, even though this repo
> is public. The generator only logs counts and task text — never employee
> names or emails — because Actions logs on a public repo are world-readable.

## 3. Test it

- Actions tab → **Generate daily PM tasks** → **Run workflow** (manual trigger).
- Watch the log: `Created N, skipped M already-existing, 0 failed.`
- Open PM Hub and confirm today's tasks appear for the right people/roles.

Dedup is safe to re-run: it skips any task that already exists for that
person + period, matching on either the recurring-task id or the task text,
so it won't double up with anything Power Automate already created.

## 4. Retire Power Automate

Once you've confirmed a scheduled (not just manual) run worked, turn **off**
the old Power Automate flow so the two don't both run. (Leaving both on is
safe — dedup prevents duplicates — but the flow is the thing you're trying
to get rid of.)

## Schedule

`.github/workflows/generate-tasks.yml` runs at **09:00 UTC** (4 AM EST /
5 AM EDT). GitHub cron is UTC and ignores daylight saving; change the `cron:`
hour if you want it later in the morning.
