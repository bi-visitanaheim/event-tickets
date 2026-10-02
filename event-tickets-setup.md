# Event Tickets · Visit Anaheim

Honda Center and Angels ticket requests for Visit Anaheim leadership, managed by Nana Cho.

- **Requester view:** events calendar, request form (Nana's 10 questions, with a big 1-4 ticket picker and optional guest names), My requests, Approvals for C-suite approvers, and a How it works tab.
- **Nana's view:** workflow (New → Option confirmed → With approver → Approved → Tickets sent), calendar & events, tracker, Joy's allocation report, How it works, questions & answers, settings.
- **Master spreadsheet:** `Event Tickets - Master.xlsx` has the tracker (Nana's four-rows-per-event layout, guests included), every request (with a Guests column), Joy's allocation report, the event list and the color key. It's built from the live data and is always current.

## Ticket counts
Each request picks 1 to 4 tickets (and the parking pass yes/no). Approved requests hold seats, so two requests of 2 can share one event. The calendar, event panel, approval cards, emails and tracker all show the count, and the event panel shows how many are open and how many are requested but not yet approved. Guest names are optional and can be added later by the requester (My requests) or Nana (Edit details).

## How it's built

A plain static page (`index.html` + `shim.js`) with Vercel Functions in `api/`, the same pattern as the Vans order form.

| File | What it does |
| --- | --- |
| `index.html` | The whole app (VA Branded: Sharp Sans, Escapism teal) |
| `shim.js` | No sign-in screen (a code screen appears only if TEAM_ACCESS_CODE / NANA_ACCESS_CODE are set); live data; downloads; email |
| `api/db.js` | Reads and saves events, requests and settings |
| `api/login.js` | Reports which codes are required and checks them when set |
| `api/mail.js` | Sends the app's emails |
| `api/export.js` | Downloads the master spreadsheet (Nana's code only); Excel can also connect to it live |
| `api/cron.js` | Every morning: refreshes the SharePoint copy and emails Nana the spreadsheet |
| `api/_lib.js` | Storage, access codes, spreadsheet builder, SharePoint and email helpers |
| `api/_seed.js` | Starting data imported from Nana's sheets: 181 events and 154 requests (see sheet-import.md) |

## Setup (one time)

### 1. GitHub
Publish this folder as `bi-visitanaheim/event-tickets` (GitHub Desktop: Add local repository → Publish), or create the empty repo and add it to the Claude session so Claude can push.

### 2. Vercel project
The Vercel project `event-tickets` imports the repo. No build settings needed; `vercel.json` has them.

### 3. Storage (required)
Vercel → project → **Storage** → **Connect Database** → Upstash for Redis (create new, or connect the existing one). Vercel adds `KV_REST_API_URL` and `KV_REST_API_TOKEN` itself. All keys are stored under `et:` so a shared database is fine.
The first time the app opens, it loads Nana's real events and requests automatically.

### 4. Environment variables
| Name | Needed? | Value |
| --- | --- | --- |
| `TEAM_ACCESS_CODE` | No | Leave this out and there is no sign-in at all. Add it later to require a code again |
| `NANA_ACCESS_CODE` | No | Leave this out and Nana opens her view from her own link (the app address followed by `?nana`). Add it later to require a code for Nana's view, the tracker and downloads |
| `SENDGRID_API_KEY` | For email | Same SendGrid key as the Vans form |
| `MAIL_FROM` | For email | Verified sender, e.g. `events@visitanaheim.org` (default) |
| `NANA_EMAIL` | Optional | Where the morning spreadsheet goes (otherwise Nana's email in Settings) |
| `CRON_SECRET` | Recommended | Any long random string; protects the morning job |

### 5. SharePoint master spreadsheet (turn on when IT is ready)
Ask IT for a Microsoft Entra app registration with Microsoft Graph application permissions **Sites.Selected** (granted on the Business Intelligence site) and optionally **Mail.Send**. Then add:

| Name | Value |
| --- | --- |
| `MS_TENANT_ID` | Directory (tenant) ID |
| `MS_CLIENT_ID` | Application (client) ID |
| `MS_CLIENT_SECRET` | Client secret |
| `MS_DRIVE_ID` | Drive ID of the document library that holds the Event Ticket Request folder |
| `MS_FILE_PATH` | Path inside that library, e.g. `Event Ticket Request/Event Tickets - Master.xlsx` |

Once these are set, every change in the app rewrites `Event Tickets - Master.xlsx` in that SharePoint folder, and the morning job refreshes it again. With `MS_*` and `MAIL_FROM` set but no SendGrid key, email goes through Microsoft 365 instead.

Until then, Nana always has the data three ways: the **Download master spreadsheet** button on her Tracker tab, the morning email with the spreadsheet attached, and the live link (Tracker tab › Copy live link for Excel), which a SharePoint workbook can use through Data › From Web.

## Before sharing with leadership
1. Open Nana's link (ends in `?nana`) → Settings → add each approver's email (Ronnie Collins is primary) and Nana's email.
2. Delete any test requests made while trying the app.
3. Send leadership the link. The app opens straight to the events, with no sign-in. Send Nana her own link: the same address followed by `?nana`. Without codes, anyone who has the link can open the app, and anyone who has Nana's link can open her view, so share each link only with the people it is for.
