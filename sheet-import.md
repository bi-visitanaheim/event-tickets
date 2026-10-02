# Event Tickets: which sheet the data comes from (Oct 1, 2026)

The app's starting data is built from the two sheets Nana populated. Both live on SharePoint, site Business Intelligence - Apps, folder Shared Documents > Event Ticket Request:

- **Ducks - Angels Tickets.xlsx** (Nana's tracker): the 2026-27 Honda Center tab with options and notes, and the earlier tabs.
- **2026 Honda Center-Angels Allocation Report.xlsx (Nana and Joy Only)**: Joy's report for January to September 2026.

## What was imported
- 181 events: 97 at Honda Center (Jan 2026 to the 2026-27 season) and 84 Angels games.
- 154 requests: 140 past (marked Used), 13 upcoming Approved (Option 1) and 1 upcoming Option confirmed.
- Past donations, Visit Newport Beach and Visit Huntington Beach allocations and staff raffle entries come in as allocation entries, so Joy's report and the tracker match her sheets.
- Upcoming donations (Boston 10/16 and Colorado 12/27) stay as "not available" holds on the event.

## Assumptions to confirm with Nana
1. Upcoming claims with no ticket note were set to 4 tickets and 1 parking pass.
2. Every imported request is treated as approved by Ronnie (her sheet does not record who approved).
3. The 9/23 preseason row has a donation color but a team requester; it was treated as a sent request.
4. Individual donation recipients and past client emails were not imported on purpose.
5. Anything she changes in her sheet after the import does not flow in by itself. Make changes in the app from now on.

## Re-running the import
The script import_sheet.py (kept in the build folder) reads both workbooks and writes the starting data. To refresh: download the two files from SharePoint, run the script, rebuild, and load the result into the app database. Ask Claude to do this; it takes a few minutes and keeps any requests entered in the app.

## Single source of truth going forward
- The app is the working copy. Nana's master workbook (Event Tickets - Master.xlsx) is generated from it: Tracker in her four-rows-per-event layout, All Requests (now with a Guests column), Joy's report, events and the color key.
- Until the SharePoint auto-sync is on (needs IT), Nana gets the workbook three ways: the Download button on the Tracker tab, the morning email, and a live link for Excel.
