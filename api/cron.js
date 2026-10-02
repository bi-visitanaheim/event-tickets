// Runs every morning (see vercel.json): refreshes the SharePoint spreadsheet and emails Nana
// a copy of the master workbook, so she always has the latest data even if the app is down.
import { readAll, buildWorkbook, spReady, syncSharePoint, mailReady, sendMail, fail } from "./_lib.js";
export default async function handler(req, res) {
  if (process.env.CRON_SECRET && req.headers.authorization !== "Bearer " + process.env.CRON_SECRET) return res.status(401).end();
  try {
    const data = await readAll(); const out = {};
    if (spReady()) out.sharepoint = await syncSharePoint(data);
    const s = data.settings.app || {}; const to = process.env.NANA_EMAIL || s.coordinatorEmail;
    if (mailReady() && to) {
      const buf = Buffer.from(await buildWorkbook(data)); const today = new Date().toLocaleDateString("en-US", { timeZone: "America/Los_Angeles", month: "long", day: "numeric", year: "numeric" });
      await sendMail({ to, subject: "Event Tickets spreadsheet · " + today, text: `Good morning,\n\nAttached is today's copy of the Event Tickets master spreadsheet: your tracker, every request, Joy's allocation report and the event list.\n\nSave it to SharePoint whenever you like; the app keeps its own copy too.`, attachment: { filename: "Event Tickets - Master.xlsx", base64: buf.toString("base64") } });
      out.emailed = to;
    }
    res.json({ ok: true, ...out });
  } catch (e) { fail(res, e); }
}
