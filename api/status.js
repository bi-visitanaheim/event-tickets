// GET /api/status -> what's switched on (email, SharePoint spreadsheet) and when the spreadsheet last updated
import { role, fail, mailReady, spReady, getMeta } from "./_lib.js";
export default async function handler(req, res) {
  if (!role(req)) return res.status(401).json({ error: "Please sign in again." });
  try { res.json({ email: mailReady(), sharepoint: spReady(), sync: spReady() ? await getMeta("sync") : null }); } catch (e) { fail(res, e); }
}
