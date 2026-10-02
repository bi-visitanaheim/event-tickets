// GET  /api/db?since=<version>  -> everything the app needs (events, requests, settings)
// POST /api/db {op, col, id, data} -> save one change. Events and settings are Nana-only.
// Every change also refreshes the master spreadsheet in SharePoint when that's switched on.
import { waitUntil } from "@vercel/functions";
import { COLS, ID_RE, role, fail, version, readAll, write, spReady, syncSharePoint } from "./_lib.js";

export default async function handler(req, res) {
  const r = role(req);
  if (!r) return res.status(401).json({ error: "Please sign in again." });
  try {
    if (req.method === "GET") {
      const since = Number(req.query.since || 0), v = await version();
      if (since && since === v) return res.json({ ver: v, unchanged: true });
      const data = await readAll();
      return res.json({ ver: await version(), data });
    }
    if (req.method === "POST") {
      const { op, col, id, data } = req.body || {};
      if (!COLS.includes(col) || !ID_RE.test(String(id || ""))) return res.status(400).json({ error: "Bad request" });
      if (col !== "requests" && r !== "nana") return res.status(403).json({ error: "Only Nana can change this." });
      if (op !== "delete" && (typeof data !== "object" || data === null || Array.isArray(data))) return res.status(400).json({ error: "Bad data" });
      if (op !== "delete" && JSON.stringify(data).length > 100000) return res.status(413).json({ error: "Too large" });
      await write(op, col, id, data);
      if (spReady()) waitUntil(syncSharePoint());
      return res.json({ ok: true, ver: await version() });
    }
    res.status(405).end();
  } catch (e) { fail(res, e); }
}
