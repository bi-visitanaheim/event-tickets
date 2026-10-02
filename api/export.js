// GET /api/export?code=<Nana's code>          -> the master spreadsheet (.xlsx), always current
// GET /api/export?code=<Nana's code>&sync=1   -> also push a fresh copy to SharePoint now
// Excel can connect to this link (Data > From Web) so a workbook in SharePoint refreshes from the app.
import { role, fail, readAll, buildWorkbook, spReady, syncSharePoint } from "./_lib.js";
export default async function handler(req, res) {
  if (role(req) !== "nana") return res.status(401).json({ error: "Nana's access code is needed." });
  try {
    const data = await readAll();
    if (req.query.sync && spReady()) await syncSharePoint(data);
    const buf = await buildWorkbook(data);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", 'attachment; filename="Event Tickets - Master.xlsx"');
    res.send(Buffer.from(buf));
  } catch (e) { fail(res, e); }
}
