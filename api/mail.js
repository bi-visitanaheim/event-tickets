// POST /api/mail {to, subject, text, replyTo} -> sends from the Visit Anaheim address
import { role, fail, sendMail } from "./_lib.js";
export default async function handler(req, res) {
  if (!role(req)) return res.status(401).json({ error: "Please sign in again." });
  if (req.method !== "POST") return res.status(405).end();
  try { const { to, subject, text, replyTo } = req.body || {}; await sendMail({ to, subject: String(subject || "Event Tickets").slice(0, 200), text: String(text || "").slice(0, 20000), replyTo }); res.json({ ok: true }); }
  catch (e) { fail(res, e); }
}
