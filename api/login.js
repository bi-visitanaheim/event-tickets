// GET  /api/login        -> {open, nanaCode}: open = leadership needs no code, nanaCode = Nana's view needs a code
// POST /api/login {code, want} -> {role: "team" | "nana"}
import { role, openMode, nanaCodeSet } from "./_lib.js";
export default function handler(req, res) {
  if (req.method === "GET") return res.json({ open: openMode(), nanaCode: nanaCodeSet() });
  if (req.method !== "POST") return res.status(405).end();
  const b = req.body || {}, code = String(b.code || "").trim();
  const r = role({ headers: { "x-access-code": code, "x-role": String(b.want || "") }, query: {} });
  if (!r) return res.status(401).json({ error: code ? "That code didn't work. Check with Nana." : "Please enter the access code." });
  res.json({ role: r });
}
