// POST /api/login {code} -> {role: "team" | "nana"}
import { role } from "./_lib.js";
export default function handler(req, res) {
  if (req.method !== "POST") return res.status(405).end();
  const code = String((req.body || {}).code || "").trim();
  const r = role({ headers: { "x-access-code": code }, query: {} });
  if (!r) return res.status(401).json({ error: "That access code didn't work. Check with Nana." });
  res.json({ role: r });
}
