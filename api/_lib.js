// Shared server code: storage (Upstash Redis / Vercel KV), access codes,
// the master Excel workbook, SharePoint sync and email.
import ExcelJS from "exceljs";
import SEED from "./_seed.js";

const P = "et:";
export const COLS = ["events", "requests", "settings"];
export const ID_RE = /^[A-Za-z0-9_\-.~:@+]{1,200}$/;

/* ---------- storage ---------- */
export async function redis(cmds) {
  const clean = v => String(v || "").trim().replace(/^["']+|["']+$/g, "").trim();
  const pick = (...names) => { for (const n of names) { const v = clean(process.env[n]); if (v) return v; } return ""; };
  /* Accepts Vercel's names (KV_REST_API_*) or the names Upstash shows (UPSTASH_REDIS_REST_*), with or without quotes. */
  let url = pick("KV_REST_API_URL", "UPSTASH_REDIS_REST_URL"); const tok = pick("KV_REST_API_TOKEN", "UPSTASH_REDIS_REST_TOKEN");
  if (url && !/^https?:\/\//i.test(url)) url = "https://" + url;
  if (!url || !tok) { const e = new Error("Storage is not connected yet" + (!url && !tok ? " (no storage address or token found)" : !url ? " (storage address missing)" : " (storage token missing)")); e.status = 503; throw e; }
  const r = await fetch(url.replace(/\/$/, "") + "/pipeline", {
    method: "POST", headers: { Authorization: "Bearer " + tok, "Content-Type": "application/json" }, body: JSON.stringify(cmds)
  });
  if (!r.ok) throw new Error("Storage error " + r.status);
  const j = await r.json();
  return j.map(x => { if (x.error) throw new Error(x.error); return x.result; });
}
export async function version() { const [v] = await redis([["GET", P + "ver"]]); return Number(v || 0); }
export async function readAll() {
  let out = await redis(COLS.map(c => ["HGETALL", P + c]));
  if (!out[0] || out[0].length === 0) { await seedOnce(); out = await redis(COLS.map(c => ["HGETALL", P + c])); }
  const data = {};
  COLS.forEach((c, i) => { const a = out[i] || [], o = {}; for (let k = 0; k < a.length; k += 2) o[a[k]] = JSON.parse(a[k + 1]); data[c] = o; });
  return data;
}
async function seedOnce() {
  const [ok] = await redis([["SETNX", P + "seeded", String(Date.now())]]);
  if (!ok) return;
  const cmds = [];
  Object.entries(SEED.events).forEach(([id, d]) => cmds.push(["HSET", P + "events", id, JSON.stringify(d)]));
  Object.entries(SEED.requests).forEach(([id, d]) => cmds.push(["HSET", P + "requests", id, JSON.stringify(d)]));
  cmds.push(["HSET", P + "settings", "app", JSON.stringify(SEED.settings)], ["INCR", P + "ver"]);
  await redis(cmds);
}
export async function write(op, col, id, data) {
  const k = P + col;
  if (op === "delete") return redis([["HDEL", k, id], ["INCR", P + "ver"]]);
  if (op === "set") return redis([["HSET", k, id, JSON.stringify(data)], ["INCR", P + "ver"]]);
  if (op === "update") {
    const [cur] = await redis([["HGET", k, id]]);
    if (!cur) { const e = new Error("Not found"); e.status = 404; throw e; }
    return redis([["HSET", k, id, JSON.stringify(Object.assign(JSON.parse(cur), data))], ["INCR", P + "ver"]]);
  }
  const e = new Error("Unknown operation"); e.status = 400; throw e;
}
export async function setMeta(key, val) { await redis([["SET", P + "meta:" + key, JSON.stringify(val)]]); }
export async function getMeta(key) { const [v] = await redis([["GET", P + "meta:" + key]]); return v ? JSON.parse(v) : null; }

/* ---------- access ---------- */
/* Codes are optional. With no codes set, the app is open to anyone with the link and Nana opens her view from
   her own link (it ends in ?nana). Set NANA_ACCESS_CODE and/or TEAM_ACCESS_CODE in Vercel to require codes again. */
export function openMode() { return !process.env.TEAM_ACCESS_CODE; }
export function nanaCodeSet() { return !!process.env.NANA_ACCESS_CODE; }
export function role(req) {
  const c = String(req.headers["x-access-code"] || req.query?.code || "").trim();
  if (c && process.env.NANA_ACCESS_CODE && c === process.env.NANA_ACCESS_CODE) return "nana";
  if (openMode()) {
    if (c) return null;
    const want = String(req.headers["x-role"] || req.query?.role || "");
    return want === "nana" && !nanaCodeSet() ? "nana" : "team";
  }
  if (c && c === process.env.TEAM_ACCESS_CODE) return "team";
  return null;
}
export function fail(res, e) { console.error(e); res.status(e.status || 500).json({ error: e.message || "Something went wrong" }); }

/* ---------- workbook (same layouts as Nana's spreadsheets) ---------- */
const MON = ["January","February","March","April","May","June","July","August","September","October","November","December"];
const DOW = ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"];
const HOLDS = ["approved","sent","used"], ACTIVE = ["new","confirmed","approval","backup"];
const RL = { new:"New request", confirmed:"Option confirmed", approval:"Waiting for approval", approved:"Approved", backup:"Approved as backup", sent:"Tickets sent", used:"Used", returned:"Came back", not_used:"Did not use", traded:"Traded", declined:"Not approved", cancelled:"Cancelled" };
const KEY = { "Donation":"B7DEE8","VA Team Ticket Request":"F2B600","VA Staff Raffle":"E4DFEC","Allocated to Visit Newport Beach":"B7DEE8","Allocated to Visit Huntington Beach":"03AEED","Sold":"A5A5A5","Did not use":"595959","Traded tickets in for another game":"FF0000" };
const pd = s => { const [y,m,d] = String(s).split("-").map(Number); return new Date(y, m - 1, d); };
const pdU = s => { const [y,m,d] = String(s).split("-").map(Number); return new Date(Date.UTC(y, m - 1, d)); };
const md = d => (d.getMonth() + 1) + "/" + d.getDate();
const t12 = t => { if (!t) return "TBA"; let [h, m] = t.split(":").map(Number); const ap = h >= 12 ? "PM" : "AM"; h = h % 12 || 12; return h + ":" + String(m).padStart(2, "0") + " " + ap; };
const amount = (t, p) => `(${t}) ticket${Number(t) > 1 ? "s" : ""}${p ? " + (1) parking pass" : ""}`;
const first = n => String(n || "").trim().split(" ")[0];
function typeFor(r) { const d = r.dept || ""; if (/Market|Social/.test(d)) return "Marketing Request"; if (/Destination|Services/.test(d)) return "Destination Services & Events Request"; if (/Sales/.test(d)) return "Sales Request"; if (/Commun/.test(d)) return "Communications Request"; if (/Tourism/.test(d)) return "Tourism Request"; if (/Sport/.test(d)) return "Sports Request"; return "VA Team Request"; }
function sentNote(r, ini) { if (!r.sentMethod) return ""; const d = r.sentDate ? pd(r.sentDate) : null; return (r.sentMethod === "physical" ? `Physical tickets to ${first(r.clientName)}` : `Transferred (${r.tickets}) tickets to ${first(r.clientName)}`) + ` - ${r.sentBy || ini} ${d ? md(d) : ""}`; }
const title = e => e.kind === "ducks" ? "Ducks vs. " + e.name : e.name;
const fill = hex => ({ type: "pattern", pattern: "solid", fgColor: { argb: "FF" + hex } });
function head(ws, cols) { const r = ws.addRow(cols); r.eachCell(c => { c.font = { bold: true, color: { argb: "FFFFFFFF" } }; c.fill = fill("125C60"); }); return r; }

export async function buildWorkbook(data) {
  const settings = data.settings.app || {}; const ini = settings.initials || "NC";
  const events = Object.entries(data.events).map(([id, e]) => ({ id, ...e })).sort((a, b) => (a.date + (a.time || "99")).localeCompare(b.date + (b.time || "99")));
  const reqs = Object.entries(data.requests).map(([id, r]) => ({ id, ...r }));
  const of = e => reqs.filter(r => r.eventId === e.id);
  const wb = new ExcelJS.Workbook(); wb.creator = "Event Tickets"; wb.created = new Date();

  // 1) Tracker — Nana's layout, four rows per event
  const ws = wb.addWorksheet("Tracker", { views: [{ state: "frozen", ySplit: 2 }] });
  ws.columns = [{ width: 7 }, { width: 9 }, { width: 30 }, { width: 10 }, { width: 28 }, { width: 28 }, { width: 28 }, { width: 32 }, { width: 50 }];
  ws.mergeCells("A1:I1"); ws.getCell("A1").value = "DUCKS / HONDA CENTER / ANGELS SEASON TICKETS  ·  updated " + new Date().toLocaleString("en-US", { timeZone: "America/Los_Angeles" });
  ws.getCell("A1").font = { bold: true, size: 13, color: { argb: "FF125C60" } };
  head(ws, ["DAY", "DATE", "OPPONENT", "TIME", "VISIT ANAHEIM REQUESTER", "ATTENDEE NAME", "TITLE", "COMPANY", "ADDITIONAL NOTES"]);
  let last = "";
  for (const e of events) {
    const d = pd(e.date), mk = MON[d.getMonth()].toUpperCase() + " " + d.getFullYear();
    if (mk !== last) { ws.addRow([mk]).font = { bold: true }; last = mk; }
    const rs = of(e), holders = rs.filter(r => HOLDS.includes(r.status)), line = rs.filter(r => ACTIVE.includes(r.status)).sort((a, b) => a.option - b.option);
    const L = [];
    if (e.hold) L.push([KEY[e.hold.type], [e.hold.type, e.hold.org || "", "", "", amount(e.hold.tickets, Number(e.hold.parking))]]);
    holders.forEach(r => { L.push(["F2B600", [r.requesterName, r.clientName, r.clientTitle, r.clientCompany, r.sentMethod ? sentNote(r, ini) : `(${r.tickets}) Tickets${r.parking ? " and (1) Parking Pass" : ""}`]]); L.push(["F2B600", ["", r.sendTo, "", "", [r.approver ? "Approved by " + r.approver : "", (r.attendees || []).filter(g => g && g.name).length ? "Guests: " + r.attendees.filter(g => g && g.name).map(g => g.name).join(", ") : ""].filter(Boolean).join(" · ")]]); });
    rs.filter(r => ["returned", "not_used", "traded"].includes(r.status)).forEach(r => L.push([r.status === "not_used" ? "595959" : r.status === "traded" ? "FF0000" : "", [r.requesterName, r.clientName, r.clientTitle, r.clientCompany, RL[r.status]]]));
    line.forEach(r => L.push(["", [r.option + ") " + r.requesterName, "", "", r.clientCompany, RL[r.status]]]));
    while (L.length < 4) L.push(["", ["", "", "", "", L.length === 0 && e.note ? e.note : ""]]);
    L.forEach(([c, v]) => { const row = ws.addRow([DOW[d.getDay()], md(d), e.name, t12(e.time), ...v]); if (c) for (let k = 1; k <= 4; k++) row.getCell(k).fill = fill(c); });
  }

  // 2) All requests — one row per request, as an Excel table (filter, sort, pivot)
  const wr = wb.addWorksheet("All Requests");
  const ev = id => events.find(e => e.id === id) || {};
  wr.addTable({ name: "Requests", ref: "A1", headerRow: true, style: { theme: "TableStyleMedium2", showRowStripes: true },
    columns: ["Request ID","Submitted","Event date","Event","Venue","Option","Status","VA requester","Requester title","Department","Requester email","Client","Client title","Client company","Business reason","Tickets","Parking","Send tickets to","Approved by","Decision note","Sent how","Sent date","Sent by","Notes","Guests","Sample"].map(name => ({ name, filterButton: true })),
    rows: reqs.sort((a, b) => String(ev(a.eventId).date).localeCompare(String(ev(b.eventId).date)) || a.option - b.option).map(r => { const e = ev(r.eventId); return [r.id, r.createdAt ? new Date(r.createdAt) : "", e.date ? pdU(e.date) : "", e.name ? title(e) : "", e.venue || "", r.option, RL[r.status] || r.status, r.requesterName, r.requesterTitle, r.dept, r.requesterEmail || "", r.clientName, r.clientTitle, r.clientCompany, r.reason, r.tickets, r.parking ? "Yes" : "No", r.sendTo, r.approver || "", r.decisionNote || "", r.sentMethod === "physical" ? "Physical tickets" : r.sentMethod ? "Transferred" : "", r.sentDate ? pdU(r.sentDate) : "", r.sentBy || "", r.notes || "", (r.attendees || []).filter(g => g && g.name).map(g => g.name + (g.title ? " (" + g.title + ")" : "")).join("; "), r.sample ? "Yes" : ""]; }) });
  ["B","C","V"].forEach(c => wr.getColumn(c).numFmt = "m/d/yyyy");
  wr.columns.forEach((c, i) => { c.width = [16, 12, 11, 30, 14, 8, 20, 20, 24, 26, 26, 20, 24, 28, 50, 8, 8, 26, 18, 24, 16, 11, 8, 30, 36, 8][i] || 14; });

  // 3) Allocation report (Joy) — one row per allocation, with a Month column
  const wa = wb.addWorksheet("Allocation Report");
  wa.columns = [{ width: 10 }, { width: 11 }, { width: 14 }, { width: 30 }, { width: 10 }, { width: 34 }, { width: 38 }, { width: 30 }];
  head(wa, ["MONTH", "DATE", "VENUE", "EVENT", "TIME", "COMPANY/ORGANIZATION", "TYPE", "Ticket Amount"]);
  for (const e of events) {
    const d = pd(e.date), m = MON[d.getMonth()].slice(0, 3).toUpperCase() + " " + d.getFullYear(), venue = e.venue === "Angel Stadium" ? "ANGELS" : "HONDA CENTER";
    const rows = [];
    if (e.hold) rows.push([e.hold.org || e.hold.type, /Newport|Huntington/.test(e.hold.type) ? "Allocated to DMO" : e.hold.type, amount(e.hold.tickets, Number(e.hold.parking))]);
    of(e).filter(r => HOLDS.includes(r.status)).forEach(r => rows.push(["Visit Anaheim", typeFor(r), amount(r.tickets, r.parking)]));
    if (!rows.length) rows.push(["", "", ""]);
    rows.forEach(x => wa.addRow([m, md(d) + "/" + String(d.getFullYear()).slice(2), venue, e.name, t12(e.time), ...x]));
  }

  // 4) Events
  const we = wb.addWorksheet("Events");
  we.columns = [{ width: 11 }, { width: 9 }, { width: 32 }, { width: 14 }, { width: 13 }, { width: 12 }, { width: 8 }, { width: 8 }, { width: 34 }, { width: 34 }, { width: 28 }];
  head(we, ["DATE", "TIME", "EVENT", "VENUE", "TYPE", "COMING SOON", "TICKETS", "PARKING", "NOT AVAILABLE BECAUSE", "ORGANIZATION", "NOTE"]);
  events.forEach(e => we.addRow([pdU(e.date), t12(e.time), title(e), e.venue, e.kind, e.comingSoon ? "Yes" : "", e.tickets, e.parking, e.hold ? e.hold.type : "", e.hold ? e.hold.org || "" : "", e.note || ""]));

  we.getColumn(1).numFmt = "m/d/yyyy";
  // 5) Color key
  const wk = wb.addWorksheet("Color Key"); wk.getColumn(2).width = 40; wk.addRow(["Color Coding"]).font = { bold: true };
  Object.entries(KEY).forEach(([l, c]) => { const r = wk.addRow(["", l]); r.getCell(1).fill = fill(c); });
  return wb.xlsx.writeBuffer();
}

/* ---------- Microsoft 365: SharePoint file + email (turn on with env vars) ---------- */
export const msReady = () => !!(process.env.MS_TENANT_ID && process.env.MS_CLIENT_ID && process.env.MS_CLIENT_SECRET);
export const spReady = () => msReady() && !!process.env.MS_DRIVE_ID;
async function msToken() {
  const r = await fetch(`https://login.microsoftonline.com/${process.env.MS_TENANT_ID}/oauth2/v2.0/token`, { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: process.env.MS_CLIENT_ID, client_secret: process.env.MS_CLIENT_SECRET, scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials" }) });
  const j = await r.json(); if (!j.access_token) throw new Error("Microsoft sign-in failed: " + (j.error_description || j.error)); return j.access_token;
}
export async function syncSharePoint(data) {
  if (!spReady()) return { skipped: true };
  const path = process.env.MS_FILE_PATH || "Event Tickets - Master.xlsx";
  try {
    const buf = await buildWorkbook(data || await readAll()); const tok = await msToken();
    const r = await fetch(`https://graph.microsoft.com/v1.0/drives/${process.env.MS_DRIVE_ID}/root:/${encodeURIComponent(path).replace(/%2F/g, "/")}:/content`, { method: "PUT", headers: { Authorization: "Bearer " + tok, "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }, body: Buffer.from(buf) });
    if (!r.ok) throw new Error("SharePoint upload " + r.status + (r.status === 423 ? " (the file is open and locked; it will update on the next change)" : ""));
    const j = await r.json(); await setMeta("sync", { at: Date.now(), ok: true, url: j.webUrl || "" }); return { ok: true, url: j.webUrl };
  } catch (e) { console.error(e); await setMeta("sync", { at: Date.now(), ok: false, error: e.message }).catch(() => {}); return { ok: false, error: e.message }; }
}
export const mailReady = () => !!process.env.SENDGRID_API_KEY || (msReady() && !!process.env.MAIL_FROM);
const escH = s => String(s || "").replace(/[&<>]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" })[c]);
function htmlMail(text) {
  return `<div style="font-family:Arial,sans-serif;color:#231F20;max-width:560px;margin:0 auto;line-height:1.5"><p style="font-size:12px;letter-spacing:1px;text-transform:uppercase;color:#43A3A3;font-weight:bold;margin:0 0 12px">Visit Anaheim · Event Tickets</p><div style="white-space:pre-wrap;font-size:15px">${escH(text)}</div><p style="color:#6B6668;font-size:12px;border-top:1px solid #B4D9E3;padding-top:10px;margin-top:22px">Visit Anaheim · visitanaheim.org</p></div>`;
}
export async function sendMail({ to, subject, text, replyTo, attachment }) {
  const list = String(to || "").split(/[;,]\s*/).map(s => s.trim()).filter(s => /@/.test(s));
  if (!list.length) { const e = new Error("Add an email address"); e.status = 400; throw e; }
  const from = process.env.MAIL_FROM || "events@visitanaheim.org";
  if (process.env.SENDGRID_API_KEY) {
    const body = { personalizations: [{ to: list.map(email => ({ email })) }], from: { email: from, name: process.env.MAIL_FROM_NAME || "Visit Anaheim Event Tickets" }, subject, content: [{ type: "text/plain", value: text }, { type: "text/html", value: htmlMail(text) }] };
    if (replyTo) body.reply_to = { email: replyTo };
    if (attachment) body.attachments = [{ content: attachment.base64, filename: attachment.filename, type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", disposition: "attachment" }];
    const r = await fetch("https://api.sendgrid.com/v3/mail/send", { method: "POST", headers: { Authorization: "Bearer " + process.env.SENDGRID_API_KEY, "Content-Type": "application/json" }, body: JSON.stringify(body) });
    if (!r.ok) throw new Error("Email didn't send (" + r.status + ")");
    return;
  }
  if (msReady() && process.env.MAIL_FROM) {
    const tok = await msToken();
    const message = { subject, body: { contentType: "HTML", content: htmlMail(text) }, toRecipients: list.map(address => ({ emailAddress: { address } })) };
    if (replyTo) message.replyTo = [{ emailAddress: { address: replyTo } }];
    if (attachment) message.attachments = [{ "@odata.type": "#microsoft.graph.fileAttachment", name: attachment.filename, contentBytes: attachment.base64 }];
    const r = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(process.env.MAIL_FROM)}/sendMail`, { method: "POST", headers: { Authorization: "Bearer " + tok, "Content-Type": "application/json" }, body: JSON.stringify({ message, saveToSentItems: true }) });
    if (!r.ok) throw new Error("Email didn't send (" + r.status + ")");
    return;
  }
  const e = new Error("Email isn't set up yet"); e.status = 503; throw e;
}
