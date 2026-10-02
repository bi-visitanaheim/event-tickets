/* Event Tickets — connects the app to its Vercel back end.
   Provides sign-in (name, work email, access code), live data (checks for changes every 10 seconds),
   Excel downloads, email sending and the link to the master spreadsheet. */
(function () {
  "use strict";
  const KEY = "vaet-session";
  let sess = null;
  try { sess = JSON.parse(localStorage.getItem(KEY) || "null"); } catch (e) {}
  window.__VERCEL = true;
  const hdr = () => ({ "Content-Type": "application/json", "x-access-code": sess ? sess.code : "" });

  /* ---------- data ---------- */
  let ver = 0, cache = { events: {}, requests: {}, settings: {} }, loaded = false, timer = null;
  const listeners = [];
  const fire = () => listeners.forEach(f => { try { f(); } catch (e) { console.error(e); } });
  async function pull(force) {
    try {
      const r = await fetch("/api/db" + (force || !ver ? "" : "?since=" + ver), { headers: hdr(), cache: "no-store" });
      if (r.status === 401) { signOut(); return; }
      const j = await r.json();
      if (!r.ok) { showBanner(j.error || "The app can't reach its data right now."); return; }
      hideBanner();
      if (j.unchanged) return;
      ver = j.ver; cache = Object.assign({ events: {}, requests: {}, settings: {} }, j.data); loaded = true; fire();
    } catch (e) { showBanner("You're offline. Changes will show when you reconnect."); }
  }
  function start() {
    if (timer) return;
    pull(true);
    timer = setInterval(() => { if (document.visibilityState === "visible") pull(false); }, 10000);
    document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible") pull(false); });
  }
  const snap = (col, id) => ({ id, exists: !!(cache[col] || {})[id], data: () => (cache[col] || {})[id] ? JSON.parse(JSON.stringify(cache[col][id])) : undefined, metadata: {} });
  async function write(op, col, id, data) {
    const r = await fetch("/api/db", { method: "POST", headers: hdr(), body: JSON.stringify({ op, col, id, data }) });
    if (!r.ok) { const e = new Error("save failed"); e.code = r.status === 403 ? "invalid_argument" : "unavailable"; throw e; }
    await pull(true);
  }
  const db = {
    collection(col) {
      return {
        doc: id => db.doc(col + "/" + id),
        onSnapshot(next) { const f = () => { const docs = Object.keys(cache[col] || {}).sort().map(id => snap(col, id)); next({ docs, size: docs.length, empty: !docs.length, metadata: {} }); }; listeners.push(f); start(); if (loaded) setTimeout(f, 0); return () => {}; }
      };
    },
    doc(path) {
      const [col, id] = path.split("/");
      return {
        set: d => write("set", col, id, d), update: d => write("update", col, id, d), delete: () => write("delete", col, id),
        onSnapshot(next) { const f = () => next(snap(col, id)); listeners.push(f); start(); if (loaded) setTimeout(f, 0); return () => {}; }
      };
    }
  };

  /* ---------- who's signed in ---------- */
  function initialsAvatar(name) {
    const i = String(name || "?").split(" ").map(s => s[0]).join("").slice(0, 2).toUpperCase();
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="56" height="56"><rect width="56" height="56" rx="28" fill="#43A3A3"/><text x="28" y="35" font-family="Arial" font-size="22" font-weight="700" fill="#fff" text-anchor="middle">${i}</text></svg>`;
    return "data:image/svg+xml;base64," + btoa(svg);
  }
  const user = {
    async me() { const nana = sess.role === "nana"; return { id: sess.email.toLowerCase(), name: sess.name, email: sess.email.toLowerCase(), avatarUrl: initialsAvatar(sess.name), color: "#43A3A3", isOwner: nana, canEdit: nana }; },
    async can() { return true; }, async canEdit() { return sess.role === "nana"; }, async isOwner() { return sess.role === "nana"; },
    async search() { return []; }, async profiles(ids) { const o = {}; [].concat(ids).forEach(id => o[id] = { id, name: "", avatarUrl: initialsAvatar("?"), email: null, isMe: false, guest: false }); return o; }
  };

  /* ---------- downloads ---------- */
  const downloads = {
    async save({ filename, data }) {
      const blob = data instanceof Blob ? data : new Blob([data], { type: filename.endsWith(".xlsx") ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" : "application/octet-stream" });
      const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = filename; document.body.appendChild(a); a.click();
      setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000); return { status: "saved" };
    }
  };

  /* ---------- email + spreadsheet status ---------- */
  window.__mail = { enabled: false, async send(o) { const r = await fetch("/api/mail", { method: "POST", headers: hdr(), body: JSON.stringify(o) }); const j = await r.json().catch(() => ({})); if (!r.ok) throw new Error(j.error || "Email didn't send"); } };
  window.__sheet = { status: null, masterUrl: () => "/api/export?code=" + encodeURIComponent(sess ? sess.code : ""), async refresh() { try { const r = await fetch("/api/status", { headers: hdr() }); if (r.ok) { const j = await r.json(); window.__mail.enabled = !!j.email; window.__sheet.status = j; } } catch (e) {} } };
  window.__signOut = signOut;
  function signOut() { try { localStorage.removeItem(KEY); } catch (e) {} location.reload(); }

  /* ---------- banner ---------- */
  let bannerEl = null;
  function showBanner(t) { if (!bannerEl) { bannerEl = document.createElement("div"); bannerEl.style.cssText = "position:fixed;left:50%;top:12px;transform:translateX(-50%);z-index:70;background:#9A3F0A;color:#fff;padding:10px 16px;border-radius:999px;font:700 14px/1.3 system-ui,sans-serif;box-shadow:0 8px 24px rgba(0,0,0,.2);max-width:calc(100% - 32px)"; document.body.appendChild(bannerEl); } bannerEl.textContent = t; bannerEl.hidden = false; }
  function hideBanner() { if (bannerEl) bannerEl.hidden = true; }

  /* ---------- sign-in screen ---------- */
  function signIn() {
    return new Promise(resolve => {
      const w = document.createElement("div");
      w.innerHTML = `<div style="position:fixed;inset:0;z-index:80;background:var(--bg,#F9F9F2);display:grid;place-items:center;padding:16px;overflow:auto">
      <form id="si" style="background:var(--surface,#fff);border:1px solid var(--line,#E3E2DA);border-radius:18px;box-shadow:0 18px 40px rgba(35,31,32,.14);padding:28px;width:min(420px,100%);display:grid;gap:14px;font-family:var(--font,system-ui)">
       <img src="${window.__LOGO || ""}" alt="Visit Anaheim" style="height:34px;width:auto;justify-self:start">
       <h1 style="margin:0;font-weight:800;font-style:italic;font-size:30px;color:var(--brand,#125C60)">Event Tickets</h1>
       <p style="margin:0;color:var(--ink2,#4A4647)">Honda Center and Angels tickets for Visit Anaheim clients. Sign in to see what's open and request tickets.</p>
       <label style="display:grid;gap:5px;font-weight:700;font-size:13px">Your name<input id="si-n" required autocomplete="name" style="border:1px solid var(--line,#E3E2DA);border-radius:9px;padding:10px 12px;font:inherit;font-weight:400"></label>
       <label style="display:grid;gap:5px;font-weight:700;font-size:13px">Your work email<input id="si-e" type="email" required autocomplete="email" placeholder="name@visitanaheim.org" style="border:1px solid var(--line,#E3E2DA);border-radius:9px;padding:10px 12px;font:inherit;font-weight:400"></label>
       <label style="display:grid;gap:5px;font-weight:700;font-size:13px">Access code<input id="si-c" required autocomplete="off" style="border:1px solid var(--line,#E3E2DA);border-radius:9px;padding:10px 12px;font:inherit;font-weight:400"><span style="font-weight:400;color:var(--muted,#6B6668)">Nana shares this with leadership.</span></label>
       <div id="si-err" style="color:#A10009;font-weight:700;font-size:13px" hidden></div>
       <button style="border:0;background:#125C60;color:#F9F9F2;border-radius:999px;padding:12px 18px;font:inherit;font-weight:700;font-size:15px;cursor:pointer">Sign in</button>
      </form></div>`;
      document.body.appendChild(w);
      const f = w.querySelector("#si"), err = w.querySelector("#si-err");
      f.addEventListener("submit", async ev => {
        ev.preventDefault();
        const name = f.querySelector("#si-n").value.trim(), email = f.querySelector("#si-e").value.trim(), code = f.querySelector("#si-c").value.trim();
        if (!name || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || !code) { err.textContent = "Fill in your name, email and the access code."; err.hidden = false; return; }
        try {
          const r = await fetch("/api/login", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ code }) });
          const j = await r.json().catch(() => ({}));
          if (!r.ok) { err.textContent = j.error || "That didn't work. Try again."; err.hidden = false; return; }
          sess = { name, email, code, role: j.role };
          try { localStorage.setItem(KEY, JSON.stringify(sess)); } catch (e) {}
          w.remove(); resolve();
        } catch (e) { err.textContent = "Couldn't reach the app. Check your connection."; err.hidden = false; }
      });
      setTimeout(() => f.querySelector("#si-n").focus(), 50);
    });
  }
  const ready = (sess && sess.code && sess.email) ? Promise.resolve() : signIn();
  const caps = { db, user, downloads };
  window.claude = { use: async n => { await ready; await window.__sheet.refresh(); return caps[n] || null; } };
})();
