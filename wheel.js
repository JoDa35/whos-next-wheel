(function () {
  const COLORS = ["#1f2a44","#8b1a1a","#c9b48a","#4a5a7a","#b5651d","#6b7f5a"];
  const DEFAULT = { names: [], inactive: [], removeDrawn: true, sidebar: true };
  let state = structuredClone(DEFAULT);
  let inOffice = false;
  const STORE = "zufallsrad" + location.pathname;
  let angle = 0, spinning = false;

  const $ = id => document.getElementById(id);
  const canvas = $("wheel"), ctx = canvas.getContext("2d");

  // Names live in the add-in's document settings (saved inside the .pptx), never in this hosted code.
  function load() {
    let s = null;
    if (inOffice) s = Office.context.document.settings.get("zufallsrad");
    else { try { s = JSON.parse(localStorage.getItem(STORE)); } catch (e) {} }
    if (s) state = Object.assign(structuredClone(DEFAULT), s);
  }
  function save() {
    if (inOffice) {
      Office.context.document.settings.set("zufallsrad", state);
      Office.context.document.settings.saveAsync(() => {});
    } else { try { localStorage.setItem(STORE, JSON.stringify(state)); } catch (e) {} }
  }

  const active = () => state.names.filter(n => !state.inactive.includes(n));
  function rand() {
    try { const u = new Uint32Array(1); crypto.getRandomValues(u); return u[0] / 4294967296; } catch (e) { return Math.random(); }
  }

  function size() {
    const wrap = $("wheelWrap");
    const d = Math.max(80, Math.min(wrap.clientWidth - 60, wrap.clientHeight - 8));
    const dpr = window.devicePixelRatio || 1;
    canvas.width = canvas.height = Math.round(d * dpr);
    canvas.style.width = canvas.style.height = d + "px";
    document.documentElement.style.setProperty("--r", d / 2 + "px");
    draw();
  }

  function draw() {
    const names = active(), n = names.length, R = canvas.width / 2;
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.save(); ctx.translate(R, R);
    if (!n) {
      ctx.fillStyle = "#555"; ctx.textAlign = "center"; ctx.textBaseline = "middle";
      ctx.font = `${R / 10}px Garamond, Georgia, serif`;
      ctx.fillText(state.names.length ? "No names left" : "Add names in the list", 0, 0);
      ctx.restore(); return;
    }
    const seg = 2 * Math.PI / n;
    const fs = Math.max(R / 22, Math.min(R / 11, (2 * R) / n));
    for (let i = 0; i < n; i++) {
      const a0 = angle + i * seg;
      ctx.beginPath(); ctx.moveTo(0, 0); ctx.arc(0, 0, R * .985, a0, a0 + seg); ctx.closePath();
      let col = COLORS[i % COLORS.length];
      if (n % COLORS.length === 1 && i === n - 1) col = COLORS[2];
      ctx.fillStyle = col; ctx.fill();
      ctx.strokeStyle = "#f4ecd8"; ctx.lineWidth = R / 150; ctx.stroke();
      ctx.save(); ctx.rotate(a0 + seg / 2);
      ctx.textAlign = "right"; ctx.textBaseline = "middle";
      ctx.fillStyle = col === "#c9b48a" ? "#1f2a44" : "#fff";
      ctx.font = `bold ${fs}px Garamond, Georgia, serif`;
      ctx.fillText(names[i], R * .94, 0);
      ctx.restore();
    }
    ctx.beginPath(); ctx.arc(0, 0, R * .11, 0, 2 * Math.PI); ctx.fillStyle = "#f4ecd8"; ctx.fill();
    ctx.strokeStyle = "#8b6f47"; ctx.lineWidth = R / 90; ctx.stroke();
    ctx.restore();
  }

  // Pointer sits at angle 0 (right side).
  function winnerIndex(n) {
    const seg = 2 * Math.PI / n;
    const a = ((-angle) % (2 * Math.PI) + 2 * Math.PI) % (2 * Math.PI);
    return Math.floor(a / seg) % n;
  }

  // Pages can pass SHA-256 hashes of names that stay on the wheel but are never picked
  // (hashes, so the names never appear in this public code).
  const EXCLUDED_HASHES = window.WHEEL_EXCLUDED || [];
  let excluded = new Set();
  async function updateExcluded() {
    if (!EXCLUDED_HASHES.length) return;
    const hits = new Set();
    for (const n of state.names) {
      const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode("wnw:" + n.trim().toLowerCase()));
      const hex = [...new Uint8Array(buf)].map(b => b.toString(16).padStart(2, "0")).join("");
      if (EXCLUDED_HASHES.includes(hex)) hits.add(n);
    }
    excluded = hits;
  }

  function spin() {
    const names = active();
    const eligible = names.filter(n => !excluded.has(n));
    if (spinning || !eligible.length) return;
    spinning = true; $("spin").disabled = true; $("result").innerHTML = "&nbsp;";
    // Choose the winner first, then spin so the pointer lands inside its segment.
    const pick = names.indexOf(eligible[Math.floor(rand() * eligible.length)]);
    const seg = 2 * Math.PI / names.length, TAU = 2 * Math.PI;
    const landing = pick * seg + (0.15 + rand() * 0.7) * seg;
    const start = angle;
    const total = (5 + Math.floor(rand() * 3)) * TAU + ((-landing - start) % TAU + TAU) % TAU;
    const dur = 4200 + rand() * 1200, t0 = performance.now();
    (function step(t) {
      const p = Math.min(1, (t - t0) / dur);
      angle = start + total * (1 - Math.pow(1 - p, 4)); draw();
      if (p < 1) return requestAnimationFrame(step);
      angle %= TAU;
      const name = names[winnerIndex(names.length)];
      $("result").textContent = name;
      spinning = false; $("spin").disabled = false;
      if (state.removeDrawn) setTimeout(() => {
        if (spinning || state.inactive.includes(name)) return;
        state.inactive.push(name); save(); render();
      }, 1400);
    })(t0);
  }

  function render() {
    $("side").classList.toggle("hidden", !state.sidebar);
    $("removeDrawn").checked = state.removeDrawn;
    $("count").textContent = `${active().length}/${state.names.length}`;
    const list = $("list"); list.innerHTML = "";
    state.names.forEach(n => {
      const l = document.createElement("label"), cb = document.createElement("input");
      cb.type = "checkbox"; cb.checked = !state.inactive.includes(n);
      cb.onchange = () => {
        if (spinning) { cb.checked = !cb.checked; return; }
        state.inactive = cb.checked ? state.inactive.filter(x => x !== n) : [...state.inactive, n];
        save(); render();
      };
      l.classList.toggle("out", !cb.checked);
      l.append(cb, document.createTextNode(n)); list.append(l);
    });
    if (!state.names.length) openEdit();
    size();
  }
  function openEdit() { $("names").value = state.names.join("\n"); $("editBox").classList.add("open"); }

  $("spin").onclick = spin;
  $("toggle").onclick = () => { state.sidebar = !state.sidebar; save(); render(); };
  $("removeDrawn").onchange = e => { state.removeDrawn = e.target.checked; save(); };
  $("reset").onclick = () => { if (spinning) return; state.inactive = []; $("result").innerHTML = "&nbsp;"; save(); render(); };
  $("editBtn").onclick = openEdit;
  $("cancel").onclick = () => $("editBox").classList.remove("open");
  $("save").onclick = () => {
    const names = [...new Set($("names").value.split("\n").map(s => s.trim()).filter(Boolean))];
    state.names = names; state.inactive = state.inactive.filter(n => names.includes(n));
    $("editBox").classList.remove("open"); save(); render(); updateExcluded();
  };
  document.addEventListener("keydown", e => {
    if (e.target.tagName === "TEXTAREA") return;
    if (e.code === "Space" || e.code === "Enter") { e.preventDefault(); spin(); }
  });
  window.addEventListener("resize", size);

  function start() { load(); render(); updateExcluded(); }
  if (window.Office && Office.onReady) Office.onReady(info => { inOffice = !!info.host; start(); });
  else start();
})();
