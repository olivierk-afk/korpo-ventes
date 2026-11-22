/* Korpo — Optimiseur de trajet (v1)
   Le représentant colle ses adresses du jour ; l'app trouve l'ordre le plus court en temps de conduite
   et ouvre le trajet dans Google Maps / Waze. Aucune donnée n'est envoyée à Supabase :
   tout reste dans le navigateur (localStorage). Services publics utilisés : Nominatim + OSRM (OpenStreetMap). */
(function () {
  'use strict';
  const KEY = 'korpo_route_v1', MAX = 8;
  const $ = id => document.getElementById(id);
  const E = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const read = () => { try { return JSON.parse(localStorage.getItem(KEY) || '{}'); } catch (e) { return {}; } };
  const write = o => { try { localStorage.setItem(KEY, JSON.stringify(o)); } catch (e) {} };
  const say = (m, t) => (typeof toast === 'function' ? toast(m, t) : alert(m));
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const km = m => (m / 1000).toLocaleString('fr-CA', { maximumFractionDigits: 1 }) + ' km';
  const hm = s => { const m = Math.round(s / 60); return m >= 60 ? Math.floor(m / 60) + ' h ' + String(m % 60).padStart(2, '0') : m + ' min'; };
  let lastGeo = 0;
  let model = { start: { t: '' }, stops: [{ t: '' }, { t: '' }] };
  const BBOX = '-79.8,44.9,-57,62.6'; // Québec
  const F = 'w-full px-3.5 py-2.5 text-sm font-medium bg-white border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-600';
  const persist = () => write(Object.assign(read(), { start: model.start, ret: $('rt-ret') ? $('rt-ret').checked : true, stops: model.stops }));
  const labelOf = f => {
    const q = f.properties || {};
    const first = q.street ? (q.housenumber ? q.housenumber + ' ' : '') + q.street : q.name;
    return [first, q.city || q.locality || q.district || q.county, q.state, q.postcode].filter(Boolean).join(', ');
  };

  // Suggestions pendant la saisie (Photon / OpenStreetMap) ; le choix enregistre aussi les coordonnées
  function attach(inp, slot) {
    const box = document.createElement('ul');
    box.className = 'hidden absolute z-30 left-0 right-0 top-full mt-1 bg-white border border-slate-200 rounded-xl shadow-lg overflow-hidden';
    inp.parentNode.appendChild(box);
    let timer, seq = 0, items = [], cur = -1;
    const close = () => { box.classList.add('hidden'); box.innerHTML = ''; items = []; cur = -1; };
    const mark = () => [...box.children].forEach((li, i) => li.classList.toggle('bg-indigo-50', i === cur));
    const pick = i => { const it = items[i]; if (!it) return; inp.value = it.t; slot.t = it.t; slot.p = it.p; close(); persist(); };
    inp.addEventListener('input', () => {
      slot.t = inp.value; slot.p = null; persist();
      clearTimeout(timer);
      const q = inp.value.trim();
      if (q.length < 3) return close();
      timer = setTimeout(async () => {
        const my = ++seq;
        try {
          const r = await fetch('https://photon.komoot.io/api/?lang=fr&limit=5&lat=45.5&lon=-73.57&bbox=' + BBOX + '&q=' + encodeURIComponent(q));
          const j = await r.json();
          if (my !== seq) return;
          const all = (j.features || []).map(f => ({ t: labelOf(f), p: { lat: f.geometry.coordinates[1], lon: f.geometry.coordinates[0], label: labelOf(f) } })).filter(x => x.t);
          items = all.filter((x, i) => all.findIndex(y => y.t === x.t) === i);
          cur = -1;
          box.innerHTML = items.map((x, i) => '<li data-i="' + i + '" class="px-3.5 py-2 text-sm text-slate-700 cursor-pointer hover:bg-indigo-50">' + E(x.t) + '</li>').join('');
          box.classList.toggle('hidden', !items.length);
        } catch (e) { close(); }
      }, 350);
    });
    box.addEventListener('pointerdown', e => { const li = e.target.closest('li'); if (li) { e.preventDefault(); pick(+li.dataset.i); } });
    inp.addEventListener('keydown', e => {
      if (box.classList.contains('hidden')) return;
      if (e.key === 'ArrowDown') { cur = Math.min(cur + 1, items.length - 1); mark(); e.preventDefault(); }
      else if (e.key === 'ArrowUp') { cur = Math.max(cur - 1, 0); mark(); e.preventDefault(); }
      else if (e.key === 'Enter' && cur >= 0) { pick(cur); e.preventDefault(); }
      else if (e.key === 'Escape') close();
    });
    inp.addEventListener('blur', () => setTimeout(close, 150));
  }

  async function geocode(addr, cache) {
    const k = addr.toLowerCase().replace(/\s+/g, ' ').trim();
    if (cache[k]) return cache[k];
    const w = 1100 - (Date.now() - lastGeo); // Nominatim : 1 requête / seconde max
    if (w > 0) await sleep(w);
    lastGeo = Date.now();
    const q = /qu[ée]bec|\bqc\b/i.test(addr) ? addr : addr + ', Québec';
    const r = await fetch('https://nominatim.openstreetmap.org/search?format=jsonv2&limit=1&countrycodes=ca&accept-language=fr&q=' + encodeURIComponent(q));
    if (!r.ok) throw new Error('geocode');
    const j = await r.json();
    if (!j.length) return null;
    return (cache[k] = { lat: +j[0].lat, lon: +j[0].lon, label: j[0].display_name });
  }

  async function matrix(pts) {
    try {
      const c = new AbortController(), t = setTimeout(() => c.abort(), 12000);
      const u = 'https://router.project-osrm.org/table/v1/driving/' + pts.map(p => p.lon + ',' + p.lat).join(';') + '?annotations=duration,distance';
      const r = await fetch(u, { signal: c.signal });
      clearTimeout(t);
      const j = await r.json();
      if (j.code === 'Ok' && j.durations && j.distances && !j.durations.flat().includes(null) && !j.distances.flat().includes(null)) {
        return { T: j.durations, D: j.distances, approx: false };
      }
    } catch (e) {}
    // Repli : distance à vol d'oiseau x 1,35, à 55 km/h
    const f = x => x * Math.PI / 180;
    const hav = (a, b) => {
      const dp = f(b.lat - a.lat), dl = f(b.lon - a.lon);
      const h = Math.sin(dp / 2) ** 2 + Math.cos(f(a.lat)) * Math.cos(f(b.lat)) * Math.sin(dl / 2) ** 2;
      return 2 * 6371e3 * Math.asin(Math.sqrt(h)) * 1.35;
    };
    const D = pts.map(a => pts.map(b => hav(a, b)));
    return { D, T: D.map(r => r.map(d => d / 15.3)), approx: true };
  }

  // Solution exacte (Held-Karp) : nœud 0 = départ, 1..n = arrêts
  function solve(T, n, ret) {
    const F = 1 << n;
    const dp = Array.from({ length: F }, () => new Array(n + 1).fill(Infinity));
    const pa = Array.from({ length: F }, () => new Array(n + 1).fill(-1));
    for (let j = 1; j <= n; j++) dp[1 << (j - 1)][j] = T[0][j];
    for (let m = 1; m < F; m++) {
      for (let j = 1; j <= n; j++) {
        if (!(m & (1 << (j - 1))) || dp[m][j] === Infinity) continue;
        for (let k = 1; k <= n; k++) {
          if (m & (1 << (k - 1))) continue;
          const nm = m | (1 << (k - 1)), c = dp[m][j] + T[j][k];
          if (c < dp[nm][k]) { dp[nm][k] = c; pa[nm][k] = j; }
        }
      }
    }
    let best = Infinity, bj = 1;
    for (let j = 1; j <= n; j++) {
      const c = dp[F - 1][j] + (ret ? T[j][0] : 0);
      if (c < best) { best = c; bj = j; }
    }
    const o = [];
    let m = F - 1, j = bj;
    while (j > 0) { o.unshift(j); const p = pa[m][j]; m &= ~(1 << (j - 1)); j = p; }
    return o;
  }

  function total(M, order, ret) {
    let t = 0, d = 0, p = 0;
    for (const j of order) { t += M.T[p][j]; d += M.D[p][j]; p = j; }
    if (ret) { t += M.T[p][0]; d += M.D[p][0]; }
    return { t, d };
  }

  const ll = p => p.lat + ',' + p.lon;
  const mapsUrl = (pts, order, ret) => {
    const seq = order.map(i => pts[i]);
    const dest = ret ? pts[0] : seq[seq.length - 1];
    const via = ret ? seq : seq.slice(0, -1);
    return 'https://www.google.com/maps/dir/?api=1&travelmode=driving&origin=' + encodeURIComponent(ll(pts[0])) +
      '&destination=' + encodeURIComponent(ll(dest)) + (via.length ? '&waypoints=' + via.map(p => encodeURIComponent(ll(p))).join('%7C') : '');
  };

  async function run() {
    const st = read(), ret = $('rt-ret').checked, S = model.stops.filter(x => x.t.trim()), start = model.start.t.trim(), lines = S.map(x => x.t.trim());
    const out = $('rt-out'), btn = $('rt-go');
    const stat = m => { out.innerHTML = '<p class="text-sm text-slate-500">' + E(m) + '</p>'; };
    if (!start) return say('Entre ton adresse de départ', 'error');
    if (lines.length < 2) return say('Ajoute au moins 2 arrêts', 'error');
    if (lines.length > MAX) return say('Maximum ' + MAX + ' arrêts par trajet', 'error');
    btn.disabled = true;
    try {
      st.cache = st.cache || {};
      const all = [model.start, ...S], names = [start, ...lines], pts = [], bad = [];
      for (let i = 0; i < names.length; i++) {
        stat('Recherche des adresses… (' + (i + 1) + '/' + names.length + ')');
        const p = all[i].p || await geocode(names[i], st.cache);
        if (p) pts.push(p); else bad.push(names[i]);
      }
      write(Object.assign(read(), { cache: st.cache }));
      if (bad.length) {
        out.innerHTML = '<div class="p-3.5 rounded-2xl bg-rose-50 border border-rose-200 text-sm text-rose-800"><strong>Adresses introuvables :</strong><ul class="list-disc pl-5 mt-1">' +
          bad.map(b => '<li>' + E(b) + '</li>').join('') + '</ul><p class="mt-1.5 text-xs">Ajoute le numéro civique, la ville ou le code postal, puis relance.</p></div>';
        return;
      }
      stat('Calcul du meilleur ordre…');
      const M = await matrix(pts), n = lines.length;
      const order = solve(M.T, n, ret), orig = lines.map((_, i) => i + 1);
      const a = total(M, order, ret), b = total(M, orig, ret);
      const gT = b.t - a.t, gD = b.d - a.d, gain = gT > 60 || gD > 1000;
      let p = 0;
      const rows = order.map((j, i) => {
        const row = '<li class="flex items-start gap-3 py-3"><span class="w-7 h-7 rounded-full bg-indigo-600 text-white text-xs font-extrabold flex items-center justify-center shrink-0">' + (i + 1) + '</span>' +
          '<div class="min-w-0 flex-1"><div class="text-sm font-bold text-slate-900 break-words">' + E(lines[j - 1]) + '</div>' +
          '<div class="text-[11px] text-slate-400 truncate">Repéré : ' + E(pts[j].label) + '</div>' +
          '<div class="text-xs text-slate-600 mt-0.5 tabular-nums">' + km(M.D[p][j]) + ' • ' + hm(M.T[p][j]) + ' depuis ' + (i ? 'l\'arrêt précédent' : 'le départ') + '</div></div>' +
          '<a href="https://waze.com/ul?ll=' + encodeURIComponent(ll(pts[j])) + '&navigate=yes" target="_blank" rel="noopener" class="shrink-0 px-2.5 py-1.5 text-xs font-bold text-indigo-700 bg-indigo-50 hover:bg-indigo-100 border border-indigo-200 rounded-xl">Waze</a></li>';
        p = j;
        return row;
      }).join('');
      out.innerHTML =
        '<div class="rounded-2xl p-4 border ' + (gain ? 'bg-emerald-50 border-emerald-200' : 'bg-slate-50 border-slate-200') + '">' +
        '<div class="text-sm font-extrabold ' + (gain ? 'text-emerald-800' : 'text-slate-800') + '">' +
        (gain ? 'Économie : ' + km(Math.max(gD, 0)) + ' et ' + hm(Math.max(gT, 0)) : 'Ton ordre était déjà le meilleur') + '</div>' +
        '<div class="text-xs text-slate-600 mt-1 tabular-nums">Trajet optimisé : ' + km(a.d) + ' • ' + hm(a.t) + ' &nbsp;|&nbsp; Ordre saisi : ' + km(b.d) + ' • ' + hm(b.t) + '</div>' +
        (M.approx ? '<div class="text-[11px] text-amber-700 mt-1.5">Service de routes indisponible : estimation à vol d\'oiseau, moins précise.</div>' : '') + '</div>' +
        '<div class="mt-3 text-xs text-slate-500">Départ : <strong class="text-slate-700">' + E(start) + '</strong></div>' +
        '<ol class="divide-y divide-slate-100">' + rows + '</ol>' +
        (ret ? '<div class="text-xs text-slate-500">Retour au départ : ' + km(M.D[order[n - 1]][0]) + ' • ' + hm(M.T[order[n - 1]][0]) + '</div>' : '') +
        '<div class="flex flex-wrap gap-2 pt-3"><a href="' + mapsUrl(pts, order, ret) + '" target="_blank" rel="noopener" class="px-4 py-2.5 text-sm font-extrabold text-white bg-indigo-600 hover:bg-indigo-700 rounded-xl">Ouvrir tout le trajet dans Google Maps</a>' +
        '<button type="button" id="rt-copy" class="px-4 py-2.5 text-sm font-bold text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-xl">Copier la liste</button></div>';
      $('rt-copy').onclick = () => {
        const txt = order.map((j, i) => (i + 1) + '. ' + lines[j - 1]).join('\n');
        navigator.clipboard?.writeText(txt).then(() => say('Liste copiée'), () => say('Copie impossible', 'error'));
      };
    } catch (e) {
      out.innerHTML = '<div class="p-3.5 rounded-2xl bg-rose-50 border border-rose-200 text-sm text-rose-800">Impossible de joindre le service de cartes. Vérifie ta connexion et réessaie dans un instant.</div>';
    } finally {
      btn.disabled = false;
    }
  }

  function draw(focusLast) {
    const list = $('rt-list');
    list.innerHTML = '';
    model.stops.forEach((st, i) => {
      const row = document.createElement('div');
      row.className = 'flex items-center gap-2';
      row.innerHTML = '<span class="w-7 h-7 rounded-full bg-slate-200 text-slate-700 text-xs font-extrabold flex items-center justify-center shrink-0">' + (i + 1) + '</span>' +
        '<div class="relative flex-1 min-w-0"><input autocomplete="off" placeholder="Adresse de l\'arrêt ' + (i + 1) + '" class="' + F + '" /></div>' +
        '<button type="button" aria-label="Retirer cet arrêt" class="w-9 h-9 shrink-0 text-slate-400 hover:text-rose-600 text-lg rounded-xl hover:bg-slate-100">×</button>';
      const inp = row.querySelector('input');
      inp.value = st.t;
      attach(inp, st);
      row.querySelector('button').onclick = () => { model.stops.splice(i, 1); while (model.stops.length < 2) model.stops.push({ t: '' }); persist(); draw(); };
      list.appendChild(row);
    });
    const add = $('rt-add');
    add.disabled = model.stops.length >= MAX;
    add.textContent = model.stops.length >= MAX ? 'Maximum de ' + MAX + ' arrêts atteint' : '+ Ajouter un arrêt';
    if (focusLast) list.lastElementChild.querySelector('input').focus();
  }

  function render() {
    const sec = $('tab-route');
    if (!sec || $('rt-start')) return;
    sec.innerHTML =
      '<div class="max-w-3xl mx-auto space-y-5"><div class="bg-white rounded-3xl border border-slate-200/80 p-5 shadow-xs space-y-4">' +
      '<div class="flex items-center gap-3"><span class="w-10 h-10 rounded-2xl bg-indigo-600 text-white flex items-center justify-center text-lg shrink-0">🧭</span><div>' +
      '<h2 class="text-lg sm:text-xl font-extrabold text-slate-900 tracking-tight">Optimiseur de trajet</h2>' +
      '<p class="text-xs text-slate-500">Entre les adresses de ta journée : l\'ordre le plus court en temps de conduite t\'est proposé.</p></div></div>' +
      '<div><label for="rt-start" class="block text-xs font-bold text-slate-700 mb-1.5">Adresse de départ</label><div class="relative"><input id="rt-start" autocomplete="off" placeholder="Commence à écrire une adresse…" class="' + F + '" /></div></div>' +
      '<label class="flex items-center gap-2 text-xs font-bold text-slate-700"><input id="rt-ret" type="checkbox" class="rounded" /> Revenir au point de départ à la fin</label>' +
      '<div><div class="text-xs font-bold text-slate-700 mb-1.5">Arrêts de la journée (' + MAX + ' maximum)</div><div id="rt-list" class="space-y-2"></div>' +
      '<button type="button" id="rt-add" class="mt-2 px-3.5 py-2 text-xs font-bold text-indigo-700 bg-indigo-50 hover:bg-indigo-100 disabled:opacity-50 border border-indigo-200 rounded-xl"></button></div>' +
      '<div class="flex flex-wrap items-center gap-2"><button type="button" id="rt-go" class="px-4 py-2.5 text-sm font-extrabold text-white bg-indigo-600 hover:bg-indigo-700 disabled:opacity-50 rounded-xl">Optimiser mon trajet</button>' +
      '<button type="button" id="rt-clear" class="px-3.5 py-2.5 text-sm font-bold text-slate-600 bg-slate-100 hover:bg-slate-200 rounded-xl">Effacer les arrêts</button></div>' +
      '<p class="text-[11px] text-slate-400">Pendant que tu écris, le texte est envoyé à des services de cartes publics (OpenStreetMap) pour suggérer des adresses et calculer le trajet. Rien n\'est enregistré dans le portail. Les heures de rendez-vous ne sont pas prises en compte : garde tes rendez-vous fixes dans le bon ordre.</p></div>' +
      '<div id="rt-out" class="bg-white rounded-3xl border border-slate-200/80 p-5 shadow-xs empty:hidden"></div></div>';
    const st = read();
    const slot = v => (typeof v === 'string' ? { t: v } : v && typeof v === 'object' ? { t: v.t || '', p: v.p || null } : { t: '' });
    const old = typeof st.stops === 'string' ? st.stops.split('\n').filter(Boolean) : st.stops;
    model.start = slot(st.start);
    model.stops = (Array.isArray(old) ? old.slice(0, MAX).map(slot) : []);
    while (model.stops.length < 2) model.stops.push({ t: '' });
    $('rt-start').value = model.start.t;
    $('rt-ret').checked = st.ret !== false;
    attach($('rt-start'), model.start);
    $('rt-ret').addEventListener('change', persist);
    $('rt-add').addEventListener('click', () => { if (model.stops.length < MAX) { model.stops.push({ t: '' }); persist(); draw(true); } });
    $('rt-go').addEventListener('click', run);
    $('rt-clear').addEventListener('click', () => { model.stops = [{ t: '' }, { t: '' }]; $('rt-out').innerHTML = ''; persist(); draw(); });
    draw();
  }

  window.KorpoRoute = { render };
})();
