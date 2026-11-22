/* Korpo — Optimiseur de trajet (v2)
   Le représentant saisit ses adresses du jour avec autocomplétion en ligne (Québec / Canada) ;
   l'app trouve l'ordre le plus court en temps de conduite (max 8 arrêts pour compatibilité 100 % iOS/Android Google Maps)
   et ouvre le trajet dans Google Maps / Waze. Aucune donnée n'est envoyée à Supabase :
   tout reste dans le navigateur (localStorage). Services publics utilisés : Photon (Komoot/OSM) + Nominatim + OSRM. */
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

  // Formatage propre d'une adresse québécoise / canadienne depuis Photon (GeoJSON)
  function formatPhotonFeature(f) {
    const p = f && f.properties ? f.properties : {};
    const num = p.housenumber || '';
    const street = p.street || p.name || '';
    const line1 = [num, street].filter(Boolean).join(' ').trim();
    const city = p.city || p.town || p.village || p.municipality || p.district || '';
    const postcode = p.postcode || '';
    const state = p.state || '';
    const parts = [];
    if (line1) parts.push(line1);
    else if (p.name) parts.push(p.name);
    if (city && city.toLowerCase() !== (p.name || '').toLowerCase()) parts.push(city);
    if (state && /qu[ée]bec/i.test(state)) parts.push('QC');
    else if (state) parts.push(state);
    if (postcode) parts.push(postcode);
    const label = parts.join(', ');
    const coords = Array.isArray(f?.geometry?.coordinates) ? f.geometry.coordinates : null;
    if (!label || !coords || coords.length < 2) return null;
    return { label, lat: Number(coords[1]), lon: Number(coords[0]) };
  }

  // Recherche d'autocomplétion rapide en ligne (Photon biaisé sur le grand Montréal / Québec, avec repli Nominatim)
  const suggestCache = new Map();
  async function searchAddressSuggestions(query, signal) {
    const raw = String(query || '').trim();
    if (raw.length < 3) return [];
    const normKey = raw.toLowerCase().replace(/\s+/g, ' ');
    if (suggestCache.has(normKey)) return suggestCache.get(normKey);

    const q = /qu[ée]bec|\bqc\b|\bcanada\b/i.test(raw) ? raw : raw + ', Québec';
    try {
      // Photon (Komoot / OpenStreetMap) : rapide, tolérant aux fautes de frappe, biaisé sur Terrebonne / Montréal / Québec
      const u = 'https://photon.komoot.io/api/?q=' + encodeURIComponent(q) + '&lat=45.70&lon=-73.64&limit=6&lang=fr';
      const r = await fetch(u, { signal });
      if (r.ok) {
        const j = await r.json();
        const seen = new Set();
        const out = [];
        for (const feat of (j.features || [])) {
          const cc = String(feat?.properties?.countrycode || '').toUpperCase();
          if (cc && cc !== 'CA') continue;
          const item = formatPhotonFeature(feat);
          if (!item) continue;
          const k = item.label.toLowerCase();
          if (seen.has(k)) continue;
          seen.add(k);
          out.push(item);
          if (out.length >= 5) break;
        }
        if (out.length) {
          suggestCache.set(normKey, out);
          return out;
        }
      }
    } catch (e) {
      if (e && e.name === 'AbortError') return [];
    }

    // Repli Nominatim si Photon ne retourne rien
    try {
      const r2 = await fetch('https://nominatim.openstreetmap.org/search?format=jsonv2&limit=5&countrycodes=ca&accept-language=fr&q=' + encodeURIComponent(q), { signal });
      if (!r2.ok) return [];
      const j2 = await r2.json();
      const out2 = (j2 || []).map(x => ({
        label: String(x.display_name || '').split(',').slice(0, 4).join(',').trim(),
        lat: +x.lat,
        lon: +x.lon
      })).filter(x => x.label);
      suggestCache.set(normKey, out2);
      return out2;
    } catch (e2) {
      return [];
    }
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

  // Gestion des suggestions d'autocomplétion sur un champ texte ou sur la dernière ligne du textarea
  function bindAddressAutocomplete(inputEl, dropdownEl, opts = {}) {
    let debounceTimer = null;
    let activeAbort = null;
    let currentItems = [];
    let selectedIdx = -1;

    const getQueryText = () => {
      if (opts.multilineLastLine) {
        const lines = inputEl.value.split('\n');
        return (lines[lines.length - 1] || '').trim();
      }
      return inputEl.value.trim();
    };

    const applyChoice = (item) => {
      if (!item) return;
      // Sauvegarder directement les coordonnées dans le cache local pour accélérer le calcul ensuite
      const st = read();
      st.cache = st.cache || {};
      st.cache[item.label.toLowerCase().replace(/\s+/g, ' ').trim()] = { lat: item.lat, lon: item.lon, label: item.label };
      write(st);

      if (opts.multilineLastLine) {
        const lines = inputEl.value.split('\n');
        lines[lines.length - 1] = item.label;
        const nonEmptyCount = lines.map(s => s.trim()).filter(Boolean).length;
        // Ajouter automatiquement un saut de ligne si on n'a pas encore atteint MAX arrêts
        inputEl.value = nonEmptyCount < MAX ? lines.join('\n') + '\n' : lines.join('\n');
      } else {
        inputEl.value = item.label;
      }
      closeDropdown();
      inputEl.dispatchEvent(new Event('input'));
      inputEl.focus();
    };

    const closeDropdown = () => {
      dropdownEl.classList.add('hidden');
      dropdownEl.innerHTML = '';
      currentItems = [];
      selectedIdx = -1;
    };

    const renderDropdown = (items, isLoading = false) => {
      currentItems = items;
      selectedIdx = items.length ? 0 : -1;
      if (isLoading) {
        dropdownEl.innerHTML = '<div class="px-3.5 py-2.5 text-xs text-slate-500 flex items-center gap-2"><span class="w-2 h-2 rounded-full bg-indigo-500 animate-ping"></span><span>Recherche d\'adresse en ligne...</span></div>';
        dropdownEl.classList.remove('hidden');
        return;
      }
      if (!items.length) {
        closeDropdown();
        return;
      }
      dropdownEl.innerHTML = items.map((it, idx) =>
        '<button type="button" data-ac-idx="' + idx + '" class="w-full text-left px-3.5 py-2.5 text-xs flex items-start gap-2.5 transition cursor-pointer ' +
        (idx === selectedIdx ? 'bg-indigo-50 text-indigo-950 font-bold' : 'text-slate-700 hover:bg-slate-50 font-medium') + '">' +
        '<span class="text-sm shrink-0 mt-0.5">📍</span>' +
        '<span class="min-w-0 flex-1 break-words leading-snug">' + E(it.label) + '</span>' +
        '</button>'
      ).join('');
      dropdownEl.classList.remove('hidden');

      dropdownEl.querySelectorAll('button[data-ac-idx]').forEach(btn => {
        btn.addEventListener('mousedown', (ev) => {
          ev.preventDefault();
          const i = Number(btn.getAttribute('data-ac-idx'));
          applyChoice(currentItems[i]);
        });
      });
    };

    inputEl.addEventListener('input', () => {
      clearTimeout(debounceTimer);
      if (activeAbort) {
        try { activeAbort.abort(); } catch (e) {}
        activeAbort = null;
      }
      const q = getQueryText();
      if (q.length < 3) {
        closeDropdown();
        return;
      }
      debounceTimer = setTimeout(async () => {
        const ctrl = new AbortController();
        activeAbort = ctrl;
        renderDropdown([], true);
        const results = await searchAddressSuggestions(q, ctrl.signal);
        if (ctrl.signal.aborted) return;
        renderDropdown(results, false);
      }, 280);
    });

    inputEl.addEventListener('keydown', (e) => {
      if (dropdownEl.classList.contains('hidden') || !currentItems.length) return;
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        selectedIdx = (selectedIdx + 1) % currentItems.length;
        dropdownEl.querySelectorAll('button[data-ac-idx]').forEach((b, i) => {
          b.className = 'w-full text-left px-3.5 py-2.5 text-xs flex items-start gap-2.5 transition cursor-pointer ' +
            (i === selectedIdx ? 'bg-indigo-50 text-indigo-950 font-bold' : 'text-slate-700 hover:bg-slate-50 font-medium');
        });
      } else if (e.key === 'ArrowUp') {
        e.preventDefault();
        selectedIdx = (selectedIdx - 1 + currentItems.length) % currentItems.length;
        dropdownEl.querySelectorAll('button[data-ac-idx]').forEach((b, i) => {
          b.className = 'w-full text-left px-3.5 py-2.5 text-xs flex items-start gap-2.5 transition cursor-pointer ' +
            (i === selectedIdx ? 'bg-indigo-50 text-indigo-950 font-bold' : 'text-slate-700 hover:bg-slate-50 font-medium');
        });
      } else if (e.key === 'Enter' && selectedIdx >= 0) {
        e.preventDefault();
        applyChoice(currentItems[selectedIdx]);
      } else if (e.key === 'Escape') {
        closeDropdown();
      }
    });

    inputEl.addEventListener('blur', () => {
      setTimeout(closeDropdown, 180);
    });
  }

  async function run() {
    const st = read(), start = $('rt-start').value.trim(), ret = $('rt-ret').checked;
    const lines = $('rt-stops').value.split('\n').map(s => s.trim()).filter(Boolean);
    const out = $('rt-out'), btn = $('rt-go');
    const stat = m => { out.innerHTML = '<p class="text-sm text-slate-500">' + E(m) + '</p>'; };
    if (!start) return say('Entre ton adresse de départ', 'error');
    if (lines.length < 2) return say('Ajoute au moins 2 arrêts', 'error');
    if (lines.length > MAX) return say('Maximum ' + MAX + ' arrêts par trajet', 'error');
    btn.disabled = true;
    try {
      st.cache = st.cache || {};
      const names = [start, ...lines], pts = [], bad = [];
      for (let i = 0; i < names.length; i++) {
        stat('Recherche des adresses… (' + (i + 1) + '/' + names.length + ')');
        const p = await geocode(names[i], st.cache);
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

  function updateStopsCounter() {
    const badge = $('rt-stops-counter');
    const ta = $('rt-stops');
    if (!badge || !ta) return;
    const count = ta.value.split('\n').map(s => s.trim()).filter(Boolean).length;
    badge.textContent = count + ' / ' + MAX + ' arrêts';
    badge.className = count > MAX
      ? 'px-2 py-0.5 text-[10px] font-extrabold rounded-full bg-rose-100 text-rose-700 border border-rose-200 tabular-nums'
      : 'px-2 py-0.5 text-[10px] font-extrabold rounded-full bg-indigo-50 text-indigo-700 border border-indigo-200 tabular-nums';
  }

  function render() {
    const sec = $('tab-route');
    if (!sec || $('rt-start')) return;
    const f = 'w-full px-3.5 py-2.5 text-sm font-medium bg-white border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-600';
    sec.innerHTML =
      '<div class="max-w-3xl mx-auto space-y-5"><div class="bg-white rounded-3xl border border-slate-200/80 p-5 shadow-xs space-y-4">' +
      '<div class="flex items-center gap-3"><span class="w-10 h-10 rounded-2xl bg-indigo-600 text-white flex items-center justify-center text-lg shrink-0">🧭</span><div>' +
      '<h2 class="text-lg sm:text-xl font-extrabold text-slate-900 tracking-tight">Optimiseur de trajet</h2>' +
      '<p class="text-xs text-slate-500">Commence à taper une adresse : elle se complète automatiquement en ligne, puis l\'app calcule le trajet le plus court.</p></div></div>' +
      '<div class="relative"><label for="rt-start" class="block text-xs font-bold text-slate-700 mb-1.5">Adresse de départ</label>' +
      '<input id="rt-start" autocomplete="off" placeholder="Ex. 123 rue Principale, Mascouche" class="' + f + '" />' +
      '<div id="rt-start-ac" class="hidden absolute left-0 right-0 top-full mt-1 bg-white border border-slate-200 rounded-2xl shadow-2xl overflow-hidden divide-y divide-slate-100 z-30"></div></div>' +
      '<label class="flex items-center gap-2 text-xs font-bold text-slate-700 cursor-pointer"><input id="rt-ret" type="checkbox" class="rounded" /> Revenir au point de départ à la fin</label>' +
      '<div class="relative"><div class="flex items-center justify-between mb-1.5"><label for="rt-stops" class="block text-xs font-bold text-slate-700">Arrêts (une adresse par ligne, ' + MAX + ' maximum)</label>' +
      '<span id="rt-stops-counter" class="px-2 py-0.5 text-[10px] font-extrabold rounded-full bg-indigo-50 text-indigo-700 border border-indigo-200 tabular-nums">0 / ' + MAX + ' arrêts</span></div>' +
      '<textarea id="rt-stops" rows="6" placeholder="Tape une adresse par ligne (ex. 456 boul. des Laurentides, Laval)..." class="' + f + '"></textarea>' +
      '<div id="rt-stops-ac" class="hidden absolute left-0 right-0 top-full mt-1 bg-white border border-slate-200 rounded-2xl shadow-2xl overflow-hidden divide-y divide-slate-100 z-30"></div></div>' +
      '<div class="flex flex-wrap items-center gap-2"><button type="button" id="rt-go" class="px-4 py-2.5 text-sm font-extrabold text-white bg-indigo-600 hover:bg-indigo-700 disabled:opacity-50 rounded-xl cursor-pointer">Optimiser mon trajet</button>' +
      '<button type="button" id="rt-clear" class="px-3.5 py-2.5 text-sm font-bold text-slate-600 bg-slate-100 hover:bg-slate-200 rounded-xl cursor-pointer">Effacer les arrêts</button></div>' +
      '<p class="text-[11px] text-slate-400">Autocomplétion et calcul via OpenStreetMap (Photon / OSRM). Les adresses restent uniquement dans ton navigateur. Les heures de rendez-vous fixes ne sont pas prises en compte.</p></div>' +
      '<div id="rt-out" class="bg-white rounded-3xl border border-slate-200/80 p-5 shadow-xs empty:hidden"></div></div>';

    const st = read();
    $('rt-start').value = st.start || '';
    $('rt-ret').checked = st.ret !== false;
    $('rt-stops').value = st.stops || '';
    updateStopsCounter();

    const keep = () => {
      updateStopsCounter();
      write(Object.assign(read(), { start: $('rt-start').value, ret: $('rt-ret').checked, stops: $('rt-stops').value }));
    };
    ['rt-start', 'rt-ret', 'rt-stops'].forEach(id => $(id).addEventListener('input', keep));
    $('rt-ret').addEventListener('change', keep);
    $('rt-go').addEventListener('click', run);
    $('rt-clear').addEventListener('click', () => { $('rt-stops').value = ''; $('rt-out').innerHTML = ''; keep(); });

    bindAddressAutocomplete($('rt-start'), $('rt-start-ac'), { multilineLastLine: false });
    bindAddressAutocomplete($('rt-stops'), $('rt-stops-ac'), { multilineLastLine: true });
  }

  window.KorpoRoute = { render };
})();
