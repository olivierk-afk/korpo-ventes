/* Korpo — Guide de démarrage (v3, avec mise en évidence animée)
   S'affiche une seule fois par utilisateur à la première ouverture, puis reste accessible
   depuis le menu du compte (« Revoir le guide de démarrage »). Mémorisé dans le navigateur. */
(function () {
  'use strict';
  const $ = id => document.getElementById(id);
  const nav = t => ['.nav-btn[data-tab="' + t + '"]'];
  const W = { // étapes communes
    welcome: { tab: 'tab-mes-ventes', icon: '👋', t: 'Bienvenue sur Korpo Ventes', d: 'Ce guide rapide te montre où déclarer tes ventes, suivre tes commissions et planifier ta journée. Tu pourras le revoir en tout temps depuis le menu de ton compte.' },
    route: { tab: 'tab-route', icon: '🧭', t: 'Optimiseur de trajet', sel: nav('tab-route'), d: 'Entre ton adresse de départ et jusqu\'à 8 arrêts : l\'app trouve l\'ordre le plus court et ouvre le trajet dans Waze ou Google Maps. Choisis chaque adresse dans les suggestions pour plus de précision.' },
    chat: { icon: '🏆', t: 'Chat d\'équipe', sel: ['#btn-floating-chat-bubble', '#desktop-sidebar button[onclick="openTeamChatModal()"]'], d: 'Korpo Top Earner : le chat de l\'équipe, pour échanger et partager les bonnes nouvelles.' },
    phone: { icon: '☎️', t: 'Téléphone rouge', req: 1, sel: ['#btn-header-red-phone'], d: 'Un deal urgent à closer ? Ce bouton appelle directement le directeur pour t\'aider.' },
    bell: { icon: '🔔', t: 'Notifications', req: 1, sel: ['#btn-admin-notif'], d: 'Les mentions du chat et les ventes à confirmer apparaissent ici.' },
    refresh: { icon: '🔄', t: 'Actualiser', sel: ['#btn-refresh'], d: 'Si les données semblent en retard, appuie ici pour les recharger.' },
    account: { icon: '👤', t: 'Ton compte', sel: ['#btn-user-menu'], d: 'Change ton mot de passe, consulte le glossaire des termes, revois ce guide ou déconnecte-toi.' }
  };
  const REP = [
    W.welcome,
    { tab: 'tab-mes-ventes', icon: '💵', t: 'Déclarer une vente', sel: ['#btn-mobile-fab-sale', '#desktop-sidebar button[onclick="openNewSaleDrawer()"]'], d: 'Appuie sur « Déclarer une vente ». Tous les champs sont obligatoires, car ils servent à la comptabilité. Ta commission estimée s\'affiche avant d\'enregistrer. Tu peux aussi dicter la vente à la voix.' },
    { tab: 'tab-mes-commissions', icon: '💰', t: 'Mes commissions', sel: nav('tab-mes-commissions'), d: 'Le détail de tes commissions, vente par vente, avec leur statut (payé ou non payé).' },
    { tab: 'tab-kpi', icon: '📊', t: 'Tableau KPI', sel: nav('tab-kpi'), d: 'Suis tes objectifs du mois. Chaque jour, note tes rendez-vous bookés et réalisés : les compteurs du jour se remettent à zéro à minuit.' },
    { tab: 'tab-classement', icon: '🏆', t: 'Classement de l\'équipe', sel: nav('tab-classement'), d: 'Les volumes signés et la progression vers les objectifs de l\'équipe. Les commissions de chacun restent privées.' },
    W.route, W.chat, W.phone, W.bell, W.refresh, W.account
  ];
  const adm = (tab, icon, t, d) => ({ tab, icon, t, d, sel: nav(tab).concat(['#btn-mobile-admin-trigger']) });
  const ADMIN = [
    W.welcome,
    adm('tab-admin-ventes', '📋', 'Registre des ventes', 'Toutes les ventes de l\'équipe : valide et gère les dossiers.'),
    adm('tab-admin-dashboard', '📈', 'Tableau de bord', 'L\'analytique de la direction, la paie et les courriels.'),
    adm('tab-admin-finances', '🏦', 'Finances Korpo', 'Les encaissements, les dépenses et le profit réel.'),
    adm('tab-admin-entreprises', '🏢', 'Entreprises partenaires', 'Les contrats, la facturation et les marges. La grille des taux de commission se trouve dans la fiche de chaque entreprise.'),
    adm('tab-admin-equipe', '👥', 'Représentants et équipe', 'Les profils, les objectifs et la paie.'),
    W.route, W.chat, W.phone, W.bell, W.refresh, W.account
  ];
  let steps = [], i = 0, cur = null, placed = false, prevFocus = null;
  const who = () => (typeof currentProfile !== 'undefined' && currentProfile) || {};
  const key = () => 'korpo_guide_v2_' + (who().id || who().nom || 'anon');
  const isDone = () => { try { return localStorage.getItem(key()) === '1'; } catch (e) { return false; } };
  const markDone = () => { try { localStorage.setItem(key(), '1'); } catch (e) {} };
  const tabExists = t => !t || !!document.querySelector('.nav-btn[data-tab="' + t + '"]');
  const find = s => {
    for (const sel of s.sel || []) for (const el of document.querySelectorAll(sel)) {
      const r = el.getBoundingClientRect();
      if (r.width > 0 && r.height > 0) return el;
    }
    return null;
  };

  const CSS = '@keyframes guideIn{from{opacity:0;transform:translateY(14px) scale(.97)}to{opacity:1;transform:none}}' +
    '@keyframes guidePulse{0%{transform:scale(1);opacity:.9}100%{transform:scale(1.16);opacity:0}}' +
    '@keyframes guideBob{0%,100%{transform:translateY(0)}50%{transform:translateY(-4px)}}' +
    '#guide-overlay{position:fixed;inset:0;z-index:75;transition:background .3s}#guide-overlay.guide-dim{background:rgba(15,23,42,.6)}' +
    '#guide-spot{position:fixed;border-radius:16px;box-shadow:0 0 0 9999px rgba(15,23,42,.6);pointer-events:none;opacity:0;' +
    'transition:top .35s cubic-bezier(.22,1,.36,1),left .35s cubic-bezier(.22,1,.36,1),width .35s cubic-bezier(.22,1,.36,1),height .35s cubic-bezier(.22,1,.36,1),opacity .25s}' +
    '#guide-spot::after{content:"";position:absolute;inset:-3px;border-radius:inherit;border:2px solid #a5b4fc;animation:guidePulse 1.5s ease-out infinite}' +
    '#guide-card>div{animation:guideIn .32s cubic-bezier(.22,1,.36,1)}.guide-ico{animation:guideBob 1.4s ease-in-out 2}' +
    '@media (prefers-reduced-motion:reduce){#guide-card>div,.guide-ico,#guide-spot::after{animation:none}#guide-spot{transition:none}}';

  function place() {
    const sp = $('guide-spot'), ov = $('guide-overlay');
    if (!sp) return;
    if (!cur) { sp.style.opacity = '0'; ov.classList.add('guide-dim'); return; }
    const r = cur.getBoundingClientRect(), p = 6;
    if (!placed) { sp.style.transition = 'none'; requestAnimationFrame(() => requestAnimationFrame(() => { sp.style.transition = ''; })); placed = true; }
    Object.assign(sp.style, { opacity: '1', top: r.top - p + 'px', left: r.left - p + 'px', width: r.width + 2 * p + 'px', height: r.height + 2 * p + 'px' });
    ov.classList.remove('guide-dim');
  }
  let raf = 0;
  const onMove = () => { cancelAnimationFrame(raf); raf = requestAnimationFrame(place); };
  function onKey(e) {
    if (e.key === 'Escape') close();
    else if (e.key === 'Tab') {
      const f = [...document.querySelectorAll('#guide-card button')];
      if (!f.length) return;
      const a = document.activeElement, first = f[0], last = f[f.length - 1];
      if (e.shiftKey && (a === first || !f.includes(a))) { e.preventDefault(); last.focus(); }
      else if (!e.shiftKey && (a === last || !f.includes(a))) { e.preventDefault(); first.focus(); }
    }
    else if (e.key === 'ArrowRight') $('guide-next')?.click();
    else if (e.key === 'ArrowLeft') $('guide-prev')?.click();
  }
  function close() {
    markDone();
    document.removeEventListener('keydown', onKey);
    window.removeEventListener('resize', onMove);
    window.removeEventListener('scroll', onMove, true);
    $('guide-card')?.remove(); $('guide-overlay')?.remove(); $('guide-style')?.remove();
    cur = null; placed = false;
    try { prevFocus?.focus?.(); } catch (e) {}
    prevFocus = null;
    if (typeof switchTab === 'function') switchTab('tab-mes-ventes');
  }

  function show() {
    const s = steps[i], last = i === steps.length - 1;
    if (s.tab && typeof switchTab === 'function') switchTab(s.tab);
    cur = find(s);
    if (cur) cur.scrollIntoView({ block: 'nearest' });
    place();
    const low = cur && cur.getBoundingClientRect().top + cur.offsetHeight / 2 > innerHeight * 0.55;
    const el = $('guide-card');
    el.className = 'fixed z-[80] inset-x-3 lg:inset-x-auto lg:right-6 lg:w-96 no-print ' + (low ? 'top-20 lg:top-24' : 'bottom-24 lg:bottom-6');
    const dots = steps.map((_, k) => '<span class="h-1.5 rounded-full ' + (k === i ? 'w-4 bg-indigo-600' : 'w-1.5 bg-slate-300') + '"></span>').join('');
    const b = 'px-3.5 py-2 text-xs font-bold rounded-xl ';
    el.innerHTML = '<div class="bg-white rounded-3xl border border-slate-200 shadow-2xl p-5 space-y-3">' +
      '<div class="flex items-start gap-3"><span class="guide-ico w-10 h-10 rounded-2xl bg-indigo-600 text-white flex items-center justify-center text-lg shrink-0">' + s.icon + '</span>' +
      '<div class="min-w-0"><h3 id="guide-title" class="text-base font-extrabold text-slate-900">' + s.t + '</h3><p class="text-sm text-slate-600 mt-1 leading-relaxed">' + s.d + '</p></div></div>' +
      '<div class="flex items-center justify-between gap-2"><span class="sr-only" aria-live="polite">Étape ' + (i + 1) + ' sur ' + steps.length + '</span><div class="flex items-center gap-1" aria-hidden="true">' + dots + '</div><div class="flex items-center gap-1.5">' +
      (last ? '' : '<button type="button" id="guide-skip" class="' + b + 'text-slate-500 hover:bg-slate-100">Passer</button>') +
      (i ? '<button type="button" id="guide-prev" class="' + b + 'text-slate-700 bg-slate-100 hover:bg-slate-200">Précédent</button>' : '') +
      '<button type="button" id="guide-next" class="' + b + 'text-white bg-indigo-600 hover:bg-indigo-700">' + (last ? 'Terminer' : 'Suivant') + '</button></div></div></div>';
    $('guide-skip')?.addEventListener('click', close);
    $('guide-prev')?.addEventListener('click', () => { i--; show(); });
    $('guide-next').addEventListener('click', () => { if (last) close(); else { i++; show(); } });
    $('guide-next').focus();
  }

  function start() {
    if ($('guide-card')) return;
    steps = (who().role === 'admin' ? ADMIN : REP).filter(s => tabExists(s.tab) && (!s.req || find(s)));
    i = 0; placed = false; prevFocus = document.activeElement;
    const st = document.createElement('style'); st.id = 'guide-style'; st.textContent = CSS; document.head.appendChild(st);
    const ov = document.createElement('div'); ov.id = 'guide-overlay'; ov.className = 'guide-dim no-print';
    ov.innerHTML = '<div id="guide-spot"></div>';
    document.body.appendChild(ov);
    const card = document.createElement('div');
    card.id = 'guide-card'; card.setAttribute('role', 'dialog'); card.setAttribute('aria-modal', 'true'); card.setAttribute('aria-labelledby', 'guide-title');
    document.body.appendChild(card);
    document.addEventListener('keydown', onKey);
    window.addEventListener('resize', onMove);
    window.addEventListener('scroll', onMove, true);
    show();
  }

  window.KorpoGuide = { maybeStart() { if (!isDone()) start(); }, start };
})();
