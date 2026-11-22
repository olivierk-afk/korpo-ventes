/* Korpo — Guide de démarrage (v1)
   S'affiche une seule fois par utilisateur à la première ouverture, puis reste accessible
   depuis le menu du compte (« Revoir le guide de démarrage »). Mémorisé dans le navigateur. */
(function () {
  'use strict';
  const $ = id => document.getElementById(id);
  const REP = [
    { tab: 'tab-mes-ventes', icon: '👋', t: 'Bienvenue sur Korpo Ventes', d: 'Ce guide d\'une minute te montre où déclarer tes ventes, suivre tes commissions et planifier ta journée. Tu pourras le revoir en tout temps depuis le menu de ton compte.' },
    { tab: 'tab-mes-ventes', icon: '💵', t: 'Déclarer une vente', d: 'Dans « Mes ventes », appuie sur « Déclarer une vente » (le bouton flottant sur téléphone). Tous les champs sont obligatoires, car ils servent à la comptabilité. Ta commission estimée s\'affiche avant d\'enregistrer. Tu peux aussi dicter la vente à la voix.' },
    { tab: 'tab-mes-commissions', icon: '💰', t: 'Mes commissions', d: 'Le détail de tes commissions, vente par vente, avec leur statut (payé ou non payé).' },
    { tab: 'tab-kpi', icon: '📊', t: 'Tableau KPI', d: 'Suis tes objectifs du mois. Chaque jour, note tes rendez-vous bookés et réalisés : les compteurs du jour se remettent à zéro à minuit.' },
    { tab: 'tab-classement', icon: '🏆', t: 'Classement de l\'équipe', d: 'Les volumes signés et la progression vers les objectifs de l\'équipe. Les commissions de chacun restent privées.' },
    { tab: 'tab-route', icon: '🧭', t: 'Optimiseur de trajet', d: 'Entre ton adresse de départ et jusqu\'à 8 arrêts : l\'app trouve l\'ordre le plus court et ouvre le trajet dans Waze ou Google Maps. Choisis chaque adresse dans les suggestions pour plus de précision.' },
    { tab: 'tab-mes-ventes', icon: '🛠️', t: 'Outils rapides', d: 'La bulle de chat « Korpo Top Earner » pour échanger avec l\'équipe, le téléphone rouge pour appeler le directeur sur un deal urgent, la cloche pour tes notifications et le bouton d\'actualisation si les données semblent en retard. Le menu de ton compte contient un glossaire des termes.' }
  ];
  const ADMIN = [
    REP[0],
    { tab: 'tab-admin-ventes', icon: '📋', t: 'Registre des ventes', d: 'Toutes les ventes de l\'équipe : valide et gère les dossiers.' },
    { tab: 'tab-admin-dashboard', icon: '📈', t: 'Tableau de bord', d: 'L\'analytique de la direction, la paie et les courriels.' },
    { tab: 'tab-admin-taux', icon: '⚙️', t: 'Grille des taux', d: 'Les commissions et les règles par client.' },
    { tab: 'tab-admin-equipe', icon: '🏢', t: 'Entreprises et équipe', d: 'Les objectifs, les représentants et les entreprises partenaires.' },
    REP[5],
    REP[6]
  ];
  let steps = [], i = 0;
  const who = () => (typeof currentProfile !== 'undefined' && currentProfile) || {};
  const key = () => 'korpo_guide_v1_' + (who().id || who().nom || 'anon');
  const isDone = () => { try { return localStorage.getItem(key()) === '1'; } catch (e) { return false; } };
  const markDone = () => { try { localStorage.setItem(key(), '1'); } catch (e) {} };

  function close() {
    markDone();
    document.removeEventListener('keydown', onKey);
    $('guide-card')?.remove();
    if (typeof switchTab === 'function') switchTab('tab-mes-ventes');
  }
  function onKey(e) { if (e.key === 'Escape') close(); }

  function show() {
    const s = steps[i], last = i === steps.length - 1;
    if (s.tab && typeof switchTab === 'function') switchTab(s.tab);
    let el = $('guide-card');
    if (!el) {
      el = document.createElement('div');
      el.id = 'guide-card';
      el.setAttribute('role', 'dialog');
      el.setAttribute('aria-label', 'Guide de démarrage');
      el.className = 'fixed z-[80] inset-x-3 bottom-24 lg:inset-x-auto lg:right-6 lg:bottom-6 lg:w-96 no-print';
      document.body.appendChild(el);
      document.addEventListener('keydown', onKey);
    }
    const dots = steps.map((_, k) => '<span class="w-1.5 h-1.5 rounded-full ' + (k === i ? 'bg-indigo-600' : 'bg-slate-300') + '"></span>').join('');
    const b = 'px-3.5 py-2 text-xs font-bold rounded-xl ';
    el.innerHTML = '<div class="bg-white rounded-3xl border border-slate-200 shadow-2xl p-5 space-y-3">' +
      '<div class="flex items-start gap-3"><span class="w-10 h-10 rounded-2xl bg-indigo-600 text-white flex items-center justify-center text-lg shrink-0">' + s.icon + '</span>' +
      '<div class="min-w-0"><h3 class="text-base font-extrabold text-slate-900">' + s.t + '</h3><p class="text-sm text-slate-600 mt-1 leading-relaxed">' + s.d + '</p></div></div>' +
      '<div class="flex items-center justify-between gap-2"><div class="flex gap-1" aria-hidden="true">' + dots + '</div><div class="flex items-center gap-1.5">' +
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
    steps = who().role === 'admin' ? ADMIN : REP;
    i = 0;
    show();
  }

  window.KorpoGuide = { maybeStart() { if (!isDone()) start(); }, start };
})();
