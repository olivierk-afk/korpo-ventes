// =========================================================================
    // CONFIGURATION SUPABASE (la clé « anon » est publique par conception :
    // la sécurité repose sur les politiques RLS de la base, voir les notes de livraison)
    // =========================================================================
    const SUPABASE_URL = "https://fdxqmmudhhofvtuxntqu.supabase.co";
    const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZkeHFtbXVkaGhvZnZ0dXhudHF1Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3OTQ5NDc0NTQsImV4cCI6MjExMDUyMzQ1NH0.2ogPrW1B4Zc3KECTtOO_Rj3UYF5_2TZX2oeL3CnJWBE";

    // =========================================================================
    // UTILITAIRES : échappement HTML, stockage local sûr, dates locales, CSV, pagination
    // =========================================================================
    let _xlsxLoadPromise = null;
    function ensureXlsxLoaded() {
      if (window.XLSX) return Promise.resolve(window.XLSX);
      if (_xlsxLoadPromise) return _xlsxLoadPromise;
      _xlsxLoadPromise = new Promise((resolve, reject) => {
        const s = document.createElement('script');
        s.src = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';
        s.async = true;
        s.onload = () => resolve(window.XLSX);
        s.onerror = () => { _xlsxLoadPromise = null; reject(new Error('Impossible de charger le lecteur Excel')); };
        document.head.appendChild(s);
      });
      return _xlsxLoadPromise;
    }

    function esc(s) {
      return String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
    }
    function lsSet(key, value) {
      try { localStorage.setItem(key, value); }
      catch (e) { console.warn('localStorage indisponible ou plein :', key, e && e.message); }
      if (KV_STORES[key]) kvSchedule(key, value);
    }
    function lsGetJson(key, fallbackJson) {
      try { return JSON.parse(localStorage.getItem(key) || fallbackJson); }
      catch (e) { return JSON.parse(fallbackJson); }
    }
    // =========================================================================
    // SYNCHRONISATION SUPABASE DES DONNÉES MÉTIER (audit, corbeille, taux, objectifs…)
    // Le localStorage reste un cache : chaque écriture est répliquée dans la table korpo_store,
    // et la base est relue à la connexion. Voir migration_korpo_store.sql pour la table et le RLS.
    // =========================================================================
    const KV_TABLE = 'korpo_store';
    const KV_STORES = {
      'korpo_sale_audit_v1':            { store: 'sale_audit',           shared: true,  type: 'map' },
      'korpo_sale_trash_v1':            { store: 'sale_trash',           shared: true,  type: 'trash' },
      'korpo_sale_bonus_overrides_v1':  { store: 'sale_bonus_overrides', shared: true,  type: 'map' },
      'korpo_carried_over_sales_v1':    { store: 'carried_over_sales',   shared: true,  type: 'map' },
      'korpo_rep_targets_v1':           { store: 'rep_targets',          shared: true,  type: 'map' },
      'korpo_company_targets_meta_v1':  { store: 'company_targets_meta', shared: true,  type: 'map' },
      'korpo_client_targets_v1':        { store: 'client_targets',       shared: true,  type: 'array' },
      'korpo_kpi_condensed_v1':         { store: 'kpi_condensed',        shared: true,  type: 'map' },
      'korpo_kpi_monthly_rdv_v1':       { store: 'kpi_monthly_rdv',      shared: true,  type: 'map' },
      'korpo_daily_rdv_v1':             { store: 'daily_rdv_logs',       shared: true,  type: 'map' },
      'korpo_rep_mileage_v1':           { store: 'rep_mileage_logs',     shared: true,  type: 'map' },
      'korpo_kpi_team_snapshot_v1':     { store: 'kpi_team_snapshot',    shared: true,  type: 'map' },
      'korpo_wave_company_seq_v1':      { store: 'wave_company_seq',     shared: true,  type: 'map' },
      'korpo_fin_expenses_v1':          { store: 'fin_expenses',         shared: true,  type: 'map' },
      'korpo_fin_partner_paid_v1':      { store: 'fin_partner_paid',     shared: true,  type: 'map' },
      'korpo_archived_reps_v1':         { store: 'archived_reps',        shared: true,  type: 'map' },
      'korpo_archived_entreprises_v1':  { store: 'archived_entreprises', shared: true,  type: 'map' },
      'korpo_personal_monthly_targets_v1': { store: 'personal_monthly_targets', shared: false, type: 'map' },
      'korpo_personal_comm_targets_v1':    { store: 'personal_comm_targets',    shared: false, type: 'map' },
      'korpo_personal_bonus_pct_v1':       { store: 'personal_bonus_pct',       shared: false, type: 'map' },
      'korpo_read_mention_ids_v1':         { store: 'read_mention_ids',         shared: false, type: 'array' },
      'korpo_chat_last_seen_v1':         { store: 'chat_last_seen',           shared: false, type: 'map' }
    };
    // Magasins qu'un représentant (non admin) peut écrire : les autres sont réservés à l'admin (voir le RLS)
    const KV_REP_SHARED = new Set(['sale_audit', 'sale_bonus_overrides', 'sale_trash', 'carried_over_sales', 'daily_rdv_logs', 'rep_mileage_logs']);
    const kvState = { ready: false, uid: null, last: {}, timers: {}, pending: {}, dirty: new Set(), lastHydrate: 0 };
    const kvOwner = (cfg) => (cfg.shared ? 'shared' : kvState.uid);
    const kvAllowed = (cfg) => !cfg.shared || currentProfile?.role === 'admin' || KV_REP_SHARED.has(cfg.store);

    function kvToEntries(cfg, raw) {
      let v; try { v = JSON.parse(raw); } catch (e) { return null; }
      const m = new Map();
      if (cfg.type === 'map') Object.keys(v && typeof v === 'object' ? v : {}).forEach(k => m.set(k, v[k]));
      else if (cfg.type === 'trash') (Array.isArray(v) ? v : []).forEach((it, i) => m.set(String(it?.sale?.id ?? ('i' + i)), it));
      else m.set('_', Array.isArray(v) ? v : []);
      return m;
    }
    function kvFromEntries(cfg, m) {
      if (cfg.type === 'map') { const o = {}; m.forEach((v, k) => { o[k] = v; }); return o; }
      if (cfg.type === 'trash') return Array.from(m.values()).sort((a, b) => String(b?.deletedAt || '').localeCompare(String(a?.deletedAt || '')));
      return m.get('_') ?? [];
    }
    function kvSchedule(lsKey, raw) {
      if (!kvState.ready) return;
      kvState.pending[lsKey] = raw;
      clearTimeout(kvState.timers[lsKey]);
      kvState.timers[lsKey] = setTimeout(() => { const r = kvState.pending[lsKey]; delete kvState.pending[lsKey]; kvPush(lsKey, r); }, 700);
    }
    async function kvFlush() {
      for (const k of Object.keys(kvState.pending)) {
        clearTimeout(kvState.timers[k]);
        const r = kvState.pending[k]; delete kvState.pending[k];
        await kvPush(k, r);
      }
    }
    async function kvPush(lsKey, raw) {
      const cfg = KV_STORES[lsKey];
      if (!cfg || !kvState.ready || !kvAllowed(cfg)) return;
      const next = kvToEntries(cfg, raw); if (!next) return;
      const owner = kvOwner(cfg), prev = kvState.last[lsKey] || new Map(), up = [], del = [];
      next.forEach((val, k) => { if (prev.get(k) !== JSON.stringify(val)) up.push({ store: cfg.store, item_key: k, owner, value: val, updated_at: new Date().toISOString() }); });
      prev.forEach((_, k) => { if (!next.has(k)) del.push(k); }); // on ne supprime que ce que cet appareil connaissait
      try {
        for (let i = 0; i < up.length; i += 200) {
          const { error } = await sb.from(KV_TABLE).upsert(up.slice(i, i + 200), { onConflict: 'store,item_key,owner' });
          if (error) throw error;
        }
        for (let i = 0; i < del.length; i += 200) {
          const { error } = await sb.from(KV_TABLE).delete().eq('store', cfg.store).eq('owner', owner).in('item_key', del.slice(i, i + 200));
          if (error) throw error;
        }
        const nl = new Map(prev);
        up.forEach(r => nl.set(r.item_key, JSON.stringify(r.value))); del.forEach(k => nl.delete(k));
        kvState.last[lsKey] = nl; kvState.dirty.delete(lsKey);
      } catch (e) { kvState.dirty.add(lsKey); console.warn('Synchronisation Supabase impossible :', cfg.store, e && e.message); }
    }
    function kvApplyToVariables() {
      condensedBaselineStore = lsGetJson('korpo_kpi_condensed_v1', '{}');
      monthlyRdvOverrides = lsGetJson('korpo_kpi_monthly_rdv_v1', '{}');
      dailyRdvLogsStore = lsGetJson('korpo_daily_rdv_v1', '{}');
      teamKpiSnapshotStore = lsGetJson('korpo_kpi_team_snapshot_v1', '{}');
      saleAuditStore = lsGetJson('korpo_sale_audit_v1', '{}');
      saleTrashStore = lsGetJson('korpo_sale_trash_v1', '[]');
      saleBonusOverridesStore = lsGetJson('korpo_sale_bonus_overrides_v1', '{}');
      carriedOverSalesStore = lsGetJson('korpo_carried_over_sales_v1', '{}');
      clientTargetsStore = lsGetJson('korpo_client_targets_v1', '[]');
      companyTargetsMetaStore = lsGetJson('korpo_company_targets_meta_v1', '{}');
      repTargetsOverrides = lsGetJson('korpo_rep_targets_v1', '{}');
      archivedRepsStore = lsGetJson('korpo_archived_reps_v1', '{}');
      archivedEntreprisesStore = lsGetJson('korpo_archived_entreprises_v1', '{}');
      readMentionIds = new Set(lsGetJson('korpo_read_mention_ids_v1', '[]'));
      lastSeenChatTimestamp = readSavedChatLastSeenTimestamp();
      personalMonthlyTargets = lsGetJson('korpo_personal_monthly_targets_v1', '{}');
      personalCommTargets = lsGetJson('korpo_personal_comm_targets_v1', '{}');
      personalBonusPctTargets = lsGetJson('korpo_personal_bonus_pct_v1', '{}');
      window.KorpoRoute?.renderMileageLog?.();
    }
    // Relit la base. Au premier passage sur un appareil, les données locales existantes sont ajoutées à la base
    // (la base l'emporte en cas de conflit) ; ensuite la base fait foi. Retourne true si les données locales ont changé.
    async function hydrateKv() {
      if (!currentUser || !currentProfile) return false;
      kvState.uid = currentUser.id;
      await kvFlush();
      let rows;
      try { rows = await fetchAllRows(() => sb.from(KV_TABLE).select('store,item_key,owner,value')); }
      catch (e) { console.warn('Table korpo_store indisponible (migration SQL non exécutée ?) :', e && e.message); return false; }
      kvState.ready = false;
      const flag = `korpo_kv_migrated_${kvState.uid}`;
      const migrated = localStorage.getItem(flag) === '1';
      let changed = false;
      for (const [lsKey, cfg] of Object.entries(KV_STORES)) {
        if (!kvAllowed(cfg)) continue;
        const owner = kvOwner(cfg), rawLocal = localStorage.getItem(lsKey);
        if (kvState.dirty.has(lsKey)) continue; // écriture locale pas encore envoyée : on ne l'écrase pas
        const db = new Map();
        rows.forEach(r => { if (r.store === cfg.store && r.owner === owner) db.set(r.item_key, r.value); });
        const local = rawLocal ? kvToEntries(cfg, rawLocal) : new Map();
        const merged = new Map(db);
        if (!migrated && local) local.forEach((v, k) => { if (!merged.has(k)) merged.set(k, v); });
        const rawNew = JSON.stringify(kvFromEntries(cfg, merged));
        if (rawNew !== JSON.stringify(kvFromEntries(cfg, local || new Map()))) changed = true;
        try { localStorage.setItem(lsKey, rawNew); } catch (e) {}
        kvState.last[lsKey] = new Map(Array.from(db, ([k, v]) => [k, JSON.stringify(v)]));
      }
      kvApplyToVariables();
      kvState.ready = true;
      if (!migrated) {
        for (const lsKey of Object.keys(KV_STORES)) await kvPush(lsKey, localStorage.getItem(lsKey) || '{}');
        if (!kvState.dirty.size) { try { localStorage.setItem(flag, '1'); } catch (e) {} }
      }
      kvState.lastHydrate = Date.now();
      return changed;
    }
    function kvReset() {
      Object.values(kvState.timers).forEach(clearTimeout);
      kvState.ready = false; kvState.uid = null; kvState.last = {}; kvState.timers = {}; kvState.pending = {}; kvState.dirty = new Set();
      ['korpo_personal_monthly_targets_v1', 'korpo_personal_comm_targets_v1', 'korpo_personal_bonus_pct_v1', 'korpo_read_mention_ids_v1', 'korpo_chat_last_seen_v1']
        .forEach(k => { try { localStorage.removeItem(k); } catch (e) {} }); // pas de fuite entre comptes sur un même appareil
    }
    document.addEventListener('visibilitychange', async () => {
      if (document.visibilityState !== 'visible' || !kvState.ready || Date.now() - kvState.lastHydrate < 60000) return;
      if (await hydrateKv()) { try { await loadAllData(); } catch (e) {} }
    });
    window.addEventListener('online', () => { kvState.dirty.forEach(k => kvPush(k, localStorage.getItem(k) || '{}')); });

    // =========================================================================
    // PREUVES PHOTO : stockées dans Supabase Storage (bucket privé « preuves-ventes »).
    // ventes.preuve_url contient le chemin du fichier ; l'affichage passe par des URL signées (1 h).
    // Les anciennes preuves en base64 restent lisibles et sont migrées automatiquement vers Storage.
    // =========================================================================
    const PROOF_BUCKET = 'preuves-ventes';
    const proofUrlCache = new Map();          // chemin Storage -> { url, exp }
    const proofMigrationTried = new Set();    // ventes déjà tentées pendant cette session
    const isProofPath = (r) => typeof r === 'string' && r !== '' && !/^(data:|https?:|blob:)/i.test(r);
    function dataUrlToBlob(dataUrl) {
      const [head, b64] = String(dataUrl).split(',');
      const mime = (head.match(/data:([^;]+)/) || [])[1] || 'image/jpeg';
      const bin = atob(b64 || ''), u8 = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
      return new Blob([u8], { type: mime });
    }
    // Retourne le chemin Storage ; si Storage est indisponible, retourne le base64 d'origine (ancien mode)
    async function uploadProofDataUrl(dataUrl) {
      try {
        const path = `${currentUser?.id || 'anon'}/${Date.now()}-${Math.random().toString(36).slice(2, 8)}.jpg`;
        const { error } = await sb.storage.from(PROOF_BUCKET).upload(path, dataUrlToBlob(dataUrl), { contentType: 'image/jpeg', upsert: false });
        if (error) throw error;
        return path;
      } catch (e) { console.warn('Storage indisponible, preuve conservée en base64 :', e && e.message); return dataUrl; }
    }
    async function refreshProofUrls(rows) {
      const now = Date.now();
      const need = [...new Set((rows || []).map(v => v && v.preuve_url).filter(isProofPath))]
        .filter(p => !(proofUrlCache.get(p)?.exp > now + 60000));
      for (let i = 0; i < need.length; i += 100) {
        try {
          const { data, error } = await sb.storage.from(PROOF_BUCKET).createSignedUrls(need.slice(i, i + 100), 3600);
          if (error) throw error;
          (data || []).forEach(d => { if (d.signedUrl) proofUrlCache.set(d.path, { url: d.signedUrl, exp: now + 3600 * 1000 }); });
        } catch (e) { console.warn('URL signées indisponibles :', e && e.message); }
      }
    }
    // Valeur à écrire dans preuve_url : chemin (nouvelle photo), undefined (inchangée) ou null (supprimée)
    async function resolveProofForSave(hadProof) {
      const p = pendingProofDataUrl;
      if (p && p.startsWith('data:')) return await uploadProofDataUrl(p);
      if (p) return undefined;
      return hadProof ? null : undefined;
    }
    // Migration en arrière-plan des anciennes preuves (base64 en base ou dans le navigateur) vers Storage
    async function migrateLegacyProofs(rows) {
      let done = 0;
      const seen = new Set();
      for (const v of rows || []) {
        if (!v || seen.has(String(v.id)) || done >= 15) continue;
        seen.add(String(v.id));
        const sid = String(v.id);
        const legacy = (typeof v.preuve_url === 'string' && v.preuve_url.startsWith('data:')) ? v.preuve_url
          : (!v.preuve_url && typeof saleProofsStore[sid] === 'string' && saleProofsStore[sid].startsWith('data:') ? saleProofsStore[sid] : '');
        if (!legacy || proofMigrationTried.has(sid)) continue;
        proofMigrationTried.add(sid);
        const path = await uploadProofDataUrl(legacy);
        if (path === legacy) continue; // Storage indisponible
        const { error } = await sb.from('ventes').update({ preuve_url: path }).eq('id', v.id);
        if (error) { try { await sb.storage.from(PROOF_BUCKET).remove([path]); } catch (e) {} continue; }
        delete saleProofsStore[sid]; v.preuve_url = path; done++;
      }
      if (done) { saveSaleProofsStore(); try { await loadAllData(); } catch (e) {} }
    }

    // Date locale AAAA-MM-JJ (toISOString() donne la date UTC : le soir, au Québec, c'est déjà « demain »)
    function localYmd(d = new Date()) {
      return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    }
    // Cellule CSV : guillemets/virgules/retours à la ligne + neutralisation des formules Excel (=, +, -, @)
    function csvCell(v, isText = true) {
      let s = String(v ?? '');
      if (isText && /^[=+\-@\t\r]/.test(s)) s = "'" + s;
      return /[",;\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
    }
    // Montant saisi en format FR ou EN : "1 234,56", "1,234.56", "12,345", "$1.234.567"
    function parseAmount(raw) {
      let s = String(raw ?? '').replace(/[\s\u00A0$€]/g, '').replace(/CAD/i, '');
      if (!s) return NaN;
      const neg = /^\(.*\)$/.test(s) || s.startsWith('-');
      s = s.replace(/[()\-]/g, '');
      const lastDot = s.lastIndexOf('.'), lastComma = s.lastIndexOf(',');
      if (lastDot >= 0 && lastComma >= 0) {
        const dec = lastDot > lastComma ? '.' : ',';
        const thou = dec === '.' ? ',' : '.';
        s = s.split(thou).join('').replace(dec, '.');
      } else if (lastComma >= 0) {
        s = /^\d{1,3}(,\d{3})+$/.test(s) ? s.replace(/,/g, '') : s.replace(',', '.');
      } else if ((s.match(/\./g) || []).length > 1) {
        s = s.replace(/\./g, '');
      }
      const n = parseFloat(s);
      return isNaN(n) ? NaN : (neg ? -n : n);
    }
    // Lecture paginée (Supabase limite chaque requête à 1 000 lignes par défaut)
    async function fetchAllRows(buildQuery) {
      const PAGE = 1000;
      const out = [];
      let from = 0;
      for (;;) {
        const { data, error } = await buildQuery().range(from, from + PAGE - 1);
        if (error) throw error;
        if (!data || !data.length) break;
        out.push(...data);
        from += data.length;
      }
      return out;
    }

    // =========================================================================
    // ARCHITECTURE KPI UNIFIÉE : BASE FICHIER CONDENSÉ + FICHIER DE VENTES
    // =========================================================================
    const DEFAULT_KPI_REPS = ['Olivier', 'Jean-Sébastien', 'Gabryel', 'Samuel', 'Alexandre'];
    let ALL_KPI_REPS = [...DEFAULT_KPI_REPS];
    let archivedRepsStore = lsGetJson('korpo_archived_reps_v1', '{}');
    let archivedEntreprisesStore = lsGetJson('korpo_archived_entreprises_v1', '{}');

    function saveArchivedRepsStore() {
      lsSet('korpo_archived_reps_v1', JSON.stringify(archivedRepsStore));
    }

    function saveArchivedEntreprisesStore() {
      lsSet('korpo_archived_entreprises_v1', JSON.stringify(archivedEntreprisesStore));
    }

    function findArchivedRepEntry(repIdOrName) {
      if (!repIdOrName) return null;
      const key = String(repIdOrName).trim().toLowerCase();
      return Object.values(archivedRepsStore || {}).find(entry => {
        if (!entry) return false;
        if (entry.id && String(entry.id) === String(repIdOrName)) return true;
        if (entry.nom && String(entry.nom).trim().toLowerCase() === key) return true;
        return false;
      }) || null;
    }

    function isRepArchived(repIdOrName) {
      return Boolean(findArchivedRepEntry(repIdOrName));
    }

    function getRepArchiveMonthFactor(repIdOrName, ym) {
      const entry = findArchivedRepEntry(repIdOrName);
      if (!entry || !entry.archivedAt || !isValidYearMonth(ym)) return 1;
      const archDate = String(entry.archivedAt).slice(0, 10);
      const archYm = archDate.slice(0, 7);
      if (!isValidYearMonth(archYm)) return 0;
      if (ym < archYm) return 1;
      if (ym > archYm) return 0;
      const archDay = parseInt(archDate.slice(8, 10), 10) || 1;
      const totalBiz = Math.max(1, countBusinessDaysInMonth(ym));
      const activeBiz = countBusinessDaysInMonth(ym, archDay);
      return Math.min(1, Math.max(0, activeBiz / totalBiz));
    }

    function getArchivedRepNameById(repId) {
      if (!repId) return '';
      const found = Object.values(archivedRepsStore || {}).find(entry => entry && String(entry.id) === String(repId));
      return found?.nom || '';
    }

    function findArchivedEntrepriseEntry(entIdOrName) {
      if (!entIdOrName) return null;
      const key = String(entIdOrName).trim().toLowerCase();
      return Object.values(archivedEntreprisesStore || {}).find(entry => {
        if (!entry) return false;
        if (entry.id && String(entry.id) === String(entIdOrName)) return true;
        if (entry.nom && String(entry.nom).trim().toLowerCase() === key) return true;
        return false;
      }) || null;
    }

    function isEntrepriseArchived(entIdOrName) {
      return Boolean(findArchivedEntrepriseEntry(entIdOrName));
    }

    function getEntrepriseArchiveMonthFactor(entIdOrName, ym) {
      const entry = findArchivedEntrepriseEntry(entIdOrName);
      if (!entry || !entry.archivedAt || !isValidYearMonth(ym)) return 1;
      const archDate = String(entry.archivedAt).slice(0, 10);
      const archYm = archDate.slice(0, 7);
      if (!isValidYearMonth(archYm)) return 0;
      if (ym < archYm) return 1;
      if (ym > archYm) return 0;
      const archDay = parseInt(archDate.slice(8, 10), 10) || 1;
      const totalBiz = Math.max(1, countBusinessDaysInMonth(ym));
      const activeBiz = countBusinessDaysInMonth(ym, archDay);
      return Math.min(1, Math.max(0, activeBiz / totalBiz));
    }

    function getArchivedEntrepriseNameById(entId) {
      if (!entId) return '';
      const found = Object.values(archivedEntreprisesStore || {}).find(entry => entry && String(entry.id) === String(entId));
      return found?.nom || '';
    }

    function resolveSaleArchivedRepName(v) {
      if (!v) return '';
      if (v.representant_id) {
        const byId = getArchivedRepNameById(v.representant_id);
        if (byId) return byId;
      }
      const sid = v.id ? String(v.id) : '';
      if (sid) {
        const bySaleId = Object.values(archivedRepsStore || {}).find(entry => Array.isArray(entry?.saleIds) && entry.saleIds.includes(sid));
        if (bySaleId?.nom) return bySaleId.nom;
      }
      const mTag = String(v.note || '').match(/\[REP:([^\]]+)\]/i);
      if (mTag && mTag[1]) return mTag[1].trim();
      // Fallback si la vente a representant_id === null et provient d'un représentant archivé (ex. Gabryel)
      if (v.representant_id === null || v.representant_id === undefined || v.representant_id === '') {
        const archivedList = Object.values(archivedRepsStore || {}).filter(e => e && e.nom);
        if (archivedList.length === 1) return archivedList[0].nom;
        const auditEvs = sid ? (saleAuditStore[sid] || []) : [];
        const creatorEv = auditEvs.find(ev => ev && ev.by && !state.representants.some(r => r.nom === ev.by));
        if (creatorEv?.by) return creatorEv.by;
        const activeNamesLow = new Set((state.representants || []).map(r => String(r?.nom || '').trim().toLowerCase()));
        const missingDefaultReps = DEFAULT_KPI_REPS.filter(n => !activeNamesLow.has(n.toLowerCase()));
        if (missingDefaultReps.length === 1) return missingDefaultReps[0];
      }
      return '';
    }

    function resolveSaleArchivedEntName(v) {
      if (!v) return '';
      if (v.entreprise_id) {
        const byId = getArchivedEntrepriseNameById(v.entreprise_id);
        if (byId) return byId;
      }
      const sid = v.id ? String(v.id) : '';
      if (sid) {
        const bySaleId = Object.values(archivedEntreprisesStore || {}).find(entry => Array.isArray(entry?.saleIds) && entry.saleIds.includes(sid));
        if (bySaleId?.nom) return bySaleId.nom;
      }
      const mTag = String(v.note || '').match(/\[ENT:([^\]]+)\]/i);
      if (mTag && mTag[1]) return mTag[1].trim();
      if (v.entreprise_id === null || v.entreprise_id === undefined || v.entreprise_id === '') {
        const archivedEnts = Object.values(archivedEntreprisesStore || {}).filter(e => e && e.nom);
        if (archivedEnts.length === 1) return archivedEnts[0].nom;
      }
      return '';
    }

    function getActiveKpiReps() {
      return ALL_KPI_REPS.filter(rn => !isRepArchived(rn));
    }

    function syncAllKpiRepsFromState() {
      const seen = new Set();
      const merged = [];
      (state.representants || []).forEach(r => {
        const nom = String(r?.nom || '').trim();
        if (nom && !isRepArchived(nom) && !seen.has(nom)) {
          seen.add(nom);
          merged.push(nom);
        }
      });
      DEFAULT_KPI_REPS.forEach(nom => {
        if (!isRepArchived(nom) && !seen.has(nom)) {
          seen.add(nom);
          merged.push(nom);
        }
      });
      Object.values(archivedRepsStore || {}).forEach(entry => {
        const nom = String(entry?.nom || '').trim();
        if (nom && !seen.has(nom)) {
          seen.add(nom);
          merged.push(nom);
        }
      });
      ALL_KPI_REPS = merged;
      const activeRepsList = getActiveKpiReps();

      // Synchroniser aussi toutes les listes déroulantes de représentants dans le Tableau KPI (uniquement les actifs)
      const kpiSelRep = document.getElementById('kpi-select-rep');
      if (kpiSelRep) {
        const cur = kpiSelRep.value || 'ALL';
        const archivedWithHistory = Object.values(archivedRepsStore || {}).map(e => String(e?.nom || '').trim()).filter(n => n && !activeRepsList.includes(n));
        kpiSelRep.innerHTML = `<option value="ALL">👥 Toute l'équipe</option>` +
          activeRepsList.map(n => `<option value="${esc(n)}">${esc(n === 'Jean-Sébastien' ? 'Jean-Sébastien (J-S)' : n)}</option>`).join('') +
          (archivedWithHistory.length ? `<optgroup label="Anciens membres (Historique)">${archivedWithHistory.map(n => `<option value="${esc(n)}">${esc(n)} (Archivé)</option>`).join('')}</optgroup>` : '');
        const allAllowed = ['ALL', ...activeRepsList, ...archivedWithHistory];
        kpiSelRep.value = allAllowed.includes(cur) ? cur : 'ALL';
      }
      ['kpi-edit-rep', 'rep-target-select'].forEach(selId => {
        const el = document.getElementById(selId);
        if (!el) return;
        const cur = el.value;
        el.innerHTML = activeRepsList.map(n => `<option value="${esc(n)}">${esc(n)}</option>`).join('');
        if (cur && activeRepsList.includes(cur)) el.value = cur;
      });
      const targetRepSel = document.getElementById('target-rep-select');
      if (targetRepSel) {
        const cur = targetRepSel.value || 'ALL';
        targetRepSel.innerHTML = `<option value="ALL">Toute l'équipe</option>` +
          activeRepsList.map(n => `<option value="${esc(n)}">${esc(n)}</option>`).join('');
        targetRepSel.value = (cur === 'ALL' || activeRepsList.includes(cur)) ? cur : 'ALL';
      }
    }
    const isValidYearMonth = (ym) => typeof ym === 'string' && /^\d{4}-(0[1-9]|1[0-2])$/.test(ym.trim());

    let kpiStore = {};
    let kpiHebdoRows = [];
    let condensedBaselineStore = lsGetJson('korpo_kpi_condensed_v1', '{}');
    let monthlyRdvOverrides = lsGetJson('korpo_kpi_monthly_rdv_v1', '{}');
    let teamKpiSnapshotStore = lsGetJson('korpo_kpi_team_snapshot_v1', '{}');
    let lastSyncedKpiSnapshotHashes = {};
    let kpiSnapshotSyncTimer = null;

    function saveCondensedBaselineStore() {
      lsSet('korpo_kpi_condensed_v1', JSON.stringify(condensedBaselineStore));
    }

    let dailyRdvLogsStore = lsGetJson('korpo_daily_rdv_v1', '{}');

    function saveMonthlyRdvOverrides() {
      lsSet('korpo_kpi_monthly_rdv_v1', JSON.stringify(monthlyRdvOverrides));
    }

    function saveDailyRdvLogsStore() {
      lsSet('korpo_daily_rdv_v1', JSON.stringify(dailyRdvLogsStore));
    }

    async function persistDailyRdvToCloud(dateStr, repName, entryObj) {
      if (!sb || !currentProfile) return;
      const key = `${dateStr}__${repName}`;
      await safeUpsertKpiHebdoRow({
        mois_commission: key,
        semaine_index: -16,
        representant_nom: JSON.stringify(entryObj),
        booke: Number(entryObj.booke || 0),
        rdv_realise: Number(entryObj.rdv || 0),
        updated_at: new Date().toISOString()
      });
    }

    function getActiveDailyRdvRepName() {
      if (!currentProfile) return '';
      return resolveKpiRepName(currentProfile.nom || '') || String(currentProfile.nom || '').trim();
    }

    function getTodayDailyRdvEntry(repName = getActiveDailyRdvRepName()) {
      const todayKey = localYmd(new Date());
      const entry = dailyRdvLogsStore[`${todayKey}__${repName}`];
      return {
        dateStr: todayKey,
        booke: Math.max(0, Number(entry?.booke || 0)),
        rdv: Math.max(0, Number(entry?.rdv || 0))
      };
    }

    function renderDailyRdvTrackerCard() {
      const card = document.getElementById('daily-rdv-tracker-card');
      if (!card || !currentProfile) return;
      const repName = getActiveDailyRdvRepName();
      if (!repName) {
        card.classList.add('hidden');
        return;
      }
      card.classList.remove('hidden');
      const todayData = getTodayDailyRdvEntry(repName);
      const badgeEl = document.getElementById('daily-rdv-date-badge');
      if (badgeEl) {
        badgeEl.textContent = `Aujourd'hui (${formatDateLabelFr(todayData.dateStr)})`;
      }
      const bookeIn = document.getElementById('daily-rdv-booke-input');
      const rdvIn = document.getElementById('daily-rdv-realise-input');
      if (bookeIn && document.activeElement !== bookeIn) bookeIn.value = String(todayData.booke);
      if (rdvIn && document.activeElement !== rdvIn) rdvIn.value = String(todayData.rdv);

      // Cumul du mois calculé STRICTEMENT pour le représentant connecté (ses propres RDV bookés & réalisés du mois)
      const ym = todayData.dateStr.slice(0, 7);
      const repNorm = resolveKpiRepName(repName) || repName;
      let monthBooke = 0;
      let monthRdv = 0;
      let hasManualMonthOverride = false;

      Object.entries(monthlyRdvOverrides || {}).forEach(([k, rdvOv]) => {
        const [kYm, kRep] = k.split('_');
        if (kYm === ym && (resolveKpiRepName(kRep) || kRep) === repNorm && rdvOv) {
          monthBooke += Number(rdvOv.booke || 0);
          monthRdv += Number(rdvOv.rdv || 0);
          hasManualMonthOverride = true;
        }
      });

      if (!hasManualMonthOverride) {
        (kpiHebdoRows || []).forEach(row => {
          const wIdx = Number(row.semaine_index);
          if (wIdx >= 0 && wIdx < 4 && row.mois_commission === ym) {
            const rNorm = resolveKpiRepName(row.representant_nom) || String(row.representant_nom || '').trim();
            if (rNorm === repNorm) {
              monthBooke += Number(row.booke || 0);
              monthRdv += Number(row.rdv_realise || 0);
            }
          }
        });
      }

      const condEntry = condensedBaselineStore?.[ym]?.[repNorm];
      if (condEntry && typeof condEntry === 'object') {
        monthBooke += Number(condEntry.booke || 0);
        monthRdv += Number(condEntry.rdv || 0);
      }

      Object.entries(dailyRdvLogsStore || {}).forEach(([key, dEntry]) => {
        if (!dEntry || typeof dEntry !== 'object') return;
        const parts = key.split('__');
        const dStr = dEntry.date || parts[0] || '';
        const rNorm = resolveKpiRepName(dEntry.rep || parts[1] || '') || String(dEntry.rep || parts[1] || '').trim();
        if (dStr.slice(0, 7) === ym && rNorm === repNorm) {
          monthBooke += Number(dEntry.booke || 0);
          monthRdv += Number(dEntry.rdv || 0);
        }
      });

      const sumEl = document.getElementById('daily-rdv-month-summary');
      if (sumEl) {
        sumEl.textContent = `${Math.max(0, monthBooke)} booké(s) • ${Math.max(0, monthRdv)} réalisé(s)`;
      }
    }

    async function saveTodayDailyRdvValues(newBooke, newRdv) {
      if (!currentProfile) return;
      const repName = getActiveDailyRdvRepName();
      if (!repName) return;
      const todayStr = localYmd(new Date());
      const safeBooke = Math.max(0, Math.round(Number(newBooke) || 0));
      const safeRdv = Math.max(0, Math.round(Number(newRdv) || 0));
      const nowIso = new Date().toISOString();

      const payload = {
        date: todayStr,
        rep: repName,
        booke: safeBooke,
        rdv: safeRdv,
        updatedAt: nowIso
      };
      dailyRdvLogsStore[`${todayStr}__${repName}`] = payload;
      saveDailyRdvLogsStore();

      rebuildKpiStore();
      renderDailyRdvTrackerCard();
      renderKpiDashboard();

      await persistDailyRdvToCloud(todayStr, repName, payload);
      toast(`Rendez-vous du jour enregistrés (${safeBooke} bookés, ${safeRdv} réalisés) ✓`);
    }

    window.adjustDailyRdvCounter = async (field, delta) => {
      const cur = getTodayDailyRdvEntry();
      const nextBooke = field === 'booke' ? Math.max(0, cur.booke + delta) : cur.booke;
      const nextRdv = field === 'rdv' ? Math.max(0, cur.rdv + delta) : cur.rdv;
      await saveTodayDailyRdvValues(nextBooke, nextRdv);
    };

    window.setDailyRdvCounterFromInput = async (field, rawVal) => {
      const cur = getTodayDailyRdvEntry();
      const parsed = Math.max(0, Math.round(parseFloat(rawVal) || 0));
      const nextBooke = field === 'booke' ? parsed : cur.booke;
      const nextRdv = field === 'rdv' ? parsed : cur.rdv;
      await saveTodayDailyRdvValues(nextBooke, nextRdv);
    };

    // Vérification automatique du passage à minuit (00:00) pour remettre les compteurs du jour à 0
    let _lastTrackedDailyRdvDate = localYmd(new Date());
    setInterval(() => {
      const nowYmd = localYmd(new Date());
      if (nowYmd !== _lastTrackedDailyRdvDate) {
        _lastTrackedDailyRdvDate = nowYmd;
        if (typeof currentDateStr !== 'undefined') currentDateStr = nowYmd;
        if (typeof currentMonthStr !== 'undefined') currentMonthStr = nowYmd.slice(0, 7);
        rebuildKpiStore();
        renderDailyRdvTrackerCard();
        renderKpiDashboard();
      }
    }, 30000);

    function hasActiveCondensedBaseline() {
      if (!condensedBaselineStore || typeof condensedBaselineStore !== 'object') return false;
      return Object.keys(condensedBaselineStore).some(ym => {
        if (!isValidYearMonth(ym)) return false;
        const m = condensedBaselineStore[ym];
        return m && typeof m === 'object' && Object.keys(m).some(k => !k.startsWith('_'));
      });
    }

    function normalizeKpiStoreReps() {
      if (!kpiStore || typeof kpiStore !== 'object') kpiStore = {};
      Object.keys(kpiStore).forEach(ym => {
        if (!isValidYearMonth(ym)) {
          delete kpiStore[ym];
          return;
        }
        const mObj = kpiStore[ym];
        if (!mObj || typeof mObj !== 'object') {
          delete kpiStore[ym];
          return;
        }
        const nw = mObj.num_weeks || 4;
        if (!mObj.reps || typeof mObj.reps !== 'object') mObj.reps = {};
        if (!mObj.team || typeof mObj.team !== 'object') {
          mObj.team = {
            'Booké': { weeks: Array(nw).fill(0), cumul: 0, objectif: 200 },
            'RDV RÉA': { weeks: Array(nw).fill(0), cumul: 0, objectif: 150 },
            'VENTE OTS': { weeks: Array(nw).fill(0), cumul: 0, objectif: 27 },
            'VENTE SUIVI': { weeks: Array(nw).fill(0), cumul: 0, objectif: 49 },
            'VENTE TOTAL': { weeks: Array(nw).fill(0), cumul: 0, objectif: 76 },
            'VENTE $': { weeks: Array(nw).fill(0), cumul: 0, objectif: 568000 }
          };
        }
        if (!mObj.entreprises || typeof mObj.entreprises !== 'object') mObj.entreprises = {};
        ALL_KPI_REPS.forEach(rn => {
          if (!mObj.reps[rn]) {
            mObj.reps[rn] = {
              'Booké': { weeks: Array(nw).fill(0), cumul: 0, objectif: 80 },
              'RDV Réalisé': { weeks: Array(nw).fill(0), cumul: 0, objectif: 60 },
              'Ventes OTS': { weeks: Array(nw).fill(0), cumul: 0, objectif: 10 },
              'Ventes SUIVI': { weeks: Array(nw).fill(0), cumul: 0, objectif: 20 },
              'Vente $': { weeks: Array(nw).fill(0), cumul: 0, objectif: 150000 },
              '% Close': { weeks: Array(nw).fill(0), cumul: 0, objectif: 35 }
            };
          } else {
            ['Booké', 'RDV Réalisé', 'Ventes OTS', 'Ventes SUIVI', 'Vente $', '% Close'].forEach(mk => {
              if (!mObj.reps[rn][mk]) {
                mObj.reps[rn][mk] = { weeks: Array(nw).fill(0), cumul: 0, objectif: 0 };
              }
            });
          }
        });
        Object.values(mObj.entreprises || {}).forEach(eObj => {
          if (!eObj.reps || typeof eObj.reps !== 'object') eObj.reps = {};
          const ew = eObj.total?.weeks?.length || nw;
          ALL_KPI_REPS.forEach(rn => {
            if (!eObj.reps[rn]) {
              eObj.reps[rn] = { weeks: Array(ew).fill(0), cumul: 0, objectif: 0 };
            }
          });
        });
      });
    }

    async function safeUpsertKpiHebdoRow(rowObj) {
      if (!sb || !currentProfile) return { error: null };
      if (Number(rowObj.semaine_index) >= -1) {
        try {
          const upRes = await sb.from('kpi_hebdo').upsert(rowObj, {
            onConflict: 'mois_commission,semaine_index,representant_nom'
          });
          if (!upRes.error) return upRes;
        } catch (e) {}
      }

      try {
        let q = sb.from('kpi_hebdo').select('id').eq('mois_commission', rowObj.mois_commission).eq('semaine_index', rowObj.semaine_index);
        if (Number(rowObj.semaine_index) >= -1) {
          q = q.eq('representant_nom', rowObj.representant_nom);
        }
        const { data: existingRows } = await q.limit(1);
        if (existingRows && existingRows.length > 0) {
          return await sb.from('kpi_hebdo').update({
            representant_nom: rowObj.representant_nom,
            booke: rowObj.booke,
            rdv_realise: rowObj.rdv_realise,
            updated_at: rowObj.updated_at || new Date().toISOString()
          }).eq('id', existingRows[0].id);
        } else {
          return await sb.from('kpi_hebdo').insert(rowObj);
        }
      } catch (err) {
        return { error: err };
      }
    }

    async function persistCondensedEntryToCloud(ym, repName, entryObj) {
      if (!sb || !currentProfile) return;
      const key = `${ym}__${repName}`;
      await safeUpsertKpiHebdoRow({
        mois_commission: key,
        semaine_index: -9,
        representant_nom: JSON.stringify(entryObj),
        booke: Number(entryObj.booke || 0),
        rdv_realise: Number(entryObj.rdv || 0),
        updated_at: new Date().toISOString()
      });
    }

    async function persistMonthlyRdvToCloud(ym, repName, rdvObj) {
      if (!sb || !currentProfile) return;
      const key = `${ym}_${repName}`;
      await safeUpsertKpiHebdoRow({
        mois_commission: key,
        semaine_index: -10,
        representant_nom: JSON.stringify(rdvObj),
        booke: Number(rdvObj.booke || 0),
        rdv_realise: Number(rdvObj.rdv || 0),
        updated_at: new Date().toISOString()
      });
    }

    // =========================================================================
    // JOURNAL D'AUDIT DES VENTES & CORBEILLE DE RESTAURATION (30 JOURS)
    // =========================================================================
    let saleAuditStore = lsGetJson('korpo_sale_audit_v1', '{}');
    let saleTrashStore = lsGetJson('korpo_sale_trash_v1', '[]');
    const TRASH_RETENTION_MS = 30 * 24 * 60 * 60 * 1000;

    function pruneExpiredTrash() {
      const now = Date.now();
      const before = saleTrashStore.length;
      saleTrashStore = saleTrashStore.filter(item => {
        const delTime = item.deletedAt ? new Date(item.deletedAt).getTime() : now;
        return (now - delTime) <= TRASH_RETENTION_MS;
      });
      if (saleTrashStore.length !== before) {
        lsSet('korpo_sale_trash_v1', JSON.stringify(saleTrashStore));
      }
    }

    function mergeCloudAuditAndTrash() {
      if (!Array.isArray(kpiHebdoRows)) return;
      [...kpiHebdoRows].sort((a, b) => String(a.updated_at || '').localeCompare(String(b.updated_at || ''))).forEach(r => {
        const sIdx = Number(r.semaine_index);
        if (sIdx === -2 && r.mois_commission && r.representant_nom) {
          try {
            const parsedEvents = JSON.parse(r.representant_nom);
            if (Array.isArray(parsedEvents)) {
              const sid = String(r.mois_commission);
              const existing = saleAuditStore[sid] || [];
              const merged = [...existing];
              parsedEvents.forEach(ev => {
                if (!merged.some(m => m.at === ev.at && m.action === ev.action && m.by === ev.by)) {
                  merged.push(ev);
                }
              });
              merged.sort((a, b) => (a.at || '').localeCompare(b.at || ''));
              saleAuditStore[sid] = merged;
            }
          } catch (e) {}
        } else if (sIdx === -4 && r.mois_commission && r.representant_nom) {
          try {
            const ovData = JSON.parse(r.representant_nom);
            if (ovData && typeof ovData === 'object') {
              saleBonusOverridesStore[String(r.mois_commission)] = ovData;
            }
          } catch (e) {}
        } else if (sIdx === -5 && r.mois_commission && r.representant_nom) {
          try {
            const coData = JSON.parse(r.representant_nom);
            if (coData && coData.origYm) {
              carriedOverSalesStore[String(r.mois_commission)] = coData;
            }
          } catch (e) {}
        } else if (sIdx === -8 && r.mois_commission && r.representant_nom) {
          try {
            const cMeta = JSON.parse(r.representant_nom);
            if (cMeta && typeof cMeta === 'object') {
              companyTargetsMetaStore[String(r.mois_commission)] = cMeta;
              lsSet('korpo_company_targets_meta_v1', JSON.stringify(companyTargetsMetaStore));
              applyCompanyTargetsMetaToStore();
            }
          } catch (e) {}
        } else if (sIdx === -7 && r.mois_commission && r.representant_nom) {
          try {
            const bData = JSON.parse(r.representant_nom);
            if (bData && typeof bData.bonusPct === 'number') {
              personalBonusPctTargets[String(r.mois_commission)] = bData.bonusPct;
              lsSet('korpo_personal_bonus_pct_v1', JSON.stringify(personalBonusPctTargets));
            }
          } catch (e) {}
        } else if (sIdx === -9 && r.mois_commission && r.representant_nom) {
          try {
            const cData = JSON.parse(r.representant_nom);
            const [cYm, cRep] = String(r.mois_commission).split('__');
            if (isValidYearMonth(cYm) && cRep && cData && typeof cData === 'object') {
              if (!condensedBaselineStore[cYm]) condensedBaselineStore[cYm] = {};
              const localEntry = condensedBaselineStore[cYm][cRep];
              if (!localEntry || !localEntry.updatedAt || (cData.updatedAt && cData.updatedAt >= localEntry.updatedAt)) {
                condensedBaselineStore[cYm][cRep] = cData;
              }
            }
          } catch (e) {}
        } else if (sIdx === -16 && r.mois_commission && r.representant_nom) {
          try {
            const dData = JSON.parse(r.representant_nom);
            const key = String(r.mois_commission);
            if (dData && typeof dData === 'object') {
              const localD = dailyRdvLogsStore[key];
              if (!localD || !localD.updatedAt || (dData.updatedAt && dData.updatedAt >= localD.updatedAt)) {
                dailyRdvLogsStore[key] = dData;
              }
            }
          } catch (e) {}
        } else if (sIdx === -10 && r.mois_commission && r.representant_nom) {
          try {
            const rdvData = JSON.parse(r.representant_nom);
            const key = String(r.mois_commission);
            if (rdvData && typeof rdvData === 'object') {
              const localR = monthlyRdvOverrides[key];
              if (!localR || !localR.updatedAt || (rdvData.updatedAt && rdvData.updatedAt >= localR.updatedAt)) {
                monthlyRdvOverrides[key] = rdvData;
              }
            }
          } catch (e) {}
        } else if (sIdx === -14 && r.mois_commission === 'red_phone_cfg' && r.representant_nom) {
          try {
            const parsedRp = JSON.parse(r.representant_nom);
            if (parsedRp && typeof parsedRp === 'object') {
              redPhoneConfig = { ...redPhoneConfig, ...parsedRp };
              lsSet('korpo_red_phone_cfg_v1', JSON.stringify(redPhoneConfig));
              syncRedPhoneUI();
            }
          } catch (e) {}
        } else if (sIdx === -12 && r.mois_commission === 'wave_company_seq_v1' && r.representant_nom) {
          try {
            const parsedSeq = JSON.parse(r.representant_nom);
            if (parsedSeq && typeof parsedSeq === 'object') {
              lsSet('korpo_wave_company_seq_v1', JSON.stringify(parsedSeq));
              renderAutoEmailCompanySeqInputs();
            }
          } catch (e) {}
        } else if (sIdx === -11 && r.mois_commission && r.representant_nom) {
          try {
            const snapData = JSON.parse(r.representant_nom);
            const ym = String(r.mois_commission);
            if (isValidYearMonth(ym) && snapData && typeof snapData === 'object') {
              teamKpiSnapshotStore[ym] = snapData;
            }
          } catch (e) {}
        } else if (sIdx === -3 && r.mois_commission && r.representant_nom) {
          try {
            const trashItem = JSON.parse(r.representant_nom);
            if (trashItem && trashItem.sale && trashItem.sale.id) {
              const tid = String(trashItem.sale.id);
              const stillActive = (state.ventes || []).some(v => String(v.id) === tid);
              if (!stillActive && !saleTrashStore.some(x => String(x.sale?.id) === tid)) {
                saleTrashStore.push(trashItem);
              }
            }
          } catch (e) {}
        }
      });
      pruneExpiredTrash();
      lsSet('korpo_sale_audit_v1', JSON.stringify(saleAuditStore));
      lsSet('korpo_sale_trash_v1', JSON.stringify(saleTrashStore));
      lsSet('korpo_sale_bonus_overrides_v1', JSON.stringify(saleBonusOverridesStore));
      lsSet('korpo_carried_over_sales_v1', JSON.stringify(carriedOverSalesStore));
      saveCondensedBaselineStore();
      saveMonthlyRdvOverrides();
      lsSet('korpo_kpi_team_snapshot_v1', JSON.stringify(teamKpiSnapshotStore));
      applySaleBonusOverridesToState();
      applyCarryOverMetadataToState();
      // Récupérer l'attribution des ventes déjà détachées (representant_id / entreprise_id null) lors d'une suppression précédente
      const archivedRepEntries = Object.values(archivedRepsStore || {});
      const archivedEntEntries = Object.values(archivedEntreprisesStore || {});
      if (archivedRepEntries.length === 1 || archivedEntEntries.length === 1) {
        (state.ventes || []).forEach(v => {
          if (v.representant_id === null && archivedRepEntries.length === 1) {
            const onlyRep = archivedRepEntries[0];
            if (onlyRep && Array.isArray(onlyRep.saleIds) && !onlyRep.saleIds.includes(String(v.id))) {
              onlyRep.saleIds.push(String(v.id));
            }
          }
          if (v.entreprise_id === null && archivedEntEntries.length === 1) {
            const onlyEnt = archivedEntEntries[0];
            if (onlyEnt && Array.isArray(onlyEnt.saleIds) && !onlyEnt.saleIds.includes(String(v.id))) {
              onlyEnt.saleIds.push(String(v.id));
            }
          }
        });
      }

      updateTrashBadges();
    }

    let saleBonusOverridesStore = lsGetJson('korpo_sale_bonus_overrides_v1', '{}');

    async function saveSaleBonusOverride(saleId, overrideObjOrNull) {
      if (!saleId) return;
      const sid = String(saleId);
      if (overrideObjOrNull) {
        saleBonusOverridesStore[sid] = overrideObjOrNull;
      } else {
        delete saleBonusOverridesStore[sid];
      }
      lsSet('korpo_sale_bonus_overrides_v1', JSON.stringify(saleBonusOverridesStore));

      if (sb && currentProfile) {
        try {
          if (overrideObjOrNull) {
            await safeUpsertKpiHebdoRow({
              mois_commission: sid,
              semaine_index: -4,
              representant_nom: JSON.stringify(overrideObjOrNull),
              booke: 0,
              rdv_realise: 0,
              updated_at: new Date().toISOString()
            });
          } else {
            await sb.from('kpi_hebdo').delete().eq('mois_commission', sid).eq('semaine_index', -4);
          }
        } catch (e) {}
      }
    }

    function mapAdminPersonalSaleRow(v) {
      const cOli = Number(v.commission_olivier || 0);
      const cRep = Number(v.commission_rep || 0);
      const isDirectOwnSale = currentProfile && String(v.representant_id) === String(currentProfile.id);
      const personalComm = (cOli !== 0 || !isDirectOwnSale) ? cOli : cRep;
      return {
        ...v,
        client_nom: state.clients.find(c => c.id === v.client_id)?.nom || '—',
        entreprise_nom: state.entreprises.find(e => e.id === v.entreprise_id)?.nom || resolveSaleArchivedEntName(v) || '—',
        rep_nom: state.representants.find(r => r.id === v.representant_id)?.nom || resolveSaleArchivedRepName(v) || '—',
        commission_rep: personalComm,
        _repCommAmount: cRep
      };
    }

    function normalizeRepPersonalSaleRow(v) {
      const rawComm = v.commission_rep ?? v.commission ?? v.commission_representant;
      const resolvedComm = (rawComm !== undefined && rawComm !== null && !Number.isNaN(Number(rawComm)))
        ? Number(rawComm)
        : Math.round(Number(v.valeur || 0) * Number(v.taux_rep_applique || 0) * 100) / 100;
      return {
        ...v,
        commission_rep: resolvedComm,
        _repCommAmount: resolvedComm
      };
    }

    function applySaleBonusOverridesToState() {
      const isAdmin = currentProfile?.role === 'admin';
      const applyToRow = (row, isMesVentesForAdmin = false) => {
        const sid = String(row.id);
        const ov = saleBonusOverridesStore[sid];
        if (!ov) return row;
        const baseR = getRateForSale(row.representant_id, row.entreprise_id, row.client_id);
        const isBonusHigher = baseR ? (Number(ov.tRep) > Number(baseR.taux) + 0.0001) : !ov.noBonus;
        const isDirectOwn = currentProfile && String(row.representant_id) === String(currentProfile.id);
        const adminPersonalOv = (Number(ov.cOli) !== 0 || !isDirectOwn) ? Number(ov.cOli) : Number(ov.cRep);
        const updated = {
          ...row,
          taux_rep_applique: Number(ov.tRep),
          taux_olivier_applique: Number(ov.tOli),
          commission_rep: isMesVentesForAdmin ? adminPersonalOv : Number(ov.cRep),
          _repCommAmount: Number(ov.cRep),
          commission_olivier: Number(ov.cOli),
          _isHit: Boolean(ov.isHit || Number(ov.cRep) < 0 || Number(ov.cOli) < 0),
          _hitBearer: ov.hitBearer || null,
          _bonusApplied: Boolean(!ov.isHit && isBonusHigher && !ov.noBonus),
          _customRateApplied: Boolean(!ov.isHit && (ov.customManual || ov.noBonus))
        };
        return updated;
      };
      if (Array.isArray(state.ventes)) {
        state.ventes = state.ventes.map(v => applyToRow(v, false));
      }
      if (Array.isArray(state.mesVentes)) {
        state.mesVentes = state.mesVentes.map(v => applyToRow(v, isAdmin));
      }
    }

    // =========================================================================
    // REPORT AUTOMATIQUE MENSUEL DES VENTES NON CONFIRMÉES (HORS STATS DU NOUVEAU MOIS)
    // =========================================================================
    let carriedOverSalesStore = lsGetJson('korpo_carried_over_sales_v1', '{}');

    function getSaleStatsMonth(v) {
      if (!v) return currentMonthStr;
      const sid = v.id ? String(v.id) : '';
      const co = sid ? carriedOverSalesStore[sid] : null;
      const origYm = co?.origYm || (v.date_vente && v.date_vente.length >= 7 ? v.date_vente.slice(0, 7) : '');
      const commYm = co?.movedToYm && (!v.mois_commission || v.mois_commission < co.movedToYm)
        ? co.movedToYm
        : (v.mois_commission || '');
      if (isValidYearMonth(commYm) && commYm > currentMonthStr) return commYm;
      if (origYm) return origYm;
      if (v._statsMonth) return v._statsMonth;
      return commYm || currentMonthStr;
    }

    function applyCarryOverMetadataToState() {
      const applyMeta = (row) => {
        const sid = String(row.id);
        const co = carriedOverSalesStore[sid];
        const origYm = co?.origYm || (row.date_vente && row.date_vente.length >= 7 ? row.date_vente.slice(0, 7) : row.mois_commission);
        let effectiveCommYm = co?.movedToYm && row.mois_commission < co.movedToYm ? co.movedToYm : row.mois_commission;
        if (row.statut_depot === 'a_confirmer' && ((origYm && origYm < currentMonthStr) || (effectiveCommYm && effectiveCommYm < currentMonthStr))) {
          effectiveCommYm = currentMonthStr;
        }
        const isCarried = Boolean(co || (origYm && effectiveCommYm && origYm < effectiveCommYm));
        return {
          ...row,
          mois_commission: effectiveCommYm,
          _statsMonth: isCarried ? effectiveCommYm : origYm,
          _carriedOverFrom: isCarried ? origYm : null
        };
      };
      const filterOutSupersededUnconfirmed = (rows) => {
        if (!Array.isArray(rows)) return [];
        const confirmedList = rows.filter(r => r.statut_depot === 'depot_recu');
        const unconfirmedList = rows.filter(r => r.statut_depot === 'a_confirmer');

        return rows.filter(r => {
          if (r.statut_depot !== 'a_confirmer') return true;
          // 1. Masquer si un dossier confirmé (`depot_recu`) existe déjà pour ce même client + entreprise + rep
          if (confirmedList.some(cRow => areSameClientSales(cRow, r))) return false;
          // 2. Si plusieurs entrées `a_confirmer` existent sur des mois différents pour la même vente (même client + entreprise + rep + même montant), ne garder que la plus récente
          const rKey = `${r.date_vente || r.mois_commission || ''}_${r.created_at || ''}_${r.id || ''}`;
          const hasNewerUnconfirmedCopy = unconfirmedList.some(other => {
            if (String(other.id) === String(r.id)) return false;
            if (!areSameClientSales(other, r)) return false;
            if (Math.abs(Number(other.valeur || 0) - Number(r.valeur || 0)) >= 0.01) return false;
            const oKey = `${other.date_vente || other.mois_commission || ''}_${other.created_at || ''}_${other.id || ''}`;
            return oKey > rKey;
          });
          return !hasNewerUnconfirmedCopy;
        });
      };

      if (Array.isArray(state.ventes)) {
        state.ventes = filterOutSupersededUnconfirmed(state.ventes.map(applyMeta));
      }
      if (Array.isArray(state.mesVentes)) {
        state.mesVentes = filterOutSupersededUnconfirmed(state.mesVentes.map(applyMeta));
      }
    }

    async function runMonthlyAutoRolloverAndCarryOver() {
      // 1. S'assurer que le mois courant existe toujours dans le Tableau KPI dès le 1er du mois
      ensureKpiMonthStructure(currentMonthStr);
      lsSet('korpo_kpi_store_v1', JSON.stringify(kpiStore));
      populateKpiMonthSelect(true);

      if (!sb || !currentProfile) return;

      // 2. Détecter toutes les ventes non confirmées (`a_confirmer`) des mois précédents (< currentMonthStr)
      const isAdminUser = currentProfile.role === 'admin';
      const allKnownSales = [...(state.ventes || [])];
      const toCarry = allKnownSales.filter(v => {
        if (!isAdminUser && String(v.representant_id) !== String(currentProfile.id)) return false;
        if (v.statut_depot !== 'a_confirmer') return false;
        const curCommYm = v.mois_commission || '';
        return curCommYm && curCommYm < currentMonthStr;
      });

      if (!toCarry.length) {
        applyCarryOverMetadataToState();
        return;
      }

      let movedCount = 0;
      for (const v of toCarry) {
        const sid = String(v.id);
        const origYm = carriedOverSalesStore[sid]?.origYm
          || (v.date_vente && v.date_vente.length >= 7 ? v.date_vente.slice(0, 7) : v.mois_commission);

        const record = {
          origYm,
          movedToYm: currentMonthStr,
          movedAt: new Date().toISOString()
        };
        carriedOverSalesStore[sid] = record;

        // Mettre à jour `mois_commission` vers le mois courant dans Supabase tout en conservant `date_vente` d'origine
        const { error } = await sb.from('ventes').update({ mois_commission: currentMonthStr }).eq('id', v.id);
        if (!error) {
          movedCount++;
          await recordSaleAudit(
            sid,
            `Transférée automatiquement vers ${currentMonthStr}`,
            `Dossier non confirmé de ${origYm} transféré vers ${currentMonthStr} (comptabilisé uniquement dans ${currentMonthStr})`,
            '↪️'
          );
          try {
            await safeUpsertKpiHebdoRow({
              mois_commission: sid,
              semaine_index: -5,
              representant_nom: JSON.stringify(record),
              booke: 0,
              rdv_realise: 0,
              updated_at: new Date().toISOString()
            });
          } catch (e) {}
        }
      }

      lsSet('korpo_carried_over_sales_v1', JSON.stringify(carriedOverSalesStore));
      applyCarryOverMetadataToState();

      if (movedCount > 0) {
        try {
          state.classement = await fetchAllRows(() => sb.from('vue_classement_ventes').select('*'));
        } catch (e) {}
        toast(`↪️ ${movedCount} vente(s) non confirmée(s) mise(s) à 0 sur le mois précédent et transférée(s) vers ${currentMonthStr}`);
      }
    }

    function updateTrashBadges() {
      pruneExpiredTrash();
      const isAdmin = currentProfile?.role === 'admin';
      const visibleTrash = isAdmin
        ? saleTrashStore
        : saleTrashStore.filter(t => String(t.sale?.representant_id) === String(currentProfile?.id));
      document.querySelectorAll('.trash-badge-count').forEach(el => {
        el.textContent = String(visibleTrash.length);
        el.className = visibleTrash.length > 0
          ? 'trash-badge-count px-1.5 py-0.2 text-[10px] font-extrabold rounded-full bg-rose-100 text-rose-700'
          : 'trash-badge-count px-1.5 py-0.2 text-[10px] font-extrabold rounded-full bg-slate-100 text-slate-600';
      });
    }

    async function recordSaleAudit(saleId, action, detail = '', icon = '📝') {
      if (!saleId) return;
      const sid = String(saleId);
      const actor = currentProfile?.nom || 'Système';
      const entry = {
        at: new Date().toISOString(),
        by: actor,
        action,
        detail,
        icon
      };
      if (!Array.isArray(saleAuditStore[sid])) saleAuditStore[sid] = [];
      saleAuditStore[sid].push(entry);
      lsSet('korpo_sale_audit_v1', JSON.stringify(saleAuditStore));

      if (sb && currentProfile) {
        try {
          await safeUpsertKpiHebdoRow({
            mois_commission: sid,
            semaine_index: -2,
            representant_nom: JSON.stringify(saleAuditStore[sid].slice(-25)),
            booke: 0,
            rdv_realise: 0,
            updated_at: new Date().toISOString()
          });
        } catch (e) {}
      }
    }

    async function moveSalesToTrash(salesArray) {
      if (!Array.isArray(salesArray) || !salesArray.length) return true;
      const actor = currentProfile?.nom || 'Utilisateur';
      const nowIso = new Date().toISOString();
      let cloudFailed = 0;

      for (const v of salesArray) {
        const sid = String(v.id);
        await recordSaleAudit(sid, 'Déplacée dans la corbeille', `Supprimée par ${actor} (restaurable pendant 30 jours)`, '🗑️');
        const repNom = v.rep_nom || state.representants.find(r => String(r.id) === String(v.representant_id))?.nom || '—';
        const entNom = v.entreprise_nom || state.entreprises.find(e => String(e.id) === String(v.entreprise_id))?.nom || '—';
        const cliNom = v.client_nom || state.clients.find(c => String(c.id) === String(v.client_id))?.nom || '—';
        const trashObj = {
          deletedAt: nowIso,
          deletedBy: actor,
          repNom,
          entNom,
          cliNom,
          proof: getProofRef(v) || '',
          audit: saleAuditStore[sid] || [],
          sale: {
            id: v.id,
            representant_id: v.representant_id,
            entreprise_id: v.entreprise_id,
            client_id: v.client_id,
            valeur: Number(v.valeur || 0),
            date_vente: v.date_vente,
            mois_commission: v.mois_commission,
            taux_rep_applique: Number(v.taux_rep_applique || 0),
            taux_olivier_applique: Number(v.taux_olivier_applique || 0),
            commission_rep: Number(v.commission_rep || 0),
            commission_olivier: Number(v.commission_olivier || 0),
            statut_depot: v.statut_depot || 'a_confirmer',
            statut_paiement: v.statut_paiement || 'non_paye',
            note: v.note || ''
          }
        };
        saleTrashStore = [trashObj, ...saleTrashStore.filter(x => String(x.sale?.id) !== sid)];

        if (sb && currentProfile) {
          try {
            // Copie cloud allégée : une grosse image base64 ferait échouer l'écriture (limite d'index)
            const cloudObj = { ...trashObj, proof: (trashObj.proof && trashObj.proof.length > 1500) ? '' : trashObj.proof };
            const cr = await safeUpsertKpiHebdoRow({
              mois_commission: sid,
              semaine_index: -3,
              representant_nom: JSON.stringify(cloudObj),
              booke: 0,
              rdv_realise: 0,
              updated_at: nowIso
            });
            if (cr && cr.error) cloudFailed++;
          } catch (e) { cloudFailed++; }
        } else {
          cloudFailed++;
        }
      }
      lsSet('korpo_sale_trash_v1', JSON.stringify(saleTrashStore));
      updateTrashBadges();
      return cloudFailed === 0;
    }

    function fmtAuditDateTime(isoStr) {
      if (!isoStr) return '';
      try {
        const d = new Date(isoStr);
        return d.toLocaleDateString('fr-CA', { year: 'numeric', month: 'short', day: 'numeric' }) +
          ' à ' + d.toLocaleTimeString('fr-CA', { hour: '2-digit', minute: '2-digit' });
      } catch (e) {
        return isoStr;
      }
    }

    window.openSaleAuditModal = (saleId) => {
      const sid = String(saleId);
      const v = state.ventes.find(x => String(x.id) === sid) || state.mesVentes.find(x => String(x.id) === sid);
      if (!v) return;

      const repNom = v.rep_nom || state.representants.find(r => String(r.id) === String(v.representant_id))?.nom || '—';
      const entNom = v.entreprise_nom || state.entreprises.find(e => String(e.id) === String(v.entreprise_id))?.nom || '—';
      const cliNom = v.client_nom || state.clients.find(c => String(c.id) === String(v.client_id))?.nom || '—';

      document.getElementById('audit-modal-title').textContent = `Journal d'audit — ${cliNom}`;
      document.getElementById('audit-modal-subtitle').textContent = `${entNom} • Rep : ${repNom} • ${fmtMoney(v.valeur)} (${v.date_vente})`;

      const recorded = [...(saleAuditStore[sid] || [])];
      // Si la vente n'a pas encore d'entrée de création enregistrée (dossier importé ou antérieur), synthétiser sa trace initiale
      if (!recorded.some(e => e.action === 'Vente créée')) {
        recorded.unshift({
          at: v.created_at || (v.date_vente ? `${v.date_vente}T12:00:00` : new Date().toISOString()),
          by: repNom,
          action: 'Vente enregistrée',
          detail: `Dossier ${entNom} — ${cliNom} (${fmtMoney(v.valeur)})`,
          icon: '💵'
        });
      }

      const body = document.getElementById('audit-modal-body');
      body.innerHTML = `
        <div class="bg-white rounded-xl border border-slate-200/90 p-3.5 flex flex-wrap items-center justify-between gap-2 text-xs">
          <div>
            <span class="text-slate-400">Statut actuel :</span>
            <strong class="${v.statut_depot === 'depot_recu' ? 'text-emerald-700' : 'text-amber-700'}">${v.statut_depot === 'depot_recu' ? (isNoDepositSale(v) ? 'Sans dépôt requis' : 'Dépôt reçu') : 'À confirmer'}</strong>
            <span class="mx-1 text-slate-300">•</span>
            <strong class="${v.statut_paiement === 'paye' ? 'text-indigo-700' : 'text-slate-700'}">${v.statut_paiement === 'paye' ? 'Payé' : 'Non payé'}</strong>
          </div>
          <span class="text-[11px] font-semibold text-slate-500">${recorded.length} événement(s)</span>
        </div>

        <div class="relative pl-6 space-y-3 before:content-[''] before:absolute before:left-2.5 before:top-2 before:bottom-2 before:w-0.5 before:bg-slate-200">
          ${recorded.slice().reverse().map(ev => `
            <div class="relative bg-white rounded-xl border border-slate-200/80 p-3 shadow-2xs">
              <span class="absolute -left-6 top-3 w-5 h-5 rounded-full bg-indigo-50 border border-indigo-200 flex items-center justify-center text-[11px]">${ev.icon || '📝'}</span>
              <div class="flex items-center justify-between gap-2">
                <span class="text-xs font-extrabold text-slate-900">${esc(ev.action)}</span>
                <span class="text-[10px] font-semibold text-slate-400 tabular-nums">${fmtAuditDateTime(ev.at)}</span>
              </div>
              ${ev.detail ? `<div class="text-xs text-slate-600 mt-0.5">${esc(ev.detail)}</div>` : ''}
              <div class="text-[11px] font-semibold text-indigo-600 mt-1">Par : ${esc(ev.by)}</div>
            </div>
          `).join('')}
        </div>
      `;
      document.getElementById('audit-modal').classList.remove('hidden');
    };

    window.openTrashModal = () => {
      pruneExpiredTrash();
      const isAdmin = currentProfile?.role === 'admin';
      const visibleTrash = isAdmin
        ? saleTrashStore
        : saleTrashStore.filter(t => String(t.sale?.representant_id) === String(currentProfile?.id));

      document.getElementById('trash-count-summary').textContent = `${visibleTrash.length} vente(s) dans la corbeille`;
      const emptyBtn = document.getElementById('btn-empty-trash');
      if (emptyBtn) emptyBtn.classList.toggle('hidden', !(isAdmin && visibleTrash.length > 0));

      const body = document.getElementById('trash-modal-body');
      if (!visibleTrash.length) {
        body.innerHTML = `<div class="py-12 text-center text-xs text-slate-400">La corbeille est vide. Les ventes supprimées apparaîtront ici pendant 30 jours.</div>`;
      } else {
        body.innerHTML = visibleTrash.map(item => {
          const s = item.sale || {};
          const sid = String(s.id);
          const daysLeft = Math.max(1, 30 - Math.floor((Date.now() - new Date(item.deletedAt).getTime()) / (1000 * 60 * 60 * 24)));
          return `
            <div class="bg-white rounded-2xl border border-slate-200/90 p-4 shadow-2xs flex flex-col sm:flex-row sm:items-center justify-between gap-3">
              <div class="space-y-1 min-w-0">
                <div class="flex items-center flex-wrap gap-2">
                  ${entBadge(item.entNom || '—')}
                  <span class="font-extrabold text-sm text-slate-900 truncate">${esc(item.cliNom || 'Client')}</span>
                  <span class="text-xs font-bold text-indigo-600">(${esc(item.repNom || '—')})</span>
                </div>
                <div class="text-xs text-slate-600 tabular-nums">
                  Date vente : <strong>${s.date_vente || '—'}</strong> • Montant : <strong>${fmtMoney(s.valeur)}</strong> • Comm. Rep : <strong class="text-emerald-700">${fmtMoney(s.commission_rep)}</strong>
                </div>
                <div class="text-[11px] text-slate-400">
                  Supprimée par <strong>${esc(item.deletedBy)}</strong> le ${fmtAuditDateTime(item.deletedAt)} • <span class="text-amber-700 font-semibold">${daysLeft}j restant(s)</span>
                </div>
              </div>
              <div class="flex items-center gap-2 shrink-0">
                <button type="button" onclick="restoreSaleFromTrash('${sid}')" class="px-3.5 py-2 text-xs font-extrabold text-white bg-emerald-600 hover:bg-emerald-700 rounded-xl shadow-xs transition">
                  ♻️ Restaurer
                </button>
                ${isAdmin ? `<button type="button" onclick="permanentlyDeleteFromTrash('${sid}')" class="px-2.5 py-2 text-xs font-semibold text-rose-600 hover:bg-rose-50 rounded-xl transition">Supprimer</button>` : ''}
              </div>
            </div>
          `;
        }).join('');
      }
      document.getElementById('trash-modal').classList.remove('hidden');
    };

    window.restoreSaleFromTrash = async (saleId) => {
      const sid = String(saleId);
      const item = saleTrashStore.find(x => String(x.sale?.id) === sid);
      if (!item || !item.sale) return;

      const s = item.sale;
      const clientId = await getOrCreateClient(item.cliNom || 'Client restauré', s.entreprise_id);
      const insertPayload = {
        representant_id: s.representant_id,
        entreprise_id: s.entreprise_id,
        client_id: clientId,
        valeur: s.valeur,
        date_vente: s.date_vente,
        mois_commission: s.mois_commission,
        taux_rep_applique: s.taux_rep_applique,
        taux_olivier_applique: s.taux_olivier_applique,
        commission_rep: s.commission_rep,
        commission_olivier: s.commission_olivier,
        statut_depot: s.statut_depot || 'a_confirmer',
        statut_paiement: s.statut_paiement || 'non_paye',
        note: s.note || '[OTS]'
      };
      if (item.proof) insertPayload.preuve_url = item.proof;

      let { data: restoredRow, error } = await sb.from('ventes').insert(insertPayload).select().single();
      if (error && item.proof && (error.message || '').includes('preuve_url')) {
        delete insertPayload.preuve_url;
        const retry = await sb.from('ventes').insert(insertPayload).select().single();
        restoredRow = retry.data;
        error = retry.error;
      }
      if (error) return toast('Erreur de restauration : ' + error.message, 'error');

      const newId = String(restoredRow.id);
      saleAuditStore[newId] = [...(item.audit || [])];
      await recordSaleAudit(newId, 'Restaurée depuis la corbeille', `Restaurée par ${currentProfile?.nom || 'Utilisateur'}`, '♻️');

      saleTrashStore = saleTrashStore.filter(x => String(x.sale?.id) !== sid);
      lsSet('korpo_sale_trash_v1', JSON.stringify(saleTrashStore));
      if (sb && currentProfile) {
        try {
          await sb.from('kpi_hebdo').delete().eq('mois_commission', sid).eq('semaine_index', -3);
        } catch (e) {}
      }

      updateTrashBadges();
      document.getElementById('trash-modal').classList.add('hidden');
      toast('Vente restaurée avec succès ♻️');
      await loadAllData();
    };

    window.permanentlyDeleteFromTrash = async (saleId) => {
      const sid = String(saleId);
      const ok = await askConfirm('Supprimer définitivement ?', 'Cette vente ne pourra plus être restaurée.');
      if (!ok) return;
      saleTrashStore = saleTrashStore.filter(x => String(x.sale?.id) !== sid);
      lsSet('korpo_sale_trash_v1', JSON.stringify(saleTrashStore));
      if (sb && currentProfile) {
        try {
          await sb.from('kpi_hebdo').delete().eq('mois_commission', sid).eq('semaine_index', -3);
        } catch (e) {}
      }
      updateTrashBadges();
      openTrashModal();
      toast('Supprimée définitivement');
    };

    document.getElementById('vente-rate-mode')?.addEventListener('change', () => {
      const mode = document.getElementById('vente-rate-mode').value;
      if (mode === 'custom') {
        const repId = document.getElementById('vente-rep').value;
        const entId = document.getElementById('vente-entreprise').value;
        const ym = document.getElementById('vente-mois-comm')?.value || currentMonthStr;
        const clientNom = document.getElementById('vente-client-nom').value.trim().toLowerCase();
        const matchedClient = state.clients.find(c => String(c.entreprise_id) === String(entId) && c.nom.toLowerCase() === clientNom);
        const eff = getEffectiveRatesWithTargetBonus(repId, entId, matchedClient?.id || null, ym, document.getElementById('vente-edit-id')?.value || null);
        const inR = document.getElementById('vente-custom-taux-rep');
        const inO = document.getElementById('vente-custom-taux-oli');
        if (inR && !inR.value) inR.value = Math.round(eff.tRep * 10000) / 100;
        if (inO && !inO.value) inO.value = Math.round(eff.tOli * 10000) / 100;
      }
      updateDrawerCommissionPreview();
    });
    document.getElementById('vente-custom-taux-rep')?.addEventListener('input', updateDrawerCommissionPreview);
    document.getElementById('vente-custom-taux-oli')?.addEventListener('input', updateDrawerCommissionPreview);
    document.getElementById('vente-is-hit')?.addEventListener('change', (e) => {
      updateDrawerCommissionPreview();
      if (e.target.checked) {
        const bVal = document.getElementById('vente-hit-bearer')?.value || 'rep';
        const targetEl = bVal === 'both' ? document.getElementById('vente-hit-amount-rep') : document.getElementById('vente-hit-amount');
        setTimeout(() => targetEl?.focus(), 50);
      }
    });
    document.getElementById('vente-hit-bearer')?.addEventListener('change', () => {
      updateDrawerCommissionPreview();
      const bVal = document.getElementById('vente-hit-bearer')?.value || 'rep';
      const targetEl = bVal === 'both' ? document.getElementById('vente-hit-amount-rep') : document.getElementById('vente-hit-amount');
      setTimeout(() => targetEl?.focus(), 30);
    });
    document.getElementById('vente-hit-amount')?.addEventListener('input', updateDrawerCommissionPreview);
    document.getElementById('vente-hit-amount-rep')?.addEventListener('input', updateDrawerCommissionPreview);
    document.getElementById('vente-hit-amount-oli')?.addEventListener('input', updateDrawerCommissionPreview);

    function updateRateAdjustModalPreview() {
      const sid = document.getElementById('rate-adjust-sale-id')?.value;
      const v = state.ventes.find(x => String(x.id) === String(sid));
      if (!v) return;
      const val = Number(v.valeur || 0);
      const rPct = parseFloat(document.getElementById('rate-adjust-in-rep')?.value) || 0;
      const oPct = parseFloat(document.getElementById('rate-adjust-in-oli')?.value) || 0;
      document.getElementById('rate-adjust-preview-rep').textContent = `Comm. Rep : ${fmtMoney(val * (rPct / 100))}`;
      document.getElementById('rate-adjust-preview-oli').textContent = `Comm. Olivier : ${fmtMoney(val * (oPct / 100))}`;
    }

    window.openRateAdjustModal = (saleId) => {
      if (currentProfile?.role !== 'admin') return;
      const v = state.ventes.find(x => String(x.id) === String(saleId));
      if (!v) return;
      const rep = state.representants.find(r => String(r.id) === String(v.representant_id))?.nom || '—';
      const ent = state.entreprises.find(e => String(e.id) === String(v.entreprise_id))?.nom || '—';
      const cli = state.clients.find(c => String(c.id) === String(v.client_id))?.nom || '—';

      document.getElementById('rate-adjust-sale-id').value = v.id;
      document.getElementById('rate-adjust-subtitle').textContent = `${rep} • ${ent} — ${cli} (${v.date_vente})`;
      document.getElementById('rate-adjust-sale-val').textContent = fmtMoney(v.valeur);
      document.getElementById('rate-adjust-in-rep').value = Math.round(Number(v.taux_rep_applique || 0) * 10000) / 100;
      document.getElementById('rate-adjust-in-oli').value = Math.round(Number(v.taux_olivier_applique || 0) * 10000) / 100;
      updateRateAdjustModalPreview();
      document.getElementById('rate-adjust-modal').classList.remove('hidden');
    };

    document.getElementById('rate-adjust-in-rep')?.addEventListener('input', updateRateAdjustModalPreview);
    document.getElementById('rate-adjust-in-oli')?.addEventListener('input', updateRateAdjustModalPreview);

    document.getElementById('btn-rate-adjust-reset-base')?.addEventListener('click', () => {
      const sid = document.getElementById('rate-adjust-sale-id')?.value;
      const v = state.ventes.find(x => String(x.id) === String(sid));
      if (!v) return;
      const eff = getEffectiveRatesWithTargetBonus(v.representant_id, v.entreprise_id, v.client_id, getSaleStatsMonth(v), v.id);
      document.getElementById('rate-adjust-in-rep').value = Math.round(eff.baseRep * 10000) / 100;
      document.getElementById('rate-adjust-in-oli').value = Math.round(eff.baseOli * 10000) / 100;
      updateRateAdjustModalPreview();
    });

    document.getElementById('btn-rate-adjust-reset-auto')?.addEventListener('click', () => {
      const sid = document.getElementById('rate-adjust-sale-id')?.value;
      const v = state.ventes.find(x => String(x.id) === String(sid));
      if (!v) return;
      const eff = getEffectiveRatesWithTargetBonus(v.representant_id, v.entreprise_id, v.client_id, getSaleStatsMonth(v), v.id);
      document.getElementById('rate-adjust-in-rep').value = Math.round(eff.tRep * 10000) / 100;
      document.getElementById('rate-adjust-in-oli').value = Math.round(eff.tOli * 10000) / 100;
      updateRateAdjustModalPreview();
    });

    document.getElementById('form-rate-adjust')?.addEventListener('submit', async (e) => {
      e.preventDefault();
      if (currentProfile?.role !== 'admin') return;
      const sid = document.getElementById('rate-adjust-sale-id')?.value;
      const v = state.ventes.find(x => String(x.id) === String(sid));
      if (!v) return;

      const tRep = Math.round(Math.max(0, (parseFloat(document.getElementById('rate-adjust-in-rep').value) || 0) / 100) * 100000) / 100000;
      const tOli = Math.round(Math.max(0, (parseFloat(document.getElementById('rate-adjust-in-oli').value) || 0) / 100) * 100000) / 100000;
      const val = Number(v.valeur || 0);
      const cRep = Math.round(val * tRep * 100) / 100;
      const cOli = Math.round(val * tOli * 100) / 100;

      const baseR = getRateForSale(v.representant_id, v.entreprise_id, v.client_id);
      const isExactlyBase = baseR && Math.abs(Number(baseR.taux) - tRep) < 0.00005 && Math.abs(Number(baseR.taux_olivier) - tOli) < 0.00005;

      const { error } = await sb.from('ventes').update({
        taux_rep_applique: tRep,
        taux_olivier_applique: tOli,
        commission_rep: cRep,
        commission_olivier: cOli
      }).eq('id', sid);
      if (error) return toast(error.message, 'error');

      await saveSaleBonusOverride(sid, {
        tRep,
        tOli,
        cRep,
        cOli,
        noBonus: Boolean(isExactlyBase),
        customManual: Boolean(!isExactlyBase)
      });

      await recordSaleAudit(
        sid,
        'Taux de commission ajusté',
        `Rep : ${fmtPct(tRep)} (${fmtMoney(cRep)}) • Olivier : ${fmtPct(tOli)} (${fmtMoney(cOli)})`,
        '⚖️'
      );

      document.getElementById('rate-adjust-modal').classList.add('hidden');
      toast(`Taux mis à jour : ${fmtPct(tRep)} Rep / ${fmtPct(tOli)} Olivier ✓`);
      await loadAllData();
    });

    document.getElementById('btn-close-rate-adjust')?.addEventListener('click', () => {
      document.getElementById('rate-adjust-modal').classList.add('hidden');
    });
    document.getElementById('btn-cancel-rate-adjust')?.addEventListener('click', () => {
      document.getElementById('rate-adjust-modal').classList.add('hidden');
    });
    document.getElementById('rate-adjust-modal')?.addEventListener('click', (e) => {
      if (e.target.id === 'rate-adjust-modal') document.getElementById('rate-adjust-modal').classList.add('hidden');
    });

    document.getElementById('btn-close-audit')?.addEventListener('click', () => {
      document.getElementById('audit-modal').classList.add('hidden');
    });
    document.getElementById('audit-modal')?.addEventListener('click', (e) => {
      if (e.target.id === 'audit-modal') document.getElementById('audit-modal').classList.add('hidden');
    });
    document.getElementById('btn-close-trash')?.addEventListener('click', () => {
      document.getElementById('trash-modal').classList.add('hidden');
    });
    document.getElementById('trash-modal')?.addEventListener('click', (e) => {
      if (e.target.id === 'trash-modal') document.getElementById('trash-modal').classList.add('hidden');
    });
    document.getElementById('btn-empty-trash')?.addEventListener('click', async () => {
      if (currentProfile?.role !== 'admin') return;
      const ok = await askConfirm('Vider la corbeille ?', 'Toutes les ventes dans la corbeille seront définitivement supprimées.');
      if (!ok) return;
      const ids = saleTrashStore.map(x => String(x.sale?.id)).filter(Boolean);
      saleTrashStore = [];
      lsSet('korpo_sale_trash_v1', '[]');
      if (sb && ids.length) {
        try {
          await sb.from('kpi_hebdo').delete().in('mois_commission', ids).eq('semaine_index', -3);
        } catch (e) {}
      }
      updateTrashBadges();
      openTrashModal();
      toast('Corbeille vidée');
    });

    function syncGoogleAgendaIntoKpi() {
      rebuildKpiStore();
    }

    let sb = null;
    let currentUser = null;
    let currentProfile = null;
    let selectedSaleIds = new Set();
    let currentFilteredAdminSales = [];
    let state = {
      representants: [],
      entreprises: [],
      clients: [],
      taux: [],
      ventes: [],
      mesVentes: [],
      classement: []
    };

    const fmtMoney = (n) => new Intl.NumberFormat('fr-CA', { style: 'currency', currency: 'CAD' }).format(Number(n || 0));
    const fmtMoney0 = (n) => new Intl.NumberFormat('fr-CA', { style: 'currency', currency: 'CAD', maximumFractionDigits: 0 }).format(Number(n || 0));
    const fmtPct = (n) => `${(Number(n || 0) * 100).toFixed(2)} %`;

    const entBadgeColors = {
      'Toitex': 'bg-sky-50 text-sky-700 border-sky-200/80',
      'MaClim': 'bg-indigo-50 text-indigo-700 border-indigo-200/80',
      'MonSpray': 'bg-emerald-50 text-emerald-700 border-emerald-200/80',
      'VerreVDM': 'bg-purple-50 text-purple-700 border-purple-200/80',
      'Groupe Diamant Pro': 'bg-amber-50 text-amber-800 border-amber-200/80',
      'Ornithorynque': 'bg-slate-100 text-slate-700 border-slate-200'
    };

    function entBadge(nom) {
      const cls = entBadgeColors[nom] || 'bg-slate-100 text-slate-700 border-slate-200';
      return `<span class="inline-flex items-center px-2.5 py-0.5 text-xs font-semibold rounded-lg border ${cls}">${esc(nom)}</span>`;
    }

    function playSaleCelebrationFanfare() {
      try {
        const ctx = sharedAudioCtx || new (window.AudioContext || window.webkitAudioContext)();
        if (ctx.state === 'suspended') ctx.resume();
        const notes = [523.25, 659.25, 783.99, 1046.5];
        notes.forEach((freq, idx) => {
          const osc = ctx.createOscillator();
          const gain = ctx.createGain();
          osc.type = 'triangle';
          const startAt = ctx.currentTime + idx * 0.08;
          osc.frequency.setValueAtTime(freq, startAt);
          gain.gain.setValueAtTime(0.14, startAt);
          gain.gain.exponentialRampToValueAtTime(0.001, startAt + 0.35);
          osc.connect(gain);
          gain.connect(ctx.destination);
          osc.start(startAt);
          osc.stop(startAt + 0.35);
        });
      } catch (e) {}
    }

    function ensureRocketCelebrationStyles() {
      if (document.getElementById('korpo-rocket-celebration-css')) return;
      const st = document.createElement('style');
      st.id = 'korpo-rocket-celebration-css';
      st.textContent = `
        @keyframes korpoRocketThrust {
          0%, 100% { transform: translate(-50%, -50%) rotate(45deg) scale(1); }
          50% { transform: translate(-50%, -53%) rotate(47deg) scale(1.08); }
        }
        @keyframes korpoRocketBlastOff {
          0% { transform: translate(-50%, -50%) rotate(45deg) scale(1.15); opacity: 1; }
          100% { transform: translate(180px, -180px) rotate(45deg) scale(1.45); opacity: 0; }
        }
        @keyframes korpoFlameFlicker {
          0%, 100% { opacity: 0.95; transform: scaleX(1); }
          50% { opacity: 0.55; transform: scaleX(0.75); }
        }
        @keyframes korpoCelebrationPop {
          0% { transform: scale(0.84) translateY(18px); opacity: 0; }
          65% { transform: scale(1.03) translateY(-2px); opacity: 1; }
          100% { transform: scale(1) translateY(0); opacity: 1; }
        }
        @keyframes korpoGlowPulse {
          0%, 100% { box-shadow: 0 0 28px rgba(16, 185, 129, 0.35), 0 0 60px rgba(99, 102, 241, 0.25); }
          50% { box-shadow: 0 0 46px rgba(16, 185, 129, 0.65), 0 0 85px rgba(245, 158, 11, 0.35); }
        }
        .korpo-rocket-card-anim { animation: korpoCelebrationPop 0.45s cubic-bezier(0.22, 1, 0.36, 1) forwards, korpoGlowPulse 2.2s ease-in-out infinite; }
        .korpo-rocket-active { animation: korpoRocketThrust 0.22s ease-in-out infinite; }
        .korpo-rocket-blast { animation: korpoRocketBlastOff 0.65s cubic-bezier(0.16, 1, 0.3, 1) forwards; }
        .korpo-flame-trail { animation: korpoFlameFlicker 0.16s ease-in-out infinite; transform-origin: right center; }
      `;
      document.head.appendChild(st);
    }

    function playRocketIgnitionSweep() {
      try {
        const ctx = sharedAudioCtx || new (window.AudioContext || window.webkitAudioContext)();
        if (ctx.state === 'suspended') ctx.resume();
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        osc.type = 'sawtooth';
        const t0 = ctx.currentTime;
        osc.frequency.setValueAtTime(110, t0);
        osc.frequency.exponentialRampToValueAtTime(520, t0 + 1.25);
        gain.gain.setValueAtTime(0.06, t0);
        gain.gain.exponentialRampToValueAtTime(0.001, t0 + 1.3);
        osc.connect(gain);
        gain.connect(ctx.destination);
        osc.start(t0);
        osc.stop(t0 + 1.3);
      } catch (e) {}
    }

    function triggerSaleCelebration(repName, clientNom, entNom, valeur, commAmount, isTeamBroadcast = false) {
      ensureRocketCelebrationStyles();
      let overlay = document.getElementById('sale-celebration-overlay');
      if (!overlay) {
        overlay = document.createElement('div');
        overlay.id = 'sale-celebration-overlay';
        overlay.className = 'fixed inset-0 bg-slate-950/75 backdrop-blur-sm flex items-center justify-center p-4';
        document.body.appendChild(overlay);
      } else if (overlay.parentNode !== document.body) {
        document.body.appendChild(overlay);
      }

      const safeVal = Math.max(0, Number(valeur || 0));
      const safeComm = Number(commAmount || 0);
      const showCommRow = !isTeamBroadcast && Math.abs(safeComm) > 0.009;
      const clientLine = clientNom && clientNom !== 'Client confidentiel' ? ` — ${clientNom}` : '';
      const headline = isTeamBroadcast
        ? `🚀 ${repName || 'Un membre'} vient de closer un deal !`
        : `🚀 DEAL CONCLU, ${String(repName || '').toUpperCase()} !`;

      overlay.style.cssText = 'position:fixed;inset:0;z-index:2147483647;display:flex;align-items:center;justify-content:center;padding:1rem;background:rgba(2,6,23,0.78);backdrop-filter:blur(6px);';
      overlay.innerHTML = `
        <div class="korpo-rocket-card-anim relative w-full max-w-md mx-4 rounded-3xl bg-gradient-to-br from-slate-950 via-indigo-950 to-slate-900 border-2 border-emerald-400/50 p-6 sm:p-7 text-center text-white shadow-2xl overflow-hidden">
          <div class="absolute -top-24 -right-24 w-56 h-56 rounded-full bg-emerald-500/20 blur-3xl pointer-events-none"></div>
          <div class="absolute -bottom-24 -left-24 w-56 h-56 rounded-full bg-amber-400/15 blur-3xl pointer-events-none"></div>

          <div class="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-emerald-500/20 border border-emerald-400/40 text-emerald-300 text-[11px] font-extrabold uppercase tracking-wider mb-3">
            <span>🔥</span><span>Nouvelle vente enregistrée</span>
          </div>

          <h3 id="celebration-title" class="text-xl sm:text-2xl font-black tracking-tight text-white">${esc(headline)}</h3>
          <p id="celebration-subtitle" class="text-xs sm:text-sm font-semibold text-indigo-200 mt-1">${esc((entNom || 'Korpo') + clientLine)}</p>

          <!-- PISTE DE DÉCOLLAGE DE LA FUSÉE -->
          <div class="mt-6 mb-5 px-2">
            <div class="flex items-center justify-between text-[11px] font-extrabold uppercase tracking-wider mb-2">
              <span id="rocket-stage-label" class="text-amber-300">Allumage des moteurs...</span>
              <span id="rocket-pct-label" class="text-emerald-300 tabular-nums">0 %</span>
            </div>
            <div class="relative h-5 rounded-full bg-slate-800/95 border border-white/15 p-1 overflow-visible shadow-inner">
              <div id="rocket-progress-fill" class="h-full rounded-full bg-gradient-to-r from-indigo-500 via-amber-400 to-emerald-400 transition-none" style="width: 4%;"></div>
              <!-- Cible d'arrivée -->
              <span class="absolute right-2 top-1/2 -translate-y-1/2 text-xs select-none" title="Objectif atteint">🏁</span>
              <!-- Fusée animée -->
              <div id="rocket-ship-runner" class="korpo-rocket-active absolute top-1/2 select-none pointer-events-none flex items-center" style="left: 4%;">
                <span class="korpo-flame-trail text-base -mr-1.5">🔥</span>
                <span class="text-2xl sm:text-3xl drop-shadow-[0_0_12px_rgba(251,191,36,0.9)]">🚀</span>
              </div>
            </div>
          </div>

          <!-- COMPTEURS DYNAMIQUES MONTANT + COMMISSION -->
          <div class="grid ${showCommRow ? 'grid-cols-2' : 'grid-cols-1'} gap-3 mt-4">
            <div class="rounded-2xl bg-white/10 border border-white/15 p-3.5 backdrop-blur-xs">
              <div class="text-[10.5px] font-bold uppercase tracking-wider text-indigo-200">Montant du deal</div>
              <div id="celebration-amount" class="text-xl sm:text-2xl font-black text-white tabular-nums mt-0.5">${fmtMoney(0)}</div>
            </div>
            ${showCommRow ? `
              <div class="rounded-2xl bg-emerald-500/20 border border-emerald-400/45 p-3.5 backdrop-blur-xs">
                <div class="text-[10.5px] font-bold uppercase tracking-wider text-emerald-200">Ta commission</div>
                <div id="celebration-comm" class="text-xl sm:text-2xl font-black text-emerald-300 tabular-nums mt-0.5">+ ${fmtMoney(0)}</div>
              </div>
            ` : ''}
          </div>

          <p id="rocket-final-quote" class="opacity-0 transition-opacity duration-300 text-xs font-bold text-amber-300 mt-4">
            💥 Boom ! Dossier propulsé dans tes statistiques !
          </p>

          <button type="button" id="btn-close-celebration" class="mt-4 w-full py-2.5 px-4 rounded-xl bg-emerald-500 hover:bg-emerald-400 text-slate-950 font-black text-xs uppercase tracking-wider shadow-lg transition cursor-pointer">
            Continuer sur ma lancée 🔥
          </button>
        </div>
      `;

      overlay.style.display = 'flex';
      overlay.classList.remove('hidden');
      playRocketIgnitionSweep();

      const fillEl = document.getElementById('rocket-progress-fill');
      const shipEl = document.getElementById('rocket-ship-runner');
      const pctEl = document.getElementById('rocket-pct-label');
      const stageEl = document.getElementById('rocket-stage-label');
      const amtEl = document.getElementById('celebration-amount');
      const commEl = document.getElementById('celebration-comm');
      const quoteEl = document.getElementById('rocket-final-quote');

      const durationMs = 1350;
      const startTime = performance.now();
      let fanfarePlayed = false;

      if (window._rocketAnimFrame) cancelAnimationFrame(window._rocketAnimFrame);

      const spawnConfettiBurst = () => {
        const emojis = ['🚀', '💵', '💰', '🔥', '🎉', '⭐', '💎'];
        for (let i = 0; i < 34; i++) {
          const p = document.createElement('div');
          p.className = 'celebration-particle text-2xl select-none';
          p.style.zIndex = '10000';
          p.textContent = emojis[i % emojis.length];
          p.style.left = `${Math.random() * 96 + 2}vw`;
          p.style.animationDuration = `${1.7 + Math.random() * 1.4}s`;
          p.style.animationDelay = `${Math.random() * 0.25}s`;
          document.body.appendChild(p);
          setTimeout(() => p.remove(), 3500);
        }
      };

      const stepAnimation = (now) => {
        if (overlay.classList.contains('hidden')) return;
        const rawT = Math.min(1, (now - startTime) / durationMs);
        // Courbe d'accélération progressive façon fusée (ease-in-out cubic)
        const eased = rawT < 0.5
          ? 2 * rawT * rawT
          : 1 - Math.pow(-2 * rawT + 2, 2) / 2;

        const pct = Math.round(eased * 100);
        const trackPct = Math.min(96, Math.max(4, Math.round(eased * 96)));
        if (fillEl) fillEl.style.width = `${Math.max(4, pct)}%`;
        if (shipEl) shipEl.style.left = `${trackPct}%`;
        if (pctEl) pctEl.textContent = `${pct} %`;
        if (amtEl) amtEl.textContent = fmtMoney(safeVal * eased);
        if (commEl) commEl.textContent = `${safeComm >= 0 ? '+ ' : ''}${fmtMoney(safeComm * eased)}`;

        if (stageEl) {
          if (pct < 35) stageEl.textContent = '🔥 Allumage des propulseurs...';
          else if (pct < 75) stageEl.textContent = '⚡ Pleine puissance...';
          else if (pct < 100) stageEl.textContent = '🚀 Approche de la cible...';
          else stageEl.textContent = '🏆 OBJECTIF ATTEINT !';
        }

        if (rawT < 1) {
          window._rocketAnimFrame = requestAnimationFrame(stepAnimation);
        } else if (!fanfarePlayed) {
          fanfarePlayed = true;
          if (amtEl) amtEl.textContent = fmtMoney(safeVal);
          if (commEl) commEl.textContent = `${safeComm >= 0 ? '+ ' : ''}${fmtMoney(safeComm)}`;
          if (quoteEl) quoteEl.classList.remove('opacity-0');
          if (shipEl) {
            shipEl.classList.remove('korpo-rocket-active');
            shipEl.classList.add('korpo-rocket-blast');
          }
          playSaleCelebrationFanfare();
          spawnConfettiBurst();
        }
      };

      window._rocketAnimFrame = requestAnimationFrame(stepAnimation);

      const closeOverlay = () => {
        if (window._rocketAnimFrame) cancelAnimationFrame(window._rocketAnimFrame);
        overlay.style.cssText = 'display:none;';
        overlay.classList.add('hidden');
      };
      const closeBtn = document.getElementById('btn-close-celebration');
      if (closeBtn) closeBtn.onclick = closeOverlay;
      overlay.onclick = (e) => { if (e.target === overlay) closeOverlay(); };
      clearTimeout(window._celebrationTimer);
      window._celebrationTimer = setTimeout(closeOverlay, 5600);
    }

    function toast(msg, type = 'success') {
      const container = document.getElementById('toast-container');
      const el = document.createElement('div');
      const colors = type === 'error'
        ? 'bg-rose-600 text-white shadow-rose-600/20'
        : 'bg-slate-900 text-white shadow-slate-900/20';
      el.className = `${colors} text-xs font-semibold px-4 py-3 rounded-xl shadow-xl transition-opacity flex items-center gap-2`;
      el.innerHTML = `<span>${type === 'error' ? '⚠' : '✓'}</span><span>${esc(msg)}</span>`;
      container.appendChild(el);
      setTimeout(() => { el.style.opacity = '0'; setTimeout(() => el.remove(), 300); }, 3500);
    }

    function askConfirm(title, message) {
      return new Promise((resolve) => {
        const modal = document.getElementById('confirm-modal');
        if (modal && modal.parentNode !== document.body) {
          document.body.appendChild(modal);
        }
        if (modal) {
          modal.style.zIndex = '9999';
        }
        document.getElementById('confirm-title').textContent = title;
        document.getElementById('confirm-message').textContent = message;
        modal.classList.remove('hidden');
        const cleanup = (val) => {
          modal.classList.add('hidden');
          document.getElementById('confirm-ok-btn').onclick = null;
          document.getElementById('confirm-cancel-btn').onclick = null;
          resolve(val);
        };
        document.getElementById('confirm-ok-btn').onclick = () => cleanup(true);
        document.getElementById('confirm-cancel-btn').onclick = () => cleanup(false);
      });
    }

    function showScreen(name) {
      ['screen-config', 'screen-login', 'screen-app'].forEach(id => {
        document.getElementById(id).classList.toggle('hidden', id !== name);
      });
      const isAppActive = name === 'screen-app' && Boolean(currentProfile);
      document.getElementById('btn-mobile-fab-sale')?.classList.toggle('hidden', !isAppActive);
      document.getElementById('btn-floating-chat-bubble')?.classList.toggle('hidden', !isAppActive);
      if (!isAppActive) {
        closeTeamChatModal();
        document.getElementById('sale-drawer-backdrop')?.classList.add('hidden');
      }
    }

    function getCredentials() {
      const url = (SUPABASE_URL && !SUPABASE_URL.includes('VOTRE_'))
        ? SUPABASE_URL
        : localStorage.getItem('korpo_sb_url');
      const key = (SUPABASE_ANON_KEY && !SUPABASE_ANON_KEY.includes('VOTRE_'))
        ? SUPABASE_ANON_KEY
        : localStorage.getItem('korpo_sb_key');
      return { url, key };
    }

    function syncThemeToggleButtons() {
      const isDark = document.documentElement.classList.contains('dark');
      document.querySelectorAll('.theme-toggle-icon').forEach(el => {
        el.textContent = isDark ? '☀️' : '🌙';
      });
      document.querySelectorAll('.theme-toggle-label').forEach(el => {
        el.textContent = isDark ? 'Clair' : 'Sombre';
      });
      document.querySelectorAll('.theme-menu-text').forEach(el => {
        el.textContent = isDark ? 'Passer en mode clair' : 'Passer en mode sombre';
      });
      const metaTheme = document.querySelector('meta[name="theme-color"]');
      if (metaTheme) metaTheme.setAttribute('content', isDark ? '#090d16' : '#0f172a');
    }

    window.toggleColorTheme = () => {
      const isDark = document.documentElement.classList.toggle('dark');
      lsSet('korpo_theme_v1', isDark ? 'dark' : 'light');
      syncThemeToggleButtons();
      initMentionPushServiceWorker();
    };

    // =========================================================================
    // NOTIFICATIONS PUSH POUR LES MENTIONS (@) DANS KORPO TOP EARNER
    // =========================================================================
    // Clé publique VAPID (à remplacer par votre clé publique VAPID générée pour le push en arrière-plan)
    const KORPO_VAPID_PUBLIC_KEY = "BFqGwmm2MtWS26Db2ro4fAcLn7o9pIMRbjTS4GerOgric-YH4_uZLzftNRO9hs4pMba8RIdWz0DwSfx2W0c16IY";
    let korpoSwRegistration = null;

    function urlBase64ToUint8Array(base64String) {
      const padding = '='.repeat((4 - (base64String.length % 4)) % 4);
      const base64 = (base64String + padding).replace(/-/g, '+').replace(/_/g, '/');
      const rawData = window.atob(base64);
      const outputArray = new Uint8Array(rawData.length);
      for (let i = 0; i < rawData.length; ++i) outputArray[i] = rawData.charCodeAt(i);
      return outputArray;
    }

    function isMentionPushUserDisabled() {
      return localStorage.getItem('korpo_mention_push_disabled_v1') === '1';
    }

    function isMentionPushCurrentlyEnabled() {
      if (!('Notification' in window)) return false;
      return Notification.permission === 'granted' && !isMentionPushUserDisabled();
    }

    async function initMentionPushServiceWorker() {
      if ('serviceWorker' in navigator) {
        try {
          korpoSwRegistration = await navigator.serviceWorker.register('./sw.js', { updateViaCache: 'none' });
          if (korpoSwRegistration && typeof korpoSwRegistration.update === 'function') {
            korpoSwRegistration.update().catch(() => {});
          }
        } catch (e) {
          console.warn('Service Worker non enregistré :', e && e.message);
        }
      }
      updateMentionPushBtnUI();
    }

    function updateMentionPushBtnUI() {
      const menuIcon = document.getElementById('menu-push-icon');
      const menuBadge = document.getElementById('menu-push-badge');
      const supported = ('Notification' in window);
      const active = isMentionPushCurrentlyEnabled();

      if (menuIcon) menuIcon.textContent = active ? '🔔' : '🔕';
      if (menuBadge) {
        if (!supported) {
          menuBadge.textContent = 'N/D';
          menuBadge.className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-slate-100 text-slate-500';
        } else if (active) {
          menuBadge.textContent = 'Actif';
          menuBadge.className = 'text-[10px] font-extrabold px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-800';
        } else {
          menuBadge.textContent = 'Désactivé';
          menuBadge.className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-slate-100 text-slate-600';
        }
      }
    }

    window.disableMentionPushNotifications = async () => {
      try {
        localStorage.setItem('korpo_mention_push_disabled_v1', '1');
      } catch (e) {}

      if (!korpoSwRegistration && 'serviceWorker' in navigator) {
        try { korpoSwRegistration = await navigator.serviceWorker.getRegistration('./sw.js') || await navigator.serviceWorker.ready; } catch (e) {}
      }

      if (korpoSwRegistration && 'PushManager' in window) {
        try {
          const sub = await korpoSwRegistration.pushManager.getSubscription();
          if (sub) {
            const endpointToRemove = sub.endpoint;
            await sub.unsubscribe();
            if (sb && endpointToRemove) {
              await sb.from('push_subscriptions').delete().eq('endpoint', endpointToRemove);
            }
          }
        } catch (err) {
          console.warn('Désabonnement PushManager:', err && err.message);
        }
      }

      updateMentionPushBtnUI();
      toast('Notifications push désactivées sur cet appareil 🔕');
    };

    window.enableMentionPushNotifications = async () => {
      if (!('Notification' in window)) {
        return toast('Les notifications ne sont pas supportées sur ce navigateur.', 'error');
      }
      const perm = await Notification.requestPermission();
      if (perm !== 'granted') {
        updateMentionPushBtnUI();
        return toast('Autorisez les notifications dans les réglages de votre téléphone / navigateur.', 'error');
      }

      try {
        localStorage.removeItem('korpo_mention_push_disabled_v1');
      } catch (e) {}

      if (!korpoSwRegistration && 'serviceWorker' in navigator) {
        try { korpoSwRegistration = await navigator.serviceWorker.register('./sw.js'); } catch (e) {}
      }

      if (korpoSwRegistration && 'PushManager' in window && KORPO_VAPID_PUBLIC_KEY && sb && currentProfile) {
        try {
          let sub = await korpoSwRegistration.pushManager.getSubscription();
          if (!sub) {
            sub = await korpoSwRegistration.pushManager.subscribe({
              userVisibleOnly: true,
              applicationServerKey: urlBase64ToUint8Array(KORPO_VAPID_PUBLIC_KEY)
            });
          }
          await sb.from('push_subscriptions').upsert({
            representant_id: currentProfile.id,
            representant_nom: currentProfile.nom,
            endpoint: sub.endpoint,
            subscription: sub.toJSON(),
            updated_at: new Date().toISOString()
          }, { onConflict: 'endpoint' });
        } catch (err) {
          console.warn('Abonnement PushManager:', err && err.message);
        }
      }

      updateMentionPushBtnUI();
      toast('Notifications activées ! Vous recevrez une alerte dès qu’on vous mentionne (@) 🔔');
    };

    window.toggleMentionPushNotifications = async () => {
      if (isMentionPushCurrentlyEnabled()) {
        await disableMentionPushNotifications();
      } else {
        await enableMentionPushNotifications();
      }
    };

    function sendChatMentionSystemNotification(msgObj) {
      if (!isMentionPushCurrentlyEnabled()) return;
      const title = `💬 ${msgObj.authorName} vous a mentionné`;
      const body = String(msgObj.text || '').slice(0, 180);
      const tag = `korpo-mention-${msgObj.id}`;

      if (korpoSwRegistration && typeof korpoSwRegistration.showNotification === 'function') {
        korpoSwRegistration.showNotification(title, {
          body,
          tag,
          renotify: true,
          data: { msgId: msgObj.id, url: './' }
        }).catch(() => {});
      } else {
        try {
          const n = new Notification(title, { body, tag });
          n.onclick = () => {
            window.focus();
            jumpToChatMessage(msgObj.id);
            n.close();
          };
        } catch (e) {}
      }
    }

    async function triggerServerMentionWebPush(chatObj, mentionedNames) {
      if (!sb || !Array.isArray(mentionedNames) || !mentionedNames.length) return;
      try {
        await sb.functions.invoke('send-mention-push', {
          body: {
            msgId: chatObj.id,
            authorId: chatObj.authorId,
            authorName: chatObj.authorName,
            text: chatObj.text,
            mentionedNames
          }
        });
      } catch (e) {
        // Silencieux tant que la Edge Function n'est pas déployée sur Supabase
      }
    }

    async function triggerAdminSaleWebPush(chatObj) {
      if (!sb || !chatObj || !chatObj.saleData) return;
      // N'envoyer un push admin que si ce n'est pas l'admin lui-même qui vient d'enregistrer sa propre vente
      if (currentProfile?.role === 'admin' && String(chatObj.authorId) === String(currentProfile.id)) return;
      const adminNames = (state.representants || [])
        .filter(r => r.role === 'admin' && r.nom)
        .map(r => r.nom);
      if (!adminNames.includes('Olivier')) adminNames.push('Olivier');

      const s = chatObj.saleData;
      const cliPart = s.clientNom ? ` (${s.clientNom})` : '';
      const saleSummaryText = `${s.repNom} • ${s.entNom}${cliPart} — ${fmtMoney(s.valeur)} [${s.typeKpi || 'OTS'}]`;

      try {
        await sb.functions.invoke('send-mention-push', {
          body: {
            msgId: chatObj.id,
            authorId: chatObj.authorId,
            authorName: chatObj.authorName,
            text: saleSummaryText,
            mentionedNames: adminNames,
            customTitle: `💵 Nouvelle vente — ${s.repNom}`,
            eventType: 'sale'
          }
        });
      } catch (e) {
        // Silencieux tant que la Edge Function n'est pas déployée sur Supabase
      }
    }

    // =========================================================================
    // TÉLÉPHONE ROUGE — APPEL DIRECTEUR URGENT POUR CLOSER UN DEAL
    // =========================================================================
    let redPhoneConfig = lsGetJson('korpo_red_phone_cfg_v1', '{"directorName":"Olivier (Directeur)","phone":""}');
    let selectedRedPhoneReason = 'Client prêt à closer maintenant';
    let seenRedPhoneCallIds = new Set(lsGetJson('korpo_seen_red_calls_v1', '[]'));
    let notifiedRedPhoneCallIds = new Set(lsGetJson('korpo_notified_red_calls_v1', '[]'));
    let activeIncomingRedCallId = null;

    function formatPhoneForDial(rawPhone) {
      const digits = String(rawPhone || '').replace(/[^\d+]/g, '');
      if (!digits) return '';
      if (digits.startsWith('+')) return digits;
      if (digits.length === 10) return `+1${digits}`;
      if (digits.length === 11 && digits.startsWith('1')) return `+${digits}`;
      return digits;
    }

    function isRedPhoneWithinAllowedHours() {
      const now = new Date();
      const day = now.getDay(); // 0 = Dimanche, 6 = Samedi, 1..5 = Lundi..Vendredi
      const nowMins = now.getHours() * 60 + now.getMinutes();
      const isWeekend = (day === 0 || day === 6);
      const startMins = 8 * 60; // 08h00
      const endMins = isWeekend ? (12 * 60) : (18 * 60); // 12h00 fin de semaine, 18h00 semaine
      return nowMins >= startMins && nowMins <= endMins;
    }

    function formatRedPhoneScheduleLabel() {
      return 'Lun–Ven 8h à 18h • Sam–Dim 8h à 12h';
    }

    function syncRedPhoneUI() {
      const isRep = currentProfile && currentProfile.role !== 'admin';
      const isOpenNow = isRedPhoneWithinAllowedHours();
      const hdrBtn = document.getElementById('btn-header-red-phone');
      if (hdrBtn) {
        hdrBtn.classList.toggle('hidden', !isRep);
        hdrBtn.classList.toggle('opacity-55', isRep && !isOpenNow);
        hdrBtn.title = isOpenNow
          ? 'Téléphone Rouge — Appel Directeur'
          : `Téléphone Rouge (Fermé — ${formatRedPhoneScheduleLabel()})`;
      }
      const nameIn = document.getElementById('red-phone-director-name');
      const phoneIn = document.getElementById('red-phone-director-number');
      if (nameIn && document.activeElement !== nameIn) nameIn.value = redPhoneConfig.directorName || 'Olivier (Directeur)';
      if (phoneIn && document.activeElement !== phoneIn) phoneIn.value = redPhoneConfig.phone || '';

      const previewEl = document.getElementById('red-phone-number-preview');
      if (previewEl) {
        previewEl.textContent = redPhoneConfig.phone
          ? `${redPhoneConfig.directorName || 'Directeur'} • ${redPhoneConfig.phone}`
          : 'Numéro non encore configuré par le directeur';
      }
      const hoursEl = document.getElementById('red-phone-hours-preview');
      if (hoursEl) {
        hoursEl.textContent = formatRedPhoneScheduleLabel();
      }
      const closedNotice = document.getElementById('red-phone-closed-notice');
      const callBtn = document.getElementById('btn-trigger-red-phone-now');
      if (closedNotice) closedNotice.classList.toggle('hidden', isOpenNow);
      if (callBtn) {
        callBtn.disabled = !isOpenNow;
        callBtn.classList.toggle('opacity-45', !isOpenNow);
        callBtn.classList.toggle('cursor-not-allowed', !isOpenNow);
        callBtn.classList.toggle('cursor-pointer', isOpenNow);
      }
    }

    window.selectRedPhoneReason = (btnEl, reasonText) => {
      selectedRedPhoneReason = reasonText;
      document.querySelectorAll('.red-reason-chip').forEach(b => {
        b.className = 'red-reason-chip px-3 py-1.5 text-xs font-bold rounded-xl bg-slate-100 text-slate-700 hover:bg-rose-50 border border-slate-200 transition cursor-pointer';
      });
      if (btnEl) {
        btnEl.className = 'red-reason-chip px-3 py-1.5 text-xs font-bold rounded-xl bg-rose-600 text-white border border-rose-600 transition cursor-pointer';
      }
    };

    let _redPhoneCountdownTimer = null;
    let _redPhonePendingDialNumber = '';

    function clearRedPhoneCountdown() {
      if (_redPhoneCountdownTimer) {
        clearInterval(_redPhoneCountdownTimer);
        _redPhoneCountdownTimer = null;
      }
      _redPhonePendingDialNumber = '';
    }

    window.dialRedPhoneNowImmediately = () => {
      const targetNum = _redPhonePendingDialNumber || formatPhoneForDial(redPhoneConfig.phone);
      clearRedPhoneCountdown();
      document.getElementById('red-phone-rep-modal')?.classList.add('hidden');
      if (targetNum) {
        window.location.href = `tel:${targetNum}`;
      }
    };

    window.openRedPhoneRepModal = () => {
      clearRedPhoneCountdown();
      syncRedPhoneUI();
      const noteEl = document.getElementById('red-phone-custom-note');
      if (noteEl) noteEl.value = '';
      document.getElementById('red-phone-step-compose')?.classList.remove('hidden');
      document.getElementById('red-phone-step-countdown')?.classList.add('hidden');
      document.getElementById('red-phone-rep-modal')?.classList.remove('hidden');
    };

    window.closeRedPhoneRepModal = () => {
      clearRedPhoneCountdown();
      document.getElementById('red-phone-rep-modal')?.classList.add('hidden');
    };

    function playRedPhoneRingtone() {
      try {
        const ctx = sharedAudioCtx || new (window.AudioContext || window.webkitAudioContext)();
        if (ctx.state === 'suspended') ctx.resume();
        const pattern = [0, 0.16, 0.32, 0.65, 0.81, 0.97];
        pattern.forEach((offset, idx) => {
          const osc = ctx.createOscillator();
          const gain = ctx.createGain();
          osc.type = 'triangle';
          const startAt = ctx.currentTime + offset;
          osc.frequency.setValueAtTime(idx % 2 === 0 ? 987.77 : 1318.51, startAt);
          gain.gain.setValueAtTime(0.22, startAt);
          gain.gain.exponentialRampToValueAtTime(0.001, startAt + 0.14);
          osc.connect(gain);
          gain.connect(ctx.destination);
          osc.start(startAt);
          osc.stop(startAt + 0.14);
        });
      } catch (e) {}
    }

    window.triggerRedPhoneDirectorCall = async () => {
      if (!currentProfile) return;
      if (!isRedPhoneWithinAllowedHours()) {
        syncRedPhoneUI();
        return toast(`⏰ Le Téléphone Rouge est disponible uniquement selon l'horaire : ${formatRedPhoneScheduleLabel()}.`, 'error');
      }
      const dialNumber = formatPhoneForDial(redPhoneConfig.phone);
      if (!dialNumber) {
        return toast('Le numéro du directeur n’est pas encore configuré. Demandez-lui de l’ajouter dans Entreprises & Équipe.', 'error');
      }

      const customNote = document.getElementById('red-phone-custom-note')?.value.trim() || '';
      const reason = customNote || selectedRedPhoneReason || 'Besoin d’aide pour closer un deal';
      const callId = `redcall_${Date.now()}_${currentProfile.id}`;
      const nowIso = new Date().toISOString();

      const payloadObj = {
        id: callId,
        at: nowIso,
        repId: String(currentProfile.id),
        repNom: currentProfile.nom,
        reason,
        status: 'ringing'
      };

      // Basculer immédiatement vers l'écran de décompte de 15 secondes
      clearRedPhoneCountdown();
      _redPhonePendingDialNumber = dialNumber;
      let secondsLeft = 15;
      const timerEl = document.getElementById('red-phone-countdown-timer');
      const barEl = document.getElementById('red-phone-countdown-bar');
      if (timerEl) timerEl.textContent = `${secondsLeft} s`;
      if (barEl) barEl.style.width = '100%';
      document.getElementById('red-phone-step-compose')?.classList.add('hidden');
      document.getElementById('red-phone-step-countdown')?.classList.remove('hidden');

      _redPhoneCountdownTimer = setInterval(() => {
        secondsLeft -= 1;
        if (timerEl) timerEl.textContent = `${Math.max(0, secondsLeft)} s`;
        if (barEl) barEl.style.width = `${Math.max(0, Math.round((secondsLeft / 15) * 100))}%`;
        if (secondsLeft <= 0) {
          dialRedPhoneNowImmediately();
        }
      }, 1000);

      // 1) Enregistrer l'alerte urgente en temps réel dans Supabase (semaine_index = -15) dès 0s
      if (sb) {
        try {
          await safeUpsertKpiHebdoRow({
            mois_commission: callId,
            semaine_index: -15,
            representant_nom: JSON.stringify(payloadObj),
            booke: 0,
            rdv_realise: 0,
            updated_at: nowIso
          });
        } catch (e) {}

        // 2) Déclencher immédiatement la notification Push serveur vers le téléphone du directeur
        try {
          const adminNames = (state.representants || [])
            .filter(r => r.role === 'admin' && r.nom)
            .map(r => r.nom);
          if (!adminNames.includes('Olivier')) adminNames.push('Olivier');
          await sb.functions.invoke('send-mention-push', {
            body: {
              msgId: callId,
              authorId: String(currentProfile.id),
              authorName: currentProfile.nom,
              text: `${currentProfile.nom} t'appelle dans 15s pour un deal (${reason}) !`,
              mentionedNames: adminNames,
              customTitle: `☎️ TÉLÉPHONE ROUGE — ${currentProfile.nom}`,
              eventType: 'red_phone'
            }
          });
        } catch (ePush) {}
      }
    };

    window.acknowledgeIncomingRedPhone = () => {
      if (activeIncomingRedCallId) {
        seenRedPhoneCallIds.add(String(activeIncomingRedCallId));
        lsSet('korpo_seen_red_calls_v1', JSON.stringify(Array.from(seenRedPhoneCallIds).slice(-60)));
        activeIncomingRedCallId = null;
      }
      document.getElementById('red-phone-incoming-modal')?.classList.add('hidden');
      renderUnifiedBellNotifications();
    };

    function getRecentRedPhoneCallsForAdmin() {
      if (!currentProfile || currentProfile.role !== 'admin') return [];
      const cutoffMs = Date.now() - 24 * 3600 * 1000;
      const calls = [];
      (kpiHebdoRows || []).forEach(r => {
        if (Number(r.semaine_index) !== -15 || !r.representant_nom) return;
        try {
          const c = JSON.parse(r.representant_nom);
          if (c && c.id && c.at && new Date(c.at).getTime() >= cutoffMs) {
            calls.push(c);
          }
        } catch (e) {}
      });
      return calls.sort((a, b) => String(b.at || '').localeCompare(String(a.at || '')));
    }

    function checkIncomingRedPhoneAlerts() {
      if (!currentProfile || currentProfile.role !== 'admin') return;
      const calls = getRecentRedPhoneCallsForAdmin();
      const freshCutoff = Date.now() - 4 * 60 * 1000; // Alerte modale active pour les appels des 4 dernières minutes
      const unacknowledged = calls.find(c => !seenRedPhoneCallIds.has(String(c.id)) && new Date(c.at).getTime() >= freshCutoff);
      if (!unacknowledged) return;

      const callIdStr = String(unacknowledged.id);
      activeIncomingRedCallId = callIdStr;
      document.getElementById('red-phone-incoming-title').textContent = `☎️ ${unacknowledged.repNom} vous appelle !`;
      document.getElementById('red-phone-incoming-subtitle').textContent = `${unacknowledged.repNom} a besoin de vous MAINTENANT pour closer un deal`;
      document.getElementById('red-phone-incoming-meta').innerHTML = `
        <div><strong>Motif :</strong> ${esc(unacknowledged.reason || 'Aide closing deal')}</div>
        <div class="text-[11px] opacity-80">Déclenché à ${esc(formatChatHourOnly(unacknowledged.at))}</div>
      `;
      document.getElementById('red-phone-incoming-modal')?.classList.remove('hidden');

      // Ne jouer la sonnerie qu'UNE SEULE fois par appel (la notification Push serveur avec la raison est déjà envoyée par Supabase)
      if (!notifiedRedPhoneCallIds.has(callIdStr)) {
        notifiedRedPhoneCallIds.add(callIdStr);
        lsSet('korpo_notified_red_calls_v1', JSON.stringify(Array.from(notifiedRedPhoneCallIds).slice(-60)));
        playRedPhoneRingtone();
      }
    }

    document.getElementById('form-red-phone-config')?.addEventListener('submit', async (e) => {
      e.preventDefault();
      if (currentProfile?.role !== 'admin') return;
      const directorName = document.getElementById('red-phone-director-name')?.value.trim() || 'Olivier (Directeur)';
      const phone = document.getElementById('red-phone-director-number')?.value.trim() || '';
      if (!formatPhoneForDial(phone)) {
        return toast('Veuillez saisir un numéro de téléphone valide.', 'error');
      }
      redPhoneConfig = { directorName, phone, updatedAt: new Date().toISOString() };
      lsSet('korpo_red_phone_cfg_v1', JSON.stringify(redPhoneConfig));
      syncRedPhoneUI();

      if (sb) {
        const res = await safeUpsertKpiHebdoRow({
          mois_commission: 'red_phone_cfg',
          semaine_index: -14,
          representant_nom: JSON.stringify(redPhoneConfig),
          booke: 0,
          rdv_realise: 0,
          updated_at: new Date().toISOString()
        });
        if (res && res.error) {
          return toast('Erreur sauvegarde réseau : ' + res.error.message, 'error');
        }
      }
      toast(`Numéro Téléphone Rouge enregistré (${phone}) ✓`);
    });

    let _startupSplashDismissed = false;
    let _startupSafetyTimer = null;

    function showStartupSplash(title = 'Bienvenue sur votre portail', subtitle = 'Connexion sécurisée en cours...') {
      _startupSplashDismissed = false;
      clearTimeout(_startupSafetyTimer);
      clearTimeout(window._korpoSplashWatchdog);
      const el = document.getElementById('korpo-startup-splash');
      if (el) {
        el.style.display = 'flex';
        el.classList.remove('hidden');
        el.style.opacity = '1';
        el.style.pointerEvents = 'auto';
      }
      updateStartupSplash(28, title, subtitle, 'Vérification de la session...');
      _startupSafetyTimer = setTimeout(() => hideStartupSplash(0), 4000);
    }

    function updateStartupSplash(pct, title, subtitle, stepText) {
      if (_startupSplashDismissed) return;
      const bar = document.getElementById('startup-splash-bar');
      const tEl = document.getElementById('startup-splash-title');
      const sEl = document.getElementById('startup-splash-subtitle');
      const stEl = document.getElementById('startup-splash-step');
      if (bar && typeof pct === 'number') bar.style.width = `${Math.min(100, Math.max(5, pct))}%`;
      if (tEl && title) tEl.textContent = title;
      if (sEl && subtitle) sEl.textContent = subtitle;
      if (stEl && stepText) stEl.textContent = stepText;
    }

    function hideStartupSplash(delayMs = 150) {
      _startupSplashDismissed = true;
      clearTimeout(_startupSafetyTimer);
      clearTimeout(window._korpoSplashWatchdog);
      const bar = document.getElementById('startup-splash-bar');
      const stEl = document.getElementById('startup-splash-step');
      if (bar) bar.style.width = '100%';
      if (stEl) stEl.textContent = 'Portail prêt ✓';
      setTimeout(() => {
        if (typeof window._korpoForceHideSplash === 'function') {
          window._korpoForceHideSplash();
        } else {
          const el = document.getElementById('korpo-startup-splash');
          if (!el) return;
          el.style.opacity = '0';
          el.style.pointerEvents = 'none';
          setTimeout(() => { el.style.display = 'none'; el.classList.add('hidden'); }, 300);
        }
      }, delayMs);
    }

    _startupSafetyTimer = setTimeout(() => hideStartupSplash(0), 4000);

    async function initApp() {
      ['comm-bar-fill', 'goal-bar-fill'].forEach(fid => {
        const el = document.getElementById(fid);
        if (el) syncRocketOnProgressBar(el, parseFloat(el.style.width) || 0);
      });
      syncThemeToggleButtons();
      initKpiControls();
      const { url, key } = getCredentials();
      if (!url || !key) {
        hideStartupSplash(0);
        showScreen('screen-config');
        return;
      }

      restoreRememberedCredentials();
      sb = window.supabase.createClient(url, key, {
        auth: {
          persistSession: true,
          autoRefreshToken: true,
          detectSessionInUrl: true
        }
      });

      const { data: { session } } = await sb.auth.getSession();
      if (session?.user) {
        updateStartupSplash(38, null, 'Profil détecté, préparation de votre espace...', 'Chargement du profil...');
        await handleAuthenticatedUser(session.user);
      } else {
        hideStartupSplash(0);
        showScreen('screen-login');
      }

      // Pas d'appels Supabase « await » directement dans ce callback (risque de blocage connu de supabase-js v2)
      sb.auth.onAuthStateChange((event, session) => {
        if (event === 'SIGNED_IN' && session?.user) {
          if (currentProfile && currentUser && currentUser.id === session.user.id) return; // déjà chargé (retour sur l'onglet)
          setTimeout(() => handleAuthenticatedUser(session.user), 0);
        } else if (event === 'PASSWORD_RECOVERY') {
          // Lien de réinitialisation reçu par courriel : ouvrir la fenêtre « nouveau mot de passe »
          setTimeout(() => { document.getElementById('pwd-modal')?.classList.remove('hidden'); }, 1800);
        } else if (event === 'SIGNED_OUT') {
          stopAdminBackgroundSync();
          currentUser = null;
          currentProfile = null;
          showScreen('screen-login');
        }
      });
    }

    let authLoadingUserId = null;
    async function handleAuthenticatedUser(user) {
      if (authLoadingUserId === user.id) return; // évite le double chargement (getSession + SIGNED_IN)
      authLoadingUserId = user.id;
      try { await handleAuthenticatedUserInner(user); }
      finally { authLoadingUserId = null; }
    }

    async function handleAuthenticatedUserInner(user) {
      currentUser = user;
      const cachedProfilesMap = lsGetJson('korpo_cached_profiles_v1', '{}');
      const emailKey = String(user.email || '').trim().toLowerCase();
      let profile = cachedProfilesMap[emailKey] || null;
      let error = null;

      if (!profile) {
        const res = await sb
          .from('representants')
          .select('*')
          .ilike('courriel', String(user.email || '').replace(/[\%_]/g, '\\$&'))
          .single();
        profile = res.data;
        error = res.error;
      } else {
        // Rafraîchir le profil en arrière-plan pour détecter tout changement de rôle
        sb.from('representants')
          .select('*')
          .ilike('courriel', String(user.email || '').replace(/[\%_]/g, '\\$&'))
          .single()
          .then(({ data: freshProf }) => {
            if (freshProf) {
              const m = lsGetJson('korpo_cached_profiles_v1', '{}');
              m[emailKey] = freshProf;
              try { localStorage.setItem('korpo_cached_profiles_v1', JSON.stringify(m)); } catch (e) {}
            }
          })
          .catch(() => {});
      }

      if (profile) {
        cachedProfilesMap[emailKey] = profile;
        try { localStorage.setItem('korpo_cached_profiles_v1', JSON.stringify(cachedProfilesMap)); } catch (e) {}
      }

      if (error || !profile) {
        const errBox = document.getElementById('login-error');
        errBox.textContent = `Profil introuvable dans la table representants pour ${user.email}.`;
        errBox.classList.remove('hidden');
        await sb.auth.signOut();
        hideStartupSplash(0);
        showScreen('screen-login');
        return;
      }

      currentProfile = profile;
      document.getElementById('header-user-name').textContent = profile.nom;
      document.getElementById('header-user-email').textContent = profile.courriel;
      const menuNameEl = document.getElementById('menu-user-name');
      const menuEmailEl = document.getElementById('menu-user-email');
      if (menuNameEl) menuNameEl.textContent = profile.nom;
      if (menuEmailEl) menuEmailEl.textContent = profile.courriel;
      document.getElementById('header-avatar').textContent = profile.nom.slice(0, 2).toUpperCase();
      document.getElementById('header-role-badge').textContent = profile.role === 'admin' ? 'Administrateur' : 'Représentant';

      const isAdminRole = profile.role === 'admin';
      ['btn-kpi-targets-toggle', 'btn-kpi-edit-toggle', 'btn-kpi-import-open', 'btn-kpi-export-csv', 'btn-kpi-rep-targets-inline', 'btn-kpi-client-targets-inline'].forEach(btnId => {
        const el = document.getElementById(btnId);
        if (el) el.classList.toggle('hidden', !isAdminRole);
      });

      if (!isAdminRole) {
        document.getElementById('admin-pending-banner')?.classList.add('hidden');
        document.getElementById('kpi-editor-panel')?.classList.add('hidden');
        document.getElementById('kpi-targets-panel')?.classList.add('hidden');
        document.getElementById('kpi-import-modal')?.classList.add('hidden');
        const repSel = document.getElementById('kpi-select-rep');
        if (repSel) {
          repSel.value = 'ALL';
        }
      }

      buildNavigation();
      kvApplyToVariables();
      syncAllKpiRepsFromState();
      renderMesVentes();
      renderClassement();
      renderKpiDashboard();
      showScreen('screen-app');

      const firstName = String(profile.nom || '').trim().split(/\s+/)[0] || profile.nom;
      updateStartupSplash(
        60,
        `Bonjour ${firstName} 👋`,
        'Synchronisation de vos ventes, commissions et objectifs...',
        'Connexion à la base de données...'
      );

      // Synchroniser le cloud et charger les dernières données avant de lever le rideau
      const refreshBtn = document.getElementById('btn-refresh');
      setRefreshButtonSpinning(refreshBtn, true);
      (async () => {
        try {
          await hydrateKv();
          updateStartupSplash(82, null, null, 'Mise à jour des ventes et du tableau KPI...');
          await loadAllData();
          window._initialDataSynced = true;
          renderMesVentes();
          renderKpiDashboard();
          hideStartupSplash(220);
          maybeTriggerFridayWeeklyDigest();
        } catch (bgErr) {
          console.warn('Synchronisation initiale en arrière-plan :', bgErr?.message || bgErr);
          hideStartupSplash(0);
        } finally {
          setRefreshButtonSpinning(refreshBtn, false);
        }
      })();
    }

    let activeMobilePersonalView = 'ventes'; // 'ventes' | 'commissions'

    window.openMobileAdminDrawer = () => {
      document.getElementById('mobile-admin-drawer-backdrop')?.classList.remove('hidden');
    };
    window.closeMobileAdminDrawer = () => {
      document.getElementById('mobile-admin-drawer-backdrop')?.classList.add('hidden');
    };

    function buildNavigation() {
      const isAdmin = currentProfile.role === 'admin';

      const primaryItems = [
        { id: 'tab-mes-ventes', icon: '💵', shortLabel: 'Ventes', label: 'Mes ventes' },
        { id: 'tab-mes-commissions', icon: '💰', shortLabel: 'Comm.', label: 'Mes commissions' },
        { id: 'tab-kpi', icon: '📊', shortLabel: 'KPI', label: 'Tableau KPI' },
        { id: 'tab-classement', icon: '🏆', shortLabel: 'Équipe', label: 'Classement équipe' },
        { id: 'tab-route', icon: '🧭', shortLabel: 'Trajet', label: 'Optimiseur de trajet' }
      ];

      const adminItems = isAdmin ? [
        { id: 'tab-admin-ventes', icon: '📋', shortLabel: 'Registre', label: 'Toutes les ventes', desc: 'Valider & gérer les dossiers' },
        { id: 'tab-admin-dashboard', icon: '📈', shortLabel: 'Direction', label: 'Tableau de bord', desc: 'Analytique, paie & courriels' },
        { id: 'tab-admin-finances', icon: '🏦', shortLabel: 'Finances', label: 'Finances Korpo', desc: 'Encaissements, profit réel & Wave' },
        { id: 'tab-admin-taux', icon: '⚙️', shortLabel: 'Taux', label: 'Grille des taux', desc: 'Commissions & règles clients' },
        { id: 'tab-admin-equipe', icon: '🏢', shortLabel: 'Équipe', label: 'Entreprises & Équipe', desc: 'Objectifs, reps & partenaires' }
      ] : [];

      // 1. Barre basse mobile ergonomique (4 boutons pour rep, 5 boutons avec tiroir Admin pour admin)
      const mobileGrid = document.getElementById('mobile-bottom-nav-grid');
      if (mobileGrid) {
        mobileGrid.className = isAdmin
          ? 'grid grid-cols-6 gap-1 max-w-lg mx-auto'
          : 'grid grid-cols-5 gap-1 max-w-lg mx-auto';

        const primaryHtml = primaryItems.map(t => `
          <button type="button" data-tab="${t.id}" style="touch-action: manipulation; -webkit-tap-highlight-color: transparent;" class="nav-btn relative flex flex-col items-center justify-center py-1.5 px-1 rounded-2xl text-[11px] font-bold transition text-slate-300 hover:text-white active:scale-95">
            <span class="text-lg leading-none mb-0.5">${t.icon}</span>
            <span class="truncate max-w-full">${t.shortLabel}</span>
          </button>
        `).join('');

        const adminTriggerHtml = isAdmin ? `
          <button type="button" id="btn-mobile-admin-trigger" onclick="openMobileAdminDrawer()" style="touch-action: manipulation; -webkit-tap-highlight-color: transparent;" class="relative flex flex-col items-center justify-center py-1.5 px-1 rounded-2xl text-[11px] font-bold transition text-amber-300 hover:text-white active:scale-95">
            <span class="text-lg leading-none mb-0.5">🛡️</span>
            <span class="truncate max-w-full">Admin</span>
          </button>
        ` : '';

        mobileGrid.innerHTML = primaryHtml + adminTriggerHtml;
      }

      // 2. Tiroir mobile Admin (grosses tuiles tactiles)
      const mobileAdminBox = document.getElementById('mobile-admin-drawer-items');
      if (mobileAdminBox) {
        mobileAdminBox.innerHTML = adminItems.map(t => `
          <button type="button" data-tab="${t.id}" onclick="closeMobileAdminDrawer()" class="nav-btn text-left p-3.5 rounded-2xl border border-slate-200 bg-slate-50 hover:bg-indigo-50 transition flex flex-col gap-1">
            <div class="flex items-center gap-2">
              <span class="text-lg">${t.icon}</span>
              <span class="text-xs font-extrabold text-slate-900">${t.label}</span>
            </div>
            <span class="text-[11px] text-slate-500 leading-tight">${t.desc}</span>
          </button>
        `).join('');
      }

      // 3. Menu latéral gauche sur ordinateur (>= 1024px)
      const desktopNav = document.getElementById('desktop-sidebar-nav');
      if (desktopNav) {
        const renderGroup = (title, items) => `
          <div class="space-y-1">
            <div class="px-3 text-[10px] font-extrabold uppercase tracking-wider text-slate-400">${title}</div>
            ${items.map(t => `
              <button type="button" data-tab="${t.id}" class="nav-btn w-full px-3 py-2.5 text-xs font-bold rounded-xl transition flex items-center gap-2.5 text-slate-300 hover:text-white hover:bg-slate-800">
                <span class="text-base leading-none">${t.icon}</span>
                <span class="truncate">${t.label}</span>
              </button>
            `).join('')}
          </div>
        `;
        desktopNav.innerHTML =
          renderGroup('Mon espace', primaryItems.slice(0, 2)) +
          renderGroup('Performance', primaryItems.slice(2, 4)) +
          renderGroup('Terrain', primaryItems.slice(4, 5)) +
          (adminItems.length ? renderGroup('Administration', adminItems) : '');
      }

      document.querySelectorAll('.nav-btn').forEach(btn => {
        btn.addEventListener('click', () => {
          if (window._globalClientFilterQuery) {
            window._globalClientFilterQuery = '';
            showActiveClientSearchBanner('');
          }
          switchTab(btn.dataset.tab);
        });
      });
      switchTab('tab-mes-ventes');
      setTimeout(() => window.KorpoGuide?.maybeStart(), 900);
    }

    function applyMobilePersonalSplitLayout() {
      const sec = document.getElementById('tab-mes-ventes');
      if (!sec) return;
      const isCommView = activeMobilePersonalView === 'commissions';
      const isAdmin = currentProfile?.role === 'admin';

      const titleTextEl = document.getElementById('mes-ventes-main-title');
      const titleIconEl = document.getElementById('mes-ventes-main-icon');
      if (titleTextEl && titleIconEl) {
        if (isCommView) {
          titleIconEl.textContent = '💰';
          titleTextEl.textContent = 'Mes commissions';
        } else {
          titleIconEl.textContent = '💵';
          titleTextEl.textContent = 'Mes ventes';
        }
      }

      // Séparer proprement les sections Ventes vs Commissions sur mobile ET ordinateur
      const isDesktop = window.innerWidth >= 768;
      sec.querySelectorAll('[data-mobile-group="ventes"]').forEach(el => {
        if (el.id === 'btn-rep-open-drawer') {
          el.classList.add('hidden');
        } else if (el.classList.contains('md:hidden')) {
          el.classList.toggle('hidden', isDesktop || isCommView);
        } else if (el.classList.contains('md:block')) {
          el.classList.toggle('hidden', !isDesktop || isCommView);
        } else {
          el.classList.toggle('hidden', isCommView);
        }
      });
      sec.querySelectorAll('[data-mobile-group="commissions"]').forEach(el => {
        if (el.id === 'filter-admin-scope') {
          el.classList.toggle('hidden', !isAdmin || !isCommView);
        } else if (el.id === 'btn-rep-future-comm') {
          el.classList.toggle('hidden', !isCommView);
        } else if (el.classList.contains('md:hidden')) {
          el.classList.toggle('hidden', isDesktop || !isCommView);
        } else if (el.classList.contains('md:block')) {
          el.classList.toggle('hidden', !isDesktop || !isCommView);
        } else {
          el.classList.toggle('hidden', !isCommView);
        }
      });

      const futureOpt = document.querySelector('#filter-rep-period option[value="future"]');
      if (futureOpt) {
        futureOpt.hidden = !isCommView;
        futureOpt.disabled = !isCommView;
      }
      if (!isCommView && document.getElementById('filter-rep-period')?.value === 'future') {
        document.getElementById('filter-rep-period').value = 'this_month';
      }

      const subEl = document.getElementById('mes-ventes-subtitle');
      if (subEl) {
        subEl.textContent = isCommView
          ? 'Vos gains confirmés, paiements reçus et soldes à venir'
          : 'Suivi confidentiel de vos dossiers fermés et de vos dépôts';
      }
    }

    window.addEventListener('resize', applyMobilePersonalSplitLayout);

    function switchTab(tabId) {
      const isMobileCommTab = tabId === 'tab-mes-commissions';
      const targetSectionId = isMobileCommTab ? 'tab-mes-ventes' : tabId;
      if (tabId === 'tab-mes-ventes') activeMobilePersonalView = 'ventes';
      if (isMobileCommTab) activeMobilePersonalView = 'commissions';

      document.querySelectorAll('main > section').forEach(sec => {
        sec.classList.toggle('hidden', sec.id !== targetSectionId);
      });
      const isAdminTabActive = tabId.startsWith('tab-admin-');
      const mobileAdminTrigger = document.getElementById('btn-mobile-admin-trigger');
      if (mobileAdminTrigger) {
        mobileAdminTrigger.classList.toggle('bg-amber-500', isAdminTabActive);
        mobileAdminTrigger.classList.toggle('text-slate-950', isAdminTabActive);
        mobileAdminTrigger.classList.toggle('text-amber-300', !isAdminTabActive);
      }
      document.querySelectorAll('.nav-btn').forEach(btn => {
        const active = btn.dataset.tab === tabId;
        const isMobileSheetCard = Boolean(btn.closest('#mobile-admin-drawer-items'));
        if (isMobileSheetCard) {
          btn.classList.toggle('border-indigo-500', active);
          btn.classList.toggle('bg-indigo-50', active);
          btn.classList.toggle('ring-1', active);
          btn.classList.toggle('ring-indigo-400', active);
        } else {
          btn.classList.toggle('bg-indigo-600', active);
          btn.classList.toggle('text-white', active);
          btn.classList.toggle('shadow-xs', active);
          btn.classList.toggle('text-slate-300', !active);
          btn.classList.toggle('hover:text-white', !active);
          btn.classList.toggle('hover:bg-slate-800', !active);
        }
      });
      if (targetSectionId === 'tab-mes-ventes') applyMobilePersonalSplitLayout();
      if (tabId === 'tab-kpi') renderKpiDashboard();
      if (tabId === 'tab-route') window.KorpoRoute?.render();
      if (tabId === 'tab-admin-finances') renderAdminFinances();
    }

    function setChatBackgroundScrollLocked(locked) {
      document.documentElement.style.overflow = locked ? 'hidden' : '';
      document.body.style.overflow = locked ? 'hidden' : '';
      document.body.style.overscrollBehavior = locked ? 'none' : '';
    }

    window.openTeamChatModal = () => {
      const backdrop = document.getElementById('chat-modal-backdrop');
      if (!backdrop) return;
      window._chatOpenedPrevSeenTs = lastSeenChatTimestamp || '';
      backdrop.classList.remove('hidden');
      setChatBackgroundScrollLocked(true);
      activeChatTabOpen = true;
      markTeamChatAsRead();
      renderTeamChatFeed(true);
      refreshTeamChatMessages(true);
      setTimeout(() => {
        document.getElementById('chat-message-input')?.focus({ preventScroll: true });
      }, 120);
    };

    window.closeTeamChatModal = () => {
      if (activeChatTabOpen) {
        markTeamChatAsRead();
      }
      const backdrop = document.getElementById('chat-modal-backdrop');
      if (backdrop) backdrop.classList.add('hidden');
      setChatBackgroundScrollLocked(false);
      activeChatTabOpen = false;
    };
    window.addEventListener('pagehide', () => {
      if (activeChatTabOpen) markTeamChatAsRead();
    });

    document.getElementById('chat-modal-backdrop')?.addEventListener('click', (e) => {
      if (e.target.id === 'chat-modal-backdrop') closeTeamChatModal();
    });
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && activeChatTabOpen) closeTeamChatModal();
    });

    // =========================================================================
    // LOGIQUE DU TABLEAU KPI (RDV, OTS, SUIVI, % CLOSE & ENTREPRISES)
    // =========================================================================
    function populateKpiMonthSelect(keepSelection = true) {
      const selMonth = document.getElementById('kpi-select-month');
      if (!selMonth) return;
      const curSel = selMonth.value;
      normalizeKpiStoreReps();
      const months = Object.keys(kpiStore).filter(isValidYearMonth).sort().reverse();
      const years = Array.from(new Set(months.map(m => m.slice(0, 4)))).sort().reverse();
      const quarters = Array.from(new Set(months.map(m => getQuarterForMonth(m)).filter(Boolean))).sort().reverse();
      const curQ = getQuarterForMonth(currentMonthStr);
      const yearOptions = years.map(y => `<option value="YEAR_${y}">🗓️ ${y === String(today.getFullYear()) ? `Cette année (${y})` : `Toute l'année ${y}`}</option>`).join('');
      const quarterOptions = quarters.map(q => {
        const [qy, qn] = q.split('-Q');
        const isCurQ = q === curQ;
        return `<option value="QUARTER_${q}">📊 ${isCurQ ? `Ce trimestre (T${qn} ${qy})` : `Trimestre T${qn} ${qy}`}</option>`;
      }).join('');
      const monthOptions = months.map(ym => `<option value="${ym}">${ym === currentMonthStr ? `📅 Ce mois-ci (${esc(kpiStore[ym].label)})` : esc(kpiStore[ym].label)}</option>`).join('');
      selMonth.innerHTML = yearOptions + quarterOptions + monthOptions;
      if (keepSelection && curSel && (curSel.startsWith('YEAR_') || curSel.startsWith('QUARTER_') || months.includes(curSel))) {
        selMonth.value = curSel;
      } else if (typeof currentMonthStr !== 'undefined' && months.includes(currentMonthStr)) {
        selMonth.value = currentMonthStr;
      } else if (months.length) {
        selMonth.value = months[0];
      }
    }

    // =========================================================================
    // GESTION DES OBJECTIFS PAR CLIENT & COMPAGNIE (ACTIVABLES)
    // =========================================================================
    let clientTargetsStore = lsGetJson('korpo_client_targets_v1', '[]');
    let companyTargetsMetaStore = lsGetJson('korpo_company_targets_meta_v1', '{}');

    function saveClientTargetsStore() {
      lsSet('korpo_client_targets_v1', JSON.stringify(clientTargetsStore));
    }

    async function saveCompanyTargetMeta(entNom, repNom, metaObj, scopeYm = 'ALL') {
      const key = scopeYm && scopeYm !== 'ALL'
        ? `${entNom}__${repNom || 'ALL'}__${scopeYm}`
        : `${entNom}__${repNom || 'ALL'}`;
      companyTargetsMetaStore[key] = { ...metaObj, entNom, repNom: repNom || 'ALL', monthScope: scopeYm || 'ALL' };
      lsSet('korpo_company_targets_meta_v1', JSON.stringify(companyTargetsMetaStore));
      applyCompanyTargetsMetaToStore();
      if (sb && currentProfile) {
        await safeUpsertKpiHebdoRow({
          mois_commission: key,
          semaine_index: -8,
          representant_nom: JSON.stringify(companyTargetsMetaStore[key]),
          booke: 0,
          rdv_realise: 0,
          updated_at: new Date().toISOString()
        });
      }
    }

    async function removeCompanyTargetMeta(key) {
      delete companyTargetsMetaStore[key];
      lsSet('korpo_company_targets_meta_v1', JSON.stringify(companyTargetsMetaStore));
      if (sb && currentProfile) {
        try {
          await sb.from('kpi_hebdo').delete().eq('mois_commission', key).eq('semaine_index', -8);
        } catch (e) {}
      }
    }

    function applyCompanyTargetsMetaToStore() {
      if (!companyTargetsMetaStore || typeof companyTargetsMetaStore !== 'object') return;
      const entries = Object.entries(companyTargetsMetaStore);
      if (!entries.length) return;

      // D'abord appliquer les règles globales ('ALL'), puis les règles spécifiques à un mois ('YYYY-MM')
      const sortedEntries = entries.sort((a, b) => {
        const scA = a[1]?.monthScope || (a[0].split('__')[2] || 'ALL');
        const scB = b[1]?.monthScope || (b[0].split('__')[2] || 'ALL');
        if (scA === scB) return 0;
        return scA === 'ALL' ? -1 : 1;
      });

      sortedEntries.forEach(([key, meta]) => {
        if (!meta || typeof meta !== 'object') return;
        const parts = key.split('__');
        const entNom = meta.entNom || parts[0];
        const repNom = meta.repNom || parts[1] || 'ALL';
        const monthScope = meta.monthScope || parts[2] || 'ALL';
        const monthlyAmt = Number(meta.monthlyAmount || (meta.annualAmount ? Math.round((Number(meta.annualAmount) / 12) * 100) / 100 : 0));
        if (!entNom || monthlyAmt <= 0) return;

        const targetMonths = monthScope === 'ALL' ? Object.keys(kpiStore) : (kpiStore[monthScope] ? [monthScope] : []);
        targetMonths.forEach(ym => {
          const mObj = kpiStore[ym];
          if (!mObj || !mObj.entreprises) return;
          if (!mObj.entreprises[entNom]) {
            const nw = mObj.num_weeks || 4;
            mObj.entreprises[entNom] = {
              reps: {
                'Olivier': { weeks: Array(nw).fill(0), cumul: 0, objectif: 0 },
                'Jean-Sébastien': { weeks: Array(nw).fill(0), cumul: 0, objectif: 0 },
                'Gabryel': { weeks: Array(nw).fill(0), cumul: 0, objectif: 0 },
                'Samuel': { weeks: Array(nw).fill(0), cumul: 0, objectif: 0 },
                'Alexandre': { weeks: Array(nw).fill(0), cumul: 0, objectif: 0 }
              },
              total: { weeks: Array(nw).fill(0), cumul: 0, objectif: 0 }
            };
          }
          if (repNom === 'ALL') {
            mObj.entreprises[entNom].total.objectif = monthlyAmt;
          } else if (mObj.entreprises[entNom].reps[repNom]) {
            mObj.entreprises[entNom].reps[repNom].objectif = monthlyAmt;
            // Recalculer le total compagnie uniquement si aucune règle globale 'ALL' explicite n'a été posée sur cette compagnie
            const hasExplicitAll = Boolean(companyTargetsMetaStore[`${entNom}__ALL`] || companyTargetsMetaStore[`${entNom}__ALL__${ym}`]);
            if (!hasExplicitAll) {
              mObj.entreprises[entNom].total.objectif = Object.values(mObj.entreprises[entNom].reps)
                .reduce((s, r) => s + Number(r.objectif || 0), 0);
            }
          }
        });
      });
    }

    window.openClientTargetsPanel = (presetEnt = '', presetClient = '') => {
      if (currentProfile?.role !== 'admin') return;
      const panel = document.getElementById('kpi-targets-panel');
      panel.classList.remove('hidden');
      populateTargetFormOptions();
      if (presetEnt) document.getElementById('target-ent-select').value = presetEnt;
      updateTargetClientDatalist();
      if (presetClient) document.getElementById('target-client-input').value = presetClient;
      renderActiveClientTargetsList();
      panel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    };

    function populateTargetFormOptions() {
      const entSel = document.getElementById('target-ent-select');
      if (!entSel) return;
      const entNames = new Set(['Toitex', 'MonSpray', 'VerreVDM', 'Groupe Diamant Pro', 'MaClim', ...state.entreprises.map(e => e.nom)].filter(en => !isEntrepriseArchived(en)));
      const cur = entSel.value;
      entSel.innerHTML = Array.from(entNames).map(n => `<option value="${esc(n)}">${esc(n)}</option>`).join('');
      if (cur && entNames.has(cur)) entSel.value = cur;
      updateTargetClientDatalist();
    }

    function uniqueClientsByName(list) {
      const seen = new Set();
      return (list || []).filter(c => {
        const k = normClientKey(c.nom);
        if (!k || seen.has(k)) return false;
        seen.add(k);
        return true;
      });
    }

    function updateTargetClientDatalist() {
      const entName = document.getElementById('target-ent-select')?.value;
      const dl = document.getElementById('target-clients-datalist');
      if (!dl) return;
      const entObj = state.entreprises.find(e => e.nom === entName);
      const matchingClients = uniqueClientsByName(
        entObj ? state.clients.filter(c => String(c.entreprise_id) === String(entObj.id)) : state.clients
      );
      dl.innerHTML = matchingClients.map(c => `<option value="${esc(c.nom)}">`).join('');
    }

    function renderActiveClientTargetsList() {
      const box = document.getElementById('active-client-targets-list');
      if (!box) return;
      const companyEntries = Object.entries(companyTargetsMetaStore || {}).filter(([, m]) => m && Number(m.monthlyAmount || m.annualAmount) > 0);
      if (!clientTargetsStore.length && !companyEntries.length) {
        box.innerHTML = `<div class="col-span-full text-xs text-slate-400 italic py-2">Aucun objectif personnalisé par compagnie ou client pour le moment. Utilisez le formulaire ci-dessus pour en activer un.</div>`;
        return;
      }

      const companyCardsHtml = companyEntries.map(([key, m]) => {
        const entN = m.entNom || key.split('__')[0];
        const repN = m.repNom || key.split('__')[1] || 'ALL';
        const isAnn = m.periodicity === 'annual';
        const labelVal = isAnn
          ? `${fmtMoney0(m.annualAmount)}/an (${fmtMoney0(m.monthlyAmount)}/mois)`
          : `${fmtMoney0(m.monthlyAmount)}/mois (${fmtMoney0(m.annualAmount || (m.monthlyAmount * 12))}/an)`;
        const safeKey = key.replace(/'/g, "\'");
        return `
          <div class="flex items-center justify-between p-3 rounded-xl border bg-indigo-50/50 border-indigo-200">
            <div class="min-w-0 pr-2">
              <div class="text-xs font-extrabold text-indigo-950 truncate">🏢 ${esc(entN)} <span class="text-[10px] font-bold px-1.5 py-0.2 rounded bg-indigo-100 text-indigo-700">${isAnn ? 'Annuel' : 'Mensuel'}</span></div>
              <div class="text-[11px] text-slate-700 tabular-nums mt-0.5">Cible : <strong>${labelVal}</strong> • ${esc(repN === 'ALL' ? 'Équipe' : repN)}</div>
            </div>
            <button type="button" onclick="deleteCompanyTargetEntry('${safeKey}')" title="Supprimer cet objectif compagnie" class="px-2 py-1 text-[10px] font-bold text-rose-600 hover:bg-rose-50 rounded-lg shrink-0">✕</button>
          </div>
        `;
      }).join('');

      const clientCardsHtml = clientTargetsStore.map((t, idx) => `
        <div class="flex items-center justify-between p-3 rounded-xl border ${t.active !== false ? 'bg-emerald-50/50 border-emerald-200' : 'bg-slate-50 border-slate-200 opacity-60'}">
          <div class="min-w-0 pr-2">
            <div class="text-xs font-extrabold text-slate-900 truncate">👤 ${esc(t.clientNom || 'Toute la compagnie')} <span class="text-[10px] font-semibold text-slate-500">(${esc(t.entNom)})</span></div>
            <div class="text-[11px] text-slate-600 tabular-nums">Cible : <strong>${t.periodicity === 'annual' ? `${fmtMoney0(t.annualObjectif || (t.objectif * 12))}/an` : `${fmtMoney0(t.objectif)}/mois`}</strong> • ${esc(t.repNom === 'ALL' ? 'Équipe' : t.repNom)}</div>
          </div>
          <div class="flex items-center gap-1.5 shrink-0">
            <button type="button" onclick="toggleClientTargetActive(${idx})" title="Activer / Désactiver" class="px-2 py-1 text-[10px] font-bold rounded-lg ${t.active !== false ? 'bg-emerald-600 text-white' : 'bg-slate-200 text-slate-700'}">
              ${t.active !== false ? 'Actif' : 'Inactif'}
            </button>
            <button type="button" onclick="removeClientTarget(${idx})" title="Supprimer" class="px-2 py-1 text-[10px] font-bold text-rose-600 hover:bg-rose-50 rounded-lg">✕</button>
          </div>
        </div>
      `).join('');

      box.innerHTML = companyCardsHtml + clientCardsHtml;
    }

    window.deleteCompanyTargetEntry = async (key) => {
      await removeCompanyTargetMeta(key);
      renderActiveClientTargetsList();
      renderKpiDashboard();
      toast('Objectif compagnie supprimé');
    };

    window.toggleClientTargetActive = (idx) => {
      if (!clientTargetsStore[idx]) return;
      clientTargetsStore[idx].active = clientTargetsStore[idx].active === false ? true : false;
      saveClientTargetsStore();
      renderActiveClientTargetsList();
      renderKpiDashboard();
      toast(clientTargetsStore[idx].active ? 'Objectif client activé' : 'Objectif client mis en pause');
    };

    window.removeClientTarget = (idx) => {
      clientTargetsStore.splice(idx, 1);
      saveClientTargetsStore();
      renderActiveClientTargetsList();
      renderKpiDashboard();
      toast('Objectif client supprimé');
    };

    let repTargetsOverrides = lsGetJson('korpo_rep_targets_v1', '{}');

    function applyRepTargetsOverridesToStore() {
      Object.entries(repTargetsOverrides).forEach(([key, vals]) => {
        const [scopeYm, rn] = key.split('_');
        const targetMonths = scopeYm === 'ALL' ? Object.keys(kpiStore) : (kpiStore[scopeYm] ? [scopeYm] : []);
        targetMonths.forEach(ym => {
          const mObj = kpiStore[ym];
          if (!mObj || !mObj.reps || !mObj.reps[rn]) return;
          mObj.reps[rn]['Vente $'].objectif = Number(vals.dollars || 0);
          mObj.reps[rn]['Booké'].objectif = Number(vals.booke || 0);
          mObj.reps[rn]['RDV Réalisé'].objectif = Number(vals.rdv || 0);
          mObj.reps[rn]['Ventes OTS'].objectif = Number(vals.ots || 0);
          mObj.reps[rn]['Ventes SUIVI'].objectif = Number(vals.suivi || 0);
          mObj.reps[rn]['% Close'].objectif = Number(vals.close || 0);
          recalcTeamObjectivesForMonth(mObj);
        });
      });
    }

    function recalcTeamObjectivesForMonth(mObj) {
      if (!mObj || !mObj.team || !mObj.reps) return;
      const ym = mObj.ym || '';
      const reps = Object.keys(mObj.reps);
      reps.forEach(rn => {
        const factor = isRepArchived(rn) ? (isValidYearMonth(ym) ? getRepArchiveMonthFactor(rn, ym) : 0) : 1;
        if (factor < 1) {
          const entry = findArchivedRepEntry(rn);
          const snapObj = entry?.targets || {};
          const applyFactor = (mk, baseVal) => {
            if (!mObj.reps[rn][mk]) return;
            const rawBase = Number(baseVal || mObj.reps[rn][mk].objectif || 0);
            mObj.reps[rn][mk].objectif = Math.round(rawBase * factor * 100) / 100;
          };
          applyFactor('Vente $', snapObj.dollars);
          applyFactor('Booké', snapObj.booke);
          applyFactor('RDV Réalisé', snapObj.rdv);
          applyFactor('Ventes OTS', snapObj.ots);
          applyFactor('Ventes SUIVI', snapObj.suivi);
          if (factor === 0 && mObj.reps[rn]['% Close']) mObj.reps[rn]['% Close'].objectif = 0;
        }
      });
      mObj.team['Booké'].objectif = Math.round(reps.reduce((s, rn) => s + Number(mObj.reps[rn]['Booké']?.objectif || 0), 0));
      mObj.team['RDV RÉA'].objectif = Math.round(reps.reduce((s, rn) => s + Number(mObj.reps[rn]['RDV Réalisé']?.objectif || 0), 0));
      mObj.team['VENTE OTS'].objectif = Math.round(reps.reduce((s, rn) => s + Number(mObj.reps[rn]['Ventes OTS']?.objectif || 0), 0));
      mObj.team['VENTE SUIVI'].objectif = Math.round(reps.reduce((s, rn) => s + Number(mObj.reps[rn]['Ventes SUIVI']?.objectif || 0), 0));
      mObj.team['VENTE TOTAL'].objectif = mObj.team['VENTE OTS'].objectif + mObj.team['VENTE SUIVI'].objectif;
      mObj.team['VENTE $'].objectif = Math.round(reps.reduce((s, rn) => s + Number(mObj.reps[rn]['Vente $']?.objectif || 0), 0) * 100) / 100;
    }

    function syncRepTargetInputs() {
      const repSel = document.getElementById('rep-target-select');
      if (!repSel) return;
      const rn = repSel.value;
      const curSelMonth = document.getElementById('kpi-select-month')?.value || currentMonthStr;
      const ym = curSelMonth.startsWith('YEAR_') ? currentMonthStr : curSelMonth;
      const mObj = ensureKpiMonthStructure(ym);
      const rData = mObj.reps[rn] || {};
      const commEl = document.getElementById('rep-obj-comm');
      if (commEl) commEl.value = getCommissionTargetForRep(rn, ym);
      const bonusEl = document.getElementById('rep-obj-bonus');
      if (bonusEl) bonusEl.value = Math.round(getBonusPctForRep(rn, ym) * 10000) / 100;
      document.getElementById('rep-obj-dollars').value = rData['Vente $']?.objectif ?? 150000;
      document.getElementById('rep-obj-booke').value = rData['Booké']?.objectif ?? 80;
      document.getElementById('rep-obj-rdv').value = rData['RDV Réalisé']?.objectif ?? 60;
      document.getElementById('rep-obj-ots').value = rData['Ventes OTS']?.objectif ?? 10;
      document.getElementById('rep-obj-suivi').value = rData['Ventes SUIVI']?.objectif ?? 20;
      document.getElementById('rep-obj-close').value = rData['% Close']?.objectif ?? 35;
    }

    window.openRepTargetsPanel = (presetRep = '') => {
      if (currentProfile?.role !== 'admin') return;
      const panel = document.getElementById('kpi-targets-panel');
      panel.classList.remove('hidden');
      const activeFilterRep = presetRep || document.getElementById('kpi-select-rep')?.value;
      const repSel = document.getElementById('rep-target-select');
      if (repSel && activeFilterRep && activeFilterRep !== 'ALL' && ALL_KPI_REPS.includes(activeFilterRep)) {
        repSel.value = activeFilterRep;
      }
      syncRepTargetInputs();
      populateTargetFormOptions();
      renderActiveClientTargetsList();
      panel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    };

    function initClientTargetsControls() {
      applyRepTargetsOverridesToStore();
      const btnToggle = document.getElementById('btn-kpi-targets-toggle');
      if (btnToggle) {
        btnToggle.addEventListener('click', () => {
          const panel = document.getElementById('kpi-targets-panel');
          panel.classList.toggle('hidden');
          if (!panel.classList.contains('hidden')) {
            const activeFilterRep = document.getElementById('kpi-select-rep')?.value;
            const repSel = document.getElementById('rep-target-select');
            if (repSel && activeFilterRep && activeFilterRep !== 'ALL' && ALL_KPI_REPS.includes(activeFilterRep)) {
              repSel.value = activeFilterRep;
            }
            syncRepTargetInputs();
            populateTargetFormOptions();
            renderActiveClientTargetsList();
          }
        });
      }

      const repTargetSel = document.getElementById('rep-target-select');
      if (repTargetSel) repTargetSel.addEventListener('change', syncRepTargetInputs);

      const formRepTarget = document.getElementById('form-rep-target');
      if (formRepTarget) {
        formRepTarget.addEventListener('submit', async (e) => {
          e.preventDefault();
          const rn = document.getElementById('rep-target-select').value;
          const commGoal = parseFloat(document.getElementById('rep-obj-comm')?.value) || 0;
          const dollars = parseFloat(document.getElementById('rep-obj-dollars').value) || 0;
          const bonusPctInput = parseFloat(document.getElementById('rep-obj-bonus')?.value);
          const bonusPctDec = isNaN(bonusPctInput) ? getBonusPctForRep(rn, currentMonthStr) : Math.max(0, bonusPctInput / 100);
          const booke = parseFloat(document.getElementById('rep-obj-booke').value) || 0;
          const rdv = parseFloat(document.getElementById('rep-obj-rdv').value) || 0;
          const ots = parseFloat(document.getElementById('rep-obj-ots').value) || 0;
          const suivi = parseFloat(document.getElementById('rep-obj-suivi').value) || 0;
          const close = parseFloat(document.getElementById('rep-obj-close').value) || 0;
          const applyAll = document.getElementById('rep-target-apply-all').checked;

          const curSelMonth = document.getElementById('kpi-select-month')?.value || currentMonthStr;
          const activeYm = curSelMonth.startsWith('YEAR_') ? currentMonthStr : curSelMonth;
          ensureKpiMonthStructure(activeYm);

          const scopeKey = applyAll ? `ALL_${rn}` : `${activeYm}_${rn}`;
          repTargetsOverrides[scopeKey] = { dollars, booke, rdv, ots, suivi, close, comm: commGoal, bonusPct: bonusPctDec };
          lsSet('korpo_rep_targets_v1', JSON.stringify(repTargetsOverrides));

          await saveRepTargetsToStorageAndCloud(rn, activeYm, dollars, commGoal, applyAll, bonusPctDec);

          const monthsToUpdate = applyAll ? Object.keys(kpiStore) : [activeYm];
          monthsToUpdate.forEach(ym => {
            const mObj = ensureKpiMonthStructure(ym);
            if (!mObj.reps[rn]) return;
            mObj.reps[rn]['Vente $'].objectif = dollars;
            mObj.reps[rn]['Booké'].objectif = booke;
            mObj.reps[rn]['RDV Réalisé'].objectif = rdv;
            mObj.reps[rn]['Ventes OTS'].objectif = ots;
            mObj.reps[rn]['Ventes SUIVI'].objectif = suivi;
            mObj.reps[rn]['% Close'].objectif = close;
            recalcTeamObjectivesForMonth(mObj);
          });

          lsSet('korpo_kpi_store_v1', JSON.stringify(kpiStore));
          renderKpiDashboard();
          renderMesVentes();
          toast(`Objectifs mis à jour pour ${rn} : ${fmtMoney0(commGoal)} comm. / ${fmtMoney0(dollars)} ventes ✓`);
        });
      }
      const btnClose = document.getElementById('btn-kpi-close-targets');
      if (btnClose) {
        btnClose.addEventListener('click', () => {
          document.getElementById('kpi-targets-panel').classList.add('hidden');
        });
      }
      const entSel = document.getElementById('target-ent-select');
      if (entSel) entSel.addEventListener('change', updateTargetClientDatalist);

      const perSel = document.getElementById('target-periodicity-select');
      if (perSel) {
        perSel.addEventListener('change', () => {
          const isAnnual = perSel.value === 'annual';
          const wrap = document.getElementById('target-all-months-wrapper');
          if (wrap) wrap.classList.toggle('hidden', isAnnual);
        });
      }

      const form = document.getElementById('form-client-target');
      if (form) {
        form.addEventListener('submit', async (e) => {
          e.preventDefault();
          const entNom = document.getElementById('target-ent-select').value;
          const clientNom = document.getElementById('target-client-input').value.trim();
          const repNom = document.getElementById('target-rep-select').value;
          const periodicity = document.getElementById('target-periodicity-select')?.value || 'monthly';
          const isAnnualInput = periodicity === 'annual';
          const rawAmount = parseAmount(document.getElementById('target-amount-input').value);
          const allMonths = isAnnualInput ? true : document.getElementById('target-apply-all-months').checked;
          const curSelMonth = document.getElementById('kpi-select-month').value;
          const targetMonth = allMonths ? 'ALL' : (curSelMonth.startsWith('YEAR_') ? currentMonthStr : curSelMonth);

          if (!rawAmount || isNaN(rawAmount) || rawAmount <= 0) return toast('Entrez un objectif supérieur à 0 $', 'error');

          const monthlyEquivalent = isAnnualInput ? Math.round((rawAmount / 12) * 100) / 100 : rawAmount;
          const annualEquivalent = isAnnualInput ? rawAmount : (rawAmount * 12);

          // Si aucun client précis n'est saisi, mettre à jour l'objectif de la compagnie dans kpiStore
          if (!clientNom) {
            const monthsToUpdate = allMonths ? Object.keys(kpiStore) : [targetMonth];
            monthsToUpdate.forEach(ym => ensureKpiMonthStructure(ym));
            await saveCompanyTargetMeta(entNom, repNom, {
              periodicity,
              monthlyAmount: monthlyEquivalent,
              annualAmount: annualEquivalent
            }, targetMonth);
            lsSet('korpo_kpi_store_v1', JSON.stringify(kpiStore));
            toast(isAnnualInput
              ? `Objectif annuel mis à jour pour ${entNom} : ${fmtMoney0(annualEquivalent)}/an (${fmtMoney0(monthlyEquivalent)}/mois)`
              : `Objectif mensuel mis à jour pour ${entNom} : ${fmtMoney0(monthlyEquivalent)}/mois`);
          } else {
            // Enregistrer ou remplacer l'objectif spécifique à ce client
            const existingIdx = clientTargetsStore.findIndex(x =>
              x.entNom.toLowerCase() === entNom.toLowerCase() &&
              x.clientNom.toLowerCase() === clientNom.toLowerCase() &&
              x.repNom === repNom
            );
            const entry = {
              entNom,
              clientNom,
              repNom,
              objectif: monthlyEquivalent,
              annualObjectif: isAnnualInput ? annualEquivalent : null,
              periodicity,
              monthScope: targetMonth,
              active: true
            };
            if (existingIdx >= 0) clientTargetsStore[existingIdx] = entry;
            else clientTargetsStore.push(entry);
            saveClientTargetsStore();
            toast(isAnnualInput
              ? `Objectif annuel activé pour ${clientNom} (${fmtMoney0(annualEquivalent)}/an)`
              : `Objectif mensuel activé pour ${clientNom} (${fmtMoney0(monthlyEquivalent)}/mois)`);
          }

          document.getElementById('target-client-input').value = '';
          document.getElementById('target-amount-input').value = '';
          renderActiveClientTargetsList();
          renderKpiDashboard();
        });
      }
    }

    function renderClientObjectivesGrid(selVal, repFilter) {
      const grid = document.getElementById('kpi-client-objectives-grid');
      const badge = document.getElementById('kpi-client-obj-badge');
      if (!grid || !badge) return;

      const kpiMode = document.getElementById('kpi-period-mode')?.value || '';
      const isAll = kpiMode === 'all';
      const isWeek = kpiMode === 'this_week';
      const isCustomRange = kpiMode === 'custom_range';
      const isYear = !isWeek && !isCustomRange && selVal && selVal.startsWith('YEAR_');
      const isQuarter = !isWeek && !isCustomRange && selVal && selVal.startsWith('QUARTER_');
      const yearPrefix = isYear ? selVal.replace('YEAR_', '') : null;
      const quarterTarget = isQuarter ? selVal.replace('QUARTER_', '') : null;
      const activeTargets = clientTargetsStore.filter(t => {
        if (t.active === false) return false;
        if (repFilter !== 'ALL' && t.repNom !== 'ALL' && t.repNom !== repFilter) return false;
        if (!isAll && !isYear && !isQuarter && !isWeek && !isCustomRange && t.monthScope !== 'ALL' && t.monthScope !== selVal) return false;
        return true;
      });

      badge.textContent = `${activeTargets.length} actif(s)`;

      if (!activeTargets.length) {
        grid.innerHTML = `
          <div class="col-span-full rounded-2xl border border-dashed border-slate-300 p-6 text-center space-y-2">
            <div class="text-xs font-bold text-slate-700">Aucun objectif client activé pour cette sélection</div>
            <p class="text-[11px] text-slate-500 max-w-md mx-auto">Cliquez sur « + Ajouter / Modifier un objectif client » pour définir une cible mensuelle en dollars sur un client précis (ex. un compte clé Toitex, MaClim ou MonSpray).</p>
          </div>
        `;
        return;
      }

      const customRangeFactor = isCustomRange
        ? computeCustomRangeMonthFactor(
            document.getElementById('kpi-custom-start')?.value || `${currentMonthStr}-01`,
            document.getElementById('kpi-custom-end')?.value || currentDateStr
          )
        : (isWeek ? (5 / Math.max(1, countBusinessDaysInMonth(currentMonthStr))) : 1);

      const monthsCount = isAll
        ? Object.keys(kpiStore).filter(isValidYearMonth).length || 1
        : (isYear
          ? Object.keys(kpiStore).filter(ym => ym.startsWith(yearPrefix)).length || 1
          : (isQuarter
            ? Object.keys(kpiStore).filter(ym => getQuarterForMonth(ym) === quarterTarget).length || 1
            : ((isWeek || isCustomRange) ? customRangeFactor : 1)));

      grid.innerHTML = activeTargets.map(t => {
        const targetAmount = Math.round((t.monthScope === 'ALL' ? (t.objectif * monthsCount) : (t.objectif * ((isWeek || isCustomRange) ? customRangeFactor : 1))));
        const matchedSales = state.ventes.filter(v => {
          const ent = state.entreprises.find(e => e.id === v.entreprise_id)?.nom || '';
          const cli = state.clients.find(c => c.id === v.client_id)?.nom || '';
          const rep = state.representants.find(r => r.id === v.representant_id)?.nom || '';
          if (ent.toLowerCase() !== t.entNom.toLowerCase()) return false;
          if (cli.toLowerCase() !== t.clientNom.toLowerCase()) return false;
          if (t.repNom !== 'ALL' && rep !== t.repNom) return false;
          if (repFilter !== 'ALL' && rep !== repFilter) return false;
          if (isWeek) return matchesSalePeriod(v, { mode: 'this_week', weekBounds: getCurrentWeekBounds(today) }, true);
          if (isCustomRange) {
            const sVal = document.getElementById('kpi-custom-start')?.value || '';
            const eVal = document.getElementById('kpi-custom-end')?.value || '';
            return matchesSalePeriod(v, { mode: 'custom_range', startVal: sVal, endVal: eVal }, true);
          }
          if (isAll) return true;
          const ym = getSaleStatsMonth(v);
          if (isYear) return ym.startsWith(yearPrefix);
          if (isQuarter) return getQuarterForMonth(ym) === quarterTarget;
          return ym === selVal;
        });

        const realized = matchedSales.reduce((s, v) => s + Number(v.valeur || 0), 0);
        const pct = targetAmount > 0 ? Math.round((realized / targetAmount) * 100) : 0;

        return `
          <div class="rounded-2xl border border-slate-200/90 p-4 bg-slate-50/40 flex flex-col justify-between space-y-3">
            <div class="space-y-2">
              <div class="flex items-center justify-between gap-2">
                <div class="font-extrabold text-sm text-slate-900 truncate">${esc(t.clientNom)}</div>
                ${entBadge(t.entNom)}
              </div>
              <div class="flex items-center justify-between text-[11px] text-slate-500">
                <span>Rep : <strong>${esc(t.repNom === 'ALL' ? 'Tous' : t.repNom)}</strong></span>
                <span class="font-extrabold ${pct >= 100 ? 'text-emerald-600' : 'text-slate-700'} tabular-nums">${pct} %</span>
              </div>
              <div class="flex items-baseline justify-between pt-0.5">
                <span class="text-xl font-extrabold text-slate-900 tabular-nums">${fmtMoney0(realized)}</span>
                <span class="text-[11px] text-slate-400">Obj : ${fmtMoney0(targetAmount)}</span>
              </div>
              <div class="w-full h-2 bg-slate-200/70 rounded-full overflow-hidden">
                <div class="h-full ${progressColor(pct)} rounded-full" style="width: ${Math.min(100, pct)}%"></div>
              </div>
            </div>
            <div class="flex items-center justify-between text-[11px] text-slate-400 pt-1 border-t border-slate-100">
              <span>${matchedSales.length} vente(s) enregistrée(s)</span>
              ${statusBadge(realized, targetAmount)}
            </div>
          </div>
        `;
      }).join('');
    }

    // =========================================================================
    // SYSTÈME DE NOTIFICATIONS EN TEMPS RÉEL + SECOURS AUTOMATIQUE (ADMIN)
    // =========================================================================
    let realtimeVentesChannel = null;
    let adminPollingTimer = null;
    let knownPendingSaleIds = new Set();
    let hasInitializedPendingBaseline = false;
    let sharedAudioCtx = null;

    function unlockAudioContext() {
      try {
        if (!sharedAudioCtx) {
          sharedAudioCtx = new (window.AudioContext || window.webkitAudioContext)();
        }
        if (sharedAudioCtx.state === 'suspended') {
          sharedAudioCtx.resume();
        }
      } catch (e) {}
    }
    document.addEventListener('click', unlockAudioContext, { once: false, passive: true });

    function playNotificationChime() {
      try {
        const ctx = sharedAudioCtx || new (window.AudioContext || window.webkitAudioContext)();
        if (ctx.state === 'suspended') ctx.resume();
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        osc.type = 'sine';
        osc.frequency.setValueAtTime(880, ctx.currentTime);
        osc.frequency.setValueAtTime(1174.66, ctx.currentTime + 0.12);
        gain.gain.setValueAtTime(0.15, ctx.currentTime);
        gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.48);
        osc.connect(gain);
        gain.connect(ctx.destination);
        osc.start();
        osc.stop(ctx.currentTime + 0.48);
      } catch (e) {}
    }

    function sendBrowserPushNotification(title, body) {
      if (!isMentionPushCurrentlyEnabled()) return;
      if (korpoSwRegistration && typeof korpoSwRegistration.showNotification === 'function') {
        korpoSwRegistration.showNotification(title, {
          body,
          tag: 'korpo-new-sale',
          renotify: true,
          data: { url: './' }
        }).catch(() => {});
        return;
      }
      if (Notification.permission === 'granted') {
        const n = new Notification(title, { body, tag: 'korpo-new-sale' });
        n.onclick = () => {
          window.focus();
          goToPendingSales();
          n.close();
        };
      }
    }

    function renderUnifiedBellNotifications() {
      if (!currentProfile) return;
      const isAdmin = currentProfile.role === 'admin';
      const wrapper = document.getElementById('admin-notif-wrapper');
      const badge = document.getElementById('admin-notif-badge');
      const listEl = document.getElementById('admin-notif-list');
      const banner = document.getElementById('admin-pending-banner');
      const markAllBtn = document.getElementById('btn-mark-all-mentions-bell');
      const pendingFooterBtn = document.getElementById('btn-notif-footer-pending');
      if (!wrapper || !badge || !listEl) return;

      wrapper.classList.remove('hidden');

      const myMentions = (typeof teamChatMessages !== 'undefined' ? teamChatMessages : [])
        .filter(m => isCurrentUserMentionedInMsg(m))
        .sort((a, b) => String(b.at || '').localeCompare(String(a.at || '')));
      const unreadMentions = myMentions.filter(m => !readMentionIds.has(String(m.id)));

      const recentRedCalls = isAdmin ? getRecentRedPhoneCallsForAdmin() : [];
      const unreadRedCalls = recentRedCalls.filter(c => !seenRedPhoneCallIds.has(String(c.id)));

      const pendingSales = isAdmin
        ? (state.ventes || [])
            .filter(v => v.statut_depot === 'a_confirmer')
            .sort((a, b) => (b.created_at || b.date_vente || '').localeCompare(a.created_at || a.date_vente || ''))
        : [];

      const totalBadgeCount = unreadMentions.length + pendingSales.length + unreadRedCalls.length;
      if (totalBadgeCount > 0) {
        badge.textContent = totalBadgeCount > 99 ? '99+' : String(totalBadgeCount);
        badge.classList.remove('hidden');
      } else {
        badge.classList.add('hidden');
      }

      if (markAllBtn) markAllBtn.classList.toggle('hidden', unreadMentions.length === 0);
      if (pendingFooterBtn) pendingFooterBtn.classList.toggle('hidden', !(isAdmin && pendingSales.length > 0));

      if (isAdmin && pendingSales.length > 0) {
        if (banner) {
          banner.classList.remove('hidden');
          const latest = pendingSales[0];
          const repN = state.representants.find(r => r.id === latest.representant_id)?.nom || 'Un représentant';
          const entN = state.entreprises.find(e => e.id === latest.entreprise_id)?.nom || '';
          const cliN = state.clients.find(c => c.id === latest.client_id)?.nom || '';
          document.getElementById('admin-pending-banner-title').textContent = `${pendingSales.length} vente(s) en attente de confirmation`;
          document.getElementById('admin-pending-banner-desc').textContent = `Dernière : ${repN} • ${entN} — ${cliN} (${fmtMoney(latest.valeur)})`;
        }
      } else if (banner) {
        banner.classList.add('hidden');
      }

      const subParts = [];
      if (unreadMentions.length > 0) subParts.push(`${unreadMentions.length} mention(s) @`);
      if (isAdmin && pendingSales.length > 0) subParts.push(`${pendingSales.length} vente(s) à confirmer`);
      document.getElementById('admin-notif-subtitle').textContent = subParts.length
        ? subParts.join(' • ')
        : 'Aucune nouvelle notification';

      const sectionsHtml = [];

      if (isAdmin && recentRedCalls.length > 0) {
        sectionsHtml.push(`
          <div class="px-3.5 py-2 bg-rose-50 text-[10px] font-extrabold uppercase tracking-wider text-rose-800 flex items-center justify-between">
            <span>☎️ Appels Directeur (Téléphone Rouge)</span>
            <span>${recentRedCalls.length} récent(s)</span>
          </div>
          ${recentRedCalls.slice(0, 6).map(c => {
            const isFresh = !seenRedPhoneCallIds.has(String(c.id));
            return `
              <div class="p-3.5 ${isFresh ? 'bg-rose-50/60' : ''} flex items-start justify-between gap-3">
                <div class="min-w-0 space-y-0.5">
                  <div class="flex items-center gap-1.5">
                    ${isFresh ? '<span class="w-2 h-2 rounded-full bg-rose-600 shrink-0"></span>' : ''}
                    <span class="text-xs font-extrabold text-rose-900">☎️ ${esc(c.repNom)}</span>
                    <span class="text-[10px] text-slate-400">• ${esc(formatChatTime(c.at))}</span>
                  </div>
                  <div class="text-xs text-slate-700 break-words">${esc(c.reason || 'Aide closing deal')}</div>
                </div>
                ${isFresh ? `<button type="button" onclick="seenRedPhoneCallIds.add('${esc(c.id)}'); lsSet('korpo_seen_red_calls_v1', JSON.stringify(Array.from(seenRedPhoneCallIds).slice(-60))); renderUnifiedBellNotifications();" class="px-2.5 py-1 text-[10px] font-bold bg-rose-600 text-white rounded-lg shrink-0 cursor-pointer">Vu ✓</button>` : ''}
              </div>
            `;
          }).join('')}
        `);
      }

      if (myMentions.length > 0) {
        sectionsHtml.push(`
          <div class="px-3.5 py-2 bg-indigo-50/70 text-[10px] font-extrabold uppercase tracking-wider text-indigo-900 flex items-center justify-between">
            <span>💬 Mentions (@${esc(currentProfile.nom)})</span>
            <span>${unreadMentions.length} non lue(s)</span>
          </div>
          ${myMentions.slice(0, 12).map(m => {
            const isUnread = !readMentionIds.has(String(m.id));
            return `
              <button type="button" onclick="jumpToChatMessage('${esc(m.id)}')" class="w-full text-left p-3.5 hover:bg-indigo-50/60 transition flex items-start justify-between gap-3 ${isUnread ? 'bg-indigo-50/40' : ''}">
                <div class="min-w-0 space-y-0.5">
                  <div class="flex items-center gap-1.5">
                    ${isUnread ? '<span class="w-2 h-2 rounded-full bg-indigo-600 shrink-0"></span>' : ''}
                    <span class="text-xs font-extrabold text-slate-900">${esc(m.authorName)}</span>
                    <span class="text-[10px] text-slate-400">• ${esc(formatChatTime(m.at))}</span>
                  </div>
                  <div class="text-xs text-slate-700 line-clamp-2 break-words">${esc(m.text || '')}</div>
                </div>
                <span class="text-[11px] font-bold text-indigo-600 shrink-0">Voir →</span>
              </button>
            `;
          }).join('')}
        `);
      }

      if (isAdmin && pendingSales.length > 0) {
        sectionsHtml.push(`
          <div class="px-3.5 py-2 bg-amber-50/80 text-[10px] font-extrabold uppercase tracking-wider text-amber-900 flex items-center justify-between">
            <span>💵 Ventes à confirmer</span>
            <span>${pendingSales.length} dossier(s)</span>
          </div>
          ${pendingSales.slice(0, 10).map(v => {
            const repN = state.representants.find(r => r.id === v.representant_id)?.nom || '—';
            const entN = state.entreprises.find(e => e.id === v.entreprise_id)?.nom || '—';
            const cliN = state.clients.find(c => c.id === v.client_id)?.nom || '—';
            return `
              <div class="p-3.5 hover:bg-slate-50 transition flex items-start justify-between gap-3">
                <div class="min-w-0">
                  <div class="flex items-center gap-1.5">
                    <span class="text-xs font-extrabold text-slate-900">${esc(repN)}</span>
                    <span class="text-[10px] text-slate-400">• ${v.date_vente}</span>
                  </div>
                  <div class="text-xs font-semibold text-slate-700 truncate mt-0.5">${esc(entN)} — ${esc(cliN)}</div>
                  <div class="text-xs font-extrabold text-indigo-600 tabular-nums mt-0.5">${fmtMoney(v.valeur)}</div>
                </div>
                <div class="flex flex-col gap-1.5 shrink-0">
                  <button type="button" onclick="quickConfirmSale('${v.id}')" class="px-2.5 py-1 text-[11px] font-bold bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg shadow-xs transition">
                    ✓ Confirmer
                  </button>
                  <button type="button" onclick="openEditVenteDrawer('${v.id}', false); document.getElementById('admin-notif-dropdown').classList.add('hidden');" class="px-2.5 py-1 text-[10px] font-semibold bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg transition">
                    Inspecter
                  </button>
                </div>
              </div>
            `;
          }).join('')}
        `);
      }

      if (!sectionsHtml.length) {
        listEl.innerHTML = `<div class="p-6 text-center text-xs text-slate-400">Aucune notification pour le moment.</div>`;
      } else {
        listEl.innerHTML = sectionsHtml.join('');
      }
    }

    function renderAdminNotifications() {
      renderUnifiedBellNotifications();
    }

    window.quickConfirmSale = async (saleId) => {
      const targetV = (state.ventes || []).find(x => String(x.id) === String(saleId));
      const { error } = await sb.from('ventes').update({ statut_depot: 'depot_recu' }).eq('id', saleId);
      if (error) return toast(error.message, 'error');
      if (targetV && !isHitSale(targetV)) {
        const repN = state.representants.find(r => String(r.id) === String(targetV.representant_id))?.nom || 'Représentant';
        const entN = state.entreprises.find(e => String(e.id) === String(targetV.entreprise_id))?.nom || 'Korpo';
        const cliN = state.clients.find(c => String(c.id) === String(targetV.client_id))?.nom || '';
        triggerSaleCelebration(repN, cliN, entN, Number(targetV.valeur || 0), Number(targetV.commission_olivier || targetV.commission_rep || 0));
      }
      await recordSaleAudit(saleId, 'Dépôt confirmé (Reçu)', 'Validation rapide depuis les notifications', '✅');
      const mergedCount = await reconcilePreviousUnconfirmedDuplicateSales([saleId]);
      toast(mergedCount > 0
        ? `Vente confirmée ✓ (${mergedCount} doublon antérieur non confirmé retiré)`
        : 'Vente confirmée (Dépôt reçu) ✓');
      await loadAllData();
    };

    window.goToPendingSales = () => {
      document.getElementById('admin-notif-dropdown')?.classList.add('hidden');
      switchTab('tab-admin-ventes');
      const statSel = document.getElementById('admin-filter-statut');
      if (statSel) {
        statSel.value = 'depot_attente';
        renderAdminVentes();
      }
    };

    function triggerAdminNewSaleAlert(saleRow) {
      if (!saleRow || !saleRow.id) return;
      const saleKey = String(saleRow.id);
      if (knownPendingSaleIds.has(saleKey)) return;
      knownPendingSaleIds.add(saleKey);

      // Ignorer l'alerte sonore/push uniquement si l'admin lui-même vient de créer cette vente dans sa session actuelle
      if (window._justCreatedSaleClientSide && currentProfile && saleRow.representant_id === currentProfile.id) {
        window._justCreatedSaleClientSide = false;
        return;
      }

      const repN = state.representants.find(r => String(r.id) === String(saleRow.representant_id))?.nom || 'Un représentant';
      const entN = state.entreprises.find(e => String(e.id) === String(saleRow.entreprise_id))?.nom || '';
      const cliN = state.clients.find(c => String(c.id) === String(saleRow.client_id))?.nom || '';
      const msg = `${repN} a déclaré une vente ${entN}${cliN ? ` (${cliN})` : ''} de ${fmtMoney(saleRow.valeur)} — à confirmer`;

      playNotificationChime();
      toast(`🔔 ${msg}`);
      sendBrowserPushNotification('Nouvelle vente à confirmer — Korpo', msg);
    }

    function syncPendingBaselineAndNotify() {
      if (!currentProfile || currentProfile.role !== 'admin') return;
      const pending = (state.ventes || []).filter(v => v.statut_depot === 'a_confirmer');
      if (!hasInitializedPendingBaseline) {
        pending.forEach(v => knownPendingSaleIds.add(String(v.id)));
        hasInitializedPendingBaseline = true;
        return;
      }
      pending.forEach(v => {
        if (!knownPendingSaleIds.has(String(v.id))) {
          triggerAdminNewSaleAlert(v);
        }
      });
    }

    let checkingPendingNow = false;
    let adminVisibilityBound = false;

    async function checkNewPendingSalesNow() {
      if (!sb || !currentProfile || currentProfile.role !== 'admin') return;
      if (checkingPendingNow) return;
      checkingPendingNow = true;
      try {
        // 1) Requête légère (id + statut) ; la requête complète n'a lieu que s'il y a du nouveau
        const light = await fetchAllRows(() => sb.from('ventes').select('id,statut_depot').order('id'));
        const hasNew = light.some(v => v.statut_depot === 'a_confirmer' && !knownPendingSaleIds.has(String(v.id)));
        if (!hasNew) return;

        const [latestVentes, latestClients] = await Promise.all([
          fetchAllRows(() => sb.from('ventes').select('*').order('date_vente', { ascending: false }).order('created_at', { ascending: false }).order('id', { ascending: false })),
          fetchAllRows(() => sb.from('clients').select('*').order('nom').order('id'))
        ]);
        state.clients = latestClients;

        const newPending = latestVentes.filter(v => v.statut_depot === 'a_confirmer' && !knownPendingSaleIds.has(String(v.id)));
        if (newPending.length > 0) {
          state.ventes = latestVentes;
          state.mesVentes = state.ventes.map(mapAdminPersonalSaleRow);
          applySaleBonusOverridesToState();
          syncSalesIntoKpi();
          renderAdminNotifications();
          renderAdminVentes();
          renderAdminDashboard();
        renderAdminFinances();
          renderMesVentes();
          renderKpiDashboard();
          newPending.forEach(v => triggerAdminNewSaleAlert(v));
        }
      } catch (e) {
        console.warn('Vérification nouvelles ventes:', e.message);
      } finally {
        checkingPendingNow = false;
      }
    }

    function stopAdminBackgroundSync() {
      if (adminPollingTimer) { clearInterval(adminPollingTimer); adminPollingTimer = null; }
      if (chatPollingTimer) { clearInterval(chatPollingTimer); chatPollingTimer = null; }
      if (realtimeVentesChannel && sb) { try { sb.removeChannel(realtimeVentesChannel); } catch (e) {} }
      if (chatRealtimeChannel && sb) { try { sb.removeChannel(chatRealtimeChannel); } catch (e) {} }
      chatRealtimeChannel = null;
      realtimeVentesChannel = null;
      knownPendingSaleIds = new Set();
      hasInitializedPendingBaseline = false;
    }

    function subscribeToNewSalesRealtime() {
      if (!sb || !currentProfile || currentProfile.role !== 'admin') return;
      syncPendingBaselineAndNotify();

      if (!realtimeVentesChannel) {
        realtimeVentesChannel = sb
          .channel('admin-ventes-notifications-' + Date.now())
          .on('postgres_changes', { event: '*', schema: 'public', table: 'ventes' }, async () => {
            await checkNewPendingSalesNow();
          })
          .subscribe((status) => {
            if (status === 'CHANNEL_ERROR' || status === 'TIMED_OUT') {
              realtimeVentesChannel = null;
            }
          });
      }

      // Filet de sécurité toutes les 15 s (le temps réel fait l'essentiel) + dès le retour sur l'onglet
      if (!adminPollingTimer) {
        adminPollingTimer = setInterval(checkNewPendingSalesNow, 15000);
      }
      if (!adminVisibilityBound) {
        adminVisibilityBound = true;
        document.addEventListener('visibilitychange', () => {
          if (!document.hidden) checkNewPendingSalesNow();
        });
      }
    }

    function initAdminNotificationControls() {
      const btnBell = document.getElementById('btn-admin-notif');
      const dropdown = document.getElementById('admin-notif-dropdown');
      if (btnBell && dropdown) {
        btnBell.addEventListener('click', (e) => {
          e.stopPropagation();
          dropdown.classList.toggle('hidden');
        });
        document.addEventListener('click', (e) => {
          if (!dropdown.contains(e.target) && !btnBell.contains(e.target)) {
            dropdown.classList.add('hidden');
          }
        });
      }

      const btnPush = document.getElementById('btn-enable-browser-notif');
      if (btnPush) {
        const updatePushBtnState = () => {
          if (!('Notification' in window)) {
            btnPush.textContent = 'Non supporté';
            btnPush.disabled = true;
          } else if (Notification.permission === 'granted') {
            btnPush.textContent = '✓ Alertes actives';
            btnPush.classList.replace('bg-indigo-600', 'bg-emerald-600');
          }
        };
        updatePushBtnState();
        btnPush.addEventListener('click', async () => {
          if (!('Notification' in window)) return;
          const perm = await Notification.requestPermission();
          updatePushBtnState();
          if (perm === 'granted') {
            toast('Notifications bureau activées ! Vous recevrez une alerte dès qu’une vente est déclarée.');
          } else {
            toast('Autorisez les notifications dans les paramètres de votre navigateur.', 'error');
          }
        });
      }
    }

    // =========================================================================
    // PREUVES PHOTO DE DÉPÔT / CONTRAT & ANCIENNETÉ DES DOSSIERS (CHANTIER 2)
    // =========================================================================
    let saleProofsStore = lsGetJson('korpo_sale_proofs_v1', '{}');
    let pendingProofDataUrl = '';

    function saveSaleProofsStore() {
      try {
        lsSet('korpo_sale_proofs_v1', JSON.stringify(saleProofsStore));
      } catch (e) {
        console.warn('Stockage local saturé pour les aperçus photo');
      }
    }

    function compressImageToDataUrl(file, maxWidth = 640, quality = 0.62) {
      return new Promise((resolve, reject) => {
        const objectUrl = URL.createObjectURL(file);
        const img = new Image();
        img.onload = () => {
          try {
            const scale = Math.min(1, maxWidth / Math.max(img.width, img.height, 1));
            const canvas = document.createElement('canvas');
            canvas.width = Math.max(1, Math.round(img.width * scale));
            canvas.height = Math.max(1, Math.round(img.height * scale));
            const ctx = canvas.getContext('2d');
            ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
            const dataUrl = canvas.toDataURL('image/jpeg', quality);
            URL.revokeObjectURL(objectUrl);
            resolve(dataUrl);
          } catch (err) {
            URL.revokeObjectURL(objectUrl);
            reject(err);
          }
        };
        img.onerror = (err) => {
          URL.revokeObjectURL(objectUrl);
          reject(err);
        };
        img.src = objectUrl;
      });
    }

    function getDaysSince(dateStr) {
      if (!dateStr) return 0;
      const d = new Date(dateStr + 'T00:00:00');
      const now = new Date();
      return Math.max(0, Math.floor((now - d) / (1000 * 60 * 60 * 24)));
    }

    function isHitSale(v) {
      if (!v) return false;
      if (v._isHit) return true;
      if (Number(v.commission_rep) < 0 || Number(v.commission_olivier) < 0) return true;
      return /\[HIT(?::(?:REP|OLIVIER|BOTH))?\]/i.test(String(v.note || ''));
    }

    function hitBadgeHtml(v) {
      if (!isHitSale(v)) return '';
      const rawN = String(v.note || '');
      const m = rawN.match(/\[HIT:(REP|OLIVIER|BOTH)\]/i);
      const bearer = v._hitBearer || (m ? m[1].toLowerCase() : (
        Number(v.commission_rep) < 0 && Number(v.commission_olivier) >= 0 ? 'rep' :
        (Number(v.commission_olivier) < 0 && Number(v.commission_rep) >= 0 ? 'olivier' : 'both')
      ));
      const lbl = bearer === 'rep' ? 'HIT Rep' : (bearer === 'olivier' ? 'HIT Olivier' : 'HIT Rep+Oli');
      return `<span class="px-1.5 py-0.5 text-[10px] font-extrabold rounded-md bg-rose-100 text-rose-700 border border-rose-200" title="Remboursement / reprise de commission">🚨 ${lbl}</span>`;
    }

    function cleanDisplaySaleNote(noteStr) {
      return String(noteStr || '')
        .replace(/^\[(OTS|SUIVI)\]\s*/i, '').replace(/\[(REP|ENT):[^\]]+\]\s*/ig, '').replace(/\[SANSDEPOT\]\s*/ig, '')
        .replace(/\[HIT(?::(?:REP|OLIVIER|BOTH))?\]\s*/ig, '')
        .trim();
    }

    function agingBadgeForDeposit(v) {
      if (v.statut_depot === 'depot_recu') return '';
      const days = getDaysSince(v.date_vente);
      if (days >= 7) {
        return `<span class="ml-1.5 inline-flex items-center px-2 py-0.5 text-[10px] font-extrabold rounded-full bg-rose-100 text-rose-700 border border-rose-200" title="En attente depuis ${days} jours">⚠ ${days}j</span>`;
      }
      if (days >= 3) {
        return `<span class="ml-1.5 inline-flex items-center px-1.5 py-0.5 text-[10px] font-bold rounded-full bg-amber-100 text-amber-800">${days}j</span>`;
      }
      return '';
    }

    function getProofForSale(sale) {
      const r = getProofRef(sale);
      return isProofPath(r) ? (proofUrlCache.get(r)?.url || '') : r;
    }

    // Valeur brute stockée (chemin Storage, base64 ancien format ou cache local)
    function getProofRef(sale) {
      if (sale.preuve_url) return sale.preuve_url;
      const keyById = sale.id ? String(sale.id) : '';
      if (keyById && saleProofsStore[keyById]) return saleProofsStore[keyById];
      const compositeKey = `${sale.date_vente}_${sale.client_nom || ''}_${sale.valeur}`;
      return saleProofsStore[compositeKey] || '';
    }

    window.removeProofFromSale = async (saleId) => {
      const sid = String(saleId);
      const v = state.ventes.find(x => String(x.id) === sid) || state.mesVentes.find(x => String(x.id) === sid);
      if (!v) return;
      const ok = await askConfirm('Supprimer cette preuve ?', 'La photo de preuve sera retirée de cette vente.');
      if (!ok) return;

      const oldProofRef = getProofRef(v);
      delete saleProofsStore[sid];
      const cliNom = v.client_nom || state.clients.find(c => String(c.id) === String(v.client_id))?.nom || '';
      const compositeKey = `${v.date_vente}_${cliNom}_${v.valeur}`;
      delete saleProofsStore[compositeKey];
      saveSaleProofsStore();

      if (sb && currentProfile) {
        try {
          await sb.from('ventes').update({ preuve_url: null }).eq('id', v.id);
          if (isProofPath(oldProofRef)) await sb.storage.from(PROOF_BUCKET).remove([oldProofRef]);
        } catch (e) {}
      }
      await recordSaleAudit(sid, 'Preuve photo supprimée', `Retirée par ${currentProfile?.nom || 'Utilisateur'}`, '🗑️');
      document.getElementById('proof-modal')?.classList.add('hidden');
      toast('Preuve photo supprimée ✓');
      await loadAllData();
    };

    window.openProofModal = (saleId) => {
      const v = state.ventes.find(x => String(x.id) === String(saleId)) || state.mesVentes.find(x => String(x.id) === String(saleId));
      if (!v) return;
      const cli = v.client_nom || state.clients.find(c => c.id === v.client_id)?.nom || 'Client';
      const ent = v.entreprise_nom || state.entreprises.find(e => e.id === v.entreprise_id)?.nom || '';
      const proofSrc = getProofForSale(v);
      document.getElementById('proof-modal-title').textContent = `Preuve de dépôt — ${cli}`;
      document.getElementById('proof-modal-subtitle').textContent = `${ent} • ${v.date_vente} • ${fmtMoney(v.valeur)}`;
      const body = document.getElementById('proof-modal-body');
      const delBtn = document.getElementById('btn-delete-proof-modal');
      const canRemoveProof = Boolean(proofSrc) && (
        currentProfile?.role === 'admin' ||
        (String(v.representant_id) === String(currentProfile?.id) && v.statut_paiement !== 'paye')
      );
      if (delBtn) {
        delBtn.classList.toggle('hidden', !canRemoveProof);
        delBtn.onclick = canRemoveProof ? () => removeProofFromSale(v.id) : null;
      }
      if (proofSrc) {
        body.innerHTML = `<img src="${esc(proofSrc)}" alt="Preuve de dépôt" class="max-h-[70vh] rounded-xl shadow-md object-contain" />`;
      } else {
        body.innerHTML = `<div class="text-xs text-slate-500">Aucune photo jointe à ce dossier.</div>`;
      }
      document.getElementById('proof-modal').classList.remove('hidden');
    };

    window.attachProofToExistingSale = async (saleId, file) => {
      if (!file) return;
      try {
        toast('Compression de la photo...');
        const dataUrl = await compressImageToDataUrl(file, 1280, 0.72);
        const proofRef = await uploadProofDataUrl(dataUrl);
        const { error } = await sb.from('ventes').update({ preuve_url: proofRef }).eq('id', saleId);
        await recordSaleAudit(saleId, 'Preuve photo ajoutée', 'Capture / reçu téléversé sur le dossier', '📷');
        if (error) {
          toast('Photo enregistrée localement (vérifiez la colonne preuve_url dans Supabase)', 'error');
        } else {
          toast('Preuve photo ajoutée et synchronisée ✓');
        }
        await loadAllData();
      } catch (err) {
        toast('Impossible de lire cette photo.', 'error');
      }
    };

    function initProofUploadControls() {
      const fileIn = document.getElementById('vente-proof-file');
      const btnClear = document.getElementById('btn-clear-proof');
      const prevBox = document.getElementById('vente-proof-preview-box');
      const prevImg = document.getElementById('vente-proof-preview-img');
      const btnCloseProof = document.getElementById('btn-close-proof');

      if (btnCloseProof) {
        btnCloseProof.addEventListener('click', () => document.getElementById('proof-modal').classList.add('hidden'));
      }

      const camIn = document.getElementById('vente-camera-input');
      const statusBox = document.getElementById('vente-proof-status');
      const handleProofSelected = async (e) => {
        const file = e.target.files && e.target.files[0];
        if (!file) return;
        if (statusBox) {
          statusBox.textContent = 'Compression de la photo en cours...';
          statusBox.classList.remove('hidden');
        }
        try {
          pendingProofDataUrl = await compressImageToDataUrl(file, 1280, 0.72);
          const kb = Math.round((pendingProofDataUrl.length * 3) / 4 / 1024);
          prevImg.src = pendingProofDataUrl;
          prevBox.classList.remove('hidden');
          btnClear.classList.remove('hidden');
          if (statusBox) {
            statusBox.textContent = `✓ Preuve prête à enregistrer (${kb} Ko)`;
            statusBox.classList.remove('hidden');
          }
          toast(`Photo prête (${kb} Ko) ✓`);
        } catch (err) {
          pendingProofDataUrl = '';
          if (statusBox) statusBox.classList.add('hidden');
          toast('Format photo non supporté par ce navigateur. Essayez une capture d’écran.', 'error');
        }
      };
      if (fileIn) fileIn.addEventListener('change', handleProofSelected);
      if (camIn) camIn.addEventListener('change', handleProofSelected);

      if (btnClear) {
        btnClear.addEventListener('click', () => {
          pendingProofDataUrl = '';
          if (fileIn) fileIn.value = '';
          if (camIn) camIn.value = '';
          prevImg.src = '';
          prevBox.classList.add('hidden');
          btnClear.classList.add('hidden');
          if (statusBox) statusBox.classList.add('hidden');
        });
      }
    }

    // =========================================================================
    // CADENCE QUOTIDIENNE & SUIVI DES TARGETS / BONUS ÉQUIPE
    // =========================================================================
    function countBusinessDaysInMonth(ym, maxDay = null) {
      const [yStr, mStr] = ym.split('-');
      const y = parseInt(yStr, 10);
      const m = parseInt(mStr, 10) - 1;
      const daysInMonth = new Date(y, m + 1, 0).getDate();
      const limitDay = maxDay !== null ? Math.min(Math.max(1, maxDay), daysInMonth) : daysInMonth;
      let count = 0;
      for (let d = 1; d <= limitDay; d++) {
        const dow = new Date(y, m, d).getDay();
        if (dow !== 0 && dow !== 6) count++;
      }
      return count;
    }

    function getPreviousYearMonth(ym) {
      const [yStr, mStr] = ym.split('-');
      let y = parseInt(yStr, 10);
      let m = parseInt(mStr, 10) - 1;
      if (m < 1) { m = 12; y -= 1; }
      return `${y}-${String(m).padStart(2, '0')}`;
    }

    function formatYearMonthLabelFr(ym, shortMonth = false) {
      if (!ym || typeof ym !== 'string' || !ym.includes('-')) return String(ym || '');
      const [yStr, mStr] = ym.split('-');
      const shortNames = {
        '01': 'Janv.', '02': 'Févr.', '03': 'Mars', '04': 'Avr.',
        '05': 'Mai', '06': 'Juin', '07': 'Juill.', '08': 'Août',
        '09': 'Sept.', '10': 'Oct.', '11': 'Nov.', '12': 'Déc.'
      };
      if (shortMonth) {
        return `${shortNames[mStr] || mStr} ${yStr.slice(2)}`;
      }
      return `${MONTH_NAMES_FR[mStr] || mStr} ${yStr}`;
    }

    function formatDateLabelFr(ymdStr) {
      if (!ymdStr || typeof ymdStr !== 'string' || ymdStr.length < 10) return String(ymdStr || '');
      const y = ymdStr.slice(0, 4);
      const m = ymdStr.slice(5, 7);
      const d = parseInt(ymdStr.slice(8, 10), 10);
      const monthName = (MONTH_NAMES_FR[m] || m).toLowerCase();
      return `${d} ${monthName} ${y}`;
    }

    function renderExecutiveForecast() {
      const dashPeriod = readPeriodFilterState('dash-filter');
      const isSingleMonthMode = dashPeriod.mode === 'this_month' || dashPeriod.mode === 'specific_month';
      let targetYm = dashPeriod.singleMonth || currentMonthStr;
      const mObj = ensureKpiMonthStructure(targetYm);
      const isCurrentMonth = isSingleMonthMode && targetYm === currentMonthStr;

      // Déterminer les mois couverts par le filtre actif du tableau de bord
      let coveredMonths = [targetYm];
      if (!isSingleMonthMode) {
        const allKnownMonths = new Set([
          ...Object.keys(kpiStore).filter(isValidYearMonth),
          ...(state.ventes || []).map(v => getSaleStatsMonth(v)).filter(isValidYearMonth)
        ]);
        if (dashPeriod.mode === 'this_quarter') {
          coveredMonths = Array.from(allKnownMonths).filter(ym => getQuarterForMonth(ym) === dashPeriod.currentQuarter).sort();
        } else if (dashPeriod.mode === 'this_year') {
          coveredMonths = Array.from(allKnownMonths).filter(ym => ym.startsWith(dashPeriod.currentYear)).sort();
        } else if (dashPeriod.mode === 'future') {
          const futureSaleMonths = Array.from(new Set(
            (state.ventes || [])
              .map(v => getSaleTargetFutureMonth(v))
              .filter(isValidYearMonth)
          )).sort();
          coveredMonths = futureSaleMonths.length ? futureSaleMonths : [currentMonthStr];
          targetYm = coveredMonths[0];
        } else if (dashPeriod.mode === 'all') {
          coveredMonths = Array.from(allKnownMonths).sort();
        } else {
          coveredMonths = [targetYm];
        }
        if (!coveredMonths.length) coveredMonths = [targetYm];
      }

      const periodDisplayTitle = isSingleMonthMode
        ? formatYearMonthLabelFr(targetYm)
        : dashPeriod.label;

      const titleEl = document.getElementById('forecast-title');
      if (titleEl) titleEl.textContent = `Rythme quotidien & Suivi des bonus — ${periodDisplayTitle}`;

      const [yStr, mStr] = targetYm.split('-');
      const totalDaysInMonth = new Date(parseInt(yStr, 10), parseInt(mStr, 10), 0).getDate();
      const currentDayOfMonth = isCurrentMonth
        ? Math.min(today.getDate(), totalDaysInMonth)
        : (dashPeriod.mode === 'specific_date' && dashPeriod.dateVal
          ? Math.min(parseInt(dashPeriod.dateVal.slice(8, 10), 10) || 1, totalDaysInMonth)
          : totalDaysInMonth);

      let totalBizDays = 1;
      let elapsedBizDays = 1;
      let remainingBizDays = 0;

      if (isSingleMonthMode) {
        totalBizDays = Math.max(1, countBusinessDaysInMonth(targetYm));
        elapsedBizDays = Math.max(1, countBusinessDaysInMonth(targetYm, currentDayOfMonth));
        remainingBizDays = isCurrentMonth ? Math.max(0, totalBizDays - elapsedBizDays) : 0;
      } else if (dashPeriod.mode === 'specific_date') {
        totalBizDays = 1;
        elapsedBizDays = 1;
        remainingBizDays = 0;
      } else if (dashPeriod.mode === 'custom_range') {
        const sR = dashPeriod.startVal <= dashPeriod.endVal ? dashPeriod.startVal : dashPeriod.endVal;
        const eR = dashPeriod.startVal <= dashPeriod.endVal ? dashPeriod.endVal : dashPeriod.startVal;
        let rangeBiz = 0;
        if (sR && eR) {
          const cD = new Date(sR + 'T00:00:00');
          const eD = new Date(eR + 'T00:00:00');
          while (!isNaN(cD.getTime()) && !isNaN(eD.getTime()) && cD <= eD) {
            if (cD.getDay() !== 0 && cD.getDay() !== 6) rangeBiz++;
            cD.setDate(cD.getDate() + 1);
          }
        }
        totalBizDays = Math.max(1, rangeBiz);
        elapsedBizDays = totalBizDays;
        remainingBizDays = 0;
      } else if (dashPeriod.mode === 'this_week') {
        totalBizDays = 5;
        const dow = today.getDay();
        elapsedBizDays = dow === 0 ? 5 : Math.min(5, Math.max(1, dow));
        remainingBizDays = Math.max(0, 5 - elapsedBizDays);
      } else {
        totalBizDays = Math.max(1, coveredMonths.reduce((s, ym) => s + countBusinessDaysInMonth(ym), 0));
        elapsedBizDays = Math.max(1, coveredMonths.reduce((s, ym) => {
          if (ym === currentMonthStr) return s + countBusinessDaysInMonth(ym, today.getDate());
          if (ym < currentMonthStr) return s + countBusinessDaysInMonth(ym);
          return s;
        }, 0));
        remainingBizDays = Math.max(0, totalBizDays - elapsedBizDays);
      }

      // Ventes filtrées selon la période choisie dans le tableau de bord (exactement synchronisées avec les cartes du haut)
      const periodSales = (state.ventes || []).filter(v => matchesSalePeriod(v, dashPeriod, true));
      const rawSalesSigned = periodSales.reduce((s, v) => s + Number(v.valeur || 0), 0);
      const condensedBaselineSum = (dashPeriod.mode === 'this_week' || dashPeriod.mode === 'specific_date' || dashPeriod.mode === 'future')
        ? 0
        : coveredMonths.reduce((s, ym) => {
            const mBase = condensedBaselineStore?.[ym];
            if (!mBase || typeof mBase !== 'object') return s;
            return s + ALL_KPI_REPS.reduce((acc, rn) => acc + Number(mBase[rn]?.dollars || 0), 0);
          }, 0);
      const totSigned = Math.round((rawSalesSigned + condensedBaselineSum) * 100) / 100;

      // Objectif global proportionnel à la période sélectionnée
      if (dashPeriod.mode === 'custom_range' && dashPeriod.startVal) {
        targetYm = dashPeriod.startVal.slice(0, 7);
      }
      const monthMultiplier = dashPeriod.mode === 'custom_range'
        ? computeCustomRangeMonthFactor(dashPeriod.startVal, dashPeriod.endVal)
        : (dashPeriod.mode === 'specific_date'
          ? (1 / Math.max(1, countBusinessDaysInMonth(targetYm)))
          : (dashPeriod.mode === 'this_week' ? (5 / Math.max(1, countBusinessDaysInMonth(targetYm))) : coveredMonths.length));

      const activeForecastReps = getActiveKpiReps();
      const repTargetsMap = {};
      ALL_KPI_REPS.forEach(rn => {
        if (isSingleMonthMode) {
          repTargetsMap[rn] = getSalesTargetForRep(rn, targetYm);
        } else if (dashPeriod.mode === 'specific_date' || dashPeriod.mode === 'this_week' || dashPeriod.mode === 'custom_range') {
          repTargetsMap[rn] = Math.round(getSalesTargetForRep(rn, targetYm) * monthMultiplier);
        } else {
          repTargetsMap[rn] = coveredMonths.reduce((s, ym) => s + getSalesTargetForRep(rn, ym), 0);
        }
      });
      const totObjective = ALL_KPI_REPS.reduce((s, rn) => s + (repTargetsMap[rn] || 0), 0);

      // Cadence par jour ouvrable
      const dailyPace = Math.round(totSigned / elapsedBizDays);
      const remainingToTarget = Math.max(0, totObjective - totSigned);
      const requiredPace = remainingBizDays > 0
        ? Math.round(remainingToTarget / remainingBizDays)
        : (remainingToTarget > 0 ? remainingToTarget : 0);

      // Comparatif période précédente
      const prevYm = getPreviousYearMonth(targetYm);
      const prevLabelFr = formatYearMonthLabelFr(prevYm);
      const prevMonthObj = kpiStore[prevYm];
      const prevSalesSameCutoff = (state.ventes || [])
        .filter(v => {
          if (getSaleStatsMonth(v) !== prevYm) return false;
          if (!isCurrentMonth) return true;
          const dayNum = v.date_vente && v.date_vente.length >= 10 ? parseInt(v.date_vente.slice(8, 10), 10) : 1;
          return dayNum <= currentDayOfMonth;
        })
        .reduce((s, v) => s + Number(v.valeur || 0), 0);
      const prevTotalFallback = !isCurrentMonth ? Number(prevMonthObj?.team?.['VENTE $']?.cumul || 0) : 0;
      const prevComparable = Math.max(prevSalesSameCutoff, prevTotalFallback);
      const diffVsPrev = totSigned - prevComparable;
      const diffVsPrevPct = prevComparable > 0 ? Math.round((diffVsPrev / prevComparable) * 100) : null;

      // Construction des lignes par représentant (Targets & Bonus)
      let unlockedCount = 0;
      const repRowsData = ALL_KPI_REPS.map(rn => {
        const repObj = state.representants.find(r => r.nom === rn);
        const repSalesList = periodSales.filter(v => {
          if (repObj && String(v.representant_id) === String(repObj.id)) return true;
          const rN = state.representants.find(r => String(r.id) === String(v.representant_id))?.nom || resolveSaleArchivedRepName(v) || v.rep_nom || '';
          return resolveKpiRepName(rN) === rn;
        });
        const signedFromSales = repSalesList.reduce((s, v) => s + Number(v.valeur || 0), 0);
        const signedFromCondensed = (dashPeriod.mode === 'this_week' || dashPeriod.mode === 'specific_date' || dashPeriod.mode === 'future')
          ? 0
          : coveredMonths.reduce((s, ym) => s + Number(condensedBaselineStore?.[ym]?.[rn]?.dollars || 0), 0);
        const signed = Math.round((signedFromSales + signedFromCondensed) * 100) / 100;

        const target = repTargetsMap[rn] || 150000;
        const pct = target > 0 ? Math.round((signed / target) * 100) : 0;
        const gap = Math.max(0, target - signed);
        const bonusDec = getBonusPctForRep(rn, targetYm);
        const bonusStr = formatBonusPctShort(bonusDec);
        const isAdminRep = repObj?.role === 'admin';
        const unlocked = !isAdminRep && target > 0 && signed >= target;
        if (unlocked) unlockedCount++;

        const commEarned = isAdminRep
          ? periodSales.reduce((s, v) => s + Number(v.commission_olivier || 0), 0)
          : repSalesList.reduce((s, v) => s + Number(v.commission_rep || 0), 0);

        return { rn, signed, target, pct, gap, bonusStr, isAdminRep, unlocked, commEarned };
      }).filter(item => !isRepArchived(item.rn) || item.signed > 0 || item.target > 0).sort((a, b) => b.pct - a.pct);

      const topBadgeLabel = document.getElementById('forecast-cadence-badge-label');
      if (topBadgeLabel) {
        topBadgeLabel.textContent = (isCurrentMonth || remainingBizDays > 0)
          ? `Cadence (${elapsedBizDays}j / ${totalBizDays}j ouvrables)`
          : `Cadence moyenne (${totalBizDays}j ouvrable${totalBizDays > 1 ? 's' : ''})`;
      }
      const projEl = document.getElementById('forecast-total-projected');
      if (projEl) projEl.textContent = `${fmtMoney0(dailyPace)} / jour`;

      const progLbl = document.getElementById('forecast-month-progress-label');
      if (progLbl) {
        if (isCurrentMonth) {
          progLbl.textContent = `${formatDateLabelFr(currentDateStr)} • ${remainingBizDays} jour(s) ouvrable(s) restant(s)`;
        } else if (dashPeriod.mode === 'specific_date') {
          progLbl.textContent = `Journée du ${formatDateLabelFr(dashPeriod.dateVal)}`;
        } else if (isSingleMonthMode) {
          progLbl.textContent = `${formatYearMonthLabelFr(targetYm)} (${totalBizDays} jours ouvrables)`;
        } else {
          progLbl.textContent = `${periodDisplayTitle} (${coveredMonths.length} mois • ${elapsedBizDays}j ouvrables)`;
        }
      }

      const cardsEl = document.getElementById('forecast-kpi-cards');
      if (cardsEl) {
        const diffSign = diffVsPrev >= 0 ? '+' : '';
        const diffColor = diffVsPrev >= 0 ? 'text-emerald-300' : 'text-rose-300';
        cardsEl.innerHTML = `
          <div class="bg-white/10 rounded-xl p-3.5 border border-white/10 space-y-1">
            <div class="text-[11px] text-indigo-200 font-semibold">1. Cadence actuelle</div>
            <div class="text-xl font-extrabold text-white tabular-nums">${fmtMoney0(dailyPace)} <span class="text-xs font-normal text-slate-300">/ jour</span></div>
            <div class="text-[11px] text-slate-300 tabular-nums">${fmtMoney0(totSigned)} sur ${elapsedBizDays}j ouvrable(s)</div>
          </div>
          <div class="bg-white/10 rounded-xl p-3.5 border border-white/10 space-y-1">
            <div class="text-[11px] text-indigo-200 font-semibold">2. Cadence requise (Target)</div>
            <div class="text-xl font-extrabold ${remainingToTarget === 0 ? 'text-emerald-400' : 'text-amber-300'} tabular-nums">
              ${remainingToTarget === 0 ? 'Objectif atteint !' : `${fmtMoney0(requiredPace)} / jour`}
            </div>
            <div class="text-[11px] text-slate-300 tabular-nums">
              ${remainingToTarget === 0 ? `Surplus : + ${fmtMoney0(totSigned - totObjective)}` : `Reste ${fmtMoney0(remainingToTarget)}${remainingBizDays > 0 ? ` (${remainingBizDays}j rest.)` : ''}`}
            </div>
          </div>
          <div class="bg-white/10 rounded-xl p-3.5 border border-white/10 space-y-1">
            <div class="text-[11px] text-indigo-200 font-semibold">3. Vs ${esc(prevLabelFr)}</div>
            <div class="text-xl font-extrabold ${diffColor} tabular-nums">${diffSign}${fmtMoney0(diffVsPrev)}</div>
            <div class="text-[11px] text-slate-300 tabular-nums">
              ${prevComparable > 0 ? `${diffSign}${diffVsPrevPct}% (${esc(prevLabelFr)} : ${fmtMoney0(prevComparable)})` : `Aucune donnée en ${esc(prevLabelFr)}`}
            </div>
          </div>
          <div class="bg-white/10 rounded-xl p-3.5 border border-white/10 space-y-1">
            <div class="text-[11px] text-indigo-200 font-semibold">4. Bonus débloqués</div>
            <div class="text-xl font-extrabold text-emerald-400 tabular-nums">${unlockedCount} / ${activeForecastReps.filter(n => n !== 'Olivier').length} reps</div>
            <div class="text-[11px] text-slate-300">Bonus actif dès 100 % du target</div>
          </div>
        `;
      }

      const tbodyEl = document.getElementById('tbody-forecast-reps');
      if (tbodyEl) {
        tbodyEl.innerHTML = repRowsData.map(r => {
          const statusBadgeHtml = r.isAdminRep
            ? `<span class="px-2 py-0.5 rounded-full bg-indigo-500/20 text-indigo-200 font-bold text-[10px]">Direction</span>`
            : (r.unlocked
              ? `<span class="px-2.5 py-0.5 rounded-full bg-emerald-500/25 text-emerald-300 border border-emerald-400/40 font-extrabold text-[10px]">🔥 +${esc(r.bonusStr)} ACTIF</span>`
              : `<span class="px-2 py-0.5 rounded-full bg-white/10 text-slate-300 font-semibold text-[10px]">+${esc(r.bonusStr)} à débloquer</span>`);

          return `
            <tr class="hover:bg-white/5 transition">
              <td class="py-2.5 px-4 font-bold text-white">${esc(r.rn)}</td>
              <td class="py-2.5 px-3 text-right tabular-nums font-bold text-slate-100">${fmtMoney0(r.signed)}</td>
              <td class="py-2.5 px-3 text-right tabular-nums text-slate-300">${fmtMoney0(r.target)}</td>
              <td class="py-2.5 px-3 text-center tabular-nums">
                <span class="font-extrabold ${r.pct >= 100 ? 'text-emerald-400' : (r.pct >= 65 ? 'text-amber-300' : 'text-slate-200')}">${r.pct} %</span>
              </td>
              <td class="py-2.5 px-3 text-right tabular-nums ${r.gap === 0 ? 'text-emerald-300 font-bold' : 'text-slate-300'}">
                ${r.gap === 0 ? '✓ Atteint' : fmtMoney0(r.gap)}
              </td>
              <td class="py-2.5 px-3 text-center">${statusBadgeHtml}</td>
              <td class="py-2.5 px-4 text-right tabular-nums font-extrabold text-emerald-300">${fmtMoney(r.commEarned)}</td>
            </tr>
          `;
        }).join('');
      }

      const mobileBox = document.getElementById('mobile-forecast-reps');
      if (mobileBox) {
        mobileBox.innerHTML = repRowsData.map(r => {
          const statusTag = r.isAdminRep
            ? `<span class="px-2 py-0.5 rounded-full bg-indigo-500/20 text-indigo-200 font-bold text-[10px]">Direction</span>`
            : (r.unlocked
              ? `<span class="px-2 py-0.5 rounded-full bg-emerald-500/25 text-emerald-300 border border-emerald-400/40 font-extrabold text-[10px]">🔥 +${esc(r.bonusStr)} ACTIF</span>`
              : `<span class="px-2 py-0.5 rounded-full bg-white/10 text-slate-300 font-semibold text-[10px]">Reste ${fmtMoney0(r.gap)}</span>`);

          return `
            <div class="bg-white/10 rounded-xl border border-white/10 p-3.5 space-y-2">
              <div class="flex items-center justify-between gap-2">
                <div class="font-extrabold text-sm text-white">${esc(r.rn)}</div>
                ${statusTag}
              </div>
              <div class="flex items-center justify-between text-xs tabular-nums">
                <span class="text-slate-300">${fmtMoney0(r.signed)} / ${fmtMoney0(r.target)}</span>
                <span class="font-extrabold ${r.pct >= 100 ? 'text-emerald-400' : 'text-amber-300'}">${r.pct} %</span>
              </div>
              <div class="w-full h-1.5 bg-white/10 rounded-full overflow-hidden">
                <div class="h-full ${r.pct >= 100 ? 'bg-emerald-400' : 'bg-indigo-400'} rounded-full" style="width: ${Math.min(100, r.pct)}%"></div>
              </div>
              <div class="flex items-center justify-between text-[11px] text-slate-300 pt-0.5 tabular-nums">
                <span>Commission cumulée :</span>
                <strong class="text-emerald-300">${fmtMoney(r.commEarned)}</strong>
              </div>
            </div>
          `;
        }).join('');
      }
    }

    function initKpiControls() {
      initProofUploadControls();
      initAdminNotificationControls();
      initClientTargetsControls();
      rebuildKpiStore();
      populateKpiMonthSelect(false);
      const selMonth = document.getElementById('kpi-select-month');

      selMonth.addEventListener('change', () => {
        const v = selMonth.value || '';
        const modeEl = document.getElementById('kpi-period-mode');
        if (modeEl) {
          if (v === currentMonthStr) modeEl.value = 'this_month';
          else if (v === `QUARTER_${getQuarterForMonth(currentMonthStr)}`) modeEl.value = 'this_quarter';
          else if (v === `YEAR_${today.getFullYear()}`) modeEl.value = 'this_year';
          else modeEl.value = 'select_dropdown';
        }
        document.getElementById('kpi-custom-range-box')?.classList.add('hidden');
        if (v.startsWith('YEAR_') || v.startsWith('QUARTER_')) {
          document.getElementById('kpi-editor-panel').classList.add('hidden');
        } else {
          syncKpiEditorInputs();
        }
        renderKpiDashboard();
      });
      document.getElementById('kpi-select-rep')?.addEventListener('change', renderKpiDashboard);

      document.getElementById('btn-kpi-edit-toggle')?.addEventListener('click', () => {
        if (currentProfile?.role !== 'admin') return;
        const selM = document.getElementById('kpi-select-month');
        if (selM.value.startsWith('YEAR_')) {
          const months = Object.keys(kpiStore).filter(isValidYearMonth).sort().reverse();
          if (months.length) {
            selM.value = months[0];
            renderKpiDashboard();
            toast(`Basculé sur ${kpiStore[months[0]].label} pour modifier les RDV`);
          }
        }
        const panel = document.getElementById('kpi-editor-panel');
        const wasHidden = panel.classList.contains('hidden');
        panel.classList.toggle('hidden');
        if (wasHidden) {
          const activeFilterRep = document.getElementById('kpi-select-rep')?.value;
          const editRepSel = document.getElementById('kpi-edit-rep');
          if (activeFilterRep && activeFilterRep !== 'ALL' && ALL_KPI_REPS.includes(activeFilterRep)) {
            editRepSel.value = activeFilterRep;
          }
        }
        syncKpiEditorInputs();
      });
      document.getElementById('btn-kpi-close-editor')?.addEventListener('click', () => {
        document.getElementById('kpi-editor-panel').classList.add('hidden');
      });

      document.getElementById('kpi-edit-rep')?.addEventListener('change', syncKpiEditorInputs);

      document.getElementById('form-kpi-edit')?.addEventListener('submit', async (e) => {
        e.preventDefault();
        if (currentProfile?.role !== 'admin') return toast('Action réservée aux administrateurs.', 'error');
        const ym = document.getElementById('kpi-select-month').value;
        if (!isValidYearMonth(ym)) return toast('Sélectionnez un mois précis.', 'error');
        const rep = document.getElementById('kpi-edit-rep').value;
        const totalBooke = Math.max(0, parseFloat(document.getElementById('kpi-in-booke').value) || 0);
        const totalRdv = Math.max(0, parseFloat(document.getElementById('kpi-in-rdv').value) || 0);

        const baseEntry = condensedBaselineStore[ym]?.[rep] || { booke: 0, rdv: 0, ots: 0, suivi: 0, dollars: 0 };
        const extraBooke = totalBooke - Number(baseEntry.booke || 0);
        const extraRdv = totalRdv - Number(baseEntry.rdv || 0);
        const nowIso = new Date().toISOString();

        const rdvPayload = { booke: extraBooke, rdv: extraRdv, updatedAt: nowIso };
        monthlyRdvOverrides[`${ym}_${rep}`] = rdvPayload;
        saveMonthlyRdvOverrides();
        await persistMonthlyRdvToCloud(ym, rep, rdvPayload);

        rebuildKpiStore();
        renderKpiDashboard();
        toast(`RDV mis à jour pour ${rep} (${kpiStore[ym]?.label || ym}) ✓`);
      });

      document.getElementById('btn-kpi-export-csv')?.addEventListener('click', exportKpiCsv);
    }

    function syncKpiEditorInputs() {
      const ym = document.getElementById('kpi-select-month').value;
      const mObj = ensureKpiMonthStructure(ym);
      if (!mObj) return;

      const editRepSel = document.getElementById('kpi-edit-rep');
      if (currentProfile && currentProfile.role !== 'admin' && ALL_KPI_REPS.includes(currentProfile.nom)) {
        editRepSel.value = currentProfile.nom;
        editRepSel.disabled = true;
      } else {
        editRepSel.disabled = false;
      }

      const rep = editRepSel.value;
      const rObj = mObj.reps[rep];
      if (!rObj) return;

      document.getElementById('kpi-in-booke').value = rObj['Booké'].cumul || 0;
      document.getElementById('kpi-in-rdv').value = rObj['RDV Réalisé'].cumul || 0;
    }

    function statusBadge(cumul, obj, isPct = false) {
      if (!obj || obj <= 0) {
        return `<span class="px-2 py-0.5 text-[11px] font-semibold rounded-full bg-slate-100 text-slate-500">—</span>`;
      }
      const ratio = cumul / obj;
      if (ratio >= 0.999) {
        return `<span class="px-2.5 py-0.5 text-[11px] font-bold rounded-full bg-emerald-50 text-emerald-700 border border-emerald-200">Réussi</span>`;
      }
      if (ratio >= 0.65) {
        return `<span class="px-2.5 py-0.5 text-[11px] font-bold rounded-full bg-amber-50 text-amber-700 border border-amber-200">En progression</span>`;
      }
      return `<span class="px-2.5 py-0.5 text-[11px] font-bold rounded-full bg-rose-50 text-rose-700 border border-rose-200">Non réussi</span>`;
    }

    // =========================================================================
    // MOTEUR DÉTERMINISTE DU TABLEAU KPI (BASE CONDENSÉE + FICHIER DE VENTES)
    // =========================================================================
    const MONTH_NAMES_FR = {
      '01': 'Janvier', '02': 'Février', '03': 'Mars', '04': 'Avril',
      '05': 'Mai', '06': 'Juin', '07': 'Juillet', '08': 'Août',
      '09': 'Septembre', '10': 'Octobre', '11': 'Novembre', '12': 'Décembre'
    };

    function getWeekIndexFromDate(dateStr, numWeeks = 4) {
      if (!dateStr || dateStr.length < 10) return 0;
      const day = parseInt(dateStr.slice(8, 10), 10) || 1;
      const idx = Math.floor((day - 1) / 7);
      return Math.min(Math.max(0, idx), numWeeks - 1);
    }

    function createEmptyMonthKpiObject(ym) {
      const [y, m] = ym.split('-');
      const label = `${MONTH_NAMES_FR[m] || m} ${y}`;
      const numWeeks = 4;
      const makeEmptyRep = (objDol = 150000, objClose = 35) => ({
        'Booké': { weeks: [0, 0, 0, 0], cumul: 0, objectif: 80 },
        'RDV Réalisé': { weeks: [0, 0, 0, 0], cumul: 0, objectif: 60 },
        'Ventes OTS': { weeks: [0, 0, 0, 0], cumul: 0, objectif: 10 },
        'Ventes SUIVI': { weeks: [0, 0, 0, 0], cumul: 0, objectif: 20 },
        'Vente $': { weeks: [0, 0, 0, 0], cumul: 0, objectif: objDol },
        '% Close': { weeks: [0, 0, 0, 0], cumul: 0, objectif: objClose }
      });
      const makeEmptyEnt = () => ({
        reps: {
          'Olivier': { weeks: [0, 0, 0, 0], cumul: 0, objectif: 0 },
          'Jean-Sébastien': { weeks: [0, 0, 0, 0], cumul: 0, objectif: 0 },
          'Gabryel': { weeks: [0, 0, 0, 0], cumul: 0, objectif: 0 },
          'Samuel': { weeks: [0, 0, 0, 0], cumul: 0, objectif: 0 },
          'Alexandre': { weeks: [0, 0, 0, 0], cumul: 0, objectif: 0 }
        },
        total: { weeks: [0, 0, 0, 0], cumul: 0, objectif: 0 }
      });
      return {
        ym,
        label,
        num_weeks: numWeeks,
        reps: {
          'Olivier': makeEmptyRep(118000, 40),
          'Jean-Sébastien': makeEmptyRep(250000, 32),
          'Gabryel': makeEmptyRep(180000, 35),
          'Samuel': makeEmptyRep(200000, 33),
          'Alexandre': makeEmptyRep(150000, 35)
        },
        team: {
          'Booké': { weeks: [0, 0, 0, 0], cumul: 0, objectif: 200 },
          'RDV RÉA': { weeks: [0, 0, 0, 0], cumul: 0, objectif: 150 },
          'VENTE OTS': { weeks: [0, 0, 0, 0], cumul: 0, objectif: 27 },
          'VENTE SUIVI': { weeks: [0, 0, 0, 0], cumul: 0, objectif: 49 },
          'VENTE TOTAL': { weeks: [0, 0, 0, 0], cumul: 0, objectif: 76 },
          'VENTE $': { weeks: [0, 0, 0, 0], cumul: 0, objectif: 568000 }
        },
        entreprises: {
          'Toitex': makeEmptyEnt(),
          'MonSpray': makeEmptyEnt(),
          'VerreVDM': makeEmptyEnt(),
          'Groupe Diamant Pro': makeEmptyEnt(),
          'MaClim': makeEmptyEnt()
        }
      };
    }

    function ensureKpiMonthStructure(ym) {
      if (!isValidYearMonth(ym)) {
        ym = (typeof currentMonthStr !== 'undefined' && isValidYearMonth(currentMonthStr)) ? currentMonthStr : '2026-11';
      }
      if (!kpiStore[ym]) {
        kpiStore[ym] = createEmptyMonthKpiObject(ym);
      }
      return kpiStore[ym];
    }

    function resolveKpiRepName(raw) {
      if (!raw) return '';
      const norm = normalizeImportedRepName(raw);
      if (norm) return norm;
      const trimmed = String(raw).trim();
      const exact = ALL_KPI_REPS.find(r => r.toLowerCase() === trimmed.toLowerCase());
      return exact || trimmed;
    }

    function rebuildKpiStore() {
      const allVentes = [...(state.ventes || [])].sort((a, b) => {
        const dA = a.date_vente || `${a.mois_commission || '9999-99'}-01`;
        const dB = b.date_vente || `${b.mois_commission || '9999-99'}-01`;
        if (dA !== dB) return dA.localeCompare(dB);
        return String(a.created_at || a.id || '').localeCompare(String(b.created_at || b.id || ''));
      });

      // 1. Collecter tous les mois pertinents (fichier condensé + fichier de ventes + RDV saisis + mois courant)
      const validMonths = new Set();
      if (typeof currentMonthStr !== 'undefined' && isValidYearMonth(currentMonthStr)) {
        validMonths.add(currentMonthStr);
      }
      Object.keys(condensedBaselineStore || {}).forEach(ym => {
        if (isValidYearMonth(ym)) validMonths.add(ym);
      });
      allVentes.forEach(v => {
        const ym = getSaleStatsMonth(v);
        if (isValidYearMonth(ym)) validMonths.add(ym);
      });
      (state.classement || []).forEach(row => {
        if (isValidYearMonth(row.mois_commission)) validMonths.add(row.mois_commission);
      });
      Object.keys(monthlyRdvOverrides || {}).forEach(k => {
        const ym = k.split('_')[0];
        if (isValidYearMonth(ym)) validMonths.add(ym);
      });
      Object.entries(dailyRdvLogsStore || {}).forEach(([k, entry]) => {
        const dStr = entry?.date || k.split('__')[0] || '';
        const ym = dStr.slice(0, 7);
        if (isValidYearMonth(ym)) validMonths.add(ym);
      });
      (kpiHebdoRows || []).forEach(r => {
        if (Number(r.semaine_index) >= 0 && isValidYearMonth(r.mois_commission)) {
          validMonths.add(r.mois_commission);
        }
      });

      // 2. Reconstruire `kpiStore` à neuf sans aucun résidu de cache
      const freshStore = {};
      validMonths.forEach(ym => {
        freshStore[ym] = createEmptyMonthKpiObject(ym);
      });

      // 3. Appliquer la base du fichier condensé (si importé)
      Object.entries(condensedBaselineStore || {}).forEach(([ym, monthData]) => {
        if (!freshStore[ym] || !monthData || typeof monthData !== 'object') return;
        const mObj = freshStore[ym];
        ALL_KPI_REPS.forEach(rn => {
          const entry = monthData[rn];
          if (!entry || typeof entry !== 'object') return;
          mObj.reps[rn]['Booké'].weeks[0] = Number(entry.booke || 0);
          mObj.reps[rn]['RDV Réalisé'].weeks[0] = Number(entry.rdv || 0);
          mObj.reps[rn]['Ventes OTS'].weeks[0] = Number(entry.ots || 0);
          mObj.reps[rn]['Ventes SUIVI'].weeks[0] = Number(entry.suivi || 0);
          mObj.reps[rn]['Vente $'].weeks[0] = Math.round(Number(entry.dollars || 0) * 100) / 100;
        });

        // Entreprises partenaires éventuellement présentes dans le fichier condensé
        if (monthData._entreprises && typeof monthData._entreprises === 'object') {
          Object.entries(monthData._entreprises).forEach(([entName, repMap]) => {
            if (!mObj.entreprises[entName]) {
              mObj.entreprises[entName] = {
                reps: Object.fromEntries(ALL_KPI_REPS.map(rn => [rn, { weeks: [0, 0, 0, 0], cumul: 0, objectif: 0 }])),
                total: { weeks: [0, 0, 0, 0], cumul: 0, objectif: 0 }
              };
            }
            ALL_KPI_REPS.forEach(rn => {
              const val = Number(repMap?.[rn] || 0);
              if (val > 0) {
                mObj.entreprises[entName].reps[rn].weeks[0] = val;
              }
            });
          });
        }
      });

      // 4. Appliquer les RDV saisis manuellement (monthlyRdvOverrides ou kpi_hebdo semaine_index >= 0) + les saisies journalières des représentants
      Object.entries(monthlyRdvOverrides || {}).forEach(([key, rdvOv]) => {
        const [ym, rn] = key.split('_');
        if (!freshStore[ym] || !freshStore[ym].reps[rn] || !rdvOv) return;
        freshStore[ym].reps[rn]['Booké'].weeks[0] = Math.max(0, (freshStore[ym].reps[rn]['Booké'].weeks[0] || 0) + Number(rdvOv.booke || 0));
        freshStore[ym].reps[rn]['RDV Réalisé'].weeks[0] = Math.max(0, (freshStore[ym].reps[rn]['RDV Réalisé'].weeks[0] || 0) + Number(rdvOv.rdv || 0));
      });

      Object.entries(dailyRdvLogsStore || {}).forEach(([key, dEntry]) => {
        if (!dEntry || typeof dEntry !== 'object') return;
        const parts = key.split('__');
        const dStr = dEntry.date || parts[0] || '';
        const rn = resolveKpiRepName(dEntry.rep || parts[1] || '');
        const ym = dStr.slice(0, 7);
        if (!isValidYearMonth(ym) || !freshStore[ym] || !freshStore[ym].reps[rn]) return;
        const wIdx = getWeekIndexFromDate(dStr, freshStore[ym].num_weeks || 4);
        freshStore[ym].reps[rn]['Booké'].weeks[wIdx] = Math.max(0, (freshStore[ym].reps[rn]['Booké'].weeks[wIdx] || 0) + Number(dEntry.booke || 0));
        freshStore[ym].reps[rn]['RDV Réalisé'].weeks[wIdx] = Math.max(0, (freshStore[ym].reps[rn]['RDV Réalisé'].weeks[wIdx] || 0) + Number(dEntry.rdv || 0));
      });

      (kpiHebdoRows || []).forEach(row => {
        const wIdx = Number(row.semaine_index);
        if (wIdx < 0 || wIdx >= 4) return;
        const ym = row.mois_commission;
        const rn = resolveKpiRepName(row.representant_nom);
        if (!freshStore[ym] || !freshStore[ym].reps[rn]) return;
        // Ignorer si déjà géré par monthlyRdvOverrides pour ce mois/rep
        if (monthlyRdvOverrides[`${ym}_${rn}`]) return;
        freshStore[ym].reps[rn]['Booké'].weeks[wIdx] += Number(row.booke || 0);
        freshStore[ym].reps[rn]['RDV Réalisé'].weeks[wIdx] += Number(row.rdv_realise || 0);
      });

      // 5. Additionner chaque vente active du fichier de ventes (`state.ventes`)
      allVentes.forEach(v => {
        const ym = getSaleStatsMonth(v);
        if (!freshStore[ym]) return;
        const mObj = freshStore[ym];
        const nw = mObj.num_weeks || 4;

        const repObj = state.representants.find(r => String(r.id) === String(v.representant_id));
        const entObj = state.entreprises.find(e => String(e.id) === String(v.entreprise_id));
        const repName = resolveKpiRepName(repObj?.nom || resolveSaleArchivedRepName(v) || v.rep_nom || '');
        const entName = (entObj?.nom || resolveSaleArchivedEntName(v) || v.entreprise_nom || '').trim();
        if (!repName) return;

        const val = Number(v.valeur || 0);
        const wIdx = getWeekIndexFromDate(v.date_vente, nw);
        const noteUp = (v.note || '').toUpperCase();
        const isSuivi = noteUp.includes('[SUIVI]');

        if (mObj.reps[repName]) {
          mObj.reps[repName]['Vente $'].weeks[wIdx] = Math.round(((mObj.reps[repName]['Vente $'].weeks[wIdx] || 0) + val) * 100) / 100;
          if (isSuivi) {
            mObj.reps[repName]['Ventes SUIVI'].weeks[wIdx] += 1;
          } else {
            mObj.reps[repName]['Ventes OTS'].weeks[wIdx] += 1;
          }
        }

        if (entName) {
          if (!mObj.entreprises[entName]) {
            mObj.entreprises[entName] = {
              reps: Object.fromEntries(ALL_KPI_REPS.map(rn => [rn, { weeks: Array(nw).fill(0), cumul: 0, objectif: 0 }])),
              total: { weeks: Array(nw).fill(0), cumul: 0, objectif: 0 }
            };
          }
          if (!mObj.entreprises[entName].reps[repName]) {
            mObj.entreprises[entName].reps[repName] = { weeks: Array(nw).fill(0), cumul: 0, objectif: 0 };
          }
          mObj.entreprises[entName].reps[repName].weeks[wIdx] = Math.round(((mObj.entreprises[entName].reps[repName].weeks[wIdx] || 0) + val) * 100) / 100;
        }
      });

      // 5b. Pour un représentant connecté : synchroniser avec précision les statistiques des collègues
      const isAdminUser = currentProfile?.role === 'admin';
      const myRepName = resolveKpiRepName(currentProfile?.nom || '');
      const classementByYmAndRep = new Map();
      (state.classement || []).forEach(row => {
        const ym = row.mois_commission;
        const repName = resolveKpiRepName(row.representant_nom || '');
        if (!isValidYearMonth(ym) || !repName) return;
        const k = `${ym}__${repName}`;
        classementByYmAndRep.set(k, (classementByYmAndRep.get(k) || 0) + Number(row.total_ventes || 0));
      });

      if (!isAdminUser) {
        // Ajouter les mois présents dans teamKpiSnapshotStore s'ils ne sont pas encore dans freshStore
        Object.keys(teamKpiSnapshotStore || {}).forEach(ym => {
          if (isValidYearMonth(ym) && !freshStore[ym]) {
            freshStore[ym] = createEmptyMonthKpiObject(ym);
          }
        });

        Object.entries(freshStore).forEach(([ym, mObj]) => {
          ALL_KPI_REPS.forEach(rn => {
            if (rn === myRepName) return;
            const k = `${ym}__${rn}`;
            const snap = teamKpiSnapshotStore[ym]?.[rn];
            const liveClDol = Math.round((classementByYmAndRep.get(k) || 0) * 100) / 100;
            const directDol = Math.round((mObj.reps[rn]['Vente $'].weeks.reduce((a, b) => a + b, 0)) * 100) / 100;
            const directOts = mObj.reps[rn]['Ventes OTS'].weeks.reduce((a, b) => a + b, 0);
            const directSuivi = mObj.reps[rn]['Ventes SUIVI'].weeks.reduce((a, b) => a + b, 0);

            if (snap && typeof snap === 'object') {
              const snapDol = Math.round(Number(snap.dol || 0) * 100) / 100;
              const snapClDol = Math.round(Number(snap.clDol || 0) * 100) / 100;
              const liveDeltaDol = liveClDol > snapClDol + 0.01 ? Math.round((liveClDol - snapClDol) * 100) / 100 : 0;
              const bestDol = Math.max(directDol, Math.round((snapDol + liveDeltaDol) * 100) / 100);
              const bestOts = Math.max(directOts, Number(snap.ots || 0) + (liveDeltaDol > 0 && directOts <= Number(snap.ots || 0) ? 1 : 0));
              const bestSuivi = Math.max(directSuivi, Number(snap.suivi || 0));

              mObj.reps[rn]['Booké'].weeks = [Math.max(mObj.reps[rn]['Booké'].weeks[0] || 0, Number(snap.booke || 0)), 0, 0, 0];
              mObj.reps[rn]['RDV Réalisé'].weeks = [Math.max(mObj.reps[rn]['RDV Réalisé'].weeks[0] || 0, Number(snap.rdv || 0)), 0, 0, 0];
              mObj.reps[rn]['Ventes OTS'].weeks = [bestOts, 0, 0, 0];
              mObj.reps[rn]['Ventes SUIVI'].weeks = [bestSuivi, 0, 0, 0];
              mObj.reps[rn]['Vente $'].weeks = [bestDol, 0, 0, 0];

              if (Number(snap.objDol) > 0) mObj.reps[rn]['Vente $'].objectif = Number(snap.objDol);
              if (Number(snap.objClose) > 0) mObj.reps[rn]['% Close'].objectif = Number(snap.objClose);
              if (Number(snap.objBooke) > 0) mObj.reps[rn]['Booké'].objectif = Number(snap.objBooke);
              if (Number(snap.objRdv) > 0) mObj.reps[rn]['RDV Réalisé'].objectif = Number(snap.objRdv);
            } else if (liveClDol > directDol) {
              mObj.reps[rn]['Vente $'].weeks = [liveClDol, 0, 0, 0];
              mObj.reps[rn]['Ventes OTS'].weeks = [Math.max(directOts, 1), 0, 0, 0];
            }
          });
        });
      }

      // 6. Calculer tous les cumuls, % Close, totaux entreprises et totaux équipe
      Object.values(freshStore).forEach(mObj => {
        const nw = mObj.num_weeks || 4;
        ALL_KPI_REPS.forEach(rn => {
          const rData = mObj.reps[rn];
          ['Booké', 'RDV Réalisé', 'Ventes OTS', 'Ventes SUIVI', 'Vente $'].forEach(k => {
            rData[k].cumul = Math.round(rData[k].weeks.reduce((a, b) => a + b, 0) * 100) / 100;
          });
          for (let wi = 0; wi < nw; wi++) {
            const rdvW = rData['RDV Réalisé'].weeks[wi] || 0;
            const vW = (rData['Ventes OTS'].weeks[wi] || 0) + (rData['Ventes SUIVI'].weeks[wi] || 0);
            rData['% Close'].weeks[wi] = rdvW > 0 ? Math.round((vW / rdvW) * 10000) / 100 : 0;
          }
          const totRdv = rData['RDV Réalisé'].cumul;
          const totV = rData['Ventes OTS'].cumul + rData['Ventes SUIVI'].cumul;
          rData['% Close'].cumul = totRdv > 0 ? Math.round((totV / totRdv) * 10000) / 100 : 0;
        });

        Object.entries(mObj.entreprises).forEach(([eKey, eObj]) => {
          if (isEntrepriseArchived(eKey)) {
            const eFactor = getEntrepriseArchiveMonthFactor(eKey, mObj.ym || '');
            const archEnt = findArchivedEntrepriseEntry(eKey);
            const baseEntObj = Number(archEnt?.monthlyObjective || eObj.total?.objectif || 0);
            if (eObj.total) eObj.total.objectif = Math.round(baseEntObj * eFactor * 100) / 100;
            Object.values(eObj.reps || {}).forEach(rpObj => {
              if (rpObj) rpObj.objectif = Math.round(Number(rpObj.objectif || 0) * eFactor * 100) / 100;
            });
          }
          const sumW = Array(nw).fill(0);
          Object.values(eObj.reps).forEach(rpObj => {
            rpObj.cumul = Math.round(rpObj.weeks.reduce((a, b) => a + b, 0) * 100) / 100;
            for (let wi = 0; wi < nw; wi++) {
              sumW[wi] = Math.round((sumW[wi] + (rpObj.weeks[wi] || 0)) * 100) / 100;
            }
          });
          eObj.total.weeks = sumW;
          eObj.total.cumul = Math.round(sumW.reduce((a, b) => a + b, 0) * 100) / 100;
        });

        for (let wi = 0; wi < nw; wi++) {
          mObj.team['Booké'].weeks[wi] = ALL_KPI_REPS.reduce((s, rn) => s + (mObj.reps[rn]['Booké'].weeks[wi] || 0), 0);
          mObj.team['RDV RÉA'].weeks[wi] = ALL_KPI_REPS.reduce((s, rn) => s + (mObj.reps[rn]['RDV Réalisé'].weeks[wi] || 0), 0);
          mObj.team['VENTE OTS'].weeks[wi] = ALL_KPI_REPS.reduce((s, rn) => s + (mObj.reps[rn]['Ventes OTS'].weeks[wi] || 0), 0);
          mObj.team['VENTE SUIVI'].weeks[wi] = ALL_KPI_REPS.reduce((s, rn) => s + (mObj.reps[rn]['Ventes SUIVI'].weeks[wi] || 0), 0);
          mObj.team['VENTE TOTAL'].weeks[wi] = mObj.team['VENTE OTS'].weeks[wi] + mObj.team['VENTE SUIVI'].weeks[wi];
          mObj.team['VENTE $'].weeks[wi] = Math.round(ALL_KPI_REPS.reduce((s, rn) => s + (mObj.reps[rn]['Vente $'].weeks[wi] || 0), 0) * 100) / 100;
        }
        ['Booké', 'RDV RÉA', 'VENTE OTS', 'VENTE SUIVI', 'VENTE TOTAL', 'VENTE $'].forEach(k => {
          mObj.team[k].cumul = Math.round(mObj.team[k].weeks.reduce((a, b) => a + b, 0) * 100) / 100;
        });
      });

      kpiStore = freshStore;
      applyRepTargetsOverridesToStore();
      applyCompanyTargetsMetaToStore();
      updateKpiSourceBadge();
      populateKpiMonthSelect(true);

      if (currentProfile?.role === 'admin' && sb) {
        scheduleAdminKpiSnapshotSync();
      }
    }

    function scheduleAdminKpiSnapshotSync() {
      clearTimeout(kpiSnapshotSyncTimer);
      kpiSnapshotSyncTimer = setTimeout(async () => {
        if (!sb || currentProfile?.role !== 'admin') return;
        const classementByYmAndRep = new Map();
        (state.classement || []).forEach(row => {
          const ym = row.mois_commission;
          const rn = resolveKpiRepName(row.representant_nom || '');
          if (isValidYearMonth(ym) && rn) {
            const k = `${ym}__${rn}`;
            classementByYmAndRep.set(k, (classementByYmAndRep.get(k) || 0) + Number(row.total_ventes || 0));
          }
        });

        for (const [ym, mObj] of Object.entries(kpiStore)) {
          if (!isValidYearMonth(ym) || !mObj?.reps) continue;
          const snapPayload = {};
          ALL_KPI_REPS.forEach(rn => {
            const r = mObj.reps[rn];
            if (!r) return;
            snapPayload[rn] = {
              booke: Number(r['Booké']?.cumul || 0),
              rdv: Number(r['RDV Réalisé']?.cumul || 0),
              ots: Number(r['Ventes OTS']?.cumul || 0),
              suivi: Number(r['Ventes SUIVI']?.cumul || 0),
              dol: Math.round(Number(r['Vente $']?.cumul || 0) * 100) / 100,
              objDol: Number(r['Vente $']?.objectif || 0),
              objClose: Number(r['% Close']?.objectif || 35),
              objBooke: Number(r['Booké']?.objectif || 80),
              objRdv: Number(r['RDV Réalisé']?.objectif || 60),
              clDol: Math.round((classementByYmAndRep.get(`${ym}__${rn}`) || 0) * 100) / 100
            };
          });
          const serialized = JSON.stringify(snapPayload);
          const existsInDb = (kpiHebdoRows || []).some(r => Number(r.semaine_index) === -11 && String(r.mois_commission) === ym && r.representant_nom === serialized);
          if (existsInDb && lastSyncedKpiSnapshotHashes[ym] === serialized) continue;
          lastSyncedKpiSnapshotHashes[ym] = serialized;
          teamKpiSnapshotStore[ym] = snapPayload;
          await safeUpsertKpiHebdoRow({
            mois_commission: ym,
            semaine_index: -11,
            representant_nom: serialized,
            booke: 0,
            rdv_realise: 0,
            updated_at: new Date().toISOString()
          });
        }
        lsSet('korpo_kpi_team_snapshot_v1', JSON.stringify(teamKpiSnapshotStore));
      }, 600);
    }

    function updateKpiSourceBadge() {
      const badge = document.getElementById('kpi-source-badge');
      if (!badge) return;
      if (hasActiveCondensedBaseline()) {
        const countMonths = Object.keys(condensedBaselineStore).filter(isValidYearMonth).length;
        badge.textContent = `📄 Base condensée (${countMonths} mois) + 💵 Fichier de ventes`;
        badge.className = 'px-2.5 py-0.5 text-[11px] font-bold rounded-full bg-indigo-50 text-indigo-700 border border-indigo-200';
      } else {
        badge.textContent = '💵 100 % basé sur le fichier de ventes';
        badge.className = 'px-2.5 py-0.5 text-[11px] font-bold rounded-full bg-emerald-50 text-emerald-700 border border-emerald-200';
      }
    }

    function syncSalesIntoKpi() {
      rebuildKpiStore();
    }

    function progressColor(pct) {
      if (pct >= 100) return 'bg-emerald-500';
      if (pct >= 65) return 'bg-amber-400';
      return 'bg-rose-500';
    }

    window.setKpiPeriodQuick = (mode) => {
      const modeEl = document.getElementById('kpi-period-mode');
      const selMonth = document.getElementById('kpi-select-month');
      const customBox = document.getElementById('kpi-custom-range-box');
      if (modeEl) modeEl.value = mode;
      if (customBox) {
        const isCustom = mode === 'custom_range';
        customBox.classList.toggle('hidden', !isCustom);
        if (isCustom) {
          const sIn = document.getElementById('kpi-custom-start');
          const eIn = document.getElementById('kpi-custom-end');
          if (sIn && !sIn.value) sIn.value = `${currentMonthStr}-01`;
          if (eIn && !eIn.value) eIn.value = currentDateStr;
        }
      }
      if (selMonth) {
        if (mode === 'this_month' && Array.from(selMonth.options).some(o => o.value === currentMonthStr)) {
          selMonth.value = currentMonthStr;
        } else if (mode === 'this_quarter') {
          const qVal = `QUARTER_${getQuarterForMonth(currentMonthStr)}`;
          if (Array.from(selMonth.options).some(o => o.value === qVal)) selMonth.value = qVal;
        } else if (mode === 'this_year') {
          const yVal = `YEAR_${today.getFullYear()}`;
          if (Array.from(selMonth.options).some(o => o.value === yVal)) selMonth.value = yVal;
        } else if (mode === 'all') {
          if (Array.from(selMonth.options).some(o => o.value === currentMonthStr)) selMonth.value = currentMonthStr;
        }
      }
      renderKpiDashboard();
    };

    function computeCustomRangeMonthFactor(startYmd, endYmd) {
      if (!startYmd || !endYmd) return 1;
      const s = startYmd <= endYmd ? startYmd : endYmd;
      const e = startYmd <= endYmd ? endYmd : startYmd;
      const startDt = new Date(s + 'T00:00:00');
      const endDt = new Date(e + 'T00:00:00');
      if (isNaN(startDt.getTime()) || isNaN(endDt.getTime())) return 1;

      const bizDaysByMonth = new Map();
      const cur = new Date(startDt);
      while (cur <= endDt) {
        const dow = cur.getDay();
        const ym = `${cur.getFullYear()}-${String(cur.getMonth() + 1).padStart(2, '0')}`;
        if (!bizDaysByMonth.has(ym)) bizDaysByMonth.set(ym, 0);
        if (dow !== 0 && dow !== 6) {
          bizDaysByMonth.set(ym, bizDaysByMonth.get(ym) + 1);
        }
        cur.setDate(cur.getDate() + 1);
      }

      if (!bizDaysByMonth.size) return 1;
      let totalFactor = 0;
      bizDaysByMonth.forEach((selectedBizDays, ym) => {
        const totalMonthBizDays = Math.max(1, countBusinessDaysInMonth(ym));
        totalFactor += (selectedBizDays / totalMonthBizDays);
      });
      return totalFactor > 0 ? totalFactor : (1 / Math.max(1, countBusinessDaysInMonth(s.slice(0, 7))));
    }

    function buildSubMonthKpiObject(label, fakePeriodState) {
      const nw = 4;
      const refYm = fakePeriodState?.mode === 'custom_range' && fakePeriodState.startVal
        ? fakePeriodState.startVal.slice(0, 7)
        : currentMonthStr;
      const factor = fakePeriodState?.mode === 'custom_range'
        ? computeCustomRangeMonthFactor(fakePeriodState.startVal, fakePeriodState.endVal)
        : (fakePeriodState?.mode === 'this_week'
          ? (5 / Math.max(1, countBusinessDaysInMonth(currentMonthStr)))
          : 1);

      const scaleObj = (val, isMoney = false) => {
        const raw = Number(val || 0) * factor;
        return isMoney ? Math.round(raw) : Math.max(0, Math.round(raw * 10) / 10);
      };

      const refStoreMonth = kpiStore[refYm] || kpiStore[currentMonthStr];
      const makeRep = (rn) => {
        const refRep = refStoreMonth?.reps?.[rn];
        const salesTarget = getSalesTargetForRep(rn, refYm);
        return {
          'Booké': { weeks: [0, 0, 0, 0], cumul: 0, objectif: scaleObj(refRep?.['Booké']?.objectif || 80, false) },
          'RDV Réalisé': { weeks: [0, 0, 0, 0], cumul: 0, objectif: scaleObj(refRep?.['RDV Réalisé']?.objectif || 60, false) },
          'Ventes OTS': { weeks: [0, 0, 0, 0], cumul: 0, objectif: scaleObj(refRep?.['Ventes OTS']?.objectif || 10, false) },
          'Ventes SUIVI': { weeks: [0, 0, 0, 0], cumul: 0, objectif: scaleObj(refRep?.['Ventes SUIVI']?.objectif || 20, false) },
          'Vente $': { weeks: [0, 0, 0, 0], cumul: 0, objectif: scaleObj(salesTarget, true) },
          '% Close': { weeks: [0, 0, 0, 0], cumul: 0, objectif: refRep?.['% Close']?.objectif || 35 }
        };
      };
      const obj = {
        ym: refYm,
        label,
        num_weeks: nw,
        monthLabels: [label],
        reps: Object.fromEntries(ALL_KPI_REPS.map(rn => [rn, makeRep(rn)])),
        team: {
          'Booké': { weeks: [0, 0, 0, 0], cumul: 0, objectif: 0 },
          'RDV RÉA': { weeks: [0, 0, 0, 0], cumul: 0, objectif: 0 },
          'VENTE OTS': { weeks: [0, 0, 0, 0], cumul: 0, objectif: 0 },
          'VENTE SUIVI': { weeks: [0, 0, 0, 0], cumul: 0, objectif: 0 },
          'VENTE TOTAL': { weeks: [0, 0, 0, 0], cumul: 0, objectif: 0 },
          'VENTE $': { weeks: [0, 0, 0, 0], cumul: 0, objectif: 0 }
        },
        entreprises: {}
      };
      const entNames = new Set(['Toitex', 'MonSpray', 'VerreVDM', 'Groupe Diamant Pro', 'MaClim', ...(state.entreprises || []).map(e => e.nom)]);
      entNames.forEach(en => {
        const refEnt = refStoreMonth?.entreprises?.[en];
        obj.entreprises[en] = {
          reps: Object.fromEntries(ALL_KPI_REPS.map(rn => [rn, {
            weeks: [0, 0, 0, 0],
            cumul: 0,
            objectif: scaleObj(refEnt?.reps?.[rn]?.objectif || 0, true)
          }])),
          total: {
            weeks: [0, 0, 0, 0],
            cumul: 0,
            objectif: scaleObj(refEnt?.total?.objectif || 0, true)
          }
        };
      });

      (state.ventes || []).forEach(v => {
        if (!matchesSalePeriod(v, fakePeriodState, true)) return;
        const repObj = state.representants.find(r => String(r.id) === String(v.representant_id));
        const entObj = state.entreprises.find(e => String(e.id) === String(v.entreprise_id));
        const rn = resolveKpiRepName(repObj?.nom || v.rep_nom || '');
        const en = (entObj?.nom || v.entreprise_nom || '').trim();
        const val = Number(v.valeur || 0);
        const isSui = (v.note || '').toUpperCase().includes('[SUIVI]');
        if (rn && obj.reps[rn]) {
          obj.reps[rn]['Vente $'].cumul = Math.round((obj.reps[rn]['Vente $'].cumul + val) * 100) / 100;
          obj.reps[rn]['Vente $'].weeks[0] = obj.reps[rn]['Vente $'].cumul;
          if (isSui) {
            obj.reps[rn]['Ventes SUIVI'].cumul += 1;
            obj.reps[rn]['Ventes SUIVI'].weeks[0] += 1;
          } else {
            obj.reps[rn]['Ventes OTS'].cumul += 1;
            obj.reps[rn]['Ventes OTS'].weeks[0] += 1;
          }
        }
        if (en) {
          if (!obj.entreprises[en]) {
            obj.entreprises[en] = {
              reps: Object.fromEntries(ALL_KPI_REPS.map(r => [r, { weeks: [0, 0, 0, 0], cumul: 0, objectif: 0 }])),
              total: { weeks: [0, 0, 0, 0], cumul: 0, objectif: 0 }
            };
          }
          if (rn && obj.entreprises[en].reps[rn]) {
            obj.entreprises[en].reps[rn].cumul = Math.round((obj.entreprises[en].reps[rn].cumul + val) * 100) / 100;
            obj.entreprises[en].reps[rn].weeks[0] = obj.entreprises[en].reps[rn].cumul;
          }
          obj.entreprises[en].total.cumul = Math.round((obj.entreprises[en].total.cumul + val) * 100) / 100;
          obj.entreprises[en].total.weeks[0] = obj.entreprises[en].total.cumul;
        }
      });

      // Calculer les RDV Bookés et RDV Réalisés (et le % Close) sur la période sous-mensuelle / sur mesure
      const rangeStart = fakePeriodState?.mode === 'custom_range'
        ? (fakePeriodState.startVal <= fakePeriodState.endVal ? fakePeriodState.startVal : fakePeriodState.endVal)
        : (fakePeriodState?.weekBounds?.start || `${currentMonthStr}-01`);
      const rangeEnd = fakePeriodState?.mode === 'custom_range'
        ? (fakePeriodState.startVal <= fakePeriodState.endVal ? fakePeriodState.endVal : fakePeriodState.startVal)
        : (fakePeriodState?.weekBounds?.end || currentDateStr);

      const bizDaysPerMonth = new Map();
      if (rangeStart && rangeEnd) {
        const curDt = new Date(rangeStart + 'T00:00:00');
        const endDt = new Date(rangeEnd + 'T00:00:00');
        if (!isNaN(curDt.getTime()) && !isNaN(endDt.getTime())) {
          while (curDt <= endDt) {
            const ym = `${curDt.getFullYear()}-${String(curDt.getMonth() + 1).padStart(2, '0')}`;
            if (!bizDaysPerMonth.has(ym)) bizDaysPerMonth.set(ym, 0);
            const dow = curDt.getDay();
            if (dow !== 0 && dow !== 6) bizDaysPerMonth.set(ym, bizDaysPerMonth.get(ym) + 1);
            curDt.setDate(curDt.getDate() + 1);
          }
        }
      }
      if (!bizDaysPerMonth.size) {
        bizDaysPerMonth.set(refYm, Math.max(1, Math.round(countBusinessDaysInMonth(refYm) * factor)));
      }

      ALL_KPI_REPS.forEach(rn => {
        let sumBooke = 0;
        let sumRdv = 0;
        let sumBaseOts = 0;
        let sumBaseSuivi = 0;
        let sumBaseDol = 0;

        // 1) Additionner les RDV journaliers exacts compris entre rangeStart et rangeEnd
        const dailyMonthTotals = new Map();
        Object.entries(dailyRdvLogsStore || {}).forEach(([k, dEntry]) => {
          if (!dEntry || typeof dEntry !== 'object') return;
          const parts = k.split('__');
          const dStr = dEntry.date || parts[0] || '';
          const entryRep = resolveKpiRepName(dEntry.rep || parts[1] || '');
          if (entryRep !== rn || !dStr) return;
          const dYm = dStr.slice(0, 7);
          const prevM = dailyMonthTotals.get(dYm) || { booke: 0, rdv: 0 };
          prevM.booke += Number(dEntry.booke || 0);
          prevM.rdv += Number(dEntry.rdv || 0);
          dailyMonthTotals.set(dYm, prevM);

          if ((!rangeStart || dStr >= rangeStart) && (!rangeEnd || dStr <= rangeEnd)) {
            sumBooke += Number(dEntry.booke || 0);
            sumRdv += Number(dEntry.rdv || 0);
          }
        });

        bizDaysPerMonth.forEach((selBiz, ym) => {
          const totBiz = Math.max(1, countBusinessDaysInMonth(ym));
          const ratio = selBiz / totBiz;
          const mData = kpiStore[ym]?.reps?.[rn];
          const dTot = dailyMonthTotals.get(ym) || { booke: 0, rdv: 0 };
          if (mData) {
            const nonDailyBooke = Math.max(0, Number(mData['Booké']?.cumul || 0) - dTot.booke);
            const nonDailyRdv = Math.max(0, Number(mData['RDV Réalisé']?.cumul || 0) - dTot.rdv);
            sumBooke += nonDailyBooke * ratio;
            sumRdv += nonDailyRdv * ratio;
          }
          const cEntry = condensedBaselineStore?.[ym]?.[rn];
          if (cEntry) {
            sumBaseOts += Number(cEntry.ots || 0) * ratio;
            sumBaseSuivi += Number(cEntry.suivi || 0) * ratio;
            sumBaseDol += Number(cEntry.dollars || 0) * ratio;
          }
        });

        const rObj = obj.reps[rn];
        rObj['Booké'].cumul = Math.round(sumBooke);
        rObj['Booké'].weeks[0] = rObj['Booké'].cumul;
        rObj['RDV Réalisé'].cumul = Math.round(sumRdv);
        rObj['RDV Réalisé'].weeks[0] = rObj['RDV Réalisé'].cumul;

        if (sumBaseOts > 0 || sumBaseSuivi > 0 || sumBaseDol > 0) {
          rObj['Ventes OTS'].cumul += Math.round(sumBaseOts);
          rObj['Ventes OTS'].weeks[0] = rObj['Ventes OTS'].cumul;
          rObj['Ventes SUIVI'].cumul += Math.round(sumBaseSuivi);
          rObj['Ventes SUIVI'].weeks[0] = rObj['Ventes SUIVI'].cumul;
          rObj['Vente $'].cumul = Math.round((rObj['Vente $'].cumul + sumBaseDol) * 100) / 100;
          rObj['Vente $'].weeks[0] = rObj['Vente $'].cumul;
        }

        const totV = rObj['Ventes OTS'].cumul + rObj['Ventes SUIVI'].cumul;
        const totRdv = rObj['RDV Réalisé'].cumul;
        rObj['% Close'].cumul = totRdv > 0 ? Math.round((totV / totRdv) * 10000) / 100 : 0;
        rObj['% Close'].weeks[0] = rObj['% Close'].cumul;
      });

      recalcTeamObjectivesForMonth(obj);
      obj.team['Booké'].cumul = ALL_KPI_REPS.reduce((s, rn) => s + obj.reps[rn]['Booké'].cumul, 0);
      obj.team['Booké'].weeks[0] = obj.team['Booké'].cumul;
      obj.team['RDV RÉA'].cumul = ALL_KPI_REPS.reduce((s, rn) => s + obj.reps[rn]['RDV Réalisé'].cumul, 0);
      obj.team['RDV RÉA'].weeks[0] = obj.team['RDV RÉA'].cumul;
      obj.team['VENTE OTS'].cumul = ALL_KPI_REPS.reduce((s, rn) => s + obj.reps[rn]['Ventes OTS'].cumul, 0);
      obj.team['VENTE SUIVI'].cumul = ALL_KPI_REPS.reduce((s, rn) => s + obj.reps[rn]['Ventes SUIVI'].cumul, 0);
      obj.team['VENTE TOTAL'].cumul = obj.team['VENTE OTS'].cumul + obj.team['VENTE SUIVI'].cumul;
      obj.team['VENTE $'].cumul = Math.round(ALL_KPI_REPS.reduce((s, rn) => s + obj.reps[rn]['Vente $'].cumul, 0) * 100) / 100;
      return obj;
    }

    function getAggregatedKpiForSelection(selVal) {
      const kpiMode = document.getElementById('kpi-period-mode')?.value || '';
      if (kpiMode === 'this_week') {
        const wb = getCurrentWeekBounds(today);
        const lbl = `Cette semaine (${formatDateLabelFr(wb.start)} → ${formatDateLabelFr(wb.end)})`;
        return {
          isYear: false,
          isCustomSubMonth: true,
          data: buildSubMonthKpiObject(lbl, { mode: 'this_week', weekBounds: wb }),
          monthsIncluded: [currentMonthStr]
        };
      }
      if (kpiMode === 'custom_range') {
        const sVal = document.getElementById('kpi-custom-start')?.value || `${currentMonthStr}-01`;
        const eVal = document.getElementById('kpi-custom-end')?.value || currentDateStr;
        const s = sVal <= eVal ? sVal : eVal;
        const e = sVal <= eVal ? eVal : sVal;
        const lbl = `Du ${formatDateLabelFr(s)} au ${formatDateLabelFr(e)}`;
        return {
          isYear: false,
          isCustomSubMonth: true,
          data: buildSubMonthKpiObject(lbl, { mode: 'custom_range', startVal: s, endVal: e }),
          monthsIncluded: [s.slice(0, 7)]
        };
      }

      const isAllSel = kpiMode === 'all' || selVal === 'ALL';
      const isYearSel = selVal && selVal.startsWith('YEAR_');
      const isQuarterSel = selVal && selVal.startsWith('QUARTER_');
      if (!isAllSel && !isYearSel && !isQuarterSel) {
        return { isYear: false, data: kpiStore[selVal], monthsIncluded: [selVal] };
      }
      const yearPrefix = isYearSel ? selVal.replace('YEAR_', '') : (isQuarterSel ? selVal.replace('QUARTER_', '').slice(0, 4) : '');
      const quarterTarget = isQuarterSel ? selVal.replace('QUARTER_', '') : null;
      const matchingMonths = Object.keys(kpiStore)
        .filter(ym => isAllSel ? isValidYearMonth(ym) : (isYearSel ? ym.startsWith(yearPrefix) : getQuarterForMonth(ym) === quarterTarget))
        .sort();

      const allReps = ALL_KPI_REPS;
      const repMetrics = ['Booké', 'RDV Réalisé', 'Ventes OTS', 'Ventes SUIVI', 'Vente $', '% Close'];
      const teamMetrics = ['Booké', 'RDV RÉA', 'VENTE OTS', 'VENTE SUIVI', 'VENTE TOTAL', 'VENTE $'];

      const agg = {
        label: isAllSel
          ? `Cumul global (${matchingMonths.length} mois)`
          : (isQuarterSel
            ? `Trimestre ${quarterTarget.replace('-Q', ' T')} (${matchingMonths.length} mois)`
            : `Année ${yearPrefix} (${matchingMonths.length} mois)`),
        num_weeks: matchingMonths.length,
        monthLabels: matchingMonths.map(ym => yearPrefix ? kpiStore[ym].label.replace(` ${yearPrefix}`, '') : formatYearMonthLabelFr(ym, true)),
        reps: {},
        team: {},
        entreprises: {}
      };

      allReps.forEach(rn => {
        agg.reps[rn] = {};
        repMetrics.forEach(mk => {
          agg.reps[rn][mk] = { weeks: [], cumul: 0, objectif: 0 };
        });
      });
      teamMetrics.forEach(tk => {
        agg.team[tk] = { weeks: [], cumul: 0, objectif: 0 };
      });

      const allEntNames = new Set();
      matchingMonths.forEach(ym => {
        Object.keys(kpiStore[ym].entreprises || {}).forEach(en => allEntNames.add(en));
      });
      allEntNames.forEach(en => {
        agg.entreprises[en] = {
          reps: {},
          total: { weeks: [], cumul: 0, objectif: 0 }
        };
        allReps.forEach(rn => {
          agg.entreprises[en].reps[rn] = { weeks: [], cumul: 0, objectif: 0 };
        });
      });

      matchingMonths.forEach(ym => {
        const mObj = kpiStore[ym];
        allReps.forEach(rn => {
          const rSrc = mObj.reps[rn] || {};
          repMetrics.forEach(mk => {
            const cVal = Number(rSrc[mk]?.cumul || 0);
            const oVal = Number(rSrc[mk]?.objectif || 0);
            agg.reps[rn][mk].weeks.push(cVal);
            if (mk === '% Close') {
              if (oVal > 0) agg.reps[rn][mk].objectif = oVal;
            } else {
              agg.reps[rn][mk].cumul += cVal;
              agg.reps[rn][mk].objectif += oVal;
            }
          });
        });

        teamMetrics.forEach(tk => {
          const cVal = Number(mObj.team[tk]?.cumul || 0);
          const oVal = Number(mObj.team[tk]?.objectif || 0);
          agg.team[tk].weeks.push(cVal);
          agg.team[tk].cumul += cVal;
          agg.team[tk].objectif += oVal;
        });

        allEntNames.forEach(en => {
          const eSrc = mObj.entreprises[en];
          const tCumul = Number(eSrc?.total?.cumul || 0);
          const tObj = Number(eSrc?.total?.objectif || 0);
          agg.entreprises[en].total.weeks.push(tCumul);
          agg.entreprises[en].total.cumul += tCumul;
          agg.entreprises[en].total.objectif += tObj;

          allReps.forEach(rn => {
            const rCumul = Number(eSrc?.reps?.[rn]?.cumul || 0);
            const rObj = Number(eSrc?.reps?.[rn]?.objectif || 0);
            agg.entreprises[en].reps[rn].weeks.push(rCumul);
            agg.entreprises[en].reps[rn].cumul += rCumul;
            agg.entreprises[en].reps[rn].objectif += rObj;
          });
        });
      });

      allReps.forEach(rn => {
        const totRdv = agg.reps[rn]['RDV Réalisé'].cumul;
        const totV = agg.reps[rn]['Ventes OTS'].cumul + agg.reps[rn]['Ventes SUIVI'].cumul;
        agg.reps[rn]['% Close'].cumul = totRdv > 0 ? Math.round((totV / totRdv) * 10000) / 100 : 0;
      });

      return { isYear: true, data: agg, monthsIncluded: matchingMonths };
    }

    function getCompanyAnnualObjective(entName, repFilter, fallbackSumObj) {
      if (repFilter && repFilter !== 'ALL') {
        const repMeta = companyTargetsMetaStore[`${entName}__${repFilter}`];
        if (repMeta && Number(repMeta.annualAmount) > 0) return Number(repMeta.annualAmount);
        if (repMeta && Number(repMeta.monthlyAmount) > 0) return Number(repMeta.monthlyAmount) * 12;
        return fallbackSumObj;
      }
      const allMeta = companyTargetsMetaStore[`${entName}__ALL`];
      if (allMeta && Number(allMeta.annualAmount) > 0) return Number(allMeta.annualAmount);
      if (allMeta && Number(allMeta.monthlyAmount) > 0) return Number(allMeta.monthlyAmount) * 12;

      const repMetas = ALL_KPI_REPS
        .map(rn => companyTargetsMetaStore[`${entName}__${rn}`])
        .filter(Boolean);
      if (repMetas.length > 0) {
        return repMetas.reduce((sum, m) => sum + Number(m.annualAmount || (Number(m.monthlyAmount || 0) * 12)), 0);
      }
      return fallbackSumObj;
    }

    function renderKpiDashboard() {
      rebuildKpiStore();
      const activeKpiMode = document.getElementById('kpi-period-mode')?.value || 'this_month';
      document.querySelectorAll('.kpi-pill-btn').forEach(btn => {
        const isAct = btn.dataset.kpiMode === activeKpiMode;
        btn.className = isAct
          ? 'kpi-pill-btn shrink-0 px-3 py-2 sm:py-1.5 text-xs font-extrabold rounded-xl bg-indigo-600 text-white shadow-2xs transition whitespace-nowrap cursor-pointer'
          : 'kpi-pill-btn shrink-0 px-3 py-2 sm:py-1.5 text-xs font-bold rounded-xl bg-slate-100 text-slate-600 hover:bg-slate-200 transition whitespace-nowrap cursor-pointer';
      });
      const selMonthEl = document.getElementById('kpi-select-month');
      if (!selMonthEl.value) populateKpiMonthSelect(false);
      const selVal = activeKpiMode === 'all' ? 'ALL' : (selMonthEl.value || currentMonthStr);
      if (activeKpiMode !== 'all' && !selVal.startsWith('YEAR_') && !selVal.startsWith('QUARTER_') && !kpiStore[selVal]) {
        ensureKpiMonthStructure(selVal);
        populateKpiMonthSelect(true);
      }
      const repSelEl = document.getElementById('kpi-select-rep');
      const repFilter = repSelEl ? repSelEl.value : 'ALL';
      const { isYear, data: mObj } = getAggregatedKpiForSelection(selVal);
      if (!mObj) return;

      // 1. Cartes synthèse (4 étapes de l'entonnoir de vente)
      let booke = 0, rdv = 0, ots = 0, suivi = 0, ventesDol = 0, objDol = 0, objRdv = 0, objVentesCnt = 0;
      if (repFilter === 'ALL' || !mObj.reps?.[repFilter]) {
        booke = mObj.team?.['Booké']?.cumul || 0;
        rdv = mObj.team?.['RDV RÉA']?.cumul || 0;
        ots = mObj.team?.['VENTE OTS']?.cumul || 0;
        suivi = mObj.team?.['VENTE SUIVI']?.cumul || 0;
        ventesDol = mObj.team?.['VENTE $']?.cumul || 0;
        objDol = mObj.team?.['VENTE $']?.objectif || 0;
        objRdv = mObj.team?.['RDV RÉA']?.objectif || 0;
        objVentesCnt = mObj.team?.['VENTE TOTAL']?.objectif || 76;
      } else {
        const r = mObj.reps[repFilter];
        booke = r['Booké']?.cumul || 0;
        rdv = r['RDV Réalisé']?.cumul || 0;
        ots = r['Ventes OTS']?.cumul || 0;
        suivi = r['Ventes SUIVI']?.cumul || 0;
        ventesDol = r['Vente $']?.cumul || 0;
        objDol = r['Vente $']?.objectif || 0;
        objRdv = r['RDV Réalisé']?.objectif || 0;
        objVentesCnt = (r['Ventes OTS']?.objectif || 0) + (r['Ventes SUIVI']?.objectif || 0);
      }

      const totalVentesCnt = ots + suivi;
      const closeRate = rdv > 0 ? ((totalVentesCnt / rdv) * 100) : 0;
      const targetClose = repFilter === 'ALL' ? 30 : (mObj.reps[repFilter]?.['% Close']?.objectif || 35);
      const showRate = booke > 0 ? Math.round((rdv / booke) * 100) : 0;
      const dolPct = objDol > 0 ? Math.round((ventesDol / objDol) * 100) : 0;

      document.getElementById('kpi-summary-cards').innerHTML = `
        <div class="bg-white p-5 rounded-2xl border border-slate-200/80 shadow-xs space-y-2">
          <div class="flex items-center justify-between">
            <span class="text-xs font-bold text-slate-500">1. Rendez-vous réalisés</span>
            <span class="px-2 py-0.5 text-[11px] font-bold rounded-full bg-indigo-50 text-indigo-700">${showRate}% présence</span>
          </div>
          <div class="flex items-baseline gap-2">
            <span class="text-3xl font-extrabold text-slate-900 tabular-nums">${rdv}</span>
            <span class="text-xs font-semibold text-slate-400">/ ${booke} bookés</span>
          </div>
          <div class="w-full h-2 bg-slate-100 rounded-full overflow-hidden">
            <div class="h-full bg-indigo-600 rounded-full" style="width: ${Math.min(100, objRdv > 0 ? Math.round((rdv / objRdv) * 100) : 0)}%"></div>
          </div>
          <div class="text-[11px] text-slate-400">Objectif RDV : ${objRdv}</div>
        </div>

        <div class="bg-white p-5 rounded-2xl border border-slate-200/80 shadow-xs space-y-2">
          <div class="flex items-center justify-between">
            <span class="text-xs font-bold text-slate-500">2. Contrats signés</span>
            <span class="px-2 py-0.5 text-[11px] font-bold rounded-full bg-sky-50 text-sky-700">${ots} OTS • ${suivi} Suivi</span>
          </div>
          <div class="flex items-baseline gap-2">
            <span class="text-3xl font-extrabold text-slate-900 tabular-nums">${totalVentesCnt}</span>
            <span class="text-xs font-semibold text-slate-400">/ ${objVentesCnt} objectif</span>
          </div>
          <div class="w-full h-2 bg-slate-100 rounded-full overflow-hidden">
            <div class="h-full bg-sky-500 rounded-full" style="width: ${Math.min(100, objVentesCnt > 0 ? Math.round((totalVentesCnt / objVentesCnt) * 100) : 0)}%"></div>
          </div>
          <div class="text-[11px] text-slate-400">One-Time-Sit vs Suivis</div>
        </div>

        <div class="bg-white p-5 rounded-2xl border border-slate-200/80 shadow-xs space-y-2">
          <div class="flex items-center justify-between">
            <span class="text-xs font-bold text-slate-500">3. Taux de closing</span>
            ${statusBadge(closeRate, targetClose, true)}
          </div>
          <div class="flex items-baseline gap-2">
            <span class="text-3xl font-extrabold ${closeRate >= targetClose ? 'text-emerald-600' : 'text-slate-900'} tabular-nums">${closeRate.toFixed(1)} %</span>
            <span class="text-xs font-semibold text-slate-400">cible ${targetClose}%</span>
          </div>
          <div class="w-full h-2 bg-slate-100 rounded-full overflow-hidden">
            <div class="h-full ${closeRate >= targetClose ? 'bg-emerald-500' : 'bg-amber-400'} rounded-full" style="width: ${Math.min(100, Math.round((closeRate / Math.max(targetClose, 1)) * 100))}%"></div>
          </div>
          <div class="text-[11px] text-slate-400">Ventes ÷ RDV réalisés</div>
        </div>

        <div class="bg-gradient-to-br from-slate-900 to-indigo-950 p-5 rounded-2xl text-white shadow-md space-y-2">
          <div class="flex items-center justify-between">
            <span class="text-xs font-bold text-indigo-200">4. Volume de ventes</span>
            <span class="px-2 py-0.5 text-[11px] font-extrabold rounded-full bg-emerald-500/20 text-emerald-300">${dolPct} %</span>
          </div>
          <div class="text-2xl sm:text-3xl font-extrabold text-emerald-400 tabular-nums">${fmtMoney0(ventesDol)}</div>
          <div class="w-full h-2 bg-white/10 rounded-full overflow-hidden">
            <div class="h-full bg-emerald-400 rounded-full" style="width: ${Math.min(100, dolPct)}%"></div>
          </div>
          <div class="text-[11px] text-slate-300">Objectif : ${fmtMoney0(objDol)}</div>
        </div>
      `;

      // 2. Tableau comparatif des représentants : inclure les actifs + les anciens qui ont des stats/objectifs sur la période affichée
      const allReps = ALL_KPI_REPS.filter(rn => {
        if (!isRepArchived(rn)) return true;
        const r = mObj.reps?.[rn];
        if (!r) return false;
        const hasActivity = (r['Booké']?.cumul || 0) > 0
          || (r['RDV Réalisé']?.cumul || 0) > 0
          || (r['Ventes OTS']?.cumul || 0) > 0
          || (r['Ventes SUIVI']?.cumul || 0) > 0
          || (r['Vente $']?.cumul || 0) > 0
          || (r['Vente $']?.objectif || 0) > 0;
        return hasActivity;
      });
      const compareRows = allReps.map(rn => {
        const r = mObj.reps[rn];
        if (!r) return '';
        const b = r['Booké']?.cumul || 0;
        const rea = r['RDV Réalisé']?.cumul || 0;
        const vOts = r['Ventes OTS']?.cumul || 0;
        const vSui = r['Ventes SUIVI']?.cumul || 0;
        const vTot = vOts + vSui;
        const cl = rea > 0 ? ((vTot / rea) * 100) : 0;
        const clObj = r['% Close']?.objectif || 35;
        const dol = r['Vente $']?.cumul || 0;
        const dolObj = r['Vente $']?.objectif || 0;
        const pctObj = dolObj > 0 ? Math.round((dol / dolObj) * 100) : 0;
        const isSelected = repFilter === rn;

        return `
          <tr onclick="selectKpiRep('${esc(rn)}')" class="cursor-pointer transition ${isSelected ? 'bg-indigo-50/70' : 'hover:bg-slate-50'}">
            <td class="py-3.5 px-5">
              <div class="flex items-center gap-3">
                <span class="w-8 h-8 rounded-xl ${isSelected ? 'bg-indigo-600 text-white' : 'bg-slate-100 text-slate-700'} text-xs font-extrabold flex items-center justify-center">${esc(rn.slice(0, 2).toUpperCase())}</span>
                <div>
                  <div class="font-bold text-slate-900 flex items-center gap-1.5"><span>${esc(rn)}</span>${isRepArchived(rn) ? `<span class="px-1.5 py-0.2 text-[9px] font-bold rounded bg-slate-100 text-slate-500">Archivé</span>` : ''}</div>
                  <div class="text-[11px] text-slate-400">${isYear ? 'Voir mois par mois →' : 'Voir le détail du mois →'}</div>
                </div>
              </div>
            </td>
            <td class="py-3.5 px-4 text-center tabular-nums">
              <span class="font-bold text-slate-900">${rea}</span>
              <span class="text-xs text-slate-400"> / ${b}</span>
            </td>
            <td class="py-3.5 px-4 text-center tabular-nums">
              <span class="font-bold text-slate-900">${vTot}</span>
              <span class="text-[11px] text-slate-400 block">${vOts} OTS • ${vSui} Suivi</span>
            </td>
            <td class="py-3.5 px-4 text-center tabular-nums">
              <span class="font-extrabold ${cl >= clObj ? 'text-emerald-600' : 'text-slate-800'}">${cl.toFixed(1)} %</span>
              <span class="text-[11px] text-slate-400 block">cible ${clObj}%</span>
            </td>
            <td class="py-3.5 px-4 text-right tabular-nums">
              <div class="font-extrabold text-slate-900">${fmtMoney0(dol)}</div>
              <div class="text-[11px] text-slate-400">Obj : ${dolObj > 0 ? fmtMoney0(dolObj) : '—'}</div>
            </td>
            <td class="py-3.5 px-5 w-44">
              <div class="flex items-center justify-between text-[11px] font-bold mb-1">
                <span>${dolObj > 0 ? pctObj + ' %' : '—'}</span>
                ${statusBadge(dol, dolObj)}
              </div>
              <div class="w-full h-2 bg-slate-100 rounded-full overflow-hidden">
                <div class="h-full ${progressColor(pctObj)} rounded-full" style="width: ${Math.min(100, pctObj)}%"></div>
              </div>
            </td>
          </tr>
        `;
      }).join('');

      document.getElementById('table-kpi-team-compare').innerHTML = `
        <thead>
          <tr class="bg-slate-50 border-b border-slate-200 text-[11px] font-bold uppercase text-slate-500">
            <th class="py-3 px-5">🧑‍💼 Représentant</th>
            <th class="py-3 px-4 text-center">📅 RDV Réa / Bookés</th>
            <th class="py-3 px-4 text-center">✍️ Ventes fermées</th>
            <th class="py-3 px-4 text-center">📈 % Close</th>
            <th class="py-3 px-4 text-right">💵 Volume généré</th>
            <th class="py-3 px-5">🎯 Progression objectif</th>
          </tr>
        </thead>
        <tbody class="divide-y divide-slate-100">${compareRows}</tbody>
      `;

      const mobileKpiCardsEl = document.getElementById('mobile-kpi-team-cards');
      if (mobileKpiCardsEl) {
        mobileKpiCardsEl.innerHTML = allReps.map(rn => {
          const r = mObj.reps[rn];
          if (!r) return '';
          const b = r['Booké']?.cumul || 0;
          const rea = r['RDV Réalisé']?.cumul || 0;
          const vOts = r['Ventes OTS']?.cumul || 0;
          const vSui = r['Ventes SUIVI']?.cumul || 0;
          const vTot = vOts + vSui;
          const cl = rea > 0 ? ((vTot / rea) * 100) : 0;
          const clObj = r['% Close']?.objectif || 35;
          const dol = r['Vente $']?.cumul || 0;
          const dolObj = r['Vente $']?.objectif || 0;
          const pctObj = dolObj > 0 ? Math.round((dol / dolObj) * 100) : 0;
          const isSelected = repFilter === rn;
          return `
            <div onclick="selectKpiRep('${esc(rn)}')" class="p-4 space-y-2.5 cursor-pointer transition ${isSelected ? 'bg-indigo-50/70' : 'bg-white active:bg-slate-50'}">
              <div class="flex items-center justify-between gap-2">
                <div class="flex items-center gap-2.5">
                  <span class="w-8 h-8 rounded-xl ${isSelected ? 'bg-indigo-600 text-white' : 'bg-slate-100 text-slate-700'} text-xs font-extrabold flex items-center justify-center">${esc(rn.slice(0, 2).toUpperCase())}</span>
                  <div>
                    <div class="font-extrabold text-sm text-slate-900 flex items-center gap-1.5"><span>${esc(rn)}</span>${isRepArchived(rn) ? `<span class="px-1.5 py-0.2 text-[9px] font-bold rounded bg-slate-100 text-slate-500">Archivé</span>` : ''}</div>
                    <div class="text-[11px] text-indigo-600 font-semibold">${isSelected ? 'Fermer le détail ↑' : 'Toucher pour le détail ↓'}</div>
                  </div>
                </div>
                <div class="text-right tabular-nums">
                  <div class="font-extrabold text-sm text-slate-900">${fmtMoney0(dol)}</div>
                  <div class="text-[10px] text-slate-400">Obj : ${dolObj > 0 ? fmtMoney0(dolObj) : '—'}</div>
                </div>
              </div>
              <div class="grid grid-cols-3 gap-2 bg-slate-50 rounded-xl p-2 text-center text-[11px] tabular-nums">
                <div>
                  <span class="text-slate-400 block text-[10px]">RDV Réa</span>
                  <strong class="text-slate-800">${rea} / ${b}</strong>
                </div>
                <div>
                  <span class="text-slate-400 block text-[10px]">Ventes</span>
                  <strong class="text-slate-800">${vTot}</strong> <span class="text-[10px] text-slate-500">(${vOts}/${vSui})</span>
                </div>
                <div>
                  <span class="text-slate-400 block text-[10px]">% Close</span>
                  <strong class="${cl >= clObj ? 'text-emerald-600' : 'text-slate-800'}">${cl.toFixed(1)}%</strong>
                </div>
              </div>
              <div class="space-y-1">
                <div class="flex items-center justify-between text-[11px] font-bold">
                  <span>Progression : ${dolObj > 0 ? pctObj + ' %' : '—'}</span>
                  ${statusBadge(dol, dolObj)}
                </div>
                <div class="w-full h-2 bg-slate-100 rounded-full overflow-hidden">
                  <div class="h-full ${progressColor(pctObj)} rounded-full" style="width: ${Math.min(100, pctObj)}%"></div>
                </div>
              </div>
            </div>
          `;
        }).join('');
      }

      // 3. Détail ciblé d'un représentant (Mois par mois sur vue annuelle, OU décomposition Base condensée + Ventes enregistrées sur vue mensuelle)
      const weeklyContainer = document.getElementById('kpi-rep-weekly-detail');
      if (repFilter === 'ALL') {
        weeklyContainer.innerHTML = '';
      } else if (isYear) {
        const rData = mObj.reps[repFilter];
        const colLabels = mObj.monthLabels || [];
        const wHeaders = colLabels.map(lbl => `<th class="py-2.5 px-3 text-right whitespace-nowrap">${esc(lbl)}</th>`).join('');
        const metricsOrder = ['Booké', 'RDV Réalisé', 'Ventes OTS', 'Ventes SUIVI', 'Vente $', '% Close'];
        const rowsHtml = metricsOrder.map(mKey => {
          const item = rData[mKey];
          const isMoney = mKey === 'Vente $';
          const isPct = mKey === '% Close';
          const fmtCell = (val) => isMoney ? fmtMoney0(val) : (isPct ? `${Number(val).toFixed(1)} %` : val);
          return `
            <tr class="${isMoney ? 'bg-indigo-50/40 font-bold' : 'hover:bg-slate-50'}">
              <td class="py-2.5 px-5 font-semibold text-slate-800 whitespace-nowrap">${mKey}</td>
              ${item.weeks.map(w => `<td class="py-2.5 px-3 text-right tabular-nums text-slate-700">${fmtCell(w)}</td>`).join('')}
              <td class="py-2.5 px-4 text-right tabular-nums font-extrabold text-slate-900 bg-indigo-50/30">${fmtCell(item.cumul)}</td>
              <td class="py-2.5 px-5 text-right tabular-nums text-slate-500">${item.objectif > 0 ? fmtCell(item.objectif) : '—'}</td>
            </tr>
          `;
        }).join('');

        const mobileYearCardsHtml = colLabels.map((lbl, idx) => {
          const bW = rData['Booké'].weeks[idx] || 0;
          const rW = rData['RDV Réalisé'].weeks[idx] || 0;
          const oW = rData['Ventes OTS'].weeks[idx] || 0;
          const sW = rData['Ventes SUIVI'].weeks[idx] || 0;
          const dW = rData['Vente $'].weeks[idx] || 0;
          const cW = rData['% Close'].weeks[idx] || 0;
          return `
            <div class="p-3.5 bg-white space-y-2">
              <div class="flex items-center justify-between">
                <span class="text-xs font-extrabold text-slate-900">${esc(lbl)}</span>
                <span class="text-sm font-extrabold text-indigo-700 tabular-nums">${fmtMoney0(dW)}</span>
              </div>
              <div class="grid grid-cols-3 gap-1.5 bg-slate-50 rounded-xl p-2 text-center text-[11px] tabular-nums">
                <div><span class="text-slate-400 block text-[10px]">RDV Réa</span><strong>${rW} / ${bW}</strong></div>
                <div><span class="text-slate-400 block text-[10px]">Ventes</span><strong>${oW + sW}</strong> <span class="text-[10px] text-slate-500">(${oW}/${sW})</span></div>
                <div><span class="text-slate-400 block text-[10px]">% Close</span><strong>${Number(cW).toFixed(1)}%</strong></div>
              </div>
            </div>
          `;
        }).join('');

        weeklyContainer.innerHTML = `
          <div class="bg-white rounded-2xl border-2 border-indigo-200 shadow-xs overflow-hidden">
            <div class="px-4 sm:px-5 py-3.5 bg-indigo-950 text-white flex items-center justify-between gap-2">
              <span class="text-xs sm:text-sm font-extrabold truncate">Évolution mois par mois — ${esc(repFilter)}</span>
              <button onclick="selectKpiRep('ALL')" class="shrink-0 px-2.5 sm:px-3 py-1 text-[11px] sm:text-xs font-bold bg-white/10 hover:bg-white/20 rounded-lg transition">← Équipe</button>
            </div>
            <div class="md:hidden divide-y divide-slate-100">${mobileYearCardsHtml}</div>
            <div class="hidden md:block overflow-x-auto">
              <table class="w-full text-left border-collapse text-xs">
                <thead>
                  <tr class="bg-slate-50 border-b border-slate-200 text-[11px] font-bold uppercase text-slate-500">
                    <th class="py-2.5 px-5">Indicateur</th>
                    ${wHeaders}
                    <th class="py-2.5 px-4 text-right text-indigo-700">Cumul Année</th>
                    <th class="py-2.5 px-5 text-right">Objectif</th>
                  </tr>
                </thead>
                <tbody class="divide-y divide-slate-100">${rowsHtml}</tbody>
              </table>
            </div>
          </div>
        `;
      } else {
        const baseEntry = condensedBaselineStore[selVal]?.[repFilter] || { booke: 0, rdv: 0, ots: 0, suivi: 0, dollars: 0 };
        const repMonthSales = (state.ventes || []).filter(v => {
          if (getSaleStatsMonth(v) !== selVal) return false;
          const rObj = state.representants.find(r => String(r.id) === String(v.representant_id));
          return resolveKpiRepName(rObj?.nom || v.rep_nom || '') === repFilter;
        });
        const appOts = repMonthSales.filter(v => !(v.note || '').toUpperCase().includes('[SUIVI]')).length;
        const appSuivi = repMonthSales.filter(v => (v.note || '').toUpperCase().includes('[SUIVI]')).length;
        const appDol = repMonthSales.reduce((s, v) => s + Number(v.valeur || 0), 0);
        const rData = mObj.reps[repFilter];

        const hideOtherRepClientNames = currentProfile?.role !== 'admin' && resolveKpiRepName(currentProfile?.nom || '') !== repFilter;
        const salesListHtml = repMonthSales.length
          ? repMonthSales.map(v => {
              const rawCli = v.client_nom || state.clients.find(c => String(c.id) === String(v.client_id))?.nom || '—';
              const cli = hideOtherRepClientNames ? 'Dossier confidentiel' : rawCli;
              const ent = v.entreprise_nom || state.entreprises.find(e => String(e.id) === String(v.entreprise_id))?.nom || '—';
              const isSui = (v.note || '').toUpperCase().includes('[SUIVI]');
              return `
                <tr class="hover:bg-slate-50">
                  <td class="py-2 px-4 text-slate-600 tabular-nums">${esc(v.date_vente || selVal)}</td>
                  <td class="py-2 px-4">${entBadge(ent)}</td>
                  <td class="py-2 px-4 font-bold text-slate-900">${esc(cli)}</td>
                  <td class="py-2 px-4">
                    <span class="px-2 py-0.5 text-[10px] font-extrabold rounded-full ${isSui ? 'bg-purple-50 text-purple-700 border border-purple-200' : 'bg-sky-50 text-sky-700 border border-sky-200'}">${isSui ? 'SUIVI' : 'OTS'}</span>
                  </td>
                  <td class="py-2 px-4 text-right font-extrabold text-slate-900 tabular-nums">${fmtMoney(v.valeur)}</td>
                </tr>
              `;
            }).join('')
          : `<tr><td colspan="5" class="py-4 px-4 text-center text-xs text-slate-400">Aucune vente individuelle enregistrée dans l'application pour ${esc(repFilter)} ce mois-ci.</td></tr>`;

        const mobileSalesCardsHtml = repMonthSales.length
          ? repMonthSales.map(v => {
              const rawCli = v.client_nom || state.clients.find(c => String(c.id) === String(v.client_id))?.nom || '—';
              const cli = hideOtherRepClientNames ? 'Dossier confidentiel' : rawCli;
              const ent = v.entreprise_nom || state.entreprises.find(e => String(e.id) === String(v.entreprise_id))?.nom || '—';
              const isSui = (v.note || '').toUpperCase().includes('[SUIVI]');
              return `
                <div class="p-3.5 bg-white flex items-center justify-between gap-2.5">
                  <div class="min-w-0 space-y-1">
                    <div class="font-bold text-xs text-slate-900 truncate">${esc(cli)}</div>
                    <div class="flex items-center flex-wrap gap-1.5">
                      ${entBadge(ent)}
                      <span class="px-1.5 py-0.5 text-[10px] font-extrabold rounded-md border ${isSui ? 'bg-purple-50 text-purple-700 border-purple-200' : 'bg-sky-50 text-sky-700 border-sky-200'}">${isSui ? 'SUIVI' : 'OTS'}</span>
                      <span class="text-[11px] text-slate-400 tabular-nums">${esc(v.date_vente || selVal)}</span>
                    </div>
                  </div>
                  <div class="text-right shrink-0 font-extrabold text-xs text-slate-900 tabular-nums">
                    ${fmtMoney(v.valeur)}
                  </div>
                </div>
              `;
            }).join('')
          : `<div class="p-4 text-center text-xs text-slate-400">Aucune vente individuelle enregistrée dans l'application pour ${esc(repFilter)} ce mois-ci.</div>`;

        weeklyContainer.innerHTML = `
          <div class="bg-white rounded-2xl border-2 border-indigo-200 shadow-xs overflow-hidden">
            <div class="px-4 sm:px-5 py-3.5 bg-indigo-950 text-white flex items-center justify-between gap-2">
              <span class="text-xs sm:text-sm font-extrabold truncate">Détail (${esc(mObj.label)}) — ${esc(repFilter)}</span>
              <button onclick="selectKpiRep('ALL')" class="shrink-0 px-2.5 sm:px-3 py-1 text-[11px] sm:text-xs font-bold bg-white/10 hover:bg-white/20 rounded-lg transition">← Équipe</button>
            </div>

            <!-- VUE MOBILE COMPACTE SANS DÉBORDEMENT (< 768px) -->
            <div class="md:hidden divide-y divide-slate-100">
              <div class="p-3.5 space-y-2.5 bg-slate-50/60">
                <div class="bg-white rounded-xl border border-slate-200/80 p-3 space-y-1.5">
                  <div class="flex items-center justify-between text-xs">
                    <span class="font-semibold text-slate-600">📄 Base fichier condensé</span>
                    <strong class="text-slate-800 tabular-nums">${fmtMoney0(baseEntry.dollars || 0)}</strong>
                  </div>
                  <div class="grid grid-cols-4 gap-1 text-center text-[10px] bg-slate-50 rounded-lg p-1.5 tabular-nums">
                    <div><span class="text-slate-400 block">Bookés</span><strong>${Number(baseEntry.booke || 0)}</strong></div>
                    <div><span class="text-slate-400 block">RDV Réa</span><strong>${Number(baseEntry.rdv || 0)}</strong></div>
                    <div><span class="text-sky-700 block">OTS</span><strong class="text-sky-700">${Number(baseEntry.ots || 0)}</strong></div>
                    <div><span class="text-purple-700 block">SUIVI</span><strong class="text-purple-700">${Number(baseEntry.suivi || 0)}</strong></div>
                  </div>
                </div>

                <div class="bg-white rounded-xl border border-emerald-200/80 p-3 space-y-1.5">
                  <div class="flex items-center justify-between text-xs">
                    <span class="font-semibold text-emerald-700">💵 Ventes App (${repMonthSales.length})</span>
                    <strong class="text-emerald-700 tabular-nums">+ ${fmtMoney0(appDol)}</strong>
                  </div>
                  <div class="grid grid-cols-2 gap-1.5 text-center text-[10px] bg-emerald-50/50 rounded-lg p-1.5 tabular-nums">
                    <div><span class="text-sky-700 block">Ventes OTS</span><strong class="text-sky-700">+ ${appOts}</strong></div>
                    <div><span class="text-purple-700 block">Ventes SUIVI</span><strong class="text-purple-700">+ ${appSuivi}</strong></div>
                  </div>
                </div>

                <div class="bg-indigo-50/80 rounded-xl border border-indigo-200 p-3 space-y-1.5">
                  <div class="flex items-center justify-between text-xs">
                    <span class="font-extrabold text-indigo-950">Total consolidé</span>
                    <strong class="text-sm font-extrabold text-indigo-900 tabular-nums">${fmtMoney0(rData['Vente $'].cumul)}</strong>
                  </div>
                  <div class="grid grid-cols-4 gap-1 text-center text-[10px] bg-white/80 rounded-lg p-1.5 tabular-nums">
                    <div><span class="text-slate-500 block">Bookés</span><strong>${rData['Booké'].cumul}</strong></div>
                    <div><span class="text-slate-500 block">RDV Réa</span><strong>${rData['RDV Réalisé'].cumul}</strong></div>
                    <div><span class="text-sky-800 block">OTS</span><strong class="text-sky-800">${rData['Ventes OTS'].cumul}</strong></div>
                    <div><span class="text-purple-800 block">SUIVI</span><strong class="text-purple-800">${rData['Ventes SUIVI'].cumul}</strong></div>
                  </div>
                </div>
              </div>

              <div class="px-3.5 py-2 bg-slate-100/80 text-[10px] font-extrabold uppercase tracking-wider text-slate-500">
                Ventes individuelles du mois (${repMonthSales.length})
              </div>
              <div class="divide-y divide-slate-100">
                ${mobileSalesCardsHtml}
              </div>
            </div>

            <!-- VUE TABLEAU DESKTOP (>= 768px) -->
            <div class="hidden md:block">
              <div class="overflow-x-auto">
                <table class="w-full text-left border-collapse text-xs">
                  <thead>
                    <tr class="bg-slate-50 border-b border-slate-200 text-[11px] font-bold uppercase text-slate-500">
                      <th class="py-2.5 px-4">Source</th>
                      <th class="py-2.5 px-3 text-center">RDV Bookés</th>
                      <th class="py-2.5 px-3 text-center">RDV Réalisés</th>
                      <th class="py-2.5 px-3 text-center text-sky-700">Ventes OTS</th>
                      <th class="py-2.5 px-3 text-center text-purple-700">Ventes SUIVI</th>
                      <th class="py-2.5 px-4 text-right">Volume ($)</th>
                    </tr>
                  </thead>
                  <tbody class="divide-y divide-slate-100 tabular-nums">
                    <tr>
                      <td class="py-2.5 px-4 font-semibold text-slate-600">📄 Base fichier condensé</td>
                      <td class="py-2.5 px-3 text-center">${Number(baseEntry.booke || 0)}</td>
                      <td class="py-2.5 px-3 text-center">${Number(baseEntry.rdv || 0)}</td>
                      <td class="py-2.5 px-3 text-center font-bold text-sky-700">${Number(baseEntry.ots || 0)}</td>
                      <td class="py-2.5 px-3 text-center font-bold text-purple-700">${Number(baseEntry.suivi || 0)}</td>
                      <td class="py-2.5 px-4 text-right font-bold text-slate-700">${fmtMoney0(baseEntry.dollars || 0)}</td>
                    </tr>
                    <tr>
                      <td class="py-2.5 px-4 font-semibold text-emerald-700">💵 Fichier de ventes (${repMonthSales.length} vente(s))</td>
                      <td class="py-2.5 px-3 text-center text-slate-400">—</td>
                      <td class="py-2.5 px-3 text-center text-slate-400">—</td>
                      <td class="py-2.5 px-3 text-center font-bold text-sky-700">+ ${appOts}</td>
                      <td class="py-2.5 px-3 text-center font-bold text-purple-700">+ ${appSuivi}</td>
                      <td class="py-2.5 px-4 text-right font-bold text-emerald-700">+ ${fmtMoney0(appDol)}</td>
                    </tr>
                    <tr class="bg-indigo-50/50 font-extrabold text-slate-900">
                      <td class="py-2.5 px-4">Total consolidé</td>
                      <td class="py-2.5 px-3 text-center">${rData['Booké'].cumul}</td>
                      <td class="py-2.5 px-3 text-center">${rData['RDV Réalisé'].cumul}</td>
                      <td class="py-2.5 px-3 text-center text-sky-800">${rData['Ventes OTS'].cumul}</td>
                      <td class="py-2.5 px-3 text-center text-purple-800">${rData['Ventes SUIVI'].cumul}</td>
                      <td class="py-2.5 px-4 text-right text-indigo-900">${fmtMoney0(rData['Vente $'].cumul)}</td>
                    </tr>
                  </tbody>
                </table>
              </div>
              <div class="border-t border-slate-200 overflow-x-auto">
                <table class="w-full text-left border-collapse text-xs">
                  <thead>
                    <tr class="bg-slate-50/70 border-b border-slate-200 text-[10px] font-bold uppercase text-slate-400">
                      <th class="py-2 px-4">Date</th>
                      <th class="py-2 px-4">Compagnie</th>
                      <th class="py-2 px-4">Client</th>
                      <th class="py-2 px-4">Type KPI</th>
                      <th class="py-2 px-4 text-right">Montant</th>
                    </tr>
                  </thead>
                  <tbody class="divide-y divide-slate-100">${salesListHtml}</tbody>
                </table>
              </div>
            </div>
          </div>
        `;
      }

      // 4. Cartes visuelles par entreprise partenaire
      const entSubEl = document.getElementById('kpi-entreprises-subtitle');
      if (entSubEl) {
        entSubEl.textContent = isYear
          ? "Progression du volume de ventes ($) par rapport aux objectifs annuels de chaque entreprise"
          : "Progression du volume de ventes ($) par rapport aux objectifs mensuels de chaque entreprise";
      }

      const isAdminRole = currentProfile?.role === 'admin';
      const entList = Object.keys(mObj.entreprises).filter(en => {
        if (!isEntrepriseArchived(en)) return true;
        const eObj = mObj.entreprises[en];
        return (eObj?.total?.cumul || 0) > 0 || (eObj?.total?.objectif || 0) > 0;
      });
      document.getElementById('kpi-entreprises-cards').innerHTML = entList.map(entName => {
        const eObj = mObj.entreprises[entName];
        const totVal = repFilter === 'ALL' ? (eObj.total?.cumul || 0) : (eObj.reps[repFilter]?.cumul || 0);
        const rawSumObj = repFilter === 'ALL' ? (eObj.total?.objectif || 0) : (eObj.reps[repFilter]?.objectif || 0);
        const totObj = isYear
          ? getCompanyAnnualObjective(entName, repFilter, rawSumObj)
          : rawSumObj;
        const annualRef = isYear ? totObj : getCompanyAnnualObjective(entName, repFilter, rawSumObj * 12);
        const monthlyRef = isYear ? Math.round(totObj / 12) : totObj;
        const pct = totObj > 0 ? Math.round((totVal / totObj) * 100) : 0;
        const objScopeTag = isYear ? 'Obj. annuel' : 'Obj. mensuel';

        const activeReps = Object.entries(eObj.reps)
          .filter(([rn, rv]) => (repFilter === 'ALL' || rn === repFilter) && (rv.cumul > 0 || rv.objectif > 0))
          .map(([rn, rv]) => {
            const repObjVal = isYear
              ? getCompanyAnnualObjective(entName, rn, rv.objectif)
              : rv.objectif;
            return `
              <div class="flex items-center justify-between text-xs py-1 border-t border-slate-100">
                <span class="font-medium text-slate-600">${esc(rn)}</span>
                <span class="font-bold text-slate-900 tabular-nums">${fmtMoney0(rv.cumul)}${repObjVal > 0 ? ` <span class="text-slate-400 font-normal">/ ${fmtMoney0(repObjVal)}</span>` : ''}</span>
              </div>
            `;
          }).join('');

        return `
          <div class="rounded-2xl border border-slate-200/90 p-4 bg-slate-50/40 flex flex-col justify-between space-y-3">
            <div class="space-y-2">
              <div class="flex items-center justify-between gap-2">
                ${entBadge(entName)}
                <div class="flex items-center gap-1.5">
                  <span class="px-2 py-0.5 text-[10px] font-bold rounded-full ${isYear ? 'bg-indigo-50 text-indigo-700 border border-indigo-200' : 'bg-slate-100 text-slate-600'}">${objScopeTag}</span>
                  <span class="text-xs font-extrabold ${pct >= 100 ? 'text-emerald-600' : 'text-slate-600'} tabular-nums">${totObj > 0 ? pct + ' %' : '—'}</span>
                  ${isAdminRole ? `<button type="button" onclick="openClientTargetsPanel('${esc(entName)}')" title="Configurer l'objectif mensuel ou annuel de ${esc(entName)}" class="px-1.5 py-0.5 text-[11px] font-bold text-slate-500 hover:text-indigo-700 hover:bg-indigo-50 rounded-md transition">✎</button>` : ''}
                </div>
              </div>
              <div class="flex items-baseline justify-between pt-1">
                <span class="text-xl font-extrabold text-slate-900 tabular-nums">${fmtMoney0(totVal)}</span>
                <span class="text-[11px] text-slate-600 font-bold tabular-nums">${objScopeTag} : ${totObj > 0 ? fmtMoney0(totObj) : 'Aucun'}</span>
              </div>
              ${totObj > 0 ? `
                <div class="text-[10px] text-slate-400 flex justify-between tabular-nums">
                  <span>Mensuel : ${fmtMoney0(monthlyRef)}</span>
                  <span>Annuel (12 mois) : ${fmtMoney0(annualRef)}</span>
                </div>
              ` : ''}
              <div class="w-full h-2 bg-slate-200/70 rounded-full overflow-hidden">
                <div class="h-full ${progressColor(pct)} rounded-full" style="width: ${Math.min(100, pct)}%"></div>
              </div>
            </div>
            <div class="pt-1">
              ${activeReps || `<div class="text-[11px] text-slate-400 italic">${isYear ? 'Aucune vente cette année' : 'Aucune vente ce mois-ci'}</div>`}
            </div>
          </div>
        `;
      }).join('');
      renderClientObjectivesGrid(selVal, repFilter);
    }

    window.selectKpiRep = (repName) => {
      const sel = document.getElementById('kpi-select-rep');
      sel.value = (sel.value === repName && repName !== 'ALL') ? 'ALL' : repName;
      renderKpiDashboard();
    };

    function exportKpiCsv() {
      if (currentProfile?.role !== 'admin') return;
      const selVal = document.getElementById('kpi-select-month').value;
      const { isYear, data: mObj } = getAggregatedKpiForSelection(selVal);
      if (!mObj) return;
      const rows = [['Periode', 'Representant', 'RDV_Bookes', 'RDV_Realises', 'Ventes_OTS', 'Ventes_SUIVI', 'Ventes_Dollars', 'Pct_Close', 'Objectif_Dollars']];
      ALL_KPI_REPS.forEach(rep => {
        const r = mObj.reps?.[rep];
        if (!r) return;
        rows.push([
          selVal,
          rep,
          r['Booké'].cumul,
          r['RDV Réalisé'].cumul,
          r['Ventes OTS'].cumul,
          r['Ventes SUIVI'].cumul,
          r['Vente $'].cumul,
          r['% Close'].cumul,
          r['Vente $'].objectif
        ]);
      });
      const csv = rows.map(r => r.join(',')).join('\n');
      const a = document.createElement('a');
      a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8;' }));
      a.download = `korpo-kpi-${selVal.toLowerCase()}.csv`;
      a.click();
    }

    // =========================================================================
    // IMPORTATION DU FICHIER CONDENSÉ (EXCEL .XLSX / .XLS OU .CSV)
    // =========================================================================
    let parsedCondensedRows = []; // [{ ym, rep, booke, rdv, ots, suivi, dollars }]

    const kpiImportModal = document.getElementById('kpi-import-modal');
    document.getElementById('btn-kpi-import-open')?.addEventListener('click', () => {
      if (currentProfile?.role !== 'admin') return;
      kpiImportModal.classList.remove('hidden');
    });
    document.getElementById('btn-close-kpi-import')?.addEventListener('click', () => {
      kpiImportModal.classList.add('hidden');
    });
    kpiImportModal?.addEventListener('click', (e) => {
      if (e.target === kpiImportModal) kpiImportModal.classList.add('hidden');
    });

    document.getElementById('btn-download-kpi-template')?.addEventListener('click', () => {
      const sample = [
        ['Periode', 'Representant', 'RDV_Bookes', 'RDV_Realises', 'Ventes_OTS', 'Ventes_Suivi', 'Ventes_Dollars'],
        ['2026-03', 'Olivier', 13, 13, 2, 3, 50165],
        ['2026-03', 'Jean-Sébastien', 32, 28, 5, 4, 82559],
        ['2026-03', 'Gabryel', 31, 31, 6, 4, 46480],
        ['2026-04', 'Olivier', 31, 24, 7, 4, 164870],
        ['2026-04', 'Jean-Sébastien', 66, 62, 9, 15, 232152],
        ['2026-04', 'Gabryel', 74, 68, 13, 19, 189688]
      ];
      const csv = sample.map(r => r.join(',')).join('\n');
      const a = document.createElement('a');
      a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8;' }));
      a.download = 'modele-fichier-condense-korpo.csv';
      a.click();
    });

    function normalizeImportedRepName(raw) {
      if (!raw) return null;
      const s = String(raw).trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
      if (s.includes('olivier')) return 'Olivier';
      if (s.includes('jean') || s === 'j-s' || s === 'js') return 'Jean-Sébastien';
      if (s.includes('gabryel') || s.includes('gabriel')) return 'Gabryel';
      if (s.includes('samuel')) return 'Samuel';
      if (s.includes('alexandre') || s === 'alex') return 'Alexandre';
      return null;
    }

    function normalizeImportedMetricName(raw) {
      if (!raw) return null;
      const s = String(raw).trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
      if (s.includes('vente total') || s.includes('ventes total')) return null;
      if (s === 'booke' || s.includes('book')) return 'booke';
      if (s.includes('rdv') && (s.includes('rea') || s.includes('fait') || s.includes('comple'))) return 'rdv';
      if (s.includes('ots')) return 'ots';
      if (s.includes('suivi')) return 'suivi';
      if (s.includes('vente') && (s.includes('$') || s.includes('dol') || s.includes('vol') || s.includes('ca'))) return 'dollars';
      return null;
    }

    function parseSheetNameToYm(sheetName, fallbackYear = '2026') {
      const s = String(sheetName || '').trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
      if (/^\d{4}-(0[1-9]|1[0-2])$/.test(s)) return s;
      const yearMatch = s.match(/(20\d{2})/);
      const year = yearMatch ? yearMatch[1] : fallbackYear;
      const monthMap = {
        'janvier': '01', 'jan': '01', 'fevrier': '02', 'fev': '02', 'mars': '03', 'mar': '03',
        'avril': '04', 'avr': '04', 'mai': '05', 'juin': '06', 'juillet': '07', 'juil': '07',
        'aout': '08', 'septembre': '09', 'sept': '09', 'sep': '09', 'octobre': '10', 'oct': '10',
        'novembre': '11', 'nov': '11', 'decembre': '12', 'dec': '12'
      };
      for (const [k, v] of Object.entries(monthMap)) {
        if (s.includes(k)) return `${year}-${v}`;
      }
      return null;
    }

    function cleanNumCell(val) {
      if (val === null || val === undefined || val === '') return 0;
      if (typeof val === 'number') return Math.round(val * 100) / 100;
      const n = parseAmount(val);
      return isNaN(n) ? 0 : Math.round(n * 100) / 100;
    }

    function parseCondensedSheetToRecords(aoa, sheetName, fallbackYear = '2026') {
      const map = new Map(); // key: `${ym}__${rep}` -> { ym, rep, booke, rdv, ots, suivi, dollars }
      const getRec = (ym, rep) => {
        const k = `${ym}__${rep}`;
        if (!map.has(k)) {
          map.set(k, { ym, rep, booke: 0, rdv: 0, ots: 0, suivi: 0, dollars: 0 });
        }
        return map.get(k);
      };

      const sheetRep = normalizeImportedRepName(sheetName);
      const sheetYm = parseSheetNameToYm(sheetName, fallbackYear);

      let currentRep = sheetRep || null;
      let currentYm = sheetYm || null;
      let monthCols = []; // [{ idx, ym }]
      let metricCols = {}; // { booke: idx, rdv: idx, ots: idx, suivi: idx, dollars: idx }
      let skipPartnerSection = false;

      for (let r = 0; r < aoa.length; r++) {
        const row = aoa[r] || [];
        const nonEmpty = row.map((c, idx) => ({ val: String(c ?? '').trim(), idx })).filter(x => x.val !== '');
        if (!nonEmpty.length) continue;

        // 1. Format CSV condensé direct (1 ligne = Periode, Representant, Booke, RDV, OTS, Suivi, Dollars)
        if (nonEmpty.length >= 6) {
          const ymDirect = parseSheetNameToYm(nonEmpty[0].val, fallbackYear);
          const repDirect = normalizeImportedRepName(nonEmpty[1].val);
          if (ymDirect && repDirect && !normalizeImportedMetricName(nonEmpty[2].val)) {
            const c0 = nonEmpty[0].idx;
            const rec = getRec(ymDirect, repDirect);
            rec.booke = cleanNumCell(row[c0 + 2]);
            rec.rdv = cleanNumCell(row[c0 + 3]);
            rec.ots = cleanNumCell(row[c0 + 4]);
            rec.suivi = cleanNumCell(row[c0 + 5]);
            rec.dollars = cleanNumCell(row[c0 + 6]);
            continue;
          }
        }

        // 2. Format CSV Korpo exporté (Periode, Representant, Indicateur, ..., Cumul)
        if (nonEmpty.length >= 4) {
          const ymFlat = parseSheetNameToYm(nonEmpty[0].val, fallbackYear);
          const repFlat = normalizeImportedRepName(nonEmpty[1].val);
          const metFlat = normalizeImportedMetricName(nonEmpty[2].val);
          if (ymFlat && repFlat && metFlat) {
            const c0 = nonEmpty[0].idx;
            const rec = getRec(ymFlat, repFlat);
            // Prendre la colonne Cumul (index +8) si présente, sinon la 4e colonne (index +3)
            const valCell = (row[c0 + 8] !== undefined && String(row[c0 + 8]).trim() !== '') ? row[c0 + 8] : row[c0 + 3];
            rec[metFlat] = cleanNumCell(valCell);
            continue;
          }
        }

        // 3. Détecter une ligne d'en-tête avec les MOIS en colonnes (Mars, Avril, Mai...)
        const rowMonthCols = [];
        nonEmpty.forEach(cell => {
          const ymCol = parseSheetNameToYm(cell.val, fallbackYear);
          if (ymCol && !normalizeImportedRepName(cell.val) && !normalizeImportedMetricName(cell.val)) {
            rowMonthCols.push({ idx: cell.idx, ym: ymCol });
          }
        });
        if (rowMonthCols.length >= 2) {
          monthCols = rowMonthCols;
          metricCols = {};
          skipPartnerSection = false;
          const repOnHeader = nonEmpty.map(c => normalizeImportedRepName(c.val)).find(Boolean);
          if (repOnHeader) currentRep = repOnHeader;
          continue;
        }

        // 4. Détecter une ligne d'en-tête avec les INDICATEURS en colonnes (Booké, RDV Réalisé, OTS, Suivi, Vente $)
        const detectedMetricHeaders = {};
        nonEmpty.forEach(cell => {
          const mk = normalizeImportedMetricName(cell.val);
          if (mk && detectedMetricHeaders[mk] === undefined) {
            detectedMetricHeaders[mk] = cell.idx;
          }
        });
        if (Object.keys(detectedMetricHeaders).length >= 3) {
          metricCols = detectedMetricHeaders;
          monthCols = [];
          skipPartnerSection = false;
          const repOnHeader = nonEmpty.map(c => normalizeImportedRepName(c.val)).find(Boolean);
          if (repOnHeader) currentRep = repOnHeader;
          continue;
        }

        const firstLow = nonEmpty[0].val.toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');

        // Ignorer les blocs récapitulatifs d'équipe ou entreprises
        if (
          firstLow.includes('total equipe') ||
          firstLow.includes('equipe') ||
          firstLow.includes('partenaire') ||
          firstLow.includes('compagnie') ||
          firstLow.includes('toitex') ||
          firstLow.includes('maclim') ||
          firstLow.includes('monspray') ||
          firstLow.includes('verrevdm') ||
          firstLow.includes('diamant')
        ) {
          skipPartnerSection = true;
          currentRep = null;
          continue;
        }

        // Détecter un titre de représentant sur une ligne dédiée
        if (nonEmpty.length <= 2) {
          const repTitle = normalizeImportedRepName(nonEmpty[0].val);
          if (repTitle && !normalizeImportedMetricName(nonEmpty[0].val)) {
            currentRep = repTitle;
            skipPartnerSection = false;
            continue;
          }
          const ymTitle = parseSheetNameToYm(nonEmpty[0].val, fallbackYear);
          if (ymTitle && Object.keys(metricCols).length >= 3) {
            currentYm = ymTitle;
            continue;
          }
        }

        if (skipPartnerSection) continue;

        // Cas Layout A : Mois en colonnes (`monthCols.length >= 2`), Indicateur en ligne
        if (monthCols.length >= 2) {
          const inlineRep = nonEmpty.length >= 3 ? normalizeImportedRepName(nonEmpty[0].val) : null;
          const activeRep = inlineRep || currentRep;
          const metricStr = inlineRep ? nonEmpty[1]?.val : nonEmpty[0]?.val;
          const mk = normalizeImportedMetricName(metricStr);
          if (activeRep && mk) {
            monthCols.forEach(mc => {
              const rec = getRec(mc.ym, activeRep);
              rec[mk] = cleanNumCell(row[mc.idx]);
            });
          }
          continue;
        }

        // Cas Layout B : Indicateurs en colonnes (`Object.keys(metricCols).length >= 3`), Mois (ou Rep) en ligne
        if (Object.keys(metricCols).length >= 3) {
          const inlineRep = nonEmpty.map(c => normalizeImportedRepName(c.val)).find(Boolean);
          const inlineYm = nonEmpty.map(c => parseSheetNameToYm(c.val, fallbackYear)).find(Boolean);
          const activeRep = inlineRep || currentRep;
          const activeYm = inlineYm || currentYm;
          if (activeRep && activeYm) {
            const rec = getRec(activeYm, activeRep);
            Object.entries(metricCols).forEach(([mk, cIdx]) => {
              rec[mk] = cleanNumCell(row[cIdx]);
            });
          }
        }
      }

      return Array.from(map.values()).filter(r =>
        r.booke > 0 || r.rdv > 0 || r.ots > 0 || r.suivi > 0 || r.dollars > 0
      );
    }

    function showCondensedImportPreview() {
      const box = document.getElementById('kpi-import-preview-box');
      const tbody = document.getElementById('tbody-kpi-import-preview');
      const sumEl = document.getElementById('kpi-import-summary');

      if (!parsedCondensedRows.length) {
        box.classList.add('hidden');
        return false;
      }

      parsedCondensedRows.sort((a, b) => a.ym.localeCompare(b.ym) || a.rep.localeCompare(b.rep));
      const monthsCount = new Set(parsedCondensedRows.map(r => r.ym)).size;
      const repsCount = new Set(parsedCondensedRows.map(r => r.rep)).size;
      sumEl.textContent = `${parsedCondensedRows.length} ligne(s) détectée(s) (${monthsCount} mois • ${repsCount} représentant(s))`;

      tbody.innerHTML = parsedCondensedRows.map((r, idx) => `
        <tr class="hover:bg-slate-50">
          <td class="py-2 px-3 font-bold text-slate-800">${esc(r.ym)}</td>
          <td class="py-2 px-3 font-semibold text-indigo-700">${esc(r.rep)}</td>
          <td class="py-2 px-3 text-center">
            <input type="number" min="0" value="${r.booke}" onchange="parsedCondensedRows[${idx}].booke = parseFloat(this.value) || 0" class="w-16 px-1.5 py-0.5 text-center font-bold bg-slate-50 border border-slate-200 rounded-md tabular-nums" />
          </td>
          <td class="py-2 px-3 text-center">
            <input type="number" min="0" value="${r.rdv}" onchange="parsedCondensedRows[${idx}].rdv = parseFloat(this.value) || 0" class="w-16 px-1.5 py-0.5 text-center font-bold bg-slate-50 border border-slate-200 rounded-md tabular-nums" />
          </td>
          <td class="py-2 px-3 text-center">
            <input type="number" min="0" value="${r.ots}" onchange="parsedCondensedRows[${idx}].ots = parseFloat(this.value) || 0" class="w-16 px-1.5 py-0.5 text-center font-extrabold text-sky-700 bg-sky-50/50 border border-sky-200 rounded-md tabular-nums" />
          </td>
          <td class="py-2 px-3 text-center">
            <input type="number" min="0" value="${r.suivi}" onchange="parsedCondensedRows[${idx}].suivi = parseFloat(this.value) || 0" class="w-16 px-1.5 py-0.5 text-center font-extrabold text-purple-700 bg-purple-50/50 border border-purple-200 rounded-md tabular-nums" />
          </td>
          <td class="py-2 px-3 text-right">
            <input type="number" step="0.01" min="0" value="${r.dollars}" onchange="parsedCondensedRows[${idx}].dollars = parseFloat(this.value) || 0" class="w-28 px-2 py-0.5 text-right font-extrabold text-slate-900 bg-slate-50 border border-slate-200 rounded-md tabular-nums" />
          </td>
        </tr>
      `).join('');

      box.classList.remove('hidden');
      return true;
    }

    document.getElementById('kpi-import-file-input')?.addEventListener('change', async (e) => {
      const file = e.target.files[0];
      if (!file) return;
      parsedCondensedRows = [];

      try {
        const ext = file.name.split('.').pop().toLowerCase();
        const selYear = document.getElementById('kpi-import-default-year')?.value || '2026';
        const fileYearMatch = file.name.match(/(20\d{2})/);
        const fallbackYear = fileYearMatch ? fileYearMatch[1] : selYear;

        const mergedMap = new Map();
        const mergeRows = (rows) => {
          rows.forEach(r => {
            const k = `${r.ym}__${r.rep}`;
            mergedMap.set(k, { ...r });
          });
        };

        if (ext === 'xlsx' || ext === 'xls') {
          await ensureXlsxLoaded();
          const buf = await file.arrayBuffer();
          const wb = XLSX.read(buf, { type: 'array' });
          wb.SheetNames.forEach(sName => {
            const aoa = XLSX.utils.sheet_to_json(wb.Sheets[sName], { header: 1, defval: '' });
            const sheetRecs = parseCondensedSheetToRecords(aoa, sName, fallbackYear);
            mergeRows(sheetRecs);
          });
        } else {
          const text = await file.text();
          const lines = text.split(/\r?\n/).filter(l => l.trim());
          const sep = lines[0]?.includes(';') ? ';' : ',';
          const aoa = lines.map(l => parseCsvLine(l, sep));
          const csvRecs = parseCondensedSheetToRecords(aoa, file.name, fallbackYear);
          mergeRows(csvRecs);
        }

        parsedCondensedRows = Array.from(mergedMap.values());
        if (!showCondensedImportPreview()) {
          return toast('Aucune donnée reconnue dans ce fichier condensé. Vérifiez que les noms des mois et des représentants sont lisibles.', 'error');
        }
        toast(`${parsedCondensedRows.length} ligne(s) du fichier condensé prêtes à enregistrer ✓`);
      } catch (err) {
        toast('Erreur de lecture du fichier : ' + err.message, 'error');
      } finally {
        e.target.value = '';
      }
    });

    document.getElementById('btn-confirm-kpi-import')?.addEventListener('click', async () => {
      if (currentProfile?.role !== 'admin') return toast('Action réservée aux administrateurs.', 'error');
      if (!parsedCondensedRows.length) return;
      const btn = document.getElementById('btn-confirm-kpi-import');
      btn.disabled = true;
      btn.textContent = 'Enregistrement en cours...';

      try {
        const replaceAll = document.getElementById('kpi-import-replace-all')?.checked;
        if (replaceAll) {
          condensedBaselineStore = {};
          if (sb && currentProfile) {
            try {
              await sb.from('kpi_hebdo').delete().eq('semaine_index', -9);
            } catch (e) {}
          }
        }

        const nowIso = new Date().toISOString();
        const impactedMonths = new Set();

        for (const r of parsedCondensedRows) {
          if (!isValidYearMonth(r.ym) || !r.rep) continue;
          impactedMonths.add(r.ym);
          if (!condensedBaselineStore[r.ym]) condensedBaselineStore[r.ym] = {};
          const entryObj = {
            booke: Number(r.booke || 0),
            rdv: Number(r.rdv || 0),
            ots: Number(r.ots || 0),
            suivi: Number(r.suivi || 0),
            dollars: Math.round(Number(r.dollars || 0) * 100) / 100,
            updatedAt: nowIso
          };
          condensedBaselineStore[r.ym][r.rep] = entryObj;
          await persistCondensedEntryToCloud(r.ym, r.rep, entryObj);
        }

        saveCondensedBaselineStore();
        rebuildKpiStore();
        kpiImportModal.classList.add('hidden');
        document.getElementById('kpi-import-preview-box').classList.add('hidden');
        renderKpiDashboard();
        renderExecutiveForecast();
      renderFutureMonthsBreakdown();
        toast(`Fichier condensé enregistré sur ${impactedMonths.size} mois ! Vos futures ventes s'y ajouteront automatiquement ✓`);
      } catch (err) {
        toast('Erreur lors de l’enregistrement : ' + err.message, 'error');
      } finally {
        btn.disabled = false;
        btn.textContent = '✓ Enregistrer ce fichier condensé comme base KPI';
      }
    });

    window.switchKpiImportTab = (tabName) => {
      ['file', 'grid'].forEach(t => {
        const view = document.getElementById(`kpi-import-view-${t}`);
        const btn = document.getElementById(`kpi-modal-tab-${t}`);
        const active = t === tabName;
        if (view) view.classList.toggle('hidden', !active);
        if (btn) {
          btn.classList.toggle('border-indigo-600', active);
          btn.classList.toggle('text-indigo-700', active);
          btn.classList.toggle('bg-white', active);
          btn.classList.toggle('border-transparent', !active);
          btn.classList.toggle('text-slate-600', !active);
        }
      });
      if (tabName === 'grid') {
        populateKpiDirectGridMonths();
        renderKpiDirectGrid();
      }
    };

    function populateKpiDirectGridMonths() {
      const sel = document.getElementById('kpi-grid-month-select');
      if (!sel) return;
      const curMain = document.getElementById('kpi-select-month')?.value;
      const months = Object.keys(kpiStore).filter(isValidYearMonth).sort().reverse();
      sel.innerHTML = months.map(ym => `<option value="${ym}">${esc(kpiStore[ym].label)}</option>`).join('');
      if (curMain && !curMain.startsWith('YEAR_') && months.includes(curMain)) {
        sel.value = curMain;
      }
    }

    function renderKpiDirectGrid() {
      const ym = document.getElementById('kpi-grid-month-select')?.value || currentMonthStr;
      ensureKpiMonthStructure(ym);
      const tbody = document.getElementById('tbody-kpi-direct-grid');
      if (!tbody) return;

      tbody.innerHTML = ALL_KPI_REPS.map(rn => {
        const base = condensedBaselineStore[ym]?.[rn] || { booke: 0, rdv: 0, ots: 0, suivi: 0, dollars: 0 };
        const appSalesDol = (state.ventes || [])
          .filter(v => {
            if (getSaleStatsMonth(v) !== ym) return false;
            const rObj = state.representants.find(r => String(r.id) === String(v.representant_id));
            return resolveKpiRepName(rObj?.nom || v.rep_nom || '') === rn;
          })
          .reduce((s, v) => s + Number(v.valeur || 0), 0);

        return `
          <tr class="hover:bg-slate-50" data-rep="${esc(rn)}">
            <td class="py-2.5 px-4 font-bold text-slate-900">${esc(rn)}</td>
            <td class="py-2 px-3 text-center">
              <input type="number" min="0" data-field="booke" value="${Number(base.booke || 0)}" class="w-16 px-2 py-1 text-center font-bold bg-slate-50 border border-slate-200 rounded-lg tabular-nums" />
            </td>
            <td class="py-2 px-3 text-center">
              <input type="number" min="0" data-field="rdv" value="${Number(base.rdv || 0)}" class="w-16 px-2 py-1 text-center font-bold bg-slate-50 border border-slate-200 rounded-lg tabular-nums" />
            </td>
            <td class="py-2 px-3 text-center">
              <input type="number" min="0" data-field="ots" value="${Number(base.ots || 0)}" class="w-16 px-2 py-1 text-center font-extrabold text-sky-700 bg-sky-50/50 border border-sky-200 rounded-lg tabular-nums" />
            </td>
            <td class="py-2 px-3 text-center">
              <input type="number" min="0" data-field="suivi" value="${Number(base.suivi || 0)}" class="w-16 px-2 py-1 text-center font-extrabold text-purple-700 bg-purple-50/50 border border-purple-200 rounded-lg tabular-nums" />
            </td>
            <td class="py-2 px-3 text-right">
              <input type="number" step="0.01" min="0" data-field="dollars" value="${Number(base.dollars || 0)}" class="w-28 px-2 py-1 text-right font-extrabold text-slate-900 bg-slate-50 border border-slate-200 rounded-lg tabular-nums" />
            </td>
            <td class="py-2 px-3 text-right font-bold text-emerald-700 tabular-nums">+ ${fmtMoney0(appSalesDol)}</td>
          </tr>
        `;
      }).join('');
    }

    document.getElementById('kpi-grid-month-select')?.addEventListener('change', renderKpiDirectGrid);

    document.getElementById('btn-save-kpi-direct-grid')?.addEventListener('click', async () => {
      if (currentProfile?.role !== 'admin') return toast('Action réservée aux administrateurs.', 'error');
      const ym = document.getElementById('kpi-grid-month-select')?.value;
      if (!isValidYearMonth(ym)) return;
      const rows = document.querySelectorAll('#tbody-kpi-direct-grid tr[data-rep]');
      const nowIso = new Date().toISOString();
      if (!condensedBaselineStore[ym]) condensedBaselineStore[ym] = {};

      for (const tr of rows) {
        const rn = tr.getAttribute('data-rep');
        const entryObj = {
          booke: Math.max(0, parseFloat(tr.querySelector('input[data-field="booke"]').value) || 0),
          rdv: Math.max(0, parseFloat(tr.querySelector('input[data-field="rdv"]').value) || 0),
          ots: Math.max(0, parseFloat(tr.querySelector('input[data-field="ots"]').value) || 0),
          suivi: Math.max(0, parseFloat(tr.querySelector('input[data-field="suivi"]').value) || 0),
          dollars: Math.max(0, Math.round((parseFloat(tr.querySelector('input[data-field="dollars"]').value) || 0) * 100) / 100),
          updatedAt: nowIso
        };
        condensedBaselineStore[ym][rn] = entryObj;
        await persistCondensedEntryToCloud(ym, rn, entryObj);
      }

      saveCondensedBaselineStore();
      rebuildKpiStore();
      renderKpiDashboard();
      renderExecutiveForecast();
      toast(`Base condensée mise à jour pour ${kpiStore[ym]?.label || ym} ✓`);
    });

    document.getElementById('btn-clear-condensed-baseline')?.addEventListener('click', async () => {
      if (currentProfile?.role !== 'admin') return toast('Action réservée aux administrateurs.', 'error');
      const ok = await askConfirm(
        'Effacer toute la base condensée ?',
        'Le Tableau KPI s’appuiera désormais uniquement sur les ventes enregistrées dans le fichier de ventes.'
      );
      if (!ok) return;
      condensedBaselineStore = {};
      saveCondensedBaselineStore();
      if (sb && currentProfile) {
        try {
          await sb.from('kpi_hebdo').delete().eq('semaine_index', -9);
        } catch (e) {}
      }
      rebuildKpiStore();
      renderKpiDirectGrid();
      renderKpiDashboard();
      renderExecutiveForecast();
      toast('Base condensée effacée : mode 100 % fichier de ventes actif ✓');
    });

    document.getElementById('btn-export-editable-kpi-csv')?.addEventListener('click', () => {
      const rows = [['Periode', 'Representant', 'RDV_Bookes', 'RDV_Realises', 'Ventes_OTS', 'Ventes_Suivi', 'Ventes_Dollars']];
      Object.keys(kpiStore).filter(isValidYearMonth).sort().forEach(ym => {
        const mObj = kpiStore[ym];
        ALL_KPI_REPS.forEach(rep => {
          const r = mObj.reps?.[rep];
          if (!r) return;
          rows.push([
            ym,
            rep,
            r['Booké'].cumul,
            r['RDV Réalisé'].cumul,
            r['Ventes OTS'].cumul,
            r['Ventes SUIVI'].cumul,
            r['Vente $'].cumul
          ]);
        });
      });
      const csv = rows.map(r => r.join(',')).join('\n');
      const a = document.createElement('a');
      a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8;' }));
      a.download = 'korpo-kpi-condense-export.csv';
      a.click();
      toast('Export CSV condensé téléchargé ✓');
    });

    // =========================================================================
    // CANAL DE CONVERSATION ÉQUIPE, MENTIONS @ & PUBLICATION AUTOMATIQUE DES VENTES
    // =========================================================================
    let teamChatMessages = []; // [{ id, kind: 'sale'|'msg', at, authorId, authorName, authorRole, text, mentions, saleData }]
    let activeChatTabOpen = false;
    const chatSalesOnlyMode = false;
    let chatRealtimeChannel = null;
    let chatPollingTimer = null;
    function readSavedChatLastSeenTimestamp() {
      try {
        const raw = localStorage.getItem('korpo_chat_last_seen_v1') || '';
        if (!raw) return '';
        if (raw.startsWith('{')) {
          const obj = JSON.parse(raw);
          return String(obj?.ts || '');
        }
        return raw;
      } catch (e) {
        return '';
      }
    }
    let lastSeenChatTimestamp = readSavedChatLastSeenTimestamp();
    let readMentionIds = new Set(lsGetJson('korpo_read_mention_ids_v1', '[]'));
    let alertedMentionIds = new Set();
    let hasInitializedMentionsBaseline = false;
    let mentionAcCandidates = [];
    let mentionAcSelectedIndex = 0;

    function saveReadMentionIds() {
      const arr = Array.from(readMentionIds).slice(-150);
      lsSet('korpo_read_mention_ids_v1', JSON.stringify(arr));
    }

    function normalizeMentionStr(s) {
      return String(s || '')
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/[^a-z0-9]/g, '');
    }

    function getMentionableTeamMembers() {
      const map = new Map();
      (state.representants || []).forEach(r => {
        if (r && r.nom) {
          map.set(String(r.nom).trim().toLowerCase(), {
            id: String(r.id || ''),
            nom: String(r.nom).trim(),
            role: r.role || 'representant'
          });
        }
      });
      ALL_KPI_REPS.forEach(rn => {
        const key = rn.toLowerCase();
        if (!map.has(key)) {
          map.set(key, { id: '', nom: rn, role: rn === 'Olivier' ? 'admin' : 'representant' });
        }
      });
      return Array.from(map.values()).sort((a, b) => a.nom.localeCompare(b.nom));
    }

    function extractMentionedNamesFromText(text) {
      const raw = String(text || '');
      if (!raw.includes('@')) return [];
      const members = getMentionableTeamMembers();
      const found = new Set();

      // Match @Name (including hyphenated names like @Jean-Sébastien or @Olivier)
      const regex = /@([A-Za-zÀ-ÖØ-öø-ÿ0-9_-]+)/g;
      let match;
      while ((match = regex.exec(raw)) !== null) {
        const tokenNorm = normalizeMentionStr(match[1]);
        if (!tokenNorm) continue;
        if (tokenNorm === 'tous' || tokenNorm === 'equipe' || tokenNorm === 'all') {
          members.forEach(m => found.add(m.nom));
          continue;
        }
        const exact = members.find(m => normalizeMentionStr(m.nom) === tokenNorm);
        if (exact) {
          found.add(exact.nom);
        } else {
          const partial = members.find(m => normalizeMentionStr(m.nom).startsWith(tokenNorm) && tokenNorm.length >= 2);
          if (partial) found.add(partial.nom);
        }
      }
      return Array.from(found);
    }

    function isCurrentUserMentionedInMsg(m) {
      if (!m || m.kind !== 'msg' || !currentProfile) return false;
      if (String(m.authorId) === String(currentProfile.id)) return false;
      const myNorm = normalizeMentionStr(currentProfile.nom);
      if (!myNorm) return false;
      const explicitList = Array.isArray(m.mentions) && m.mentions.length
        ? m.mentions
        : extractMentionedNamesFromText(m.text);
      return explicitList.some(name => normalizeMentionStr(name) === myNorm);
    }

    function formatChatMessageWithMentions(text, isMeBubble = false) {
      const members = getMentionableTeamMembers();
      const myNorm = normalizeMentionStr(currentProfile?.nom || '');
      const escaped = esc(text || '');
      return escaped.replace(/@([A-Za-zÀ-ÖØ-öø-ÿ0-9_-]+)/g, (full, handle) => {
        const normH = normalizeMentionStr(handle);
        const matched = members.find(m =>
          normalizeMentionStr(m.nom) === normH ||
          (normH.length >= 2 && normalizeMentionStr(m.nom).startsWith(normH))
        ) || (['tous', 'equipe', 'all'].includes(normH) ? { nom: handle } : null);

        if (!matched) return full;
        const isMeTag = myNorm && normalizeMentionStr(matched.nom) === myNorm;
        const pillCls = isMeBubble
          ? 'chat-mention-pill-own-bubble bg-white/25 text-white font-extrabold px-1.5 py-0.5 rounded-md'
          : (isMeTag
            ? 'chat-mention-pill-me bg-amber-200/90 text-amber-950 font-extrabold px-1.5 py-0.5 rounded-md ring-1 ring-amber-400'
            : 'chat-mention-pill-other bg-indigo-100 text-indigo-800 font-extrabold px-1.5 py-0.5 rounded-md');
        return `<span class="${pillCls}">@${esc(handle)}</span>`;
      });
    }

    function formatChatTime(isoStr) {
      if (!isoStr) return '';
      try {
        const d = new Date(isoStr);
        const isToday = localYmd(d) === localYmd(new Date());
        const timePart = d.toLocaleTimeString('fr-CA', { hour: '2-digit', minute: '2-digit' });
        if (isToday) return `Aujourd'hui à ${timePart}`;
        return d.toLocaleDateString('fr-CA', { day: 'numeric', month: 'short' }) + ` • ${timePart}`;
      } catch (e) {
        return isoStr;
      }
    }

    function updateChatUnreadBadges() {
      const unreadCount = teamChatMessages.filter(m => {
        if (String(m.authorId) === String(currentProfile?.id)) return false;
        if (!lastSeenChatTimestamp) return true;
        return String(m.at || '') > lastSeenChatTimestamp;
      }).length;

      document.querySelectorAll('.chat-unread-badge').forEach(el => {
        if (unreadCount > 0 && !activeChatTabOpen) {
          el.textContent = unreadCount > 99 ? '99+' : String(unreadCount);
          el.classList.remove('hidden');
        } else {
          el.classList.add('hidden');
        }
      });
    }

    function renderChatMentionNotifications() {
      if (!currentProfile) return;
      const myMentions = teamChatMessages
        .filter(m => isCurrentUserMentionedInMsg(m))
        .sort((a, b) => String(b.at || '').localeCompare(String(a.at || '')));

      const unreadMentions = myMentions.filter(m => !readMentionIds.has(String(m.id)));
      const topBanner = document.getElementById('chat-mention-top-banner');

      if (!hasInitializedMentionsBaseline) {
        unreadMentions.forEach(m => alertedMentionIds.add(String(m.id)));
        hasInitializedMentionsBaseline = true;
      } else {
        unreadMentions.forEach(m => {
          const mid = String(m.id);
          if (!alertedMentionIds.has(mid)) {
            alertedMentionIds.add(mid);
            playNotificationChime();
            toast(`@ ${m.authorName} vous a identifié : "${String(m.text || '').slice(0, 65)}"`);
            sendChatMentionSystemNotification(m);
          }
        });
      }

      if (topBanner) {
        if (unreadMentions.length > 0) {
          const latest = unreadMentions[0];
          topBanner.dataset.targetMsgId = String(latest.id);
          document.getElementById('chat-mention-banner-title').textContent =
            unreadMentions.length === 1
              ? `${latest.authorName} vous a tagué dans le canal`
              : `${unreadMentions.length} messages vous identifient dans le canal (dernier par ${latest.authorName})`;
          document.getElementById('chat-mention-banner-desc').textContent =
            `« ${String(latest.text || '')} » — Cliquez pour ouvrir directement le message`;
          topBanner.classList.remove('hidden');
        } else {
          topBanner.classList.add('hidden');
        }
      }

      renderUnifiedBellNotifications();
    }

    window.markAllChatMentionsRead = () => {
      teamChatMessages.forEach(m => {
        if (isCurrentUserMentionedInMsg(m)) {
          readMentionIds.add(String(m.id));
        }
      });
      saveReadMentionIds();
      renderChatMentionNotifications();
    };

    window.jumpToLatestUnreadMention = () => {
      const banner = document.getElementById('chat-mention-top-banner');
      const targetId = banner?.dataset?.targetMsgId;
      if (targetId) {
        jumpToChatMessage(targetId);
      } else {
        openTeamChatModal();
      }
    };

    window.jumpToChatMessage = (msgId) => {
      document.getElementById('admin-notif-dropdown')?.classList.add('hidden');
      const sid = String(msgId);
      readMentionIds.add(sid);
      saveReadMentionIds();

      // Si le filtre "Ventes seulement" était actif, le désactiver pour afficher le message texte tagué
      if (chatSalesOnlyMode) {
        chatSalesOnlyMode = false;
        const btn = document.getElementById('btn-chat-filter-sales');
        if (btn) {
          btn.classList.remove('bg-emerald-500', 'text-slate-950');
          btn.classList.add('bg-white/10', 'text-slate-200');
        }
      }

      const backdrop = document.getElementById('chat-modal-backdrop');
      if (backdrop) backdrop.classList.remove('hidden');
      setChatBackgroundScrollLocked(true);
      activeChatTabOpen = true;
      markTeamChatAsRead();
      renderTeamChatFeed(false);
      renderChatMentionNotifications();

      setTimeout(() => {
        const el = document.getElementById(`chat-item-${sid}`);
        if (el) {
          el.scrollIntoView({ behavior: 'smooth', block: 'center' });
          el.classList.add('ring-2', 'ring-amber-400', 'bg-amber-50/70', 'dark:bg-amber-500/15', 'rounded-2xl', 'p-2', '-mx-2', 'transition-all');
          setTimeout(() => {
            el.classList.remove('ring-2', 'ring-amber-400', 'bg-amber-50/70', 'dark:bg-amber-500/15', 'p-2', '-mx-2');
          }, 2800);
        }
      }, 90);
    };

    function markTeamChatAsRead() {
      const latestAt = teamChatMessages.length
        ? teamChatMessages[teamChatMessages.length - 1].at
        : new Date().toISOString();
      let changedSeen = false;
      if (latestAt && (!lastSeenChatTimestamp || latestAt > lastSeenChatTimestamp)) {
        lastSeenChatTimestamp = latestAt;
        changedSeen = true;
      }
      if (lastSeenChatTimestamp) {
        lsSet('korpo_chat_last_seen_v1', JSON.stringify({ ts: lastSeenChatTimestamp }));
      }
      teamChatMessages.forEach(m => {
        if (isCurrentUserMentionedInMsg(m)) {
          readMentionIds.add(String(m.id));
        }
      });
      saveReadMentionIds();
      if (changedSeen && typeof kvFlush === 'function') {
        kvFlush().catch(() => {});
      }
      updateChatUnreadBadges();
      renderChatMentionNotifications();
    }

    function buildTeamChatMessagesFromState() {
      const itemsMap = new Map();

      const isAdminViewer = currentProfile?.role === 'admin';
      const unconfirmedOtherSaleIds = new Set(
        (state.ventes || [])
          .filter(v => String(v.representant_id) !== String(currentProfile?.id) && v.statut_depot === 'a_confirmer')
          .map(v => `sale_${v.id}`)
      );

      // 1. Messages & annonces stockés dans kpi_hebdo (semaine_index = -12) — fenêtre glissante des 150 plus récents par défaut
      const rawChatRecords = [];
      (kpiHebdoRows || []).forEach(r => {
        if (Number(r.semaine_index) !== -12 || !r.representant_nom) return;
        try {
          const parsed = JSON.parse(r.representant_nom);
          if (parsed && parsed.id) {
            if (!isAdminViewer && parsed.kind === 'sale' && unconfirmedOtherSaleIds.has(String(parsed.id))) {
              return;
            }
            rawChatRecords.push(parsed);
          }
        } catch (e) {}
      });
      rawChatRecords.sort((a, b) => String(a.at || '').localeCompare(String(b.at || '')));
      const chatLimit = window._chatShowAllHistory ? rawChatRecords.length : 150;
      window._chatHiddenOlderCount = Math.max(0, rawChatRecords.length - chatLimit);
      rawChatRecords.slice(-chatLimit).forEach(parsed => {
        itemsMap.set(String(parsed.id), parsed);
      });

      // 2. Garantir que toutes les ventes récentes apparaissent dans le canal (les ventes non confirmées des autres reps restent invisibles pour un représentant)
      const isAdminUser = currentProfile?.role === 'admin';
      const sortedSales = [...(state.ventes || [])]
        .filter(v => {
          if (isAdminUser) return true;
          if (String(v.representant_id) === String(currentProfile?.id)) return true;
          return v.statut_depot === 'depot_recu';
        })
        .sort((a, b) => String(a.created_at || a.date_vente || '').localeCompare(String(b.created_at || b.date_vente || '')))
        .slice(-40);

      sortedSales.forEach(v => {
        const autoId = `sale_${v.id}`;
        if (itemsMap.has(autoId)) return;
        const repObj = state.representants.find(r => String(r.id) === String(v.representant_id));
        const repNom = repObj?.nom || v.rep_nom || 'Représentant';
        const entNom = state.entreprises.find(e => String(e.id) === String(v.entreprise_id))?.nom || v.entreprise_nom || 'Entreprise';
        const isOwnSale = String(v.representant_id) === String(currentProfile?.id);
        const rawCli = v.client_nom || state.clients.find(c => String(c.id) === String(v.client_id))?.nom || '';
        const cliDisplay = (isAdminUser || isOwnSale) && rawCli && rawCli !== 'Client confidentiel' ? rawCli : '';
        const noteUp = String(v.note || '').toUpperCase();
        const typeKpi = noteUp.includes('[SUIVI]') ? 'SUIVI' : 'OTS';
        const atIso = v.created_at || (v.date_vente ? `${v.date_vente}T12:00:00.000Z` : new Date().toISOString());

        itemsMap.set(autoId, {
          id: autoId,
          kind: 'sale',
          at: atIso,
          authorId: String(v.representant_id || ''),
          authorName: repNom,
          authorRole: repObj?.role || 'representant',
          saleData: {
            saleId: String(v.id),
            repNom,
            entNom,
            clientNom: cliDisplay,
            valeur: Number(v.valeur || 0),
            typeKpi,
            dateVente: v.date_vente || ''
          }
        });
      });

      teamChatMessages = Array.from(itemsMap.values()).sort((a, b) =>
        String(a.at || '').localeCompare(String(b.at || ''))
      );

      // Détecter une nouvelle vente d'un collègue en temps réel pour lancer la fusée d'équipe
      if (!window._knownChatSaleIds) {
        window._knownChatSaleIds = new Set(teamChatMessages.filter(m => m.kind === 'sale').map(m => String(m.id)));
      } else if (currentProfile) {
        const freshCutoffIso = new Date(Date.now() - 2 * 60 * 1000).toISOString();
        teamChatMessages.forEach(m => {
          if (m.kind !== 'sale') return;
          const sid = String(m.id);
          if (window._knownChatSaleIds.has(sid)) return;
          window._knownChatSaleIds.add(sid);
          if (String(m.authorId) === String(currentProfile.id)) return;
          if (m.at && String(m.at) < freshCutoffIso) return;
          const sd = m.saleData || {};
          const localSale = (state.ventes || []).find(v => `sale_${v.id}` === sid);
          const commForViewer = currentProfile.role === 'admin'
            ? Number(localSale?.commission_olivier || 0)
            : 0;
          triggerSaleCelebration(
            sd.repNom || m.authorName || 'Équipe',
            sd.clientNom || '',
            sd.entNom || 'Korpo',
            Number(sd.valeur || 0),
            commForViewer,
            commForViewer === 0
          );
        });
      }

      if (activeChatTabOpen) {
        markTeamChatAsRead();
      } else {
        updateChatUnreadBadges();
        renderChatMentionNotifications();
      }
    }

    async function publishAutoSaleToTeamChat(saleRow, repNom, entNom, clientNom, typeKpi) {
      if (!sb || !currentProfile || !saleRow?.id) return;
      const msgId = `sale_${saleRow.id}`;
      const atIso = saleRow.created_at || new Date().toISOString();
      const chatObj = {
        id: msgId,
        kind: 'sale',
        at: atIso,
        authorId: String(saleRow.representant_id || currentProfile.id),
        authorName: repNom,
        authorRole: currentProfile.role || 'representant',
        saleData: {
          saleId: String(saleRow.id),
          repNom,
          entNom,
          clientNom: '',
          valeur: Number(saleRow.valeur || 0),
          typeKpi: typeKpi || 'OTS',
          dateVente: saleRow.date_vente || localYmd()
        }
      };

      kpiHebdoRows.push({
        mois_commission: msgId,
        semaine_index: -12,
        representant_nom: JSON.stringify(chatObj),
        booke: 0,
        rdv_realise: 0,
        updated_at: atIso
      });
      buildTeamChatMessagesFromState();
      renderTeamChatFeed(true);

      await safeUpsertKpiHebdoRow({
        mois_commission: msgId,
        semaine_index: -12,
        representant_nom: JSON.stringify(chatObj),
        booke: 0,
        rdv_realise: 0,
        updated_at: atIso
      });
      triggerAdminSaleWebPush(chatObj);
    }

    let chatReplyTarget = null; // { id, authorName, snippet }
    let pendingChatPhotoDataUrl = '';
    let chatMediaRecorder = null;
    let chatAudioChunks = [];
    let chatVoiceStartTs = 0;
    let chatVoiceTimerInterval = null;
    let chatVoiceCancelled = false;
    const chatMediaUrlCache = new Map();

    // Médias du chat (vocaux & photos) : stockés dans Supabase Storage (bucket privé « preuves-ventes »,
    // dossier <id utilisateur>/chat/). Le message ne contient que le chemin du fichier.
    // Repli (rare) : une seule ligne `korpo_store` si Storage est indisponible.
    // Les anciens médias (refs `media:` et `chunk:`) restent lisibles.
    const chatMediaFailAt = new Map(); // ref -> timestamp du dernier échec de résolution

    async function uploadChatMediaBlob(blob, ext, mimeType) {
      const mediaId = `cm_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
      window._lastChatMediaError = '';

      // Tentative 1 : Supabase Storage
      try {
        const path = `${currentUser?.id || 'anon'}/chat/${mediaId}.${ext}`;
        const { error } = await sb.storage.from(PROOF_BUCKET).upload(path, blob, { contentType: mimeType, upsert: false });
        if (error) throw error;
        chatMediaUrlCache.set(path, { url: URL.createObjectURL(blob), exp: Date.now() + 2 * 60 * 1000 });
        return path;
      } catch (e) {
        window._lastChatMediaError = 'Storage : ' + String((e && (e.message || e.error)) || e);
        console.warn('Échec Storage média chat :', e);
      }

      // Tentative 2 : une seule ligne korpo_store (petits fichiers seulement)
      try {
        if (blob.size > 1500000) throw new Error('fichier trop volumineux pour le repli');
        const dataUrl = await new Promise((resolve) => {
          const reader = new FileReader();
          reader.onloadend = () => resolve(String(reader.result || ''));
          reader.readAsDataURL(blob);
        });
        if (!dataUrl) throw new Error('lecture du fichier impossible');
        const { error } = await sb.from(KV_TABLE).upsert({
          store: 'chat_media',
          item_key: mediaId,
          owner: 'shared',
          value: { dataUrl, mimeType, ext, at: new Date().toISOString() },
          updated_at: new Date().toISOString()
        }, { onConflict: 'store,item_key,owner' });
        if (error) throw error;
        chatMediaUrlCache.set(`media:${mediaId}`, { url: dataUrl, exp: Date.now() + 2 * 60 * 1000 });
        return `media:${mediaId}`;
      } catch (e2) {
        window._lastChatMediaError += ' | Repli : ' + String((e2 && (e2.message || e2.error)) || e2);
        console.warn('Échec repli média chat :', e2);
        return '';
      }
    }

    async function resolveChatMediaSignedUrls() {
      const now = Date.now();
      const mediaKeys = [];
      const chunkPrefixes = [];
      const storagePaths = [];
      const pendingRefs = [];

      (teamChatMessages || []).forEach(m => {
        [m.imageRef, m.audioRef].forEach(ref => {
          if (!ref) return;
          const sRef = String(ref);
          if (chatMediaUrlCache.get(sRef)?.exp > now + 60000) return;
          if ((chatMediaFailAt.get(sRef) || 0) > now - 20000) return; // on réessaie au plus toutes les 20 s
          pendingRefs.push(sRef);
          if (sRef.startsWith('media:')) {
            mediaKeys.push(sRef.slice(6));
          } else if (sRef.startsWith('chunk:')) {
            const parts = sRef.split(':');
            if (parts[1]) chunkPrefixes.push({ fullRef: sRef, mediaId: parts[1], total: Number(parts[2] || 0) });
          } else if (isProofPath(sRef)) {
            storagePaths.push(sRef);
          }
        });
      });

      const uniqMediaKeys = [...new Set(mediaKeys)];
      if (uniqMediaKeys.length && sb) {
        try {
          const { data, error } = await sb
            .from(KV_TABLE)
            .select('item_key,value')
            .eq('store', 'chat_media')
            .eq('owner', 'shared')
            .in('item_key', uniqMediaKeys);
          if (!error && Array.isArray(data)) {
            data.forEach(row => {
              const url = row?.value?.dataUrl;
              if (url) chatMediaUrlCache.set(`media:${row.item_key}`, { url, exp: now + 24 * 3600 * 1000 });
            });
          }
        } catch (e) {}
      }

      // Anciens vocaux/photos découpés en blocs (lecture seule, avec pagination)
      if (chunkPrefixes.length && sb) {
        for (const item of chunkPrefixes) {
          try {
            const data = await fetchAllRows(() =>
              sb.from('kpi_hebdo')
                .select('mois_commission,representant_nom,booke')
                .eq('semaine_index', -13)
                .like('mois_commission', `${item.mediaId}__%`)
                .order('booke')
            );
            if (Array.isArray(data) && data.length > 0 && (!item.total || data.length >= item.total)) {
              data.sort((a, b) => Number(a.booke || 0) - Number(b.booke || 0));
              const fullDataUrl = data.map(r => r.representant_nom || '').join('');
              if (fullDataUrl.startsWith('data:')) {
                chatMediaUrlCache.set(item.fullRef, { url: fullDataUrl, exp: now + 24 * 3600 * 1000 });
              }
            }
          } catch (e) {}
        }
      }

      const uniqStoragePaths = [...new Set(storagePaths)];
      if (uniqStoragePaths.length && sb) {
        try {
          const { data, error } = await sb.storage.from(PROOF_BUCKET).createSignedUrls(uniqStoragePaths, 3600);
          if (error) throw error;
          if (Array.isArray(data)) {
            data.forEach(d => {
              if (d.signedUrl) chatMediaUrlCache.set(d.path, { url: d.signedUrl, exp: now + 3600 * 1000 });
            });
          }
        } catch (e) {
          console.warn('URL signées du chat indisponibles :', e && e.message);
        }
      }

      pendingRefs.forEach(r => {
        if (chatMediaUrlCache.get(r)?.url) chatMediaFailAt.delete(r);
        else chatMediaFailAt.set(r, now);
      });
    }

    function chatFeedSignature() {
      return (teamChatMessages || []).map(m => `${m.id}:${m.audioRef || ''}:${m.imageRef || ''}`).join('|')
        + `|c${chatMediaUrlCache.size}|f${chatMediaFailAt.size}`;
    }

    function chatMediaPendingHtml(ref, label) {
      if (!ref) return '';
      if (chatMediaFailAt.has(String(ref))) {
        return `<div class="py-1.5 px-2 text-xs opacity-90 flex items-center gap-2"><span>⚠️ ${label} indisponible</span><button type="button" onclick="retryChatMediaLoad()" class="underline font-extrabold cursor-pointer">Réessayer</button></div>`;
      }
      return `<div class="py-1.5 px-2 text-xs italic opacity-75 flex items-center gap-2"><span class="w-2 h-2 rounded-full bg-current animate-ping"></span><span>Chargement…</span></div>`;
    }

    window.retryChatMediaLoad = async () => {
      chatMediaFailAt.clear();
      await resolveChatMediaSignedUrls();
      renderTeamChatFeed(false);
    };

    function getChatMediaUrl(ref) {
      if (!ref) return '';
      const sRef = String(ref);
      if (sRef.startsWith('media:') || sRef.startsWith('chunk:') || sRef.startsWith('local_audio_') || isProofPath(sRef)) {
        return chatMediaUrlCache.get(sRef)?.url || '';
      }
      return sRef;
    }

    function autoResizeChatTextarea() {
      const ta = document.getElementById('chat-message-input');
      if (!ta) return;
      ta.style.height = 'auto';
      ta.style.height = Math.min(ta.scrollHeight, 112) + 'px';
    }

    window.setChatReplyTarget = (msgId) => {
      const target = teamChatMessages.find(m => String(m.id) === String(msgId));
      if (!target) return;
      let snippet = '';
      if (target.kind === 'sale') {
        const sd = target.saleData || {};
        snippet = `💵 Vente ${sd.entNom || ''} — ${fmtMoney(sd.valeur)}`;
      } else if (target.audioRef) {
        snippet = `🎤 Message vocal${target.audioDuration ? ` (${target.audioDuration})` : ''}`;
      } else if (target.imageRef && !target.text) {
        snippet = '📷 Photo';
      } else {
        snippet = (target.imageRef ? '📷 ' : '') + String(target.text || '');
      }
      chatReplyTarget = {
        id: String(target.id),
        authorName: target.authorName || target.saleData?.repNom || 'Membre',
        snippet: snippet.slice(0, 120)
      };
      const bar = document.getElementById('chat-reply-preview-bar');
      const authEl = document.getElementById('chat-reply-preview-author');
      const snipEl = document.getElementById('chat-reply-preview-snippet');
      if (authEl) authEl.textContent = `Répondre à ${chatReplyTarget.authorName}`;
      if (snipEl) snipEl.textContent = chatReplyTarget.snippet;
      bar?.classList.remove('hidden');
      document.getElementById('chat-message-input')?.focus();
    };

    window.clearChatReplyTarget = () => {
      chatReplyTarget = null;
      document.getElementById('chat-reply-preview-bar')?.classList.add('hidden');
    };

    window.handleChatPhotoSelected = async (inputEl) => {
      const file = inputEl?.files?.[0];
      if (!file) return;
      try {
        const dataUrl = await compressImageToDataUrl(file, 900, 0.68);
        pendingChatPhotoDataUrl = dataUrl;
        const img = document.getElementById('chat-photo-preview-img');
        if (img) img.src = dataUrl;
        document.getElementById('chat-photo-preview-bar')?.classList.remove('hidden');
        document.getElementById('chat-message-input')?.focus();
      } catch (e) {
        toast('Impossible de charger cette photo.', 'error');
      } finally {
        inputEl.value = '';
      }
    };

    window.clearPendingChatPhoto = () => {
      pendingChatPhotoDataUrl = '';
      const img = document.getElementById('chat-photo-preview-img');
      if (img) img.src = '';
      document.getElementById('chat-photo-preview-bar')?.classList.add('hidden');
    };

    window.openChatImageLightbox = (src) => {
      if (!src) return;
      const modal = document.getElementById('proof-modal');
      const body = document.getElementById('proof-modal-body');
      const delBtn = document.getElementById('btn-delete-proof-modal');
      if (!modal || !body) return;
      document.getElementById('proof-modal-title').textContent = 'Photo partagée — Korpo Top Earner';
      document.getElementById('proof-modal-subtitle').textContent = '';
      if (delBtn) delBtn.classList.add('hidden');
      body.innerHTML = `<img src="${esc(src)}" alt="Photo partagée" class="max-h-[75vh] rounded-xl shadow-md object-contain" />`;
      modal.classList.remove('hidden');
    };

    // Encodeur PCM -> WAV mono 8 000 Hz ultra-léger (~8 Ko/s) + enveloppe visuelle 24 barres façon WhatsApp
    function encodeFloat32SamplesToWavBlob(floatChunks, inputSampleRate) {
      const targetRate = 8000;
      let totalLen = 0;
      for (const c of floatChunks) totalLen += c.length;
      const merged = new Float32Array(totalLen);
      let pos = 0;
      for (const c of floatChunks) {
        merged.set(c, pos);
        pos += c.length;
      }

      const barCount = 24;
      const wavePeaks = [];
      const step = Math.max(1, Math.floor(merged.length / barCount));
      let maxPeak = 0.015;
      const rawRms = [];
      for (let b = 0; b < barCount; b++) {
        let sumSq = 0;
        const start = b * step;
        const end = Math.min(merged.length, start + step);
        const count = Math.max(1, end - start);
        for (let j = start; j < end; j += 4) {
          const v = merged[j];
          sumSq += v * v;
        }
        const rms = Math.sqrt(sumSq / Math.max(1, Math.ceil(count / 4)));
        rawRms.push(rms);
        if (rms > maxPeak) maxPeak = rms;
      }
      for (let b = 0; b < barCount; b++) {
        const norm = Math.round(18 + Math.min(1, rawRms[b] / maxPeak) * 82);
        wavePeaks.push(norm);
      }

      const ratio = Math.max(1, inputSampleRate / targetRate);
      const outLen = Math.max(1, Math.floor(merged.length / ratio));
      const buffer = new ArrayBuffer(44 + outLen * 2);
      const view = new DataView(buffer);
      const writeStr = (off, str) => { for (let i = 0; i < str.length; i++) view.setUint8(off + i, str.charCodeAt(i)); };

      writeStr(0, 'RIFF');
      view.setUint32(4, 36 + outLen * 2, true);
      writeStr(8, 'WAVE');
      writeStr(12, 'fmt ');
      view.setUint32(16, 16, true);
      view.setUint16(20, 1, true); // PCM
      view.setUint16(22, 1, true); // Mono
      view.setUint32(24, targetRate, true);
      view.setUint32(28, targetRate * 2, true);
      view.setUint16(32, 2, true);
      view.setUint16(34, 16, true);
      writeStr(36, 'data');
      view.setUint32(40, outLen * 2, true);

      let offset = 44;
      for (let i = 0; i < outLen; i++, offset += 2) {
        const srcIdx = Math.min(merged.length - 1, Math.floor(i * ratio));
        const s = Math.max(-1, Math.min(1, merged[srcIdx]));
        view.setInt16(offset, s < 0 ? s * 0x8000 : s * 0x7FFF, true);
      }
      return { blob: new Blob([buffer], { type: 'audio/wav' }), wavePeaks };
    }

    let chatVoiceStream = null;
    let chatVoiceAudioCtx = null;
    let chatVoiceSource = null;
    let chatVoiceProcessor = null;
    let chatVoiceSilentGain = null;
    const _pendingVoiceUploads = new Map();
    let chatVoicePcmChunks = [];
    let chatVoiceInputRate = 44100;
    let chatVoiceLiveLevels = new Array(22).fill(18);

    function renderLiveVoiceWaveBars(levelPct = null) {
      const box = document.getElementById('chat-voice-live-wave');
      if (!box) return;
      if (levelPct !== null) {
        chatVoiceLiveLevels.push(levelPct);
        if (chatVoiceLiveLevels.length > 22) chatVoiceLiveLevels.shift();
      }
      box.innerHTML = chatVoiceLiveLevels.map(h =>
        `<span class="wa-live-bar" style="height:${Math.max(16, Math.min(100, h))}%"></span>`
      ).join('');
    }

    function cleanupChatVoiceHardware() {
      clearInterval(chatVoiceTimerInterval);
      if (chatVoiceSource) {
        try { chatVoiceSource.disconnect(); } catch (e) {}
        chatVoiceSource = null;
      }
      if (chatVoiceProcessor) {
        try { chatVoiceProcessor.disconnect(); } catch (e) {}
        chatVoiceProcessor.onaudioprocess = null;
        chatVoiceProcessor = null;
      }
      if (chatVoiceSilentGain) {
        try { chatVoiceSilentGain.disconnect(); } catch (e) {}
        chatVoiceSilentGain = null;
      }
      if (chatVoiceAudioCtx) {
        try { chatVoiceAudioCtx.close(); } catch (e) {}
        chatVoiceAudioCtx = null;
      }
      if (chatVoiceStream) {
        try { chatVoiceStream.getTracks().forEach(t => t.stop()); } catch (e) {}
        chatVoiceStream = null;
      }
      document.getElementById('chat-voice-recording-bar')?.classList.add('hidden');
      document.getElementById('form-team-chat')?.classList.remove('hidden');
    }

    window.startChatVoiceRecording = async () => {
      if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
        return toast('Microphone indisponible sur ce navigateur (HTTPS requis).', 'error');
      }
      const AudioCtx = window.AudioContext || window.webkitAudioContext;
      if (!AudioCtx) {
        return toast('AudioContext non supporté sur cet appareil.', 'error');
      }

      try {
        cleanupChatVoiceHardware();
        const stream = await navigator.mediaDevices.getUserMedia({
          audio: { echoCancellation: true, noiseSuppression: true, autoGainControl: true }
        });
        chatVoiceStream = stream;
        chatVoicePcmChunks = [];
        chatVoiceCancelled = false;
        chatVoiceLiveLevels = new Array(22).fill(18);

        const ctx = new AudioCtx();
        if (ctx.state === 'suspended') {
          await ctx.resume();
        }
        chatVoiceAudioCtx = ctx;
        chatVoiceInputRate = ctx.sampleRate || 44100;

        const source = ctx.createMediaStreamSource(stream);
        const processor = ctx.createScriptProcessor(4096, 1, 1);
        const silentGain = ctx.createGain();
        silentGain.gain.value = 0; // Empêche tout écho / larsen dans le haut-parleur pendant l'enregistrement
        chatVoiceSource = source;
        chatVoiceProcessor = processor;
        chatVoiceSilentGain = silentGain;

        processor.onaudioprocess = (e) => {
          if (chatVoiceCancelled) return;
          const input = e.inputBuffer.getChannelData(0);
          chatVoicePcmChunks.push(new Float32Array(input));
          let peak = 0;
          for (let i = 0; i < input.length; i += 16) {
            const a = Math.abs(input[i]);
            if (a > peak) peak = a;
          }
          renderLiveVoiceWaveBars(Math.round(18 + Math.min(1, peak * 3.8) * 82));
        };

        source.connect(processor);
        processor.connect(silentGain);
        silentGain.connect(ctx.destination);

        chatVoiceStartTs = Date.now();
        const timerEl = document.getElementById('chat-voice-timer');
        if (timerEl) timerEl.textContent = '0:00';
        renderLiveVoiceWaveBars();
        document.getElementById('form-team-chat')?.classList.add('hidden');
        document.getElementById('chat-voice-recording-bar')?.classList.remove('hidden');

        chatVoiceTimerInterval = setInterval(() => {
          const sec = Math.floor((Date.now() - chatVoiceStartTs) / 1000);
          if (timerEl) timerEl.textContent = `${Math.floor(sec / 60)}:${String(sec % 60).padStart(2, '0')}`;
          if (sec >= 60) stopAndSendChatVoiceRecording();
        }, 350);
      } catch (err) {
        cleanupChatVoiceHardware();
        const msg = String(err?.name || err?.message || '');
        if (msg.includes('NotAllowed') || msg.includes('Permission')) {
          toast('Accès au micro refusé : autorisez le microphone dans les réglages.', 'error');
        } else {
          toast('Erreur micro : ' + (err?.message || 'vérifiez les permissions audio'), 'error');
        }
      }
    };

    window.cancelChatVoiceRecording = () => {
      chatVoiceCancelled = true;
      chatVoicePcmChunks = [];
      cleanupChatVoiceHardware();
    };

    window.stopAndSendChatVoiceRecording = async () => {
      chatVoiceCancelled = false;
      const chunksCopy = chatVoicePcmChunks.slice();
      const sampleRate = chatVoiceInputRate;
      const elapsedSec = Math.max(1, Math.round((Date.now() - chatVoiceStartTs) / 1000));
      chatVoicePcmChunks = [];
      cleanupChatVoiceHardware();

      if (!chunksCopy.length) {
        return toast('Aucun son capté. Parlez au moins 1 seconde.', 'error');
      }

      const mins = Math.floor(elapsedSec / 60);
      const secs = String(elapsedSec % 60).padStart(2, '0');
      const durationLabel = `${mins}:${secs}`;

      const { blob: wavBlob, wavePeaks } = encodeFloat32SamplesToWavBlob(chunksCopy, sampleRate);
      const instantLocalUrl = URL.createObjectURL(wavBlob);
      const tempRef = `local_audio_${Date.now()}`;
      chatMediaUrlCache.set(tempRef, { url: instantLocalUrl, exp: Date.now() + 24 * 3600 * 1000 });

      // 1) Affichage immédiat (0 ms) dans la conversation avec lecture instantanée
      const msgId = `msg_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
      const nowIso = new Date().toISOString();
      const replyPayload = chatReplyTarget ? { ...chatReplyTarget } : null;
      clearChatReplyTarget();

      const optimisticObj = {
        id: msgId,
        kind: 'msg',
        at: nowIso,
        authorId: String(currentProfile.id),
        authorName: currentProfile.nom,
        authorRole: currentProfile.role || 'representant',
        text: '',
        mentions: [],
        ...(replyPayload ? { replyTo: replyPayload } : {}),
        audioRef: tempRef,
        audioDuration: durationLabel,
        wavePeaks
      };

      const optimisticRow = {
        mois_commission: msgId,
        semaine_index: -12,
        representant_nom: JSON.stringify(optimisticObj),
        booke: 0,
        rdv_realise: 0,
        updated_at: nowIso
      };
      _pendingVoiceUploads.set(msgId, optimisticRow);
      kpiHebdoRows.push(optimisticRow);
      buildTeamChatMessagesFromState();
      renderTeamChatFeed(true);

      // 2) Envoi cloud en arrière-plan sans jamais interrompre la lecture locale
      try {
        const uploadedRef = await uploadChatMediaBlob(wavBlob, 'wav', 'audio/wav');
        if (!uploadedRef) {
          kpiHebdoRows = kpiHebdoRows.filter(r => !(Number(r.semaine_index) === -12 && String(r.mois_commission) === msgId));
          buildTeamChatMessagesFromState();
          renderTeamChatFeed(false);
          return toast('Vocal non envoyé : ' + (window._lastChatMediaError || 'réseau indisponible'), 'error');
        }
        chatMediaUrlCache.set(uploadedRef, { url: instantLocalUrl, exp: Date.now() + 24 * 3600 * 1000 });
        optimisticObj.audioRef = uploadedRef;
        optimisticRow.representant_nom = JSON.stringify(optimisticObj);
        const rowIdx = kpiHebdoRows.findIndex(r => Number(r.semaine_index) === -12 && String(r.mois_commission) === msgId);
        if (rowIdx >= 0) kpiHebdoRows[rowIdx].representant_nom = JSON.stringify(optimisticObj);

        const voiceRes = await safeUpsertKpiHebdoRow({
          mois_commission: msgId,
          semaine_index: -12,
          representant_nom: JSON.stringify(optimisticObj),
          booke: 0,
          rdv_realise: 0,
          updated_at: nowIso
        });
        if (voiceRes && voiceRes.error) {
          toast('Vocal non enregistré : ' + (voiceRes.error.message || 'erreur'), 'error');
        }
      } finally {
        _pendingVoiceUploads.delete(msgId);
      }
    };

    // Contrôleur du lecteur vocal WhatsApp (1 seule lecture active à la fois + vitesse 1x / 1.5x / 2x + barre interactive)
    let _activeWaAudioEl = null;
    let _activeWaMsgId = null;
    const _waPlaybackRates = [1, 1.5, 2];

    function fmtAudioSec(sec) {
      if (!isFinite(sec) || sec < 0) return '0:00';
      const s = Math.floor(sec);
      return `${Math.floor(s / 60)}:${String(s % 60).padStart(2, '0')}`;
    }

    function syncWaPlayerVisuals(msgId, isPlaying) {
      const playIcon = document.getElementById(`wa-play-icon-${msgId}`);
      const pauseIcon = document.getElementById(`wa-pause-icon-${msgId}`);
      if (playIcon) playIcon.classList.toggle('hidden', isPlaying);
      if (pauseIcon) pauseIcon.classList.toggle('hidden', !isPlaying);
    }

    window.toggleWaVoicePlay = (msgId) => {
      const audio = document.getElementById(`wa-audio-el-${msgId}`);
      if (!audio) return;
      if (_activeWaAudioEl && _activeWaAudioEl !== audio) {
        _activeWaAudioEl.pause();
        if (_activeWaMsgId) syncWaPlayerVisuals(_activeWaMsgId, false);
      }
      if (audio.paused) {
        _activeWaAudioEl = audio;
        _activeWaMsgId = msgId;
        audio.play().then(() => syncWaPlayerVisuals(msgId, true)).catch(() => {
          toast('Impossible de lire ce message vocal.', 'error');
        });
      } else {
        audio.pause();
        syncWaPlayerVisuals(msgId, false);
      }
    };

    window.cycleWaVoiceSpeed = (msgId, btnEl) => {
      const audio = document.getElementById(`wa-audio-el-${msgId}`);
      if (!audio) return;
      const cur = audio.playbackRate || 1;
      const idx = _waPlaybackRates.findIndex(r => Math.abs(r - cur) < 0.05);
      const next = _waPlaybackRates[(idx + 1) % _waPlaybackRates.length];
      audio.playbackRate = next;
      if (btnEl) btnEl.textContent = `${next}x`;
    };

    window.seekWaVoiceWave = (event, msgId) => {
      const audio = document.getElementById(`wa-audio-el-${msgId}`);
      const container = event.currentTarget;
      if (!audio || !container || !isFinite(audio.duration) || audio.duration <= 0) return;
      const rect = container.getBoundingClientRect();
      const ratio = Math.max(0, Math.min(1, (event.clientX - rect.left) / Math.max(1, rect.width)));
      audio.currentTime = ratio * audio.duration;
      updateWaVoiceProgress(msgId);
    };

    window.updateWaVoiceProgress = (msgId) => {
      const audio = document.getElementById(`wa-audio-el-${msgId}`);
      const barsWrap = document.getElementById(`wa-wave-${msgId}`);
      const timeEl = document.getElementById(`wa-time-${msgId}`);
      if (!audio || !barsWrap) return;
      const dur = isFinite(audio.duration) && audio.duration > 0 ? audio.duration : 0;
      const cur = audio.currentTime || 0;
      const pct = dur > 0 ? (cur / dur) : 0;
      const isMe = barsWrap.dataset.isme === '1';
      const bars = barsWrap.querySelectorAll('.wa-wave-bar');
      const activeCount = Math.round(pct * bars.length);
      bars.forEach((b, i) => {
        if (i < activeCount) {
          b.style.backgroundColor = isMe ? '#ffffff' : '#4f46e5';
        } else {
          b.style.backgroundColor = isMe ? 'rgba(255,255,255,0.38)' : 'rgba(100,116,139,0.38)';
        }
      });
      if (timeEl) {
        timeEl.textContent = cur > 0.05 ? fmtAudioSec(cur) : (timeEl.dataset.defaultDur || fmtAudioSec(dur));
      }
    };

    window.onWaVoiceEnded = (msgId) => {
      syncWaPlayerVisuals(msgId, false);
      const audio = document.getElementById(`wa-audio-el-${msgId}`);
      if (audio) audio.currentTime = 0;
      updateWaVoiceProgress(msgId);
    };

    function renderWhatsAppVoicePlayerHtml(msg, audioUrl, isMe) {
      const mid = esc(msg.id);
      const defaultPeaks = [28, 52, 74, 42, 86, 96, 62, 38, 70, 92, 50, 34, 78, 98, 66, 46, 84, 58, 40, 76, 52, 36, 60, 30];
      const peaks = Array.isArray(msg.wavePeaks) && msg.wavePeaks.length === 24 ? msg.wavePeaks : defaultPeaks;
      const durLabel = esc(msg.audioDuration || '0:05');
      const timeLabel = esc(formatChatHourOnly(msg.at));
      const playBtnCls = isMe
        ? 'bg-white text-indigo-700 hover:bg-indigo-50 shadow-sm'
        : 'bg-indigo-600 text-white hover:bg-indigo-700 shadow-sm';
      const speedPillCls = isMe
        ? 'bg-white/20 hover:bg-white/30 text-white'
        : 'bg-slate-100 hover:bg-slate-200 text-slate-700 dark:bg-slate-700/90 dark:text-slate-200 border border-slate-200/80 dark:border-slate-600';
      const barIdleColor = isMe ? 'rgba(255,255,255,0.40)' : 'rgba(100,116,139,0.38)';

      const barsHtml = peaks.map(h =>
        `<span class="wa-wave-bar" style="height:${Math.max(20, Math.min(100, Number(h || 40)))}%;background-color:${barIdleColor};"></span>`
      ).join('');

      return `
        <div class="wa-player-wrap">
          <audio id="wa-audio-el-${mid}" preload="metadata" src="${esc(audioUrl)}"
            ontimeupdate="updateWaVoiceProgress('${mid}')"
            onloadedmetadata="updateWaVoiceProgress('${mid}')"
            onended="onWaVoiceEnded('${mid}')"></audio>
          <button type="button" onclick="toggleWaVoicePlay('${mid}')" aria-label="Lire ou mettre en pause" class="wa-play-btn ${playBtnCls}">
            <svg id="wa-play-icon-${mid}" class="w-4 h-4 translate-x-0.5" viewBox="0 0 24 24" fill="currentColor"><path d="M8 5v14l11-7z"/></svg>
            <svg id="wa-pause-icon-${mid}" class="hidden w-4 h-4" viewBox="0 0 24 24" fill="currentColor"><path d="M6 19h4V5H6v14zm8-14v14h4V5h-4z"/></svg>
          </button>
          <div class="flex-1 min-w-0 flex flex-col gap-1.5">
            <div id="wa-wave-${mid}" data-isme="${isMe ? '1' : '0'}" onclick="seekWaVoiceWave(event, '${mid}')" class="wa-wave-bars">
              ${barsHtml}
            </div>
            <div class="flex items-center justify-between gap-2 text-[10px] font-bold tabular-nums leading-none">
              <div class="flex items-center gap-1.5">
                <span id="wa-time-${mid}" data-default-dur="${durLabel}" class="opacity-90">${durLabel}</span>
                <button type="button" onclick="cycleWaVoiceSpeed('${mid}', this)" title="Changer la vitesse de lecture" class="wa-speed-pill ${speedPillCls}">1x</button>
              </div>
              <span class="opacity-70 font-semibold">${timeLabel}${isMe ? ' ✓✓' : ''}</span>
            </div>
          </div>
        </div>
      `;
    }

    async function postMessageToTeamChat(text, extraMedia = {}) {
      const cleanText = String(text || '').trim();
      const hasPhoto = Boolean(pendingChatPhotoDataUrl);
      const hasAudio = Boolean(extraMedia.audioRef);
      if (!cleanText && !hasPhoto && !hasAudio) return;
      if (!sb || !currentProfile) return;

      const nowIso = new Date().toISOString();
      const msgId = `msg_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
      const mentionedNames = extractMentionedNamesFromText(cleanText);
      const replyPayload = chatReplyTarget ? { ...chatReplyTarget } : null;

      let imageRef = '';
      if (hasPhoto) {
        const photoData = pendingChatPhotoDataUrl;
        const sendBtn = document.getElementById('btn-send-chat');
        if (sendBtn) sendBtn.disabled = true;
        toast('📷 Envoi de la photo…');
        try {
          imageRef = await uploadChatMediaBlob(dataUrlToBlob(photoData), 'jpg', 'image/jpeg');
        } finally {
          if (sendBtn) sendBtn.disabled = false;
        }
        if (!imageRef) {
          toast('Photo non envoyée : ' + (window._lastChatMediaError || 'réseau indisponible'), 'error');
          const inp = document.getElementById('chat-message-input');
          if (inp && cleanText && !inp.value) { inp.value = cleanText; autoResizeChatTextarea(); }
          return; // la photo reste en aperçu pour réessayer
        }
        clearPendingChatPhoto();
      }
      clearChatReplyTarget();

      const chatObj = {
        id: msgId,
        kind: 'msg',
        at: nowIso,
        authorId: String(currentProfile.id),
        authorName: currentProfile.nom,
        authorRole: currentProfile.role || 'representant',
        text: cleanText,
        mentions: mentionedNames,
        ...(replyPayload ? { replyTo: replyPayload } : {}),
        ...(imageRef ? { imageRef } : {}),
        ...(hasAudio ? { audioRef: extraMedia.audioRef, audioDuration: extraMedia.audioDuration || '', ...(Array.isArray(extraMedia.wavePeaks) ? { wavePeaks: extraMedia.wavePeaks } : {}) } : {})
      };

      kpiHebdoRows.push({
        mois_commission: msgId,
        semaine_index: -12,
        representant_nom: JSON.stringify(chatObj),
        booke: 0,
        rdv_realise: 0,
        updated_at: nowIso
      });
      buildTeamChatMessagesFromState();
      await resolveChatMediaSignedUrls();
      renderTeamChatFeed(true);

      const res = await safeUpsertKpiHebdoRow({
        mois_commission: msgId,
        semaine_index: -12,
        representant_nom: JSON.stringify(chatObj),
        booke: 0,
        rdv_realise: 0,
        updated_at: nowIso
      });
      if (res && res.error) {
        toast('Erreur lors de l’envoi du message : ' + res.error.message, 'error');
      } else if (mentionedNames && mentionedNames.length > 0) {
        triggerServerMentionWebPush(chatObj, mentionedNames);
      }
    }

    window.deleteTeamChatMessage = async (msgId) => {
      if (!sb || !currentProfile) return;
      const target = teamChatMessages.find(m => String(m.id) === String(msgId));
      if (!target) return;
      const canDelete = currentProfile.role === 'admin' || String(target.authorId) === String(currentProfile.id);
      if (!canDelete) return;

      kpiHebdoRows = kpiHebdoRows.filter(r => !(Number(r.semaine_index) === -12 && String(r.mois_commission) === String(msgId)));
      buildTeamChatMessagesFromState();
      renderTeamChatFeed(false);

      try {
        const filesToRemove = [target.imageRef, target.audioRef]
          .map(r => String(r || ''))
          .filter(r => r.includes('/chat/') && isProofPath(r));
        if (filesToRemove.length) await sb.storage.from(PROOF_BUCKET).remove(filesToRemove);
      } catch (e) {}

      try {
        await sb.from('kpi_hebdo').delete().eq('mois_commission', String(msgId)).eq('semaine_index', -12);
      } catch (e) {}
    };

    window.sendQuickBravoInChat = async (repNom, entNom, montant) => {
      await postMessageToTeamChat(`🔥 Félicitations @${repNom} pour la vente ${entNom} (${fmtMoney0(montant)}) ! 👏`);
    };

    function formatChatDayDividerLabel(isoStr) {
      if (!isoStr) return '';
      try {
        const d = new Date(isoStr);
        const now = new Date();
        const todayStr = localYmd(now);
        const yest = new Date(now);
        yest.setDate(now.getDate() - 1);
        const yestStr = localYmd(yest);
        const msgDayStr = localYmd(d);
        if (msgDayStr === todayStr) return "Aujourd'hui";
        if (msgDayStr === yestStr) return 'Hier';
        const raw = d.toLocaleDateString('fr-CA', { weekday: 'long', day: 'numeric', month: 'long' });
        return raw.charAt(0).toUpperCase() + raw.slice(1);
      } catch (e) {
        return '';
      }
    }

    function formatChatHourOnly(isoStr) {
      if (!isoStr) return '';
      try {
        return new Date(isoStr).toLocaleTimeString('fr-CA', { hour: '2-digit', minute: '2-digit' });
      } catch (e) {
        return '';
      }
    }

    window.scrollTeamChatToBottom = () => {
      const feed = document.getElementById('chat-messages-feed');
      if (feed) {
        feed.scrollTo({ top: feed.scrollHeight, behavior: 'smooth' });
      }
    };

    window.scrollToQuotedChatMessage = (targetId) => {
      const el = document.getElementById(`chat-item-${String(targetId)}`);
      if (!el) return;
      el.scrollIntoView({ behavior: 'smooth', block: 'center' });
      el.classList.add('ring-2', 'ring-indigo-400', 'rounded-2xl', 'transition-all');
      setTimeout(() => el.classList.remove('ring-2', 'ring-indigo-400'), 1800);
    };

    // Glisser vers la droite sur un message (mobile/tactile) pour y répondre directement
    let _chatSwipeStartX = 0;
    let _chatSwipeStartY = 0;
    let _chatSwipeEl = null;

    window.handleChatTouchStart = (e, el) => {
      const t = e.touches && e.touches[0];
      if (!t) return;
      _chatSwipeStartX = t.clientX;
      _chatSwipeStartY = t.clientY;
      _chatSwipeEl = el;
    };

    window.handleChatTouchMove = (e) => {
      if (!_chatSwipeEl) return;
      const t = e.touches && e.touches[0];
      if (!t) return;
      const dx = t.clientX - _chatSwipeStartX;
      const dy = Math.abs(t.clientY - _chatSwipeStartY);
      if (dy > 28) {
        _chatSwipeEl.style.transform = '';
        _chatSwipeEl = null;
        return;
      }
      if (dx > 6 && dx < 72) {
        _chatSwipeEl.style.transform = `translateX(${dx}px)`;
      }
    };

    window.handleChatTouchEnd = (e, msgId) => {
      if (!_chatSwipeEl) return;
      const t = e.changedTouches && e.changedTouches[0];
      const dx = t ? (t.clientX - _chatSwipeStartX) : 0;
      _chatSwipeEl.style.transition = 'transform 0.18s ease';
      _chatSwipeEl.style.transform = '';
      setTimeout(() => { if (_chatSwipeEl) _chatSwipeEl.style.transition = ''; }, 200);
      _chatSwipeEl = null;
      if (dx >= 46) {
        setChatReplyTarget(msgId);
      }
    };

    const CHAT_SENDER_COLORS = [
      'text-indigo-700 dark:text-indigo-300',
      'text-emerald-700 dark:text-emerald-300',
      'text-sky-700 dark:text-sky-300',
      'text-amber-700 dark:text-amber-300',
      'text-purple-700 dark:text-purple-300',
      'text-rose-700 dark:text-rose-300'
    ];

    function getWaSenderColorClass(name) {
      const str = String(name || '');
      let hash = 0;
      for (let i = 0; i < str.length; i++) {
        hash = ((hash << 5) - hash) + str.charCodeAt(i);
        hash |= 0;
      }
      return CHAT_SENDER_COLORS[Math.abs(hash) % CHAT_SENDER_COLORS.length];
    }

    function renderTeamChatFeed(scrollToBottom = false) {
      const feed = document.getElementById('chat-messages-feed');
      if (!feed) return;
      const isAdminUser = currentProfile?.role === 'admin';

      const subEl = document.getElementById('wa-chat-members-subtitle');
      if (subEl) {
        const names = getMentionableTeamMembers().map(m => m.nom.split(' ')[0]);
        subEl.textContent = names.length ? names.join(', ') : 'Équipe Korpo • En direct';
      }

      const list = chatSalesOnlyMode
        ? teamChatMessages.filter(m => m.kind === 'sale')
        : teamChatMessages;

      if (!list.length) {
        feed.innerHTML = `
          <div class="py-16 text-center space-y-2">
            <div class="inline-block px-4 py-2 rounded-2xl bg-white dark:bg-slate-800 text-xs font-semibold text-slate-500 dark:text-slate-300 border border-slate-200/80 dark:border-slate-700">
              Aucun message pour le moment. Écrivez ci-dessous ou envoyez un vocal !
            </div>
          </div>
        `;
        return;
      }

      const prevSeenCutoff = window._chatOpenedPrevSeenTs || '';
      let insertedUnreadDivider = false;
      let lastDayKey = '';

      const olderBannerHtml = (!window._chatShowAllHistory && Number(window._chatHiddenOlderCount || 0) > 0)
        ? `<div class="text-center py-2"><button type="button" onclick="window._chatShowAllHistory = true; buildTeamChatMessagesFromState(); renderTeamChatFeed(false);" class="px-3.5 py-1.5 text-[11px] font-bold rounded-full bg-white dark:bg-slate-800 text-indigo-700 dark:text-indigo-300 border border-indigo-200 dark:border-slate-700 shadow-2xs hover:bg-indigo-50 transition cursor-pointer">⬆️ Charger les ${window._chatHiddenOlderCount} message(s) plus ancien(s)</button></div>`
        : '';

      feed.innerHTML = olderBannerHtml + list.map((m, idx) => {
        const isMe = String(m.authorId) === String(currentProfile?.id);
        const canDel = isAdminUser || (m.kind === 'msg' && isMe);
        const isMentionedMe = isCurrentUserMentionedInMsg(m);

        let dayDividerHtml = '';
        const dayKey = m.at ? localYmd(new Date(m.at)) : '';
        if (dayKey && dayKey !== lastDayKey) {
          lastDayKey = dayKey;
          dayDividerHtml = `
            <div class="flex items-center justify-center my-2.5">
              <span class="px-3 py-0.5 text-[10.5px] font-bold rounded-full bg-slate-200/80 dark:bg-slate-800 text-slate-600 dark:text-slate-300">
                ${esc(formatChatDayDividerLabel(m.at))}
              </span>
            </div>
          `;
        }

        let unreadDividerHtml = '';
        if (!insertedUnreadDivider && prevSeenCutoff && !isMe && String(m.at || '') > prevSeenCutoff) {
          insertedUnreadDivider = true;
          unreadDividerHtml = `
            <div id="chat-unread-divider" class="flex items-center justify-center my-2.5">
              <span class="px-3 py-0.5 text-[10.5px] font-extrabold rounded-full bg-indigo-600 text-white shadow-2xs">
                Nouveaux messages
              </span>
            </div>
          `;
        }

        // CARTE DE NOUVELLE VENTE COMPACTE (MÊME TAILLE QU'UN MESSAGE NORMAL)
        if (m.kind === 'sale') {
          const sd = m.saleData || {};
          const localSaleMatch = (isAdminUser || isMe)
            ? ((state.mesVentes || []).find(v => String(v.id) === String(sd.saleId)) || (state.ventes || []).find(v => String(v.id) === String(sd.saleId)))
            : null;
          const resolvedLocalCli = localSaleMatch
            ? (localSaleMatch.client_nom || state.clients.find(c => String(c.id) === String(localSaleMatch.client_id))?.nom || '')
            : '';
          const displayCli = (isAdminUser || isMe)
            ? ((resolvedLocalCli && resolvedLocalCli !== 'Client confidentiel') ? resolvedLocalCli : (sd.clientNom && sd.clientNom !== 'Client confidentiel' ? sd.clientNom : ''))
            : '';
          const showClient = Boolean(displayCli);
          return dayDividerHtml + unreadDividerHtml + `
            <div id="chat-item-${esc(m.id)}" ontouchstart="handleChatTouchStart(event, this)" ontouchmove="handleChatTouchMove(event)" ontouchend="handleChatTouchEnd(event, '${esc(m.id)}')" class="my-1.5 flex justify-center">
              <div class="chat-sale-pill max-w-[94%] sm:max-w-md px-3 py-1.5 flex items-center gap-2 text-xs shadow-2xs">
                <span class="text-sm shrink-0">💵</span>
                <div class="min-w-0 flex-1 truncate text-slate-700 dark:text-slate-200">
                  <strong class="font-extrabold text-slate-900 dark:text-white">${esc(sd.repNom || m.authorName)}</strong>
                  <span class="text-slate-400 mx-1">•</span>
                  <span class="font-bold text-indigo-700 dark:text-indigo-300">${esc(sd.entNom || '')}</span>
                  ${showClient ? `<span class="text-slate-500 dark:text-slate-400"> (${esc(displayCli)})</span>` : ''}
                  <span class="text-slate-400 mx-1">•</span>
                  <strong class="font-extrabold text-emerald-700 dark:text-emerald-300 tabular-nums">${fmtMoney(sd.valeur)}</strong>
                </div>
                <div class="flex items-center gap-1.5 shrink-0">
                  <button type="button" onclick="setChatReplyTarget('${esc(m.id)}')" class="text-[10.5px] font-bold text-indigo-600 dark:text-indigo-300 hover:underline cursor-pointer">Répondre</button>
                  ${!isMe ? `<button type="button" onclick="sendQuickBravoInChat('${esc(sd.repNom || m.authorName)}', '${esc(sd.entNom || '')}', ${Number(sd.valeur || 0)})" class="px-2 py-0.5 text-[10px] font-extrabold rounded-md bg-amber-100 hover:bg-amber-200 text-amber-900 transition cursor-pointer">👏 Bravo</button>` : ''}
                  <span class="text-[10px] text-slate-400 tabular-nums">${esc(formatChatHourOnly(m.at))}</span>
                </div>
              </div>
            </div>
          `;
        }

        const prevMsg = idx > 0 ? list[idx - 1] : null;
        const prevDayKey = prevMsg?.at ? localYmd(new Date(prevMsg.at)) : '';
        const isGroupedWithPrev = prevMsg && prevMsg.kind === 'msg' && String(prevMsg.authorId) === String(m.authorId) && prevDayKey === dayKey && !unreadDividerHtml && !dayDividerHtml;

        const isAuthorAdmin = m.authorRole === 'admin';
        const senderColor = getWaSenderColorClass(m.authorName);
        const formattedBody = m.text ? formatChatMessageWithMentions(m.text, isMe) : '';
        const imgUrl = getChatMediaUrl(m.imageRef);
        const audioUrl = getChatMediaUrl(m.audioRef);

        const quotedHtml = m.replyTo ? `
          <div onclick="scrollToQuotedChatMessage('${esc(m.replyTo.id)}')" class="mb-1.5 px-2.5 py-1.5 rounded-xl text-[11px] cursor-pointer border-l-3 ${isMe ? 'bg-white/15 border-l-white text-indigo-100' : 'bg-slate-100 dark:bg-slate-800/90 border-l-indigo-500 text-slate-600 dark:text-slate-300'}">
            <div class="font-extrabold truncate">${esc(m.replyTo.authorName || '')}</div>
            <div class="truncate opacity-90">${esc(m.replyTo.snippet || '')}</div>
          </div>
        ` : '';

        const photoHtml = imgUrl ? `
          <div class="${m.text ? 'mb-1.5' : 'mb-0.5'}">
            <img src="${esc(imgUrl)}" alt="Photo" onclick="openChatImageLightbox('${esc(imgUrl)}')" class="max-h-56 w-auto rounded-xl object-cover cursor-pointer hover:opacity-95 transition" />
          </div>
        ` : '';

        const audioHtml = audioUrl
          ? renderWhatsAppVoicePlayerHtml(m, audioUrl, isMe)
          : chatMediaPendingHtml(m.audioRef, 'Vocal');

        const bubbleClass = isMe
          ? 'chat-bubble-out'
          : `chat-bubble-in ${isMentionedMe ? 'ring-2 ring-amber-400' : ''}`;

        return dayDividerHtml + unreadDividerHtml + `
          <div id="chat-item-${esc(m.id)}" ondblclick="setChatReplyTarget('${esc(m.id)}')" ontouchstart="handleChatTouchStart(event, this)" ontouchmove="handleChatTouchMove(event)" ontouchend="handleChatTouchEnd(event, '${esc(m.id)}')" class="group flex flex-col ${isMe ? 'items-end' : 'items-start'} ${isGroupedWithPrev ? 'mt-0.5' : 'mt-2'}">
            <div class="max-w-[86%] sm:max-w-[76%] px-3.5 py-2.5 text-[13.5px] leading-snug break-words text-center flex flex-col items-center ${bubbleClass}">
              ${!isMe && !isGroupedWithPrev ? `
                <div class="flex items-center justify-center gap-1.5 mb-1 w-full">
                  <span class="text-[11px] font-extrabold ${senderColor}">${esc(m.authorName)}</span>
                  ${isAuthorAdmin ? '<span class="text-[9.5px] font-bold text-slate-400">Admin</span>' : ''}
                  ${isMentionedMe ? '<span class="px-1.5 py-0.2 text-[9px] font-extrabold rounded bg-amber-200 text-amber-950">@Vous</span>' : ''}
                </div>
              ` : ''}
              ${quotedHtml}
              ${photoHtml || chatMediaPendingHtml(m.imageRef, 'Photo')}
              ${audioHtml}
              ${formattedBody ? `<div class="whitespace-pre-wrap text-center w-full">${formattedBody}</div>` : ''}
              ${!(audioUrl && !formattedBody && !imgUrl) ? `
                <div class="flex items-center justify-center gap-1 mt-1 text-[9.5px] opacity-70 tabular-nums select-none w-full">
                  <span>${esc(formatChatHourOnly(m.at))}</span>
                  ${isMe ? '<span>✓✓</span>' : ''}
                </div>
              ` : ''}
            </div>
            <div class="flex items-center justify-center gap-2.5 mt-0.5 px-1.5 text-[10.5px] text-slate-400">
              <button type="button" onclick="setChatReplyTarget('${esc(m.id)}')" class="hover:text-indigo-600 dark:hover:text-indigo-300 font-semibold transition cursor-pointer">Répondre</button>
              ${canDel ? `<button type="button" onclick="deleteTeamChatMessage('${esc(m.id)}')" class="hover:text-rose-600 transition cursor-pointer">Supprimer</button>` : ''}
            </div>
          </div>
        `;
      }).join('');

      if (scrollToBottom) {
        const unreadDiv = document.getElementById('chat-unread-divider');
        if (unreadDiv) {
          unreadDiv.scrollIntoView({ behavior: 'smooth', block: 'center' });
        } else {
          feed.scrollTop = feed.scrollHeight;
        }
      }
    }

    // Autocomplétion @Nom dans la zone de saisie du chat
    function hideMentionAutocomplete() {
      const box = document.getElementById('chat-mention-autocomplete');
      if (box) box.classList.add('hidden');
      mentionAcCandidates = [];
      mentionAcSelectedIndex = 0;
    }

    function renderMentionAutocompleteList() {
      const box = document.getElementById('chat-mention-autocomplete');
      const listEl = document.getElementById('chat-mention-autocomplete-list');
      if (!box || !listEl) return;
      if (!mentionAcCandidates.length) {
        hideMentionAutocomplete();
        return;
      }
      listEl.innerHTML = mentionAcCandidates.map((m, idx) => `
        <button type="button" onmousedown="event.preventDefault(); applyMentionAutocomplete(${idx});" class="w-full px-3.5 py-2 text-left text-xs flex items-center justify-between gap-2 transition ${idx === mentionAcSelectedIndex ? 'bg-indigo-50 text-indigo-900 font-extrabold' : 'hover:bg-slate-50 text-slate-700 font-semibold'}">
          <span class="flex items-center gap-2">
            <span class="w-6 h-6 rounded-full bg-indigo-600 text-white text-[10px] font-extrabold flex items-center justify-center">${esc(m.nom.slice(0, 2).toUpperCase())}</span>
            <span>@${esc(m.nom)}</span>
          </span>
          <span class="text-[10px] text-slate-400">${m.role === 'admin' ? 'Admin' : 'Rep'}</span>
        </button>
      `).join('');
      box.classList.remove('hidden');
    }

    window.applyMentionAutocomplete = (idx) => {
      const input = document.getElementById('chat-message-input');
      const chosen = mentionAcCandidates[idx];
      if (!input || !chosen) return;
      const curPos = input.selectionStart ?? input.value.length;
      const before = input.value.slice(0, curPos);
      const after = input.value.slice(curPos);
      const replacedBefore = before.replace(/@([A-Za-zÀ-ÖØ-öø-ÿ0-9_-]*)$/, `@${chosen.nom} `);
      input.value = replacedBefore + after;
      const newPos = replacedBefore.length;
      input.focus();
      input.setSelectionRange(newPos, newPos);
      autoResizeChatTextarea();
      hideMentionAutocomplete();
    };

    window.insertAtSymbolInChat = () => {
      const input = document.getElementById('chat-message-input');
      if (!input) return;
      const needsSpace = input.value.length > 0 && !/\s$/.test(input.value);
      input.value += (needsSpace ? ' @' : '@');
      input.focus();
      autoResizeChatTextarea();
      input.dispatchEvent(new Event('input'));
    };

    const chatInputEl = document.getElementById('chat-message-input');
    if (chatInputEl) {
      chatInputEl.addEventListener('input', () => {
        autoResizeChatTextarea();
        const curPos = chatInputEl.selectionStart ?? chatInputEl.value.length;
        const before = chatInputEl.value.slice(0, curPos);
        const match = before.match(/(?:^|\s)@([A-Za-zÀ-ÖØ-öø-ÿ0-9_-]*)$/);
        if (!match) {
          hideMentionAutocomplete();
          return;
        }
        const queryNorm = normalizeMentionStr(match[1]);
        const members = getMentionableTeamMembers();
        mentionAcCandidates = members.filter(m =>
          !queryNorm || normalizeMentionStr(m.nom).includes(queryNorm)
        );
        mentionAcSelectedIndex = 0;
        renderMentionAutocompleteList();
      });

      chatInputEl.addEventListener('keydown', (e) => {
        if (mentionAcCandidates.length) {
          if (e.key === 'ArrowDown') {
            e.preventDefault();
            mentionAcSelectedIndex = (mentionAcSelectedIndex + 1) % mentionAcCandidates.length;
            renderMentionAutocompleteList();
            return;
          } else if (e.key === 'ArrowUp') {
            e.preventDefault();
            mentionAcSelectedIndex = (mentionAcSelectedIndex - 1 + mentionAcCandidates.length) % mentionAcCandidates.length;
            renderMentionAutocompleteList();
            return;
          } else if (e.key === 'Enter' || e.key === 'Tab') {
            e.preventDefault();
            applyMentionAutocomplete(mentionAcSelectedIndex);
            return;
          } else if (e.key === 'Escape') {
            e.stopPropagation();
            hideMentionAutocomplete();
            return;
          }
        }
        // Envoyer avec Entrée (Shift+Entrée pour sauter une ligne)
        if (e.key === 'Enter' && !e.shiftKey) {
          e.preventDefault();
          document.getElementById('form-team-chat')?.dispatchEvent(new Event('submit', { cancelable: true, bubbles: true }));
        }
      });

      chatInputEl.addEventListener('blur', () => {
        setTimeout(hideMentionAutocomplete, 150);
      });
    }

    // Bouton cloche @mentions dans le header
    const mentionBellBtn = document.getElementById('btn-chat-mention-notif');
    const mentionDropdown = document.getElementById('chat-mention-notif-dropdown');
    if (mentionBellBtn && mentionDropdown) {
      mentionBellBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        document.getElementById('admin-notif-dropdown')?.classList.add('hidden');
        document.getElementById('user-menu-dropdown')?.classList.add('hidden');
        mentionDropdown.classList.toggle('hidden');
      });
      document.addEventListener('click', (e) => {
        if (!mentionDropdown.contains(e.target) && !mentionBellBtn.contains(e.target)) {
          mentionDropdown.classList.add('hidden');
        }
      });
    }

    let refreshingChatNow = false;
    window.refreshTeamChatMessages = async (forceScroll = false) => {
      if (!sb || !currentProfile || refreshingChatNow) return;
      // Ne jamais reconstruire le DOM pendant qu'un message vocal est en cours d'écoute
      if (_activeWaAudioEl && !_activeWaAudioEl.paused) return;
      refreshingChatNow = true;
      try {
        const prevSig = chatFeedSignature();
        const rows = await fetchAllRows(() =>
          sb.from('kpi_hebdo').select('*').in('semaine_index', [-12, -14, -15, -16]).order('id')
        );
        const chatRowsOnly = rows.filter(r => Number(r.semaine_index) === -12);
        const fetchedIds = new Set(chatRowsOnly.map(r => String(r.mois_commission)));
        const pendingRows = Array.from(_pendingVoiceUploads.values()).filter(r => !fetchedIds.has(String(r.mois_commission)));
        kpiHebdoRows = [
          ...(kpiHebdoRows || []).filter(r => ![-12, -14, -15, -16].includes(Number(r.semaine_index))),
          ...rows,
          ...pendingRows
        ];
        rows.forEach(r => {
          if (Number(r.semaine_index) === -16 && r.mois_commission && r.representant_nom) {
            try {
              const parsedDaily = JSON.parse(r.representant_nom);
              const dKey = String(r.mois_commission);
              if (parsedDaily && typeof parsedDaily === 'object') {
                const locD = dailyRdvLogsStore[dKey];
                if (!locD || !locD.updatedAt || (parsedDaily.updatedAt && parsedDaily.updatedAt >= locD.updatedAt)) {
                  dailyRdvLogsStore[dKey] = parsedDaily;
                  saveDailyRdvLogsStore();
                  renderDailyRdvTrackerCard();
                }
              }
            } catch (e) {}
          }
          if (Number(r.semaine_index) === -14 && r.mois_commission === 'red_phone_cfg' && r.representant_nom) {
            try {
              const parsedRp = JSON.parse(r.representant_nom);
              if (parsedRp && typeof parsedRp === 'object') {
                redPhoneConfig = { ...redPhoneConfig, ...parsedRp };
                lsSet('korpo_red_phone_cfg_v1', JSON.stringify(redPhoneConfig));
                syncRedPhoneUI();
              }
            } catch (e) {}
          }
        });
        checkIncomingRedPhoneAlerts();
        buildTeamChatMessagesFromState();
        await resolveChatMediaSignedUrls();
        const nextSig = chatFeedSignature();
        // Ne redessiner le fil que si de nouveaux messages/médias sont arrivés (et qu'aucun audio ne joue)
        if (nextSig !== prevSig && !(_activeWaAudioEl && !_activeWaAudioEl.paused)) {
          renderTeamChatFeed(forceScroll);
        }
      } catch (e) {
        console.warn('Actualisation canal:', e.message);
      } finally {
        refreshingChatNow = false;
      }
    };

    function subscribeToTeamChatRealtime() {
      if (!sb || !currentProfile) return;
      buildTeamChatMessagesFromState();
      resolveChatMediaSignedUrls().then(() => renderTeamChatFeed(true));

      if (!chatRealtimeChannel) {
        chatRealtimeChannel = sb
          .channel('korpo-team-chat-' + Date.now())
          .on('postgres_changes', { event: '*', schema: 'public', table: 'kpi_hebdo' }, async (payload) => {
            const sIdx = Number(payload?.new?.semaine_index ?? payload?.old?.semaine_index);
            if (sIdx === -12 || sIdx === -14 || sIdx === -15 || sIdx === -16) {
              await refreshTeamChatMessages(activeChatTabOpen);
            }
          })
          .on('postgres_changes', { event: 'INSERT', schema: 'public', table: 'ventes' }, async () => {
            await refreshTeamChatMessages(activeChatTabOpen);
          })
          .subscribe((status) => {
            if (status === 'CHANNEL_ERROR' || status === 'TIMED_OUT') {
              chatRealtimeChannel = null;
            }
          });
      }

      if (!chatPollingTimer) {
        chatPollingTimer = setInterval(() => {
          if (!chatRealtimeChannel) subscribeToTeamChatRealtime();
          refreshTeamChatMessages(activeChatTabOpen);
        }, 4000);
      }
      if (!window._chatAutoFocusBound) {
        window._chatAutoFocusBound = true;
        document.addEventListener('visibilitychange', () => {
          if (document.visibilityState === 'visible') {
            refreshTeamChatMessages(activeChatTabOpen);
          }
        });
        window.addEventListener('focus', () => {
          refreshTeamChatMessages(activeChatTabOpen);
        });
      }
    }

    const KORPO_CHAT_EMOJIS = [
      '🔥', '💰', '💵', '🏆', '👏', '🎉', '🚀', '💪',
      '👑', '⭐', '💯', '🤝', '📈', '💎', '⚡', '🎯',
      '😀', '😂', '🥳', '😎', '🤩', '🙌', '👍', '❤️',
      '✅', '📞', '📅', '🏠', '❄️', '☀️', '👀', '🙏'
    ];

    function initChatEmojiPicker() {
      const grid = document.getElementById('chat-emoji-grid');
      if (!grid || grid.dataset.ready) return;
      grid.dataset.ready = '1';
      grid.innerHTML = KORPO_CHAT_EMOJIS.map(em => `
        <button type="button" onclick="insertEmojiInChat('${em}')" class="w-8 h-8 rounded-lg hover:bg-slate-100 active:scale-90 flex items-center justify-center transition">
          ${em}
        </button>
      `).join('');
    }

    window.toggleChatEmojiPicker = (forceState) => {
      const picker = document.getElementById('chat-emoji-picker');
      if (!picker) return;
      initChatEmojiPicker();
      hideMentionAutocomplete();
      const show = typeof forceState === 'boolean' ? forceState : picker.classList.contains('hidden');
      picker.classList.toggle('hidden', !show);
    };

    window.insertEmojiInChat = (emoji) => {
      const input = document.getElementById('chat-message-input');
      if (!input) return;
      const start = input.selectionStart ?? input.value.length;
      const end = input.selectionEnd ?? input.value.length;
      input.value = input.value.slice(0, start) + emoji + input.value.slice(end);
      const nextPos = start + emoji.length;
      input.focus();
      input.setSelectionRange(nextPos, nextPos);
      autoResizeChatTextarea();
    };

    document.addEventListener('click', (e) => {
      const picker = document.getElementById('chat-emoji-picker');
      const btn = document.getElementById('btn-toggle-chat-emoji');
      if (picker && !picker.classList.contains('hidden') && !picker.contains(e.target) && (!btn || !btn.contains(e.target))) {
        picker.classList.add('hidden');
      }
    });

    document.getElementById('form-team-chat')?.addEventListener('submit', async (e) => {
      e.preventDefault();
      hideMentionAutocomplete();
      toggleChatEmojiPicker(false);
      const input = document.getElementById('chat-message-input');
      if (!input) return;
      const val = input.value.trim();
      if (!val && !pendingChatPhotoDataUrl) return;
      input.value = '';
      autoResizeChatTextarea();
      await postMessageToTeamChat(val);
    });

    // CHARGEMENT SUPABASE (VENTES & COMMISSIONS)
    // =========================================================================
    async function loadAllDataInner() {
      const isAdmin = currentProfile.role === 'admin';
      clearBulkSelection();

      const [classementRows, entRows, clientRows, kpiRows] = await Promise.all([
        fetchAllRows(() => sb.from('vue_classement_ventes').select('*')),
        fetchAllRows(() => sb.from('entreprises').select('*').order('nom').order('id')),
        fetchAllRows(() => sb.from('clients').select('*').order('nom').order('id')),
        // Les représentants ne téléchargent pas la corbeille (index -3) des autres
        fetchAllRows(() => {
          let q = sb.from('kpi_hebdo').select('*').order('id');
          if (!isAdmin) q = q.neq('semaine_index', -3).neq('semaine_index', -14);
          q = q.neq('semaine_index', -13); // anciens blocs média du chat : lus à la demande
          return q;
        })
      ]);

      state.classement = classementRows;
      state.entreprises = entRows;
      state.clients = clientRows;
      kpiHebdoRows = kpiRows;

      if (isAdmin) {
        const [repRows, tauxRows, venteRows] = await Promise.all([
          fetchAllRows(() => sb.from('representants').select('*').order('nom').order('id')),
          fetchAllRows(() => sb.from('taux_commission').select('*').order('id')),
          fetchAllRows(() => sb.from('ventes').select('*').order('date_vente', { ascending: false }).order('created_at', { ascending: false }).order('id', { ascending: false }))
        ]);
        state.representants = repRows;
        syncAllKpiRepsFromState();
        state.taux = tauxRows;
        state.ventes = venteRows;

        state.mesVentes = state.ventes.map(mapAdminPersonalSaleRow);

        document.getElementById('filter-admin-scope').classList.remove('hidden');
        document.getElementById('th-rep-col').classList.remove('hidden');
        const thGrpVentes = document.getElementById('th-group-ventes');
        if (thGrpVentes) thGrpVentes.colSpan = 5;
        document.getElementById('mes-ventes-subtitle').textContent = 'Vos ventes directes et vos commissions sur l’ensemble de l’équipe';

        mergeCloudAuditAndTrash();
        const autoCleanedDups = await reconcilePreviousUnconfirmedDuplicateSales();
        if (autoCleanedDups > 0) {
          state.classement = await fetchAllRows(() => sb.from('vue_classement_ventes').select('*'));
          toast(`🔗 ${autoCleanedDups} doublon(s) antérieur(s) non confirmé(s) fusionné(s) automatiquement`);
        }
        await runMonthlyAutoRolloverAndCarryOver();
        mergeCloudAuditAndTrash();
        syncSalesIntoKpi();
        syncGoogleAgendaIntoKpi();
        populateAdminSelects();
        await refreshProofUrls(state.ventes);
        migrateLegacyProofs(state.ventes);
        renderAdminVentes();
        renderAdminDashboard();
        renderAdminTaux();
        renderAdminEquipe();
        renderAdminNotifications();
        subscribeToNewSalesRealtime();
      } else {
        const [mesVentesRows, myTauxRows, teamVentesPublicRows, allRepsPublicRows] = await Promise.all([
          fetchAllRows(() => sb.from('vue_ventes_representant').select('*').eq('representant_id', currentProfile.id).order('date_vente', { ascending: false }).order('created_at', { ascending: false })),
          fetchAllRows(() => sb.from('taux_commission').select('*').eq('representant_id', currentProfile.id)),
          fetchAllRows(() => sb.from('ventes').select('id,representant_id,entreprise_id,valeur,date_vente,mois_commission,statut_depot,note,created_at').order('date_vente', { ascending: false })).catch(() => []),
          fetchAllRows(() => sb.from('representants').select('id,nom,role').order('nom')).catch(() => [])
        ]);
        const noteByIdMap = new Map((teamVentesPublicRows || []).map(r => [String(r.id), r.note]));
        mesVentesRows.forEach(r => {
          if (r.note === undefined && noteByIdMap.has(String(r.id))) {
            r.note = noteByIdMap.get(String(r.id));
          }
        });
        document.getElementById('filter-admin-scope')?.classList.add('hidden');
        document.getElementById('th-rep-col')?.classList.add('hidden');
        state.representants = (allRepsPublicRows && allRepsPublicRows.length) ? allRepsPublicRows : [currentProfile];
        state.mesVentes = (mesVentesRows || []).map(normalizeRepPersonalSaleRow);
        const myOwnSaleIds = new Set(state.mesVentes.map(v => String(v.id)));
        const myOwnClientIds = new Set(state.mesVentes.map(v => String(v.client_id || '')).filter(Boolean));
        state.clients = (state.clients || []).filter(c => myOwnClientIds.has(String(c.id)));
        kpiHebdoRows = (kpiHebdoRows || []).filter(r => {
          const sIdx = Number(r.semaine_index);
          if (sIdx === -2 || sIdx === -4 || sIdx === -5) {
            return myOwnSaleIds.has(String(r.mois_commission));
          }
          if (sIdx === -14) return false; // Ne pas exposer le numéro personnel du directeur dans les données locales d'un rep
          return true;
        });
        mergeCloudAuditAndTrash();
        // Détecter automatiquement tout ancien membre présent dans l'historique KPI (-11 / -9) mais absent des représentants actifs
        const activeRepNamesLower = new Set((state.representants || []).map(r => String(r?.nom || '').trim().toLowerCase()));
        Object.values(teamKpiSnapshotStore || {}).forEach(snapMonth => {
          Object.keys(snapMonth || {}).forEach(rn => {
            const norm = resolveKpiRepName(rn);
            if (norm && !activeRepNamesLower.has(norm.toLowerCase()) && !isRepArchived(norm)) {
              archivedRepsStore[norm] = { id: `arch_rep_${norm.toLowerCase()}`, nom: norm, archivedAt: '2026-01-01' };
            }
          });
        });
        const hasUnassignedSales = (teamVentesPublicRows || []).some(v => v.representant_id === null || v.representant_id === undefined || v.representant_id === '');
        if (hasUnassignedSales) {
          DEFAULT_KPI_REPS.forEach(rn => {
            if (!activeRepNamesLower.has(rn.toLowerCase()) && !isRepArchived(rn)) {
              archivedRepsStore[rn] = { id: `arch_rep_${rn.toLowerCase()}`, nom: rn, archivedAt: '2026-01-01' };
            }
          });
        }
        syncAllKpiRepsFromState();

        const ownSalesNormalized = state.mesVentes.map(v => ({
          ...v,
          representant_id: v.representant_id || currentProfile.id,
          rep_nom: currentProfile.nom,
          entreprise_id: v.entreprise_id || state.entreprises.find(e => e.nom === v.entreprise_nom)?.id,
          client_id: v.client_id || state.clients.find(c => c.nom === v.client_nom)?.id
        }));
        const ownIds = new Set(ownSalesNormalized.map(v => String(v.id)));
        const otherRepsPublicSales = (teamVentesPublicRows || [])
          .filter(v => String(v.representant_id) !== String(currentProfile.id) && !ownIds.has(String(v.id)) && v.statut_depot === 'depot_recu')
          .map(v => {
            const rObj = state.representants.find(r => String(r.id) === String(v.representant_id));
            const cRow = (state.classement || []).find(c => String(c.representant_id) === String(v.representant_id));
            return {
              ...v,
              rep_nom: rObj?.nom || resolveSaleArchivedRepName(v) || cRow?.representant_nom || '',
              entreprise_nom: state.entreprises.find(e => String(e.id) === String(v.entreprise_id))?.nom || resolveSaleArchivedEntName(v) || '—',
              client_nom: 'Client confidentiel'
            };
          });
        state.ventes = [...ownSalesNormalized, ...otherRepsPublicSales];
        state.taux = myTauxRows;
        if (!state.taux.length) {
          const { data: rpcTaux } = await sb.rpc('get_mes_taux_commission');
          if (rpcTaux && rpcTaux.length) state.taux = rpcTaux;
        }
        mergeCloudAuditAndTrash();
        await runMonthlyAutoRolloverAndCarryOver();
        syncSalesIntoKpi();
        syncGoogleAgendaIntoKpi();
        populateDrawerSelects();
      }

      await refreshClosedMonths();
      await refreshProofUrls([...(state.mesVentes || []), ...(state.ventes || [])]);
      migrateLegacyProofs(state.mesVentes);
      renderMesVentes();
      renderClassement();
      renderKpiDashboard();
      subscribeToTeamChatRealtime();
    }

    async function loadAllData() {
      try {
        await loadAllDataInner();
      } catch (err) {
        console.error('Erreur de chargement des données :', err);
        toast(`Erreur de chargement des données : ${err?.message || err}`, 'error');
        return false;
      }
    }

    let mesVentesSort = { col: 'date_vente', dir: 'desc' };
    let adminVentesSort = { col: 'date_vente', dir: 'desc' };

    function compareSaleValues(valA, valB, dir, isNumeric = false) {
      const mult = dir === 'asc' ? 1 : -1;
      if (isNumeric) {
        return (Number(valA || 0) - Number(valB || 0)) * mult;
      }
      return String(valA || '').localeCompare(String(valB || ''), 'fr', { numeric: true, sensitivity: 'base' }) * mult;
    }

    function updateSortIndicators(prefix, cols, activeSort) {
      cols.forEach(c => {
        const el = document.getElementById(`sort-ind-${prefix}-${c}`);
        if (!el) return;
        if (activeSort.col === c) {
          el.textContent = activeSort.dir === 'asc' ? '↑' : '↓';
          el.className = 'ml-1 text-indigo-600 font-extrabold';
        } else {
          el.textContent = '⇅';
          el.className = 'ml-1 text-slate-300 group-hover:text-indigo-400';
        }
      });
    }

    window.toggleSortMesVentes = (col) => {
      if (mesVentesSort.col === col) {
        mesVentesSort.dir = mesVentesSort.dir === 'asc' ? 'desc' : 'asc';
      } else {
        mesVentesSort = { col, dir: (col === 'date_vente' || col === 'valeur' || col === 'commission_rep') ? 'desc' : 'asc' };
      }
      renderMesVentes();
    };

    window.toggleSortAdminVentes = (col) => {
      if (adminVentesSort.col === col) {
        adminVentesSort.dir = adminVentesSort.dir === 'asc' ? 'desc' : 'asc';
      } else {
        adminVentesSort = { col, dir: (col === 'date_vente' || col === 'valeur' || col === 'commission_rep' || col === 'commission_olivier') ? 'desc' : 'asc' };
      }
      renderAdminVentes();
    };

    let personalMonthlyTargets = lsGetJson('korpo_personal_monthly_targets_v1', '{}');
    let personalCommTargets = lsGetJson('korpo_personal_comm_targets_v1', '{}');
    let personalBonusPctTargets = lsGetJson('korpo_personal_bonus_pct_v1', '{}');

    function getBonusPctForRep(repName, ym) {
      if (personalBonusPctTargets[`${repName}_${ym}`] !== undefined) return Number(personalBonusPctTargets[`${repName}_${ym}`]);
      if (personalBonusPctTargets[`${repName}_ALL`] !== undefined) return Number(personalBonusPctTargets[`${repName}_ALL`]);
      if (repTargetsOverrides[`${ym}_${repName}`]?.bonusPct !== undefined) return Number(repTargetsOverrides[`${ym}_${repName}`].bonusPct);
      if (repTargetsOverrides[`ALL_${repName}`]?.bonusPct !== undefined) return Number(repTargetsOverrides[`ALL_${repName}`].bonusPct);
      return 0.01;
    }

    function formatBonusPctShort(decimalRate) {
      const pctVal = Math.round(Number(decimalRate || 0) * 10000) / 100;
      return Number.isInteger(pctVal) ? `${pctVal}%` : `${pctVal.toFixed(2).replace(/\.?0+$/, '')}%`;
    }

    function getCloudRepTargetRow(repName, ym) {
      if (!Array.isArray(kpiHebdoRows)) return null;
      return kpiHebdoRows.find(r => Number(r.semaine_index) === -1 && r.representant_nom === repName && r.mois_commission === ym)
        || kpiHebdoRows.find(r => Number(r.semaine_index) === -1 && r.representant_nom === repName && r.mois_commission === 'ALL');
    }

    function getSalesTargetForRep(repName, ym) {
      if (isRepArchived(repName)) {
        const factor = getRepArchiveMonthFactor(repName, ym || currentMonthStr);
        if (factor <= 0) return 0;
        const baseD = Number(findArchivedRepEntry(repName)?.targets?.dollars || 0);
        return Math.round(baseD * factor * 100) / 100;
      }
      const cloudRow = getCloudRepTargetRow(repName, ym);
      if (cloudRow && Number(cloudRow.booke) > 0) return Number(cloudRow.booke);
      if (personalMonthlyTargets[`${repName}_${ym}`] > 0) return Number(personalMonthlyTargets[`${repName}_${ym}`]);
      if (personalMonthlyTargets[`${repName}_ALL`] > 0) return Number(personalMonthlyTargets[`${repName}_ALL`]);
      if (repTargetsOverrides[`${ym}_${repName}`]?.dollars > 0) return Number(repTargetsOverrides[`${ym}_${repName}`].dollars);
      if (repTargetsOverrides[`ALL_${repName}`]?.dollars > 0) return Number(repTargetsOverrides[`ALL_${repName}`].dollars);
      const mObj = kpiStore[ym];
      if (mObj?.reps?.[repName]?.['Vente $']?.objectif > 0) return Number(mObj.reps[repName]['Vente $'].objectif);
      return 150000;
    }

    // Vente « sans dépôt requis » : stockée comme dépôt confirmé (aucun report ni relance), repérée par la balise [SANSDEPOT] dans la note.
    function isNoDepositSale(v) {
      return /\[SANSDEPOT\]/i.test(String(v?.note || ''));
    }

    function getSelectedEntreprise() {
      const id = document.getElementById('vente-entreprise')?.value;
      return (state.entreprises || []).find(e => String(e.id) === String(id));
    }
    function companyNoDeposit(ent) {
      const e = ent || getSelectedEntreprise();
      return !!e && e.depot_requis === false;
    }
    // La règle de dépôt vient de la fiche de l'entreprise partenaire.
    function applyCompanyDepositRule() {
      const sel = document.getElementById('vente-statut-depot');
      if (!sel) return;
      const noDep = companyNoDeposit();
      const opt = sel.querySelector('option[value="sans_depot"]');
      if (opt) { opt.disabled = !noDep; opt.hidden = !noDep; }
      if (noDep) {
        sel.value = 'sans_depot';
        sel.disabled = true;
      } else {
        sel.disabled = false;
        if (sel.value === 'sans_depot') sel.value = 'a_confirmer';
      }
      document.getElementById('drawer-depot-company-note')?.classList.toggle('hidden', !noDep);
      syncDepositFormUi();
    }
    document.getElementById('vente-entreprise')?.addEventListener('change', applyCompanyDepositRule);

    function syncDepositFormUi() {
      const isNoDep = document.getElementById('vente-statut-depot')?.value === 'sans_depot';
      document.getElementById('drawer-proof-block')?.classList.toggle('hidden', isNoDep);
      if (isNoDep) {
        pendingProofDataUrl = '';
        const pf = document.getElementById('vente-proof-file'); if (pf) pf.value = '';
        const cm = document.getElementById('vente-camera-input'); if (cm) cm.value = '';
        document.getElementById('vente-proof-preview-box')?.classList.add('hidden');
      }
    }
    document.getElementById('vente-statut-depot')?.addEventListener('change', syncDepositFormUi);

    // Vrai si un objectif a réellement été fixé (cloud, local, surcharge admin) ou préréglé pour l'équipe historique.
    function hasExplicitSalesTarget(repName, ym) {
      const c = getCloudRepTargetRow(repName, ym);
      if (c && Number(c.booke) > 0) return true;
      if (personalMonthlyTargets[`${repName}_${ym}`] > 0 || personalMonthlyTargets[`${repName}_ALL`] > 0) return true;
      if (repTargetsOverrides[`${ym}_${repName}`]?.dollars > 0 || repTargetsOverrides[`ALL_${repName}`]?.dollars > 0) return true;
      return DEFAULT_KPI_REPS.includes(repName) && Number(kpiStore[ym]?.reps?.[repName]?.['Vente $']?.objectif) > 0;
    }

    function hasExplicitCommTarget(repName, ym) {
      const c = getCloudRepTargetRow(repName, ym);
      if (c && Number(c.rdv_realise) > 0) return true;
      if (personalCommTargets[`${repName}_${ym}`] > 0 || personalCommTargets[`${repName}_ALL`] > 0) return true;
      if (repTargetsOverrides[`${ym}_${repName}`]?.comm > 0 || repTargetsOverrides[`ALL_${repName}`]?.comm > 0) return true;
      return hasExplicitSalesTarget(repName, ym); // l'estimation (~5,5 %) repose sur un objectif de ventes réel
    }

    function getCommissionTargetForRep(repName, ym) {
      if (isRepArchived(repName)) {
        const factor = getRepArchiveMonthFactor(repName, ym || currentMonthStr);
        if (factor <= 0) return 0;
        const baseC = Number(findArchivedRepEntry(repName)?.targets?.comm || 0);
        return Math.round(baseC * factor * 100) / 100;
      }
      const cloudRow = getCloudRepTargetRow(repName, ym);
      if (cloudRow && Number(cloudRow.rdv_realise) > 0) return Number(cloudRow.rdv_realise);
      if (personalCommTargets[`${repName}_${ym}`] > 0) return Number(personalCommTargets[`${repName}_${ym}`]);
      if (personalCommTargets[`${repName}_ALL`] > 0) return Number(personalCommTargets[`${repName}_ALL`]);
      if (repTargetsOverrides[`${ym}_${repName}`]?.comm > 0) return Number(repTargetsOverrides[`${ym}_${repName}`].comm);
      if (repTargetsOverrides[`ALL_${repName}`]?.comm > 0) return Number(repTargetsOverrides[`ALL_${repName}`].comm);
      // Estimation par défaut (~5.5% de l'objectif de ventes si non encore personnalisé)
      const salesT = getSalesTargetForRep(repName, ym);
      return Math.max(2500, Math.round((salesT * 0.055) / 500) * 500);
    }

    async function saveRepTargetsToStorageAndCloud(repName, activeYm, salesGoal, commGoal, applyAll = true, bonusPct = null) {
      if (currentProfile?.role !== 'admin') return toast('Seul l’administrateur peut modifier les targets.', 'error');
      const effectiveBonusPct = (bonusPct !== null && !isNaN(bonusPct))
        ? Math.max(0, Number(bonusPct))
        : getBonusPctForRep(repName, activeYm);

      personalMonthlyTargets[`${repName}_${activeYm}`] = salesGoal;
      personalCommTargets[`${repName}_${activeYm}`] = commGoal;
      personalBonusPctTargets[`${repName}_${activeYm}`] = effectiveBonusPct;
      if (applyAll) {
        personalMonthlyTargets[`${repName}_ALL`] = salesGoal;
        personalCommTargets[`${repName}_ALL`] = commGoal;
        personalBonusPctTargets[`${repName}_ALL`] = effectiveBonusPct;
      }
      lsSet('korpo_personal_monthly_targets_v1', JSON.stringify(personalMonthlyTargets));
      lsSet('korpo_personal_comm_targets_v1', JSON.stringify(personalCommTargets));
      lsSet('korpo_personal_bonus_pct_v1', JSON.stringify(personalBonusPctTargets));

      const scopes = applyAll ? [activeYm, 'ALL'] : [activeYm];
      scopes.forEach(sc => {
        const idx = kpiHebdoRows.findIndex(r => Number(r.semaine_index) === -1 && r.representant_nom === repName && r.mois_commission === sc);
        if (idx >= 0) {
          kpiHebdoRows[idx].booke = salesGoal;
          kpiHebdoRows[idx].rdv_realise = commGoal;
        } else {
          kpiHebdoRows.push({ mois_commission: sc, semaine_index: -1, representant_nom: repName, booke: salesGoal, rdv_realise: commGoal });
        }
      });

      if (sb && currentProfile) {
        const rows = scopes.map(sc => ({
          mois_commission: sc,
          semaine_index: -1,
          representant_nom: repName,
          booke: salesGoal,
          rdv_realise: commGoal,
          updated_at: new Date().toISOString()
        }));
        const { error } = await sb.from('kpi_hebdo').upsert(rows, { onConflict: 'mois_commission,semaine_index,representant_nom' });
        if (error) console.warn('Sync targets cloud:', error.message);

        for (const sc of scopes) {
          await safeUpsertKpiHebdoRow({
            mois_commission: `${repName}_${sc}`,
            semaine_index: -7,
            representant_nom: JSON.stringify({ bonusPct: effectiveBonusPct }),
            booke: 0,
            rdv_realise: 0,
            updated_at: new Date().toISOString()
          });
        }
      }
    }

    function ensureMiniBarRocketStyles() {
      if (document.getElementById('korpo-mini-bar-rocket-css')) return;
      const st = document.createElement('style');
      st.id = 'korpo-mini-bar-rocket-css';
      st.textContent = `
        @keyframes korpoMiniRocketBob {
          0%, 100% { transform: translate(-50%, -50%) rotate(45deg) scale(1); }
          50% { transform: translate(-50%, -54%) rotate(48deg) scale(1.1); }
        }
        .korpo-mini-rocket-badge {
          position: absolute;
          top: 50%;
          z-index: 10;
          pointer-events: none;
          user-select: none;
          line-height: 1;
          font-size: 15px;
          filter: drop-shadow(0 1px 4px rgba(15, 23, 42, 0.35));
          transition: left 0.7s cubic-bezier(0.22, 1, 0.36, 1);
          animation: korpoMiniRocketBob 1.3s ease-in-out infinite;
        }
      `;
      document.head.appendChild(st);
    }

    function syncRocketOnProgressBar(fillEl, pctVal) {
      if (!fillEl) return;
      ensureMiniBarRocketStyles();
      const track = fillEl.parentElement;
      if (!track) return;
      track.classList.remove('overflow-hidden');
      track.classList.add('relative', 'overflow-visible');
      let rocketEl = track.querySelector('.korpo-mini-rocket-badge');
      if (!rocketEl) {
        rocketEl = document.createElement('span');
        rocketEl.className = 'korpo-mini-rocket-badge';
        rocketEl.textContent = '🚀';
        track.appendChild(rocketEl);
      }
      const clamped = Math.min(100, Math.max(2, Number(pctVal) || 0));
      rocketEl.style.left = `${clamped}%`;
    }

    function miniRocketProgressTrackHtml(fillClass, pctVal, heightClass = 'h-2.5', trackBgClass = 'bg-slate-100') {
      ensureMiniBarRocketStyles();
      const safePct = Math.min(100, Math.max(0, Number(pctVal) || 0));
      const rocketPos = Math.min(100, Math.max(2, safePct));
      return `
        <div class="relative w-full ${heightClass} ${trackBgClass} rounded-full overflow-visible">
          <div class="h-full ${fillClass} rounded-full transition-all duration-700 ease-out" style="width: ${safePct}%"></div>
          <span class="korpo-mini-rocket-badge" style="left: ${rocketPos}%">🚀</span>
        </div>
      `;
    }

    function updatePersonalGoalProgressBar(periodState, scopeFilter) {
      if (!currentProfile) return;
      const isAdmin = currentProfile.role === 'admin';
      const isInspectingRep = isAdmin && typeof scopeFilter === 'string' && scopeFilter.startsWith('rep:');
      const inspectedRepId = isInspectingRep ? scopeFilter.slice(4) : String(currentProfile.id);
      const inspectedRepObj = isInspectingRep
        ? state.representants.find(r => String(r.id) === inspectedRepId)
        : currentProfile;
      const pMode = periodState?.mode || 'this_month';
      const isMultiMonth = pMode === 'all' || pMode === 'this_year' || pMode === 'this_quarter';
      const activeYm = periodState?.singleMonth
        || (pMode === 'custom_range' && periodState?.startVal ? periodState.startVal.slice(0, 7) : '')
        || currentMonthStr;
      const inspectedRepName = inspectedRepObj?.nom || currentProfile.nom;

      const sourceRows = isInspectingRep
        ? (state.ventes || []).filter(v => String(v.representant_id) === inspectedRepId)
        : state.mesVentes;

      const myPeriodSales = sourceRows.filter(v => {
        if (!isInspectingRep && isAdmin && scopeFilter === 'direct') {
          if (String(v.representant_id) !== String(currentProfile.id)) return false;
        }
        return matchesSalePeriod(v, periodState, true);
      });

      const myOwnDirectSales = myPeriodSales.filter(v => String(v.representant_id) === inspectedRepId);

      let salesTarget = getSalesTargetForRep(inspectedRepName, activeYm);
      let commTarget = getCommissionTargetForRep(inspectedRepName, activeYm);
      const bizDaysInActiveMonth = Math.max(1, countBusinessDaysInMonth(activeYm));

      if (pMode === 'this_week') {
        const weekRatio = 5 / bizDaysInActiveMonth;
        salesTarget = Math.round(salesTarget * weekRatio);
        commTarget = Math.round(commTarget * weekRatio);
      } else if (pMode === 'specific_date') {
        const dayRatio = 1 / bizDaysInActiveMonth;
        salesTarget = Math.round(salesTarget * dayRatio);
        commTarget = Math.round(commTarget * dayRatio);
      } else if (pMode === 'custom_range') {
        const rangeRatio = computeCustomRangeMonthFactor(periodState?.startVal, periodState?.endVal);
        salesTarget = Math.round(salesTarget * rangeRatio);
        commTarget = Math.round(commTarget * rangeRatio);
      } else if (pMode === 'this_quarter') {
        salesTarget = salesTarget * 3;
        commTarget = commTarget * 3;
      } else if (pMode === 'this_year') {
        salesTarget = salesTarget * 12;
        commTarget = commTarget * 12;
      } else if (pMode === 'all') {
        const distinctSalesMonths = new Set(myOwnDirectSales.map(v => getSaleStatsMonth(v)).filter(Boolean));
        const salesMonthsList = distinctSalesMonths.size > 0 ? Array.from(distinctSalesMonths) : [currentMonthStr];
        salesTarget = salesMonthsList.reduce((sum, ym) => sum + getSalesTargetForRep(inspectedRepName, ym), 0);

        const distinctCommMonths = new Set(myPeriodSales.map(v => getSaleStatsMonth(v)).filter(Boolean));
        const commMonthsList = distinctCommMonths.size > 0 ? Array.from(distinctCommMonths) : [currentMonthStr];
        commTarget = commMonthsList.reduce((sum, ym) => sum + getCommissionTargetForRep(inspectedRepName, ym), 0);
      }

      const achievedSales = myOwnDirectSales.reduce((s, v) => s + Number(v.valeur || 0), 0);
      const achievedComm = myPeriodSales.reduce((s, v) => s + Number(v.commission_rep || 0), 0);

      const headerTitleEl = document.getElementById('goal-card-header-title');
      if (headerTitleEl) {
        const periodTag = periodState?.label || activeYm;
        headerTitleEl.textContent = isInspectingRep
          ? `Objectifs de ${inspectedRepName} — ${periodTag}`
          : (isMultiMonth ? `Mes objectifs — ${periodTag}` : `Mes objectifs — ${periodTag}`);
      }

      const salesExplicit = hasExplicitSalesTarget(inspectedRepName, activeYm);
      const commExplicit = hasExplicitCommTarget(inspectedRepName, activeYm);
      const undefinedGoalMsg = isAdmin ? 'À configurer dans Entreprises & Équipe' : 'Votre direction le fixera bientôt';

      // 1. Mise à jour de la barre TARGET COMMISSION ($)
      const commPct = commTarget > 0 ? Math.round((achievedComm / commTarget) * 100) : 0;
      const commRem = Math.max(0, commTarget - achievedComm);
      const cTitle = document.getElementById('comm-bar-title');
      const cSub = document.getElementById('comm-bar-subtitle');
      const cBadge = document.getElementById('comm-bar-pct-badge');
      const cFill = document.getElementById('comm-bar-fill');
      const cRem = document.getElementById('comm-bar-remaining');
      const cMsg = document.getElementById('comm-bar-status-msg');

      if (cTitle && cFill && !commExplicit) {
        cTitle.textContent = `💵 Target Commission (${inspectedRepName})`;
        cSub.textContent = `${fmtMoney(achievedComm)} gagnés • objectif non défini`;
        cBadge.textContent = '—';
        cBadge.className = 'px-2 py-0.5 text-[11px] font-extrabold rounded-full bg-slate-100 text-slate-500 tabular-nums';
        cFill.style.width = '0%';
        syncRocketOnProgressBar(cFill, 0);
        cRem.textContent = 'Aucun objectif de commission pour le moment';
        cMsg.textContent = undefinedGoalMsg;
        cMsg.className = 'font-semibold text-slate-500';
      } else if (cTitle && cFill) {
        cTitle.textContent = `💵 Target Commission (${inspectedRepName})`;
        cSub.textContent = `${fmtMoney(achievedComm)} / ${fmtMoney(commTarget)}`;
        cBadge.textContent = `${commPct} %`;
        cFill.style.width = `${Math.min(100, commPct)}%`;
        syncRocketOnProgressBar(cFill, commPct);
        if (commPct >= 100) {
          cBadge.className = 'px-2 py-0.5 text-[11px] font-extrabold rounded-full bg-emerald-600 text-white tabular-nums';
          cRem.textContent = `Bonus dépassé de + ${fmtMoney(achievedComm - commTarget)} !`;
          cMsg.textContent = '🔥 Target commission atteint !';
          cMsg.className = 'font-extrabold text-emerald-700';
        } else if (commPct >= 65) {
          cBadge.className = 'px-2 py-0.5 text-[11px] font-extrabold rounded-full bg-amber-100 text-amber-800 tabular-nums';
          cRem.textContent = `Reste à gagner : ${fmtMoney(commRem)}`;
          cMsg.textContent = '💪 En feu !';
          cMsg.className = 'font-bold text-amber-700';
        } else {
          cBadge.className = 'px-2 py-0.5 text-[11px] font-extrabold rounded-full bg-emerald-100 text-emerald-800 tabular-nums';
          cRem.textContent = `Reste à gagner : ${fmtMoney(commRem)}`;
          cMsg.textContent = 'En progression';
          cMsg.className = 'font-semibold text-emerald-700/80';
        }
      }

      // 2. Mise à jour de la barre OBJECTIF VENTES ($)
      const salesPct = salesTarget > 0 ? Math.round((achievedSales / salesTarget) * 100) : 0;
      const salesRem = Math.max(0, salesTarget - achievedSales);
      const titleEl = document.getElementById('goal-bar-title');
      const subEl = document.getElementById('goal-bar-subtitle');
      const badgeEl = document.getElementById('goal-bar-pct-badge');
      const fillEl = document.getElementById('goal-bar-fill');
      const remEl = document.getElementById('goal-bar-remaining');
      const msgEl = document.getElementById('goal-bar-status-msg');

      if (!titleEl || !fillEl) return;
      titleEl.textContent = `🎯 Objectif Ventes (${inspectedRepName})`;
      if (!salesExplicit) {
        subEl.textContent = `${fmtMoney(achievedSales)} vendus • objectif non défini`;
        badgeEl.textContent = '—';
        badgeEl.className = 'px-2 py-0.5 text-[11px] font-extrabold rounded-full bg-slate-100 text-slate-500 tabular-nums';
        fillEl.style.width = '0%';
        syncRocketOnProgressBar(fillEl, 0);
        remEl.textContent = 'Aucun objectif de ventes pour le moment';
        msgEl.textContent = undefinedGoalMsg;
        msgEl.className = 'font-semibold text-slate-500';
        return;
      }
      subEl.textContent = `${fmtMoney(achievedSales)} / ${fmtMoney(salesTarget)}`;
      badgeEl.textContent = `${salesPct} %`;
      fillEl.style.width = `${Math.min(100, salesPct)}%`;
      syncRocketOnProgressBar(fillEl, salesPct);

      if (salesPct >= 100) {
        badgeEl.className = 'px-2 py-0.5 text-[11px] font-extrabold rounded-full bg-emerald-100 text-emerald-800 tabular-nums';
        fillEl.className = 'h-full rounded-full bg-gradient-to-r from-emerald-500 to-teal-400 transition-all duration-700 ease-out';
        remEl.textContent = `Objectif dépassé de + ${fmtMoney(achievedSales - salesTarget)} !`;
        const bStr = formatBonusPctShort(getBonusPctForRep(inspectedRepName, activeYm));
        msgEl.textContent = isAdmin ? `🔥 Target atteint (+${bStr} Rep / -${bStr} Oli actif)` : `🔥 Bonus +${bStr} actif sur vos nouvelles ventes !`;
        msgEl.className = 'font-extrabold text-emerald-600';
      } else if (salesPct >= 65) {
        badgeEl.className = 'px-2 py-0.5 text-[11px] font-extrabold rounded-full bg-amber-100 text-amber-800 tabular-nums';
        fillEl.className = 'h-full rounded-full bg-gradient-to-r from-indigo-600 via-amber-400 to-emerald-500 transition-all duration-700 ease-out';
        remEl.textContent = `Reste à vendre : ${fmtMoney(salesRem)}`;
        msgEl.textContent = '💪 Dernière ligne droite !';
        msgEl.className = 'font-bold text-amber-700';
      } else {
        badgeEl.className = 'px-2 py-0.5 text-[11px] font-extrabold rounded-full bg-indigo-50 text-indigo-700 tabular-nums';
        fillEl.className = 'h-full rounded-full bg-gradient-to-r from-indigo-600 to-sky-500 transition-all duration-700 ease-out';
        remEl.textContent = `Reste à vendre : ${fmtMoney(salesRem)}`;
        msgEl.textContent = 'En progression';
        msgEl.className = 'font-semibold text-slate-500';
      }
    }

    function renderEquipeTargetsTable() {
      const tbody = document.getElementById('tbody-equipe-targets');
      if (!tbody) return;
      if (currentProfile?.role !== 'admin') {
        tbody.innerHTML = '';
        return;
      }
      const repNames = Array.from(new Set([
        ...getActiveKpiReps(),
        ...state.representants.map(r => r.nom)
      ])).filter(rn => !isRepArchived(rn));

      tbody.innerHTML = repNames.map(rn => {
        const repObj = state.representants.find(r => r.nom === rn);
        const roleLabel = repObj ? (repObj.role === 'admin' ? 'Admin' : 'Rep') : 'Rep';
        const commVal = getCommissionTargetForRep(rn, currentMonthStr);
        const salesVal = getSalesTargetForRep(rn, currentMonthStr);
        const bonusValPct = Math.round(getBonusPctForRep(rn, currentMonthStr) * 10000) / 100;
        const safeKey = rn.replace(/'/g, "\'");
        return `
          <tr class="hover:bg-slate-50/80 transition" data-target-rep="${esc(rn)}">
            <td class="py-3 px-4">
              <div class="font-bold text-slate-900 flex items-center gap-2">
                <span>${esc(rn)}</span>
                <span class="text-[10px] font-semibold px-2 py-0.5 rounded-full bg-slate-100 text-slate-600">${roleLabel}</span>
              </div>
            </td>
            <td class="py-2.5 px-4">
              <div class="relative max-w-[170px]">
                <input type="number" step="100" min="0" data-comm-input="${esc(rn)}" value="${commVal}" class="w-full pl-3 pr-7 py-2 text-xs font-extrabold text-emerald-900 bg-emerald-50/50 border border-emerald-200 rounded-xl tabular-nums focus:bg-white focus:outline-none focus:ring-2 focus:ring-emerald-600" />
                <span class="absolute right-2.5 top-2 text-xs font-bold text-emerald-700">$</span>
              </div>
            </td>
            <td class="py-2.5 px-4">
              <div class="relative max-w-[170px]">
                <input type="number" step="500" min="0" data-sales-input="${esc(rn)}" value="${salesVal}" class="w-full pl-3 pr-7 py-2 text-xs font-bold text-slate-900 bg-slate-50 border border-slate-200 rounded-xl tabular-nums focus:bg-white focus:outline-none focus:ring-2 focus:ring-indigo-600" />
                <span class="absolute right-2.5 top-2 text-xs font-semibold text-slate-400">$</span>
              </div>
            </td>
            <td class="py-2.5 px-4">
              <div class="relative max-w-[130px]">
                <input type="number" step="0.25" min="0" max="50" data-bonus-input="${esc(rn)}" value="${bonusValPct}" class="w-full pl-3 pr-7 py-2 text-xs font-extrabold text-amber-900 bg-amber-50/60 border border-amber-200 rounded-xl tabular-nums focus:bg-white focus:outline-none focus:ring-2 focus:ring-amber-500" />
                <span class="absolute right-2.5 top-2 text-xs font-bold text-amber-700">%</span>
              </div>
            </td>
            <td class="py-2.5 px-4 text-right">
              <button type="button" onclick="saveSingleEquipeTarget('${safeKey}')" class="px-3 py-1.5 text-xs font-bold text-indigo-700 bg-indigo-50 hover:bg-indigo-100 border border-indigo-200 rounded-xl transition">
                Enregistrer
              </button>
            </td>
          </tr>
        `;
      }).join('');
    }

    window.saveSingleEquipeTarget = async (repName) => {
      if (currentProfile?.role !== 'admin') return toast('Réservé à l’administrateur.', 'error');
      const commIn = document.querySelector(`input[data-comm-input="${repName}"]`);
      const salesIn = document.querySelector(`input[data-sales-input="${repName}"]`);
      const bonusIn = document.querySelector(`input[data-bonus-input="${repName}"]`);
      const newComm = parseFloat(commIn?.value) || 0;
      const newSales = parseFloat(salesIn?.value) || 0;
      const rawBonusPct = parseFloat(bonusIn?.value);
      const newBonusDec = isNaN(rawBonusPct) ? getBonusPctForRep(repName, currentMonthStr) : Math.max(0, rawBonusPct / 100);
      if (newComm <= 0 || newSales <= 0) return toast('Entrez des montants supérieurs à 0 $', 'error');

      await saveRepTargetsToStorageAndCloud(repName, currentMonthStr, newSales, newComm, true, newBonusDec);
      rebuildKpiStore();
      renderMesVentes();
      renderClassement();
      renderKpiDashboard();
      renderEquipeTargetsTable();
      toast(`Targets enregistrés pour ${repName} : ${fmtMoney0(newComm)} comm. / ${fmtMoney0(newSales)} ventes / +${formatBonusPctShort(newBonusDec)} bonus 🎯`);
    };

    document.getElementById('btn-save-all-equipe-targets')?.addEventListener('click', async () => {
      if (currentProfile?.role !== 'admin') return toast('Réservé à l’administrateur.', 'error');
      const rows = document.querySelectorAll('#tbody-equipe-targets tr[data-target-rep]');
      let count = 0;
      for (const tr of rows) {
        const rn = tr.getAttribute('data-target-rep');
        const newComm = parseFloat(tr.querySelector('input[data-comm-input]')?.value) || 0;
        const newSales = parseFloat(tr.querySelector('input[data-sales-input]')?.value) || 0;
        const rawBonusPct = parseFloat(tr.querySelector('input[data-bonus-input]')?.value);
        const newBonusDec = isNaN(rawBonusPct) ? getBonusPctForRep(rn, currentMonthStr) : Math.max(0, rawBonusPct / 100);
        if (rn && newComm > 0 && newSales > 0) {
          await saveRepTargetsToStorageAndCloud(rn, currentMonthStr, newSales, newComm, true, newBonusDec);
          count++;
        }
      }
      rebuildKpiStore();
      renderMesVentes();
      renderClassement();
      renderKpiDashboard();
      renderEquipeTargetsTable();
      toast(`${count} représentant(s) mis à jour 🎯`);
    });

    let commDisplayMode = localStorage.getItem('korpo_comm_display_mode_v1') || 'vendu';

    window.setCommDisplayMode = (mode) => {
      commDisplayMode = mode === 'confirme' ? 'confirme' : 'vendu';
      lsSet('korpo_comm_display_mode_v1', commDisplayMode);
      syncCommDisplayModeButtons();
      renderMesVentes();
      if (currentProfile?.role === 'admin') {
        renderAdminDashboard();
      }
    };

    function syncCommDisplayModeButtons() {
      const isConf = commDisplayMode === 'confirme';
      const rV = document.getElementById('btn-rep-comm-mode-vendu');
      const rC = document.getElementById('btn-rep-comm-mode-confirme');
      if (rV && rC) {
        rV.className = !isConf
          ? 'px-2 py-0.5 rounded-md bg-white text-emerald-900 shadow-2xs transition'
          : 'px-2 py-0.5 rounded-md text-emerald-100 hover:text-white transition';
        rC.className = isConf
          ? 'px-2 py-0.5 rounded-md bg-white text-emerald-900 shadow-2xs transition'
          : 'px-2 py-0.5 rounded-md text-emerald-100 hover:text-white transition';
      }
      const dV = document.getElementById('btn-dash-comm-mode-vendu');
      const dC = document.getElementById('btn-dash-comm-mode-confirme');
      if (dV && dC) {
        dV.className = !isConf
          ? 'px-3 py-1 rounded-lg bg-indigo-600 text-white shadow-2xs transition'
          : 'px-3 py-1 rounded-lg text-slate-600 hover:text-slate-900 transition';
        dC.className = isConf
          ? 'px-3 py-1 rounded-lg bg-emerald-600 text-white shadow-2xs transition'
          : 'px-3 py-1 rounded-lg text-slate-600 hover:text-slate-900 transition';
      }
    }

    window.toggleRepFutureCommQuickView = () => {
      const sel = document.getElementById('filter-rep-period');
      if (!sel) return;
      sel.value = sel.value === 'future' ? 'this_month' : 'future';
      syncPeriodInputsVisibility('filter-rep', false);
      if (sel.value === 'future') {
        switchTab('tab-mes-commissions');
        toast('🔮 Affichage de vos commissions prévues dans les prochains mois');
      }
      renderMesVentes();
    };

    window.dismissRepDepositReminderBanner = () => {
      if (!currentProfile) return;
      try {
        localStorage.setItem(`korpo_deposit_reminder_dismissed_${currentProfile.id}_${currentDateStr}`, '1');
      } catch (e) {}
      document.getElementById('rep-deposit-reminder-banner')?.classList.add('hidden');
    };

    window.filterRepPendingDepositsOnly = () => {
      const statSel = document.getElementById('filter-rep-statut');
      if (statSel) {
        statSel.value = 'depot_attente';
        renderMesVentes();
      }
    };

    function isPenultimateDayOfMonth(dateObj = new Date()) {
      const y = dateObj.getFullYear();
      const m = dateObj.getMonth();
      const lastDay = new Date(y, m + 1, 0).getDate();
      return dateObj.getDate() === (lastDay - 1);
    }

    function maybeSendPenultimateDayUnconfirmedPush(unconfirmedMonthSales) {
      if (!currentProfile || currentProfile.role === 'admin') return;
      if (!Array.isArray(unconfirmedMonthSales) || !unconfirmedMonthSales.length) return;
      const now = new Date();
      if (!isPenultimateDayOfMonth(now)) return;

      const sentKey = `korpo_penultimate_push_sent_${currentProfile.id}_${currentMonthStr}`;
      if (localStorage.getItem(sentKey) === '1') return;

      const clientNames = unconfirmedMonthSales.map(v => {
        const cli = v.client_nom || state.clients.find(c => String(c.id) === String(v.client_id))?.nom || 'Client';
        return `${cli} (${fmtMoney0(v.valeur)})`;
      });
      const previewClients = clientNames.slice(0, 3).join(', ') + (clientNames.length > 3 ? ` +${clientNames.length - 3}` : '');
      const title = `⏳ Fin de mois demain — ${unconfirmedMonthSales.length} dossier(s) non confirmé(s)`;
      const body = `Dépôt à confirmer pour ${formatYearMonthLabelFr(currentMonthStr)} : ${previewClients}. Ajoute la preuve avant la clôture !`;

      try {
        localStorage.setItem(sentKey, '1');
      } catch (e) {}

      if (isMentionPushCurrentlyEnabled()) {
        if (korpoSwRegistration && typeof korpoSwRegistration.showNotification === 'function') {
          korpoSwRegistration.showNotification(title, {
            body,
            tag: `korpo-month-end-${currentMonthStr}`,
            renotify: true,
            data: { url: './' }
          }).catch(() => {});
        } else if (Notification.permission === 'granted') {
          try {
            const n = new Notification(title, { body, tag: `korpo-month-end-${currentMonthStr}` });
            n.onclick = () => { window.focus(); filterRepPendingDepositsOnly(); n.close(); };
          } catch (e) {}
        }
      }
      toast(`${title} : ${previewClients}`);
    }

    window._pushRelanceMap = new Map();

    window.closePushRelanceModal = () => {
      document.getElementById('push-relance-modal')?.classList.add('hidden');
    };

    window.updatePushRelanceSelectionCount = () => {
      const checked = document.querySelectorAll('.push-relance-rep-cb:checked').length;
      const lbl = document.getElementById('push-relance-selected-count');
      const btn = document.getElementById('btn-confirm-send-push-relance');
      if (lbl) lbl.textContent = `${checked} représentant(s) sélectionné(s)`;
      if (btn) {
        btn.disabled = checked === 0;
        btn.classList.toggle('opacity-50', checked === 0);
        btn.classList.toggle('cursor-not-allowed', checked === 0);
      }
    };

    window.toggleAllPushRelanceCheckboxes = (checkedState) => {
      document.querySelectorAll('.push-relance-rep-cb').forEach(cb => {
        cb.checked = Boolean(checkedState);
      });
      updatePushRelanceSelectionCount();
    };

    window.triggerUnconfirmedMonthEndPushNow = () => {
      if (currentProfile?.role !== 'admin') return;
      const monthUnconfirmed = (state.ventes || []).filter(v =>
        v.statut_depot === 'a_confirmer' && getSaleStatsMonth(v) === currentMonthStr
      );
      if (!monthUnconfirmed.length) {
        return toast(`Aucun dossier non confirmé pour ${formatYearMonthLabelFr(currentMonthStr)} ✓`);
      }

      const byRep = new Map();
      monthUnconfirmed.forEach(v => {
        const repObj = state.representants.find(r => String(r.id) === String(v.representant_id));
        const rName = repObj?.nom || v.rep_nom || '';
        if (!rName) return;
        if (!byRep.has(rName)) byRep.set(rName, { dossiers: [], totalVol: 0, totalComm: 0 });
        const cli = v.client_nom || state.clients.find(c => String(c.id) === String(v.client_id))?.nom || 'Client';
        const entry = byRep.get(rName);
        entry.dossiers.push(`${cli} (${fmtMoney0(v.valeur)})`);
        entry.totalVol += Number(v.valeur || 0);
        entry.totalComm += Number(v.commission_rep || 0);
      });

      window._pushRelanceMap = byRep;
      const subEl = document.getElementById('push-relance-subtitle');
      if (subEl) {
        subEl.textContent = `${formatYearMonthLabelFr(currentMonthStr)} • ${monthUnconfirmed.length} dossier(s) en attente au total`;
      }

      const listEl = document.getElementById('push-relance-reps-list');
      if (listEl) {
        listEl.innerHTML = Array.from(byRep.entries()).map(([repName, info]) => `
          <label class="flex items-start gap-3 p-3.5 rounded-2xl border border-slate-200 bg-slate-50/70 hover:bg-indigo-50/50 cursor-pointer transition select-none">
            <input type="checkbox" value="${esc(repName)}" checked onchange="updatePushRelanceSelectionCount()" class="push-relance-rep-cb mt-0.5 w-4 h-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-600" />
            <div class="min-w-0 flex-1 space-y-1">
              <div class="flex items-center justify-between gap-2">
                <span class="text-xs sm:text-sm font-extrabold text-slate-900">${esc(repName)}</span>
                <span class="px-2 py-0.5 text-[10px] font-extrabold rounded-full bg-amber-100 text-amber-900 border border-amber-200 shrink-0">
                  ${info.dossiers.length} dossier(s) • ${fmtMoney0(info.totalVol)}
                </span>
              </div>
              <div class="text-[11px] text-slate-600 leading-snug break-words">
                ${esc(info.dossiers.join(', '))}
              </div>
            </div>
          </label>
        `).join('');
      }

      updatePushRelanceSelectionCount();
      document.getElementById('push-relance-modal')?.classList.remove('hidden');
    };

    window.executeSelectedPushRelance = async () => {
      if (currentProfile?.role !== 'admin') return;
      const selectedReps = Array.from(document.querySelectorAll('.push-relance-rep-cb:checked')).map(cb => cb.value);
      if (!selectedReps.length) {
        return toast('Sélectionnez au moins un représentant.', 'error');
      }

      const btn = document.getElementById('btn-confirm-send-push-relance');
      const origTxt = btn ? btn.innerHTML : '';
      if (btn) {
        btn.disabled = true;
        btn.innerHTML = '⏳ Envoi en cours...';
      }

      let invokedCount = 0;
      let totalDossiers = 0;
      for (const repName of selectedReps) {
        const info = window._pushRelanceMap?.get(repName);
        if (!info) continue;
        totalDossiers += info.dossiers.length;
        const summary = `⏳ Rappel fin de mois (${formatYearMonthLabelFr(currentMonthStr)}) — ${info.dossiers.length} dossier(s) non confirmé(s) : ${info.dossiers.join(', ')}`;
        try {
          await sb.functions.invoke('send-mention-push', {
            body: {
              msgId: `month-end-${currentMonthStr}-${repName}-${Date.now()}`,
              authorId: currentProfile.id,
              authorName: currentProfile.nom,
              text: summary,
              mentionedNames: [repName],
              customTitle: `⏳ Dossier(s) non confirmé(s) — ${formatYearMonthLabelFr(currentMonthStr)}`,
              eventType: 'month_end_reminder'
            }
          });
          invokedCount++;
        } catch (e) {}
      }

      if (btn) {
        btn.disabled = false;
        btn.innerHTML = origTxt;
      }
      closePushRelanceModal();
      toast(`🔔 Notification push envoyée à ${invokedCount} représentant(s) (${totalDossiers} dossier(s)) !`);
    };

    function renderRepDepositReminderBanner(scopeSales, isAdmin, isInspectingRep) {
      const banner = document.getElementById('rep-deposit-reminder-banner');
      if (!banner || !currentProfile) return;
      if (isAdmin && !isInspectingRep) {
        banner.classList.add('hidden');
        return;
      }

      const isPenultimate = isPenultimateDayOfMonth(new Date());
      const allUnconfirmedThisMonth = (scopeSales || [])
        .filter(v => v.statut_depot === 'a_confirmer' && getSaleStatsMonth(v) === currentMonthStr)
        .sort((a, b) => String(a.date_vente || '').localeCompare(String(b.date_vente || '')));

      if (isPenultimate && !isInspectingRep && allUnconfirmedThisMonth.length > 0) {
        maybeSendPenultimateDayUnconfirmedPush(allUnconfirmedThisMonth);
      }

      const dismissedToday = localStorage.getItem(`korpo_deposit_reminder_dismissed_${currentProfile.id}_${currentDateStr}`) === '1';
      if (dismissedToday) {
        banner.classList.add('hidden');
        return;
      }

      const targetList = isPenultimate
        ? allUnconfirmedThisMonth
        : (scopeSales || [])
            .filter(v => v.statut_depot === 'a_confirmer' && !getProofRef(v))
            .sort((a, b) => String(a.date_vente || '').localeCompare(String(b.date_vente || '')));

      if (!targetList.length) {
        banner.classList.add('hidden');
        return;
      }

      const pendingWithMissingProof = targetList;
      const lockedCommTotal = pendingWithMissingProof.reduce((s, v) => s + Number(v.commission_rep || 0), 0);
      const oldestDays = Math.max(...pendingWithMissingProof.map(v => getDaysSince(v.date_vente)));
      const titleEl = document.getElementById('rep-deposit-reminder-title');
      const descEl = document.getElementById('rep-deposit-reminder-desc');
      const listEl = document.getElementById('rep-deposit-reminder-list');

      if (titleEl) {
        titleEl.textContent = isPenultimate
          ? `🚨 Avant-dernier jour du mois : ${pendingWithMissingProof.length} dossier(s) non confirmé(s) (${fmtMoney(lockedCommTotal)} de comm.)`
          : `${pendingWithMissingProof.length} dossier(s) attendent leur preuve Interac (${fmtMoney(lockedCommTotal)} de commission à débloquer)`;
      }
      if (descEl) {
        descEl.textContent = isPenultimate
          ? `Demain est le dernier jour de ${formatYearMonthLabelFr(currentMonthStr)} : tout dossier non confirmé sera reporté au mois prochain.`
          : (oldestDays >= 3
            ? `Le plus ancien date d'il y a ${oldestDays} jour(s). Ajoutez la capture du virement ou du chèque en 1 clic ci-dessous :`
            : 'Ajoutez la photo du virement Interac ou du chèque pour accélérer la confirmation de votre commission :');
      }
      if (listEl) {
        listEl.innerHTML = pendingWithMissingProof.slice(0, 6).map(v => {
          const cli = v.client_nom || state.clients.find(c => String(c.id) === String(v.client_id))?.nom || 'Client';
          const ent = v.entreprise_nom || state.entreprises.find(e => String(e.id) === String(v.entreprise_id))?.nom || '';
          const days = getDaysSince(v.date_vente);
          const inputId = `quick-proof-up-${v.id}`;
          return `
            <div class="bg-white rounded-xl border border-amber-200/90 p-2.5 flex items-center justify-between gap-2">
              <div class="min-w-0">
                <div class="text-xs font-extrabold text-slate-900 truncate">${esc(cli)} <span class="font-semibold text-slate-500">(${esc(ent)})</span></div>
                <div class="text-[11px] text-slate-600 tabular-nums">
                  Comm. : <strong class="text-emerald-700">${fmtMoney(v.commission_rep)}</strong> • ${days > 0 ? `il y a ${days}j` : "aujourd'hui"}
                </div>
              </div>
              <label for="${inputId}" class="px-2.5 py-1.5 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-[11px] font-extrabold shrink-0 cursor-pointer transition">
                📷 Photo
                <input id="${inputId}" type="file" accept="image/*" class="sr-only" onchange="if(this.files && this.files[0]) attachProofToExistingSale('${v.id}', this.files[0])" />
              </label>
            </div>
          `;
        }).join('');
      }
      banner.classList.remove('hidden');
    }

    function buildWeeklyDigestSnapshot() {
      const wb = getCurrentWeekBounds(today);
      const weekPeriod = { mode: 'this_week', weekBounds: wb };
      const weekSales = (state.ventes || []).filter(v => matchesSalePeriod(v, weekPeriod, true));
      const weekVol = weekSales.reduce((s, v) => s + Number(v.valeur || 0), 0);
      const weekConfirmedVol = weekSales.filter(v => v.statut_depot === 'depot_recu').reduce((s, v) => s + Number(v.valeur || 0), 0);
      const weekCommRep = weekSales.reduce((s, v) => s + Number(v.commission_rep || 0), 0);
      const weekCommOli = weekSales.reduce((s, v) => s + Number(v.commission_olivier || 0), 0);

      const repRanking = (state.representants || []).map(r => {
        const sub = weekSales.filter(v => String(v.representant_id) === String(r.id));
        return {
          nom: r.nom,
          count: sub.length,
          vol: sub.reduce((s, v) => s + Number(v.valeur || 0), 0)
        };
      }).filter(x => x.vol > 0 || x.count > 0).sort((a, b) => b.vol - a.vol);

      const allPending = (state.ventes || [])
        .filter(v => v.statut_depot === 'a_confirmer')
        .sort((a, b) => String(a.date_vente || '').localeCompare(String(b.date_vente || '')));
      const agedPending = allPending.filter(v => getDaysSince(v.date_vente) >= 3);

      // Projection fin de mois
      const [yStr, mStr] = currentMonthStr.split('-');
      const totalDaysInM = new Date(parseInt(yStr, 10), parseInt(mStr, 10), 0).getDate();
      const totalBiz = Math.max(1, countBusinessDaysInMonth(currentMonthStr));
      const elapsedBiz = Math.max(1, countBusinessDaysInMonth(currentMonthStr, Math.min(today.getDate(), totalDaysInM)));
      const monthPeriod = { mode: 'this_month', singleMonth: currentMonthStr };
      const monthSales = (state.ventes || []).filter(v => matchesSalePeriod(v, monthPeriod, true));
      const monthVol = monthSales.reduce((s, v) => s + Number(v.valeur || 0), 0);
      const projectedMonthVol = Math.round((monthVol / elapsedBiz) * totalBiz);

      return {
        wb,
        weekSales,
        weekVol,
        weekConfirmedVol,
        weekCommRep,
        weekCommOli,
        repRanking,
        allPending,
        agedPending,
        monthVol,
        projectedMonthVol
      };
    }

    function formatWeeklyDigestPlainText(snap) {
      const lines = [
        `📊 BILAN HEBDOMADAIRE KORPO SOLUTIONS`,
        `Période : Du ${formatDateLabelFr(snap.wb.start)} au ${formatDateLabelFr(snap.wb.end)}`,
        ``,
        `• Volume signé cette semaine : ${fmtMoney(snap.weekVol)} (${snap.weekSales.length} dossier${snap.weekSales.length > 1 ? 's' : ''})`,
        `• Volume confirmé (dépôt reçu) : ${fmtMoney(snap.weekConfirmedVol)}`,
        `• Commissions générées (Olivier) : ${fmtMoney(snap.weekCommOli)}`,
        `• Projection fin de mois (${formatYearMonthLabelFr(currentMonthStr)}) : ${fmtMoney0(snap.projectedMonthVol)}`,
        ``,
        `🏆 TOP REPRÉSENTANTS DE LA SEMAINE :`,
        ...(snap.repRanking.length
          ? snap.repRanking.map((r, i) => `  ${i + 1}. ${r.nom} — ${fmtMoney(r.vol)} (${r.count} vente${r.count > 1 ? 's' : ''})`)
          : [`  Aucune vente enregistrée cette semaine.`]),
        ``,
        `⏳ DÉPÔTS À CONFIRMER : ${snap.allPending.length} dossier(s) au total (dont ${snap.agedPending.length} en attente depuis 3+ jours)`
      ];
      return lines.join('\n');
    }

    window.toggleWeeklyDigestAutoFriday = (enabled) => {
      try {
        localStorage.setItem('korpo_weekly_digest_auto_friday_v1', enabled ? '1' : '0');
      } catch (e) {}
      toast(enabled ? 'Bilan automatique du vendredi 17h activé ✓' : 'Affichage automatique du vendredi désactivé');
    };

    window.openWeeklyDigestModal = () => {
      if (currentProfile?.role !== 'admin') return;
      const snap = buildWeeklyDigestSnapshot();
      const subEl = document.getElementById('weekly-digest-subtitle');
      if (subEl) {
        subEl.textContent = `Semaine du ${formatDateLabelFr(snap.wb.start)} au ${formatDateLabelFr(snap.wb.end)}`;
      }
      const autoCb = document.getElementById('weekly-digest-auto-friday');
      if (autoCb) {
        autoCb.checked = localStorage.getItem('korpo_weekly_digest_auto_friday_v1') !== '0';
      }

      const body = document.getElementById('weekly-digest-body');
      if (body) {
        const medals = ['🥇', '🥈', '🥉'];
        body.innerHTML = `
          <div class="grid grid-cols-2 sm:grid-cols-4 gap-3">
            <div class="bg-slate-50 rounded-2xl border border-slate-200/80 p-3.5">
              <div class="text-[11px] font-semibold text-slate-500">Ventes semaine</div>
              <div class="text-lg sm:text-xl font-extrabold text-slate-900 mt-1 tabular-nums">${fmtMoney0(snap.weekVol)}</div>
              <div class="text-[11px] text-slate-400 mt-0.5">${snap.weekSales.length} dossier(s)</div>
            </div>
            <div class="bg-emerald-50/70 rounded-2xl border border-emerald-200/80 p-3.5">
              <div class="text-[11px] font-semibold text-emerald-800">Confirmé semaine</div>
              <div class="text-lg sm:text-xl font-extrabold text-emerald-700 mt-1 tabular-nums">${fmtMoney0(snap.weekConfirmedVol)}</div>
              <div class="text-[11px] text-emerald-700/80 mt-0.5">Dépôts reçus</div>
            </div>
            <div class="bg-indigo-50/70 rounded-2xl border border-indigo-200/80 p-3.5">
              <div class="text-[11px] font-semibold text-indigo-800">Projection mois</div>
              <div class="text-lg sm:text-xl font-extrabold text-indigo-700 mt-1 tabular-nums">${fmtMoney0(snap.projectedMonthVol)}</div>
              <div class="text-[11px] text-indigo-700/80 mt-0.5">Cadence actuelle</div>
            </div>
            <div class="bg-amber-50/70 rounded-2xl border border-amber-200/80 p-3.5">
              <div class="text-[11px] font-semibold text-amber-800">Dépôts en attente</div>
              <div class="text-lg sm:text-xl font-extrabold text-amber-700 mt-1 tabular-nums">${snap.allPending.length}</div>
              <div class="text-[11px] text-amber-800/80 mt-0.5">${snap.agedPending.length} de 3+ jours</div>
            </div>
          </div>

          <div class="bg-slate-50 rounded-2xl border border-slate-200/80 p-4 space-y-2.5">
            <div class="text-xs font-extrabold uppercase tracking-wider text-slate-700">🏆 Podium de la semaine</div>
            ${snap.repRanking.length
              ? `<div class="divide-y divide-slate-200/70">
                  ${snap.repRanking.map((r, i) => `
                    <div class="py-2 flex items-center justify-between text-xs">
                      <div class="flex items-center gap-2 font-extrabold text-slate-900">
                        <span>${medals[i] || `#${i + 1}`}</span>
                        <span>${esc(r.nom)}</span>
                        <span class="text-[11px] font-semibold text-slate-400">(${r.count} vente${r.count > 1 ? 's' : ''})</span>
                      </div>
                      <span class="font-extrabold text-indigo-700 tabular-nums">${fmtMoney(r.vol)}</span>
                    </div>
                  `).join('')}
                </div>`
              : `<div class="text-xs text-slate-400 italic">Aucune vente enregistrée cette semaine.</div>`}
          </div>

          ${snap.agedPending.length ? `
            <div class="bg-rose-50/60 rounded-2xl border border-rose-200 p-4 space-y-2">
              <div class="flex items-center justify-between">
                <span class="text-xs font-extrabold text-rose-900">⚠️ Dossiers en attente de dépôt depuis 3+ jours (${snap.agedPending.length})</span>
                <button type="button" onclick="closeWeeklyDigestModal(); goToPendingSales();" class="text-xs font-extrabold text-rose-700 hover:underline cursor-pointer">Gérer →</button>
              </div>
              <div class="space-y-1.5">
                ${snap.agedPending.slice(0, 5).map(v => {
                  const repN = state.representants.find(r => String(r.id) === String(v.representant_id))?.nom || '—';
                  const cliN = state.clients.find(c => String(c.id) === String(v.client_id))?.nom || '—';
                  const entN = state.entreprises.find(e => String(e.id) === String(v.entreprise_id))?.nom || '';
                  return `<div class="text-xs text-slate-700 flex items-center justify-between gap-2 bg-white/90 px-3 py-1.5 rounded-xl border border-rose-200/70">
                    <span class="truncate"><strong>${esc(repN)}</strong> • ${esc(cliN)} (${esc(entN)})</span>
                    <span class="font-bold text-rose-700 shrink-0 tabular-nums">${fmtMoney(v.valeur)} • ${getDaysSince(v.date_vente)}j</span>
                  </div>`;
                }).join('')}
              </div>
            </div>
          ` : ''}
        `;
      }

      try {
        localStorage.setItem(`korpo_weekly_digest_shown_${wbKeyFromDate(snap.wb.start)}`, '1');
      } catch (e) {}
      document.getElementById('weekly-digest-modal')?.classList.remove('hidden');
    };

    function wbKeyFromDate(ymd) {
      return String(ymd || currentDateStr);
    }

    window.closeWeeklyDigestModal = () => {
      document.getElementById('weekly-digest-modal')?.classList.add('hidden');
    };

    window.copyWeeklyDigestText = async () => {
      const txt = formatWeeklyDigestPlainText(buildWeeklyDigestSnapshot());
      try {
        await navigator.clipboard.writeText(txt);
        toast('Bilan hebdomadaire copié dans le presse-papiers 📋');
      } catch (e) {
        toast('Impossible de copier automatiquement', 'error');
      }
    };

    window.publishWeeklyDigestToTeamChat = async () => {
      const snap = buildWeeklyDigestSnapshot();
      const topLine = snap.repRanking.length
        ? snap.repRanking.slice(0, 3).map((r, i) => `${['🥇', '🥈', '🥉'][i]} ${r.nom} (${fmtMoney0(r.vol)})`).join(' • ')
        : 'Aucune vente enregistrée cette semaine';
      const msg = `📊 Bilan de la semaine (${formatDateLabelFr(snap.wb.start)} → ${formatDateLabelFr(snap.wb.end)}) : ${fmtMoney0(snap.weekVol)} signés (${snap.weekSales.length} dossiers) ! Podium : ${topLine} 🚀`;
      const inputEl = document.getElementById('chat-message-input');
      const formEl = document.getElementById('form-team-chat');
      if (inputEl && formEl) {
        inputEl.value = msg;
        formEl.dispatchEvent(new Event('submit', { cancelable: true, bubbles: true }));
        closeWeeklyDigestModal();
        toast('Bilan publié dans Korpo Top Earner 🏆');
      }
    };

    window.sendWeeklyDigestEmail = () => {
      const snap = buildWeeklyDigestSnapshot();
      const body = formatWeeklyDigestPlainText(snap);
      const subject = `Bilan hebdomadaire Korpo Solutions — Semaine du ${formatDateLabelFr(snap.wb.start)}`;
      const adminEmail = currentProfile?.courriel || '';
      window.open(`mailto:${encodeURIComponent(adminEmail)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`, '_blank');
    };

    function maybeTriggerFridayWeeklyDigest() {
      if (!currentProfile || currentProfile.role !== 'admin') return;
      if (localStorage.getItem('korpo_weekly_digest_auto_friday_v1') === '0') return;
      const now = new Date();
      const isFridayAfter5pm = now.getDay() === 5 && now.getHours() >= 17;
      if (!isFridayAfter5pm) return;
      const wb = getCurrentWeekBounds(now);
      const seenKey = `korpo_weekly_digest_shown_${wbKeyFromDate(wb.start)}`;
      if (localStorage.getItem(seenKey) === '1') return;
      setTimeout(() => openWeeklyDigestModal(), 900);
    }

    function renderMesVentes() {
      syncCommDisplayModeButtons();
      syncRedPhoneUI();
      renderDailyRdvTrackerCard();
      const isAdmin = currentProfile.role === 'admin';
      const periodState = readPeriodFilterState('filter-rep');
      const futureQuickBtn = document.getElementById('btn-rep-future-comm');
      if (futureQuickBtn) {
        const isFutureActive = periodState.mode === 'future';
        futureQuickBtn.className = isFutureActive
          ? 'w-full sm:w-auto px-3.5 py-2 text-xs font-extrabold text-white bg-emerald-600 hover:bg-emerald-700 border border-emerald-600 rounded-xl transition flex items-center justify-center gap-1.5 shadow-xs'
          : 'w-full sm:w-auto px-3.5 py-2 text-xs font-extrabold text-emerald-900 bg-emerald-50 hover:bg-emerald-100 border border-emerald-300 rounded-xl transition flex items-center justify-center gap-1.5 shadow-2xs';
      }
      const statutFilter = document.getElementById('filter-rep-statut')?.value || '';
      const scopeSelEl = document.getElementById('filter-admin-scope');
      if (scopeSelEl && !isAdmin) scopeSelEl.classList.add('hidden');
      const scopeFilter = (isAdmin && scopeSelEl) ? scopeSelEl.value : 'all';
      const q = String(window._globalClientFilterQuery || '').trim().toLowerCase();
      const currentYear = new Date().getFullYear().toString();

      const isInspectingRep = isAdmin && typeof scopeFilter === 'string' && scopeFilter.startsWith('rep:');
      const inspectedRepId = isInspectingRep ? scopeFilter.slice(4) : null;
      const inspectedRepObj = isInspectingRep
        ? state.representants.find(r => String(r.id) === inspectedRepId)
        : null;
      const showRepCol = isAdmin && !isInspectingRep;
      const thRepColEl = document.getElementById('th-rep-col');
      if (thRepColEl) thRepColEl.classList.toggle('hidden', !showRepCol);
      const thRepCommColEl = document.getElementById('th-rep-comm-col');
      if (thRepCommColEl) thRepCommColEl.classList.toggle('hidden', !showRepCol);

      const subtitleEl = document.getElementById('mes-ventes-subtitle');
      const commCardTitleEl = document.getElementById('kpi-rep-comm-mois')?.parentElement?.querySelector('span');
      if (isInspectingRep && inspectedRepObj) {
        if (subtitleEl) subtitleEl.textContent = `Vue miroir de ${inspectedRepObj.nom} — ses ventes, ses commissions et ses objectifs`;
        if (commCardTitleEl) commCardTitleEl.textContent = `Commission de ${inspectedRepObj.nom} (Période)`;
      } else {
        if (subtitleEl) subtitleEl.textContent = activeMobilePersonalView === 'commissions'
          ? (isAdmin ? 'Vos commissions directes et d’équipe, paiements reçus et soldes à venir' : 'Vos gains confirmés, paiements reçus et soldes à venir')
          : (isAdmin ? 'Vos ventes directes et le suivi des dossiers de l’équipe' : 'Suivi confidentiel de vos dossiers fermés et de vos dépôts');
        if (commCardTitleEl) commCardTitleEl.textContent = 'Ma commission (Période)';
      }

      const scopeFilteredAll = isInspectingRep
        ? state.ventes
            .filter(v => String(v.representant_id) === inspectedRepId)
            .map(v => ({
              ...v,
              client_nom: state.clients.find(c => c.id === v.client_id)?.nom || '—',
              entreprise_nom: state.entreprises.find(e => e.id === v.entreprise_id)?.nom || '—',
              rep_nom: inspectedRepObj?.nom || '—',
              commission_rep: Number(v.commission_rep || 0)
            }))
        : state.mesVentes.filter(v => {
            if (isAdmin && scopeFilter === 'direct' && v.representant_id !== currentProfile.id) return false;
            return true;
          });

      renderRepDepositReminderBanner(scopeFilteredAll, isAdmin, isInspectingRep);

      const allFutureRepSales = scopeFilteredAll.filter(v => Boolean(getSaleTargetFutureMonth(v)));
      const repFutureBadge = document.getElementById('rep-future-comm-badge');
      const repFutureBtn = document.getElementById('btn-rep-future-comm');
      if (repFutureBadge) repFutureBadge.textContent = String(allFutureRepSales.length);
      if (repFutureBtn) {
        const isRepFuture = periodState.mode === 'future';
        repFutureBtn.className = isRepFuture
          ? 'w-full sm:w-auto px-3.5 py-2 text-xs font-extrabold text-white bg-emerald-600 hover:bg-emerald-700 border border-emerald-600 rounded-xl transition flex items-center justify-center gap-1.5 shadow-sm ring-2 ring-emerald-400/40'
          : 'w-full sm:w-auto px-3.5 py-2 text-xs font-extrabold text-emerald-900 bg-emerald-50 hover:bg-emerald-100 border border-emerald-300 rounded-xl transition flex items-center justify-center gap-1.5 shadow-2xs';
        if (repFutureBadge) {
          repFutureBadge.className = isRepFuture
            ? 'px-1.5 py-0.2 text-[10px] font-extrabold rounded-full bg-white text-emerald-800 tabular-nums'
            : 'px-1.5 py-0.2 text-[10px] font-extrabold rounded-full bg-emerald-600 text-white tabular-nums';
        }
      }

      const filtered = scopeFilteredAll.filter(v => {
        if (!matchesSalePeriod(v, periodState, true)) return false;
        if (statutFilter === 'en_attente' && v.statut_depot === 'depot_recu' && v.statut_paiement === 'paye') return false;
        if (statutFilter === 'depot_attente' && v.statut_depot !== 'a_confirmer') return false;
        if (statutFilter === 'non_paye' && v.statut_paiement !== 'non_paye') return false;
        if (statutFilter === 'paye' && v.statut_paiement !== 'paye') return false;
        if (q && !(v.client_nom || '').toLowerCase().includes(q) && !(v.entreprise_nom || '').toLowerCase().includes(q) && !(v.rep_nom || '').toLowerCase().includes(q)) return false;
        return true;
      });

      updateSortIndicators('mes', ['date_vente', 'entreprise_nom', 'client_nom', 'rep_nom', 'valeur', 'commission_rep', 'statut_depot', 'statut_paiement'], mesVentesSort);
      filtered.sort((a, b) => {
        const col = mesVentesSort.col;
        const dir = mesVentesSort.dir;
        if (col === 'valeur' || col === 'commission_rep') {
          return compareSaleValues(a[col], b[col], dir, true);
        }
        if (col === 'date_vente') {
          const cmp = compareSaleValues(a.date_vente, b.date_vente, dir, false);
          if (cmp !== 0) return cmp;
          const cmpCreated = compareSaleValues(a.created_at || '', b.created_at || '', dir, false);
          if (cmpCreated !== 0) return cmpCreated;
          return compareSaleValues(a.id || '', b.id || '', dir, false);
        }
        return compareSaleValues(a[col], b[col], dir, false);
      });

      const yearSales = scopeFilteredAll.filter(v => getSaleStatsMonth(v).startsWith(currentYear));
      const isConfirmedOnly = commDisplayMode === 'confirme';

      // Comptabiliser chaque dossier uniquement dans sa période active
      const statsPeriodSales = filtered;
      const carriedInPeriodCount = statsPeriodSales.filter(v => v._carriedOverFrom).length;

      const filteredConfirmed = filtered.filter(v => v.statut_depot === 'depot_recu');
      const statsPeriodConfirmed = statsPeriodSales.filter(v => v.statut_depot === 'depot_recu');
      const yearConfirmed = yearSales.filter(v => v.statut_depot === 'depot_recu');

      const commPeriodVendu = statsPeriodSales.reduce((s, v) => s + Number(v.commission_rep), 0);
      const commPeriodConfirme = filteredConfirmed.reduce((s, v) => s + Number(v.commission_rep), 0);
      const activePeriodSales = isConfirmedOnly ? filteredConfirmed : statsPeriodSales;
      const commPeriod = isConfirmedOnly ? commPeriodConfirme : commPeriodVendu;
      const paidPeriod = activePeriodSales.filter(v => v.statut_paiement === 'paye').reduce((s, v) => s + Number(v.commission_rep), 0);

      const commYearVendu = yearSales.reduce((s, v) => s + Number(v.commission_rep), 0);
      const commYearConfirme = yearConfirmed.reduce((s, v) => s + Number(v.commission_rep), 0);
      const activeYearSales = isConfirmedOnly ? yearConfirmed : yearSales;
      const commYear = isConfirmedOnly ? commYearConfirme : commYearVendu;
      const paidYear = activeYearSales.filter(v => v.statut_paiement === 'paye').reduce((s, v) => s + Number(v.commission_rep), 0);

      const periodSalesTotal = statsPeriodSales.reduce((s, v) => s + Number(v.valeur), 0);
      const periodConfirmedTotal = statsPeriodConfirmed.reduce((s, v) => s + Number(v.valeur), 0);
      document.getElementById('kpi-rep-ventes-mois').textContent = fmtMoney(periodSalesTotal);
      document.getElementById('kpi-rep-count-mois').textContent = `${statsPeriodSales.length} dossier(s) (${statsPeriodConfirmed.length} confirmé(s)${carriedInPeriodCount > 0 ? ` • dont ${carriedInPeriodCount} transféré(s)` : ''})`;
      const elMoisConf = document.getElementById('kpi-rep-ventes-mois-confirme');
      if (elMoisConf) elMoisConf.textContent = fmtMoney(periodConfirmedTotal);

      updatePersonalGoalProgressBar(periodState, scopeFilter);

      document.getElementById('kpi-rep-comm-mois').textContent = fmtMoney(commPeriod);
      const elSubMois = document.getElementById('kpi-rep-comm-mois-sublabel');
      if (elSubMois) elSubMois.textContent = isConfirmedOnly ? 'Sur total confirmé (Dépôt reçu)' : 'Sur total vendu';
      const elMoisVendu = document.getElementById('kpi-rep-comm-mois-vendu');
      if (elMoisVendu) elMoisVendu.textContent = fmtMoney(commPeriodVendu);
      const elMoisConfComm = document.getElementById('kpi-rep-comm-mois-confirme');
      if (elMoisConfComm) elMoisConfComm.textContent = fmtMoney(commPeriodConfirme);
      document.getElementById('kpi-rep-paid-split').textContent = `Payé : ${fmtMoney(paidPeriod)} • À venir : ${fmtMoney(Math.max(0, commPeriod - paidPeriod))}`;

      document.getElementById('kpi-rep-ventes-annee').textContent = fmtMoney(yearSales.reduce((s, v) => s + Number(v.valeur), 0));
      const elAnneeConf = document.getElementById('kpi-rep-ventes-annee-confirme');
      if (elAnneeConf) elAnneeConf.textContent = fmtMoney(yearConfirmed.reduce((s, v) => s + Number(v.valeur), 0));

      document.getElementById('kpi-rep-comm-annee').textContent = fmtMoney(commYear);
      const elSubAnnee = document.getElementById('kpi-rep-comm-annee-sublabel');
      if (elSubAnnee) elSubAnnee.textContent = isConfirmedOnly ? 'Sur total confirmé (Dépôt reçu)' : 'Sur total vendu';
      const elAnneeVendu = document.getElementById('kpi-rep-comm-annee-vendu');
      if (elAnneeVendu) elAnneeVendu.textContent = fmtMoney(commYearVendu);
      const elAnneeConfComm = document.getElementById('kpi-rep-comm-annee-confirme');
      if (elAnneeConfComm) elAnneeConfComm.textContent = fmtMoney(commYearConfirme);
      document.getElementById('kpi-rep-annee-paid-split').textContent = `Payé : ${fmtMoney(paidYear)} • À venir : ${fmtMoney(Math.max(0, commYear - paidYear))}`;

      const tbody = document.getElementById('tbody-mes-ventes');
      const tbodyComm = document.getElementById('tbody-mes-commissions');
      const mobileCardsBox = document.getElementById('mobile-mes-ventes-cards');
      const mobileCommCardsBox = document.getElementById('mobile-mes-commissions-cards');
      if (!filtered.length) {
        tbody.innerHTML = `<tr><td colspan="${showRepCol ? 6 : 5}" class="py-12 text-center text-sm text-slate-400">Aucune vente enregistrée pour cette période.</td></tr>`;
        if (tbodyComm) {
          tbodyComm.innerHTML = `<tr><td colspan="${showRepCol ? 7 : 6}" class="py-12 text-center text-sm text-slate-400">Aucune commission enregistrée pour cette période.</td></tr>`;
        }
        if (mobileCardsBox) {
          mobileCardsBox.innerHTML = `<div class="bg-white rounded-2xl border border-slate-200/80 p-8 text-center text-xs text-slate-400">Aucune vente enregistrée pour cette période.</div>`;
        }
        if (mobileCommCardsBox) {
          mobileCommCardsBox.innerHTML = `<div class="bg-white rounded-2xl border border-slate-200/80 p-8 text-center text-xs text-slate-400">Aucune commission enregistrée pour cette période.</div>`;
        }
        applyMobilePersonalSplitLayout();
        return;
      }

      if (mobileCardsBox) {
        mobileCardsBox.innerHTML = filtered.map(v => {
          const isSuivi = (v.note || '').toUpperCase().includes('[SUIVI]');
          const locked = isMonthClosed(v.mois_commission);
          const canToggleKpi = !locked && (isAdmin || (v.statut_depot === 'a_confirmer' && v.statut_paiement !== 'paye'));
          const canDelete = !locked && v.statut_depot === 'a_confirmer' && v.statut_paiement !== 'paye';
          const proofSrc = getProofForSale(v);
          const cleanNote = cleanDisplaySaleNote(v.note);
          const hitRow = isHitSale(v);
          return `
            <div data-sale-id="${v.id}" class="bg-white rounded-2xl border ${hitRow ? 'border-rose-300 bg-rose-50/20' : (locked ? 'border-slate-300/90 bg-slate-50/40' : 'border-slate-200/80')} p-4 shadow-xs space-y-3">
              <div class="flex items-start justify-between gap-2">
                <div class="space-y-1 min-w-0">
                  <div class="flex items-center flex-wrap gap-1.5">
                    ${entBadge(v.entreprise_nom)}
                    ${locked ? `<span class="inline-flex items-center gap-1 px-2 py-0.5 text-[10px] font-extrabold rounded-md bg-slate-800 text-amber-300 border border-slate-700" title="Mois clôturé après l'envoi du relevé">🔒 Clôturée</span>` : ''}
                    ${hitBadgeHtml(v)}
                    ${canToggleKpi
                      ? `<button type="button" onclick="toggleVenteKpiType('${v.id}')" class="px-1.5 py-0.5 text-[10px] font-extrabold rounded-md border ${isSuivi ? 'bg-purple-50 text-purple-700 border-purple-200' : 'bg-sky-50 text-sky-700 border-sky-200'}">${isSuivi ? 'SUIVI' : 'OTS'}</button>`
                      : `<span class="px-1.5 py-0.5 text-[10px] font-extrabold rounded-md border ${isSuivi ? 'bg-purple-50 text-purple-700 border-purple-200' : 'bg-sky-50 text-sky-700 border-sky-200'}">${isSuivi ? 'SUIVI' : 'OTS'}</span>`
                    }
                    <span class="text-[11px] font-medium text-slate-400">${v.date_vente}</span>
                  </div>
                  <div class="font-extrabold text-sm text-slate-900 truncate">${esc(v.client_nom)}</div>
                  ${isAdmin ? `<div class="text-[11px] font-semibold text-indigo-600">Rep : ${esc(v.rep_nom || '—')}</div>` : ''}
                  ${cleanNote ? `<div class="text-[11px] text-slate-400 truncate">${esc(cleanNote)}</div>` : ''}
                  ${v._carriedOverFrom ? `<span class="inline-block px-1.5 py-0.5 text-[10px] font-bold rounded-md bg-amber-50 text-amber-800 border border-amber-200">↪ Transférée de ${v._carriedOverFrom}</span>` : ''}
                </div>
                <div class="text-right shrink-0 tabular-nums">
                  <div class="text-base font-extrabold text-slate-900">${fmtMoney(v.valeur)}</div>
                  <div class="text-[10px] font-semibold text-slate-400">Montant vente</div>
                </div>
              </div>
              <div class="flex items-center justify-between gap-2 pt-2 border-t border-slate-100">
                <div class="flex items-center flex-wrap gap-1.5">
                  <span class="inline-flex items-center gap-1 px-2 py-0.5 text-[11px] font-semibold rounded-full ${v.statut_depot === 'depot_recu' ? 'bg-emerald-50 text-emerald-700' : 'bg-amber-50 text-amber-700'}">
                    ${v.statut_depot === 'depot_recu' ? (isNoDepositSale(v) ? '✓ Sans dépôt requis' : '✓ Dépôt reçu') : '⏳ À confirmer'}
                  </span>
                  ${agingBadgeForDeposit(v)}
                </div>
                <div class="flex items-center gap-1.5">
                  ${proofSrc
                    ? `<button type="button" onclick="openProofModal('${v.id}')" class="px-2 py-1 text-[11px] font-bold rounded-lg bg-indigo-50 text-indigo-700">📷 Preuve</button>`
                    : `<label class="cursor-pointer px-2 py-1 text-[11px] font-semibold rounded-lg bg-slate-100 text-slate-700">+ 📷<input type="file" accept="image/*" class="sr-only" onchange="attachProofToExistingSale('${v.id}', this.files[0])" /></label>`
                  }
                  <button type="button" onclick="openSaleAuditModal('${v.id}')" class="px-2 py-1 text-[11px] font-semibold rounded-lg bg-slate-100 text-slate-700">🕒</button>
                  ${locked
                    ? `<button type="button" onclick="toast('🔒 Vente verrouillée : le relevé mensuel a été envoyé et le mois est clôturé.')" title="Vente verrouillée (mois clôturé)" class="px-2 py-1 text-[11px] font-extrabold rounded-lg bg-slate-800 text-amber-300">🔒</button>`
                    : (canDelete ? `<button type="button" onclick="deleteVente('${v.id}')" class="px-2 py-1 text-[11px] font-semibold rounded-lg bg-rose-50 text-rose-600">Suppr.</button>` : '')
                  }
                </div>
              </div>
            </div>
          `;
        }).join('');
      }

      if (mobileCommCardsBox) {
        mobileCommCardsBox.innerHTML = filtered.map(v => {
          const baseR = getRateForSale(v.representant_id, v.entreprise_id, v.client_id);
          const diff = baseR ? Math.max(0, Number(v.taux_rep_applique) - Number(baseR.taux)) : 0.01;
          const myCommVal = Number(v.commission_rep || 0);
          const repCommVal = Number(v._repCommAmount || 0);
          const showTeamRepSub = isAdmin && !isInspectingRep && String(v.representant_id) !== String(currentProfile.id) && repCommVal !== 0;
          return `
            <div class="bg-white rounded-2xl border border-slate-200/80 p-4 shadow-xs space-y-3">
              <div class="space-y-1 min-w-0">
                <div class="flex items-center flex-wrap gap-1.5">
                  ${entBadge(v.entreprise_nom)}
                  ${isMonthClosed(v.mois_commission) ? `<span class="inline-flex items-center gap-1 px-1.5 py-0.5 text-[10px] font-extrabold rounded-md bg-slate-800 text-amber-300">🔒 Clôturée</span>` : ''}
                  <span class="text-[11px] font-medium text-slate-400">${v.date_vente} (${v.mois_commission})</span>
                </div>
                <div class="font-extrabold text-sm text-slate-900 truncate">${esc(v.client_nom)}</div>
                ${isAdmin ? `<div class="text-[11px] font-semibold text-indigo-600">Rep : ${esc(v.rep_nom || '—')}</div>` : ''}
              </div>

              <div class="grid grid-cols-2 gap-2.5 p-2.5 rounded-xl bg-slate-50/90 border border-slate-200/70 tabular-nums">
                <div>
                  <div class="text-[10px] font-bold uppercase tracking-wider text-slate-400">Montant vente</div>
                  <div class="text-sm font-extrabold text-slate-800 mt-0.5">${fmtMoney(v.valeur)}</div>
                </div>
                <div class="text-right border-l border-slate-200/80 pl-2.5">
                  <div class="text-[10px] font-bold uppercase tracking-wider text-emerald-700">Ma commission</div>
                  <div class="text-base font-extrabold mt-0.5 ${myCommVal < 0 ? 'text-rose-600' : 'text-emerald-600'}">${myCommVal < 0 ? fmtMoney(myCommVal) : `+ ${fmtMoney(myCommVal)}`}</div>
                  ${showTeamRepSub ? `<div class="text-[10px] font-semibold text-slate-500">Comm. Rep : ${fmtMoney(repCommVal)}</div>` : ''}
                  ${isHitSale(v) ? hitBadgeHtml(v) : (v._bonusApplied ? `<span class="inline-block mt-0.5 px-1.5 py-0.2 text-[10px] font-extrabold rounded bg-emerald-100 text-emerald-800">+${formatBonusPctShort(diff || 0.01)} bonus</span>` : '')}
                </div>
              </div>

              <div class="flex items-center justify-between gap-2 pt-1">
                <div class="flex items-center flex-wrap gap-1.5">
                  <span class="inline-flex items-center gap-1 px-2 py-0.5 text-[11px] font-semibold rounded-full ${v.statut_depot === 'depot_recu' ? 'bg-emerald-50 text-emerald-700' : 'bg-amber-50 text-amber-700'}">
                    ${v.statut_depot === 'depot_recu' ? (isNoDepositSale(v) ? 'Sans dépôt requis' : 'Dépôt reçu') : 'À confirmer'}
                  </span>
                  <span class="px-2 py-0.5 text-[11px] font-semibold rounded-full ${v.statut_paiement === 'paye' ? 'bg-indigo-100 text-indigo-800' : 'bg-slate-100 text-slate-600'}">
                    ${v.statut_paiement === 'paye' ? '✓ Payé' : 'À recevoir'}
                  </span>
                </div>
                <button type="button" onclick="openSaleAuditModal('${v.id}')" class="px-2 py-1 text-[11px] font-semibold rounded-lg bg-slate-100 text-slate-700">🕒 Historique</button>
              </div>
            </div>
          `;
        }).join('');
      }
      applyMobilePersonalSplitLayout();

      tbody.innerHTML = filtered.map(v => {
        const isSuivi = (v.note || '').toUpperCase().includes('[SUIVI]');
        const locked = isMonthClosed(v.mois_commission);
        const canToggleKpi = !locked && (isAdmin || (v.statut_depot === 'a_confirmer' && v.statut_paiement !== 'paye'));
        const cleanNote = cleanDisplaySaleNote(v.note);
        const proofSrc = getProofForSale(v);
        const canDelete = !locked && v.statut_depot === 'a_confirmer' && v.statut_paiement !== 'paye';
        const showMoisDiff = v.mois_commission && v.date_vente && !v.date_vente.startsWith(v.mois_commission);
        return `
        <tr data-sale-id="${v.id}" class="${locked ? 'bg-slate-50/60' : 'hover:bg-slate-50/90'} transition">
          <td class="py-3.5 px-4 whitespace-nowrap align-middle">
            <div class="font-semibold text-slate-800 tabular-nums">${v.date_vente}</div>
            ${showMoisDiff ? `<div class="text-[11px] text-slate-400">Paie ${v.mois_commission}</div>` : ''}
            ${v._carriedOverFrom ? `<span class="mt-0.5 inline-block px-1.5 py-0.5 text-[10px] font-bold rounded-md bg-amber-50 text-amber-800 border border-amber-200" title="Reportée de ${v._carriedOverFrom}">↪ ${v._carriedOverFrom}</span>` : ''}
          </td>
          <td class="py-3.5 px-4 align-middle">
            <div class="flex items-center flex-wrap gap-2">
              ${entBadge(v.entreprise_nom)}
              <span class="font-bold text-slate-900">${esc(v.client_nom)}</span>
              ${locked ? `<span class="inline-flex items-center gap-1 px-2 py-0.5 text-[10px] font-extrabold rounded-md bg-slate-800 text-amber-300 border border-slate-700" title="Mois clôturé après l'envoi du relevé">🔒 Clôturée</span>` : ''}
            </div>
            <div class="flex items-center flex-wrap gap-1.5 mt-1">
              ${canToggleKpi
                ? `<button type="button" onclick="toggleVenteKpiType('${v.id}')" title="Basculer OTS / SUIVI"
                    class="px-1.5 py-0.2 text-[10px] font-extrabold rounded border transition ${isSuivi ? 'bg-purple-50 text-purple-700 border-purple-200 hover:bg-purple-100' : 'bg-sky-50 text-sky-700 border-sky-200 hover:bg-sky-100'}">
                    ${isSuivi ? 'SUIVI' : 'OTS'}
                  </button>`
                : `<span class="px-1.5 py-0.2 text-[10px] font-extrabold rounded border ${isSuivi ? 'bg-purple-50 text-purple-700 border-purple-200' : 'bg-sky-50 text-sky-700 border-sky-200'}">
                    ${isSuivi ? 'SUIVI' : 'OTS'}
                  </span>`
              }
              ${hitBadgeHtml(v)}
              ${cleanNote ? `<span class="text-[11px] text-slate-400 truncate max-w-[230px]">${esc(cleanNote)}</span>` : ''}
            </div>
          </td>
          ${showRepCol ? `<td class="py-3.5 px-4 text-xs font-semibold text-slate-700 whitespace-nowrap align-middle">${esc(v.rep_nom || '—')}</td>` : ''}
          <td class="py-3.5 px-4 text-right font-bold text-slate-900 tabular-nums whitespace-nowrap align-middle">${fmtMoney(v.valeur)}</td>
          <td class="py-3.5 px-4 whitespace-nowrap align-middle">
            <div class="flex items-center gap-1.5">
              <span class="inline-flex items-center gap-1.5 px-2.5 py-1 text-xs font-semibold rounded-full ${v.statut_depot === 'depot_recu' ? 'bg-emerald-50 text-emerald-700' : 'bg-amber-50 text-amber-700'}">
                <span class="w-1.5 h-1.5 rounded-full ${v.statut_depot === 'depot_recu' ? 'bg-emerald-500' : 'bg-amber-500'}"></span>
                ${v.statut_depot === 'depot_recu' ? (isNoDepositSale(v) ? 'Sans dépôt requis' : 'Dépôt reçu') : 'À confirmer'}
              </span>
              ${agingBadgeForDeposit(v)}
            </div>
          </td>
          <td class="py-3.5 px-4 text-right whitespace-nowrap align-middle">
            <div class="inline-flex items-center justify-end gap-1.5">
              ${proofSrc
                ? `<button type="button" onclick="openProofModal('${v.id}')" title="Voir la preuve photo" class="px-2 py-1 text-xs font-bold rounded-lg bg-indigo-50 text-indigo-700 hover:bg-indigo-100 border border-indigo-200/70 transition">📷</button>`
                : `<label title="Ajouter une preuve photo" class="cursor-pointer px-2 py-1 text-xs font-semibold rounded-lg bg-slate-100 text-slate-600 hover:bg-indigo-50 hover:text-indigo-700 transition">+📷<input type="file" accept="image/*" class="sr-only" onchange="attachProofToExistingSale('${v.id}', this.files[0])" /></label>`
              }
              <button type="button" onclick="openSaleAuditModal('${v.id}')" title="Historique du dossier" class="px-2 py-1 text-xs rounded-lg bg-slate-100 text-slate-600 hover:text-indigo-700 hover:bg-indigo-50 transition">🕒</button>
              ${locked
                ? `<button type="button" onclick="toast('🔒 Vente verrouillée : le relevé mensuel a été envoyé et le mois est clôturé.')" title="Vente verrouillée (mois clôturé)" class="px-2 py-1 text-xs font-extrabold rounded-lg bg-slate-800 text-amber-300 hover:bg-slate-700 transition">🔒</button>`
                : (canDelete
                  ? `<button type="button" onclick="deleteVente('${v.id}')" title="Supprimer" class="px-2 py-1 text-xs rounded-lg bg-rose-50 text-rose-600 hover:bg-rose-100 transition">🗑️</button>`
                  : '')
              }
            </div>
          </td>
        </tr>
      `;
      }).join('');

      if (tbodyComm) {
        tbodyComm.innerHTML = filtered.map(v => {
          const locked = isMonthClosed(v.mois_commission);
          const baseR = getRateForSale(v.representant_id, v.entreprise_id, v.client_id);
          const diff = baseR ? Math.max(0, Number(v.taux_rep_applique) - Number(baseR.taux)) : 0.01;
          return `
          <tr class="${locked ? 'bg-slate-50/60' : 'hover:bg-slate-50/90'} transition">
            <td class="py-3.5 px-4 whitespace-nowrap align-middle">
              <div class="font-semibold text-slate-800 tabular-nums">${v.date_vente}</div>
              <div class="text-[11px] text-slate-400">Paie ${v.mois_commission}</div>
            </td>
            <td class="py-3.5 px-4 align-middle">
              <div class="flex items-center flex-wrap gap-2">
                ${entBadge(v.entreprise_nom)}
                <span class="font-bold text-slate-900">${esc(v.client_nom)}</span>
                ${locked ? `<span class="inline-flex items-center gap-1 px-2 py-0.5 text-[10px] font-extrabold rounded-md bg-slate-800 text-amber-300 border border-slate-700">🔒 Clôturée</span>` : ''}
              </div>
            </td>
            ${showRepCol ? `<td class="py-3.5 px-4 text-xs font-semibold text-slate-700 whitespace-nowrap align-middle">${esc(v.rep_nom || '—')}</td>` : ''}
            <td class="py-3.5 px-4 text-right font-semibold text-slate-700 tabular-nums whitespace-nowrap align-middle">${fmtMoney(v.valeur)}</td>
            <td class="py-3.5 px-4 text-right tabular-nums whitespace-nowrap align-middle">
              <div class="font-extrabold ${Number(v.commission_rep) < 0 ? 'text-rose-600' : 'text-emerald-600'}">${Number(v.commission_rep) < 0 ? fmtMoney(v.commission_rep) : `+ ${fmtMoney(v.commission_rep)}`}</div>
              ${(isAdmin && !isInspectingRep && String(v.representant_id) !== String(currentProfile.id) && Number(v._repCommAmount || 0) !== 0) ? `<div class="text-[11px] font-semibold text-slate-500">Comm. Rep : ${fmtMoney(v._repCommAmount)}</div>` : ''}
              ${isHitSale(v) ? hitBadgeHtml(v) : (v._bonusApplied ? `<span class="inline-block mt-0.5 px-1.5 py-0.2 text-[10px] font-extrabold rounded bg-emerald-100 text-emerald-800">+${formatBonusPctShort(diff || 0.01)} bonus</span>` : '')}
            </td>
            <td class="py-3.5 px-4 whitespace-nowrap align-middle">
              <div class="flex items-center gap-1.5">
                <span class="inline-flex items-center gap-1.5 px-2.5 py-1 text-xs font-semibold rounded-full ${v.statut_depot === 'depot_recu' ? 'bg-emerald-50 text-emerald-700' : 'bg-amber-50 text-amber-700'}">
                  ${v.statut_depot === 'depot_recu' ? (isNoDepositSale(v) ? 'Sans dépôt requis' : 'Dépôt reçu') : 'À confirmer'}
                </span>
                <span class="px-2 py-0.5 text-[11px] font-semibold rounded-full ${v.statut_paiement === 'paye' ? 'bg-indigo-50 text-indigo-700' : 'bg-slate-100 text-slate-500'}">
                  ${v.statut_paiement === 'paye' ? '✓ Payé' : 'À recevoir'}
                </span>
              </div>
            </td>
            <td class="py-3.5 px-4 text-right whitespace-nowrap align-middle">
              <button type="button" onclick="openSaleAuditModal('${v.id}')" title="Historique du dossier" class="px-2.5 py-1 text-xs font-semibold rounded-lg bg-slate-100 text-slate-600 hover:text-indigo-700 hover:bg-indigo-50 transition">🕒 Historique</button>
            </td>
          </tr>
        `;
        }).join('');
      }
    }

    window.setRankPeriodQuick = (mode) => {
      const sel = document.getElementById('filter-rank-period');
      if (!sel) return;
      sel.value = mode;
      syncPeriodInputsVisibility('filter-rank', false);
      renderClassement();
    };

    const REP_AVATAR_GRADIENTS = [
      'from-indigo-600 to-blue-600',
      'from-emerald-600 to-teal-600',
      'from-violet-600 to-purple-600',
      'from-sky-600 to-cyan-600',
      'from-rose-600 to-pink-600',
      'from-amber-600 to-orange-600'
    ];

    function getRepAvatarGradient(name) {
      let h = 0;
      const s = String(name || '');
      for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) | 0;
      return REP_AVATAR_GRADIENTS[Math.abs(h) % REP_AVATAR_GRADIENTS.length];
    }

    function renderClassement() {
      const rankPeriod = readPeriodFilterState('filter-rank');
      const pillEl = document.getElementById('classement-period-pill');
      if (pillEl) pillEl.textContent = rankPeriod.label || 'Ce mois-ci';

      document.querySelectorAll('.rank-pill-btn').forEach(btn => {
        const isAct = btn.dataset.rankMode === rankPeriod.mode;
        btn.className = isAct
          ? 'rank-pill-btn shrink-0 px-3 py-2 sm:py-1.5 text-xs font-extrabold rounded-xl bg-indigo-600 text-white shadow-2xs transition whitespace-nowrap cursor-pointer'
          : 'rank-pill-btn shrink-0 px-3 py-2 sm:py-1.5 text-xs font-bold rounded-xl bg-slate-100 text-slate-600 hover:bg-slate-200 transition whitespace-nowrap cursor-pointer';
      });

      const totalsByRep = new Map();
      const ensureRepEntry = (id, nom) => {
        const key = String(id || nom);
        if (!totalsByRep.has(key)) {
          totalsByRep.set(key, {
            id: key,
            nom,
            total: 0,
            confirmedTotal: 0,
            dealsCount: 0,
            otsCount: 0,
            suiviCount: 0,
            bestSale: 0,
            byEnt: new Map(),
            monthsCount: new Set()
          });
        }
        return totalsByRep.get(key);
      };

      const isSubMonthMode = rankPeriod.mode === 'this_week' || rankPeriod.mode === 'specific_date' || rankPeriod.mode === 'custom_range';
      const hasFullVentesAccess = Boolean(state.ventes && state.ventes.length && currentProfile?.role === 'admin');

      if (hasFullVentesAccess) {
        state.representants.forEach(r => ensureRepEntry(r.id, r.nom));
        state.ventes.forEach(v => {
          const repObj = state.representants.find(r => String(r.id) === String(v.representant_id));
          const archRepNom = !repObj ? resolveSaleArchivedRepName(v) : '';
          const resolvedRepId = repObj ? repObj.id : (archRepNom ? `arch_rep_${archRepNom.toLowerCase()}` : '');
          const resolvedRepNom = repObj ? repObj.nom : archRepNom;
          if (!resolvedRepNom) return;
          const entry = ensureRepEntry(resolvedRepId, resolvedRepNom);
          if (matchesSalePeriod(v, rankPeriod, true)) {
            const val = Number(v.valeur || 0);
            const sYm = getSaleStatsMonth(v);
            entry.total += val;
            if (v.statut_depot === 'depot_recu') entry.confirmedTotal += val;
            entry.dealsCount += 1;
            if (val > entry.bestSale) entry.bestSale = val;
            if ((v.note || '').toUpperCase().includes('[SUIVI]')) entry.suiviCount += 1;
            else entry.otsCount += 1;
            const entName = state.entreprises.find(e => String(e.id) === String(v.entreprise_id))?.nom || resolveSaleArchivedEntName(v) || v.entreprise_nom || '';
            if (entName) entry.byEnt.set(entName, (entry.byEnt.get(entName) || 0) + val);
            if (sYm) entry.monthsCount.add(sYm);
          }
        });
      } else {
        // Pour un représentant : combiner ses propres ventes détaillées (`state.mesVentes`) + vue classement + KPI store pour les collègues
        const knownRepsMap = new Map();
        (state.representants || []).forEach(r => {
          if (r && r.nom) knownRepsMap.set(resolveKpiRepName(r.nom), { id: r.id, nom: r.nom });
        });
        ALL_KPI_REPS.forEach((n, i) => {
          const norm = resolveKpiRepName(n);
          if (norm && !knownRepsMap.has(norm)) {
            knownRepsMap.set(norm, { id: `arch_rep_${norm.toLowerCase()}`, nom: n });
          }
        });
        const knownReps = Array.from(knownRepsMap.values());
        knownReps.forEach(r => ensureRepEntry(r.id, r.nom));

        if (isSubMonthMode) {
          // Mode sous-mensuel (cette semaine / date sur mesure) pour un représentant :
          // 1) Compter ses propres ventes exactes depuis `state.mesVentes`
          const myIdStr = String(currentProfile?.id || '');
          (state.mesVentes || []).forEach(v => {
            if (!matchesSalePeriod(v, rankPeriod, true)) return;
            const entry = ensureRepEntry(v.representant_id || myIdStr, currentProfile?.nom || 'Moi');
            const val = Number(v.valeur || 0);
            entry.total += val;
            if (v.statut_depot === 'depot_recu') entry.confirmedTotal += val;
            entry.dealsCount += 1;
            if (val > entry.bestSale) entry.bestSale = val;
            if ((v.note || '').toUpperCase().includes('[SUIVI]')) entry.suiviCount += 1;
            else entry.otsCount += 1;
            const entN = v.entreprise_nom || state.entreprises.find(e => String(e.id) === String(v.entreprise_id))?.nom || '';
            if (entN) entry.byEnt.set(entN, (entry.byEnt.get(entN) || 0) + val);
          });
          // 2) Ajouter les ventes des collègues publiées dans le canal d'équipe sur cette période
          if (typeof teamChatMessages !== 'undefined' && Array.isArray(teamChatMessages)) {
            teamChatMessages.forEach(m => {
              if (!m || m.type !== 'sale' || !m.repNom) return;
              if (resolveKpiRepName(m.repNom) === resolveKpiRepName(currentProfile?.nom || '')) return;
              const dStr = (m.dateVente || m.createdAt || '').slice(0, 10);
              if (!dStr) return;
              const fakeSale = { date_vente: dStr, mois_commission: dStr.slice(0, 7) };
              if (!matchesSalePeriod(fakeSale, rankPeriod, true)) return;
              const repObj = state.representants.find(r => resolveKpiRepName(r.nom) === resolveKpiRepName(m.repNom));
              const entry = ensureRepEntry(repObj?.id || m.repNom, repObj?.nom || m.repNom);
              const val = Number(m.valeur || 0);
              entry.total += val;
              entry.dealsCount += 1;
              if (val > entry.bestSale) entry.bestSale = val;
              if ((m.typeKpi || '').toUpperCase() === 'SUIVI') entry.suiviCount += 1;
              else entry.otsCount += 1;
              if (m.entrepriseNom) entry.byEnt.set(m.entrepriseNom, (entry.byEnt.get(m.entrepriseNom) || 0) + val);
            });
          }
        } else {
          state.classement.forEach(row => {
            const normR = resolveKpiRepName(row.representant_nom || '');
            const existingByName = normR ? Array.from(totalsByRep.values()).find(e => resolveKpiRepName(e.nom) === normR) : null;
            const entry = existingByName || ensureRepEntry(row.representant_id || (normR ? `arch_rep_${normR.toLowerCase()}` : ''), row.representant_nom || normR);
            if (!entry) return;
            const fakeRow = { mois_commission: row.mois_commission, date_vente: `${row.mois_commission}-01` };
            if (matchesSalePeriod(fakeRow, rankPeriod, false)) {
              const adjTotal = Math.max(0, Number(row.total_ventes || 0));
              entry.total += adjTotal;
              entry.dealsCount += Number(row.nb_ventes || row.nombre_ventes || 0);
              if (row.mois_commission) entry.monthsCount.add(row.mois_commission);
            }
          });
        }

        // Enrichir TOUS les représentants avec les détails disponibles (state.ventes + teamChatMessages + kpiStore + state.mesVentes)
        const detailedSalesPool = new Map();
        (state.mesVentes || []).forEach(v => {
          if (v && v.id) detailedSalesPool.set(String(v.id), v);
        });
        (state.ventes || []).forEach(v => {
          if (v && v.id && !detailedSalesPool.has(String(v.id))) detailedSalesPool.set(String(v.id), v);
        });

        // 1) D'abord utiliser les ventes individuelles accessibles
        const perRepDetailed = new Map();
        detailedSalesPool.forEach(v => {
          if (!matchesSalePeriod(v, rankPeriod, true)) return;
          const rObj = state.representants.find(r => String(r.id) === String(v.representant_id));
          const rName = rObj?.nom || resolveSaleArchivedRepName(v) || v.rep_nom || '';
          const normRName = resolveKpiRepName(rName);
          if (!normRName) return;
          const entry = Array.from(totalsByRep.values()).find(e => resolveKpiRepName(e.nom) === normRName)
            || ensureRepEntry(rObj?.id || `arch_rep_${normRName.toLowerCase()}`, normRName);
          if (!entry) return;
          if (!perRepDetailed.has(entry.id)) {
            perRepDetailed.set(entry.id, { cnt: 0, conf: 0, best: 0, ots: 0, sui: 0, byEnt: new Map(), vol: 0 });
          }
          const acc = perRepDetailed.get(entry.id);
          const val = Number(v.valeur || 0);
          acc.cnt += 1;
          acc.vol += val;
          if (v.statut_depot === 'depot_recu') acc.conf += val;
          if (val > acc.best) acc.best = val;
          if ((v.note || '').toUpperCase().includes('[SUIVI]')) acc.sui += 1;
          else acc.ots += 1;
          const entN = v.entreprise_nom || state.entreprises.find(e => String(e.id) === String(v.entreprise_id))?.nom || '';
          if (entN) acc.byEnt.set(entN, (acc.byEnt.get(entN) || 0) + val);
        });

        // 2) Compléter avec les annonces de ventes du canal d'équipe (`teamChatMessages`) pour les collègues dont les lignes `ventes` sont masquées par RLS
        if (typeof teamChatMessages !== 'undefined' && Array.isArray(teamChatMessages)) {
          teamChatMessages.forEach(m => {
            if (!m || m.type !== 'sale' || !m.repNom) return;
            const dStr = (m.dateVente || m.createdAt || '').slice(0, 10);
            if (!dStr) return;
            const fakeSale = { date_vente: dStr, mois_commission: dStr.slice(0, 7) };
            if (!matchesSalePeriod(fakeSale, rankPeriod, true)) return;
            const entry = Array.from(totalsByRep.values()).find(e => resolveKpiRepName(e.nom) === resolveKpiRepName(m.repNom));
            if (!entry) return;
            // Si on a déjà les ventes individuelles de ce rep (ex: soi-même), ne pas recompter
            if (String(entry.id) === String(currentProfile?.id)) return;
            if (!perRepDetailed.has(entry.id)) {
              perRepDetailed.set(entry.id, { cnt: 0, conf: 0, best: 0, ots: 0, sui: 0, byEnt: new Map(), vol: 0 });
            }
            const acc = perRepDetailed.get(entry.id);
            const val = Number(m.valeur || 0);
            acc.cnt += 1;
            acc.vol += val;
            if (val > acc.best) acc.best = val;
            if ((m.typeKpi || '').toUpperCase() === 'SUIVI') acc.sui += 1;
            else acc.ots += 1;
            if (m.entrepriseNom) acc.byEnt.set(m.entrepriseNom, (acc.byEnt.get(m.entrepriseNom) || 0) + val);
          });
        }

        // 3) Compléter aussi OTS / Suivi / Entreprises depuis `kpiStore` si disponible sur la période mensuelle
        if (!isSubMonthMode) {
          Object.entries(kpiStore || {}).forEach(([ym, mObj]) => {
            if (!isValidYearMonth(ym) || !mObj?.reps) return;
            const fakeRow = { mois_commission: ym, date_vente: `${ym}-01` };
            if (!matchesSalePeriod(fakeRow, rankPeriod, false)) return;
            totalsByRep.forEach(entry => {
              const rn = resolveKpiRepName(entry.nom);
              const rKpi = mObj.reps[rn];
              if (!rKpi) return;
              const kOts = Number(rKpi['Ventes OTS']?.cumul || 0);
              const kSui = Number(rKpi['Ventes SUIVI']?.cumul || 0);
              if (kOts + kSui > 0) {
                entry.otsCount += kOts;
                entry.suiviCount += kSui;
                entry.dealsCount = Math.max(entry.dealsCount, entry.otsCount + entry.suiviCount);
              }
              Object.entries(mObj.entreprises || {}).forEach(([entName, eObj]) => {
                const eVal = Number(eObj?.reps?.[rn]?.cumul || 0);
                if (eVal > 0) {
                  entry.byEnt.set(entName, Math.max(entry.byEnt.get(entName) || 0, eVal));
                }
              });
            });
          });
        }

        perRepDetailed.forEach((acc, repId) => {
          const entry = totalsByRep.get(repId);
          if (!entry) return;
          if (acc.cnt > 0) {
            entry.dealsCount = Math.max(entry.dealsCount, acc.cnt);
            if (acc.best > entry.bestSale) entry.bestSale = acc.best;
            if (entry.otsCount === 0 && entry.suiviCount === 0) {
              entry.otsCount = acc.ots;
              entry.suiviCount = acc.sui;
            }
            if (entry.byEnt.size === 0 && acc.byEnt.size > 0) {
              entry.byEnt = acc.byEnt;
            }
            if (acc.conf > 0) entry.confirmedTotal = acc.conf;
            if (acc.vol > entry.total) entry.total = acc.vol;
          }
        });
        if (!isSubMonthMode) {
          totalsByRep.forEach(entry => {
            const rn = resolveKpiRepName(entry.nom);
            if (!rn) return;
            let kpiPeriodDol = 0;
            Object.entries(kpiStore || {}).forEach(([ym, mObj]) => {
              if (!isValidYearMonth(ym) || !mObj?.reps?.[rn]) return;
              const fakeRow = { mois_commission: ym, date_vente: `${ym}-01` };
              if (matchesSalePeriod(fakeRow, rankPeriod, false)) {
                kpiPeriodDol += Number(mObj.reps[rn]['Vente $']?.cumul || 0);
              }
            });
            if (kpiPeriodDol > entry.total) entry.total = kpiPeriodDol;
          });
        }
      }

      const activeYm = rankPeriod.singleMonth || (rankPeriod.mode === 'custom_range' && rankPeriod.startVal ? rankPeriod.startVal.slice(0, 7) : currentMonthStr);
      const bizDaysM = Math.max(1, countBusinessDaysInMonth(activeYm));
      const mult = rankPeriod.mode === 'custom_range'
        ? computeCustomRangeMonthFactor(rankPeriod.startVal, rankPeriod.endVal)
        : (rankPeriod.mode === 'this_week'
          ? (5 / bizDaysM)
          : (rankPeriod.mode === 'specific_date'
            ? (1 / bizDaysM)
            : (rankPeriod.mode === 'this_quarter'
              ? 3
              : (rankPeriod.mode === 'this_year'
                ? 12
                : (rankPeriod.mode === 'all' ? Math.max(1, Object.keys(kpiStore).length || 1) : 1)))));

      const rankCoveredMonths = Object.keys(kpiStore || {}).filter(ym => {
        if (!isValidYearMonth(ym)) return false;
        const fakeRow = { mois_commission: ym, date_vente: `${ym}-01` };
        return matchesSalePeriod(fakeRow, rankPeriod, false);
      });
      const enriched = Array.from(totalsByRep.values()).map(r => {
        let repTarget = 0;
        if (!isSubMonthMode && rankCoveredMonths.length > 0 && (rankPeriod.mode === 'this_year' || rankPeriod.mode === 'this_quarter' || rankPeriod.mode === 'all')) {
          repTarget = Math.round(rankCoveredMonths.reduce((s, ym) => s + getSalesTargetForRep(r.nom, ym), 0));
        } else {
          const unitTarget = getSalesTargetForRep(r.nom, activeYm);
          repTarget = Math.round(unitTarget * mult);
        }
        const targetPct = repTarget > 0 ? Math.round((r.total / repTarget) * 100) : 0;
        const remaining = Math.max(0, repTarget - r.total);
        const avgTicket = r.dealsCount > 0 ? (r.total / r.dealsCount) : 0;
        const topEntEntry = Array.from(r.byEnt.entries()).sort((a, b) => b[1] - a[1])[0];
        return {
          ...r,
          repTarget,
          targetPct,
          remaining,
          avgTicket,
          topEntName: topEntEntry ? topEntEntry[0] : '',
          topEntVal: topEntEntry ? topEntEntry[1] : 0
        };
      }).filter(r => !isRepArchived(r.nom) && !isRepArchived(r.id) || r.total > 0 || r.dealsCount > 0).sort((a, b) => b.total - a.total);

      const teamTotalVol = enriched.reduce((s, r) => s + r.total, 0);
      const teamTotalTarget = enriched.reduce((s, r) => s + r.repTarget, 0);
      const teamTotalDeals = enriched.reduce((s, r) => s + r.dealsCount, 0);
      const teamAvgTicket = teamTotalDeals > 0 ? (teamTotalVol / teamTotalDeals) : 0;
      const teamPct = teamTotalTarget > 0 ? Math.round((teamTotalVol / teamTotalTarget) * 100) : 0;
      const unlockedRepsCount = enriched.filter(r => r.targetPct >= 100).length;

      // 1. BANDEAU OBJECTIF COLLECTIF DE L'ÉQUIPE
      const heroBox = document.getElementById('classement-team-hero');
      if (heroBox) {
        heroBox.innerHTML = `
          <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div>
              <div class="flex items-center gap-2">
                <span class="px-2.5 py-0.5 text-[10px] font-extrabold uppercase tracking-wider rounded-full bg-amber-400/20 text-amber-300 border border-amber-300/30">Objectif Collectif</span>
                <span class="text-xs font-semibold text-indigo-200">${esc(rankPeriod.label || 'Période active')}</span>
              </div>
              <div class="text-2xl sm:text-3xl font-extrabold text-white mt-1.5 tabular-nums">
                ${fmtMoney(teamTotalVol)}
                <span class="text-sm font-semibold text-indigo-300">/ ${fmtMoney0(teamTotalTarget)}</span>
              </div>
            </div>
            <div class="sm:text-right">
              <span class="inline-block px-3 py-1 rounded-xl text-sm font-extrabold tabular-nums ${teamPct >= 100 ? 'bg-emerald-500/25 text-emerald-300 border border-emerald-400/40' : 'bg-white/10 text-amber-300 border border-white/15'}">
                ${teamPct >= 100 ? `🔥 ${teamPct} % — Objectif équipe atteint !` : `${teamPct} % de l'objectif équipe`}
              </span>
            </div>
          </div>

          <div class="w-full h-3 bg-white/10 rounded-full overflow-hidden p-0.5 border border-white/15">
            <div class="h-full rounded-full transition-all duration-700 ${teamPct >= 100 ? 'bg-gradient-to-r from-emerald-400 to-teal-300' : 'bg-gradient-to-r from-indigo-400 via-sky-400 to-emerald-400'}" style="width: ${Math.min(100, teamPct)}%"></div>
          </div>

          <div class="grid grid-cols-2 sm:grid-cols-4 gap-3 pt-1">
            <div class="bg-white/10 rounded-2xl p-3 border border-white/10">
              <div class="text-[11px] text-indigo-200 font-semibold">Volume équipe</div>
              <div class="text-base sm:text-lg font-extrabold text-emerald-300 tabular-nums mt-0.5">${fmtMoney0(teamTotalVol)}</div>
            </div>
            <div class="bg-white/10 rounded-2xl p-3 border border-white/10">
              <div class="text-[11px] text-indigo-200 font-semibold">Contrats signés</div>
              <div class="text-base sm:text-lg font-extrabold text-white tabular-nums mt-0.5">${teamTotalDeals > 0 ? `${teamTotalDeals} dossier${teamTotalDeals > 1 ? 's' : ''}` : '—'}</div>
            </div>
            <div class="bg-white/10 rounded-2xl p-3 border border-white/10">
              <div class="text-[11px] text-indigo-200 font-semibold">Panier moyen</div>
              <div class="text-base sm:text-lg font-extrabold text-sky-300 tabular-nums mt-0.5">${teamAvgTicket > 0 ? fmtMoney0(teamAvgTicket) : '—'}</div>
            </div>
            <div class="bg-white/10 rounded-2xl p-3 border border-white/10">
              <div class="text-[11px] text-indigo-200 font-semibold">Objectifs atteints</div>
              <div class="text-base sm:text-lg font-extrabold text-amber-300 tabular-nums mt-0.5">${unlockedRepsCount} / ${enriched.length} reps</div>
            </div>
          </div>
        `;
      }

      // 2. PODIUM TOP 3 (Sur mobile : #1 en haut, #2 puis #3 ; sur desktop : #2 à gauche, #1 au centre surélevé, #3 à droite)
      const podiumBox = document.getElementById('classement-podium');
      if (podiumBox) {
        const top3 = enriched.slice(0, 3);
        const podiumOrder = top3.length === 3
          ? [
              { r: top3[1], rank: 2, orderCls: 'order-2 md:order-1' },
              { r: top3[0], rank: 1, orderCls: 'order-1 md:order-2' },
              { r: top3[2], rank: 3, orderCls: 'order-3 md:order-3' }
            ]
          : top3.map((r, i) => ({ r, rank: i + 1, orderCls: `order-${i + 1}` }));

        podiumBox.innerHTML = podiumOrder.map(({ r, rank, orderCls }, pIdx) => {
          const isFirst = rank === 1;
          const isSecond = rank === 2;
          const isMe = r.nom === currentProfile.nom;
          const sharePct = teamTotalVol > 0 ? Math.round((r.total / teamTotalVol) * 100) : 0;
          const crownOrMedal = isFirst ? '👑' : (isSecond ? '🥈' : '🥉');
          const podiumCls = isFirst
            ? 'rank-podium-gold border-2 md:pb-7'
            : (isSecond ? 'rank-podium-silver border' : 'rank-podium-bronze border');
          const badgeStyle = isFirst
            ? 'bg-amber-400 text-slate-950 shadow-2xs'
            : (isSecond ? 'bg-slate-200 text-slate-800' : 'bg-orange-100 text-orange-900');
          const avatarBg = isFirst
            ? 'bg-gradient-to-br from-amber-400 via-amber-500 to-orange-500 text-slate-950 shadow-md'
            : `bg-gradient-to-br ${getRepAvatarGradient(r.nom)} text-white shadow-sm`;
          const barGradient = r.targetPct >= 100
            ? 'bg-gradient-to-r from-emerald-500 to-teal-400'
            : (isFirst ? 'bg-gradient-to-r from-amber-500 to-orange-400' : 'bg-gradient-to-r from-indigo-600 to-sky-500');

          return `
            <div style="animation-delay: ${(rank - 1) * 55}ms" class="${orderCls} rank-card-smooth rounded-3xl p-5 text-center space-y-3 ${podiumCls} ${isMe ? 'ring-2 ring-indigo-500/45' : ''}">
              <div class="flex items-center justify-center gap-2">
                <span class="px-3 py-0.5 text-[11px] font-extrabold rounded-full ${badgeStyle}">${crownOrMedal} #${rank}</span>
                ${isMe ? '<span class="px-2 py-0.5 text-[10px] font-extrabold uppercase rounded-full bg-indigo-600 text-white">Vous</span>' : ''}
              </div>
              <div class="w-12 h-12 mx-auto rounded-2xl ${avatarBg} font-extrabold text-base flex items-center justify-center">
                ${esc(r.nom.slice(0, 2).toUpperCase())}
              </div>
              <div>
                <div class="font-extrabold text-base text-slate-900 truncate">${esc(r.nom)}</div>
                <div class="text-[11px] font-semibold text-slate-500 mt-0.5">${sharePct} % du volume équipe</div>
              </div>
              <div class="text-lg sm:text-2xl font-extrabold ${isFirst ? 'text-amber-700' : 'text-slate-900'} tabular-nums break-words">
                ${fmtMoney(r.total)}
              </div>
              <div class="space-y-1.5 pt-1">
                <div class="flex items-center justify-between text-[11px] font-bold tabular-nums">
                  <span class="text-slate-500">Target : ${fmtMoney0(r.repTarget)}</span>
                  <span class="${r.targetPct >= 100 ? 'text-emerald-600' : 'text-indigo-700'}">${r.targetPct} %</span>
                </div>
                ${miniRocketProgressTrackHtml(barGradient, r.targetPct, 'h-2.5', 'bg-slate-100')}
              </div>
            </div>
          `;
        }).join('');
      }

      // 3. CARTES DÉTAILLÉES POUR TOUS LES REPRÉSENTANTS
      const medals = ['🥇', '🥈', '🥉'];
      const leaderVol = enriched[0]?.total || 0;

      document.getElementById('classement-cards').innerHTML = enriched.map((r, idx) => {
        const barWidth = Math.min(100, Math.max(r.total > 0 ? 4 : 0, r.targetPct));
        const isMe = r.nom === currentProfile.nom;
        const reached = r.targetPct >= 100;
        const sharePct = teamTotalVol > 0 ? Math.round((r.total / teamTotalVol) * 100) : 0;
        const gapToLeader = idx === 0 ? 0 : Math.max(0, leaderVol - r.total);
        const rankPillCls = idx === 0
          ? 'bg-amber-100 text-amber-900 border border-amber-300'
          : (idx === 1
            ? 'bg-slate-200/80 text-slate-800 border border-slate-300'
            : (idx === 2 ? 'bg-orange-100 text-orange-900 border border-orange-200' : 'bg-slate-100 text-slate-700 border border-slate-200'));
        const progressGradient = reached
          ? 'bg-gradient-to-r from-emerald-500 to-teal-400'
          : (r.targetPct >= 65
            ? 'bg-gradient-to-r from-indigo-600 via-sky-500 to-emerald-400'
            : 'bg-gradient-to-r from-indigo-600 to-sky-500');

        return `
          <div style="animation-delay: ${Math.min(idx * 45, 250)}ms" class="rank-card-smooth bg-white p-5 rounded-2xl border ${isMe ? 'border-indigo-400 ring-2 ring-indigo-600/20' : 'border-slate-200/80'} shadow-xs space-y-3.5">
            <div class="flex items-start justify-between gap-3">
              <div class="flex items-center gap-3 min-w-0">
                <div class="relative shrink-0">
                  <div class="w-11 h-11 rounded-2xl bg-gradient-to-br ${getRepAvatarGradient(r.nom)} text-white font-extrabold text-sm flex items-center justify-center shadow-xs">
                    ${esc(r.nom.slice(0, 2).toUpperCase())}
                  </div>
                  <span class="absolute -bottom-1.5 -right-1.5 px-1.5 py-0.2 rounded-full text-[10px] font-extrabold ${rankPillCls} shadow-2xs">
                    ${medals[idx] || `#${idx + 1}`}
                  </span>
                </div>
                <div class="min-w-0">
                  <div class="font-extrabold text-slate-900 flex items-center gap-2 truncate">
                    <span>${esc(r.nom)}</span>
                    ${isMe ? '<span class="text-[10px] font-extrabold uppercase px-2 py-0.5 rounded-full bg-indigo-50 text-indigo-700 border border-indigo-200">Vous</span>' : ''}
                  </div>
                  <div class="text-[11px] text-slate-500 flex flex-wrap items-center gap-1.5 mt-0.5">
                    <span class="font-semibold">${idx === 0 ? '👑 Leader' : `Rang #${idx + 1}`}</span>
                    <span>•</span>
                    <span class="font-bold text-indigo-600">${sharePct} % de l'équipe</span>
                  </div>
                </div>
              </div>
              <div class="text-right shrink-0 tabular-nums">
                <div class="text-lg font-extrabold text-slate-900">${fmtMoney(r.total)}</div>
                <div class="text-[11px] font-semibold text-slate-400">🎯 Target : ${fmtMoney0(r.repTarget)}</div>
              </div>
            </div>

            <!-- MINI INDICATEURS DE PERFORMANCE -->
            <div class="grid grid-cols-3 gap-2">
              <div class="rank-stat-tile bg-slate-50/90 rounded-xl p-2.5 border border-slate-100 text-center text-[11px] tabular-nums">
                <span class="text-slate-400 block text-[10px]">Dossiers</span>
                <strong class="text-slate-800">${r.dealsCount > 0 ? r.dealsCount : '—'}</strong>
                ${(r.otsCount > 0 || r.suiviCount > 0) ? `<span class="text-[9px] text-slate-400 block">${r.otsCount} OTS • ${r.suiviCount} Suivi</span>` : ''}
              </div>
              <div class="rank-stat-tile bg-slate-50/90 rounded-xl p-2.5 border border-slate-100 text-center text-[11px] tabular-nums">
                <span class="text-slate-400 block text-[10px]">Panier moyen</span>
                <strong class="text-slate-800">${r.avgTicket > 0 ? fmtMoney0(r.avgTicket) : '—'}</strong>
                ${r.bestSale > 0 ? `<span class="text-[9px] text-emerald-700 font-semibold block">Max : ${fmtMoney0(r.bestSale)}</span>` : ''}
              </div>
              <div class="rank-stat-tile bg-slate-50/90 rounded-xl p-2.5 border border-slate-100 text-center text-[11px] tabular-nums">
                <span class="text-slate-400 block text-[10px]">${r.topEntName ? 'Top partenaire' : 'Écart #1'}</span>
                <strong class="text-slate-800 truncate block px-1">${r.topEntName ? esc(r.topEntName) : (idx === 0 ? '👑 En tête' : `- ${fmtMoney0(gapToLeader)}`)}</strong>
                ${r.topEntVal > 0 ? `<span class="text-[9px] text-indigo-600 font-semibold block">${fmtMoney0(r.topEntVal)}</span>` : ''}
              </div>
            </div>

            <div class="space-y-1.5">
              <div class="flex items-center justify-between text-[11px] tabular-nums">
                <span class="font-bold ${reached ? 'text-emerald-600' : 'text-slate-600'}">
                  ${reached ? '🔥 Target atteint !' : `Reste ${fmtMoney0(r.remaining)} avant target`}
                </span>
                <span class="px-2 py-0.5 rounded-full font-extrabold ${reached ? 'bg-emerald-100 text-emerald-800' : (r.targetPct >= 65 ? 'bg-amber-100 text-amber-800' : 'bg-indigo-50 text-indigo-700')}">
                  ${r.targetPct} %
                </span>
              </div>
              ${miniRocketProgressTrackHtml(progressGradient, r.targetPct, 'h-2.5', 'bg-slate-100')}
            </div>
          </div>
        `;
      }).join('');
    }

    function getRateForSale(repId, entId, clientId) {
      const rStr = String(repId || '');
      const eStr = String(entId || '');
      const cStr = clientId ? String(clientId) : '';
      return (cStr && state.taux.find(t => String(t.representant_id) === rStr && String(t.entreprise_id) === eStr && String(t.client_id || '') === cStr))
        || state.taux.find(t => String(t.representant_id) === rStr && String(t.entreprise_id) === eStr && !t.client_id);
    }

    function getRepMonthPriorSalesVolume(repId, ym, excludeSaleId = null) {
      const rStr = String(repId || '');
      const targetYm = ym || currentMonthStr;
      const sourceList = (state.ventes && state.ventes.length) ? state.ventes : (state.mesVentes || []);
      return sourceList.reduce((sum, v) => {
        if (excludeSaleId && String(v.id) === String(excludeSaleId)) return sum;
        if (String(v.representant_id) !== rStr) return sum;
        const vYm = getSaleStatsMonth(v);
        if (vYm !== targetYm) return sum;
        return sum + Number(v.valeur || 0);
      }, 0);
    }

    function getEffectiveRatesWithTargetBonus(repId, entId, clientId, ym, excludeSaleId = null, customPriorVolume = null) {
      let rateObj = getRateForSale(repId, entId, clientId);
      if (!rateObj && state.mesVentes && state.mesVentes.length) {
        const pastSale = state.mesVentes.find(v => String(v.representant_id) === String(repId) && String(v.entreprise_id) === String(entId) && Number(v.taux_rep_applique) > 0);
        if (pastSale) {
          rateObj = { taux: Number(pastSale.taux_rep_applique), taux_olivier: Number(pastSale.taux_olivier_applique || 0) };
        }
      }

      const baseRep = rateObj ? Number(rateObj.taux) : 0;
      const baseOli = rateObj ? Number(rateObj.taux_olivier) : 0;

      const repObj = state.representants.find(r => String(r.id) === String(repId))
        || (currentProfile && String(currentProfile.id) === String(repId) ? currentProfile : null);
      const repName = repObj?.nom || '';
      const activeYm = ym || currentMonthStr;
      const salesTarget = repName ? getSalesTargetForRep(repName, activeYm) : 0;
      const priorVolume = customPriorVolume !== null
        ? Number(customPriorVolume)
        : getRepMonthPriorSalesVolume(repId, activeYm, excludeSaleId);

      // Le bonus configurable (+X% Rep / -X% Olivier) s'applique uniquement sur les nouvelles ventes une fois l'objectif mensuel atteint (non rétroactif)
      const isRepRole = !repObj || repObj.role !== 'admin';
      const configuredBonusPct = repName ? getBonusPctForRep(repName, activeYm) : 0.01;
      const bonusActive = Boolean(rateObj && isRepRole && configuredBonusPct > 0 && salesTarget > 0 && priorVolume >= salesTarget);
      const bonusDelta = bonusActive ? configuredBonusPct : 0;

      const finalRep = Math.round((baseRep + bonusDelta) * 100000) / 100000;
      const finalOli = bonusActive
        ? Math.max(0, Math.round((baseOli - bonusDelta) * 100000) / 100000)
        : baseOli;

      return {
        hasRule: Boolean(rateObj),
        baseRep,
        baseOli,
        tRep: finalRep,
        tOli: finalOli,
        bonusActive,
        bonusDelta,
        priorVolume,
        salesTarget
      };
    }

    function computeDrawerCommissions(valeur) {
      const eff = getDrawerComputedRates();
      const isHit = Boolean(document.getElementById('vente-is-hit')?.checked);
      const hitBearer = document.getElementById('vente-hit-bearer')?.value || 'rep';
      const rawHitAmt = parseFloat(document.getElementById('vente-hit-amount')?.value);
      const rawHitRep = parseFloat(document.getElementById('vente-hit-amount-rep')?.value);
      const rawHitOli = parseFloat(document.getElementById('vente-hit-amount-oli')?.value);

      const dedSingle = (!isNaN(rawHitAmt) && rawHitAmt > 0) ? Math.abs(rawHitAmt) : 0;
      const dedRepInput = (!isNaN(rawHitRep) && rawHitRep > 0) ? Math.abs(rawHitRep) : 0;
      const dedOliInput = (!isNaN(rawHitOli) && rawHitOli > 0) ? Math.abs(rawHitOli) : 0;

      const baseCRep = Math.round(valeur * eff.tRep * 100) / 100;
      const baseCOli = Math.round(valeur * eff.tOli * 100) / 100;
      let cRep = baseCRep;
      let cOli = baseCOli;
      let totalDedAmt = dedSingle;

      // Ne jamais mettre la commission à 0 automatiquement : déduire uniquement les montants saisis
      if (isHit) {
        if (hitBearer === 'rep' && dedSingle > 0) {
          cRep = Math.round((baseCRep - dedSingle) * 100) / 100;
          if (Object.is(cRep, -0)) cRep = 0;
          cOli = baseCOli;
          totalDedAmt = dedSingle;
        } else if (hitBearer === 'olivier' && dedSingle > 0) {
          cRep = baseCRep;
          cOli = Math.round((baseCOli - dedSingle) * 100) / 100;
          if (Object.is(cOli, -0)) cOli = 0;
          totalDedAmt = dedSingle;
        } else if (hitBearer === 'both') {
          cRep = Math.round((baseCRep - dedRepInput) * 100) / 100;
          cOli = Math.round((baseCOli - dedOliInput) * 100) / 100;
          if (Object.is(cRep, -0)) cRep = 0;
          if (Object.is(cOli, -0)) cOli = 0;
          totalDedAmt = Math.round((dedRepInput + dedOliInput) * 100) / 100;
        }
      }

      return {
        ...eff,
        cRep,
        cOli,
        isHit,
        hitBearer,
        hitAmount: totalDedAmt,
        hitAmountRep: hitBearer === 'both' ? dedRepInput : (hitBearer === 'rep' ? dedSingle : 0),
        hitAmountOli: hitBearer === 'both' ? dedOliInput : (hitBearer === 'olivier' ? dedSingle : 0)
      };
    }

    function getDrawerComputedRates() {
      const editId = document.getElementById('vente-edit-id')?.value || null;
      const repId = document.getElementById('vente-rep').value;
      const entId = document.getElementById('vente-entreprise').value;
      const ym = document.getElementById('vente-mois-comm')?.value || currentMonthStr;
      const clientNom = document.getElementById('vente-client-nom').value.trim().toLowerCase();
      const matchedClient = state.clients.find(c => String(c.entreprise_id) === String(entId) && c.nom.toLowerCase() === clientNom);
      const eff = getEffectiveRatesWithTargetBonus(repId, entId, matchedClient?.id || null, ym, editId);

      const isAdmin = currentProfile?.role === 'admin';
      const rateMode = isAdmin ? (document.getElementById('vente-rate-mode')?.value || 'auto') : 'auto';
      if (rateMode === 'no_bonus') {
        return {
          ...eff,
          rateMode: 'no_bonus',
          tRep: eff.baseRep,
          tOli: eff.baseOli,
          bonusActive: false,
          bonusDelta: 0
        };
      }
      if (rateMode === 'custom') {
        const rawR = parseFloat(document.getElementById('vente-custom-taux-rep')?.value);
        const rawO = parseFloat(document.getElementById('vente-custom-taux-oli')?.value);
        const cRepRate = isNaN(rawR) ? eff.tRep : Math.max(0, rawR / 100);
        const cOliRate = isNaN(rawO) ? eff.tOli : Math.max(0, rawO / 100);
        return {
          ...eff,
          hasRule: true,
          rateMode: 'custom',
          tRep: Math.round(cRepRate * 100000) / 100000,
          tOli: Math.round(cOliRate * 100000) / 100000,
          bonusActive: false,
          bonusDelta: 0
        };
      }
      return { ...eff, rateMode: 'auto' };
    }

    function updateDrawerCommissionPreview() {
      const val = parseFloat(document.getElementById('vente-valeur').value) || 0;
      const isAdmin = currentProfile?.role === 'admin';
      const overrideBox = document.getElementById('admin-drawer-rate-override-box');
      if (overrideBox) overrideBox.classList.toggle('hidden', !isAdmin);

      const rateMode = isAdmin ? (document.getElementById('vente-rate-mode')?.value || 'auto') : 'auto';
      const customRow = document.getElementById('vente-custom-rates-row');
      if (customRow) customRow.classList.toggle('hidden', rateMode !== 'custom');

      const isHitChecked = Boolean(document.getElementById('vente-is-hit')?.checked);
      const hitDetailsBox = document.getElementById('vente-hit-details');
      if (hitDetailsBox) hitDetailsBox.classList.toggle('hidden', !isHitChecked);

      const bearerSel = document.getElementById('vente-hit-bearer');
      if (bearerSel) {
        const oliOpt = bearerSel.querySelector('option[value="olivier"]');
        if (oliOpt) oliOpt.disabled = !isAdmin;
        if (!isAdmin && bearerSel.value === 'olivier') bearerSel.value = 'rep';
        const isBothMode = bearerSel.value === 'both';
        document.getElementById('vente-hit-single-box')?.classList.toggle('hidden', isBothMode);
        document.getElementById('vente-hit-both-box')?.classList.toggle('hidden', !isBothMode);
      }

      const comp = computeDrawerCommissions(val);
      const { hasRule, tRep, tOli, bonusActive, bonusDelta, cRep, cOli, isHit, hitBearer } = comp;
      const bStr = formatBonusPctShort(bonusDelta);

      const repEl = document.getElementById('preview-comm-rep');
      const oliEl = document.getElementById('preview-comm-olivier');
      const hitBadge = isHit
        ? ` <span class="inline-block ml-1 px-1.5 py-0.5 text-[10px] font-extrabold rounded-md bg-rose-100 text-rose-700">🚨 HIT (${hitBearer === 'rep' ? 'Rep' : (hitBearer === 'olivier' ? 'Olivier' : 'Rep + Oli')})</span>`
        : '';

      if (repEl) {
        repEl.className = cRep < 0
          ? 'text-sm font-extrabold text-rose-600 tabular-nums'
          : 'text-sm font-extrabold text-slate-900 tabular-nums';
      }
      if (oliEl) {
        oliEl.className = cOli < 0
          ? 'text-xs font-extrabold text-rose-600 tabular-nums'
          : 'text-xs font-bold text-emerald-600 tabular-nums';
      }

      if (isAdmin) {
        let tagHtml = '';
        if (rateMode === 'no_bonus') {
          tagHtml = ' <span class="inline-block ml-1 px-1.5 py-0.5 text-[10px] font-extrabold rounded-md bg-amber-100 text-amber-800">Bypass bonus</span>';
        } else if (rateMode === 'custom') {
          tagHtml = ' <span class="inline-block ml-1 px-1.5 py-0.5 text-[10px] font-extrabold rounded-md bg-indigo-100 text-indigo-800">Taux manuel</span>';
        } else if (bonusActive) {
          tagHtml = ` <span class="inline-block ml-1 px-1.5 py-0.5 text-[10px] font-extrabold rounded-md bg-emerald-100 text-emerald-800">🔥 +${bStr} Rep / -${bStr} Oli</span>`;
        }
        document.getElementById('preview-taux-label').innerHTML = hasRule
          ? `Règle : <strong>${fmtPct(tRep)}</strong> Rep / <strong>${fmtPct(tOli)}</strong> Olivier${tagHtml}${hitBadge}`
          : `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded-lg bg-amber-100 text-amber-900 border border-amber-300 font-bold text-[11px]">⚠️ Aucun taux configuré pour ce rep/entreprise (0 % par défaut)</span>${hitBadge}`;
        repEl.textContent = `Rep : ${fmtMoney(cRep)}`;
        oliEl.textContent = `Olivier : ${fmtMoney(cOli)}`;
        oliEl.classList.remove('hidden');
      } else {
        document.getElementById('preview-taux-label').innerHTML = hasRule
          ? `Votre taux applicable : <strong>${fmtPct(tRep)}</strong>${bonusActive ? ` <span class="inline-block ml-1 px-1.5 py-0.5 text-[10px] font-extrabold rounded-md bg-emerald-100 text-emerald-800">🔥 Bonus +${bStr} actif</span>` : ''}${hitBadge}`
          : `<span class="inline-block px-2 py-1 rounded-lg bg-amber-50 text-amber-900 border border-amber-200 font-semibold text-[11px]">⏳ Taux en attente de configuration par l'administration — votre vente sera bien enregistrée et votre commission recalculée dès validation.</span>${hitBadge}`;
        repEl.textContent = hasRule ? `Ma commission : ${fmtMoney(cRep)}` : 'En attente du taux';
        oliEl.classList.add('hidden');
      }
    }

    window.openEditVenteDrawer = (id, duplicate = false) => {
      const v = state.ventes.find(item => item.id === id);
      if (!v) return;
      const cli = state.clients.find(c => c.id === v.client_id)?.nom || '';
      document.getElementById('vente-edit-id').value = duplicate ? '' : v.id;
      document.getElementById('drawer-title').textContent = duplicate ? 'Dupliquer la vente' : 'Modifier la vente';
      document.getElementById('drawer-subtitle').textContent = duplicate ? 'Crée un nouveau dossier à partir de cette vente' : 'Les commissions seront recalculées selon la grille active';
      document.getElementById('btn-submit-drawer').textContent = duplicate ? 'Créer la copie' : 'Mettre à jour';

      document.getElementById('vente-rep').value = v.representant_id;
      document.getElementById('vente-entreprise').value = v.entreprise_id;
      updateClientDatalist();
      document.getElementById('vente-client-nom').value = cli;
      document.getElementById('vente-valeur').value = v.valeur;
      document.getElementById('vente-date').value = duplicate ? localYmd() : v.date_vente;
      document.getElementById('vente-mois-comm').value = v.mois_commission;
      document.getElementById('vente-statut-depot').value = duplicate ? 'a_confirmer' : ((v.statut_depot === 'depot_recu' && isNoDepositSale(v)) ? 'sans_depot' : v.statut_depot);
      document.getElementById('vente-statut-paiement').value = duplicate ? 'non_paye' : v.statut_paiement;
      const rawN = v.note || '';
      const hasHitTag = /\[HIT(?::(REP|OLIVIER|BOTH))?\]/i.test(rawN) || Number(v.commission_rep) < 0 || Number(v.commission_olivier) < 0;
      const ovHit = !duplicate ? saleBonusOverridesStore[String(v.id)] : null;
      document.getElementById('vente-type-kpi').value = rawN.toUpperCase().includes('[SUIVI]') ? 'SUIVI' : 'OTS';
      applyDrawerRoleErgonomics();
      renderDrawerClientQuickSuggestions();
      document.getElementById('vente-note').value = rawN
        .replace(/^\[(OTS|SUIVI)\]\s*/i, '').replace(/\[(REP|ENT):[^\]]+\]\s*/ig, '').replace(/\[SANSDEPOT\]\s*/ig, '')
        .replace(/\[HIT(?::(?:REP|OLIVIER|BOTH))?\]\s*/ig, '')
        .trim();

      const hitCb = document.getElementById('vente-is-hit');
      const hitBearerSel = document.getElementById('vente-hit-bearer');
      const hitAmtIn = document.getElementById('vente-hit-amount');
      const hitAmtRepIn = document.getElementById('vente-hit-amount-rep');
      const hitAmtOliIn = document.getElementById('vente-hit-amount-oli');
      if (hitCb && hitBearerSel && hitAmtIn) {
        hitCb.checked = !duplicate && hasHitTag;
        if (!duplicate && hasHitTag) {
          const tagMatch = rawN.match(/\[HIT:(REP|OLIVIER|BOTH)\]/i);
          let bearer = ovHit?.hitBearer || (tagMatch ? tagMatch[1].toLowerCase() : null);
          if (!bearer) {
            if (Number(v.commission_rep) < 0 && Number(v.commission_olivier) >= 0) bearer = 'rep';
            else if (Number(v.commission_olivier) < 0 && Number(v.commission_rep) >= 0) bearer = 'olivier';
            else bearer = 'both';
          }
          hitBearerSel.value = bearer;
          if (bearer === 'both') {
            hitAmtIn.value = '';
            if (hitAmtRepIn) hitAmtRepIn.value = ovHit?.hitAmountRep > 0 ? ovHit.hitAmountRep : (ovHit?.hitAmount > 0 ? Math.round((ovHit.hitAmount / 2) * 100) / 100 : '');
            if (hitAmtOliIn) hitAmtOliIn.value = ovHit?.hitAmountOli > 0 ? ovHit.hitAmountOli : (ovHit?.hitAmount > 0 ? Math.round((ovHit.hitAmount / 2) * 100) / 100 : '');
          } else {
            hitAmtIn.value = ovHit?.hitAmount > 0 ? ovHit.hitAmount : '';
            if (hitAmtRepIn) hitAmtRepIn.value = '';
            if (hitAmtOliIn) hitAmtOliIn.value = '';
          }
        } else {
          hitBearerSel.value = 'both';
          hitAmtIn.value = '';
          if (hitAmtRepIn) hitAmtRepIn.value = '';
          if (hitAmtOliIn) hitAmtOliIn.value = '';
        }
      }

      const modeSel = document.getElementById('vente-rate-mode');
      const customRepIn = document.getElementById('vente-custom-taux-rep');
      const customOliIn = document.getElementById('vente-custom-taux-oli');
      if (modeSel && customRepIn && customOliIn) {
        const ov = !duplicate ? saleBonusOverridesStore[String(v.id)] : null;
        if (ov?.noBonus) {
          modeSel.value = 'no_bonus';
        } else if (ov?.customManual) {
          modeSel.value = 'custom';
          customRepIn.value = Math.round(Number(ov.tRep) * 10000) / 100;
          customOliIn.value = Math.round(Number(ov.tOli) * 10000) / 100;
        } else {
          modeSel.value = 'auto';
          customRepIn.value = Math.round(Number(v.taux_rep_applique || 0) * 10000) / 100;
          customOliIn.value = Math.round(Number(v.taux_olivier_applique || 0) * 10000) / 100;
        }
      }

      const existingProof = !duplicate ? getProofForSale(v) : '';
      pendingProofDataUrl = existingProof || '';
      window._editingOriginalHadProof = Boolean(existingProof);
      const pfIn = document.getElementById('vente-proof-file');
      if (pfIn) pfIn.value = '';
      const cmIn = document.getElementById('vente-camera-input');
      if (cmIn) cmIn.value = '';
      const prevBox = document.getElementById('vente-proof-preview-box');
      const prevImg = document.getElementById('vente-proof-preview-img');
      const btnClear = document.getElementById('btn-clear-proof');
      const statusBox = document.getElementById('vente-proof-status');
      if (existingProof) {
        if (prevImg) prevImg.src = existingProof;
        prevBox?.classList.remove('hidden');
        btnClear?.classList.remove('hidden');
        if (statusBox) {
          statusBox.textContent = '✓ Preuve actuelle attachée (cliquez sur « Retirer » pour la supprimer)';
          statusBox.classList.remove('hidden');
        }
      } else {
        if (prevImg) prevImg.src = '';
        prevBox?.classList.add('hidden');
        btnClear?.classList.add('hidden');
        statusBox?.classList.add('hidden');
      }

      updateDrawerCommissionPreview();
      document.getElementById('sale-drawer-backdrop').classList.remove('hidden');
    };

    function applyDrawerRoleErgonomics() {
      applyCompanyDepositRule();
      const isAdmin = currentProfile?.role === 'admin';
      const repWrap = document.getElementById('drawer-rep-field-wrap');
      const repEntGrid = document.getElementById('drawer-rep-ent-grid');
      if (repWrap) repWrap.classList.toggle('hidden', !isAdmin);
      if (repEntGrid) {
        repEntGrid.className = isAdmin
          ? 'grid grid-cols-1 sm:grid-cols-2 gap-3.5'
          : 'grid grid-cols-1 gap-3.5';
      }

      const payWrap = document.getElementById('drawer-statut-paiement-wrap');
      const payNote = document.getElementById('drawer-pay-note');
      const statusGrid = document.getElementById('drawer-status-grid');
      if (payWrap) payWrap.classList.toggle('hidden', !isAdmin);
      if (payNote) payNote.classList.toggle('hidden', isAdmin);
      document.getElementById('drawer-pay-note')?.classList.toggle('hidden', isAdmin);
      if (statusGrid) {
        statusGrid.className = isAdmin
          ? 'grid grid-cols-1 sm:grid-cols-3 gap-3'
          : 'grid grid-cols-1 sm:grid-cols-2 gap-3';
      }

      const hitWrap = document.getElementById('drawer-hit-section-wrap');
      const isHitChecked = Boolean(document.getElementById('vente-is-hit')?.checked);
      if (hitWrap) hitWrap.classList.toggle('hidden', !isAdmin && !isHitChecked);
    }

    function populateDrawerSelects() {
      const isAdmin = currentProfile.role === 'admin';
      const entOptions = state.entreprises.map(e => `<option value="${e.id}">${esc(e.nom)}</option>`).join('');
      document.getElementById('vente-entreprise').innerHTML = entOptions;

      const repSel = document.getElementById('vente-rep');
      if (isAdmin && state.representants.length) {
        repSel.innerHTML = state.representants.map(r => `<option value="${r.id}">${esc(r.nom)}</option>`).join('');
        repSel.disabled = false;
      } else {
        repSel.innerHTML = `<option value="${currentProfile.id}">${esc(currentProfile.nom)}</option>`;
        repSel.value = currentProfile.id;
        repSel.disabled = true;
      }
      applyDrawerRoleErgonomics();
      updateClientDatalist();
      updateDrawerCommissionPreview();
    }

    function populateAdminSelects() {
      const prevFilterRep = document.getElementById('admin-filter-rep').value;
      const prevFilterEnt = document.getElementById('admin-filter-ent').value;

      const repOptions = state.representants.map(r => `<option value="${r.id}">${esc(r.nom)}</option>`).join('');
      const entOptions = state.entreprises.map(e => `<option value="${e.id}">${esc(e.nom)}</option>`).join('');

      populateDrawerSelects();
      document.getElementById('taux-rep').innerHTML = repOptions;
      document.getElementById('taux-ent').innerHTML = entOptions;

      document.getElementById('admin-filter-rep').innerHTML = `<option value="">Tous les représentants</option>${repOptions}`;
      document.getElementById('admin-filter-ent').innerHTML = `<option value="">Toutes les entreprises</option>${entOptions}`;

      const scopeSel = document.getElementById('filter-admin-scope');
      if (scopeSel) {
        const prevScope = scopeSel.value || 'all';
        const otherRepsOpts = state.representants
          .filter(r => String(r.id) !== String(currentProfile?.id))
          .map(r => `<option value="rep:${r.id}">👤 Vue de ${esc(r.nom)} (ses ventes & commissions)</option>`)
          .join('');
        scopeSel.innerHTML = `
          <option value="all">Toutes mes commissions (Directes + Équipe)</option>
          <option value="direct">Mes ventes directes uniquement</option>
          ${otherRepsOpts ? `<optgroup label="Inspecter la vue d'un représentant">${otherRepsOpts}</optgroup>` : ''}
        `;
        if (Array.from(scopeSel.options).some(o => o.value === prevScope)) {
          scopeSel.value = prevScope;
        }
      }

      if (prevFilterRep) document.getElementById('admin-filter-rep').value = prevFilterRep;
      if (prevFilterEnt) document.getElementById('admin-filter-ent').value = prevFilterEnt;

      updateTauxClientSelect();
    }

    function renderDrawerClientQuickSuggestions() {
      const box = document.getElementById('drawer-client-quick-suggestions');
      const input = document.getElementById('vente-client-nom');
      if (!box || !input) return;
      const q = input.value.trim().toLowerCase();
      if (q.length < 2) {
        box.classList.add('hidden');
        box.innerHTML = '';
        return;
      }
      const entId = document.getElementById('vente-entreprise')?.value || '';
      const pool = uniqueClientsByName(
        entId ? state.clients.filter(c => String(c.entreprise_id) === String(entId)) : state.clients
      );
      const matches = pool.filter(c => c.nom.toLowerCase().includes(q) && c.nom.toLowerCase() !== q).slice(0, 5);
      if (!matches.length) {
        box.classList.add('hidden');
        box.innerHTML = '';
        return;
      }
      box.classList.remove('hidden');
      box.innerHTML = `<span class="text-[10px] font-bold text-slate-400">Clients existants :</span>` + matches.map(c => `
        <button type="button" onclick="pickDrawerClientSuggestion('${esc(c.nom)}')" class="px-2 py-0.5 text-[11px] font-bold rounded-lg bg-indigo-50 hover:bg-indigo-100 text-indigo-800 border border-indigo-200 transition">
          ${esc(c.nom)}
        </button>
      `).join('');
    }

    window.pickDrawerClientSuggestion = (nom) => {
      const input = document.getElementById('vente-client-nom');
      if (!input) return;
      input.value = nom;
      renderDrawerClientQuickSuggestions();
      updateDrawerCommissionPreview();
    }
    document.getElementById('vente-client-nom')?.addEventListener('input', renderDrawerClientQuickSuggestions);;

    function updateClientDatalist() {
      const entId = document.getElementById('vente-entreprise').value;
      const list = uniqueClientsByName(state.clients.filter(c => String(c.entreprise_id) === String(entId)));
      document.getElementById('clients-datalist').innerHTML = list.map(c => `<option value="${esc(c.nom)}">`).join('');
      renderDrawerClientQuickSuggestions();
      updateDrawerCommissionPreview();
    }

    function updateTauxClientSelect() {
      const entId = document.getElementById('taux-ent').value;
      const list = state.clients.filter(c => c.entreprise_id === entId);
      document.getElementById('taux-client').innerHTML = `<option value="">— Toute l'entreprise —</option>` +
        list.map(c => `<option value="${c.id}">${esc(c.nom)}</option>`).join('');
    }

    function updateBulkBar() {
      const bar = document.getElementById('bulk-action-bar');
      const count = selectedSaleIds.size;
      bar.classList.toggle('hidden', count === 0);
      document.getElementById('bulk-count-label').textContent = `${count} sélectionnée(s)`;
    }

    window.toggleSelectSale = (id, checked) => {
      if (checked) selectedSaleIds.add(id);
      else selectedSaleIds.delete(id);
      updateBulkBar();
    };

    window.clearBulkSelection = () => {
      selectedSaleIds.clear();
      const master = document.getElementById('select-all-ventes');
      if (master) master.checked = false;
      document.querySelectorAll('.sale-row-cb').forEach(cb => cb.checked = false);
      updateBulkBar();
    };

    window.bulkUpdateStatut = async (field, newVal) => {
      const ids = Array.from(selectedSaleIds);
      if (!ids.length) return;
      const { error } = await sb.from('ventes').update({ [field]: newVal }).in('id', ids);
      if (error) return toast(error.message, 'error');
      const label = field === 'statut_depot'
        ? (newVal === 'depot_recu' ? 'Dépôt confirmé (Lot)' : 'Dépôt remis À confirmer (Lot)')
        : (newVal === 'paye' ? 'Commission marquée Payée (Lot)' : 'Commission remise Non payée (Lot)');
      for (const sid of ids) {
        await recordSaleAudit(sid, label, '', '📦');
      }
      let mergedCount = 0;
      if (field === 'statut_depot' && newVal === 'depot_recu') {
        mergedCount = await reconcilePreviousUnconfirmedDuplicateSales(ids);
      }
      toast(mergedCount > 0
        ? `${ids.length} vente(s) mise(s) à jour • ${mergedCount} doublon antérieur retiré ✓`
        : `${ids.length} vente(s) mise(s) à jour`);
      await loadAllData();
    };

    window.bulkDeleteVentes = async () => {
      const ids = Array.from(selectedSaleIds);
      if (!ids.length) return;
      const ok = await askConfirm(`Mettre ${ids.length} vente(s) à la corbeille ?`, 'Les ventes sélectionnées pourront être restaurées depuis la corbeille pendant 30 jours.');
      if (!ok) return;
      const toTrash = state.ventes.filter(v => ids.includes(v.id));
      if (toTrash.length) {
        const trashOk = await moveSalesToTrash(toTrash);
        if (!trashOk) return toast('La sauvegarde cloud de la corbeille a échoué : suppression annulée.', 'error');
        for (const ts of toTrash) {

        }
      }
      const { error } = await sb.from('ventes').delete().in('id', ids);
      if (error) return toast(error.message, 'error');
      toast(`${ids.length} vente(s) déplacée(s) dans la corbeille 🗑️`);
      await loadAllData();
    };

    function renderAdminVentes() {
      const q = String(window._globalClientFilterQuery || '').trim().toLowerCase();
      const fRep = document.getElementById('admin-filter-rep').value;
      const fEnt = document.getElementById('admin-filter-ent').value;
      const adminPeriod = readPeriodFilterState('admin-filter');
      const fStat = document.getElementById('admin-filter-statut').value;

      const list = state.ventes.filter(v => {
        if (fRep && v.representant_id !== fRep) return false;
        if (fEnt && v.entreprise_id !== fEnt) return false;
        if (!matchesSalePeriod(v, adminPeriod, true)) return false;
        if (fStat === 'depot_attente' && v.statut_depot !== 'a_confirmer') return false;
        if (fStat === 'non_paye' && v.statut_paiement !== 'non_paye') return false;
        if (fStat === 'paye' && v.statut_paiement !== 'paye') return false;
        if (q) {
          const cli = state.clients.find(c => c.id === v.client_id)?.nom || '';
          const note = v.note || '';
          if (!cli.toLowerCase().includes(q) && !note.toLowerCase().includes(q)) return false;
        }
        return true;
      });

      updateSortIndicators('adm', ['date_vente', 'rep_nom', 'entreprise_client', 'valeur', 'commission_rep', 'commission_olivier', 'statut_depot', 'statut_paiement'], adminVentesSort);
      list.sort((a, b) => {
        const col = adminVentesSort.col;
        const dir = adminVentesSort.dir;
        if (col === 'valeur' || col === 'commission_rep' || col === 'commission_olivier') {
          return compareSaleValues(a[col], b[col], dir, true);
        }
        if (col === 'date_vente') {
          const cmp = compareSaleValues(a.date_vente, b.date_vente, dir, false);
          if (cmp !== 0) return cmp;
          const cmpCreated = compareSaleValues(a.created_at || '', b.created_at || '', dir, false);
          if (cmpCreated !== 0) return cmpCreated;
          return compareSaleValues(a.id || '', b.id || '', dir, false);
        }
        if (col === 'rep_nom') {
          const rA = state.representants.find(r => r.id === a.representant_id)?.nom || '';
          const rB = state.representants.find(r => r.id === b.representant_id)?.nom || '';
          return compareSaleValues(rA, rB, dir, false);
        }
        if (col === 'entreprise_client') {
          const eA = state.entreprises.find(e => e.id === a.entreprise_id)?.nom || '';
          const eB = state.entreprises.find(e => e.id === b.entreprise_id)?.nom || '';
          const cmpE = compareSaleValues(eA, eB, dir, false);
          if (cmpE !== 0) return cmpE;
          const cA = state.clients.find(c => c.id === a.client_id)?.nom || '';
          const cB = state.clients.find(c => c.id === b.client_id)?.nom || '';
          return compareSaleValues(cA, cB, dir, false);
        }
        return compareSaleValues(a[col], b[col], dir, false);
      });

      currentFilteredAdminSales = list;
      const sumVal = list.reduce((s, v) => s + Number(v.valeur), 0);
      document.getElementById('admin-filter-summary').textContent = `${list.length} vente(s) affichée(s) • Total : ${fmtMoney(sumVal)}`;

      updateDedupBadge();
      const tbody = document.getElementById('tbody-admin-ventes');
      const mobileAdminBox = document.getElementById('mobile-admin-ventes-cards');
      if (!list.length) {
        tbody.innerHTML = `<tr><td colspan="8" class="py-12 text-center text-sm text-slate-400">Aucune vente ne correspond à ces critères.</td></tr>`;
        if (mobileAdminBox) {
          mobileAdminBox.innerHTML = `<div class="bg-white rounded-2xl border border-slate-200/80 p-8 text-center text-xs text-slate-400">Aucune vente ne correspond à ces critères.</div>`;
        }
        return;
      }

      if (mobileAdminBox) {
        mobileAdminBox.innerHTML = list.map(v => {
          const rep = state.representants.find(r => r.id === v.representant_id)?.nom || resolveSaleArchivedRepName(v) || '—';
          const ent = state.entreprises.find(e => e.id === v.entreprise_id)?.nom || resolveSaleArchivedEntName(v) || '—';
          const cli = state.clients.find(c => c.id === v.client_id)?.nom || '—';
          const isSuivi = (v.note || '').toUpperCase().includes('[SUIVI]');
          const proofSrc = getProofForSale(v);
          const cleanNote = cleanDisplaySaleNote(v.note);
          const hitRow = isHitSale(v);
          return `
            <div data-sale-id="${v.id}" class="bg-white rounded-2xl border ${hitRow ? 'border-rose-300 bg-rose-50/20' : 'border-slate-200/80'} p-4 shadow-xs space-y-3">
              <div class="flex items-start justify-between gap-2">
                <div class="min-w-0">
                  <div class="flex items-center flex-wrap gap-1.5">
                    ${entBadge(ent)}
                    ${hitBadgeHtml(v)}
                    <span class="text-xs font-bold text-indigo-700">${esc(rep)}</span>
                    <span class="text-[11px] text-slate-400">• ${v.date_vente}</span>
                    ${v._carriedOverFrom ? `<span class="px-1.5 py-0.5 text-[10px] font-bold rounded-md bg-amber-50 text-amber-800 border border-amber-200">↪ Transférée de ${v._carriedOverFrom}</span>` : ''}
                  </div>
                  <div class="font-extrabold text-sm text-slate-900 mt-1 truncate">${esc(cli)}</div>
                  ${cleanNote ? `<div class="text-[11px] text-slate-400 truncate">${esc(cleanNote)}</div>` : ''}
                </div>
                <div class="text-right shrink-0 tabular-nums">
                  <div class="text-sm font-extrabold text-slate-900">${fmtMoney(v.valeur)}</div>
                  <button type="button" onclick="openRateAdjustModal('${v.id}')" class="mt-0.5 inline-flex flex-col items-end text-right hover:opacity-80">
                    <span class="text-[11px] text-slate-500">Rep: ${fmtMoney(v.commission_rep)} (${fmtPct(v.taux_rep_applique)}) ✎</span>
                    <span class="text-[11px] font-bold text-emerald-600">Oli: ${fmtMoney(v.commission_olivier)} (${fmtPct(v.taux_olivier_applique)}) ✎</span>
                  </button>
                </div>
              </div>

              <div class="flex flex-wrap items-center gap-1.5 pt-2 border-t border-slate-100">
                <button type="button" onclick="toggleVenteKpiType('${v.id}')" class="px-2.5 py-1 text-[11px] font-extrabold rounded-lg border ${isSuivi ? 'bg-purple-50 text-purple-700 border-purple-200' : 'bg-sky-50 text-sky-700 border-sky-200'}">${isSuivi ? 'SUIVI' : 'OTS'}</button>
                <button type="button" onclick="toggleVenteStatut('${v.id}', 'statut_depot', '${v.statut_depot === 'depot_recu' ? 'a_confirmer' : 'depot_recu'}')" class="px-2.5 py-1 text-[11px] font-bold rounded-full ${v.statut_depot === 'depot_recu' ? 'bg-emerald-50 text-emerald-700 border border-emerald-200' : 'bg-amber-50 text-amber-700 border border-amber-200'}">
                  ${v.statut_depot === 'depot_recu' ? (isNoDepositSale(v) ? '✓ Sans dépôt requis' : '✓ Dépôt reçu') : '⏳ À confirmer'}
                </button>
                ${agingBadgeForDeposit(v)}
                <button type="button" onclick="toggleVenteStatut('${v.id}', 'statut_paiement', '${v.statut_paiement === 'paye' ? 'non_paye' : 'paye'}')" class="px-2.5 py-1 text-[11px] font-bold rounded-full ${v.statut_paiement === 'paye' ? 'bg-indigo-50 text-indigo-700 border border-indigo-200' : 'bg-slate-100 text-slate-600'}">
                  ${v.statut_paiement === 'paye' ? 'Payé' : 'Non payé'}
                </button>
              </div>

              <div class="grid grid-cols-4 gap-1.5 pt-1">
                ${proofSrc
                  ? `<button type="button" onclick="openProofModal('${v.id}')" class="py-2 px-1.5 text-[11px] font-bold rounded-xl bg-indigo-50 text-indigo-700 border border-indigo-200 text-center">📷 Preuve</button>`
                  : `<label class="cursor-pointer py-2 px-1.5 text-[11px] font-bold rounded-xl bg-slate-100 text-slate-700 text-center">+ 📷<input type="file" accept="image/*" class="sr-only" onchange="attachProofToExistingSale('${v.id}', this.files[0])" /></label>`
                }
                <button type="button" onclick="openSaleAuditModal('${v.id}')" class="py-2 px-1.5 text-[11px] font-bold rounded-xl bg-slate-100 text-slate-700 text-center">🕒 Audit</button>
                <button type="button" onclick="openEditVenteDrawer('${v.id}', false)" class="py-2 px-1.5 text-[11px] font-bold rounded-xl bg-slate-100 text-slate-800 text-center">Modifier</button>
                <button type="button" onclick="deleteVente('${v.id}')" class="py-2 px-1.5 text-[11px] font-bold rounded-xl bg-rose-50 text-rose-600 text-center">Supprimer</button>
              </div>
            </div>
          `;
        }).join('');
      }

      tbody.innerHTML = list.map(v => {
        const rep = state.representants.find(r => r.id === v.representant_id)?.nom || resolveSaleArchivedRepName(v) || '—';
        const ent = state.entreprises.find(e => e.id === v.entreprise_id)?.nom || resolveSaleArchivedEntName(v) || '—';
        const cli = state.clients.find(c => c.id === v.client_id)?.nom || '—';
        const isChecked = selectedSaleIds.has(v.id);
        const isSuivi = (v.note || '').toUpperCase().includes('[SUIVI]');
        const cleanNote = cleanDisplaySaleNote(v.note);
        const proofSrc = getProofForSale(v);
        const showMoisDiff = v.mois_commission && v.date_vente && !v.date_vente.startsWith(v.mois_commission);
        return `
          <tr data-sale-id="${v.id}" class="hover:bg-slate-50/90 transition ${isChecked ? 'bg-indigo-50/40' : ''}">
            <td class="py-3 px-3 text-center align-middle">
              <input type="checkbox" class="sale-row-cb rounded border-slate-300 text-indigo-600" ${isChecked ? 'checked' : ''} onchange="toggleSelectSale('${v.id}', this.checked)" />
            </td>
            <td class="py-3 px-3 text-xs whitespace-nowrap align-middle">
              <div class="font-semibold text-slate-900 tabular-nums">${v.date_vente}</div>
              ${showMoisDiff ? `<div class="text-[11px] text-slate-400">Paie ${v.mois_commission}</div>` : ''}
              ${v._carriedOverFrom ? `<span class="mt-0.5 inline-block px-1.5 py-0.5 text-[10px] font-bold rounded-md bg-amber-50 text-amber-800 border border-amber-200" title="Reportée de ${v._carriedOverFrom}">↪ ${v._carriedOverFrom}</span>` : ''}
            </td>
            <td class="py-3 px-3 align-middle">
              <div class="flex items-center gap-2">
                ${entBadge(ent)}
                <span class="font-bold text-slate-900">${esc(cli)}</span>
              </div>
              <div class="flex items-center flex-wrap gap-1.5 mt-1">
                <span class="text-xs font-semibold text-indigo-600">${esc(rep)}</span>
                <span class="text-slate-300">•</span>
                <button type="button" onclick="toggleVenteKpiType('${v.id}')" title="Basculer entre OTS et SUIVI"
                  class="px-1.5 py-0.2 text-[10px] font-extrabold rounded border transition ${isSuivi ? 'bg-purple-50 text-purple-700 border-purple-200 hover:bg-purple-100' : 'bg-sky-50 text-sky-700 border-sky-200 hover:bg-sky-100'}">
                  ${isSuivi ? 'SUIVI' : 'OTS'}
                </button>
                ${hitBadgeHtml(v)}
                ${cleanNote ? `<span class="text-[11px] text-slate-400 truncate max-w-[210px]">${esc(cleanNote)}</span>` : ''}
              </div>
            </td>
            <td class="py-3 px-3 text-right font-bold text-slate-900 tabular-nums whitespace-nowrap align-middle">${fmtMoney(v.valeur)}</td>
            <td class="py-3 px-3 text-right tabular-nums whitespace-nowrap align-middle">
              <div class="font-semibold ${Number(v.commission_rep) < 0 ? 'text-rose-600' : 'text-slate-900'}">${fmtMoney(v.commission_rep)}</div>
              <div class="text-[11px] text-slate-400 inline-flex items-center justify-end gap-1">
                <button type="button" onclick="openRateAdjustModal('${v.id}')" title="Ajuster le taux de commission" class="inline-flex items-center gap-0.5 hover:text-indigo-700 font-medium transition">
                  <span>${fmtPct(v.taux_rep_applique)}</span>
                  <span class="text-[10px]">✎</span>
                </button>
                ${(() => {
                  const baseR = getRateForSale(v.representant_id, v.entreprise_id, v.client_id);
                  const diff = baseR ? (Number(v.taux_rep_applique) - Number(baseR.taux)) : 0;
                  const ov = saleBonusOverridesStore[String(v.id)];
                  if (ov?.noBonus) return '<span class="px-1 text-[9px] font-extrabold rounded bg-amber-100 text-amber-800">Base</span>';
                  if (ov?.customManual) return '<span class="px-1 text-[9px] font-extrabold rounded bg-indigo-100 text-indigo-800">Ajusté</span>';
                  return (baseR && diff > 0.0001) ? `<span class="px-1 text-[9px] font-extrabold rounded bg-emerald-100 text-emerald-800">+${formatBonusPctShort(diff)}</span>` : '';
                })()}
              </div>
            </td>
            <td class="py-3 px-3 text-right tabular-nums whitespace-nowrap align-middle">
              <div class="font-bold ${Number(v.commission_olivier) < 0 ? 'text-rose-600' : 'text-emerald-600'}">${fmtMoney(v.commission_olivier)}</div>
              <button type="button" onclick="openRateAdjustModal('${v.id}')" title="Ajuster le taux de commission" class="text-[11px] text-slate-400 inline-flex items-center gap-0.5 hover:text-emerald-700 font-medium transition">
                <span>${fmtPct(v.taux_olivier_applique)}</span>
                <span class="text-[10px]">✎</span>
              </button>
            </td>
            <td class="py-3 px-3 whitespace-nowrap align-middle">
              <div class="flex items-center gap-1.5">
                <button onclick="toggleVenteStatut('${v.id}', 'statut_depot', '${v.statut_depot === 'depot_recu' ? 'a_confirmer' : 'depot_recu'}')"
                  class="px-2.5 py-1 text-xs font-semibold rounded-full transition ${v.statut_depot === 'depot_recu' ? 'bg-emerald-50 text-emerald-700 hover:bg-emerald-100' : 'bg-amber-50 text-amber-700 hover:bg-amber-100'}">
                  ${v.statut_depot === 'depot_recu' ? (isNoDepositSale(v) ? 'Sans dépôt requis' : 'Dépôt reçu') : 'À confirmer'}
                </button>
                <button onclick="toggleVenteStatut('${v.id}', 'statut_paiement', '${v.statut_paiement === 'paye' ? 'non_paye' : 'paye'}')"
                  class="px-2.5 py-1 text-xs font-semibold rounded-full transition ${v.statut_paiement === 'paye' ? 'bg-indigo-50 text-indigo-700 hover:bg-indigo-100' : 'bg-slate-100 text-slate-600 hover:bg-slate-200'}">
                  ${v.statut_paiement === 'paye' ? 'Payé' : 'Non payé'}
                </button>
                ${agingBadgeForDeposit(v)}
              </div>
            </td>
            <td class="py-3 px-3 text-right whitespace-nowrap align-middle">
              <div class="inline-flex items-center gap-1">
                ${proofSrc
                  ? `<button type="button" onclick="openProofModal('${v.id}')" title="Voir la preuve photo" class="p-1.5 text-xs font-bold rounded-lg bg-indigo-50 text-indigo-700 hover:bg-indigo-100 transition">📷</button>`
                  : `<label title="Ajouter une preuve photo" class="cursor-pointer p-1.5 text-xs font-semibold rounded-lg text-slate-400 hover:bg-slate-100 hover:text-indigo-700 transition">+📷<input type="file" accept="image/*" class="sr-only" onchange="attachProofToExistingSale('${v.id}', this.files[0])" /></label>`
                }
                <button onclick="openSaleAuditModal('${v.id}')" title="Journal d'audit" class="p-1.5 text-xs rounded-lg text-slate-400 hover:text-indigo-700 hover:bg-slate-100 transition">🕒</button>
                <button onclick="openEditVenteDrawer('${v.id}', false)" title="Modifier" class="px-2 py-1 text-xs font-semibold text-indigo-600 hover:bg-indigo-50 rounded-lg transition">Modifier</button>
                <button onclick="openEditVenteDrawer('${v.id}', true)" title="Dupliquer" class="p-1.5 text-xs rounded-lg text-slate-400 hover:text-slate-700 hover:bg-slate-100 transition">📋</button>
                <button onclick="deleteVente('${v.id}')" title="Supprimer" class="p-1.5 text-xs rounded-lg text-slate-400 hover:text-rose-600 hover:bg-rose-50 transition">🗑️</button>
              </div>
            </td>
          </tr>
        `;
      }).join('');
    }

    window.toggleDashFutureQuickView = () => {
      const sel = document.getElementById('dash-filter-period');
      if (!sel) return;
      sel.value = sel.value === 'future' ? 'all' : 'future';
      syncPeriodInputsVisibility('dash-filter', false);
      renderAdminDashboard();
      if (sel.value === 'future') {
        toast('🔮 Vue Futur activée : affichage des ventes et commissions à venir');
      }
    };

    function syncDashFutureButtonState() {
      const btn = document.getElementById('btn-dash-future-view');
      const badge = document.getElementById('dash-future-btn-badge');
      const futureCount = (state.ventes || []).filter(v => Boolean(getSaleTargetFutureMonth(v))).length;
      if (badge) badge.textContent = String(futureCount);
      if (!btn) return;
      const isFutureActive = document.getElementById('dash-filter-period')?.value === 'future';
      if (isFutureActive) {
        btn.className = 'px-3.5 py-2 text-xs font-extrabold text-white bg-emerald-600 hover:bg-emerald-700 border border-emerald-600 rounded-xl transition flex items-center gap-1.5 shadow-sm ring-2 ring-emerald-400/40';
        if (badge) badge.className = 'px-1.5 py-0.2 text-[10px] font-extrabold rounded-full bg-white text-emerald-800 tabular-nums';
      } else {
        btn.className = 'px-3.5 py-2 text-xs font-extrabold text-emerald-900 bg-emerald-50 hover:bg-emerald-100 border border-emerald-300 rounded-xl transition flex items-center gap-1.5 shadow-2xs';
        if (badge) badge.className = 'px-1.5 py-0.2 text-[10px] font-extrabold rounded-full bg-emerald-600 text-white tabular-nums';
      }
    }

    function renderAdminDashboard() {
      syncCommDisplayModeButtons();
      syncDashFutureButtonState();
      const dashPeriod = readPeriodFilterState('dash-filter');
      const m = dashPeriod.singleMonth;
      const statsList = state.ventes.filter(v => matchesSalePeriod(v, dashPeriod, true));
      const list = statsList;
      const confirmedList = list.filter(v => v.statut_depot === 'depot_recu');
      const statsConfirmedList = statsList.filter(v => v.statut_depot === 'depot_recu');
      const isConfirmedOnly = commDisplayMode === 'confirme';
      const activeCommList = isConfirmedOnly ? confirmedList : statsList;

      const totVentes = statsList.reduce((s, v) => s + Number(v.valeur), 0);
      const totVentesConf = statsConfirmedList.reduce((s, v) => s + Number(v.valeur), 0);
      const commRepVendu = statsList.reduce((s, v) => s + Number(v.commission_rep), 0);
      const commRepConf = confirmedList.reduce((s, v) => s + Number(v.commission_rep), 0);
      const commOliVendu = statsList.reduce((s, v) => s + Number(v.commission_olivier), 0);
      const commOliConf = confirmedList.reduce((s, v) => s + Number(v.commission_olivier), 0);

      document.getElementById('dash-total-ventes').textContent = fmtMoney(totVentes);
      const elDashVConf = document.getElementById('dash-total-ventes-confirme');
      if (elDashVConf) elDashVConf.textContent = fmtMoney(totVentesConf);

      document.getElementById('dash-total-comm-rep').textContent = fmtMoney(isConfirmedOnly ? commRepConf : commRepVendu);
      const elDashRepV = document.getElementById('dash-comm-rep-vendu');
      if (elDashRepV) elDashRepV.textContent = fmtMoney(commRepVendu);
      const elDashRepC = document.getElementById('dash-comm-rep-confirme');
      if (elDashRepC) elDashRepC.textContent = fmtMoney(commRepConf);

      document.getElementById('dash-total-comm-olivier').textContent = fmtMoney(isConfirmedOnly ? commOliConf : commOliVendu);
      const elDashOliV = document.getElementById('dash-comm-oli-vendu');
      if (elDashOliV) elDashOliV.textContent = fmtMoney(commOliVendu);
      const elDashOliC = document.getElementById('dash-comm-oli-confirme');
      if (elDashOliC) elDashOliC.textContent = fmtMoney(commOliConf);

      document.getElementById('dash-depots-attente').textContent = list.filter(v => v.statut_depot === 'a_confirmer').length;
      renderExecutiveForecast();
      renderFutureMonthsBreakdown();

      const byMonth = new Map();
      state.ventes.forEach(v => {
        const sYm = getSaleStatsMonth(v);
        if (!sYm) return;
        byMonth.set(sYm, (byMonth.get(sYm) || 0) + Number(v.valeur || 0));
      });
      const monthsSorted = Array.from(byMonth.entries()).sort((a, b) => a[0].localeCompare(b[0]));
      const maxVal = Math.max(...monthsSorted.map(x => x[1]), 1);

      document.getElementById('dash-monthly-chart').innerHTML = monthsSorted.length
        ? monthsSorted.map(([ym, total]) => {
            const hPct = Math.max(10, Math.round((total / maxVal) * 100));
            const isSelected = dashPeriod.mode === 'future' ? (ym > currentMonthStr) : (m === ym);
            return `
              <button onclick="selectDashMonth('${ym}')" class="group flex flex-col items-center gap-2 text-center focus:outline-none">
                <span class="text-[10px] font-bold text-slate-500 group-hover:text-indigo-600 tabular-nums">${Math.round(total / 1000)}k$</span>
                <div class="w-full h-28 bg-slate-100 rounded-xl flex items-end p-1">
                  <div class="w-full rounded-lg transition-all ${isSelected ? 'bg-emerald-500' : 'bg-indigo-600 group-hover:bg-indigo-500'}" style="height: ${hPct}%"></div>
                </div>
                <span class="text-[11px] font-bold ${isSelected ? 'text-emerald-600' : 'text-slate-600'}">${esc(formatYearMonthLabelFr(ym, true))}</span>
              </button>
            `;
          }).join('')
        : `<p class="col-span-full text-xs text-slate-400">Aucune donnée mensuelle disponible.</p>`;

      const dashRepEntities = [
        ...state.representants.map(r => ({ ...r, _isArchived: false })),
        ...Object.values(archivedRepsStore || {})
          .filter(a => a && a.nom && !state.representants.some(r => r.nom.toLowerCase() === a.nom.toLowerCase()))
          .map(a => ({ id: a.id || `arch_rep_${a.nom}`, nom: a.nom, _isArchived: true }))
      ];
      const matchDashRepSale = (v, r) => {
        if (!r._isArchived) return String(v.representant_id) === String(r.id);
        return resolveSaleArchivedRepName(v).toLowerCase() === r.nom.toLowerCase();
      };
      const rankedReps = dashRepEntities.map(r => {
        const subAll = statsList.filter(v => matchDashRepSale(v, r));
        const subComm = activeCommList.filter(v => matchDashRepSale(v, r));
        const ventesVol = subAll.reduce((s, v) => s + Number(v.valeur || 0), 0);
        const commRepVol = subComm.reduce((s, v) => s + Number(v.commission_rep || 0), 0);
        const commOliVol = subComm.reduce((s, v) => s + Number(v.commission_olivier || 0), 0);
        const commBaseVol = subComm.reduce((s, v) => s + Number(v.valeur || 0), 0);
        const margePct = commBaseVol > 0 ? ((commOliVol / commBaseVol) * 100) : 0;
        return { r, ventesVol, commRepVol, commOliVol, margePct };
      }).filter(item => !item.r._isArchived || item.ventesVol > 0 || item.commRepVol !== 0 || item.commOliVol !== 0)
        .sort((a, b) => {
          if (Math.abs(b.commOliVol - a.commOliVol) > 0.005) return b.commOliVol - a.commOliVol;
          if (Math.abs(b.margePct - a.margePct) > 0.005) return b.margePct - a.margePct;
          return b.ventesVol - a.ventesVol;
        });

      const repMedals = ['🥇', '🥈', '🥉'];
      document.getElementById('tbody-dash-rep').innerHTML = rankedReps.map((item, idx) => {
        const rankTag = item.commOliVol > 0 && idx < 3
          ? `<span class="text-sm leading-none">${repMedals[idx]}</span>`
          : `<span class="text-[11px] font-bold text-slate-400 w-4 text-center">#${idx + 1}</span>`;
        return `
          <tr class="h-11 ${idx === 0 && item.commOliVol > 0 ? 'bg-emerald-50/40' : ''}">
            <td class="py-2.5 pr-2 font-semibold text-slate-900 whitespace-nowrap align-middle">
              <div class="inline-flex items-center gap-1.5">
                ${rankTag}
                <span>${esc(item.r.nom)}</span>
                ${item.r._isArchived ? `<span class="px-1.5 py-0.2 text-[9px] font-bold rounded bg-slate-100 text-slate-500">Archivé</span>` : ''}
              </div>
            </td>
            <td class="py-2.5 px-2 text-right tabular-nums font-medium whitespace-nowrap align-middle">${fmtMoney(item.ventesVol)}</td>
            <td class="py-2.5 px-2 text-right tabular-nums text-amber-600 font-semibold whitespace-nowrap align-middle">${fmtMoney(item.commRepVol)}</td>
            <td class="py-2.5 px-2 text-right tabular-nums text-emerald-600 font-extrabold whitespace-nowrap align-middle">${fmtMoney(item.commOliVol)}</td>
            <td class="py-2.5 px-2 text-right tabular-nums whitespace-nowrap align-middle">
              <span class="px-1.5 py-0.5 text-[11px] font-bold rounded-md ${item.margePct > 0 ? 'bg-emerald-50 text-emerald-700 border border-emerald-200/80' : 'text-slate-400'}">
                ${item.margePct > 0 ? `${item.margePct.toFixed(2)} %` : '—'}
              </span>
            </td>
            <td class="py-2.5 pl-2 text-right whitespace-nowrap align-middle">
              ${item.r._isArchived ? '<span class="text-[11px] text-slate-400">—</span>' : `<button onclick="openPayslipModal('${item.r.id}')" class="px-2.5 py-1 text-[11px] font-semibold bg-slate-100 hover:bg-indigo-600 hover:text-white text-slate-700 rounded-lg transition">Relevé</button>`}
            </td>
          </tr>
        `;
      }).join('');

      const dashEntEntities = [
        ...state.entreprises.map(e => ({ ...e, _isArchived: false })),
        ...Object.values(archivedEntreprisesStore || {})
          .filter(a => a && a.nom && !state.entreprises.some(e => e.nom.toLowerCase() === a.nom.toLowerCase()))
          .map(a => ({ id: a.id || `arch_ent_${a.nom}`, nom: a.nom, _isArchived: true }))
      ];
      const matchDashEntSale = (v, e) => {
        if (!e._isArchived) return String(v.entreprise_id) === String(e.id);
        return resolveSaleArchivedEntName(v).toLowerCase() === e.nom.toLowerCase();
      };
      const rankedEntreprises = dashEntEntities.map(e => {
        const subAll = statsList.filter(v => matchDashEntSale(v, e));
        const subComm = activeCommList.filter(v => matchDashEntSale(v, e));
        const ventesVol = subAll.reduce((s, v) => s + Number(v.valeur || 0), 0);
        const commRepVol = subComm.reduce((s, v) => s + Number(v.commission_rep || 0), 0);
        const commOliVol = subComm.reduce((s, v) => s + Number(v.commission_olivier || 0), 0);
        const commBaseVol = subComm.reduce((s, v) => s + Number(v.valeur || 0), 0);
        const margePct = commBaseVol > 0 ? ((commOliVol / commBaseVol) * 100) : 0;
        return { e, ventesVol, commRepVol, commOliVol, margePct };
      }).filter(item => !item.e._isArchived || item.ventesVol > 0 || item.commRepVol !== 0 || item.commOliVol !== 0)
        .sort((a, b) => {
        if (Math.abs(b.commOliVol - a.commOliVol) > 0.005) return b.commOliVol - a.commOliVol;
        if (Math.abs(b.margePct - a.margePct) > 0.005) return b.margePct - a.margePct;
        return b.ventesVol - a.ventesVol;
      });

      const entMedals = ['🥇', '🥈', '🥉'];
      document.getElementById('tbody-dash-ent').innerHTML = rankedEntreprises.map((item, idx) => {
        const rankTag = item.commOliVol > 0 && idx < 3
          ? `<span class="text-sm leading-none">${entMedals[idx]}</span>`
          : `<span class="text-[11px] font-bold text-slate-400 w-4 text-center">#${idx + 1}</span>`;
        return `
          <tr class="h-11 ${idx === 0 && item.commOliVol > 0 ? 'bg-emerald-50/40' : ''}">
            <td class="py-2.5 pr-2 whitespace-nowrap align-middle">
              <div class="inline-flex items-center gap-1.5">
                ${rankTag}
                ${entBadge(item.e.nom)}
              </div>
            </td>
            <td class="py-2.5 px-2 text-right tabular-nums font-medium whitespace-nowrap align-middle">${fmtMoney(item.ventesVol)}</td>
            <td class="py-2.5 px-2 text-right tabular-nums text-amber-600 font-semibold whitespace-nowrap align-middle">${fmtMoney(item.commRepVol)}</td>
            <td class="py-2.5 px-2 text-right tabular-nums text-emerald-600 font-extrabold whitespace-nowrap align-middle">${fmtMoney(item.commOliVol)}</td>
            <td class="py-2.5 pl-2 text-right tabular-nums whitespace-nowrap align-middle">
              <span class="px-1.5 py-0.5 text-[11px] font-bold rounded-md ${item.margePct > 0 ? 'bg-emerald-50 text-emerald-700 border border-emerald-200/80' : 'text-slate-400'}">
                ${item.margePct > 0 ? `${item.margePct.toFixed(2)} %` : '—'}
              </span>
            </td>
          </tr>
        `;
      }).join('');
    }

    function getSaleTargetFutureMonth(v) {
      if (!v) return '';
      const dStr = (v.date_vente || '').trim().slice(0, 10);
      const commYm = (v.mois_commission || '').trim().slice(0, 7);
      const dateYm = dStr.length >= 7 ? dStr.slice(0, 7) : '';
      if (isValidYearMonth(commYm) && commYm > currentMonthStr) return commYm;
      if (isValidYearMonth(dateYm) && dateYm > currentMonthStr) return dateYm;
      return '';
    }

    function renderFutureMonthsBreakdown() {
      const kpisBox = document.getElementById('dash-future-kpis');
      const tbody = document.getElementById('tbody-dash-future-reps');
      const mobileBox = document.getElementById('mobile-dash-future-reps');
      const monthSel = document.getElementById('dash-future-month-filter');
      if (!kpisBox || !tbody || !mobileBox) return;

      const allFutureSales = (state.ventes || []).filter(v => Boolean(getSaleTargetFutureMonth(v)));
      const futureMonthsSet = new Set(allFutureSales.map(v => getSaleTargetFutureMonth(v)).filter(Boolean));
      const futureMonthsSorted = Array.from(futureMonthsSet).sort();

      if (monthSel) {
        const prevVal = monthSel.value || 'ALL';
        monthSel.innerHTML = `<option value="ALL">Tous les mois futurs (${futureMonthsSorted.length})</option>` +
          futureMonthsSorted.map(ym => `<option value="${ym}">${esc(formatYearMonthLabelFr(ym))}</option>`).join('');
        monthSel.value = (prevVal === 'ALL' || futureMonthsSorted.includes(prevVal)) ? prevVal : 'ALL';
      }

      const selectedFutureYm = monthSel?.value || 'ALL';
      const scopedFutureSales = selectedFutureYm === 'ALL'
        ? allFutureSales
        : allFutureSales.filter(v => getSaleTargetFutureMonth(v) === selectedFutureYm);

      const totCount = scopedFutureSales.length;
      const totVol = scopedFutureSales.reduce((s, v) => s + Number(v.valeur || 0), 0);
      const totCommRep = scopedFutureSales.reduce((s, v) => s + Number(v.commission_rep || 0), 0);
      const totCommOli = scopedFutureSales.reduce((s, v) => s + Number(v.commission_olivier || 0), 0);

      kpisBox.innerHTML = `
        <div class="bg-slate-50/90 rounded-xl border border-slate-200/80 p-3.5">
          <div class="text-[11px] font-semibold text-slate-500">Ventes planifiées</div>
          <div class="text-xl font-extrabold text-slate-900 mt-1 tabular-nums">${totCount} <span class="text-xs font-semibold text-slate-400">dossier(s)</span></div>
          <div class="text-[10px] text-slate-400 mt-0.5">${selectedFutureYm === 'ALL' ? `Sur ${futureMonthsSorted.length} mois futur(s)` : esc(formatYearMonthLabelFr(selectedFutureYm))}</div>
        </div>
        <div class="bg-slate-50/90 rounded-xl border border-slate-200/80 p-3.5">
          <div class="text-[11px] font-semibold text-slate-500">Volume ventes futures</div>
          <div class="text-xl font-extrabold text-indigo-700 mt-1 tabular-nums">${fmtMoney(totVol)}</div>
          <div class="text-[10px] text-slate-400 mt-0.5">Contrats reportés / à venir</div>
        </div>
        <div class="bg-amber-50/70 rounded-xl border border-amber-200/80 p-3.5">
          <div class="text-[11px] font-semibold text-amber-900">Commissions Rep à venir</div>
          <div class="text-xl font-extrabold text-amber-700 mt-1 tabular-nums">${fmtMoney(totCommRep)}</div>
          <div class="text-[10px] text-amber-800/80 mt-0.5">À verser aux représentants</div>
        </div>
        <div class="bg-emerald-50/70 rounded-xl border border-emerald-200/80 p-3.5">
          <div class="text-[11px] font-semibold text-emerald-900">Commissions Olivier à venir</div>
          <div class="text-xl font-extrabold text-emerald-700 mt-1 tabular-nums">${fmtMoney(totCommOli)}</div>
          <div class="text-[10px] text-emerald-800/80 mt-0.5">Votre part sur mois ultérieurs</div>
        </div>
      `;

      const repSummaries = state.representants.map(r => {
        const repSales = scopedFutureSales.filter(v => String(v.representant_id) === String(r.id));
        const byYmMap = new Map();
        repSales.forEach(v => {
          const fYm = getSaleTargetFutureMonth(v);
          if (!fYm) return;
          const cur = byYmMap.get(fYm) || { count: 0, val: 0, cRep: 0, cOli: 0 };
          cur.count += 1;
          cur.val += Number(v.valeur || 0);
          cur.cRep += Number(v.commission_rep || 0);
          cur.cOli += Number(v.commission_olivier || 0);
          byYmMap.set(fYm, cur);
        });
        const monthsBreakdown = Array.from(byYmMap.entries()).sort((a, b) => a[0].localeCompare(b[0]));
        return {
          rep: r,
          count: repSales.length,
          vol: repSales.reduce((s, v) => s + Number(v.valeur || 0), 0),
          cRep: repSales.reduce((s, v) => s + Number(v.commission_rep || 0), 0),
          cOli: repSales.reduce((s, v) => s + Number(v.commission_olivier || 0), 0),
          monthsBreakdown
        };
      }).sort((a, b) => b.count - a.count || b.vol - a.vol);

      if (!scopedFutureSales.length) {
        tbody.innerHTML = `<tr><td colspan="6" class="py-8 text-center text-xs text-slate-400">Aucune vente planifiée dans un mois ultérieur pour le moment.</td></tr>`;
        mobileBox.innerHTML = `<div class="py-6 text-center text-xs text-slate-400 bg-slate-50 rounded-xl border border-slate-200/80">Aucune vente planifiée dans un mois ultérieur.</div>`;
        return;
      }

      tbody.innerHTML = repSummaries.map(item => {
        const pillsHtml = item.monthsBreakdown.length
          ? item.monthsBreakdown.map(([ym, info]) => `
              <button type="button" onclick="selectDashMonth('${ym}')" title="Filtrer le tableau de bord sur ${esc(formatYearMonthLabelFr(ym))}" class="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-indigo-50 hover:bg-indigo-100 text-indigo-900 border border-indigo-200/80 text-[11px] font-bold transition tabular-nums">
                <span>${esc(formatYearMonthLabelFr(ym, true))} :</span>
                <span class="px-1.5 py-0.2 rounded bg-indigo-600 text-white text-[10px]">${info.count}</span>
                <span class="text-slate-600 font-semibold">(${fmtMoney0(info.val)})</span>
              </button>
            `).join(' ')
          : `<span class="text-xs text-slate-400 italic">Aucune vente future</span>`;

        return `
          <tr class="${item.count === 0 ? 'opacity-50' : 'hover:bg-slate-50/90'} transition">
            <td class="py-3 px-4 font-extrabold text-slate-900 whitespace-nowrap">${esc(item.rep.nom)}</td>
            <td class="py-3 px-4"><div class="flex flex-wrap gap-1.5">${pillsHtml}</div></td>
            <td class="py-3 px-3 text-center tabular-nums">
              <span class="inline-flex items-center justify-center min-w-[28px] px-2 py-0.5 rounded-full text-xs font-extrabold ${item.count > 0 ? 'bg-indigo-100 text-indigo-800' : 'bg-slate-100 text-slate-500'}">${item.count}</span>
            </td>
            <td class="py-3 px-4 text-right tabular-nums font-bold text-slate-900 whitespace-nowrap">${fmtMoney(item.vol)}</td>
            <td class="py-3 px-4 text-right tabular-nums font-extrabold text-amber-600 whitespace-nowrap">${fmtMoney(item.cRep)}</td>
            <td class="py-3 px-4 text-right tabular-nums font-extrabold text-emerald-600 whitespace-nowrap">${fmtMoney(item.cOli)}</td>
          </tr>
        `;
      }).join('');

      mobileBox.innerHTML = repSummaries.filter(x => x.count > 0).map(item => `
        <div class="bg-slate-50/80 rounded-xl border border-slate-200/90 p-3.5 space-y-2.5">
          <div class="flex items-center justify-between">
            <span class="font-extrabold text-sm text-slate-900">${esc(item.rep.nom)}</span>
            <span class="px-2.5 py-0.5 rounded-full text-xs font-extrabold bg-indigo-100 text-indigo-800 tabular-nums">${item.count} vente(s)</span>
          </div>
          <div class="flex flex-wrap gap-1.5">
            ${item.monthsBreakdown.map(([ym, info]) => `
              <button type="button" onclick="selectDashMonth('${ym}')" class="inline-flex items-center gap-1 px-2 py-0.5 rounded-md bg-white border border-indigo-200 text-[11px] font-bold text-indigo-900 tabular-nums">
                <span>${esc(formatYearMonthLabelFr(ym, true))} :</span>
                <span class="text-indigo-600">${info.count}</span>
                <span class="text-slate-500 font-normal">(${fmtMoney0(info.val)})</span>
              </button>
            `).join('')}
          </div>
          <div class="grid grid-cols-3 gap-2 pt-2 border-t border-slate-200/80 text-[11px] tabular-nums">
            <div>
              <span class="text-slate-400 block">Volume</span>
              <strong class="text-slate-900">${fmtMoney0(item.vol)}</strong>
            </div>
            <div>
              <span class="text-slate-400 block">Comm. Rep</span>
              <strong class="text-amber-700">${fmtMoney(item.cRep)}</strong>
            </div>
            <div class="text-right">
              <span class="text-slate-400 block">Comm. Olivier</span>
              <strong class="text-emerald-700">${fmtMoney(item.cOli)}</strong>
            </div>
          </div>
        </div>
      `).join('');

      const contractsBox = document.getElementById('dash-future-contracts-list-box');
      if (contractsBox) {
        const sortedContracts = [...scopedFutureSales].sort((a, b) => {
          const mA = getSaleTargetFutureMonth(a);
          const mB = getSaleTargetFutureMonth(b);
          if (mA !== mB) return mA.localeCompare(mB);
          return String(a.date_vente || '').localeCompare(String(b.date_vente || ''));
        });

        contractsBox.innerHTML = `
          <div class="flex items-center justify-between px-1 pt-2">
            <span class="text-xs font-extrabold uppercase tracking-wider text-slate-600">📋 Liste détaillée des contrats prévus (${sortedContracts.length})</span>
            <span class="text-[11px] text-slate-400">Vendus maintenant • Réalisation future</span>
          </div>
          <div class="overflow-x-auto border border-slate-200/80 rounded-xl">
            <table class="w-full text-left border-collapse text-xs">
              <thead>
                <tr class="bg-slate-50/90 border-b border-slate-200 text-[10px] font-bold uppercase text-slate-500">
                  <th class="py-2.5 px-3">Mois prévu</th>
                  <th class="py-2.5 px-3">Vendu le</th>
                  <th class="py-2.5 px-3">Représentant</th>
                  <th class="py-2.5 px-3">Entreprise & Client</th>
                  <th class="py-2.5 px-3 text-right">Montant</th>
                  <th class="py-2.5 px-3 text-right text-amber-700">Comm. Rep</th>
                  <th class="py-2.5 px-3 text-right text-emerald-700">Comm. Olivier</th>
                </tr>
              </thead>
              <tbody class="divide-y divide-slate-100">
                ${sortedContracts.map(v => {
                  const fYm = getSaleTargetFutureMonth(v);
                  const repN = state.representants.find(r => String(r.id) === String(v.representant_id))?.nom || v.rep_nom || '—';
                  const entN = state.entreprises.find(e => String(e.id) === String(v.entreprise_id))?.nom || v.entreprise_nom || '—';
                  const cliN = state.clients.find(c => String(c.id) === String(v.client_id))?.nom || v.client_nom || '—';
                  return `
                    <tr class="hover:bg-slate-50/80 transition">
                      <td class="py-2.5 px-3 whitespace-nowrap">
                        <span class="px-2 py-0.5 rounded-md bg-indigo-50 text-indigo-800 border border-indigo-200 font-extrabold">${esc(formatYearMonthLabelFr(fYm, true))}</span>
                      </td>
                      <td class="py-2.5 px-3 text-slate-500 tabular-nums whitespace-nowrap">${esc(v.date_vente || '—')}</td>
                      <td class="py-2.5 px-3 font-bold text-slate-900 whitespace-nowrap">${esc(repN)}</td>
                      <td class="py-2.5 px-3">
                        <div class="flex items-center gap-1.5 flex-wrap">
                          ${entBadge(entN)}
                          <span class="font-semibold text-slate-800">${esc(cliN)}</span>
                        </div>
                      </td>
                      <td class="py-2.5 px-3 text-right font-extrabold text-slate-900 tabular-nums whitespace-nowrap">${fmtMoney(v.valeur)}</td>
                      <td class="py-2.5 px-3 text-right font-bold text-amber-700 tabular-nums whitespace-nowrap">${fmtMoney(v.commission_rep)}</td>
                      <td class="py-2.5 px-3 text-right font-extrabold text-emerald-700 tabular-nums whitespace-nowrap">${fmtMoney(v.commission_olivier)}</td>
                    </tr>
                  `;
                }).join('')}
              </tbody>
            </table>
          </div>
        `;
      }
    }

    function getRepMonthlyMileage(rep, ym = null) {
      const store = lsGetJson('korpo_rep_mileage_v1', '{}');
      const entries = Object.values(store)
        .filter(it => it && (String(it.repId) === String(rep.id) || (it.repNom && it.repNom === rep.nom)) && (!ym || it.ym === ym))
        .sort((a, b) => String(a.date || '').localeCompare(String(b.date || '')));
      const totalKm = Math.round(entries.reduce((s, it) => s + Number(it.totalKm || 0), 0) * 10) / 10;
      return { entries, totalKm };
    }

    window.openPayslipModal = (repId) => {
      const rep = state.representants.find(r => r.id === repId);
      if (!rep) return;
      const dashPeriod = readPeriodFilterState('dash-filter');
      const m = dashPeriod.label;
      const isOlivier = rep.role === 'admin';
      const isConfirmedOnly = commDisplayMode === 'confirme';
      const sales = state.ventes.filter(v => (isOlivier ? true : v.representant_id === repId) && matchesSalePeriod(v, dashPeriod, true) && (!isConfirmedOnly || v.statut_depot === 'depot_recu'));

      const totalVentes = sales.reduce((s, v) => s + Number(v.valeur), 0);
      const totalComm = sales.reduce((s, v) => s + Number(isOlivier ? v.commission_olivier : v.commission_rep), 0);
      const totalPaye = sales.filter(v => v.statut_paiement === 'paye').reduce((s, v) => s + Number(isOlivier ? v.commission_olivier : v.commission_rep), 0);
      const mileageYm = dashPeriod.singleMonth || (dashPeriod.mode === 'all' ? null : currentMonthStr);
      const mileageInfo = getRepMonthlyMileage(rep, mileageYm);
      const totalKmFmt = mileageInfo.totalKm.toLocaleString('fr-CA', { maximumFractionDigits: 1 });

      document.getElementById('payslip-print-area').innerHTML = `
        <div class="flex items-start justify-between border-b border-slate-200 pb-5">
          <div>
            <div class="text-xl font-extrabold text-slate-900">Korpo Solutions</div>
            <div class="text-xs text-slate-500 mt-0.5">Relevé détaillé des commissions — <strong>${isConfirmedOnly ? 'Sur total confirmé (Dépôt reçu)' : 'Sur total vendu'}</strong> (${sales.length} dossier(s))</div>
          </div>
          <div class="text-right">
            <div class="text-sm font-bold text-slate-900">${esc(rep.nom)}</div>
            <div class="text-xs text-slate-500">${esc(rep.courriel)}</div>
            <div class="text-xs font-semibold text-indigo-600 mt-1">Période : ${m || 'Cumul complet'}</div>
          </div>
        </div>

        <div class="grid grid-cols-2 sm:grid-cols-4 gap-3">
          <div class="bg-slate-50 p-4 rounded-xl border border-slate-200">
            <div class="text-xs text-slate-500">Volume des ventes</div>
            <div class="text-lg font-extrabold text-slate-900 mt-1 tabular-nums">${fmtMoney(totalVentes)}</div>
          </div>
          <div class="bg-emerald-50 p-4 rounded-xl border border-emerald-200">
            <div class="text-xs text-emerald-700">Total commissions</div>
            <div class="text-lg font-extrabold text-emerald-700 mt-1 tabular-nums">${fmtMoney(totalComm)}</div>
          </div>
          <div class="bg-indigo-50 p-4 rounded-xl border border-indigo-200">
            <div class="text-xs text-indigo-700">Solde à verser</div>
            <div class="text-lg font-extrabold text-indigo-700 mt-1 tabular-nums">${fmtMoney(totalComm - totalPaye)}</div>
          </div>
          <div class="bg-sky-50 p-4 rounded-xl border border-sky-200">
            <div class="text-xs text-sky-700">Kilométrage validé</div>
            <div class="text-lg font-extrabold text-sky-800 mt-1 tabular-nums">${totalKmFmt} km</div>
          </div>
        </div>

        <table class="w-full text-left border-collapse text-xs">
          <thead>
            <tr class="border-b-2 border-slate-200 text-slate-500 uppercase">
              <th class="py-2">Date</th>
              <th class="py-2">Entreprise</th>
              <th class="py-2">Client</th>
              <th class="py-2 text-right">Vente</th>
              <th class="py-2 text-right">Commission</th>
              <th class="py-2 text-right">Statut</th>
            </tr>
          </thead>
          <tbody class="divide-y divide-slate-100">
            ${sales.map(v => {
              const ent = state.entreprises.find(e => e.id === v.entreprise_id)?.nom || resolveSaleArchivedEntName(v) || '—';
              const cli = state.clients.find(c => c.id === v.client_id)?.nom || '—';
              const comm = isOlivier ? v.commission_olivier : v.commission_rep;
              return `
                <tr>
                  <td class="py-2">${v.date_vente}</td>
                  <td class="py-2 font-semibold">${esc(ent)}</td>
                  <td class="py-2">${esc(cli)}</td>
                  <td class="py-2 text-right tabular-nums">${fmtMoney(v.valeur)}</td>
                  <td class="py-2 text-right font-bold text-emerald-700 tabular-nums">${fmtMoney(comm)}</td>
                  <td class="py-2 text-right">${v.statut_paiement === 'paye' ? 'Payé' : 'À payer'}</td>
                </tr>
              `;
            }).join('')}
          </tbody>
        </table>

        ${mileageInfo.entries.length > 0 ? `
          <div class="pt-4 border-t border-slate-200 space-y-2.5">
            <div class="flex items-center justify-between">
              <div class="text-xs font-extrabold uppercase tracking-wider text-slate-700">🧭 Registre des déplacements & adresses visitées</div>
              <div class="text-xs font-extrabold text-sky-700 tabular-nums">Total : ${totalKmFmt} km (${mileageInfo.entries.length} trajet(s))</div>
            </div>
            <table class="w-full text-left border-collapse text-xs">
              <thead>
                <tr class="border-b-2 border-slate-200 text-slate-500 uppercase">
                  <th class="py-2">Date</th>
                  <th class="py-2">Itinéraire complet (Départ → Arrêts)</th>
                  <th class="py-2 text-right">Distance</th>
                </tr>
              </thead>
              <tbody class="divide-y divide-slate-100">
                ${mileageInfo.entries.map(tr => `
                  <tr>
                    <td class="py-2 whitespace-nowrap align-top">${esc(tr.date)}</td>
                    <td class="py-2 text-slate-700"><strong>Départ :</strong> ${esc(tr.depart)} → ${(tr.arrets || []).map(a => esc(a)).join(' → ')}${tr.retour ? ` → <em>Retour (${esc(tr.depart)})</em>` : ''}</td>
                    <td class="py-2 text-right font-bold text-sky-700 tabular-nums whitespace-nowrap align-top">${Number(tr.totalKm || 0).toLocaleString('fr-CA', { maximumFractionDigits: 1 })} km</td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>
        ` : ''}
      `;

      const btnSendSingleEmail = document.getElementById('btn-send-payslip-email');
      if (btnSendSingleEmail) {
        btnSendSingleEmail.onclick = async () => {
          const ym = dashPeriod.singleMonth || currentMonthStr;
          await sendStatementsForReps([rep], ym, isConfirmedOnly);
        };
      }

      document.getElementById('btn-export-payslip-csv').onclick = () => {
        const headers = ['Date', 'Mois', 'Entreprise', 'Client', 'Valeur_vente', 'Commission', 'Statut_paiement'];
        const rows = sales.map(v => {
          const ent = state.entreprises.find(e => e.id === v.entreprise_id)?.nom || '';
          const cli = state.clients.find(c => c.id === v.client_id)?.nom || '';
          const comm = isOlivier ? v.commission_olivier : v.commission_rep;
          return [v.date_vente, v.mois_commission, csvCell(ent), csvCell(cli), v.valeur, comm, v.statut_paiement].join(',');
        });
        const blob = new Blob(['\ufeff' + [headers.join(','), ...rows].join('\n')], { type: 'text/csv;charset=utf-8;' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = `releve-${rep.nom.toLowerCase().replace(/\s+/g, '-')}-${m || 'complet'}.csv`;
        a.click();
      };

      document.getElementById('payslip-modal').classList.remove('hidden');
    };

    window.selectDashMonth = (ym) => {
      const perSel = document.getElementById('dash-filter-period');
      const input = document.getElementById('dash-filter-month');
      if (perSel && input) {
        if (perSel.value === 'specific_month' && input.value === ym) {
          perSel.value = 'all';
        } else {
          perSel.value = 'specific_month';
          input.value = ym;
        }
      }
      syncPeriodInputsVisibility('dash-filter', false);
      renderAdminDashboard();
    };

    // =========================================================================
    // ENVOI AUTOMATIQUE DES RELEVÉS DE COMMISSIONS PAR COURRIEL
    // =========================================================================
    const SUPABASE_EDGE_FUNCTION_CODE = ["// Supabase Edge Function : send-monthly-statements (Version compl\u00e8te)", "// D\u00e9clench\u00e9e automatiquement chaque fin de mois (pg_cron) ou depuis le portail Korpo", "import { serve } from \"https://deno.land/std@0.177.0/http/server.ts\";", "import { createClient } from \"https://esm.sh/@supabase/supabase-js@2.45.4\";", "", "const RESEND_API_KEY = Deno.env.get(\"RESEND_API_KEY\")!;", "const SUPABASE_URL = Deno.env.get(\"SUPABASE_URL\")!;", "const SUPABASE_SERVICE_ROLE_KEY = Deno.env.get(\"SUPABASE_SERVICE_ROLE_KEY\")!;", "const FROM_EMAIL = Deno.env.get(\"KORPO_FROM_EMAIL\") || \"Korpo Solutions <commissions@korposolutions.com>\";", "", "const fmtMoney = (n: number) =>", "  new Intl.NumberFormat(\"fr-CA\", { style: \"currency\", currency: \"CAD\" }).format(Number(n || 0));", "", "const fmtMoney0 = (n: number) =>", "  new Intl.NumberFormat(\"fr-CA\", { style: \"currency\", currency: \"CAD\", maximumFractionDigits: 0 }).format(Number(n || 0));", "", "const fmtPct = (n: number) => `${(Number(n || 0) * 100).toFixed(2).replace(/\\.00$/, \"\")} %`;", "", "const escHtml = (s: any) =>", "  String(s ?? \"\").replace(/[&<>\"']/g, (c) => ({ \"&\": \"&amp;\", \"<\": \"&lt;\", \">\": \"&gt;\", '\"': \"&quot;\", \"'\": \"&#39;\" }[c] || c));", "", "serve(async (req) => {", "  if (req.method === \"OPTIONS\") {", "    return new Response(\"ok\", { headers: { \"Access-Control-Allow-Origin\": \"*\", \"Access-Control-Allow-Headers\": \"authorization, x-client-info, apikey, content-type\" } });", "  }", "", "  try {", "    const body = await req.json().catch(() => ({}));", "    const now = new Date();", "    const defaultYm = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, \"0\")}`;", "    const targetYm = body.month || defaultYm;", "    const confirmedOnly = Boolean(body.confirmedOnly);", "    const customStatements = Array.isArray(body.statements) ? body.statements : null;", "", "    let statementsToSend = customStatements;", "", "    // Si appel\u00e9 par le d\u00e9clencheur automatique (cron), lire directement toutes les donn\u00e9es dans Supabase", "    if (!statementsToSend) {", "      const sb = createClient(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY);", "      const [{ data: reps }, { data: ventes }, { data: entreprises }, { data: clients }, { data: kpiRows }] = await Promise.all([", "        sb.from(\"representants\").select(\"*\"),", "        sb.from(\"ventes\").select(\"*\").eq(\"mois_commission\", targetYm),", "        sb.from(\"entreprises\").select(\"id,nom\"),", "        sb.from(\"clients\").select(\"id,nom\"),", "        sb.from(\"kpi_hebdo\").select(\"*\").in(\"semaine_index\", [-1, -7]),", "      ]);", "", "      const entMap = new Map((entreprises || []).map((e: any) => [String(e.id), e.nom]));", "      const cliMap = new Map((clients || []).map((c: any) => [String(c.id), c.nom]));", "", "      statementsToSend = (reps || []).map((rep: any) => {", "        const isAdmin = rep.role === \"admin\";", "        const repSales = (ventes || []).filter((v: any) => {", "          if (!isAdmin && String(v.representant_id) !== String(rep.id)) return false;", "          if (confirmedOnly && v.statut_depot !== \"depot_recu\") return false;", "          return true;", "        });", "", "        const totalVentes = repSales.reduce((s: number, v: any) => s + Number(v.valeur || 0), 0);", "        const confirmedSales = repSales.filter((v: any) => v.statut_depot === \"depot_recu\");", "        const totalVentesConfirme = confirmedSales.reduce((s: number, v: any) => s + Number(v.valeur || 0), 0);", "", "        const totalComm = repSales.reduce((s: number, v: any) => s + Number(isAdmin ? v.commission_olivier : v.commission_rep), 0);", "        const totalCommConfirme = confirmedSales.reduce((s: number, v: any) => s + Number(isAdmin ? v.commission_olivier : v.commission_rep), 0);", "        const totalPaye = repSales.filter((v: any) => v.statut_paiement === \"paye\").reduce((s: number, v: any) => s + Number(isAdmin ? v.commission_olivier : v.commission_rep), 0);", "", "        // R\u00e9cup\u00e9rer l'objectif mensuel du repr\u00e9sentant", "        const tRow = (kpiRows || []).find((r: any) => Number(r.semaine_index) === -1 && r.representant_nom === rep.nom && r.mois_commission === targetYm)", "          || (kpiRows || []).find((r: any) => Number(r.semaine_index) === -1 && r.representant_nom === rep.nom && r.mois_commission === \"ALL\");", "        const salesTarget = tRow && Number(tRow.booke) > 0 ? Number(tRow.booke) : 150000;", "        const targetPct = salesTarget > 0 ? Math.round((totalVentes / salesTarget) * 100) : 0;", "", "        const dossiers = repSales.map((v: any) => {", "          const rawNote = String(v.note || \"\");", "          const typeKpi = rawNote.toUpperCase().includes(\"[SUIVI]\") ? \"SUIVI\" : \"OTS\";", "          const cleanNote = rawNote.replace(/^\\[(OTS|SUIVI)\\]\\s*/i, \"\").trim();", "          return {", "            date: v.date_vente,", "            entreprise: entMap.get(String(v.entreprise_id)) || \"\u2014\",", "            client: cliMap.get(String(v.client_id)) || \"\u2014\",", "            typeKpi,", "            valeur: Number(v.valeur || 0),", "            taux: Number(isAdmin ? v.taux_olivier_applique : v.taux_rep_applique),", "            commission: Number(isAdmin ? v.commission_olivier : v.commission_rep),", "            depot: v.statut_depot === \"depot_recu\" ? \"Confirm\u00e9\" : \"\u00c0 confirmer\",", "            paiement: v.statut_paiement === \"paye\" ? \"Pay\u00e9\" : \"\u00c0 payer\",", "            note: cleanNote,", "          };", "        });", "", "        // R\u00e9partition par entreprise partenaire", "        const byEntMap = new Map<string, { entreprise: string; count: number; ventes: number; comm: number }>();", "        for (const d of dossiers) {", "          const cur = byEntMap.get(d.entreprise) || { entreprise: d.entreprise, count: 0, ventes: 0, comm: 0 };", "          cur.count += 1;", "          cur.ventes += d.valeur;", "          cur.comm += d.commission;", "          byEntMap.set(d.entreprise, cur);", "        }", "", "        return {", "          nom: rep.nom,", "          courriel: rep.courriel,", "          mois: targetYm,", "          confirmedOnly,", "          nbDossiers: dossiers.length,", "          nbConfirmes: confirmedSales.length,", "          nbOts: dossiers.filter((d: any) => d.typeKpi === \"OTS\").length,", "          nbSuivi: dossiers.filter((d: any) => d.typeKpi === \"SUIVI\").length,", "          totalVentes,", "          totalVentesConfirme,", "          totalComm,", "          totalCommConfirme,", "          totalPaye,", "          soldeAVerser: Math.max(0, totalComm - totalPaye),", "          soldeConfirmeAVerser: Math.max(0, totalCommConfirme - totalPaye),", "          salesTarget,", "          targetPct,", "          parEntreprise: Array.from(byEntMap.values()).sort((a, b) => b.ventes - a.ventes),", "          dossiers,", "        };", "      }).filter((st: any) => st.courriel && st.dossiers.length > 0);", "    }", "", "    const csvEsc = (v: any) => {", "      let s = String(v ?? \"\");", "      if (/^[=+\\-@\\t\\r]/.test(s)) s = \"'\" + s;", "      return /[\",;\\n\\r]/.test(s) ? `\"${s.replace(/\"/g, '\"\"')}\"` : s;", "    };", "    const toBase64Utf8 = (str: string) => {", "      const bytes = new TextEncoder().encode(\"\\uFEFF\" + str);", "      let bin = \"\";", "      for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);", "      return btoa(bin);", "    };", "", "    let sentCount = 0;", "    for (const st of statementsToSend || []) {", "      if (!st.courriel) continue;", "", "      const dossiers = Array.isArray(st.dossiers) ? st.dossiers : [];", "      const slugNom = String(st.nom || \"rep\").toLowerCase().normalize(\"NFD\").replace(/[\\u0300-\\u036f]/g, \"\").replace(/[^a-z0-9]+/g, \"-\");", "      const invDate = new Date().toISOString().slice(0, 10);", "      const edgeSeqObj: Record<string, number> = { MaClim: 107, MonSpray: 113, Toitex: 115, ...(body.companySeqMap || {}) };", "      const edgeCompMap = new Map<string, string>();", "      const fmtCompPrefix = (c: string) => { const s = String(c || \"Korpo\").replace(/[^A-Za-z0-9]/g, \"\") || \"Korpo\"; return s.charAt(0).toUpperCase() + s.slice(1).toLowerCase(); };", "      const waveCsv = st.waveCsvContent || [", "        \"Customer Name,Invoice Number,Invoice Date,Due Date,Item Name,Description,Quantity,Unit Price,Memo\",", "        ...[...dossiers].sort((a: any, b: any) => String(a.entreprise || \"\").localeCompare(String(b.entreprise || \"\"), \"fr\") || String(a.date || \"\").localeCompare(String(b.date || \"\"))).map((d: any) => {", "          const customerCompany = d.entreprise || \"Korpo Solutions\";", "          if (!edgeCompMap.has(customerCompany)) { const nxt = Number(edgeSeqObj[customerCompany] ?? 100) + 1; edgeSeqObj[customerCompany] = nxt; edgeCompMap.set(customerCompany, `${fmtCompPrefix(customerCompany)}${nxt}`); }", "          const invNum = edgeCompMap.get(customerCompany)!;", "          const desc = `${d.client}${d.note ? \" (\" + d.note + \")\" : \"\"}`.replace(/[\\t\\r\\n]+/g, \" \");", "          return [csvEsc(customerCompany), csvEsc(invNum), csvEsc(invDate), csvEsc(invDate), csvEsc(\"Commission sur vente\"), csvEsc(desc), \"1\", Number(d.commission || 0).toFixed(2), csvEsc(`Facture mois de ${st.mois || targetYm}`)].join(\",\");", "        })", "      ].join(\"\\n\");", "      const waveFilename = st.waveCsvFilename || `facture-wave-${slugNom}-${st.mois || targetYm}.csv`;", "      const attachments = dossiers.length > 0 ? [{ filename: waveFilename, content: toBase64Utf8(waveCsv) }] : [];", "      const parEnt = Array.isArray(st.parEntreprise) ? st.parEntreprise : [];", "      const nbDossiers = st.nbDossiers ?? dossiers.length;", "      const nbConfirmes = st.nbConfirmes ?? dossiers.filter((d: any) => String(d.depot).toLowerCase().includes(\"confirm\") || String(d.depot).toLowerCase().includes(\"re\u00e7u\")).length;", "      const nbOts = st.nbOts ?? dossiers.filter((d: any) => d.typeKpi === \"OTS\").length;", "      const nbSuivi = st.nbSuivi ?? dossiers.filter((d: any) => d.typeKpi === \"SUIVI\").length;", "      const totVentesConf = st.totalVentesConfirme ?? st.totalVentes;", "      const totCommConf = st.totalCommConfirme ?? st.totalComm;", "      const soldeConf = st.soldeConfirmeAVerser ?? st.soldeAVerser;", "      const salesTarget = Number(st.salesTarget || 0);", "      const targetPct = Number(st.targetPct || 0);", "      const resteTarget = Math.max(0, salesTarget - Number(st.totalVentes || 0));", "", "      const entRowsHtml = parEnt.map((e: any) => `", "        <tr>", "          <td style=\"padding:8px 10px;border-bottom:1px solid #e2e8f0;font-weight:700;color:#0f172a;\">${escHtml(e.entreprise)}</td>", "          <td style=\"padding:8px 10px;border-bottom:1px solid #e2e8f0;text-align:center;color:#475569;\">${e.count}</td>", "          <td style=\"padding:8px 10px;border-bottom:1px solid #e2e8f0;text-align:right;font-weight:600;\">${fmtMoney(e.ventes)}</td>", "          <td style=\"padding:8px 10px;border-bottom:1px solid #e2e8f0;text-align:right;font-weight:700;color:#059669;\">${fmtMoney(e.comm)}</td>", "        </tr>", "      `).join(\"\");", "", "      const rowsHtml = dossiers.map((d: any) => {", "        const isDepotOk = String(d.depot).toLowerCase().includes(\"confirm\") || String(d.depot).toLowerCase().includes(\"re\u00e7u\");", "        const isPaye = String(d.paiement).toLowerCase() === \"pay\u00e9\";", "        const depotBadge = isDepotOk", "          ? `<span style=\"display:inline-block;padding:2px 8px;border-radius:999px;background:#ecfdf5;color:#065f46;font-size:11px;font-weight:700;\">\u2713 Confirm\u00e9</span>`", "          : `<span style=\"display:inline-block;padding:2px 8px;border-radius:999px;background:#fffbeb;color:#92400e;font-size:11px;font-weight:700;\">\u23f3 \u00c0 confirmer</span>`;", "        const payeBadge = isPaye", "          ? `<span style=\"display:inline-block;padding:2px 8px;border-radius:999px;background:#eef2ff;color:#3730a3;font-size:11px;font-weight:700;\">Pay\u00e9</span>`", "          : `<span style=\"display:inline-block;padding:2px 8px;border-radius:999px;background:#f1f5f9;color:#475569;font-size:11px;font-weight:600;\">\u00c0 payer</span>`;", "        const typeTag = d.typeKpi", "          ? `<span style=\"display:inline-block;padding:1px 6px;border-radius:4px;background:${d.typeKpi === \"SUIVI\" ? \"#f3e8ff\" : \"#e0f2fe\"};color:${d.typeKpi === \"SUIVI\" ? \"#6b21a8\" : \"#0369a1\"};font-size:10px;font-weight:700;margin-left:4px;\">${escHtml(d.typeKpi)}</span>`", "          : \"\";", "        const noteLine = d.note ? `<div style=\"font-size:11px;color:#64748b;margin-top:2px;\">Note : ${escHtml(d.note)}</div>` : \"\";", "        const tauxCol = d.taux ? fmtPct(d.taux) : \"\u2014\";", "", "        return `", "          <tr>", "            <td style=\"padding:9px 8px;border-bottom:1px solid #e2e8f0;white-space:nowrap;color:#334155;\">${escHtml(d.date)}</td>", "            <td style=\"padding:9px 8px;border-bottom:1px solid #e2e8f0;font-weight:700;color:#0f172a;\">${escHtml(d.entreprise)}${typeTag}</td>", "            <td style=\"padding:9px 8px;border-bottom:1px solid #e2e8f0;color:#0f172a;\"><strong>${escHtml(d.client)}</strong>${noteLine}</td>", "            <td style=\"padding:9px 8px;border-bottom:1px solid #e2e8f0;text-align:right;white-space:nowrap;\">${fmtMoney(d.valeur)}</td>", "            <td style=\"padding:9px 8px;border-bottom:1px solid #e2e8f0;text-align:right;white-space:nowrap;color:#64748b;\">${tauxCol}</td>", "            <td style=\"padding:9px 8px;border-bottom:1px solid #e2e8f0;text-align:right;white-space:nowrap;font-weight:800;color:#059669;\">${fmtMoney(d.commission)}</td>", "            <td style=\"padding:9px 8px;border-bottom:1px solid #e2e8f0;text-align:center;white-space:nowrap;\">${depotBadge}</td>", "            <td style=\"padding:9px 8px;border-bottom:1px solid #e2e8f0;text-align:center;white-space:nowrap;\">${payeBadge}</td>", "          </tr>", "        `;", "      }).join(\"\");", "", "      const targetSectionHtml = salesTarget > 0 ? `", "        <div style=\"margin:18px 0;padding:14px 16px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;\">", "          <table style=\"width:100%;border-collapse:collapse;\">", "            <tr>", "              <td style=\"font-size:12px;font-weight:700;color:#1e293b;\">\ud83c\udfaf Progression objectif mensuel (${fmtMoney0(st.totalVentes)} / ${fmtMoney0(salesTarget)})</td>", "              <td style=\"text-align:right;font-size:13px;font-weight:800;color:${targetPct >= 100 ? \"#059669\" : \"#4f46e5\"};\">${targetPct} %</td>", "            </tr>", "          </table>", "          <div style=\"margin-top:8px;width:100%;height:10px;background:#e2e8f0;border-radius:999px;overflow:hidden;\">", "            <div style=\"height:10px;width:${Math.min(100, targetPct)}%;background:${targetPct >= 100 ? \"#10b981\" : \"#4f46e5\"};border-radius:999px;\"></div>", "          </div>", "          <div style=\"margin-top:6px;font-size:11px;color:#475569;\">", "            ${targetPct >= 100", "              ? `\ud83d\udd25 Objectif mensuel d\u00e9pass\u00e9 de + ${fmtMoney(Number(st.totalVentes || 0) - salesTarget)} !`", "              : `\u00c9cart restant avant l'objectif : ${fmtMoney(resteTarget)}`}", "          </div>", "        </div>", "      ` : \"\";", "", "      const htmlEmail = `", "        <div style=\"font-family:'Segoe UI',Arial,sans-serif;max-width:760px;margin:0 auto;color:#0f172a;background:#ffffff;\">", "          <div style=\"background:linear-gradient(135deg,#0f172a 0%,#1e1b4b 100%);color:#fff;padding:24px 28px;border-radius:14px 14px 0 0;\">", "            <div style=\"font-size:11px;text-transform:uppercase;letter-spacing:1.2px;color:#a5b4fc;font-weight:700;\">Korpo Solutions \u2014 Portail Ventes & Commissions</div>", "            <h2 style=\"margin:6px 0 0;font-size:21px;font-weight:800;\">Relev\u00e9 mensuel d\u00e9taill\u00e9 \u2014 ${escHtml(st.mois)}</h2>", "            <p style=\"margin:6px 0 0;font-size:13px;color:#cbd5e1;\">Bonjour <strong>${escHtml(st.nom)}</strong>, voici le bilan complet de vos ventes, d\u00e9p\u00f4ts et commissions pour la p\u00e9riode.</p>", "          </div>", "", "          <div style=\"padding:24px 28px;border:1px solid #e2e8f0;border-top:none;border-radius:0 0 14px 14px;\">", "            <!-- 4 CARTES R\u00c9CAPITULATIVES -->", "            <table style=\"width:100%;border-collapse:separate;border-spacing:8px;margin:0 -8px 14px;\">", "              <tr>", "                <td style=\"width:25%;padding:12px 14px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;vertical-align:top;\">", "                  <div style=\"font-size:11px;color:#64748b;font-weight:600;\">Volume vendu</div>", "                  <div style=\"font-size:16px;font-weight:800;color:#0f172a;margin-top:4px;\">${fmtMoney(st.totalVentes)}</div>", "                  <div style=\"font-size:11px;color:#059669;margin-top:4px;\">Confirm\u00e9 : ${fmtMoney(totVentesConf)}</div>", "                </td>", "                <td style=\"width:25%;padding:12px 14px;background:#ecfdf5;border:1px solid #a7f3d0;border-radius:10px;vertical-align:top;\">", "                  <div style=\"font-size:11px;color:#065f46;font-weight:600;\">Total commissions</div>", "                  <div style=\"font-size:16px;font-weight:800;color:#059669;margin-top:4px;\">${fmtMoney(st.totalComm)}</div>", "                  <div style=\"font-size:11px;color:#065f46;margin-top:4px;\">Confirm\u00e9 : ${fmtMoney(totCommConf)}</div>", "                </td>", "                <td style=\"width:25%;padding:12px 14px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;vertical-align:top;\">", "                  <div style=\"font-size:11px;color:#64748b;font-weight:600;\">D\u00e9j\u00e0 vers\u00e9</div>", "                  <div style=\"font-size:16px;font-weight:800;color:#334155;margin-top:4px;\">${fmtMoney(st.totalPaye)}</div>", "                  <div style=\"font-size:11px;color:#64748b;margin-top:4px;\">${nbConfirmes}/${nbDossiers} d\u00e9p\u00f4ts re\u00e7us</div>", "                </td>", "                <td style=\"width:25%;padding:12px 14px;background:#eef2ff;border:1px solid #c7d2fe;border-radius:10px;vertical-align:top;\">", "                  <div style=\"font-size:11px;color:#3730a3;font-weight:600;\">Solde \u00e0 verser</div>", "                  <div style=\"font-size:16px;font-weight:800;color:#4338ca;margin-top:4px;\">${fmtMoney(st.soldeAVerser)}</div>", "                  <div style=\"font-size:11px;color:#4338ca;margin-top:4px;\">Dont confirm\u00e9 : ${fmtMoney(soldeConf)}</div>", "                </td>", "              </tr>", "            </table>", "", "            <!-- INDICATEURS ACTIVIT\u00c9 -->", "            <div style=\"padding:10px 14px;background:#f1f5f9;border-radius:8px;font-size:12px;color:#334155;margin-bottom:16px;\">", "              <strong>Activit\u00e9 du mois :</strong> ${nbDossiers} dossier(s) sign\u00e9(s) \u2022 <strong>${nbOts} OTS</strong> / <strong>${nbSuivi} Suivi</strong> \u2022 <strong>${nbConfirmes} d\u00e9p\u00f4t(s) confirm\u00e9(s)</strong>", "            </div>", "", "            ${targetSectionHtml}", "", "            ${parEnt.length > 0 ? `", "              <h3 style=\"margin:20px 0 8px;font-size:13px;text-transform:uppercase;letter-spacing:0.6px;color:#334155;\">R\u00e9partition par entreprise partenaire</h3>", "              <table style=\"width:100%;border-collapse:collapse;font-size:12px;margin-bottom:20px;border:1px solid #e2e8f0;\">", "                <thead>", "                  <tr style=\"background:#f8fafc;color:#475569;text-align:left;border-bottom:1px solid #e2e8f0;\">", "                    <th style=\"padding:8px 10px;\">Entreprise</th>", "                    <th style=\"padding:8px 10px;text-align:center;\">Dossiers</th>", "                    <th style=\"padding:8px 10px;text-align:right;\">Volume ventes</th>", "                    <th style=\"padding:8px 10px;text-align:right;\">Commissions</th>", "                  </tr>", "                </thead>", "                <tbody>${entRowsHtml}</tbody>", "              </table>", "            ` : \"\"}", "", "            <h3 style=\"margin:20px 0 8px;font-size:13px;text-transform:uppercase;letter-spacing:0.6px;color:#334155;\">D\u00e9tail ligne par ligne de vos dossiers</h3>", "            <table style=\"width:100%;border-collapse:collapse;font-size:12px;border:1px solid #e2e8f0;\">", "              <thead>", "                <tr style=\"background:#0f172a;color:#f8fafc;text-align:left;\">", "                  <th style=\"padding:9px 8px;\">Date</th>", "                  <th style=\"padding:9px 8px;\">Entreprise</th>", "                  <th style=\"padding:9px 8px;\">Client</th>", "                  <th style=\"padding:9px 8px;text-align:right;\">Vente</th>", "                  <th style=\"padding:9px 8px;text-align:right;\">Taux</th>", "                  <th style=\"padding:9px 8px;text-align:right;\">Commission</th>", "                  <th style=\"padding:9px 8px;text-align:center;\">D\u00e9p\u00f4t</th>", "                  <th style=\"padding:9px 8px;text-align:center;\">Paiement</th>", "                </tr>", "              </thead>", "              <tbody>${rowsHtml}</tbody>", "              <tfoot>", "                <tr style=\"background:#f8fafc;font-weight:800;border-top:2px solid #cbd5e1;\">", "                  <td colspan=\"3\" style=\"padding:10px 8px;color:#0f172a;\">TOTAL (${nbDossiers} dossier(s))</td>", "                  <td style=\"padding:10px 8px;text-align:right;color:#0f172a;\">${fmtMoney(st.totalVentes)}</td>", "                  <td style=\"padding:10px 8px;\"></td>", "                  <td style=\"padding:10px 8px;text-align:right;color:#059669;\">${fmtMoney(st.totalComm)}</td>", "                  <td colspan=\"2\" style=\"padding:10px 8px;text-align:right;color:#4338ca;\">Solde : ${fmtMoney(st.soldeAVerser)}</td>", "                </tr>", "              </tfoot>", "            </table>", "", "            ${Array.isArray(st.trajets) && st.trajets.length > 0 ? `", "              <h3 style=\"margin:22px 0 8px;font-size:13px;text-transform:uppercase;letter-spacing:0.6px;color:#334155;\">\ud83e\udded Registre kilom\u00e9trage & adresses visit\u00e9es (${Number(st.totalKm || 0).toLocaleString(\"fr-CA\", { maximumFractionDigits: 1 })} km)</h3>", "              <table style=\"width:100%;border-collapse:collapse;font-size:12px;margin-bottom:18px;border:1px solid #e2e8f0;\">", "                <thead>", "                  <tr style=\"background:#f0f9ff;color:#0369a1;text-align:left;border-bottom:1px solid #bae6fd;\">", "                    <th style=\"padding:8px 10px;\">Date</th>", "                    <th style=\"padding:8px 10px;\">Itin\u00e9raire complet (D\u00e9part \u2192 Arr\u00eats)</th>", "                    <th style=\"padding:8px 10px;text-align:right;\">Distance</th>", "                  </tr>", "                </thead>", "                <tbody>", "                  ${st.trajets.map((tr: any) => `", "                    <tr>", "                      <td style=\"padding:8px 10px;border-bottom:1px solid #e2e8f0;white-space:nowrap;vertical-align:top;\">${escHtml(tr.date)}</td>", "                      <td style=\"padding:8px 10px;border-bottom:1px solid #e2e8f0;color:#334155;\"><strong>D\u00e9part :</strong> ${escHtml(tr.depart)} \u2192 ${(tr.arrets || []).map((a: any) => escHtml(a)).join(\" \u2192 \")}${tr.retour ? ` \u2192 <em>Retour (${escHtml(tr.depart)})</em>` : \"\"}</td>", "                      <td style=\"padding:8px 10px;border-bottom:1px solid #e2e8f0;text-align:right;font-weight:700;color:#0369a1;white-space:nowrap;vertical-align:top;\">${Number(tr.totalKm || 0).toLocaleString(\"fr-CA\", { maximumFractionDigits: 1 })} km</td>", "                    </tr>", "                  `).join(\"\")}", "                </tbody>", "              </table>", "            ` : \"\"}", "", "            ${dossiers.length > 0 ? `<div style=\"margin:18px 0 0;padding:12px 14px;background:#ecfdf5;border:1px solid #a7f3d0;border-radius:10px;font-size:12px;color:#065f46;\"><strong>\ud83d\udcce Pi\u00e8ce jointe Wave Accounting :</strong> Le fichier <code>${escHtml(waveFilename)}</code> est joint \u00e0 ce courriel. Importez-le directement dans <strong>Wave &gt; Ventes &gt; Factures &gt; Importer</strong> pour cr\u00e9er automatiquement votre facture avec une ligne par vente.</div>` : \"\"}", "            <p style=\"margin:16px 0 0;font-size:11px;color:#64748b;text-align:center;\">", "              Relev\u00e9 g\u00e9n\u00e9r\u00e9 automatiquement par le portail <strong>Korpo Solutions</strong> \u2022 <a href=\"https://ventes.korposolutions.com\" style=\"color:#4f46e5;text-decoration:none;\">ventes.korposolutions.com</a>", "            </p>", "          </div>", "        </div>", "      `;", "", "      const res = await fetch(\"https://api.resend.com/emails\", {", "        method: \"POST\",", "        headers: {", "          \"Authorization\": `Bearer ${RESEND_API_KEY}`,", "          \"Content-Type\": \"application/json\",", "        },", "        body: JSON.stringify({", "          from: FROM_EMAIL,", "          to: [st.courriel],", "          subject: `Relev\u00e9 complet ventes & commissions \u2014 ${st.mois} | Korpo Solutions`,", "          html: htmlEmail,", "          attachments,", "        }),", "      });", "      if (res.ok) sentCount++;", "    }", "", "    return new Response(JSON.stringify({ ok: true, sentCount }), {", "      headers: { \"Content-Type\": \"application/json\", \"Access-Control-Allow-Origin\": \"*\" },", "    });", "  } catch (err: any) {", "    return new Response(JSON.stringify({ ok: false, error: err.message }), {", "      status: 500,", "      headers: { \"Content-Type\": \"application/json\", \"Access-Control-Allow-Origin\": \"*\" },", "    });", "  }", "});"].join('\n');

    const DEFAULT_WAVE_COMPANY_SEQ = {
      'MaClim': 107,
      'MonSpray': 113,
      'Toitex': 115
    };

    function getWaveCompanyPrefix(companyName) {
      const clean = String(companyName || 'Korpo').trim().replace(/[^A-Za-z0-9]/g, '');
      if (!clean) return 'Korpo';
      return clean.charAt(0).toUpperCase() + clean.slice(1).toLowerCase();
    }

    function getWaveCompanySeqMap() {
      const stored = lsGetJson('korpo_wave_company_seq_v1', 'null');
      const base = { ...DEFAULT_WAVE_COMPANY_SEQ };
      if (stored && typeof stored === 'object') {
        Object.entries(stored).forEach(([k, v]) => {
          const n = parseInt(v, 10);
          if (!isNaN(n) && n >= 0) base[k] = n;
        });
      }
      (state.entreprises || []).forEach(e => {
        if (e && e.nom && base[e.nom] === undefined) {
          base[e.nom] = 100;
        }
      });
      return base;
    }

    async function saveWaveCompanySeqMap(seqMap) {
      if (!seqMap || typeof seqMap !== 'object') return;
      lsSet('korpo_wave_company_seq_v1', JSON.stringify(seqMap));
      if (sb && currentProfile) {
        try {
          await safeUpsertKpiHebdoRow({
            mois_commission: 'wave_company_seq_v1',
            semaine_index: -12,
            representant_nom: JSON.stringify(seqMap),
            booke: 0,
            rdv_realise: 0,
            updated_at: new Date().toISOString()
          });
        } catch (e) {}
      }
    }

    function renderAutoEmailCompanySeqInputs() {
      const grid = document.getElementById('auto-email-company-seq-grid');
      if (!grid) return;
      const seqMap = getWaveCompanySeqMap();
      const companies = Object.keys(seqMap);
      grid.innerHTML = companies.map(comp => {
        const curNum = Number(seqMap[comp] || 100);
        const prefix = getWaveCompanyPrefix(comp);
        const nextCode = `${prefix}${curNum + 1}`;
        return `
          <div class="inline-flex items-center gap-1.5 bg-white border border-indigo-200/90 rounded-xl px-2.5 py-1.5 shadow-2xs">
            <span class="text-[11px] font-extrabold text-slate-800">${esc(comp)} :</span>
            <input type="number" min="0" step="1" data-wave-seq-company="${esc(comp)}" value="${curNum}" class="w-16 px-1.5 py-0.5 text-xs font-extrabold text-indigo-950 bg-slate-50 border border-slate-200 rounded-lg tabular-nums text-center" />
            <span class="text-[10px] font-bold text-emerald-700 bg-emerald-50 border border-emerald-200 px-1.5 py-0.5 rounded-md tabular-nums">→ ${esc(nextCode)}</span>
          </div>
        `;
      }).join('');

      grid.querySelectorAll('input[data-wave-seq-company]').forEach(inp => {
        inp.addEventListener('input', () => {
          const comp = inp.dataset.waveSeqCompany;
          const val = parseInt(inp.value, 10);
          const mapNow = getWaveCompanySeqMap();
          if (!isNaN(val) && val >= 0) {
            mapNow[comp] = val;
            lsSet('korpo_wave_company_seq_v1', JSON.stringify(mapNow));
            const badge = inp.parentElement?.querySelector('span:last-child');
            if (badge) badge.textContent = `→ ${getWaveCompanyPrefix(comp)}${val + 1}`;
            renderAutoEmailPreviewList();
          }
        });
        inp.addEventListener('change', () => {
          saveWaveCompanySeqMap(getWaveCompanySeqMap());
        });
      });
    }

    function buildStatementForRep(rep, ym, confirmedOnly = false) {
      const isOlivier = rep.role === 'admin';
      const sales = (state.ventes || []).filter(v => {
        if (!isOlivier && String(v.representant_id) !== String(rep.id)) return false;
        if (ym && getSaleStatsMonth(v) !== ym) return false;
        if (confirmedOnly && v.statut_depot !== 'depot_recu') return false;
        return true;
      });
      const confirmedSales = sales.filter(v => v.statut_depot === 'depot_recu');
      const totalVentes = sales.reduce((s, v) => s + Number(v.valeur || 0), 0);
      const totalVentesConfirme = confirmedSales.reduce((s, v) => s + Number(v.valeur || 0), 0);
      const totalComm = sales.reduce((s, v) => s + Number(isOlivier ? v.commission_olivier : v.commission_rep), 0);
      const totalCommConfirme = confirmedSales.reduce((s, v) => s + Number(isOlivier ? v.commission_olivier : v.commission_rep), 0);
      const totalPaye = sales.filter(v => v.statut_paiement === 'paye').reduce((s, v) => s + Number(isOlivier ? v.commission_olivier : v.commission_rep), 0);
      const soldeAVerser = Math.max(0, totalComm - totalPaye);
      const soldeConfirmeAVerser = Math.max(0, totalCommConfirme - totalPaye);

      const salesTarget = getSalesTargetForRep(rep.nom, ym || currentMonthStr);
      const targetPct = salesTarget > 0 ? Math.round((totalVentes / salesTarget) * 100) : 0;

      const dossiers = sales.map(v => {
        const ent = state.entreprises.find(e => String(e.id) === String(v.entreprise_id))?.nom || '—';
        const cli = state.clients.find(c => String(c.id) === String(v.client_id))?.nom || '—';
        const comm = Number(isOlivier ? v.commission_olivier : v.commission_rep);
        const taux = Number(isOlivier ? v.taux_olivier_applique : v.taux_rep_applique);
        const rawNote = String(v.note || '');
        const typeKpi = rawNote.toUpperCase().includes('[SUIVI]') ? 'SUIVI' : 'OTS';
        const cleanNote = rawNote.replace(/^\[(OTS|SUIVI)\]\s*/i, '').replace(/\[(REP|ENT):[^\]]+\]\s*/ig, '').replace(/\[SANSDEPOT\]\s*/ig, '').trim();
        return {
          date: v.date_vente,
          entreprise: ent,
          client: cli,
          typeKpi,
          valeur: Number(v.valeur || 0),
          taux,
          commission: comm,
          depot: v.statut_depot === 'depot_recu' ? 'Confirmé' : 'À confirmer',
          paiement: v.statut_paiement === 'paye' ? 'Payé' : 'À payer',
          note: cleanNote
        };
      });

      const byEntMap = new Map();
      dossiers.forEach(d => {
        const cur = byEntMap.get(d.entreprise) || { entreprise: d.entreprise, count: 0, ventes: 0, comm: 0 };
        cur.count += 1;
        cur.ventes += d.valeur;
        cur.comm += d.commission;
        byEntMap.set(d.entreprise, cur);
      });
      const parEntreprise = Array.from(byEntMap.values()).sort((a, b) => b.ventes - a.ventes);

      const nbOts = dossiers.filter(d => d.typeKpi === 'OTS').length;
      const nbSuivi = dossiers.filter(d => d.typeKpi === 'SUIVI').length;

      // Génération du fichier CSV au format exact Wave Connect (1 facture par entreprise partenaire : ex. Maclim108, Monspray114, Toitex116)
      const invDate = localYmd();
      const workingCompanySeq = getWaveCompanySeqMap();
      const waveCsvHeaders = [
        'Customer Name',
        'Invoice Number',
        'Invoice Date',
        'Due Date',
        'Item Name',
        'Description',
        'Quantity',
        'Unit Price',
        'Memo'
      ];
      const sortedDossiersForWave = [...dossiers].sort((a, b) =>
        String(a.entreprise || '').localeCompare(String(b.entreprise || ''), 'fr') ||
        String(a.date || '').localeCompare(String(b.date || ''))
      );
      const assignedInvoiceCodeByCompany = new Map();
      const waveCsvRows = sortedDossiersForWave.map(d => {
        const customerCompany = d.entreprise || 'Korpo Solutions';
        if (!assignedInvoiceCodeByCompany.has(customerCompany)) {
          const prevNum = Number(workingCompanySeq[customerCompany] ?? 100);
          const nextNum = prevNum + 1;
          workingCompanySeq[customerCompany] = nextNum;
          assignedInvoiceCodeByCompany.set(customerCompany, `${getWaveCompanyPrefix(customerCompany)}${nextNum}`);
        }
        const invNumber = assignedInvoiceCodeByCompany.get(customerCompany);
        const clientDesc = `${d.client}${d.note ? ' (' + d.note + ')' : ''}`.replace(/[\t\r\n]+/g, ' ');
        const commAmt = Number(d.commission || 0).toFixed(2);
        return [
          csvCell(customerCompany),
          csvCell(invNumber),
          csvCell(invDate),
          csvCell(invDate),
          csvCell('Commission sur vente'),
          csvCell(clientDesc),
          '1',
          commAmt,
          csvCell(`Facture mois de ${ym || invDate.slice(0, 7)}`)
        ].join(',');
      });
      const waveCsvContent = [waveCsvHeaders.join(','), ...waveCsvRows].join('\n');
      const generatedInvoiceCodes = Array.from(assignedInvoiceCodeByCompany.values());

      const mileageInfo = getRepMonthlyMileage(rep, ym || currentMonthStr);
      const totalKmFmt = mileageInfo.totalKm.toLocaleString('fr-CA', { maximumFractionDigits: 1 });
      const mileageSectionLines = mileageInfo.entries.length > 0 ? [
        ``,
        `Registre kilométrage validé (${totalKmFmt} km au total sur ${mileageInfo.entries.length} trajet(s)) :`,
        ...mileageInfo.entries.map(tr => `- ${tr.date} (${Number(tr.totalKm || 0).toLocaleString('fr-CA', { maximumFractionDigits: 1 })} km) : Départ ${tr.depart} -> ${(tr.arrets || []).join(' -> ')}${tr.retour ? ' -> Retour départ' : ''}`)
      ] : [];

      const lines = [
        `Bonjour ${rep.nom},`,
        ``,
        `Voici votre relevé mensuel complet des ventes et commissions (${ym}) chez Korpo Solutions :`,
        ``,
        `• Dossiers signés : ${dossiers.length} (${nbOts} OTS / ${nbSuivi} Suivi) • Dépôts confirmés : ${confirmedSales.length}/${dossiers.length}`,
        `• Volume total vendu : ${fmtMoney(totalVentes)} (Confirmé : ${fmtMoney(totalVentesConfirme)})`,
        `• Objectif mensuel : ${fmtMoney0(salesTarget)} (${targetPct} % atteint)`,
        `• Total des commissions : ${fmtMoney(totalComm)} (Confirmées : ${fmtMoney(totalCommConfirme)})`,
        `• Déjà payé : ${fmtMoney(totalPaye)}`,
        `• Solde à verser : ${fmtMoney(soldeAVerser)} (Dont confirmé : ${fmtMoney(soldeConfirmeAVerser)})`,
        `• Kilométrage validé du mois : ${totalKmFmt} km (${mileageInfo.entries.length} trajet(s))`,
        ``,
        `💡 ASTUCE WAVE : Connectez-vous au portail Korpo > Mes ventes & commissions et cliquez sur « 📄 Facture Wave (CSV) » pour télécharger votre fichier prêt à importer dans Wave (Ventes > Factures > Importer) avec une ligne par vente.`,
        ``,
        `Répartition par entreprise :`,
        ...parEntreprise.map(e => `- ${e.entreprise} (${e.count} dossier(s)) : Ventes ${fmtMoney(e.ventes)} | Comm. ${fmtMoney(e.comm)}`),
        ``,
        `Détail de vos dossiers :`,
        ...dossiers.map(d => `- ${d.date} | ${d.entreprise} [${d.typeKpi}] | ${d.client} : Vente ${fmtMoney(d.valeur)} (${fmtPct(d.taux)}) -> Comm. ${fmtMoney(d.commission)} (${d.depot} / ${d.paiement})${d.note ? ' — ' + d.note : ''}`),
        ...mileageSectionLines,
        ``,
        `Cordialement,`,
        `L'équipe Korpo Solutions`
      ];

      return {
        repId: rep.id,
        nom: rep.nom,
        courriel: rep.courriel,
        mois: ym,
        confirmedOnly,
        nbDossiers: dossiers.length,
        nbConfirmes: confirmedSales.length,
        nbOts,
        nbSuivi,
        totalVentes,
        totalVentesConfirme,
        totalComm,
        totalCommConfirme,
        totalPaye,
        soldeAVerser,
        soldeConfirmeAVerser,
        totalKm: mileageInfo.totalKm,
        trajets: mileageInfo.entries,
        salesTarget,
        targetPct,
        parEntreprise,
        dossiers,
        updatedCompanySeqMap: workingCompanySeq,
        generatedInvoiceCodes,
        waveCsvContent,
        waveCsvFilename: `facture-wave-${String(rep.nom || 'rep').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]+/g, '-')}-${ym}.csv`,
        subject: `Relevé complet ventes & commissions — ${ym} | Korpo Solutions`,
        bodyText: lines.join('\n')
      };
    }

    async function sendStatementsForReps(repsList, ym, confirmedOnly = false) {
      if (!repsList || !repsList.length) return toast('Aucun représentant sélectionné.', 'error');
      const startingSeqMap = getWaveCompanySeqMap();
      const statements = [];

      for (const r of repsList) {
        const st = buildStatementForRep(r, ym, confirmedOnly);
        if (st.dossiers && st.dossiers.length > 0 && st.updatedCompanySeqMap) {
          lsSet('korpo_wave_company_seq_v1', JSON.stringify(st.updatedCompanySeqMap));
        }
        statements.push(st);
      }
      const finalSeqMap = getWaveCompanySeqMap();
      await saveWaveCompanySeqMap(finalSeqMap);
      renderAutoEmailCompanySeqInputs();

      const shouldAutoClose = Boolean(document.getElementById('auto-email-auto-close-month')?.checked);
      const finalizeAutoCloseIfNeeded = async () => {
        if (!shouldAutoClose || !isValidYearMonth(ym)) return false;
        if (typeof isMonthClosed === 'function' && isMonthClosed(ym)) return true;
        try {
          const { error: clErr } = await sb.from('mois_clotures').insert({
            ym,
            closed_by_nom: currentProfile?.nom || 'Admin',
            note: 'Clôturé automatiquement après envoi des relevés par courriel'
          });
          if (!clErr && typeof closedMonths !== 'undefined') {
            closedMonths.add(ym);
            renderMesVentes();
            renderAdminVentes();
            return true;
          }
        } catch (e) {}
        return false;
      };

      // 1. Tenter d'abord l'envoi automatique côté serveur via Supabase Edge Function (`send-monthly-statements`)
      try {
        const { data, error } = await sb.functions.invoke('send-monthly-statements', {
          body: { month: ym, confirmedOnly, autoCloseMonth: shouldAutoClose, companySeqMap: startingSeqMap, statements }
        });
        if (!error && data && data.ok) {
          const lockedNow = await finalizeAutoCloseIfNeeded();
          toast(lockedNow
            ? `✉️ ${data.sentCount || statements.length} relevé(s) envoyé(s) et mois ${ym} clôturé 🔒`
            : `✉️ ${data.sentCount || statements.length} relevé(s) envoyé(s) automatiquement par courriel ✓`);
          return true;
        }
      } catch (e) {}

      // 2. Fallback immédiat : télécharger automatiquement le fichier CSV Wave + ouvrir le brouillon courriel
      statements.forEach((st, idx) => {
        setTimeout(() => {
          if (st.dossiers && st.dossiers.length > 0 && st.waveCsvContent) {
            triggerWaveCsvDownload(st.waveCsvContent, st.waveCsvFilename || `facture-wave-${ym}.csv`);
          }
          const mailtoUrl = `mailto:${encodeURIComponent(st.courriel)}?subject=${encodeURIComponent(st.subject)}&body=${encodeURIComponent(st.bodyText)}`;
          window.open(mailtoUrl, idx === 0 ? '_self' : '_blank');
        }, idx * 450);
      });
      const lockedAfterDraft = await finalizeAutoCloseIfNeeded();
      toast(lockedAfterDraft
        ? `✉️ Relevé(s) préparé(s) et mois ${ym} clôturé 🔒 (chiffres figés à l'identique)`
        : `✉️ Brouillon(s) préparé(s) et fichier(s) CSV Wave téléchargé(s) pour ${statements.length} représentant(s) ✓`);
      return false;
    }

    function renderAutoEmailPreviewList() {
      const ym = document.getElementById('auto-email-month')?.value || currentMonthStr;
      const confirmedOnly = Boolean(document.getElementById('auto-email-confirmed-only')?.checked);
      const listBox = document.getElementById('auto-email-reps-list');
      const sumEl = document.getElementById('auto-email-summary-label');
      if (!listBox) return;

      const statements = state.representants.map(r => buildStatementForRep(r, ym, confirmedOnly));
      const activeCount = statements.filter(s => s.dossiers.length > 0).length;
      if (sumEl) sumEl.textContent = `${activeCount} représentant(s) avec ventes sur ${ym}`;

      listBox.innerHTML = statements.map(st => {
        const hasSales = st.dossiers.length > 0 || (st.trajets && st.trajets.length > 0);
        return `
          <label class="flex flex-col sm:flex-row sm:items-center justify-between gap-3 p-3.5 rounded-2xl border ${hasSales ? 'bg-white border-slate-200/90' : 'bg-slate-50/70 border-slate-200/60 opacity-65'} cursor-pointer hover:border-indigo-300 transition">
            <div class="flex items-center gap-3 min-w-0">
              <input type="checkbox" data-auto-email-rep-id="${esc(st.repId)}" ${hasSales ? 'checked' : ''} class="rounded border-slate-300 text-indigo-600 focus:ring-indigo-600" />
              <div class="min-w-0">
                <div class="text-xs font-extrabold text-slate-900 truncate">${esc(st.nom)} <span class="font-normal text-slate-500">(${esc(st.courriel)})</span></div>
                <div class="text-[11px] text-slate-500 tabular-nums mt-0.5">
                  ${st.dossiers.length} dossier(s) • Ventes : <strong class="text-slate-800">${fmtMoney(st.totalVentes)}</strong> • Comm. : <strong class="text-emerald-700">${fmtMoney(st.totalComm)}</strong> • Solde : <strong class="text-indigo-700">${fmtMoney(st.soldeAVerser)}</strong> • Route : <strong class="text-sky-700">${Number(st.totalKm || 0).toLocaleString('fr-CA', { maximumFractionDigits: 1 })} km</strong>
                </div>
              </div>
            </div>
            <div class="self-end sm:self-center flex items-center gap-1.5 shrink-0">
              <button type="button" onclick="event.preventDefault(); downloadWaveCsvForRep('${esc(st.repId)}')" title="Télécharger le fichier CSV prêt à importer dans Wave Accounting (1 ligne par vente)" class="px-2.5 py-1.5 text-[11px] font-bold text-emerald-800 bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 rounded-xl transition">
                📄 CSV Wave
              </button>
              <button type="button" onclick="event.preventDefault(); sendSingleRepStatementById('${esc(st.repId)}')" class="px-3 py-1.5 text-[11px] font-bold text-indigo-700 bg-indigo-50 hover:bg-indigo-100 rounded-xl transition">
                ✉️ Envoyer seul
              </button>
            </div>
          </label>
        `;
      }).join('');
    }

    function triggerWaveCsvDownload(csvContent, filename) {
      const blob = new Blob(['\ufeff' + csvContent], { type: 'text/csv;charset=utf-8;' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = filename;
      a.click();
    }

    window.downloadWaveCsvForRep = (repId) => {
      const rep = state.representants.find(r => String(r.id) === String(repId));
      if (!rep) return;
      const ym = document.getElementById('auto-email-month')?.value || currentMonthStr;
      const confirmedOnly = Boolean(document.getElementById('auto-email-confirmed-only')?.checked);
      const st = buildStatementForRep(rep, ym, confirmedOnly);
      if (!st.dossiers.length) return toast('Aucune commission à exporter pour ce mois.', 'error');
      if (st.updatedCompanySeqMap) {
        saveWaveCompanySeqMap(st.updatedCompanySeqMap);
        renderAutoEmailCompanySeqInputs();
      }
      const codesSummary = (st.generatedInvoiceCodes || []).join(', ');
      const slug = rep.nom.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]+/g, '-');
      triggerWaveCsvDownload(st.waveCsvContent, `facture-wave-${slug}-${ym}.csv`);
      toast(`📄 Fichier Wave téléchargé (${codesSummary || st.dossiers.length + ' ligne(s)'}) pour ${rep.nom} ✓`);
    };

    window.sendSingleRepStatementById = async (repId) => {
      const rep = state.representants.find(r => String(r.id) === String(repId));
      if (!rep) return;
      const ym = document.getElementById('auto-email-month')?.value || currentMonthStr;
      const confirmedOnly = Boolean(document.getElementById('auto-email-confirmed-only')?.checked);
      await sendStatementsForReps([rep], ym, confirmedOnly);
    };

    document.getElementById('btn-open-auto-email-modal')?.addEventListener('click', () => {
      const dashPeriod = readPeriodFilterState('dash-filter');
      const mIn = document.getElementById('auto-email-month');
      if (mIn) mIn.value = dashPeriod.singleMonth || currentMonthStr;
      renderAutoEmailCompanySeqInputs();
      const confIn = document.getElementById('auto-email-confirmed-only');
      if (confIn) confIn.checked = commDisplayMode === 'confirme';
      renderAutoEmailPreviewList();
      document.getElementById('auto-email-modal').classList.remove('hidden');
    });

    document.getElementById('auto-email-month')?.addEventListener('change', renderAutoEmailPreviewList);
    document.getElementById('auto-email-confirmed-only')?.addEventListener('change', renderAutoEmailPreviewList);
    document.getElementById('btn-close-auto-email')?.addEventListener('click', () => document.getElementById('auto-email-modal').classList.add('hidden'));
    document.getElementById('btn-cancel-auto-email')?.addEventListener('click', () => document.getElementById('auto-email-modal').classList.add('hidden'));
    document.getElementById('auto-email-modal')?.addEventListener('click', (e) => {
      if (e.target.id === 'auto-email-modal') document.getElementById('auto-email-modal').classList.add('hidden');
    });
    document.getElementById('btn-toggle-edge-guide')?.addEventListener('click', () => {
      document.getElementById('auto-email-edge-guide')?.classList.toggle('hidden');
    });
    document.getElementById('btn-copy-edge-fn-code')?.addEventListener('click', async () => {
      try {
        await navigator.clipboard.writeText(SUPABASE_EDGE_FUNCTION_CODE);
        toast('Code de la fonction Supabase copié dans le presse-papiers 📋');
      } catch (e) {
        toast('Impossible de copier automatiquement.', 'error');
      }
    });
    document.getElementById('btn-send-all-statements-now')?.addEventListener('click', async () => {
      const ym = document.getElementById('auto-email-month')?.value || currentMonthStr;
      const confirmedOnly = Boolean(document.getElementById('auto-email-confirmed-only')?.checked);
      const checkedIds = Array.from(document.querySelectorAll('#auto-email-reps-list input[data-auto-email-rep-id]:checked')).map(el => String(el.dataset.autoEmailRepId));
      const selectedReps = state.representants.filter(r => checkedIds.includes(String(r.id)));
      if (!selectedReps.length) return toast('Cochez au moins un représentant.', 'error');
      const btn = document.getElementById('btn-send-all-statements-now');
      btn.disabled = true;
      btn.textContent = 'Envoi en cours...';
      try {
        await sendStatementsForReps(selectedReps, ym, confirmedOnly);
        document.getElementById('auto-email-modal').classList.add('hidden');
      } finally {
        btn.disabled = false;
        btn.innerHTML = '<span>🚀</span> Envoyer à tous les représentants sélectionnés';
      }
    });

    function renderAdminTaux() {
      const tbody = document.getElementById('tbody-taux');
      if (!state.taux.length) {
        tbody.innerHTML = `<tr><td colspan="6" class="py-10 text-center text-sm text-slate-400">Aucun taux configuré.</td></tr>`;
        return;
      }
      tbody.innerHTML = state.taux.map(t => {
        const rep = state.representants.find(r => r.id === t.representant_id)?.nom || '—';
        const ent = state.entreprises.find(e => e.id === t.entreprise_id)?.nom || '—';
        const cli = t.client_id ? (state.clients.find(c => c.id === t.client_id)?.nom || 'Client') : 'Toute l\'entreprise';
        return `
          <tr class="hover:bg-slate-50">
            <td class="py-3.5 px-4 font-semibold text-slate-900">${esc(rep)}</td>
            <td class="py-3.5 px-4">${entBadge(ent)}</td>
            <td class="py-3.5 px-4"><span class="px-2.5 py-0.5 text-xs font-semibold rounded-full ${t.client_id ? 'bg-indigo-50 text-indigo-700' : 'bg-slate-100 text-slate-600'}">${esc(cli)}</span></td>
            <td class="py-3.5 px-4 text-right tabular-nums font-bold">${fmtPct(t.taux)}</td>
            <td class="py-3.5 px-4 text-right tabular-nums font-bold text-emerald-600">${fmtPct(t.taux_olivier)}</td>
            <td class="py-3.5 px-4 text-right"><button onclick="deleteTaux('${t.id}')" class="text-xs font-semibold text-rose-600 hover:underline">Supprimer</button></td>
          </tr>
        `;
      }).join('');
    }

    function renderAdminEquipe() {
      syncRedPhoneUI();
      renderEquipeTargetsTable();
      document.getElementById('list-entreprises').innerHTML = state.entreprises.map(e => `
        <li class="px-4 py-3 flex items-center justify-between text-sm bg-white">
          ${entBadge(e.nom)}
          <div class="flex items-center gap-3">
            <button type="button" onclick="toggleEntrepriseDepot('${e.id}')" title="Cliquer pour changer (n'affecte pas les ventes déjà enregistrées)" class="text-[11px] font-bold px-2.5 py-0.5 rounded-full border transition cursor-pointer ${e.depot_requis === false ? 'bg-slate-100 text-slate-700 border-slate-300 hover:bg-slate-200' : 'bg-emerald-50 text-emerald-700 border-emerald-200 hover:bg-emerald-100'}">${e.depot_requis === false ? '🚫 Sans dépôt' : '💵 Dépôt exigé'}</button>
            <button onclick="deleteEntreprise('${e.id}')" class="text-xs font-semibold text-rose-600 hover:underline">Supprimer</button>
          </div>
        </li>
      `).join('');

      document.getElementById('list-representants').innerHTML = state.representants.map(r => {
        const ratesCount = (state.taux || []).filter(t => String(t.representant_id) === String(r.id)).length;
        const rateStatusBadge = r.role === 'admin'
          ? ''
          : (ratesCount > 0
            ? `<span class="ml-1.5 text-[10px] font-bold px-2 py-0.5 rounded-full bg-emerald-50 text-emerald-700 border border-emerald-200">${ratesCount} taux actif(s)</span>`
            : `<button type="button" onclick="openRatesForRep('${r.id}')" class="ml-1.5 text-[10px] font-extrabold px-2 py-0.5 rounded-full bg-amber-100 text-amber-900 border border-amber-300 hover:bg-amber-200 transition cursor-pointer">⚠️ Configurer taux →</button>`);
        return `
        <li class="px-4 py-3 flex items-center justify-between gap-2 text-sm bg-white">
          <div class="min-w-0">
            <div class="font-bold text-slate-900 flex flex-wrap items-center">
              <span>${esc(r.nom)}</span>
              <span class="ml-1.5 text-[11px] font-semibold px-2 py-0.5 rounded-full bg-slate-100 text-slate-600">${esc(r.role)}</span>
              ${rateStatusBadge}
            </div>
            <div class="text-xs text-slate-500 truncate">${esc(r.courriel)}</div>
          </div>
          <div class="shrink-0">
            ${r.id !== currentProfile.id ? `<button onclick="deleteRepresentant('${r.id}')" class="text-xs font-semibold text-rose-600 hover:underline">Supprimer</button>` : '<span class="text-xs font-medium text-emerald-600">Session active</span>'}
          </div>
        </li>
      `;
      }).join('');
    }

    const pendingClientCreations = new Map();
    async function getOrCreateClient(nom, entrepriseId) {
      const clean = nom.trim();
      const lower = clean.toLowerCase();
      const existing = state.clients.find(c => String(c.entreprise_id) === String(entrepriseId) && c.nom.toLowerCase() === lower);
      if (existing) return existing.id;

      const key = `${entrepriseId}::${lower}`;
      if (pendingClientCreations.has(key)) return pendingClientCreations.get(key);
      const p = (async () => {
        const { data, error } = await sb.from('clients').insert({ nom: clean, entreprise_id: entrepriseId }).select().single();
        if (error) throw error;
        state.clients.push(data);
        return data.id;
      })();
      pendingClientCreations.set(key, p);
      try { return await p; } finally { pendingClientCreations.delete(key); }
    }

    function normClientKey(name) {
      return String(name || '')
        .trim()
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/\bsainte\b/g, 'st')
        .replace(/\bsaint\b/g, 'st')
        .replace(/[^a-z0-9]+/g, ' ')
        .trim();
    }

    function getResolvedClientNameForSale(v) {
      if (!v) return '';
      return v.client_nom || state.clients.find(c => String(c.id) === String(v.client_id))?.nom || '';
    }

    function areSameClientSales(v1, v2) {
      if (!v1 || !v2) return false;
      const sameRep = String(v1.representant_id || '') === String(v2.representant_id || '');
      const sameEnt = String(v1.entreprise_id || '') === String(v2.entreprise_id || '');
      if (!sameRep || !sameEnt) return false;
      const n1 = normClientKey(getResolvedClientNameForSale(v1));
      const n2 = normClientKey(getResolvedClientNameForSale(v2));
      if (!n1 || !n2) return false;
      if (n1 === n2) return true;
      // Correspondance tolérante si un nom complet inclut l'autre (ex. "Benoit St-Marie" vs "Benoit Sainte-Marie" ou note/prénom)
      const tokens1 = n1.split(' ').filter(t => t.length >= 3);
      const tokens2 = n2.split(' ').filter(t => t.length >= 3);
      if (tokens1.length >= 2 && tokens2.length >= 2 && tokens1.every(t => tokens2.includes(t))) return true;
      return false;
    }

    function findAllHistoricalDuplicates() {
      const sales = Array.isArray(state.ventes) ? state.ventes : [];
      const candidates = []; // [{ dup, kept, reason, defaultChecked }]
      const seenDupIds = new Set();

      const confirmedRows = sales.filter(v => v.statut_depot === 'depot_recu');
      const unconfirmedRows = sales.filter(v => v.statut_depot === 'a_confirmer');

      // 1. Toute entrée `a_confirmer` qui a un dossier `depot_recu` pour le même client + entreprise + rep
      confirmedRows.forEach(confV => {
        unconfirmedRows.forEach(cand => {
          const cid = String(cand.id);
          if (cid === String(confV.id) || seenDupIds.has(cid)) return;
          if (!areSameClientSales(confV, cand)) return;
          seenDupIds.add(cid);
          candidates.push({
            dup: cand,
            kept: confV,
            reason: 'Déjà confirmé (Dépôt reçu) sur un autre dossier',
            defaultChecked: true
          });
        });
      });

      // 2. Plusieurs entrées `a_confirmer` pour le même client + entreprise + rep + montant (ex. reporté/ressaisi sur plusieurs mois)
      unconfirmedRows.forEach(newerV => {
        const nKey = `${newerV.date_vente || newerV.mois_commission || ''}_${newerV.created_at || ''}_${newerV.id || ''}`;
        unconfirmedRows.forEach(olderV => {
          const oid = String(olderV.id);
          if (oid === String(newerV.id) || seenDupIds.has(oid)) return;
          if (!areSameClientSales(newerV, olderV)) return;
          if (Math.abs(Number(newerV.valeur || 0) - Number(olderV.valeur || 0)) >= 0.01) return;
          const oKey = `${olderV.date_vente || olderV.mois_commission || ''}_${olderV.created_at || ''}_${olderV.id || ''}`;
          if (nKey > oKey) {
            seenDupIds.add(oid);
            candidates.push({
              dup: olderV,
              kept: newerV,
              reason: 'Ancienne copie non confirmée (remplacée par une entrée plus récente)',
              defaultChecked: true
            });
          }
        });
      });

      // 3. Doublons exacts déjà marqués `depot_recu` (même client + entreprise + rep + exactement le même montant, sur des mois différents ou la même date)
      for (let i = 0; i < confirmedRows.length; i++) {
        for (let j = i + 1; j < confirmedRows.length; j++) {
          const a = confirmedRows[i];
          const b = confirmedRows[j];
          if (!areSameClientSales(a, b)) continue;
          if (Math.abs(Number(a.valeur || 0) - Number(b.valeur || 0)) >= 0.01) continue;
          const aKey = `${a.date_vente || a.mois_commission || ''}_${a.created_at || ''}_${a.id || ''}`;
          const bKey = `${b.date_vente || b.mois_commission || ''}_${b.created_at || ''}_${b.id || ''}`;
          const [kept, dup] = aKey >= bKey ? [a, b] : [b, a];
          const did = String(dup.id);
          if (seenDupIds.has(did)) continue;
          seenDupIds.add(did);
          candidates.push({
            dup,
            kept,
            reason: 'Doublon confirmé avec exactement le même montant (vérifiez s’il s’agit d’un 2e contrat distinct)',
            defaultChecked: a.date_vente === b.date_vente
          });
        }
      }

      return candidates;
    }

    function updateDedupBadge() {
      const badge = document.getElementById('dedup-badge-count');
      if (!badge) return;
      const list = findAllHistoricalDuplicates();
      badge.textContent = String(list.length);
      badge.className = list.length > 0
        ? 'px-1.5 py-0.2 text-[10px] font-extrabold rounded-full bg-rose-600 text-white'
        : 'px-1.5 py-0.2 text-[10px] font-extrabold rounded-full bg-amber-200/80 text-amber-950';
    }

    async function executeDuplicateCleanup(pairsToClean) {
      if (!sb || !Array.isArray(pairsToClean) || !pairsToClean.length) return 0;
      const dupsOnly = pairsToClean.map(x => x.dup);
      await moveSalesToTrash(dupsOnly);
      for (const { dup, kept } of pairsToClean) {
        const dupYm = dup.date_vente ? dup.date_vente.slice(0, 7) : dup.mois_commission;
        const keptYm = kept.mois_commission || (kept.date_vente ? kept.date_vente.slice(0, 7) : '');
        await recordSaleAudit(
          String(kept.id),
          'Doublon précédent fusionné automatiquement',
          `L'entrée de ${dupYm} (${fmtMoney(dup.valeur)}) a été retirée pour ne compter ce dossier qu'une seule fois en ${keptYm}`,
          '🔗'
        );
        await sb.from('ventes').delete().eq('id', dup.id);
        state.ventes = state.ventes.filter(x => String(x.id) !== String(dup.id));
        state.mesVentes = state.mesVentes.filter(x => String(x.id) !== String(dup.id));
      }
      updateDedupBadge();
      return pairsToClean.length;
    }

    async function reconcilePreviousUnconfirmedDuplicateSales(confirmedSaleIds = null) {
      if (!sb || !Array.isArray(state.ventes) || !state.ventes.length) return 0;
      const isSpecific = Array.isArray(confirmedSaleIds) && confirmedSaleIds.length > 0;
      const idSet = isSpecific ? new Set(confirmedSaleIds.map(x => String(x))) : null;

      const allDups = findAllHistoricalDuplicates().filter(item => {
        if (item.dup.statut_depot !== 'a_confirmer') return false;
        if (idSet) return idSet.has(String(item.kept.id));
        return true;
      });
      if (!allDups.length) return 0;
      return await executeDuplicateCleanup(allDups);
    }

    window.openDedupAuditModal = () => {
      if (currentProfile?.role !== 'admin') return;
      const list = findAllHistoricalDuplicates();
      window._currentDedupCandidates = list;

      const sumEl = document.getElementById('dedup-count-summary');
      const btnClean = document.getElementById('btn-dedup-clean-selected');
      const body = document.getElementById('dedup-modal-body');

      sumEl.textContent = list.length > 0
        ? `${list.length} doublon(s) potentiel(s) détecté(s) dans l'historique`
        : 'Aucun doublon détecté : votre base est 100 % propre !';
      if (btnClean) btnClean.classList.toggle('hidden', list.length === 0);

      if (!list.length) {
        body.innerHTML = `<div class="py-12 text-center text-xs text-slate-500">✅ Aucun client en double détecté dans vos ventes historiques.</div>`;
      } else {
        body.innerHTML = list.map((item, idx) => {
          const { dup, kept, reason, defaultChecked } = item;
          const repN = state.representants.find(r => String(r.id) === String(dup.representant_id))?.nom || '—';
          const entN = state.entreprises.find(e => String(e.id) === String(dup.entreprise_id))?.nom || '—';
          const cliN = getResolvedClientNameForSale(dup) || 'Client';
          return `
            <label class="block bg-white rounded-2xl border border-slate-200/90 p-4 shadow-2xs cursor-pointer hover:border-indigo-300 transition">
              <div class="flex items-start gap-3">
                <input type="checkbox" data-dedup-idx="${idx}" ${defaultChecked ? 'checked' : ''} class="mt-1 rounded border-slate-300 text-indigo-600 focus:ring-indigo-600" />
                <div class="flex-1 min-w-0 space-y-2">
                  <div class="flex flex-wrap items-center justify-between gap-2">
                    <div class="flex items-center flex-wrap gap-2">
                      ${entBadge(entN)}
                      <span class="font-extrabold text-sm text-slate-900">${esc(cliN)}</span>
                      <span class="text-xs font-bold text-indigo-600">(${esc(repN)})</span>
                    </div>
                    <span class="text-[11px] font-bold px-2.5 py-0.5 rounded-full bg-amber-50 text-amber-800 border border-amber-200">${esc(reason)}</span>
                  </div>

                  <div class="grid grid-cols-1 sm:grid-cols-2 gap-2.5 text-xs">
                    <div class="rounded-xl bg-rose-50/70 border border-rose-200 p-2.5">
                      <div class="text-[10px] font-extrabold uppercase text-rose-700">🗑️ Entrée à retirer (mise en corbeille)</div>
                      <div class="font-bold text-slate-900 mt-0.5 tabular-nums">${dup.date_vente} (Mois : ${dup.mois_commission}) — ${fmtMoney(dup.valeur)}</div>
                      <div class="text-[11px] text-slate-600">Statut : ${dup.statut_depot === 'depot_recu' ? 'Dépôt reçu' : 'À confirmer'} • ${dup.statut_paiement === 'paye' ? 'Payé' : 'Non payé'}</div>
                    </div>
                    <div class="rounded-xl bg-emerald-50/70 border border-emerald-200 p-2.5">
                      <div class="text-[10px] font-extrabold uppercase text-emerald-700">✅ Dossier conservé dans les stats</div>
                      <div class="font-bold text-slate-900 mt-0.5 tabular-nums">${kept.date_vente} (Mois : ${kept.mois_commission}) — ${fmtMoney(kept.valeur)}</div>
                      <div class="text-[11px] text-slate-600">Statut : ${kept.statut_depot === 'depot_recu' ? 'Dépôt reçu' : 'À confirmer'} • ${kept.statut_paiement === 'paye' ? 'Payé' : 'Non payé'}</div>
                    </div>
                  </div>
                </div>
              </div>
            </label>
          `;
        }).join('');
      }
      document.getElementById('dedup-modal').classList.remove('hidden');
    };

    document.getElementById('btn-close-dedup')?.addEventListener('click', () => {
      document.getElementById('dedup-modal').classList.add('hidden');
    });
    document.getElementById('dedup-modal')?.addEventListener('click', (e) => {
      if (e.target.id === 'dedup-modal') document.getElementById('dedup-modal').classList.add('hidden');
    });
    document.getElementById('btn-dedup-clean-selected')?.addEventListener('click', async () => {
      const list = window._currentDedupCandidates || [];
      const checkedBoxes = Array.from(document.querySelectorAll('#dedup-modal-body input[data-dedup-idx]:checked'));
      if (!checkedBoxes.length) return toast('Sélectionnez au moins un doublon à nettoyer.', 'error');
      const selectedPairs = checkedBoxes.map(cb => list[Number(cb.dataset.dedupIdx)]).filter(Boolean);
      const count = await executeDuplicateCleanup(selectedPairs);
      document.getElementById('dedup-modal').classList.add('hidden');
      toast(`🧹 ${count} doublon(s) retiré(s) et déplacé(s) dans la corbeille ! Statistiques mises à jour ✓`);
      await loadAllData();
    });

    window.toggleVenteStatut = async (id, field, newVal) => {
      const { error } = await sb.from('ventes').update({ [field]: newVal }).eq('id', id);
      if (error) return toast(error.message, 'error');
      const label = field === 'statut_depot'
        ? (newVal === 'depot_recu' ? 'Dépôt confirmé (Reçu)' : 'Dépôt remis À confirmer')
        : (newVal === 'paye' ? 'Commission marquée Payée' : 'Commission remise Non payée');
      await recordSaleAudit(id, label, '', field === 'statut_depot' ? '✅' : '💳');
      let mergedCount = 0;
      if (field === 'statut_depot' && newVal === 'depot_recu') {
        mergedCount = await reconcilePreviousUnconfirmedDuplicateSales([id]);
      }
      toast(mergedCount > 0
        ? `Dépôt confirmé ✓ (${mergedCount} doublon antérieur non confirmé retiré automatiquement)`
        : 'Statut mis à jour');
      await loadAllData();
    };

    window.toggleVenteKpiType = async (id) => {
      const v = state.ventes.find(x => String(x.id) === String(id));
      if (!v) return;
      const rawNote = v.note || '';
      const isCurrentlySuivi = rawNote.toUpperCase().includes('[SUIVI]');
      const nextTag = isCurrentlySuivi ? 'OTS' : 'SUIVI';
      const cleanBody = rawNote.replace(/^\[(OTS|SUIVI)\]\s*/i, '').replace(/\[(REP|ENT):[^\]]+\]\s*/ig, '').replace(/\[SANSDEPOT\]\s*/ig, '').trim();
      const updatedNote = cleanBody ? `[${nextTag}] ${cleanBody}` : `[${nextTag}]`;

      const { error } = await sb.from('ventes').update({ note: updatedNote }).eq('id', id);
      if (error) return toast(error.message, 'error');
      await recordSaleAudit(id, `Type KPI basculé en ${nextTag}`, '', '🏷️');
      toast(`Dossier classé en ${nextTag} ✓`);
      await loadAllData();
    };

    window.deleteVente = async (id) => {
      const ok = await askConfirm('Mettre cette vente à la corbeille ?', 'La vente sera retirée des totaux et conservée dans la corbeille pendant 30 jours pour restauration.');
      if (!ok) return;
      const targetSale = state.ventes.find(v => String(v.id) === String(id)) || state.mesVentes.find(v => String(v.id) === String(id));
      if (targetSale) {
        const trashOk = await moveSalesToTrash([targetSale]);
        if (!trashOk) toast('Attention : la corbeille n’a pu être sauvegardée que sur cet appareil.', 'error');

      }
      const { error } = await sb.from('ventes').delete().eq('id', id);
      if (error) return toast(error.message, 'error');
      toast('Vente déplacée dans la corbeille 🗑️');
      await loadAllData();
    };

    window.deleteTaux = async (id) => {
      const { error } = await sb.from('taux_commission').delete().eq('id', id);
      if (error) return toast(error.message, 'error');
      toast('Taux supprimé');
      await loadAllData();
    };

    window.toggleEntrepriseDepot = async (id) => {
      if (currentProfile?.role !== 'admin') return;
      const ent = (state.entreprises || []).find(e => String(e.id) === String(id));
      if (!ent) return;
      const nextRequired = ent.depot_requis === false; // bascule
      const { error } = await sb.from('entreprises').update({ depot_requis: nextRequired }).eq('id', id);
      if (error) {
        return toast(/depot_requis/i.test(error.message || '')
          ? 'Colonne « depot_requis » absente de Supabase : exécutez le SQL fourni.'
          : error.message, 'error');
      }
      toast(`${ent.nom} : ${nextRequired ? 'dépôt exigé' : 'aucun dépôt requis'} (ventes futures)`);
      await loadAllData();
    };

    window.deleteEntreprise = async (id) => {
      const entObj = state.entreprises.find(x => String(x.id) === String(id));
      const nom = entObj?.nom || 'cette entreprise';
      const matchingSales = state.ventes.filter(v => String(v.entreprise_id) === String(id));
      const hasPastSales = matchingSales.length > 0;
      const ok = await askConfirm(
        `Supprimer ${nom} des entreprises actives ?`,
        hasPastSales
          ? `Ses statistiques et ventes antérieures seront conservées dans l'historique global et ses objectifs du mois en cours seront ajustés au prorata.`
          : `Cette entreprise, ses taux et tous ses objectifs seront retirés des KPI.`
      );
      if (!ok) return;

      const curMonthEntObj = Number(kpiStore[currentMonthStr]?.entreprises?.[nom]?.total?.objectif || 0);
      const archiveKey = String(id || nom);
      archivedEntreprisesStore[archiveKey] = {
        id: String(id || ''),
        nom,
        archivedAt: new Date().toISOString(),
        saleIds: matchingSales.map(v => String(v.id)),
        monthlyObjective: curMonthEntObj
      };
      saveArchivedEntreprisesStore();

      const metaKeysToRemove = Object.keys(companyTargetsMetaStore || {}).filter(k => {
        const m = companyTargetsMetaStore[k];
        const entN = m?.entNom || k.split('__')[0];
        return String(entN || '').trim().toLowerCase() === nom.toLowerCase();
      });
      for (const k of metaKeysToRemove) {
        await removeCompanyTargetMeta(k);
      }

      if (Array.isArray(clientTargetsStore)) {
        const beforeLen = clientTargetsStore.length;
        clientTargetsStore = clientTargetsStore.filter(t => String(t.entNom || '').trim().toLowerCase() !== nom.toLowerCase());
        if (clientTargetsStore.length !== beforeLen) saveClientTargetsStore();
      }

      if (sb) {
        try {
          await sb.from('taux_commission').delete().eq('entreprise_id', id);
        } catch (e) {}

        if (hasPastSales) {
          for (const v of matchingSales) {
            const curNote = String(v.note || '').trim();
            const taggedNote = /\[ENT:/i.test(curNote) ? curNote : `${curNote} [ENT:${nom}]`.trim();
            try {
              await sb.from('ventes').update({ entreprise_id: null, note: taggedNote }).eq('id', v.id);
            } catch (e) {}
          }
          try {
            await sb.from('clients').update({ entreprise_id: null }).eq('entreprise_id', id);
          } catch (e) {}
        }

        const { error } = await sb.from('entreprises').delete().eq('id', id);
        if (error) return toast(error.message, 'error');
      }

      toast(`${nom} retirée (ventes antérieures conservées et objectifs ajustés au prorata) ✓`);
      await loadAllData();
    };

    window.deleteRepresentant = async (id) => {
      const repObj = state.representants.find(x => String(x.id) === String(id));
      const nom = repObj?.nom || 'ce membre';
      const matchingSales = state.ventes.filter(v => String(v.representant_id) === String(id));
      const hasPastSales = matchingSales.length > 0;
      const ok = await askConfirm(
        `Supprimer ${nom} de l'équipe active ?`,
        hasPastSales
          ? `Ses statistiques et ventes antérieures seront conservées dans l'historique global et ses objectifs du mois en cours seront calculés au prorata.`
          : `Ce membre et tous ses objectifs seront retirés de l'équipe et des KPI.`
      );
      if (!ok) return;

      const curRepKpi = kpiStore[currentMonthStr]?.reps?.[nom] || {};
      const snapTargets = {
        dollars: getSalesTargetForRep(nom, currentMonthStr),
        comm: getCommissionTargetForRep(nom, currentMonthStr),
        booke: Number(curRepKpi['Booké']?.objectif || 80),
        rdv: Number(curRepKpi['RDV Réalisé']?.objectif || 60),
        ots: Number(curRepKpi['Ventes OTS']?.objectif || 10),
        suivi: Number(curRepKpi['Ventes SUIVI']?.objectif || 20)
      };

      const archiveKey = String(id || nom);
      archivedRepsStore[archiveKey] = {
        id: String(id || ''),
        nom,
        archivedAt: new Date().toISOString(),
        saleIds: matchingSales.map(v => String(v.id)),
        targets: snapTargets
      };
      saveArchivedRepsStore();

      Object.keys(repTargetsOverrides || {}).forEach(k => {
        if (k.endsWith(`_${nom}`)) delete repTargetsOverrides[k];
      });
      lsSet('korpo_rep_targets_v1', JSON.stringify(repTargetsOverrides));

      Object.keys(personalMonthlyTargets || {}).forEach(k => {
        if (k.startsWith(`${nom}_`)) delete personalMonthlyTargets[k];
      });
      lsSet('korpo_personal_monthly_targets_v1', JSON.stringify(personalMonthlyTargets));

      Object.keys(personalCommTargets || {}).forEach(k => {
        if (k.startsWith(`${nom}_`)) delete personalCommTargets[k];
      });
      lsSet('korpo_personal_comm_targets_v1', JSON.stringify(personalCommTargets));

      Object.keys(personalBonusPctTargets || {}).forEach(k => {
        if (k === nom || k.startsWith(`${nom}_`)) delete personalBonusPctTargets[k];
      });
      lsSet('korpo_personal_bonus_pct_v1', JSON.stringify(personalBonusPctTargets));

      if (sb) {
        try {
          await sb.from('kpi_hebdo').delete().eq('representant_nom', nom).eq('semaine_index', -1);
          await sb.from('kpi_hebdo').delete().eq('mois_commission', nom).eq('semaine_index', -7);
          await sb.from('taux_commission').delete().eq('representant_id', id);
        } catch (e) {}

        if (hasPastSales) {
          for (const v of matchingSales) {
            const curNote = String(v.note || '').trim();
            const taggedNote = /\[REP:/i.test(curNote) ? curNote : `${curNote} [REP:${nom}]`.trim();
            try {
              await sb.from('ventes').update({ representant_id: null, note: taggedNote }).eq('id', v.id);
            } catch (e) {}
          }
        }

        const { error } = await sb.from('representants').delete().eq('id', id);
        if (error) return toast(error.message, 'error');
      }

      syncAllKpiRepsFromState();
      toast(`${nom} retiré de l'équipe (ventes antérieures conservées et objectifs ajustés au prorata) ✓`);
      await loadAllData();
    };

    document.getElementById('btn-recalc-all-rates')?.addEventListener('click', async () => {
      if (!state.ventes.length) return toast('Aucune vente à recalculer.', 'error');
      const ok = await askConfirm(
        `Recalculer les ${state.ventes.length} ventes ?`,
        'Tous les taux et montants de commission seront mis à jour selon votre grille des taux actuelle.'
      );
      if (!ok) return;

      const chronoVentes = [...state.ventes].sort((a, b) => {
        const cmpD = String(a.date_vente || '').localeCompare(String(b.date_vente || ''));
        if (cmpD !== 0) return cmpD;
        return String(a.created_at || a.id || '').localeCompare(String(b.created_at || b.id || ''));
      });
      const runningRepMonthVol = new Map();
      let updated = 0;
      for (const v of chronoVentes) {
        const ym = v.mois_commission || (v.date_vente ? v.date_vente.slice(0, 7) : currentMonthStr);
        const key = `${v.representant_id}_${ym}`;
        const priorVol = runningRepMonthVol.get(key) || 0;
        const eff = getEffectiveRatesWithTargetBonus(v.representant_id, v.entreprise_id, v.client_id, ym, v.id, priorVol);
        runningRepMonthVol.set(key, priorVol + Number(v.valeur || 0));

        const { tRep, tOli } = eff;
        const cRep = Math.round(Number(v.valeur) * tRep * 100) / 100;
        const cOli = Math.round(Number(v.valeur) * tOli * 100) / 100;

        const existingOv = saleBonusOverridesStore[String(v.id)];
        if (existingOv && (existingOv.customManual || existingOv.noBonus)) {
          continue;
        }
        if (eff.bonusActive) {
          await saveSaleBonusOverride(v.id, { tRep, tOli, cRep, cOli });
        } else if (existingOv) {
          await saveSaleBonusOverride(v.id, null);
        }
        if (Number(v.taux_rep_applique) !== tRep || Number(v.taux_olivier_applique) !== tOli || Number(v.commission_rep) !== cRep || Number(v.commission_olivier) !== cOli) {
          const { error } = await sb.from('ventes').update({
            taux_rep_applique: tRep,
            taux_olivier_applique: tOli,
            commission_rep: cRep,
            commission_olivier: cOli
          }).eq('id', v.id);
          if (!error) updated++;
        }
      }
      toast(`${updated} vente(s) recalculée(s)`);
      await loadAllData();
    });

    // =========================================================================
    // ASSISTANT VOCAL INTELLIGENT (1 PHRASE COMPLÈTE OU ÉTAPE PAR ÉTAPE + CLAVIER SIRI/GOOGLE)
    // =========================================================================
    let activeVoiceRecognition = null;
    let isVoiceListening = false;
    let isVoiceSpeaking = false;
    let voiceMuted = false;
    let currentVoiceStepIdx = 0;
    let voiceStepAdvanceTimer = null;
    let voiceInterimIdleTimer = null;
    let voiceLastHeardText = '';
    let voicePreferredLang = 'fr-CA';

    const VOICE_SALE_STEPS = [
      {
        key: 'entreprise',
        prompt: '1. Entreprise partenaire',
        spokenPrompt: 'Quelle entreprise ?',
        hint: 'Toitex, MaClim, MonSpray, VerreVDM ou Diamant Pro'
      },
      {
        key: 'client',
        prompt: '2. Nom du client',
        spokenPrompt: 'Nom du client ?',
        hint: 'Prénom et nom du client'
      },
      {
        key: 'montant',
        prompt: '3. Montant avant taxes ($)',
        spokenPrompt: 'Quel montant ?',
        hint: 'Ex. 14 500 dollars'
      },
      {
        key: 'type_kpi',
        prompt: '4. Type : OTS ou Suivi',
        spokenPrompt: 'O T S ou suivi ?',
        hint: 'Dites « OTS » ou « Suivi »'
      },
      {
        key: 'depot',
        get prompt() { return companyNoDeposit() ? '5. Dépôt : aucun requis pour cette entreprise' : '5. Dépôt : Reçu ou À confirmer'; },
        get spokenPrompt() { return companyNoDeposit() ? 'Cette entreprise ne demande aucun dépôt. Dites suivant pour continuer.' : 'Dépôt reçu ou à confirmer ?'; },
        get hint() { return companyNoDeposit() ? 'Dites « Suivant »' : 'Dites « Reçu » ou « À confirmer »'; }
      },
      {
        key: 'mois_comm',
        prompt: '6. Mois de paiement de la commission',
        spokenPrompt: 'Dans quel mois la commission doit-elle être payée ?',
        hint: 'Ex. « Ce mois-ci », « Décembre », « Janvier »...'
      },
      {
        key: 'note',
        prompt: '7. Contrat ou note (ou « Aucun »)',
        spokenPrompt: 'Numéro de contrat ? Sinon dites aucun.',
        hint: 'Ex. Est 3505 ou « Aucun »'
      }
    ];

    function playSoftVoiceBeep(freq = 660) {
      try {
        const ctx = sharedAudioCtx || new (window.AudioContext || window.webkitAudioContext)();
        if (ctx.state === 'suspended') ctx.resume();
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        osc.type = 'sine';
        osc.frequency.setValueAtTime(freq, ctx.currentTime);
        gain.gain.setValueAtTime(0.08, ctx.currentTime);
        gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.16);
        osc.connect(gain);
        gain.connect(ctx.destination);
        osc.start();
        osc.stop(ctx.currentTime + 0.16);
      } catch (e) {}
    }

    function getBestFrenchTtsVoice() {
      if (!('speechSynthesis' in window)) return null;
      const voices = window.speechSynthesis.getVoices() || [];
      return voices.find(v => /^fr[-_]CA/i.test(v.lang))
        || voices.find(v => /^fr/i.test(v.lang) && /natural|google|amelie|amélie|thomas|siri|microsoft/i.test(v.name))
        || voices.find(v => /^fr/i.test(v.lang))
        || null;
    }

    if ('speechSynthesis' in window) {
      window.speechSynthesis.getVoices();
      window.speechSynthesis.onvoiceschanged = () => { window.speechSynthesis.getVoices(); };
    }

    function cancelVoiceAssistantSpeech() {
      isVoiceSpeaking = false;
      if ('speechSynthesis' in window) {
        try { window.speechSynthesis.cancel(); } catch (e) {}
      }
    }

    function speakVoiceAssistantPrompt(textToSpeak, onFinishedCallback = null) {
      const cleanSpeech = String(textToSpeak || '').trim();
      if (!cleanSpeech || voiceMuted || !('speechSynthesis' in window)) {
        isVoiceSpeaking = false;
        playSoftVoiceBeep(740);
        if (typeof onFinishedCallback === 'function') setTimeout(onFinishedCallback, 120);
        return;
      }

      cancelVoiceAssistantSpeech();
      isVoiceSpeaking = true;

      const statusEl = document.getElementById('voice-sale-status');
      if (statusEl) statusEl.textContent = '🔊 L’assistant parle...';

      let callbackCalled = false;
      const done = () => {
        if (callbackCalled) return;
        callbackCalled = true;
        isVoiceSpeaking = false;
        if (typeof onFinishedCallback === 'function') onFinishedCallback();
      };

      try {
        const utter = new SpeechSynthesisUtterance(cleanSpeech);
        utter.lang = 'fr-CA';
        const chosenVoice = getBestFrenchTtsVoice();
        if (chosenVoice) utter.voice = chosenVoice;
        utter.rate = 1.1;
        utter.pitch = 1.0;

        const maxWaitMs = Math.min(3500, Math.max(1000, cleanSpeech.length * 65));
        const fallbackTimer = setTimeout(done, maxWaitMs);

        utter.onend = () => {
          clearTimeout(fallbackTimer);
          setTimeout(done, 320);
        };
        utter.onerror = () => {
          clearTimeout(fallbackTimer);
          done();
        };

        window.speechSynthesis.speak(utter);
      } catch (e) {
        done();
      }
    }

    window.toggleVoiceMute = () => {
      voiceMuted = !voiceMuted;
      if (voiceMuted) cancelVoiceAssistantSpeech();
      const btn = document.getElementById('btn-voice-mute-toggle');
      if (btn) {
        btn.textContent = voiceMuted ? '🔇 Bip seul' : '🔊 Voix';
        btn.className = voiceMuted
          ? 'px-2 py-0.5 text-[10px] font-extrabold rounded-lg bg-slate-200 text-slate-700 transition'
          : 'px-2 py-0.5 text-[10px] font-extrabold rounded-lg bg-indigo-100 text-indigo-800 transition';
      }
    };

    function renderVoiceQuickChoiceChips() {
      const box = document.getElementById('voice-quick-chips');
      if (!box) return;
      const step = VOICE_SALE_STEPS[currentVoiceStepIdx];
      if (!step) { box.innerHTML = ''; return; }

      let chips = [];
      if (step.key === 'entreprise') {
        const entSel = document.getElementById('vente-entreprise');
        chips = Array.from(entSel?.options || []).map(o => ({
          label: o.textContent.trim(),
          say: o.textContent.trim()
        }));
      } else if (step.key === 'type_kpi') {
        chips = [
          { label: '⚡ OTS (Sur place)', say: 'OTS' },
          { label: '🔄 Suivi', say: 'Suivi' }
        ];
      } else if (step.key === 'depot') {
        chips = companyNoDeposit()
          ? [{ label: '🚫 Aucun dépôt requis', say: 'Aucun dépôt requis' }]
          : [
            { label: '✅ Dépôt reçu', say: 'Dépôt reçu' },
            { label: '⏳ À confirmer', say: 'À confirmer' }
          ];
      } else if (step.key === 'mois_comm') {
        const baseYm = document.getElementById('vente-mois-comm')?.value || currentMonthStr;
        const [by, bm] = baseYm.split('-').map(Number);
        for (let offset = 0; offset < 4; offset++) {
          const d = new Date(by, (bm - 1) + offset, 1);
          const ym = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
          const lbl = offset === 0 ? `📅 ${formatYearMonthLabelFr(ym)} (Ce mois)` : formatYearMonthLabelFr(ym);
          chips.push({ label: lbl, say: ym });
        }
      } else if (step.key === 'note') {
        chips = [
          { label: 'Aucun', say: 'Aucun' }
        ];
      }

      if (!chips.length) {
        box.classList.add('hidden');
        box.innerHTML = '';
        return;
      }
      box.classList.remove('hidden');
      box.innerHTML = chips.map(c => `
        <button type="button" onclick="pickVoiceQuickChoice('${esc(c.say)}')" class="px-2.5 py-1 text-[11px] font-extrabold rounded-lg bg-indigo-50 hover:bg-indigo-100 text-indigo-900 border border-indigo-200 transition active:scale-95">
          ${esc(c.label)}
        </button>
      `).join('');
    }

    window.pickVoiceQuickChoice = (spokenValue) => {
      clearTimeout(voiceStepAdvanceTimer);
      clearTimeout(voiceInterimIdleTimer);
      const res = applySingleStepVoiceAnswer(currentVoiceStepIdx, [spokenValue]);
      if (!res.ok) return;

      if (activeVoiceRecognition) {
        try { activeVoiceRecognition.onend = null; activeVoiceRecognition.stop(); } catch (e) {}
        activeVoiceRecognition = null;
      }

      if (currentVoiceStepIdx < VOICE_SALE_STEPS.length - 1) {
        currentVoiceStepIdx++;
        updateVoiceStepUI();
        if (isVoiceListening) {
          startListeningForCurrentStep(true, res.spokenAck);
        } else {
          const statusEl = document.getElementById('voice-sale-status');
          if (statusEl) statusEl.textContent = `Étape ${currentVoiceStepIdx + 1} / ${VOICE_SALE_STEPS.length} — ${VOICE_SALE_STEPS[currentVoiceStepIdx].hint}`;
          speakVoiceAssistantPrompt(`${res.spokenAck}. ${VOICE_SALE_STEPS[currentVoiceStepIdx].spokenPrompt}`);
        }
      } else {
        stopVoiceSaleDictation(true);
        const statusEl = document.getElementById('voice-sale-status');
        if (statusEl) statusEl.textContent = '✅ Dictée terminée ! Cliquez sur Enregistrer.';
        speakVoiceAssistantPrompt(`${res.spokenAck}. Formulaire prêt.`);
      }
    };

    function updateVoiceStepUI() {
      const step = VOICE_SALE_STEPS[currentVoiceStepIdx] || VOICE_SALE_STEPS[0];
      const badgeEl = document.getElementById('voice-step-badge');
      const promptEl = document.getElementById('voice-step-prompt');
      if (badgeEl) badgeEl.textContent = `Étape ${currentVoiceStepIdx + 1} / ${VOICE_SALE_STEPS.length}`;
      if (promptEl) promptEl.textContent = `${step.prompt}`;
      renderVoiceQuickChoiceChips();
    }

    function stopVoiceSaleDictation(keepSpeech = false) {
      isVoiceListening = false;
      clearTimeout(voiceStepAdvanceTimer);
      clearTimeout(voiceInterimIdleTimer);
      if (!keepSpeech) cancelVoiceAssistantSpeech();
      if (activeVoiceRecognition) {
        try { activeVoiceRecognition.onend = null; activeVoiceRecognition.stop(); } catch (e) {}
        activeVoiceRecognition = null;
      }
      const btn = document.getElementById('btn-voice-sale-toggle');
      const iconEl = document.getElementById('voice-sale-btn-icon');
      const lblEl = document.getElementById('voice-sale-btn-label');
      if (btn) {
        btn.classList.remove('bg-rose-600', 'hover:bg-rose-700', 'animate-pulse');
        btn.classList.add('bg-indigo-600', 'hover:bg-indigo-700');
      }
      if (iconEl) iconEl.textContent = '🎙️';
      if (lblEl) lblEl.textContent = 'Dicter';
    }

    window.toggleVoiceManualInputBox = () => {
      const box = document.getElementById('voice-sale-manual-box');
      if (!box) return;
      box.classList.toggle('hidden');
      if (!box.classList.contains('hidden')) {
        document.getElementById('voice-sale-live-box')?.classList.remove('hidden');
        updateVoiceStepUI();
        document.getElementById('voice-sale-manual-input')?.focus();
      }
    };

    window.clearVoiceSaleTranscript = () => {
      currentVoiceStepIdx = 0;
      voiceLastHeardText = '';
      const tEl = document.getElementById('voice-sale-transcript');
      const sEl = document.getElementById('voice-sale-summary');
      const mEl = document.getElementById('voice-sale-manual-input');
      if (tEl) tEl.textContent = '';
      if (sEl) sEl.textContent = '';
      if (mEl) mEl.value = '';
      updateVoiceStepUI();
    };

    window.prevVoiceStep = () => {
      clearTimeout(voiceStepAdvanceTimer);
      clearTimeout(voiceInterimIdleTimer);
      currentVoiceStepIdx = Math.max(0, currentVoiceStepIdx - 1);
      updateVoiceStepUI();
      if (isVoiceListening) startListeningForCurrentStep(true);
    };

    window.skipVoiceStep = () => {
      clearTimeout(voiceStepAdvanceTimer);
      clearTimeout(voiceInterimIdleTimer);
      if (currentVoiceStepIdx < VOICE_SALE_STEPS.length - 1) {
        currentVoiceStepIdx++;
        updateVoiceStepUI();
        if (isVoiceListening) startListeningForCurrentStep(true);
      } else {
        stopVoiceSaleDictation(true);
        const stEl = document.getElementById('voice-sale-status');
        if (stEl) stEl.textContent = '✓ Toutes les étapes sont complétées !';
        speakVoiceAssistantPrompt('Formulaire prêt. Appuyez sur enregistrer.');
      }
    };

    window.retryCurrentVoiceStep = () => {
      clearTimeout(voiceStepAdvanceTimer);
      clearTimeout(voiceInterimIdleTimer);
      voiceLastHeardText = '';
      const tEl = document.getElementById('voice-sale-transcript');
      if (tEl) tEl.textContent = '';
      updateVoiceStepUI();
      startListeningForCurrentStep(true);
    };

    // Parseur intelligent d'une phrase complète (ex. "Toitex Jean Tremblay 14500 OTS dépôt reçu Est 3505")
    // Utilisé automatiquement si l'utilisateur dicte toute sa vente d'un coup via le micro ou via Siri/Gboard !
    function parseFullSentenceSaleIntoForm(rawPhrase) {
      const original = String(rawPhrase || '').trim();
      if (!original) return { count: 0, summary: [] };
      const norm = normalizePhoneticFr(original);
      const filledSummary = [];

      // 1. Entreprise
      const matchedEnt = matchEntrepriseFromAlternatives([original]);
      const entSel = document.getElementById('vente-entreprise');
      if (matchedEnt && entSel) {
        entSel.value = matchedEnt.value;
        flashVoiceFilledField(entSel);
        filledSummary.push(matchedEnt.textContent.trim());
      }

      // 2. Montant
      const numText = parseFrenchSpokenNumbers(original);
      const nums = Array.from(numText.matchAll(/(\d+(?:[.,]\d{1,2})?)/g))
        .map(m => parseAmount(m[1]))
        .filter(n => !isNaN(n) && n >= 50);
      if (nums.length) {
        const amt = Math.max(...nums);
        const valEl = document.getElementById('vente-valeur');
        if (valEl) {
          valEl.value = amt.toFixed(2).replace(/\.00$/, '');
          flashVoiceFilledField(valEl);
          filledSummary.push(fmtMoney(amt));
        }
      }

      // 3. Type OTS / Suivi
      const kpiSel = document.getElementById('vente-type-kpi');
      if (kpiSel) {
        if (/\b(?:suivi|suivis|relance)\b/i.test(norm)) {
          kpiSel.value = 'SUIVI';
          flashVoiceFilledField(kpiSel);
          filledSummary.push('Suivi');
        } else if (/\b(?:ots|o t s|au ts|hotesse|sur place)\b/i.test(norm)) {
          kpiSel.value = 'OTS';
          flashVoiceFilledField(kpiSel);
          filledSummary.push('OTS');
        }
      }

      // 4. Statut dépôt
      const depotSel = document.getElementById('vente-statut-depot');
      if (depotSel && !depotSel.disabled) {
        if (/\b(?:sans depot|aucun depot|pas de depot|depot non requis|aucun requis)\b/i.test(norm)) {
          depotSel.value = 'a_confirmer'; // cette entreprise exige un dépôt
          flashVoiceFilledField(depotSel);
          filledSummary.push('À confirmer');
        } else if (/\b(?:depot recu|recu|paye|virement|interac)\b/i.test(norm) && !/\b(?:pas recu|a confirmer)\b/i.test(norm)) {
          depotSel.value = 'depot_recu';
          flashVoiceFilledField(depotSel);
          filledSummary.push('Dépôt reçu');
        } else if (/\b(?:a confirmer|confirmer)\b/i.test(norm)) {
          depotSel.value = 'a_confirmer';
          flashVoiceFilledField(depotSel);
          filledSummary.push('À confirmer');
        }
      }

      // 5. Numéro de contrat (ex. Est3505 / Est 3505 / contrat 1234)
      const contractMatch = original.match(/\b(?:est|contrat)\s*[-#:]?\s*([a-z0-9-]{2,15})\b/i);
      if (contractMatch) {
        const noteEl = document.getElementById('vente-note');
        const cCode = contractMatch[0].replace(/^est\s+/i, 'Est').replace(/^contrat\s*[:#-]?\s*/i, '');
        if (noteEl && cCode) {
          noteEl.value = cCode;
          flashVoiceFilledField(noteEl);
          filledSummary.push(`Contrat ${cCode}`);
        }
      }

      // 6. Client : isoler les mots restants qui ne sont ni la compagnie, ni un nombre, ni un mot-clé
      let clientCandidate = original
        .replace(/\b(?:toitex|toit\s*ex|maclim|ma\s*clim|monspray|mon\s*spray|verrevdm|verre\s*vdm|vdm|groupe\s*diamant\s*pro|diamant\s*pro|ornithorynque)\b/gi, ' ')
        .replace(/\b(?:vente|vendu|pour|client|cliente|chez|montant|avant\s*taxes|dollars?|piastres?|cad|ots|o\.t\.s|suivi|depot|dépôt|recu|reçu|a\s*confirmer|à\s*confirmer|contrat|est\s*\d+)\b/gi, ' ')
        .replace(/[\d$.,]+/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();

      if (clientCandidate.length >= 3) {
        const entId = entSel?.value || '';
        const formattedCli = matchExistingClientFuzzy(clientCandidate, entId);
        const cliEl = document.getElementById('vente-client-nom');
        if (cliEl && formattedCli) {
          cliEl.value = formattedCli;
          flashVoiceFilledField(cliEl);
          filledSummary.push(formattedCli);
        }
      }

      if (typeof updateDrawerCommissionPreview === 'function') updateDrawerCommissionPreview();
      return { count: filledSummary.length, summary: filledSummary };
    }

    window.applyManualVoiceText = () => {
      const inEl = document.getElementById('voice-sale-manual-input');
      const val = (inEl?.value || '').trim();
      if (!val) return toast('Dictez ou écrivez une vente dans le champ.', 'error');

      // Si l'utilisateur a dicté plusieurs informations d'un coup via Siri / clavier Google (ex. "Toitex Jean Tremblay 12500 OTS"), tout remplir d'un coup !
      const wordsCount = val.split(/\s+/).length;
      if (wordsCount >= 3) {
        const multi = parseFullSentenceSaleIntoForm(val);
        if (multi.count >= 2) {
          inEl.value = '';
          const sEl = document.getElementById('voice-sale-summary');
          const stEl = document.getElementById('voice-sale-status');
          if (sEl) sEl.textContent = `✓ Rempli : ${multi.summary.join(' • ')}`;
          if (stEl) stEl.textContent = '✅ Vente remplie automatiquement !';
          toast(`✓ ${multi.count} champs remplis automatiquement`);
          return;
        }
      }

      // Sinon appliquer à l'étape en cours
      const res = applySingleStepVoiceAnswer(currentVoiceStepIdx, [val]);
      if (res.ok) {
        inEl.value = '';
        if (currentVoiceStepIdx < VOICE_SALE_STEPS.length - 1) {
          currentVoiceStepIdx++;
          updateVoiceStepUI();
          speakVoiceAssistantPrompt(`${res.spokenAck ? res.spokenAck + '. ' : ''}${VOICE_SALE_STEPS[currentVoiceStepIdx].spokenPrompt}`);
        } else {
          const stEl = document.getElementById('voice-sale-status');
          if (stEl) stEl.textContent = '✓ Formulaire complété ! Vérifiez et enregistrez.';
          speakVoiceAssistantPrompt(`${res.spokenAck ? res.spokenAck + '. ' : ''}Formulaire prêt.`);
        }
      } else {
        toast('Réponse non reconnue pour cette étape.', 'error');
      }
    };

    function normalizePhoneticFr(str) {
      return String(str || '')
        .toLowerCase()
        .normalize('NFD')
        .replace(/[̀-ͯ]/g, '')
        .replace(/[^a-z0-9\s-]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
    }

    function levenshteinDistance(a, b) {
      const s = normalizePhoneticFr(a);
      const t = normalizePhoneticFr(b);
      if (!s) return t.length;
      if (!t) return s.length;
      const dp = Array.from({ length: s.length + 1 }, (_, i) => [i]);
      for (let j = 1; j <= t.length; j++) dp[0][j] = j;
      for (let i = 1; i <= s.length; i++) {
        for (let j = 1; j <= t.length; j++) {
          const cost = s[i - 1] === t[j - 1] ? 0 : 1;
          dp[i][j] = Math.min(dp[i - 1][j] + 1, dp[i][j - 1] + 1, dp[i - 1][j - 1] + cost);
        }
      }
      return dp[s.length][t.length];
    }

    function parseFrenchSpokenNumbers(text) {
      let s = String(text || '').toLowerCase()
        .replace(/\u00a0|\u202f/g, ' ')
        // 1. Retirer le symbole $ / dollars / piastres pour éviter toute interférence
        .replace(/dollars?|piastres?|cad|\$/gi, ' ')
        // 2. Décimales dictées explicitement (« virgule 50 », « point 75 », « et 50 sous »)
        .replace(/\b(?:et\s+)?(\d{1,2})\s*(?:sous|cents?)\b/gi, '.$1')
        .replace(/\s*(?:virgule|point)\s*(\d{1,2})\b/gi, '.$1');

      // 3. Regrouper les milliers avec espaces/virgules entre chiffres :
      //    ex. "14 500" -> "14500", "12,500" -> "12500", "1 250 000" -> "1250000"
      s = s.replace(/\b(\d{1,3})(?:[\s,](\d{3}))+\b/g, (m) => m.replace(/[\s,]+/g, ''));

      // 4. Formes mixtes chiffres + mots en français québécois :
      //    ex. "14 mille 500" -> "14500", "14 mille 5" -> "14500", "14 mille et demi" -> "14500", "14k" -> "14000"
      s = s
        .replace(/(\d+(?:\.\d+)?)\s*mille\s*et\s*demi\b/gi, (_, m) => String(Math.round(parseFloat(m) * 1000 + 500)))
        .replace(/(\d+)\s*mille\s*(\d{1,3})\b/gi, (_, m, rest) => {
          const rNum = parseInt(rest, 10);
          // Au Québec, "14 mille 5" = 14 500 ; "14 mille 50" = 14 050 ; "14 mille 500" = 14 500
          const add = rest.length === 1 ? rNum * 100 : rNum;
          return String(parseInt(m, 10) * 1000 + add);
        })
        .replace(/(\d+(?:\.\d+)?)\s*mille\b/gi, (_, m) => String(Math.round(parseFloat(m) * 1000)))
        .replace(/(\d+(?:\.\d+)?)\s*k\b/gi, (_, k) => String(Math.round(parseFloat(k) * 1000)))
        .replace(/(\d+)\s*cents?\b/gi, (_, c) => String(parseInt(c, 10) * 100));

      // 5. Conversion complète des nombres écrits en toutes lettres (ex. "quatorze mille cinq cents", "neuf mille huit cent cinquante")
      const units = {
        'zero': 0, 'zéro': 0, 'un': 1, 'une': 1, 'deux': 2, 'trois': 3, 'quatre': 4,
        'cinq': 5, 'six': 6, 'sept': 7, 'huit': 8, 'neuf': 9, 'dix': 10,
        'onze': 11, 'douze': 12, 'treize': 13, 'quatorze': 14, 'quinze': 15,
        'seize': 16, 'dix-sept': 17, 'dix sept': 17, 'dix-huit': 18, 'dix huit': 18,
        'dix-neuf': 19, 'dix neuf': 19, 'vingt': 20, 'vingts': 20, 'trente': 30, 'quarante': 40,
        'cinquante': 50, 'soixante': 60, 'soixante-dix': 70, 'soixante dix': 70,
        'quatre-vingt': 80, 'quatre-vingts': 80, 'quatre vingt': 80,
        'quatre-vingt-dix': 90, 'quatre vingt dix': 90
      };

      const wordGroupRegex = /\b((?:(?:zero|zéro|un|une|deux|trois|quatre|cinq|six|sept|huit|neuf|dix|onze|douze|treize|quatorze|quinze|seize|vingts?|trente|quarante|cinquante|soixante|quatre-vingts?|quatre\s+vingts?|cents?|mille|et|\d+)[\s-]*)+)\b/gi;
      s = s.replace(wordGroupRegex, (match) => {
        const clean = match.trim();
        if (!/(mille|cent|vingt|trente|quarante|cinquante|soixante|quatorze|quinze|douze|treize|seize|onze|dix|deux|trois|quatre|cinq|six|sept|huit|neuf)/i.test(clean)) {
          return match;
        }
        const tokens = clean.split(/[\s-]+/).filter(Boolean);
        let total = 0;
        let current = 0;
        let matchedAny = false;
        for (const tok of tokens) {
          if (tok === 'et') continue;
          if (/^\d+$/.test(tok)) {
            current += parseInt(tok, 10);
            matchedAny = true;
          } else if (tok === 'cent' || tok === 'cents') {
            current = (current === 0 ? 1 : current) * 100;
            matchedAny = true;
          } else if (tok === 'mille') {
            total += (current === 0 ? 1 : current) * 1000;
            current = 0;
            matchedAny = true;
          } else if (units[tok] !== undefined) {
            current += units[tok];
            matchedAny = true;
          }
        }
        return matchedAny ? ` ${total + current} ` : match;
      });

      // 6. Si le micro a transcrit deux morceaux côte à côte pour un montant en milliers
      //    ex. "14 500" après conversion de "quatorze 500" -> "14500", ou "8 950" -> "8950"
      s = s.replace(/\b(\d{1,3})\s+(\d{3})\b/g, '$1$2');

      return s;
    }

    function flashVoiceFilledField(el) {
      if (!el) return;
      el.classList.add('ring-2', 'ring-emerald-500', 'bg-emerald-50/40');
      setTimeout(() => el.classList.remove('ring-2', 'ring-emerald-500', 'bg-emerald-50/40'), 1400);
    }

    function formatTitleCaseName(raw) {
      return String(raw || '')
        .trim()
        .split(/\s+/)
        .map(part => part.split('-').map(sub => sub ? (sub[0].toUpperCase() + sub.slice(1).toLowerCase()) : '').join('-'))
        .join(' ');
    }

    function matchExistingClientFuzzy(spokenClient, entId) {
      const clean = String(spokenClient || '').trim();
      if (!clean) return '';
      const normClean = normalizePhoneticFr(clean);
      const pool = (state.clients || []).filter(c => !entId || String(c.entreprise_id) === String(entId));
      const allPool = pool.length ? pool : (state.clients || []);

      const exact = allPool.find(c => normalizePhoneticFr(c.nom) === normClean);
      if (exact) return exact.nom;

      const fuzzy = allPool.find(c => {
        const cn = normalizePhoneticFr(c.nom);
        if (cn.length < 4) return false;
        if (cn.includes(normClean) || normClean.includes(cn)) return true;
        return levenshteinDistance(cn, normClean) <= 2;
      });
      if (fuzzy) return fuzzy.nom;

      return formatTitleCaseName(clean);
    }

    function matchEntrepriseFromAlternatives(alternativesList) {
      const entSel = document.getElementById('vente-entreprise');
      if (!entSel) return null;

      const entAliases = [
        { keys: ['toitex', 'toit ex', 'toit x', 'toit tex', 'toit tech', 'toitech', 'toiture', 'toi tex', 'toit text', 'toit texte', 'toit', 'tex'], match: 'toitex' },
        { keys: ['maclim', 'ma clim', 'ma clime', 'mac lim', 'maclime', 'clim', 'clime', 'climatisation', 'macleen', 'mcclim', 'ma clean'], match: 'maclim' },
        { keys: ['monspray', 'mon spray', 'mon spre', 'mons pre', 'spray', 'mon pre', 'mon pret', 'mon sprey', 'mon esprit'], match: 'monspray' },
        { keys: ['verrevdm', 'verre vdm', 'vers vdm', 'vdm', 'v d m', 'cloture', 'verre', 'vers', 'bden', 'bdm', 'vtm', 'vbm'], match: 'verrevdm' },
        { keys: ['diamant pro', 'groupe diamant', 'diamant', 'diamants', 'pro'], match: 'diamant' },
        { keys: ['ornithorynque', 'ornitho'], match: 'ornithorynque' }
      ];

      for (const alt of alternativesList) {
        const normNoAcc = normalizePhoneticFr(alt);
        if (!normNoAcc) continue;
        for (const item of entAliases) {
          if (item.keys.some(k => normNoAcc.includes(k))) {
            for (const opt of entSel.options) {
              const optNorm = normalizePhoneticFr(opt.textContent);
              if (optNorm.includes(item.match)) return opt;
            }
          }
        }
      }

      for (const alt of alternativesList) {
        const words = normalizePhoneticFr(alt).split(' ').filter(w => w.length >= 3);
        for (const opt of entSel.options) {
          const optNorm = normalizePhoneticFr(opt.textContent);
          if (!optNorm) continue;
          if (optNorm.includes(normalizePhoneticFr(alt)) || normalizePhoneticFr(alt).includes(optNorm)) return opt;
          for (const w of words) {
            if (levenshteinDistance(w, optNorm) <= 2) return opt;
          }
        }
      }

      return null;
    }

    function applySingleStepVoiceAnswer(stepIdx, rawAlternatives) {
      const alts = Array.isArray(rawAlternatives)
        ? rawAlternatives.map(x => String(x || '').trim()).filter(Boolean)
        : [String(rawAlternatives || '').trim()].filter(Boolean);
      const step = VOICE_SALE_STEPS[stepIdx];
      if (!step || !alts.length) return { ok: false, spokenAck: '' };

      const primaryOriginal = alts[0];
      const primaryNorm = normalizePhoneticFr(primaryOriginal);
      const summaryEl = document.getElementById('voice-sale-summary');

      if (alts.some(a => /^(?:passer|passe|suivant|prochain|skip)$/i.test(normalizePhoneticFr(a)))) {
        if (summaryEl) summaryEl.textContent = 'Étape passée';
        return { ok: true, spokenAck: 'Passé' };
      }

      if (step.key === 'entreprise') {
        const entSel = document.getElementById('vente-entreprise');
        const matchedOpt = matchEntrepriseFromAlternatives(alts);
        if (entSel && matchedOpt) {
          entSel.value = matchedOpt.value;
          flashVoiceFilledField(entSel);
          if (typeof updateDrawerCommissionPreview === 'function') updateDrawerCommissionPreview();
          const label = matchedOpt.textContent.trim();
          if (summaryEl) summaryEl.textContent = `✓ Entreprise : ${label}`;
          return { ok: true, spokenAck: label };
        }
        return { ok: false, spokenAck: '' };
      }

      if (step.key === 'client') {
        const cleanedClient = primaryOriginal
          .replace(/^(?:le\s+client\s+est|client|cliente|nom\s+du\s+client|c'est|pour|chez|monsieur|madame|m\.|mme)\s+/i, '')
          .replace(/[.,!?]+$/g, '')
          .trim();
        if (cleanedClient.length < 2) return { ok: false, spokenAck: '' };
        const entId = document.getElementById('vente-entreprise')?.value || '';
        const formattedClient = matchExistingClientFuzzy(cleanedClient, entId);
        const cliEl = document.getElementById('vente-client-nom');
        if (cliEl) {
          cliEl.value = formattedClient;
          flashVoiceFilledField(cliEl);
        }
        if (typeof updateDrawerCommissionPreview === 'function') updateDrawerCommissionPreview();
        if (summaryEl) summaryEl.textContent = `✓ Client : ${formattedClient}`;
        return { ok: true, spokenAck: formattedClient };
      }

      if (step.key === 'montant') {
        for (const alt of alts) {
          const normalized = parseFrenchSpokenNumbers(alt);
          // Capturer aussi les formats avec espaces insécables ou virgules de milliers avant normalisation
          const rawMatches = Array.from(normalized.matchAll(/(\d+(?:\.\d{1,2})?)/g))
            .map(m => parseFloat(m[1]))
            .filter(n => !isNaN(n) && n > 0);

          if (rawMatches.length) {
            // Si deux nombres consécutifs forment un millier + centaines ex: [14, 500] -> 14500
            let amount = 0;
            if (rawMatches.length === 2 && rawMatches[0] >= 1 && rawMatches[0] <= 999 && rawMatches[1] >= 100 && rawMatches[1] <= 999) {
              amount = rawMatches[0] * 1000 + rawMatches[1];
            } else {
              amount = Math.max(...rawMatches);
            }
            amount = Math.round(amount * 100) / 100;

            const valEl = document.getElementById('vente-valeur');
            if (valEl) {
              valEl.value = amount.toFixed(2).replace(/\.00$/, '');
              flashVoiceFilledField(valEl);
            }
            if (typeof updateDrawerCommissionPreview === 'function') updateDrawerCommissionPreview();
            if (summaryEl) summaryEl.textContent = `✓ Montant : ${fmtMoney(amount)}`;
            const spokenAmt = `${Math.round(amount) === amount ? amount : amount.toFixed(2).replace('.', ' virgule ')} dollars`;
            return { ok: true, spokenAck: spokenAmt };
          }
        }
        return { ok: false, spokenAck: '' };
      }

      if (step.key === 'type_kpi') {
        const kpiSel = document.getElementById('vente-type-kpi');
        if (!kpiSel) return { ok: false, spokenAck: '' };
        for (const alt of alts) {
          const n = normalizePhoneticFr(alt);
          const compact = n.replace(/\s+/g, '');
          if (n.includes('suivi') || n.includes('relance') || n.includes('suivre') || compact.includes('suivi')) {
            kpiSel.value = 'SUIVI';
            flashVoiceFilledField(kpiSel);
            if (summaryEl) summaryEl.textContent = '✓ Type : Suivi';
            return { ok: true, spokenAck: 'Suivi' };
          }
          if (
            compact.includes('ots') ||
            compact.includes('auts') ||
            compact.includes('hautes') ||
            compact.includes('hotesse') ||
            n.includes('o t s') ||
            n.includes('au ts') ||
            n.includes('aux ts') ||
            n.includes('sur place') ||
            n.includes('one time') ||
            /\b(?:ots|ts|auto)\b/i.test(n)
          ) {
            kpiSel.value = 'OTS';
            flashVoiceFilledField(kpiSel);
            if (summaryEl) summaryEl.textContent = '✓ Type : OTS';
            return { ok: true, spokenAck: 'O T S' };
          }
        }
        return { ok: false, spokenAck: '' };
      }

      if (step.key === 'depot') {
        const depotSel = document.getElementById('vente-statut-depot');
        if (!depotSel) return { ok: false, spokenAck: '' };
        if (companyNoDeposit()) {
          applyCompanyDepositRule();
          flashVoiceFilledField(depotSel);
          if (summaryEl) summaryEl.textContent = '✓ Dépôt : aucun requis (entreprise)';
          return { ok: true, spokenAck: 'Aucun dépôt requis' };
        }
        for (const alt of alts) {
          const n = normalizePhoneticFr(alt);
          if (/\b(?:aucun depot|sans depot|pas de depot|depot non requis|non requis|pas requis|aucun requis)\b/i.test(n)) {
            depotSel.value = 'a_confirmer'; // cette entreprise exige un dépôt
            applyCompanyDepositRule();
            flashVoiceFilledField(depotSel);
            if (summaryEl) summaryEl.textContent = '✓ Dépôt : À confirmer (cette entreprise exige un dépôt)';
            return { ok: true, spokenAck: 'Cette entreprise exige un dépôt, noté à confirmer' };
          }
          if (/\b(?:confirmer|non|pas|aucun|attente|venir|pas recu)\b/i.test(n)) {
            depotSel.value = 'a_confirmer';
            flashVoiceFilledField(depotSel);
            if (summaryEl) summaryEl.textContent = '✓ Dépôt : À confirmer';
            return { ok: true, spokenAck: 'À confirmer' };
          }
          if (/\b(?:recu|recue|fait|paye|oui|virement|interac|cheque|acompte|depot recu|ok)\b/i.test(n)) {
            depotSel.value = 'depot_recu';
            flashVoiceFilledField(depotSel);
            if (summaryEl) summaryEl.textContent = '✓ Dépôt : Dépôt reçu';
            return { ok: true, spokenAck: 'Dépôt reçu' };
          }
        }
        return { ok: false, spokenAck: '' };
      }

      if (step.key === 'mois_comm') {
        const moisEl = document.getElementById('vente-mois-comm');
        if (!moisEl) return { ok: false, spokenAck: '' };
        const baseYm = moisEl.value || currentMonthStr;
        const [curY, curM] = currentMonthStr.split('-').map(Number);

        const monthTokens = [
          { num: 1, keys: ['janvier', 'janv'] },
          { num: 2, keys: ['fevrier', 'fevr'] },
          { num: 3, keys: ['mars'] },
          { num: 4, keys: ['avril', 'avr'] },
          { num: 5, keys: ['mai'] },
          { num: 6, keys: ['juin'] },
          { num: 7, keys: ['juillet', 'juill'] },
          { num: 8, keys: ['aout'] },
          { num: 9, keys: ['septembre', 'sept'] },
          { num: 10, keys: ['octobre', 'oct'] },
          { num: 11, keys: ['novembre', 'nov'] },
          { num: 12, keys: ['decembre', 'dec'] }
        ];

        for (const alt of alts) {
          const rawTrim = String(alt || '').trim();
          const n = normalizePhoneticFr(rawTrim);
          let resolvedYm = '';

          if (/^\d{4}-(?:0[1-9]|1[0-2])$/.test(rawTrim)) {
            resolvedYm = rawTrim;
          } else if (/\b(?:ce mois|mois ci|courant|actuel|maintenant|meme mois)\b/i.test(n)) {
            resolvedYm = baseYm;
          } else if (/\b(?:mois prochain|prochain mois|suivant)\b/i.test(n)) {
            const dNext = new Date(curY, curM, 1);
            resolvedYm = `${dNext.getFullYear()}-${String(dNext.getMonth() + 1).padStart(2, '0')}`;
          } else {
            const foundMonth = monthTokens.find(mt => mt.keys.some(k => new RegExp(`\\b${k}\\b`, 'i').test(n)));
            if (foundMonth) {
              const yearMatch = n.match(/\b(20\d{2})\b/);
              let targetYear = yearMatch ? parseInt(yearMatch[1], 10) : curY;
              // Si aucun millésime n'est dit et que le mois nommé est déjà passé (ex: en nov 2026 on dit "janvier"), basculer sur l'année suivante (2027)
              if (!yearMatch && foundMonth.num < curM) {
                targetYear = curY + 1;
              }
              resolvedYm = `${targetYear}-${String(foundMonth.num).padStart(2, '0')}`;
            }
          }

          if (resolvedYm) {
            moisEl.value = resolvedYm;
            flashVoiceFilledField(moisEl);
            if (typeof updateDrawerCommissionPreview === 'function') updateDrawerCommissionPreview();
            const spokenMonth = formatYearMonthLabelFr(resolvedYm);
            if (summaryEl) summaryEl.textContent = `✓ Paiement : ${spokenMonth}`;
            return { ok: true, spokenAck: spokenMonth };
          }
        }
        return { ok: false, spokenAck: '' };
      }

      if (step.key === 'note') {
        const noteEl = document.getElementById('vente-note');
        if (alts.some(a => /^(?:aucun|aucune|rien|non|pas de note|pas de contrat|zero|effacer|vider)$/i.test(normalizePhoneticFr(a)))) {
          if (noteEl) {
            noteEl.value = '';
            flashVoiceFilledField(noteEl);
          }
          if (summaryEl) summaryEl.textContent = '✓ Aucun contrat / note (champ vidé)';
          return { ok: true, spokenAck: 'Aucun' };
        }
        let cleanNote = primaryOriginal
          .replace(/^(?:numero\s+de\s+contrat|numéro\s+de\s+contrat|contrat|note)\s*[:\-]?\s*/i, '')
          .replace(/[.,!?]+$/g, '')
          .trim();
        if (/^est\s+\d+$/i.test(cleanNote)) {
          cleanNote = cleanNote.replace(/^est\s+/i, 'Est');
        }
        if (cleanNote && noteEl) {
          noteEl.value = cleanNote;
          flashVoiceFilledField(noteEl);
          if (summaryEl) summaryEl.textContent = `✓ Note : ${cleanNote}`;
          return { ok: true, spokenAck: cleanNote };
        }
        return { ok: false, spokenAck: '' };
      }

      return { ok: false, spokenAck: '' };
    }

    function handleStepRecognizedAlternatives(alternativesList) {
      if (!alternativesList || !alternativesList.length) return;
      clearTimeout(voiceInterimIdleTimer);
      const statusEl = document.getElementById('voice-sale-status');

      // Si l'utilisateur a dicté une phrase complète dès l'étape 1 (ex. "Toitex Jean Tremblay 14500 OTS")
      if (currentVoiceStepIdx === 0 && alternativesList[0].split(/\s+/).length >= 4) {
        const multi = parseFullSentenceSaleIntoForm(alternativesList[0]);
        if (multi.count >= 3) {
          stopVoiceSaleDictation(true);
          if (statusEl) statusEl.textContent = `✅ Vente complète remplie (${multi.summary.join(' • ')})`;
          speakVoiceAssistantPrompt('Vente remplie dans le formulaire. Appuyez sur enregistrer.');
          return;
        }
      }

      const res = applySingleStepVoiceAnswer(currentVoiceStepIdx, alternativesList);

      if (res.ok) {
        if (activeVoiceRecognition) {
          try { activeVoiceRecognition.onend = null; activeVoiceRecognition.stop(); } catch (e) {}
          activeVoiceRecognition = null;
        }
        clearTimeout(voiceStepAdvanceTimer);
        voiceStepAdvanceTimer = setTimeout(() => {
          if (!isVoiceListening) return;
          if (currentVoiceStepIdx < VOICE_SALE_STEPS.length - 1) {
            currentVoiceStepIdx++;
            startListeningForCurrentStep(true, res.spokenAck);
          } else {
            stopVoiceSaleDictation(true);
            if (statusEl) statusEl.textContent = '✅ Dictée terminée ! Vérifiez et cliquez sur Enregistrer.';
            const entNom = document.getElementById('vente-entreprise')?.selectedOptions?.[0]?.textContent?.trim() || '';
            const cliNom = document.getElementById('vente-client-nom')?.value?.trim() || '';
            const valNum = parseFloat(document.getElementById('vente-valeur')?.value || '0') || 0;
            const recaps = [];
            if (res.spokenAck) recaps.push(res.spokenAck + '.');
            recaps.push(`Vente prête : ${entNom}, client ${cliNom}, ${Math.round(valNum)} dollars.`);
            speakVoiceAssistantPrompt(recaps.join(' '));
          }
        }, 250);
      } else {
        if (activeVoiceRecognition) {
          try { activeVoiceRecognition.onend = null; activeVoiceRecognition.stop(); } catch (e) {}
          activeVoiceRecognition = null;
        }
        if (statusEl) statusEl.textContent = `⚠️ Non compris (« ${alternativesList[0]} ») — répétez ou utilisez le micro clavier ⌨️`;
        speakVoiceAssistantPrompt(`Répétez. ${VOICE_SALE_STEPS[currentVoiceStepIdx].spokenPrompt}`, () => {
          if (isVoiceListening) beginMicrophoneRecognitionForCurrentStep();
        });
      }
    }

    function openKeyboardDictationFallback(reasonMsg = '') {
      stopVoiceSaleDictation();
      const liveBox = document.getElementById('voice-sale-live-box');
      const mBox = document.getElementById('voice-sale-manual-box');
      const statusEl = document.getElementById('voice-sale-status');
      const mInput = document.getElementById('voice-sale-manual-input');
      if (liveBox) liveBox.classList.remove('hidden');
      if (mBox) mBox.classList.remove('hidden');
      updateVoiceStepUI();
      if (statusEl && reasonMsg) statusEl.textContent = reasonMsg;
      if (mInput) {
        mInput.placeholder = 'Appuyez sur le micro 🎤 de votre clavier iPhone/Android et dictez toute la vente...';
        mInput.focus();
      }
    }

    function beginMicrophoneRecognitionForCurrentStep() {
      const SpeechRec = window.SpeechRecognition || window.webkitSpeechRecognition;
      if (!SpeechRec || !isVoiceListening) return;

      if (activeVoiceRecognition) {
        try { activeVoiceRecognition.onend = null; activeVoiceRecognition.stop(); } catch (e) {}
        activeVoiceRecognition = null;
      }

      const recognition = new SpeechRec();
      recognition.lang = voicePreferredLang;
      recognition.continuous = false;
      recognition.interimResults = true;
      recognition.maxAlternatives = 5;

      activeVoiceRecognition = recognition;
      voiceLastHeardText = '';

      const statusEl = document.getElementById('voice-sale-status');
      const transcriptEl = document.getElementById('voice-sale-transcript');
      if (statusEl) statusEl.textContent = `🔴 Micro actif (${voicePreferredLang}) : ${VOICE_SALE_STEPS[currentVoiceStepIdx].hint}`;
      if (transcriptEl) transcriptEl.textContent = '';

      let hasProcessedFinal = false;

      recognition.onresult = (event) => {
        if (isVoiceSpeaking || hasProcessedFinal) return;
        let interim = '';
        let finalAlternatives = [];

        for (let i = event.resultIndex; i < event.results.length; i++) {
          const resItem = event.results[i];
          if (resItem.isFinal) {
            for (let a = 0; a < resItem.length; a++) {
              const t = String(resItem[a].transcript || '').trim();
              if (t && !finalAlternatives.includes(t)) finalAlternatives.push(t);
            }
          } else {
            interim += resItem[0].transcript;
          }
        }

        const displayHeard = (finalAlternatives[0] || interim).trim();
        if (displayHeard) {
          voiceLastHeardText = displayHeard;
          if (transcriptEl) transcriptEl.textContent = `« ${displayHeard} »`;
        }

        if (finalAlternatives.length > 0) {
          hasProcessedFinal = true;
          handleStepRecognizedAlternatives(finalAlternatives);
        } else if (interim.trim()) {
          clearTimeout(voiceInterimIdleTimer);
          const capturedInterim = interim.trim();
          voiceInterimIdleTimer = setTimeout(() => {
            if (!hasProcessedFinal && isVoiceListening && !isVoiceSpeaking && capturedInterim) {
              hasProcessedFinal = true;
              handleStepRecognizedAlternatives([capturedInterim]);
            }
          }, 1300);
        }
      };

      recognition.onerror = (event) => {
        clearTimeout(voiceInterimIdleTimer);
        if (event.error === 'language-not-supported' && voicePreferredLang === 'fr-CA') {
          voicePreferredLang = 'fr-FR';
          if (isVoiceListening) beginMicrophoneRecognitionForCurrentStep();
          return;
        }
        if (event.error === 'not-allowed' || event.error === 'service-not-allowed') {
          openKeyboardDictationFallback('🎙️ Micro navigateur bloqué sur cet appareil — utilisez le micro du clavier ci-dessous');
          return;
        }
        if (event.error === 'no-speech' && isVoiceListening) {
          if (statusEl) statusEl.textContent = `⏳ Aucune voix captée — parlez ou utilisez le micro clavier ⌨️`;
        }
      };

      recognition.onend = () => {
        clearTimeout(voiceInterimIdleTimer);
        if (!hasProcessedFinal && voiceLastHeardText.trim() && isVoiceListening && !isVoiceSpeaking) {
          hasProcessedFinal = true;
          handleStepRecognizedAlternatives([voiceLastHeardText.trim()]);
          return;
        }
        if (isVoiceListening && !isVoiceSpeaking && !hasProcessedFinal) {
          try { recognition.start(); } catch (e) { openKeyboardDictationFallback('🎙️ Utilisez le micro de votre clavier ci-dessous'); }
        }
      };

      try {
        recognition.start();
      } catch (e) {
        openKeyboardDictationFallback('🎙️ Utilisez le micro de votre clavier ci-dessous');
      }
    }

    function startListeningForCurrentStep(speakQuestion = true, prefixAck = '') {
      const SpeechRec = window.SpeechRecognition || window.webkitSpeechRecognition;
      if (!SpeechRec) {
        openKeyboardDictationFallback('🎙️ Utilisez le micro 🎤 de votre clavier iPhone/Android ci-dessous');
        return;
      }

      if (activeVoiceRecognition) {
        try { activeVoiceRecognition.onend = null; activeVoiceRecognition.stop(); } catch (e) {}
        activeVoiceRecognition = null;
      }

      isVoiceListening = true;

      const btn = document.getElementById('btn-voice-sale-toggle');
      const iconEl = document.getElementById('voice-sale-btn-icon');
      const lblEl = document.getElementById('voice-sale-btn-label');
      const liveBox = document.getElementById('voice-sale-live-box');
      const mBox = document.getElementById('voice-sale-manual-box');
      const transcriptEl = document.getElementById('voice-sale-transcript');

      if (btn) {
        btn.classList.remove('bg-indigo-600', 'hover:bg-indigo-700');
        btn.classList.add('bg-rose-600', 'hover:bg-rose-700', 'animate-pulse');
      }
      if (iconEl) iconEl.textContent = '⏹️';
      if (lblEl) lblEl.textContent = 'Arrêter';
      if (liveBox) liveBox.classList.remove('hidden');
      if (mBox) mBox.classList.remove('hidden');
      updateVoiceStepUI();
      if (transcriptEl) transcriptEl.textContent = '';

      if (speakQuestion && !voiceMuted && 'speechSynthesis' in window) {
        const stepObj = VOICE_SALE_STEPS[currentVoiceStepIdx];
        const phrase = prefixAck ? `${prefixAck}. ${stepObj.spokenPrompt}` : stepObj.spokenPrompt;
        speakVoiceAssistantPrompt(phrase, () => {
          if (isVoiceListening) beginMicrophoneRecognitionForCurrentStep();
        });
      } else {
        playSoftVoiceBeep(740);
        beginMicrophoneRecognitionForCurrentStep();
      }
    }

    window.toggleVoiceSaleDictation = () => {
      if (isVoiceListening) {
        stopVoiceSaleDictation();
        const stEl = document.getElementById('voice-sale-status');
        if (stEl) stEl.textContent = '⏸️ Dictée en pause';
        return;
      }
      if (currentVoiceStepIdx >= VOICE_SALE_STEPS.length) currentVoiceStepIdx = 0;
      startListeningForCurrentStep(true);
    };

    // Validation automatique dès qu'on dicte via le micro natif Siri / Gboard dans `#voice-sale-manual-input`
    let manualDictationAutoTimer = null;
    document.getElementById('voice-sale-manual-input')?.addEventListener('input', (e) => {
      clearTimeout(manualDictationAutoTimer);
      const val = String(e.target.value || '').trim();
      if (!val) return;
      // Si l'utilisateur dicte une phrase complète via le clavier iPhone/Android (ex: "Toitex Jean Tremblay 14500 OTS"), remplir dès qu'il cesse de parler 1,1s
      if (val.split(/\s+/).length >= 3) {
        manualDictationAutoTimer = setTimeout(() => {
          applyManualVoiceText();
        }, 1100);
      }
    });

    const drawer = document.getElementById('sale-drawer-backdrop');
    function openNewSaleDrawer() {
      document.getElementById('vente-edit-id').value = '';
      if (document.getElementById('vente-rate-mode')) document.getElementById('vente-rate-mode').value = 'auto';
      document.getElementById('drawer-title').textContent = currentProfile?.role === 'admin' ? 'Enregistrer une nouvelle vente' : 'Déclarer une vente';
      document.getElementById('drawer-subtitle').textContent = 'Votre commission et le tableau KPI se mettent à jour automatiquement';
      document.getElementById('btn-submit-drawer').textContent = 'Enregistrer la vente';
      populateDrawerSelects();
      const nowLive = new Date();
      const liveDateStr = `${nowLive.getFullYear()}-${String(nowLive.getMonth() + 1).padStart(2, '0')}-${String(nowLive.getDate()).padStart(2, '0')}`;
      document.getElementById('vente-date').value = liveDateStr;
      document.getElementById('vente-mois-comm').value = liveDateStr.slice(0, 7);
      document.getElementById('vente-client-nom').value = '';
      document.getElementById('vente-valeur').value = '';
      document.getElementById('vente-note').value = '';
      document.getElementById('vente-statut-depot').value = 'a_confirmer';
      document.getElementById('vente-statut-paiement').value = 'non_paye';
      window._editingOriginalHadProof = false;
      const hitCb = document.getElementById('vente-is-hit');
      if (hitCb) hitCb.checked = false;
      const hitBearerSel = document.getElementById('vente-hit-bearer');
      if (hitBearerSel) hitBearerSel.value = 'both';
      const hitAmtIn = document.getElementById('vente-hit-amount');
      if (hitAmtIn) hitAmtIn.value = '';
      const hitAmtRepIn = document.getElementById('vente-hit-amount-rep');
      if (hitAmtRepIn) hitAmtRepIn.value = '';
      const hitAmtOliIn = document.getElementById('vente-hit-amount-oli');
      if (hitAmtOliIn) hitAmtOliIn.value = '';
      pendingProofDataUrl = '';
      const pfIn = document.getElementById('vente-proof-file');
      if (pfIn) pfIn.value = '';
      const cmIn = document.getElementById('vente-camera-input');
      if (cmIn) cmIn.value = '';
      document.getElementById('vente-proof-preview-box')?.classList.add('hidden');
      document.getElementById('btn-clear-proof')?.classList.add('hidden');
      document.getElementById('vente-proof-status')?.classList.add('hidden');
      stopVoiceSaleDictation();
      clearVoiceSaleTranscript();
      document.getElementById('voice-sale-live-box')?.classList.add('hidden');
      document.getElementById('voice-sale-manual-box')?.classList.add('hidden');
      drawer.classList.remove('hidden');
      updateDrawerCommissionPreview();
    }
    document.getElementById('btn-open-drawer')?.addEventListener('click', openNewSaleDrawer);
    document.getElementById('btn-rep-open-drawer')?.addEventListener('click', openNewSaleDrawer);
    document.getElementById('btn-mobile-fab-sale')?.addEventListener('click', openNewSaleDrawer);
    document.getElementById('btn-close-drawer')?.addEventListener('click', () => { stopVoiceSaleDictation(); drawer.classList.add('hidden'); });
    document.getElementById('btn-cancel-drawer')?.addEventListener('click', () => { stopVoiceSaleDictation(); drawer.classList.add('hidden'); });
    drawer.addEventListener('click', (e) => { if (e.target === drawer) { stopVoiceSaleDictation(); drawer.classList.add('hidden'); } });

    document.getElementById('btn-close-payslip')?.addEventListener('click', () => document.getElementById('payslip-modal').classList.add('hidden'));

    const userMenuBtn = document.getElementById('btn-user-menu');
    const userMenuDropdown = document.getElementById('user-menu-dropdown');
    if (userMenuBtn && userMenuDropdown) {
      userMenuBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        userMenuDropdown.classList.toggle('hidden');
      });
      document.addEventListener('click', (e) => {
        if (!userMenuDropdown.contains(e.target) && !userMenuBtn.contains(e.target)) {
          userMenuDropdown.classList.add('hidden');
        }
      });
    }

    const pwdModal = document.getElementById('pwd-modal');
    document.getElementById('btn-open-pwd')?.addEventListener('click', () => {
      userMenuDropdown?.classList.add('hidden');
      pwdModal.classList.remove('hidden');
    });
    document.getElementById('btn-close-pwd')?.addEventListener('click', () => pwdModal.classList.add('hidden'));
    document.getElementById('btn-cancel-pwd')?.addEventListener('click', () => pwdModal.classList.add('hidden'));
    document.getElementById('form-change-pwd')?.addEventListener('submit', async (e) => {
      e.preventDefault();
      const p1 = document.getElementById('new-pwd-1').value;
      const p2 = document.getElementById('new-pwd-2').value;
      if (p1 !== p2) return toast('Les deux mots de passe ne correspondent pas.', 'error');
      const { error } = await sb.auth.updateUser({ password: p1 });
      if (error) return toast(error.message, 'error');
      // Le nouveau mot de passe n'est jamais mémorisé dans le navigateur
      document.getElementById('new-pwd-1').value = '';
      document.getElementById('new-pwd-2').value = '';
      pwdModal.classList.add('hidden');
      toast('Mot de passe mis à jour avec succès');
    });

    ['vente-rep', 'vente-entreprise', 'vente-client-nom', 'vente-valeur', 'vente-mois-comm', 'vente-date'].forEach(id => {
      const el = document.getElementById(id);
      el.addEventListener('input', updateDrawerCommissionPreview);
      el.addEventListener('keyup', updateDrawerCommissionPreview);
      el.addEventListener('change', updateDrawerCommissionPreview);
    });

    document.getElementById('select-all-ventes')?.addEventListener('change', (e) => {
      const checked = e.target.checked;
      currentFilteredAdminSales.forEach(v => {
        if (checked) selectedSaleIds.add(v.id);
        else selectedSaleIds.delete(v.id);
      });
      document.querySelectorAll('.sale-row-cb').forEach(cb => cb.checked = checked);
      updateBulkBar();
    });

    document.getElementById('config-form')?.addEventListener('submit', (e) => {
      e.preventDefault();
      lsSet('korpo_sb_url', document.getElementById('cfg-url').value.trim());
      lsSet('korpo_sb_key', document.getElementById('cfg-key').value.trim());
      initApp();
    });

    // =========================================================================
    // SECTION FINANCES KORPO : ENCAISSEMENTS, PROFIT RÉEL, PRÉVISIONS & WAVE
    // =========================================================================
    const QC_TAX_RATE = 0.14975; // TPS 5 % + TVQ 9,975 %
    let finTaxMode = 'ht'; // 'ht' | 'ttc'
    let waveReconcileResults = null;

    function getFinExpensesStore() {
      return lsGetJson('korpo_fin_expenses_v1', '{}');
    }
    function saveFinExpensesStore(obj) {
      lsSet('korpo_fin_expenses_v1', JSON.stringify(obj || {}));
    }

    function getFinPartnerPaidStore() {
      return lsGetJson('korpo_fin_partner_paid_v1', '{}');
    }
    function saveFinPartnerPaidStore(obj) {
      lsSet('korpo_fin_partner_paid_v1', JSON.stringify(obj || {}));
    }

    window.setFinancesTaxMode = (mode) => {
      finTaxMode = mode === 'ttc' ? 'ttc' : 'ht';
      const btnHt = document.getElementById('btn-fin-tax-ht');
      const btnTtc = document.getElementById('btn-fin-tax-ttc');
      if (btnHt && btnTtc) {
        btnHt.className = finTaxMode === 'ht'
          ? 'px-3 py-1.5 rounded-lg bg-indigo-600 text-white shadow-2xs transition cursor-pointer'
          : 'px-3 py-1.5 rounded-lg text-slate-600 hover:text-slate-900 transition cursor-pointer';
        btnTtc.className = finTaxMode === 'ttc'
          ? 'px-3 py-1.5 rounded-lg bg-indigo-600 text-white shadow-2xs transition cursor-pointer'
          : 'px-3 py-1.5 rounded-lg text-slate-600 hover:text-slate-900 transition cursor-pointer';
      }
      renderAdminFinances();
    };

    function getFinPeriodMonths() {
      const mode = document.getElementById('fin-filter-period')?.value || 'this_month';
      const monthInp = document.getElementById('fin-filter-month');
      if (monthInp) {
        monthInp.classList.toggle('hidden', mode !== 'specific_month');
        if (mode === 'specific_month' && !monthInp.value) monthInp.value = currentMonthStr;
      }
      const allYms = new Set([currentMonthStr]);
      (state.ventes || []).forEach(v => {
        const ym = getSaleStatsMonth(v);
        if (isValidYearMonth(ym)) allYms.add(ym);
      });
      Object.values(getFinExpensesStore()).forEach(e => {
        if (e && e.ym && isValidYearMonth(e.ym)) allYms.add(e.ym);
      });
      const sortedAll = Array.from(allYms).sort();

      if (mode === 'this_month') return { label: formatYearMonthLabelFr(currentMonthStr), months: [currentMonthStr] };
      if (mode === 'last_month') {
        const prev = getPreviousYearMonth(currentMonthStr);
        return { label: formatYearMonthLabelFr(prev), months: [prev] };
      }
      if (mode === 'this_quarter') {
        const curQ = getQuarterForMonth(currentMonthStr);
        return { label: `Trimestre ${curQ.replace('-Q', ' T')}`, months: sortedAll.filter(ym => getQuarterForMonth(ym) === curQ) };
      }
      if (mode === 'this_year') {
        const y = String(today.getFullYear());
        return { label: `Année ${y}`, months: sortedAll.filter(ym => ym.startsWith(y)) };
      }
      if (mode === 'specific_month') {
        const m = monthInp?.value || currentMonthStr;
        return { label: formatYearMonthLabelFr(m), months: [m] };
      }
      return { label: 'Cumul complet', months: sortedAll };
    }

    function getPartnerBatchKey(entName, ym) {
      return `${ym}__${String(entName || '').trim()}`;
    }

    function computeKorpoGrossRevenueForSale(v) {
      // Revenu brut facturé par Korpo au partenaire = part rep + part Korpo (Olivier)
      const cRep = Number(v.commission_rep || 0);
      const cOli = Number(v.commission_olivier || 0);
      return Math.round((cRep + cOli) * 100) / 100;
    }

    window.togglePartnerBatchPaid = (batchKey) => {
      if (currentProfile?.role !== 'admin') return;
      const store = getFinPartnerPaidStore();
      if (store[batchKey]?.paid) {
        delete store[batchKey];
        toast('Facture partenaire remise en attente d’encaissement');
      } else {
        store[batchKey] = {
          paid: true,
          paidAt: localYmd(),
          by: currentProfile.nom || 'Admin'
        };
        toast('Encaissement partenaire confirmé ✓');
      }
      saveFinPartnerPaidStore(store);
      renderAdminFinances();
    };

    window.payAllReadyForRepInFinances = async (repId, monthsJson) => {
      if (currentProfile?.role !== 'admin') return;
      let months = [];
      try { months = JSON.parse(decodeURIComponent(monthsJson)); } catch (e) {}
      const partnerPaidStore = getFinPartnerPaidStore();
      const toPay = (state.ventes || []).filter(v => {
        if (String(v.representant_id) !== String(repId)) return false;
        if (v.statut_paiement === 'paye' || v.statut_depot !== 'depot_recu') return false;
        const ym = getSaleStatsMonth(v);
        if (months.length && !months.includes(ym)) return false;
        const entName = state.entreprises.find(e => String(e.id) === String(v.entreprise_id))?.nom || resolveSaleArchivedEntName(v) || '';
        return Boolean(partnerPaidStore[getPartnerBatchKey(entName, ym)]?.paid);
      });
      if (!toPay.length) return toast('Aucune commission encaissée prête à verser pour ce représentant.', 'error');
      const repNom = state.representants.find(r => String(r.id) === String(repId))?.nom || 'Représentant';
      const totalAmt = toPay.reduce((s, v) => s + Number(v.commission_rep || 0), 0);
      const ok = await askConfirm(
        `Marquer ${toPay.length} commission(s) payée(s) ?`,
        `Confirmer le versement de ${fmtMoney(totalAmt)} à ${repNom} (dossiers déjà encaissés par Korpo).`
      );
      if (!ok) return;
      const ids = toPay.map(v => v.id);
      const { error } = await sb.from('ventes').update({ statut_paiement: 'paye' }).in('id', ids);
      if (error) return toast(error.message, 'error');
      for (const id of ids) {
        await recordSaleAudit(id, 'Commission versée au représentant', `Payé depuis Finances Korpo (${fmtMoney(totalAmt)})`, '💸');
      }
      toast(`Versement de ${fmtMoney(totalAmt)} enregistré pour ${repNom} ✓`);
      await loadAllData();
    };

    window.deleteFinExpense = (expId) => {
      if (currentProfile?.role !== 'admin') return;
      const store = getFinExpensesStore();
      delete store[expId];
      saveFinExpensesStore(store);
      renderAdminFinances();
      toast('Dépense supprimée');
    };

    window.applyRecurringExpensesToCurrentMonth = () => {
      if (currentProfile?.role !== 'admin') return;
      const store = getFinExpensesStore();
      const existingThisMonth = Object.values(store).filter(e => e && e.ym === currentMonthStr);
      const recurringTemplates = new Map();
      Object.values(store).forEach(e => {
        if (e && e.recurring && e.label) {
          recurringTemplates.set(`${e.category}__${e.label.trim().toLowerCase()}`, e);
        }
      });
      if (!recurringTemplates.size) {
        return toast('Aucun frais fixe récurrent enregistré. Cochez « Frais fixe mensuel » lors de l’ajout d’une dépense.', 'error');
      }
      let added = 0;
      recurringTemplates.forEach((tpl, key) => {
        const alreadyExists = existingThisMonth.some(x => `${x.category}__${String(x.label || '').trim().toLowerCase()}` === key);
        if (!alreadyExists) {
          const id = `exp_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`;
          store[id] = {
            id,
            date: `${currentMonthStr}-01`,
            ym: currentMonthStr,
            category: tpl.category,
            label: tpl.label,
            amountHt: Number(tpl.amountHt || 0),
            recurring: true,
            createdAt: new Date().toISOString()
          };
          added++;
        }
      });
      if (!added) return toast(`Tous vos frais fixes (${recurringTemplates.size}) sont déjà présents pour ${currentMonthStr}.`);
      saveFinExpensesStore(store);
      renderAdminFinances();
      toast(`${added} frais fixe(s) généré(s) pour ${currentMonthStr} ✓`);
    };

    document.getElementById('form-fin-expense')?.addEventListener('submit', (e) => {
      e.preventDefault();
      if (currentProfile?.role !== 'admin') return;
      const dateVal = document.getElementById('fin-exp-date')?.value || localYmd();
      const catVal = document.getElementById('fin-exp-cat')?.value || 'Autre dépense';
      const labelVal = document.getElementById('fin-exp-label')?.value.trim() || '';
      const amtVal = parseFloat(document.getElementById('fin-exp-amount')?.value) || 0;
      const recVal = Boolean(document.getElementById('fin-exp-recurring')?.checked);
      if (!labelVal || amtVal <= 0) return toast('Entrez une description et un montant valide.', 'error');

      const id = `exp_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`;
      const store = getFinExpensesStore();
      store[id] = {
        id,
        date: dateVal,
        ym: dateVal.slice(0, 7),
        category: catVal,
        label: labelVal,
        amountHt: Math.round(amtVal * 100) / 100,
        recurring: recVal,
        createdAt: new Date().toISOString()
      };
      saveFinExpensesStore(store);
      document.getElementById('fin-exp-label').value = '';
      document.getElementById('fin-exp-amount').value = '';
      document.getElementById('fin-exp-recurring').checked = false;
      renderAdminFinances();
      toast(`Dépense enregistrée (${fmtMoney(amtVal)} HT) ✓`);
    });

    window.handleWaveReconcileCsvUpload = async (inputEl) => {
      const file = inputEl?.files?.[0];
      if (!file) return;
      try {
        const text = await file.text();
        const lines = text.split(/
?
/).filter(l => l.trim());
        if (lines.length < 2) return toast('Fichier CSV Wave vide.', 'error');
        const sep = lines[0].includes(';') ? ';' : ',';
        const headers = parseCsvLine(lines[0], sep).map(h => h.toLowerCase().trim());

        const idxCustomer = headers.findIndex(h => h.includes('customer') || h.includes('client') || h.includes('contact') || h.includes('description'));
        const idxAmount = headers.findIndex(h => h.includes('amount') || h.includes('total') || h.includes('montant') || h.includes('paid'));
        const idxStatus = headers.findIndex(h => h.includes('status') || h.includes('statut') || h.includes('state'));
        const idxInvNum = headers.findIndex(h => h.includes('invoice') || h.includes('facture') || h.includes('number'));

        const waveRows = [];
        for (let i = 1; i < lines.length; i++) {
          const cols = parseCsvLine(lines[i], sep);
          const amt = idxAmount >= 0 ? Math.abs(parseAmount(cols[idxAmount])) : NaN;
          if (isNaN(amt) || amt <= 0) continue;
          waveRows.push({
            customer: idxCustomer >= 0 ? String(cols[idxCustomer] || '').trim() : '',
            invoiceNum: idxInvNum >= 0 ? String(cols[idxInvNum] || '').trim() : '',
            status: idxStatus >= 0 ? String(cols[idxStatus] || '').trim() : '',
            amount: Math.round(amt * 100) / 100
          });
        }

        const partnerPaidStore = getFinPartnerPaidStore();
        const { months } = getFinPeriodMonths();
        const batches = buildPartnerBatchesForMonths(months);
        let autoMatched = 0;

        batches.forEach(b => {
          const match = waveRows.find(w => {
            const custMatch = w.customer.toLowerCase().includes(b.entName.toLowerCase()) ||
              w.invoiceNum.toLowerCase().startsWith(getWaveCompanyPrefix(b.entName).toLowerCase());
            if (!custMatch) return false;
            const isPaidStatus = !w.status || /paid|pay[eé]|complet|deposit|d[eé]p[oô]t/i.test(w.status);
            const amtMatchesHt = Math.abs(w.amount - b.grossHt) <= 2.0;
            const amtMatchesTtc = Math.abs(w.amount - b.grossTtc) <= 2.0;
            return isPaidStatus && (amtMatchesHt || amtMatchesTtc);
          });
          if (match && !partnerPaidStore[b.batchKey]?.paid) {
            partnerPaidStore[b.batchKey] = {
              paid: true,
              paidAt: localYmd(),
              waveRef: match.invoiceNum || 'Import CSV Wave',
              by: 'Rapprochement Wave'
            };
            autoMatched++;
          }
        });

        if (autoMatched > 0) saveFinPartnerPaidStore(partnerPaidStore);
        waveReconcileResults = {
          filename: file.name,
          totalRows: waveRows.length,
          autoMatched,
          totalWaveAmt: waveRows.reduce((s, r) => s + r.amount, 0)
        };
        renderAdminFinances();
        toast(`Rapprochement Wave terminé : ${autoMatched} lot(s) partenaire(s) rapproché(s) ✓`);
      } catch (err) {
        toast('Erreur lecture CSV Wave : ' + err.message, 'error');
      } finally {
        inputEl.value = '';
      }
    };

    function buildPartnerBatchesForMonths(monthsList) {
      const partnerPaidStore = getFinPartnerPaidStore();
      const map = new Map();
      (state.ventes || []).forEach(v => {
        if (v.statut_depot !== 'depot_recu') return;
        const ym = getSaleStatsMonth(v);
        if (monthsList && monthsList.length && !monthsList.includes(ym)) return;
        const entName = state.entreprises.find(e => String(e.id) === String(v.entreprise_id))?.nom || resolveSaleArchivedEntName(v) || 'Autre';
        const batchKey = getPartnerBatchKey(entName, ym);
        const cur = map.get(batchKey) || {
          batchKey,
          entName,
          ym,
          count: 0,
          salesVolume: 0,
          grossHt: 0,
          grossTtc: 0,
          repComm: 0,
          paidInfo: partnerPaidStore[batchKey] || null
        };
        const gross = computeKorpoGrossRevenueForSale(v);
        cur.count += 1;
        cur.salesVolume += Number(v.valeur || 0);
        cur.grossHt = Math.round((cur.grossHt + gross) * 100) / 100;
        cur.grossTtc = Math.round(cur.grossHt * (1 + QC_TAX_RATE) * 100) / 100;
        cur.repComm = Math.round((cur.repComm + Number(v.commission_rep || 0)) * 100) / 100;
        map.set(batchKey, cur);
      });
      return Array.from(map.values()).sort((a, b) => b.ym.localeCompare(a.ym) || b.grossHt - a.grossHt);
    }

    function renderAdminFinances() {
      if (currentProfile?.role !== 'admin') return;
      const expDateIn = document.getElementById('fin-exp-date');
      if (expDateIn && !expDateIn.value) expDateIn.value = localYmd();

      const { label: periodLabel, months } = getFinPeriodMonths();
      const partnerPaidStore = getFinPartnerPaidStore();
      const expensesStore = getFinExpensesStore();
      const taxMult = finTaxMode === 'ttc' ? (1 + QC_TAX_RATE) : 1;
      const modeTag = finTaxMode === 'ttc' ? 'TTC' : 'HT';

      // 1. Calcul des revenus, encaissements, commissions à recevoir et à payer sur la période
      const periodSales = (state.ventes || []).filter(v => {
        const ym = getSaleStatsMonth(v);
        return !months.length || months.includes(ym);
      });
      const confirmedSales = periodSales.filter(v => v.statut_depot === 'depot_recu');

      let totalFacturableHt = 0;
      let totalEncaisseHt = 0;
      let totalARecevoirPartenairesHt = 0;
      let totalCommRepsPayees = 0;
      let totalCommRepsAPayer = 0;

      confirmedSales.forEach(v => {
        const ym = getSaleStatsMonth(v);
        const entName = state.entreprises.find(e => String(e.id) === String(v.entreprise_id))?.nom || resolveSaleArchivedEntName(v) || 'Autre';
        const isPartnerPaid = Boolean(partnerPaidStore[getPartnerBatchKey(entName, ym)]?.paid);
        const grossHt = computeKorpoGrossRevenueForSale(v);
        const cRep = Number(v.commission_rep || 0);

        totalFacturableHt += grossHt;
        if (isPartnerPaid) totalEncaisseHt += grossHt;
        else totalARecevoirPartenairesHt += grossHt;

        if (v.statut_paiement === 'paye') totalCommRepsPayees += cRep;
        else totalCommRepsAPayer += cRep;
      });

      const periodExpenses = Object.values(expensesStore)
        .filter(e => e && (!months.length || months.includes(e.ym)))
        .sort((a, b) => String(b.date || '').localeCompare(String(a.date || '')));
      const totalExpensesHt = periodExpenses.reduce((s, e) => s + Number(e.amountHt || 0), 0);

      // Profit réel conservé par Korpo = Encaissé (ou facturable HT) - Commissions reps (payées + dues) - Dépenses HT
      const totalCommRepsAll = totalCommRepsPayees + totalCommRepsAPayer;
      const profitReelEncaisseHt = totalEncaisseHt - totalCommRepsAll - totalExpensesHt;
      const profitProjetePeriodHt = totalFacturableHt - totalCommRepsAll - totalExpensesHt;

      const tpsTvqPercue = Math.round(totalEncaisseHt * QC_TAX_RATE * 100) / 100;
      const tpsTvqPayeeDepenses = Math.round(totalExpensesHt * QC_TAX_RATE * 100) / 100;
      const tpsTvqNetARemettre = Math.max(0, Math.round((tpsTvqPercue - tpsTvqPayeeDepenses) * 100) / 100);

      // Rendu des 5 cartes KPI Finances
      const cardsEl = document.getElementById('fin-kpi-cards');
      if (cardsEl) {
        cardsEl.innerHTML = `
          <div class="bg-white p-4 rounded-2xl border border-emerald-200/90 shadow-xs space-y-1.5">
            <div class="text-[11px] font-bold text-emerald-700">1. Argent encaissé (${modeTag})</div>
            <div class="text-xl font-extrabold text-emerald-700 tabular-nums">${fmtMoney(totalEncaisseHt * taxMult)}</div>
            <div class="text-[10.5px] text-slate-500">Reçu des partenaires (${esc(periodLabel)})</div>
          </div>

          <div class="bg-white p-4 rounded-2xl border border-sky-200/90 shadow-xs space-y-1.5">
            <div class="text-[11px] font-bold text-sky-700">2. À recevoir partenaires (${modeTag})</div>
            <div class="text-xl font-extrabold text-sky-800 tabular-nums">${fmtMoney(totalARecevoirPartenairesHt * taxMult)}</div>
            <div class="text-[10.5px] text-slate-500">Factures en attente de règlement</div>
          </div>

          <div class="bg-white p-4 rounded-2xl border border-amber-200/90 shadow-xs space-y-1.5">
            <div class="text-[11px] font-bold text-amber-700">3. Comm. à payer aux reps</div>
            <div class="text-xl font-extrabold text-amber-700 tabular-nums">${fmtMoney(totalCommRepsAPayer)}</div>
            <div class="text-[10.5px] text-slate-500">Déjà versées : ${fmtMoney(totalCommRepsPayees)}</div>
          </div>

          <div class="bg-white p-4 rounded-2xl border border-rose-200/90 shadow-xs space-y-1.5">
            <div class="text-[11px] font-bold text-rose-700">4. Dépenses &amp; frais fixes (${modeTag})</div>
            <div class="text-xl font-extrabold text-rose-700 tabular-nums">${fmtMoney(totalExpensesHt * taxMult)}</div>
            <div class="text-[10.5px] text-slate-500">${periodExpenses.length} dépense(s) inscrite(s)</div>
          </div>

          <div class="bg-gradient-to-br from-slate-900 to-indigo-950 p-4 rounded-2xl text-white shadow-md space-y-1.5">
            <div class="text-[11px] font-bold text-indigo-200">5. Profit réel conservé (Net)</div>
            <div class="text-xl font-extrabold ${profitReelEncaisseHt >= 0 ? 'text-emerald-400' : 'text-rose-400'} tabular-nums">${fmtMoney(profitReelEncaisseHt)}</div>
            <div class="text-[10.5px] text-slate-300 tabular-nums">Après encaissement total : ${fmtMoney(profitProjetePeriodHt)}</div>
          </div>
        `;
      }

      const taxBanner = document.getElementById('fin-tax-cash-banner');
      if (taxBanner) {
        taxBanner.innerHTML = `
          <div class="flex flex-wrap items-center gap-4">
            <span>🧾 <strong>Provision TPS/TVQ (14,975 %) :</strong></span>
            <span class="text-slate-600 tabular-nums">Perçue sur encaissements : <strong>${fmtMoney(tpsTvqPercue)}</strong></span>
            <span class="text-slate-600 tabular-nums">CTI/RTI sur dépenses : <strong>− ${fmtMoney(tpsTvqPayeeDepenses)}</strong></span>
          </div>
          <div class="px-3 py-1 rounded-xl bg-amber-50 border border-amber-200 text-amber-900 font-extrabold tabular-nums">
            Solde TPS/TVQ à mettre de côté : ${fmtMoney(tpsTvqNetARemettre)}
          </div>
        `;
      }

      // 2. Tableau Commissions partenaires & Rapprochement Wave
      const batches = buildPartnerBatchesForMonths(months);
      const tbodyPartners = document.getElementById('tbody-fin-partners');
      if (tbodyPartners) {
        if (!batches.length) {
          tbodyPartners.innerHTML = `<tr><td colspan="5" class="py-6 text-center text-xs text-slate-400">Aucune vente confirmée sur cette période.</td></tr>`;
        } else {
          tbodyPartners.innerHTML = batches.map(b => {
            const isPaid = Boolean(b.paidInfo?.paid);
            const safeKey = b.batchKey.replace(/'/g, "'");
            return `
              <tr class="hover:bg-slate-50">
                <td class="py-3 px-3.5">
                  <div class="font-extrabold text-slate-900">${esc(b.entName)}</div>
                  <div class="text-[11px] text-slate-500">${esc(formatYearMonthLabelFr(b.ym))} • Volume : ${fmtMoney0(b.salesVolume)}</div>
                </td>
                <td class="py-3 px-2.5 text-center font-bold text-slate-700 tabular-nums">${b.count}</td>
                <td class="py-3 px-2.5 text-right font-extrabold text-slate-900 tabular-nums">${fmtMoney(b.grossHt)}</td>
                <td class="py-3 px-2.5 text-right font-semibold text-slate-500 tabular-nums">${fmtMoney(b.grossTtc)}</td>
                <td class="py-3 px-3.5 text-right">
                  <button type="button" onclick="togglePartnerBatchPaid('${safeKey}')" class="px-3 py-1.5 text-[11px] font-extrabold rounded-xl transition cursor-pointer ${isPaid ? 'bg-emerald-100 text-emerald-800 border border-emerald-300' : 'bg-amber-50 text-amber-900 border border-amber-300 hover:bg-amber-100'}">
                    ${isPaid ? `✓ Encaissé (${esc(b.paidInfo.paidAt || '')})` : '⏳ Marquer encaissé'}
                  </button>
                </td>
              </tr>
            `;
          }).join('');
        }
      }

      const waveBox = document.getElementById('fin-wave-reconcile-box');
      if (waveBox) {
        if (!waveReconcileResults) {
          waveBox.classList.add('hidden');
        } else {
          waveBox.classList.remove('hidden');
          waveBox.innerHTML = `
            <div class="flex items-center justify-between gap-2">
              <div class="text-xs font-extrabold text-emerald-950">✓ Fichier Wave analysé : ${esc(waveReconcileResults.filename)}</div>
              <button type="button" onclick="waveReconcileResults=null; renderAdminFinances();" class="text-xs font-bold text-slate-500 hover:text-slate-800">✕</button>
            </div>
            <div class="text-[11px] text-emerald-900 tabular-nums">
              ${waveReconcileResults.totalRows} transaction(s) lue(s) (${fmtMoney(waveReconcileResults.totalWaveAmt)}) • <strong>${waveReconcileResults.autoMatched} lot(s) partenaire(s) rapproché(s) automatiquement</strong>.
            </div>
          `;
        }
      }

      // 3. Tableau Commissions à payer aux représentants
      const tbodyPayouts = document.getElementById('tbody-fin-rep-payouts');
      if (tbodyPayouts) {
        const repRows = (state.representants || [])
          .filter(r => r.role !== 'admin')
          .map(rep => {
            const repSales = confirmedSales.filter(v => String(v.representant_id) === String(rep.id) && v.statut_paiement !== 'paye');
            let readyAmt = 0;
            let waitingPartnerAmt = 0;
            repSales.forEach(v => {
              const ym = getSaleStatsMonth(v);
              const entName = state.entreprises.find(e => String(e.id) === String(v.entreprise_id))?.nom || resolveSaleArchivedEntName(v) || '';
              const isPartnerPaid = Boolean(partnerPaidStore[getPartnerBatchKey(entName, ym)]?.paid);
              if (isPartnerPaid) readyAmt += Number(v.commission_rep || 0);
              else waitingPartnerAmt += Number(v.commission_rep || 0);
            });
            return { rep, readyAmt, waitingPartnerAmt };
          });

        const monthsEncoded = encodeURIComponent(JSON.stringify(months));
        tbodyPayouts.innerHTML = repRows.map(({ rep, readyAmt, waitingPartnerAmt }) => `
          <tr class="hover:bg-slate-50">
            <td class="py-3 px-3.5 font-extrabold text-slate-900">${esc(rep.nom)}</td>
            <td class="py-3 px-2.5 text-right font-extrabold text-emerald-700 tabular-nums">${fmtMoney(readyAmt)}</td>
            <td class="py-3 px-2.5 text-right font-semibold text-amber-700 tabular-nums">${fmtMoney(waitingPartnerAmt)}</td>
            <td class="py-3 px-3.5 text-right">
              ${readyAmt > 0
                ? `<button type="button" onclick="payAllReadyForRepInFinances('${esc(rep.id)}', '${monthsEncoded}')" class="px-2.5 py-1 text-[11px] font-extrabold text-white bg-emerald-600 hover:bg-emerald-700 rounded-lg transition cursor-pointer">Verser</button>`
                : `<span class="text-[11px] text-slate-400">—</span>`}
            </td>
          </tr>
        `).join('');
      }

      // 4. Tableau Dépenses & frais fixes
      const tbodyExp = document.getElementById('tbody-fin-expenses');
      if (tbodyExp) {
        if (!periodExpenses.length) {
          tbodyExp.innerHTML = `<tr><td colspan="8" class="py-6 text-center text-xs text-slate-400">Aucune dépense enregistrée sur cette période.</td></tr>`;
        } else {
          tbodyExp.innerHTML = periodExpenses.map(exp => {
            const ht = Number(exp.amountHt || 0);
            const tax = Math.round(ht * QC_TAX_RATE * 100) / 100;
            const ttc = Math.round((ht + tax) * 100) / 100;
            return `
              <tr class="hover:bg-slate-50">
                <td class="py-2.5 px-4 whitespace-nowrap text-slate-600 tabular-nums">${esc(exp.date)}</td>
                <td class="py-2.5 px-3 font-semibold text-slate-700">${esc(exp.category)}</td>
                <td class="py-2.5 px-3 font-bold text-slate-900">${esc(exp.label)}</td>
                <td class="py-2.5 px-3 text-center">
                  ${exp.recurring
                    ? '<span class="px-2 py-0.5 text-[10px] font-extrabold rounded-full bg-indigo-50 text-indigo-700 border border-indigo-200">Fixe mensuel</span>'
                    : '<span class="px-2 py-0.5 text-[10px] font-semibold rounded-full bg-slate-100 text-slate-600">Ponctuelle</span>'}
                </td>
                <td class="py-2.5 px-3 text-right font-bold text-slate-800 tabular-nums">${fmtMoney(ht)}</td>
                <td class="py-2.5 px-3 text-right text-slate-500 tabular-nums">${fmtMoney(tax)}</td>
                <td class="py-2.5 px-4 text-right font-extrabold text-slate-900 tabular-nums">${fmtMoney(ttc)}</td>
                <td class="py-2.5 px-3 text-right">
                  <button type="button" onclick="deleteFinExpense('${esc(exp.id)}')" class="px-2 py-1 text-[11px] font-bold text-rose-600 hover:bg-rose-50 rounded-lg cursor-pointer">✕</button>
                </td>
              </tr>
            `;
          }).join('');
        }
      }

      // 5. Prévisions financières mensuelles & annuelles (12 mois de l'année en cours)
      const curYear = String(today.getFullYear());
      const months12 = Array.from({ length: 12 }, (_, idx) => `${curYear}-${String(idx + 1).padStart(2, '0')}`);

      // Moyennes historiques pour projeter les mois futurs sans ventes
      const activeRepsCount = Math.max(1, getActiveKpiReps().length);
      const teamMonthlySalesObj = ALL_KPI_REPS.reduce((s, rn) => s + getSalesTargetForRep(rn, currentMonthStr), 0);
      const avgGrossRate = 0.10; // ~10 % marge brute globale Korpo (Rep + Olivier)
      const avgRepShareRatio = 0.58; // ~58 % reversé aux reps, ~42 % conservé par Korpo
      const recurringMonthlyFixed = Object.values(expensesStore)
        .filter(e => e && e.recurring)
        .reduce((acc, e) => {
          acc.set(`${e.category}__${String(e.label || '').trim().toLowerCase()}`, Number(e.amountHt || 0));
          return acc;
        }, new Map());
      const baseRecurringSum = Array.from(recurringMonthlyFixed.values()).reduce((s, v) => s + v, 0);

      let annualRev = 0;
      let annualRepComm = 0;
      let annualExp = 0;
      let annualNet = 0;

      const forecastRows = months12.map(ym => {
        const mSales = (state.ventes || []).filter(v => getSaleStatsMonth(v) === ym);
        const mExp = Object.values(expensesStore).filter(e => e && e.ym === ym).reduce((s, e) => s + Number(e.amountHt || 0), 0);
        const isPast = ym < currentMonthStr;
        const isCurrent = ym === currentMonthStr;

        let revHt = mSales.reduce((s, v) => s + computeKorpoGrossRevenueForSale(v), 0);
        let repComm = mSales.reduce((s, v) => s + Number(v.commission_rep || 0), 0);
        let expHt = mExp > 0 ? mExp : (!isPast ? baseRecurringSum : 0);
        let statusTag = '';

        if (isPast) {
          statusTag = '<span class="px-2 py-0.5 rounded-full bg-white/10 text-slate-300 text-[10px] font-bold">Réel</span>';
        } else if (isCurrent) {
          const dayRatio = Math.min(1, Math.max(0.15, today.getDate() / 30));
          const projectedRestRev = Math.max(0, (teamMonthlySalesObj * avgGrossRate) * (1 - dayRatio));
          revHt = Math.round((revHt + projectedRestRev) * 100) / 100;
          repComm = Math.round((repComm + projectedRestRev * avgRepShareRatio) * 100) / 100;
          statusTag = '<span class="px-2 py-0.5 rounded-full bg-amber-500/25 text-amber-300 text-[10px] font-extrabold">En cours + Proj.</span>';
        } else {
          // Mois futur : inclut le pipeline futur déjà vendu + projection à 85 % des objectifs
          const targetProjectedRev = Math.round(teamMonthlySalesObj * avgGrossRate * 0.85);
          if (revHt < targetProjectedRev) {
            const addRev = targetProjectedRev - revHt;
            revHt += addRev;
            repComm += Math.round(addRev * avgRepShareRatio);
          }
          statusTag = mSales.length > 0
            ? `<span class="px-2 py-0.5 rounded-full bg-indigo-500/30 text-indigo-200 text-[10px] font-extrabold">Pipeline (${mSales.length}) + Proj.</span>`
            : '<span class="px-2 py-0.5 rounded-full bg-white/10 text-slate-400 text-[10px] font-semibold">Prévisionnel</span>';
        }

        const netProfit = Math.round((revHt - repComm - expHt) * 100) / 100;
        annualRev += revHt;
        annualRepComm += repComm;
        annualExp += expHt;
        annualNet += netProfit;

        return { ym, label: formatYearMonthLabelFr(ym, true), fullLabel: formatYearMonthLabelFr(ym), statusTag, revHt, repComm, expHt, netProfit, isCurrent };
      });

      const annualPill = document.getElementById('fin-annual-summary-pill');
      if (annualPill) {
        annualPill.innerHTML = `
          <div class="text-[11px] uppercase tracking-wider text-indigo-300 font-semibold">Profit net annuel projeté (${curYear})</div>
          <div class="text-2xl sm:text-3xl font-extrabold ${annualNet >= 0 ? 'text-emerald-400' : 'text-rose-400'} tabular-nums">${fmtMoney0(annualNet)}</div>
          <div class="text-[11px] text-slate-300 tabular-nums">Revenus Korpo : ${fmtMoney0(annualRev)} • Charges : ${fmtMoney0(annualExp)}</div>
        `;
      }

      const maxNet = Math.max(1, ...forecastRows.map(r => Math.abs(r.netProfit)));
      const chartEl = document.getElementById('fin-forecast-chart');
      if (chartEl) {
        chartEl.innerHTML = forecastRows.map(r => {
          const hPct = Math.max(12, Math.round((Math.abs(r.netProfit) / maxNet) * 100));
          const barColor = r.netProfit >= 0
            ? (r.isCurrent ? 'bg-amber-400' : 'bg-emerald-400')
            : 'bg-rose-500';
          return `
            <div class="flex flex-col items-center gap-1.5">
              <div class="text-[10px] font-extrabold tabular-nums ${r.netProfit >= 0 ? 'text-emerald-300' : 'text-rose-300'}">${fmtMoney0(r.netProfit)}</div>
              <div class="w-full h-24 bg-white/5 rounded-xl flex items-end p-1 border border-white/10">
                <div class="w-full ${barColor} rounded-lg transition-all" style="height: ${hPct}%"></div>
              </div>
              <div class="text-[10px] font-bold ${r.isCurrent ? 'text-amber-300' : 'text-slate-300'}">${esc(r.label)}</div>
            </div>
          `;
        }).join('');
      }

      const tbodyForecast = document.getElementById('tbody-fin-forecast');
      if (tbodyForecast) {
        tbodyForecast.innerHTML = forecastRows.map(r => `
          <tr class="${r.isCurrent ? 'bg-white/10' : 'hover:bg-white/5'} transition">
            <td class="py-2.5 px-4 font-bold text-white">${esc(r.fullLabel)}</td>
            <td class="py-2.5 px-3">${r.statusTag}</td>
            <td class="py-2.5 px-3 text-right tabular-nums text-slate-200">${fmtMoney0(r.revHt)}</td>
            <td class="py-2.5 px-3 text-right tabular-nums text-amber-300">− ${fmtMoney0(r.repComm)}</td>
            <td class="py-2.5 px-3 text-right tabular-nums text-rose-300">− ${fmtMoney0(r.expHt)}</td>
            <td class="py-2.5 px-4 text-right tabular-nums font-extrabold ${r.netProfit >= 0 ? 'text-emerald-300' : 'text-rose-400'}">${fmtMoney0(r.netProfit)}</td>
          </tr>
        `).join('') + `
          <tr class="bg-indigo-950/90 font-extrabold border-t-2 border-indigo-400/40">
            <td colspan="2" class="py-3 px-4 text-white">TOTAL ANNUEL PROJETÉ (${curYear})</td>
            <td class="py-3 px-3 text-right tabular-nums text-white">${fmtMoney0(annualRev)}</td>
            <td class="py-3 px-3 text-right tabular-nums text-amber-300">− ${fmtMoney0(annualRepComm)}</td>
            <td class="py-3 px-3 text-right tabular-nums text-rose-300">− ${fmtMoney0(annualExp)}</td>
            <td class="py-3 px-4 text-right tabular-nums text-emerald-400 text-sm">${fmtMoney0(annualNet)}</td>
          </tr>
        `;
      }
    }


    initApp();

    document.getElementById('btn-reset-cfg')?.addEventListener('click', () => {
      localStorage.removeItem('korpo_sb_url');
      localStorage.removeItem('korpo_sb_key');
      showScreen('screen-config');
    });

    function restoreRememberedCredentials() {
      try {
        const raw = localStorage.getItem('korpo_saved_creds_v1');
        const emailEl = document.getElementById('login-email');
        const pwdEl = document.getElementById('login-password');
        const remEl = document.getElementById('login-remember');
        const forgetBtn = document.getElementById('btn-forget-saved-login');
        if (!raw || !emailEl || !pwdEl) return;
        const parsed = JSON.parse(raw);
        if (parsed.email) emailEl.value = parsed.email;
        if (parsed.pwd) { lsSet('korpo_saved_creds_v1', JSON.stringify({ email: parsed.email || '' })); } // purge d'un ancien mot de passe mémorisé
        if (remEl) remEl.checked = true;
        if (forgetBtn && parsed.email) forgetBtn.classList.remove('hidden');
      } catch (e) {}
    }

    document.getElementById('btn-toggle-pwd-vis')?.addEventListener('click', () => {
      const pwdEl = document.getElementById('login-password');
      if (!pwdEl) return;
      pwdEl.type = pwdEl.type === 'password' ? 'text' : 'password';
    });

    document.getElementById('btn-forget-saved-login')?.addEventListener('click', () => {
      localStorage.removeItem('korpo_saved_creds_v1');
      document.getElementById('login-email').value = '';
      document.getElementById('login-password').value = '';
      document.getElementById('login-remember').checked = false;
      document.getElementById('btn-forget-saved-login').classList.add('hidden');
      toast('Identifiants mémorisés effacés');
    });

    document.getElementById('login-form')?.addEventListener('submit', async (e) => {
      e.preventDefault();
      const errBox = document.getElementById('login-error');
      const submitBtn = document.getElementById('login-submit');
      errBox.classList.add('hidden');
      const email = document.getElementById('login-email').value.trim();
      const password = document.getElementById('login-password').value;
      const remember = document.getElementById('login-remember')?.checked;

      if (submitBtn) {
        submitBtn.disabled = true;
        submitBtn.textContent = 'Connexion en cours...';
      }

      try {
        const { data, error } = await sb.auth.signInWithPassword({ email, password });
        if (error) {
          errBox.textContent = `Échec de connexion : ${error.message}`;
          errBox.classList.remove('hidden');
          return;
        }

        if (remember) {
          try {
            lsSet('korpo_saved_creds_v1', JSON.stringify({ email })); // jamais le mot de passe
            document.getElementById('btn-forget-saved-login')?.classList.remove('hidden');
          } catch (err) {}
        } else {
          localStorage.removeItem('korpo_saved_creds_v1');
          document.getElementById('btn-forget-saved-login')?.classList.add('hidden');
        }

        if (data?.user) {
          showStartupSplash('Connexion réussie 👋', 'Préparation de votre espace Korpo...');
          await handleAuthenticatedUser(data.user);
        }
      } finally {
        if (submitBtn) {
          submitBtn.disabled = false;
          submitBtn.textContent = 'Accéder au portail';
        }
      }
    });

    async function performAppLogout() {
      document.getElementById('user-menu-dropdown')?.classList.add('hidden');
      stopAdminBackgroundSync();
      currentUser = null;
      currentProfile = null;
      authLoadingUserId = null;
      kvReset();

      // Purger les jetons de session et caches sensibles locaux immédiatement pour garantir la confidentialité sur appareil partagé
      try {
        const sensitiveKeys = new Set([
          'korpo_sale_audit_v1', 'korpo_sale_trash_v1', 'korpo_sale_bonus_overrides_v1',
          'korpo_carried_over_sales_v1', 'korpo_rep_targets_v1', 'korpo_company_targets_meta_v1',
          'korpo_client_targets_v1', 'korpo_kpi_condensed_v1', 'korpo_kpi_monthly_rdv_v1',
          'korpo_daily_rdv_v1', 'korpo_kpi_team_snapshot_v1', 'korpo_kpi_store_v1',
          'korpo_sale_proofs_v1', 'korpo_red_phone_cfg_v1', 'korpo_archived_reps_v1',
          'korpo_archived_entreprises_v1'
        ]);
        Object.keys(localStorage).forEach(k => {
          if ((k.startsWith('sb-') && k.endsWith('-auth-token')) || sensitiveKeys.has(k) || k.startsWith('korpo_kv_migrated_')) {
            localStorage.removeItem(k);
          }
        });
        window._initialDataSynced = false;
        state = { representants: [], entreprises: [], clients: [], taux: [], ventes: [], mesVentes: [], classement: [] };
        kpiStore = {};
        kpiHebdoRows = [];
        teamChatMessages = [];
        saleProofsStore = {};
      } catch (e) {}

      window._loggingOutNow = true;
      try {
        if (sb && sb.auth) {
          await Promise.race([
            sb.auth.signOut({ scope: 'local' }),
            new Promise(res => setTimeout(res, 250))
          ]);
        }
      } catch (e) {} finally {
        window._loggingOutNow = false;
      }

      restoreRememberedCredentials();
      showScreen('screen-login');
      toast('Vous êtes déconnecté');
    }

    document.getElementById('btn-logout')?.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      performAppLogout();
    });

    function setRefreshButtonSpinning(btnEl, isSpinning) {
      if (!btnEl) return;
      if (!document.getElementById('korpo-spin-keyframes')) {
        const st = document.createElement('style');
        st.id = 'korpo-spin-keyframes';
        st.textContent = '@keyframes korpoRefreshSpin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } } .korpo-spinning-icon { display: inline-block; animation: korpoRefreshSpin 0.75s linear infinite; transform-origin: center; }';
        document.head.appendChild(st);
      }
      const iconTarget = btnEl.querySelector('svg') || btnEl.querySelector('span') || btnEl;
      if (isSpinning) {
        btnEl.classList.add('opacity-80', 'cursor-wait');
        iconTarget.classList.add('korpo-spinning-icon', 'animate-spin');
      } else {
        btnEl.classList.remove('opacity-80', 'cursor-wait');
        iconTarget.classList.remove('korpo-spinning-icon', 'animate-spin');
      }
    }

    document.getElementById('btn-refresh')?.addEventListener('click', async () => {
      const btn = document.getElementById('btn-refresh');
      btn.disabled = true;
      setRefreshButtonSpinning(btn, true);
      const minSpinDelay = new Promise(res => setTimeout(res, 450));
      try {
        const [loaded] = await Promise.all([loadAllData(), minSpinDelay]);
        if (loaded !== false) toast('Données actualisées');
      } catch (err) {
        toast(err.message, 'error');
      } finally {
        setRefreshButtonSpinning(btn, false);
        btn.disabled = false;
      }
    });

    const today = new Date();
    const currentMonthStr = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}`;
    const currentDateStr = localYmd(today);
    if (document.getElementById('filter-rep-month')) document.getElementById('filter-rep-month').value = currentMonthStr;
    if (document.getElementById('filter-rep-start')) document.getElementById('filter-rep-start').value = `${currentMonthStr}-01`;
    if (document.getElementById('filter-rep-end')) document.getElementById('filter-rep-end').value = currentDateStr;
    document.getElementById('vente-date').value = currentDateStr;
    document.getElementById('vente-mois-comm').value = currentMonthStr;

    function getCurrentWeekBounds(refDate = new Date()) {
      const d = new Date(refDate.getFullYear(), refDate.getMonth(), refDate.getDate());
      const dayOfWeek = d.getDay(); // 0 = dimanche, 1 = lundi
      const diffToMonday = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
      const monday = new Date(d);
      monday.setDate(d.getDate() + diffToMonday);
      const sunday = new Date(monday);
      sunday.setDate(monday.getDate() + 6);
      return { start: localYmd(monday), end: localYmd(sunday) };
    }

    function getQuarterForMonth(ym) {
      if (!ym || ym.length < 7) return null;
      const y = ym.slice(0, 4);
      const m = parseInt(ym.slice(5, 7), 10);
      if (isNaN(m)) return null;
      return `${y}-Q${Math.ceil(m / 3)}`;
    }

    function readPeriodFilterState(prefix) {
      const mode = document.getElementById(`${prefix}-period`)?.value || 'all';
      const monthVal = document.getElementById(`${prefix}-month`)?.value || currentMonthStr;
      const dateVal = document.getElementById(`${prefix}-date`)?.value || currentDateStr;
      const startVal = document.getElementById(`${prefix}-start`)?.value || `${currentMonthStr}-01`;
      const endVal = document.getElementById(`${prefix}-end`)?.value || currentDateStr;
      const weekBounds = getCurrentWeekBounds(today);
      const currentQuarter = getQuarterForMonth(currentMonthStr);
      const currentYear = String(today.getFullYear());

      let label = 'Cumul global (Toutes les périodes)';
      let singleMonth = '';
      if (mode === 'this_week') label = `Cette semaine (${formatDateLabelFr(weekBounds.start)} → ${formatDateLabelFr(weekBounds.end)})`;
      else if (mode === 'this_month') { label = `Ce mois-ci (${formatYearMonthLabelFr(currentMonthStr)})`; singleMonth = currentMonthStr; }
      else if (mode === 'this_quarter') label = `Ce trimestre (${currentQuarter.replace('-Q', ' • T')})`;
      else if (mode === 'this_year') label = `Cette année (${currentYear})`;
      else if (mode === 'future') label = `Vue Futur (Dates à venir > ${formatDateLabelFr(currentDateStr)})`;
      else if (mode === 'specific_month') { label = formatYearMonthLabelFr(monthVal); singleMonth = monthVal; }
      else if (mode === 'specific_date') { label = formatDateLabelFr(dateVal); singleMonth = dateVal ? dateVal.slice(0, 7) : ''; }
      else if (mode === 'custom_range') {
        const s = startVal <= endVal ? startVal : endVal;
        const e = startVal <= endVal ? endVal : startVal;
        label = `Du ${formatDateLabelFr(s)} au ${formatDateLabelFr(e)}`;
        if (s && e && s.slice(0, 7) === e.slice(0, 7)) singleMonth = s.slice(0, 7);
      }

      return { mode, monthVal, dateVal, startVal, endVal, weekBounds, currentQuarter, currentYear, label, singleMonth };
    }

    function matchesSalePeriod(v, periodState, useStatsMonth = true) {
      if (!periodState || periodState.mode === 'all') return true;
      const { mode, monthVal, dateVal, weekBounds, currentQuarter, currentYear } = periodState;
      const dStr = (v.date_vente || '').trim().slice(0, 10);
      const saleDateYm = dStr.length >= 7 ? dStr.slice(0, 7) : '';
      const commYm = (v.mois_commission || '').trim().slice(0, 7);
      const futureYm = getSaleTargetFutureMonth(v);
      const ym = useStatsMonth
        ? (getSaleStatsMonth(v) || saleDateYm || commYm)
        : (commYm || saleDateYm);

      if (mode === 'this_week') {
        if (futureYm && futureYm > currentMonthStr) return false;
        return Boolean(dStr && dStr >= weekBounds.start && dStr <= weekBounds.end);
      }
      if (mode === 'this_month') {
        return ym === currentMonthStr;
      }
      if (mode === 'this_quarter') {
        return getQuarterForMonth(ym) === currentQuarter;
      }
      if (mode === 'this_year') {
        return (ym || '').startsWith(currentYear);
      }
      if (mode === 'future') {
        return Boolean(getSaleTargetFutureMonth(v));
      }
      if (mode === 'specific_month') {
        if (!monthVal) return true;
        if (monthVal > currentMonthStr) {
          return getSaleTargetFutureMonth(v) === monthVal || commYm === monthVal || saleDateYm === monthVal;
        }
        return ym === monthVal;
      }
      if (mode === 'specific_date') {
        return !dateVal || dStr === dateVal;
      }
      if (mode === 'custom_range') {
        const rawS = periodState.startVal || '';
        const rawE = periodState.endVal || '';
        const s = (rawS && rawE && rawS > rawE) ? rawE : rawS;
        const e = (rawS && rawE && rawS > rawE) ? rawS : rawE;
        if (s && dStr && dStr < s) return false;
        if (e && dStr && dStr > e) return false;
        return Boolean(dStr);
      }
      return true;
    }

    function syncPeriodInputsVisibility(prefix, autoOpenPicker = false) {
      const mode = document.getElementById(`${prefix}-period`)?.value || 'all';
      const monthInp = document.getElementById(`${prefix}-month`);
      const dateInp = document.getElementById(`${prefix}-date`);
      const customBox = document.getElementById(`${prefix}-custom-box`);
      const startInp = document.getElementById(`${prefix}-start`);
      const endInp = document.getElementById(`${prefix}-end`);
      if (monthInp) {
        const showM = mode === 'specific_month';
        monthInp.classList.toggle('hidden', !showM);
        if (showM && !monthInp.value) monthInp.value = currentMonthStr;
        if (showM && autoOpenPicker && typeof monthInp.showPicker === 'function') {
          try { monthInp.showPicker(); } catch (e) {}
        }
      }
      if (dateInp) {
        const showD = mode === 'specific_date';
        dateInp.classList.toggle('hidden', !showD);
        if (showD && !dateInp.value) dateInp.value = currentDateStr;
        if (showD && autoOpenPicker && typeof dateInp.showPicker === 'function') {
          try { dateInp.showPicker(); } catch (e) {}
        }
      }
      if (customBox) {
        const showC = mode === 'custom_range';
        customBox.classList.toggle('hidden', !showC);
        if (showC) {
          if (startInp && !startInp.value) startInp.value = `${currentMonthStr}-01`;
          if (endInp && !endInp.value) endInp.value = currentDateStr;
        }
      }
    }

    function syncAllMonthPickers() {
      ['filter-rep', 'filter-rank', 'admin-filter', 'dash-filter'].forEach(p => syncPeriodInputsVisibility(p, false));
      const kpiMode = document.getElementById('kpi-period-mode')?.value || 'this_month';
      document.getElementById('kpi-custom-range-box')?.classList.toggle('hidden', kpiMode !== 'custom_range');
    }

    function bindPeriodSelectorGroup(prefix, renderFn) {
      const sel = document.getElementById(`${prefix}-period`);
      const mInp = document.getElementById(`${prefix}-month`);
      const dInp = document.getElementById(`${prefix}-date`);
      const sInp = document.getElementById(`${prefix}-start`);
      const eInp = document.getElementById(`${prefix}-end`);
      sel?.addEventListener('change', () => {
        syncPeriodInputsVisibility(prefix, true);
        renderFn();
      });
      mInp?.addEventListener('change', renderFn);
      mInp?.addEventListener('input', renderFn);
      dInp?.addEventListener('change', renderFn);
      dInp?.addEventListener('input', renderFn);
      sInp?.addEventListener('change', renderFn);
      sInp?.addEventListener('input', renderFn);
      eInp?.addEventListener('change', renderFn);
      eInp?.addEventListener('input', renderFn);
    }

    bindPeriodSelectorGroup('filter-rep', renderMesVentes);
    bindPeriodSelectorGroup('filter-rank', renderClassement);
    bindPeriodSelectorGroup('admin-filter', renderAdminVentes);
    bindPeriodSelectorGroup('dash-filter', renderAdminDashboard);

    document.getElementById('filter-rep-statut')?.addEventListener('change', renderMesVentes);
    // =========================================================================
    // RECHERCHE GLOBALE CLIENT DEPUIS L'EN-TÊTE (LOUPE)
    // =========================================================================
    window.openGlobalClientSearch = () => {
      const modal = document.getElementById('global-client-search-modal');
      const input = document.getElementById('global-client-search-input');
      if (!modal || !input) return;
      modal.classList.remove('hidden');
      input.value = '';
      renderGlobalClientSearchResults('');
      setTimeout(() => input.focus(), 60);
    };

    window.closeGlobalClientSearch = () => {
      const modal = document.getElementById('global-client-search-modal');
      if (modal) modal.classList.add('hidden');
    };

    function getAllSearchableSalesForCurrentUser() {
      const isAdmin = currentProfile?.role === 'admin';
      const source = isAdmin ? (state.ventes || []) : (state.mesVentes || []);
      return source.map(v => {
        const cli = v.client_nom || state.clients.find(c => String(c.id) === String(v.client_id))?.nom || '—';
        const ent = v.entreprise_nom || state.entreprises.find(e => String(e.id) === String(v.entreprise_id))?.nom || '—';
        const rep = v.rep_nom || state.representants.find(r => String(r.id) === String(v.representant_id))?.nom || '—';
        const commVal = isAdmin ? Number(v.commission_olivier || 0) : Number(v.commission_rep || 0);
        return {
          ...v,
          _cli: cli,
          _ent: ent,
          _rep: rep,
          _comm: commVal
        };
      }).sort((a, b) => String(b.date_vente || '').localeCompare(String(a.date_vente || '')));
    }

    function renderGlobalClientSearchResults(rawQuery) {
      const box = document.getElementById('global-client-search-results');
      const countEl = document.getElementById('global-client-search-count');
      if (!box) return;

      const q = String(rawQuery || '').trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
      const all = getAllSearchableSalesForCurrentUser();

      const matches = !q
        ? all.slice(0, 12)
        : all.filter(v => {
            const hay = `${v._cli} ${v._ent} ${v._rep} ${v.note || ''} ${v.date_vente || ''}`
              .toLowerCase()
              .normalize('NFD')
              .replace(/[̀-ͯ]/g, '');
            return hay.includes(q);
          }).slice(0, 30);

      if (countEl) {
        countEl.textContent = q
          ? `${matches.length} dossier(s) trouvé(s)`
          : `${all.length} dossier(s) au total`;
      }

      if (!matches.length) {
        box.innerHTML = `<div class="p-8 text-center text-xs text-slate-400">Aucun client ou dossier ne correspond à « ${esc(rawQuery)} ».</div>`;
        return;
      }

      const isAdmin = currentProfile?.role === 'admin';
      box.innerHTML = matches.map(v => {
        const isSuivi = (v.note || '').toUpperCase().includes('[SUIVI]');
        const cleanNote = cleanDisplaySaleNote(v.note);
        return `
          <button type="button" data-nav-sale-id="${v.id}" onclick="navigateToSaleRecord('${v.id}')" style="touch-action: manipulation; -webkit-tap-highlight-color: transparent;" class="w-full text-left p-3.5 sm:px-5 hover:bg-indigo-50/70 active:bg-indigo-100/70 transition flex items-center justify-between gap-3 cursor-pointer">
            <div class="min-w-0 space-y-1">
              <div class="flex items-center flex-wrap gap-1.5">
                ${entBadge(v._ent)}
                <span class="font-extrabold text-sm text-slate-900 truncate">${esc(v._cli)}</span>
              </div>
              <div class="flex items-center flex-wrap gap-1.5 text-[11px] text-slate-500">
                <span class="font-semibold text-slate-700">${v.date_vente || '—'}</span>
                ${isAdmin ? `<span>•</span><span class="font-semibold text-indigo-600">${esc(v._rep)}</span>` : ''}
                <span>•</span>
                <span class="px-1.5 py-0.2 text-[10px] font-extrabold rounded border ${isSuivi ? 'bg-purple-50 text-purple-700 border-purple-200' : 'bg-sky-50 text-sky-700 border-sky-200'}">${isSuivi ? 'SUIVI' : 'OTS'}</span>
                <span class="px-1.5 py-0.2 text-[10px] font-bold rounded-full ${v.statut_depot === 'depot_recu' ? 'bg-emerald-50 text-emerald-700' : 'bg-amber-50 text-amber-700'}">${v.statut_depot === 'depot_recu' ? (isNoDepositSale(v) ? 'Sans dépôt requis' : 'Dépôt reçu') : 'À confirmer'}</span>
                ${cleanNote ? `<span class="text-slate-400 truncate max-w-[160px]">• ${esc(cleanNote)}</span>` : ''}
              </div>
            </div>
            <div class="text-right shrink-0 tabular-nums">
              <div class="text-sm font-extrabold text-slate-900">${fmtMoney(v.valeur)}</div>
              <div class="text-[11px] font-bold text-emerald-600">+ ${fmtMoney(v._comm)}</div>
              <div class="text-[10px] font-bold text-indigo-600 mt-0.5">Ouvrir la fiche →</div>
            </div>
          </button>
        `;
      }).join('');
    }

    window._globalClientFilterQuery = '';
    function showActiveClientSearchBanner(cliNom) {
      let bar = document.getElementById('active-global-client-filter-bar');
      if (!cliNom) {
        if (bar) bar.remove();
        return;
      }
      if (!bar) {
        bar = document.createElement('div');
        bar.id = 'active-global-client-filter-bar';
        bar.className = 'fixed bottom-20 sm:bottom-6 left-1/2 -translate-x-1/2 z-40 px-4 py-2 rounded-full bg-slate-900 text-white text-xs font-bold shadow-2xl border border-indigo-400/50 flex items-center gap-2.5 no-print';
        document.body.appendChild(bar);
      }
      bar.innerHTML = `<span>🔍 Client : <strong class="text-emerald-300">${esc(cliNom)}</strong></span><button type="button" onclick="clearGlobalClientFilter()" class="px-2 py-0.5 rounded-full bg-white/15 hover:bg-white/25 text-[11px] font-extrabold transition">Tout afficher ✕</button>`;
    }

    window.clearGlobalClientFilter = () => {
      window._globalClientFilterQuery = '';
      showActiveClientSearchBanner('');
      renderMesVentes();
      if (currentProfile?.role === 'admin') renderAdminVentes();
    };

    window.navigateToSaleRecord = (saleId) => {
      const sid = String(saleId);
      closeGlobalClientSearch();
      if (document.activeElement && typeof document.activeElement.blur === 'function') {
        document.activeElement.blur();
      }

      const isAdmin = currentProfile?.role === 'admin';
      const v = (isAdmin ? state.ventes : state.mesVentes).find(x => String(x.id) === sid)
        || (state.mesVentes || []).find(x => String(x.id) === sid)
        || (state.ventes || []).find(x => String(x.id) === sid);
      if (!v) return;

      const cliNom = v.client_nom || state.clients.find(c => String(c.id) === String(v.client_id))?.nom || '';

      window._globalClientFilterQuery = cliNom;
      showActiveClientSearchBanner(cliNom);

      if (isAdmin) {
        switchTab('tab-admin-ventes');
        const rIn = document.getElementById('admin-filter-rep');
        const eIn = document.getElementById('admin-filter-ent');
        const pIn = document.getElementById('admin-filter-period');
        const stIn = document.getElementById('admin-filter-statut');
        if (rIn) rIn.value = '';
        if (eIn) eIn.value = '';
        if (pIn) pIn.value = 'all';
        if (stIn) stIn.value = '';
        if (typeof syncPeriodInputsVisibility === 'function') syncPeriodInputsVisibility('admin-filter', false);
        renderAdminVentes();
      } else {
        switchTab('tab-mes-ventes');
        const stIn = document.getElementById('filter-rep-statut');
        const pIn = document.getElementById('filter-rep-period');
        if (stIn) stIn.value = '';
        if (pIn) pIn.value = 'all';
        if (typeof syncPeriodInputsVisibility === 'function') syncPeriodInputsVisibility('filter-rep', false);
        renderMesVentes();
      }

      // Défiler jusqu'à la fiche exacte du client et l'illuminer (Desktop & Mobile)
      const scrollToExactCard = () => {
        const isMobile = window.innerWidth < 768;
        const containerSel = isAdmin
          ? (isMobile ? '#mobile-admin-ventes-cards' : '#tbody-admin-ventes')
          : (isMobile ? '#mobile-mes-ventes-cards' : '#tbody-mes-ventes');
        const container = document.querySelector(containerSel);
        const exactEl = container?.querySelector(`[data-sale-id="${sid}"]`) || container?.firstElementChild;
        if (!exactEl) return;

        const headerEl = document.querySelector('#screen-app header');
        const headerOffset = (headerEl ? headerEl.offsetHeight : 112) + 16;
        const rect = exactEl.getBoundingClientRect();
        const targetY = window.scrollY + rect.top - headerOffset;
        try {
          window.scrollTo({ top: Math.max(0, targetY), behavior: 'smooth' });
        } catch (e) {
          window.scrollTo(0, Math.max(0, targetY));
        }

        exactEl.classList.add('ring-2', 'ring-emerald-500', 'bg-emerald-50/60');
        setTimeout(() => {
          exactEl.classList.remove('ring-2', 'ring-emerald-500', 'bg-emerald-50/60');
        }, 2600);
      };

      setTimeout(scrollToExactCard, 150);
      setTimeout(scrollToExactCard, 500);
    };

    document.getElementById('global-client-search-input')?.addEventListener('input', (e) => {
      renderGlobalClientSearchResults(e.target.value);
    });
    document.getElementById('global-client-search-input')?.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.keyCode === 13) {
        e.preventDefault();
        const firstBtn = document.querySelector('#global-client-search-results [data-nav-sale-id]');
        if (firstBtn) {
          navigateToSaleRecord(firstBtn.getAttribute('data-nav-sale-id'));
        }
      }
    });
    document.getElementById('global-client-search-results')?.addEventListener('pointerdown', (e) => {
      const btn = e.target.closest('[data-nav-sale-id]');
      if (btn && e.pointerType === 'touch') {
        window._touchNavSaleId = btn.getAttribute('data-nav-sale-id');
        window._touchNavStartY = e.clientY;
      }
    });
    document.getElementById('global-client-search-results')?.addEventListener('pointerup', (e) => {
      const btn = e.target.closest('[data-nav-sale-id]');
      if (btn && e.pointerType === 'touch' && window._touchNavSaleId === btn.getAttribute('data-nav-sale-id')) {
        if (Math.abs(e.clientY - (window._touchNavStartY || 0)) < 12) {
          e.preventDefault();
          navigateToSaleRecord(window._touchNavSaleId);
        }
        window._touchNavSaleId = null;
      }
    });
    document.getElementById('global-client-search-modal')?.addEventListener('click', (e) => {
      if (e.target.id === 'global-client-search-modal') closeGlobalClientSearch();
    });
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        const gModal = document.getElementById('global-client-search-modal');
        if (gModal && !gModal.classList.contains('hidden')) closeGlobalClientSearch();
      }
    });

    document.getElementById('filter-admin-scope')?.addEventListener('change', renderMesVentes);
    document.getElementById('dash-future-month-filter')?.addEventListener('change', renderFutureMonthsBreakdown);

    document.getElementById('vente-entreprise')?.addEventListener('change', updateClientDatalist);
    document.getElementById('taux-ent')?.addEventListener('change', updateTauxClientSelect);

    document.getElementById('form-new-vente')?.addEventListener('submit', async (e) => {
      e.preventDefault();
      try {
        const editId = document.getElementById('vente-edit-id').value;
        const repId = currentProfile.role === 'admin' ? document.getElementById('vente-rep').value : currentProfile.id;
        const entId = document.getElementById('vente-entreprise').value;
        const clientNom = document.getElementById('vente-client-nom').value;
        {
          const entRule = (state.entreprises || []).find(x => String(x.id) === String(entId));
          const depotSelEl = document.getElementById('vente-statut-depot');
          if (entRule && entRule.depot_requis === false) depotSelEl.value = 'sans_depot';
          else if (entRule && depotSelEl.value === 'sans_depot') throw new Error('Cette entreprise exige un dépôt : choisissez « Dépôt reçu » ou « À confirmer ».');
        }
        if (!entId) throw new Error('Choisissez l\'entreprise partenaire.');
        if (!clientNom.trim()) throw new Error('Il manque le nom du client.');
        if (!document.getElementById('vente-date').value) throw new Error('Il manque la date réelle de la vente.');
        if (!document.getElementById('vente-mois-comm').value) throw new Error('Il manque le mois de commission.');
        const clientId = await getOrCreateClient(clientNom, entId);
        const valeur = parseFloat(document.getElementById('vente-valeur').value);
        if (!Number.isFinite(valeur) || valeur <= 0) throw new Error('Montant de vente invalide.');
        {
          const ymGuard = document.getElementById('vente-mois-comm')?.value;
          const oldGuard = editId ? findSaleById(editId) : null;
          if (isMonthClosed(ymGuard) || (oldGuard && isMonthClosed(oldGuard.mois_commission))) throw new Error(LOCK_MSG);
        }

        const payload = {
          representant_id: repId,
          entreprise_id: entId,
          client_id: clientId,
          valeur,
          date_vente: document.getElementById('vente-date').value,
          mois_commission: document.getElementById('vente-mois-comm').value,
          statut_depot: document.getElementById('vente-statut-depot').value === 'sans_depot' ? 'depot_recu' : document.getElementById('vente-statut-depot').value,
          statut_paiement: document.getElementById('vente-statut-paiement').value,
          note: (() => {
            const rawNote = document.getElementById('vente-note').value
              .trim()
              .replace(/^\[(OTS|SUIVI)\]\s*/i, '').replace(/\[(REP|ENT):[^\]]+\]\s*/ig, '').replace(/\[SANSDEPOT\]\s*/ig, '')
              .replace(/\[HIT(?::(?:REP|OLIVIER|BOTH))?\]\s*/ig, '')
              .trim();
            const typeKpi = document.getElementById('vente-type-kpi').value;
            const isHit = Boolean(document.getElementById('vente-is-hit')?.checked);
            const hitBearer = (document.getElementById('vente-hit-bearer')?.value || 'both').toUpperCase();
            const hitPrefix = isHit ? `[HIT:${hitBearer}] ` : '';
            const noDepTag = document.getElementById('vente-statut-depot').value === 'sans_depot' ? '[SANSDEPOT] ' : '';
            return rawNote ? `[${typeKpi}] ${hitPrefix}${noDepTag}${rawNote}` : `[${typeKpi}]${isHit ? ` [HIT:${hitBearer}]` : ''}${noDepTag ? ' [SANSDEPOT]' : ''}`;
          })()
        };

        let insertedRow = null; // portée de toute la soumission (réconciliation des doublons plus bas)
        let proofRefSaved;      // valeur de preuve_url écrite : chemin Storage, null (supprimée) ou undefined (inchangée)
        if (editId) {
          const eff = computeDrawerCommissions(valeur);
          const { tRep, tOli, cRep, cOli, bonusActive, rateMode, isHit, hitBearer, hitAmount, hitAmountRep, hitAmountOli } = eff;
          payload.taux_rep_applique = tRep;
          payload.taux_olivier_applique = tOli;
          payload.commission_rep = cRep;
          payload.commission_olivier = cOli;
          proofRefSaved = await resolveProofForSave(window._editingOriginalHadProof);
          if (proofRefSaved !== undefined) payload.preuve_url = proofRefSaved;

          let { error } = await sb.from('ventes').update(payload).eq('id', editId);
          if (error && 'preuve_url' in payload && (error.message || '').includes('preuve_url')) {
            delete payload.preuve_url;
            const retryUp = await sb.from('ventes').update(payload).eq('id', editId);
            error = retryUp.error;
          }
          if (error) throw error;

          if (proofRefSaved === null) {
            delete saleProofsStore[String(editId)];
            const oldV = state.ventes.find(x => String(x.id) === String(editId));
            if (oldV) {
              const oldCli = oldV.client_nom || state.clients.find(c => String(c.id) === String(oldV.client_id))?.nom || '';
              delete saleProofsStore[`${oldV.date_vente}_${oldCli}_${oldV.valeur}`];
            }
            saveSaleProofsStore();
            await recordSaleAudit(editId, 'Preuve photo supprimée', 'Retirée lors de la modification du dossier', '🗑️');
          }

          if (isHit || bonusActive || rateMode === 'no_bonus' || rateMode === 'custom') {
            await sb.from('ventes').update({
              taux_rep_applique: tRep,
              taux_olivier_applique: tOli,
              commission_rep: cRep,
              commission_olivier: cOli
            }).eq('id', editId);
            await saveSaleBonusOverride(editId, {
              tRep,
              tOli,
              cRep,
              cOli,
              isHit,
              hitBearer,
              hitAmount,
              hitAmountRep,
              hitAmountOli,
              noBonus: rateMode === 'no_bonus',
              customManual: rateMode === 'custom'
            });
          } else {
            await saveSaleBonusOverride(editId, null);
          }

          const hitAuditTag = isHit
            ? ` • 🚨 HIT (${hitBearer === 'rep' ? 'Rep' : (hitBearer === 'olivier' ? 'Olivier' : 'Rep + Oli')}) : Rep ${fmtMoney(cRep)} / Oli ${fmtMoney(cOli)}`
            : '';
          const rateAuditTag = rateMode === 'no_bonus'
            ? ` • Bypass bonus (${fmtPct(tRep)} / ${fmtPct(tOli)})`
            : (rateMode === 'custom'
              ? ` • Taux manuel (${fmtPct(tRep)} / ${fmtPct(tOli)})`
              : (bonusActive ? ` • Bonus Target +${formatBonusPctShort(eff.bonusDelta)} Rep / -${formatBonusPctShort(eff.bonusDelta)} Olivier` : ''));
          await recordSaleAudit(editId, isHit ? 'Dossier modifié (HIT)' : 'Dossier modifié', `Montant : ${fmtMoney(valeur)}${hitAuditTag}${rateAuditTag} • Dépôt : ${payload.statut_depot === 'depot_recu' ? 'Reçu' : 'À confirmer'} • Paiement : ${payload.statut_paiement === 'paye' ? 'Payé' : 'Non payé'}`, isHit ? '🚨' : '✏️');
          toast(isHit ? 'Dossier HIT mis à jour 🚨' : 'Vente mise à jour');
        } else {
          const eff = computeDrawerCommissions(valeur);
          const { tRep, tOli, cRep, cOli, bonusActive, rateMode, isHit, hitBearer, hitAmount, hitAmountRep, hitAmountOli } = eff;
          payload.taux_rep_applique = tRep;
          payload.taux_olivier_applique = tOli;
          payload.commission_rep = cRep;
          payload.commission_olivier = cOli;

          window._justCreatedSaleClientSide = true;
          proofRefSaved = await resolveProofForSave(false);
          if (proofRefSaved) payload.preuve_url = proofRefSaved;
          let error;
          ({ data: insertedRow, error } = await sb.from('ventes').insert(payload).select().single());
          if (error && pendingProofDataUrl && (error.message || '').includes('preuve_url')) {
            delete payload.preuve_url;
            const retryRes = await sb.from('ventes').insert(payload).select().single();
            insertedRow = retryRes.data;
            error = retryRes.error;
          }
          if (error) throw error;

          if (insertedRow?.id && (isHit || bonusActive || rateMode === 'no_bonus' || rateMode === 'custom')) {
            await sb.from('ventes').update({
              taux_rep_applique: tRep,
              taux_olivier_applique: tOli,
              commission_rep: cRep,
              commission_olivier: cOli
            }).eq('id', insertedRow.id);
            await saveSaleBonusOverride(insertedRow.id, {
              tRep,
              tOli,
              cRep,
              cOli,
              isHit,
              hitBearer,
              hitAmount,
              hitAmountRep,
              hitAmountOli,
              noBonus: rateMode === 'no_bonus',
              customManual: rateMode === 'custom'
            });
          }

          toast('Vente enregistrée avec succès');
          const repObjCel = state.representants.find(r => String(r.id) === String(repId));
          const entObjCel = state.entreprises.find(en => String(en.id) === String(entId));
          const shownComm = (currentProfile?.role === 'admin' && String(repId) !== String(currentProfile?.id))
            ? payload.commission_olivier
            : payload.commission_rep;
          if (!isHit) {
            triggerSaleCelebration(
              repObjCel?.nom || currentProfile?.nom || '',
              clientNom.trim(),
              entObjCel?.nom || '',
              valeur,
              shownComm
            );
          }
          if (insertedRow?.id) {
            const entN = entObjCel?.nom || '';
            const bonusNote = eff.bonusActive ? ` • 🔥 Bonus Target +${formatBonusPctShort(eff.bonusDelta)} Rep / -${formatBonusPctShort(eff.bonusDelta)} Olivier` : '';
            await recordSaleAudit(insertedRow.id, 'Vente créée', `${entN} — ${clientNom.trim()} • ${fmtMoney(valeur)}${bonusNote}${pendingProofDataUrl ? ' (avec preuve photo)' : ''}`, '💵');
            const typeKpiVal = document.getElementById('vente-type-kpi')?.value || 'OTS';
            await publishAutoSaleToTeamChat(
              insertedRow,
              repObjCel?.nom || currentProfile?.nom || 'Représentant',
              entN || 'Entreprise',
              clientNom.trim(),
              typeKpiVal
            );
          }
        }

        if (editId) {
          const newStatsYm = (payload.date_vente && payload.date_vente.length >= 7) ? payload.date_vente.slice(0, 7) : payload.mois_commission;
          if (carriedOverSalesStore[String(editId)] && newStatsYm) {
            carriedOverSalesStore[String(editId)].origYm = newStatsYm;
            lsSet('korpo_carried_over_sales_v1', JSON.stringify(carriedOverSalesStore));
          }
        }

        const savedSaleYm = (payload.date_vente && payload.date_vente.length >= 7) ? payload.date_vente.slice(0, 7) : payload.mois_commission;
        if (savedSaleYm && isValidYearMonth(savedSaleYm)) {
          ensureKpiMonthStructure(savedSaleYm);
          populateKpiMonthSelect(false);
          const kpiMonthSel = document.getElementById('kpi-select-month');
          if (kpiMonthSel) kpiMonthSel.value = savedSaleYm;
        }
        if (payload.statut_depot === 'depot_recu') {
          const targetId = editId || (insertedRow?.id ? insertedRow.id : null);
          if (targetId) {
            await loadAllData();
            const merged = await reconcilePreviousUnconfirmedDuplicateSales([targetId]);
            if (merged > 0) {
              toast(`Doublon antérieur non confirmé retiré automatiquement (${merged}) ✓`);
              await loadAllData();
            }
            drawer.classList.add('hidden');
            return;
          }
        }
        drawer.classList.add('hidden');
        await loadAllData();
      } catch (err) {
        toast(err.message, 'error');
      }
    });

    ['admin-filter-rep', 'admin-filter-ent', 'admin-filter-statut'].forEach(id => {
      document.getElementById(id)?.addEventListener('change', renderAdminVentes);
    });
    document.getElementById('btn-admin-clear-filters')?.addEventListener('click', () => {
      window._globalClientFilterQuery = '';
      showActiveClientSearchBanner('');
      document.getElementById('admin-filter-rep').value = '';
      document.getElementById('admin-filter-ent').value = '';
      document.getElementById('admin-filter-period').value = 'all';
      if (document.getElementById('admin-filter-month')) document.getElementById('admin-filter-month').value = '';
      if (document.getElementById('admin-filter-start')) document.getElementById('admin-filter-start').value = '';
      if (document.getElementById('admin-filter-end')) document.getElementById('admin-filter-end').value = '';
      document.getElementById('admin-filter-custom-box')?.classList.add('hidden');
      document.getElementById('admin-filter-statut').value = '';
      syncAllMonthPickers();
      renderAdminVentes();
    });
    syncAllMonthPickers();
    document.getElementById('btn-export-payroll-batch')?.addEventListener('click', () => {
      const dashPeriod = readPeriodFilterState('dash-filter');
      const m = dashPeriod.singleMonth;
      const list = state.ventes.filter(v => matchesSalePeriod(v, dashPeriod, true));
      const headers = ['Representant', 'Courriel', 'Dossiers_Total', 'Dossiers_Depot_Recu', 'Volume_Ventes_CAD', 'Commission_Totale_CAD', 'Deja_Paye_CAD', 'Solde_A_Verser_CAD'];
      const rows = state.representants.map(r => {
        const isOli = r.role === 'admin';
        const sub = list.filter(v => isOli ? true : v.representant_id === r.id);
        const totV = sub.reduce((s, v) => s + Number(v.valeur || 0), 0);
        const totC = sub.reduce((s, v) => s + Number(isOli ? v.commission_olivier : v.commission_rep), 0);
        const paidC = sub.filter(v => v.statut_paiement === 'paye').reduce((s, v) => s + Number(isOli ? v.commission_olivier : v.commission_rep), 0);
        const depotOk = sub.filter(v => v.statut_depot === 'depot_recu').length;
        return [csvCell(r.nom), csvCell(r.courriel), sub.length, depotOk, totV.toFixed(2), totC.toFixed(2), paidC.toFixed(2), (totC - paidC).toFixed(2)].join(',');
      });
      const blob = new Blob(['\ufeff' + [headers.join(','), ...rows].join('\n')], { type: 'text/csv;charset=utf-8;' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = `korpo-paie-equipe-${m || 'cumul'}.csv`;
      a.click();
      toast('Fichier de paie consolidé exporté ✓');
    });

    document.getElementById('btn-delete-all-ventes')?.addEventListener('click', async () => {
      if (!state.ventes.length) return toast('Aucune vente à supprimer.', 'error');
      const ok = await askConfirm(
        `Mettre les ${state.ventes.length} ventes à la corbeille ?`,
        'Toutes les ventes seront archivées dans la corbeille pendant 30 jours afin que vous puissiez les restaurer si nécessaire.'
      );
      if (!ok) return;

      // On relit TOUTES les ventes de la base (pas seulement celles affichées) avant de supprimer
      let allVentes;
      try {
        allVentes = await fetchAllRows(() => sb.from('ventes').select('*').order('id'));
      } catch (err) {
        return toast(`Lecture des ventes impossible : ${err.message}`, 'error');
      }
      if (!allVentes.length) return toast('Aucune vente à supprimer.', 'error');

      const trashOk = await moveSalesToTrash(allVentes);
      if (!trashOk) return toast('La sauvegarde cloud de la corbeille a échoué : suppression annulée, aucune vente supprimée.', 'error');

      const ids = allVentes.map(v => v.id);
      for (let i = 0; i < ids.length; i += 200) {
        const { error } = await sb.from('ventes').delete().in('id', ids.slice(i, i + 200));
        if (error) {
          toast(error.message, 'error');
          await loadAllData();
          return;
        }
      }

      toast('Toutes les ventes ont été déplacées dans la corbeille 🗑️');
      await loadAllData();
    });

    document.getElementById('form-taux')?.addEventListener('submit', async (e) => {
      e.preventDefault();
      const repId = document.getElementById('taux-rep').value;
      const entId = document.getElementById('taux-ent').value;
      const clientId = document.getElementById('taux-client').value || null;
      const tauxRep = parseFloat(document.getElementById('taux-val-rep').value) / 100;
      const tauxOlivier = parseFloat(document.getElementById('taux-val-olivier').value) / 100;

      if (!isFinite(tauxRep) || !isFinite(tauxOlivier) || tauxRep < 0 || tauxOlivier < 0 || tauxRep > 1 || tauxOlivier > 1) {
        return toast('Taux invalide : entrez des pourcentages entre 0 et 100.', 'error');
      }

      // Sauvegarde de l'ancien taux : restauré si l'insertion échoue (avant : le taux était perdu)
      let selQuery = sb.from('taux_commission').select('*').eq('representant_id', repId).eq('entreprise_id', entId);
      selQuery = clientId ? selQuery.eq('client_id', clientId) : selQuery.is('client_id', null);
      const { data: backupRows } = await selQuery;

      let delQuery = sb.from('taux_commission').delete().eq('representant_id', repId).eq('entreprise_id', entId);
      delQuery = clientId ? delQuery.eq('client_id', clientId) : delQuery.is('client_id', null);
      const delRes = await delQuery;
      if (delRes.error) return toast(delRes.error.message, 'error');

      const { error } = await sb.from('taux_commission').insert({
        representant_id: repId,
        entreprise_id: entId,
        client_id: clientId,
        taux: tauxRep,
        taux_olivier: tauxOlivier
      });

      if (error) {
        if (backupRows && backupRows.length) await sb.from('taux_commission').insert(backupRows);
        return toast(`${error.message} (ancien taux conservé)`, 'error');
      }
      toast('Taux enregistré');
      await loadAllData();
    });

    document.getElementById('form-new-entreprise')?.addEventListener('submit', async (e) => {
      e.preventDefault();
      const nom = document.getElementById('new-entreprise-nom').value.trim();
      const depotChoice = document.getElementById('new-entreprise-depot')?.value || '';
      if (!depotChoice) return toast('Indiquez si cette entreprise exige un dépôt.', 'error');
      let { error } = await sb.from('entreprises').insert({ nom, depot_requis: depotChoice === 'oui' });
      let columnMissing = false;
      if (error && /depot_requis/i.test(error.message || '')) {
        columnMissing = true;
        ({ error } = await sb.from('entreprises').insert({ nom }));
      }
      if (error) return toast(error.message, 'error');
      document.getElementById('new-entreprise-nom').value = '';
      const depotSelNew = document.getElementById('new-entreprise-depot');
      if (depotSelNew) depotSelNew.value = '';
      toast(columnMissing
        ? '⚠️ Entreprise ajoutée, mais la colonne « depot_requis » est absente de Supabase : exécutez le SQL fourni pour enregistrer cette règle.'
        : 'Entreprise ajoutée', columnMissing ? 'error' : undefined);
      await loadAllData();
    });

    window.openRatesForRep = (repId) => {
      switchTab('tab-admin-taux');
      const tauxRepSel = document.getElementById('taux-rep');
      if (tauxRepSel && repId) {
        tauxRepSel.value = String(repId);
        if (typeof updateTauxClientSelect === 'function') updateTauxClientSelect();
      }
      document.getElementById('taux-val-rep')?.focus();
    };

    document.getElementById('form-new-rep')?.addEventListener('submit', async (e) => {
      e.preventDefault();
      const payload = {
        nom: document.getElementById('new-rep-nom').value.trim(),
        courriel: document.getElementById('new-rep-email').value.trim().toLowerCase(),
        role: document.getElementById('new-rep-role').value
      };
      const { data: insertedRep, error } = await sb.from('representants').insert(payload).select().single();
      if (error) return toast(error.message, 'error');
      document.getElementById('new-rep-nom').value = '';
      document.getElementById('new-rep-email').value = '';
      toast(`${payload.nom} ajouté à l'équipe ✓`);
      await loadAllData();

      const setupBanner = document.getElementById('new-rep-setup-banner');
      if (setupBanner) {
        document.getElementById('new-rep-setup-title').textContent = `✅ ${payload.nom} ajouté ! 2 étapes pour activer son profil :`;
        document.getElementById('new-rep-setup-email').textContent = payload.courriel;
        const gotoBtn = document.getElementById('btn-new-rep-goto-rates');
        if (gotoBtn) {
          const targetId = insertedRep?.id || state.representants.find(r => r.courriel === payload.courriel)?.id || '';
          gotoBtn.onclick = () => openRatesForRep(targetId);
        }
        setupBanner.classList.remove('hidden');
      }
    });

    document.getElementById('btn-csv-export')?.addEventListener('click', () => {
      const headers = ['date_vente', 'mois_commission', 'representant', 'entreprise', 'client', 'valeur', 'taux_rep', 'commission_rep', 'taux_olivier', 'commission_olivier', 'statut_depot', 'statut_paiement', 'note'];
      const rows = state.ventes.map(v => [
        csvCell(v.date_vente),
        csvCell(v.mois_commission),
        csvCell(state.representants.find(r => r.id === v.representant_id)?.nom || ''),
        csvCell(state.entreprises.find(e => e.id === v.entreprise_id)?.nom || ''),
        csvCell(state.clients.find(c => c.id === v.client_id)?.nom || ''),
        csvCell(v.valeur, false),
        csvCell(v.taux_rep_applique, false),
        csvCell(v.commission_rep, false),
        csvCell(v.taux_olivier_applique, false),
        csvCell(v.commission_olivier, false),
        csvCell(v.statut_depot),
        csvCell(v.statut_paiement),
        csvCell(v.note)
      ].join(','));

      const blob = new Blob(['\ufeff' + [headers.join(','), ...rows].join('\r\n')], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `korpo-ventes-${localYmd()}.csv`;
      a.click();
      setTimeout(() => URL.revokeObjectURL(url), 2000);
    });

    // Lecteur CSV complet : guillemets, "" échappés, séparateurs et retours à la ligne dans les champs
    function parseCsvText(text, sep) {
      const rows = [];
      let row = [], cur = '', inQ = false;
      for (let i = 0; i < text.length; i++) {
        const ch = text[i];
        if (inQ) {
          if (ch === '"') {
            if (text[i + 1] === '"') { cur += '"'; i++; } else { inQ = false; }
          } else { cur += ch; }
        } else if (ch === '"') {
          inQ = true;
        } else if (ch === sep) {
          row.push(cur); cur = '';
        } else if (ch === '\n') {
          row.push(cur); rows.push(row); row = []; cur = '';
        } else if (ch !== '\r') {
          cur += ch;
        }
      }
      if (cur !== '' || row.length) { row.push(cur); rows.push(row); }
      return rows;
    }

    function parseCsvLine(line, sep) {
      return (parseCsvText(line, sep)[0] || ['']).map(c => c.trim());
    }

    // Repère les colonnes par leur en-tête (compatible avec le fichier produit par « Exporter CSV »)
    function csvHeaderIndex(headerCols) {
      const norm = s => String(s || '').trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, '');
      const idx = {};
      headerCols.forEach((h, i) => { if (idx[norm(h)] === undefined) idx[norm(h)] = i; });
      const pick = (...names) => { for (const n of names) if (idx[n] !== undefined) return idx[n]; return -1; };
      return {
        date: pick('date_vente', 'date'),
        mois: pick('mois_commission', 'mois'),
        rep: pick('representant', 'rep'),
        ent: pick('entreprise'),
        cli: pick('client'),
        val: pick('valeur', 'montant', 'valeur_vente'),
        note: pick('note', 'notes'),
        type: pick('type', 'kpi', 'type_kpi')
      };
    }

    document.getElementById('csv-import-input')?.addEventListener('change', async (e) => {
      const input = e.target;
      const file = input.files[0];
      if (!file) return;
      let imported = 0, skipped = 0;
      try {
        const text = (await file.text()).replace(/^\uFEFF/, '');
        const firstLine = text.split(/\r?\n/, 1)[0] || '';
        const sep = (firstLine.match(/;/g) || []).length > (firstLine.match(/,/g) || []).length ? ';' : ',';
        const table = parseCsvText(text, sep).filter(r => r.some(c => String(c).trim()));
        if (table.length < 2) return toast('Fichier CSV vide', 'error');

        const H = csvHeaderIndex(table[0]);
        const headerMode = H.rep >= 0 && H.ent >= 0 && H.cli >= 0 && H.val >= 0;
        // Ancien format sans en-têtes reconnus : date, mois, représentant, entreprise, client, valeur, note, type
        const col = headerMode ? H : { date: 0, mois: 1, rep: 2, ent: 3, cli: 4, val: 5, note: 6, type: 7 };

        for (let i = 1; i < table.length; i++) {
          const cols = table[i].map(c => String(c).trim());
          if (!headerMode && cols.length < 6) { skipped++; continue; }
          const get = k => (col[k] >= 0 ? (cols[col[k]] || '') : '');

          const rep = state.representants.find(r => r.nom.toLowerCase() === get('rep').toLowerCase());
          const ent = state.entreprises.find(en => en.nom.toLowerCase().replace(/\s+/g, '') === get('ent').toLowerCase().replace(/\s+/g, ''));
          const cliNom = get('cli');
          if (!rep || !ent || !cliNom) { skipped++; continue; }

          let dateVente = get('date');
          const fr = dateVente.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/);
          if (fr) dateVente = `${fr[3]}-${fr[2].padStart(2, '0')}-${fr[1].padStart(2, '0')}`;
          if (!/^\d{4}-\d{2}-\d{2}$/.test(dateVente)) { skipped++; continue; }
          let moisComm = get('mois');
          if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(moisComm)) moisComm = dateVente.slice(0, 7);

          const cleanVal = parseAmount(get('val'));
          if (!isFinite(cleanVal)) { skipped++; continue; }

          const note = get('note');
          const rowJoined = (headerMode ? `${note} ${get('type')}` : cols.slice(6).join(' ')).toUpperCase();
          const kpiTag = rowJoined.includes('SUIVI') ? 'SUIVI' : 'OTS';
          const cleanNoteBody = note.replace(/^\[(OTS|SUIVI)\]\s*/i, '').replace(/\[(REP|ENT):[^\]]+\]\s*/ig, '').replace(/\[SANSDEPOT\]\s*/ig, '').trim();
          const finalNote = cleanNoteBody ? `[${kpiTag}] ${cleanNoteBody}` : `[${kpiTag}]`;

          try {
            const clientId = await getOrCreateClient(cliNom, ent.id);
            const { error } = await sb.from('ventes').insert({
              representant_id: rep.id,
              entreprise_id: ent.id,
              client_id: clientId,
              valeur: cleanVal,
              date_vente: dateVente,
              mois_commission: moisComm,
              note: finalNote
            });
            if (error) skipped++; else imported++;
          } catch (rowErr) {
            skipped++;
          }
        }
        toast(skipped > 0 ? `${imported} vente(s) importée(s), ${skipped} ignorée(s)` : `${imported} vente(s) importée(s)`);
      } catch (err) {
        toast(`Import impossible : ${err.message}`, 'error');
      } finally {
        input.value = '';
        await loadAllData();
      }
    });

    window.addEventListener('error', (ev) => {
      if (document.getElementById('screen-app')?.classList.contains('hidden') &&
          document.getElementById('screen-login')?.classList.contains('hidden')) {
        document.getElementById('screen-login')?.classList.remove('hidden');
      }
      console.error('Erreur application :', ev.error || ev.message);
    });

    initApp().catch((err) => {
      console.error('Erreur initApp :', err);
      showScreen('screen-login');
      const errBox = document.getElementById('login-error');
      if (errBox) {
        errBox.textContent = 'Erreur de démarrage : ' + (err?.message || err);
        errBox.classList.remove('hidden');
      }
    });
  
    // =========================================================================
    // CLÔTURE DE MOIS : un mois clôturé est verrouillé (ventes, taux, commissions, dépôts).
    // Seuls le statut de paiement et la preuve restent modifiables. Le verrou est appliqué par la base
    // (voir migration_cloture_validation.sql) ; ce code affiche des messages clairs et gère l'écran admin.
    // =========================================================================
    const closedMonths = new Set();
    const LOCK_MSG = '🔒 Ce mois est clôturé : rouvrez-le (Clôture des mois) pour modifier ce dossier.';
    const isMonthClosed = (ym) => Boolean(ym) && closedMonths.has(String(ym).slice(0, 7));
    function findSaleById(id) {
      return (state.ventes || []).find(x => String(x.id) === String(id)) || (state.mesVentes || []).find(x => String(x.id) === String(id));
    }
    const isSaleLocked = (id) => { const v = findSaleById(id); return Boolean(v) && isMonthClosed(v.mois_commission); };
    async function refreshClosedMonths() {
      try {
        const { data, error } = await sb.from('mois_clotures').select('ym');
        if (error) throw error;
        closedMonths.clear(); (data || []).forEach(r => closedMonths.add(r.ym));
      } catch (e) { /* table absente : fonction inactive tant que la migration SQL n'est pas exécutée */ }
    }
    const fmtYmLong = (ym) => { try { return new Date(ym + '-01T12:00:00').toLocaleDateString('fr-CA', { month: 'long', year: 'numeric' }); } catch (e) { return ym; } };
    function monthClosureSummary(ym) {
      const list = (state.ventes || []).filter(v => v.mois_commission === ym);
      const comm = list.reduce((s, v) => s + Number(v.commission_rep || 0) + Number(v.commission_olivier || 0), 0);
      return { n: list.length, toConfirm: list.filter(v => v.statut_depot !== 'depot_recu').length, unpaid: list.filter(v => v.statut_paiement !== 'paye').length, comm };
    }
    async function renderMonthClosureModal() {
      let m = document.getElementById('month-closure-modal');
      if (!m) {
        m = document.createElement('div'); m.id = 'month-closure-modal';
        m.className = 'fixed inset-0 z-50 bg-slate-950/60 backdrop-blur-sm flex items-center justify-center p-4 no-print';
        m.addEventListener('click', (e) => { if (e.target === m) m.remove(); });
        document.body.appendChild(m);
      }
      const known = new Set([currentMonthStr, ...closedMonths, ...(state.ventes || []).map(v => v.mois_commission).filter(Boolean)]);
      const openMonths = Array.from(known).filter(ym => !closedMonths.has(ym)).sort().reverse();
      const closed = Array.from(closedMonths).sort().reverse();
      m.innerHTML = `
        <div class="bg-white rounded-3xl border border-slate-200 shadow-2xl w-full max-w-md p-5 space-y-4 text-slate-900">
          <div class="flex items-center justify-between"><h3 class="text-base font-extrabold">🔒 Clôture des mois</h3>
            <button type="button" id="mc-close" class="w-8 h-8 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-600">✕</button></div>
          <p class="text-xs text-slate-500">Un mois clôturé ne peut plus changer : ventes, taux, commissions et dépôts sont verrouillés. Seuls le paiement et la preuve restent modifiables.</p>
          <div class="flex gap-2">
            <select id="mc-month" class="flex-1 px-3 py-2 text-xs font-bold border border-slate-200 rounded-xl bg-white">${openMonths.map(ym => `<option value="${esc(ym)}">${esc(fmtYmLong(ym))}</option>`).join('')}</select>
            <button type="button" id="mc-do-close" class="px-3 py-2 text-xs font-extrabold text-white bg-indigo-600 hover:bg-indigo-700 rounded-xl">Clôturer</button>
          </div>
          <div class="space-y-1.5"><div class="text-[11px] font-extrabold uppercase tracking-wider text-slate-400">Mois clôturés</div>
            ${closed.length ? closed.map(ym => `<div class="flex items-center justify-between px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs"><span class="font-bold">🔒 ${esc(fmtYmLong(ym))}</span><button type="button" data-reopen="${esc(ym)}" class="px-2.5 py-1 font-bold text-rose-700 bg-rose-50 hover:bg-rose-100 border border-rose-200 rounded-lg">Rouvrir</button></div>`).join('') : '<div class="text-xs text-slate-400">Aucun mois clôturé.</div>'}
          </div>
        </div>`;
      m.querySelector('#mc-close').onclick = () => m.remove();
      m.querySelector('#mc-do-close').onclick = async () => {
        const ym = m.querySelector('#mc-month').value; if (!ym) return;
        const s = monthClosureSummary(ym);
        const warn = (s.toConfirm ? `\n⚠️ ${s.toConfirm} dépôt(s) pas encore confirmé(s).` : '') + (s.unpaid ? `\n${s.unpaid} commission(s) non payée(s) (le paiement restera modifiable).` : '');
        if (!(await askConfirm(`Clôturer ${fmtYmLong(ym)} ?`, `${s.n} vente(s), commissions totales ${fmtMoney(s.comm)}.${warn}\nLe mois sera verrouillé.`))) return;
        const { error } = await sb.from('mois_clotures').insert({ ym, closed_by_nom: currentProfile?.nom || null });
        if (error) return toast(error.message, 'error');
        closedMonths.add(ym); toast(`Mois ${fmtYmLong(ym)} clôturé 🔒`); renderMonthClosureModal();
      };
      m.querySelectorAll('[data-reopen]').forEach(btn => { btn.onclick = async () => {
        const ym = btn.getAttribute('data-reopen');
        if (!(await askConfirm(`Rouvrir ${fmtYmLong(ym)} ?`, 'Les ventes de ce mois redeviendront modifiables.'))) return;
        const { error } = await sb.from('mois_clotures').delete().eq('ym', ym);
        if (error) return toast(error.message, 'error');
        closedMonths.delete(ym); toast(`Mois ${fmtYmLong(ym)} rouvert`); renderMonthClosureModal();
      }; });
    }
    window.openMonthClosureModal = async () => {
      if (currentProfile?.role !== 'admin') return;
      await refreshClosedMonths(); renderMonthClosureModal();
    };
    // Messages clairs avant d'atteindre le verrou de la base
    (function guardClosedMonths() {
      const wrap = (name, blocked) => {
        const orig = window[name]; if (typeof orig !== 'function') return;
        window[name] = function (...a) { if (blocked(...a)) { toast(LOCK_MSG, 'error'); return; } return orig.apply(this, a); };
      };
      wrap('openEditVenteDrawer', (id, dup) => !dup && isSaleLocked(id));
      wrap('deleteVente', (id) => isSaleLocked(id));
      wrap('toggleVenteStatut', (id, field) => field === 'statut_depot' && isSaleLocked(id));
      wrap('toggleVenteKpiType', (id) => isSaleLocked(id));
      wrap('openRateAdjustModal', (id) => isSaleLocked(id));
      wrap('quickConfirmSale', (id) => isSaleLocked(id));
    })();

    // =========================================================================
    // MOT DE PASSE OUBLIÉ + GLOSSAIRE
    // =========================================================================
    document.getElementById('btn-forgot-pwd')?.addEventListener('click', async () => {
      const info = document.getElementById('login-info');
      const err = document.getElementById('login-error');
      info?.classList.add('hidden'); err?.classList.add('hidden');
      const email = (document.getElementById('login-email')?.value || '').trim();
      if (!email) { if (err) { err.textContent = 'Entrez d\'abord votre courriel ci-dessus, puis cliquez sur « Mot de passe oublié ».'; err.classList.remove('hidden'); } return; }
      try {
        const { error } = await sb.auth.resetPasswordForEmail(email, { redirectTo: window.location.origin + window.location.pathname });
        if (error) throw error;
        if (info) { info.textContent = 'Si ce courriel est enregistré, un lien de réinitialisation vient d\'être envoyé. Vérifiez aussi vos courriels indésirables.'; info.classList.remove('hidden'); }
      } catch (e) {
        if (err) { err.textContent = 'Envoi impossible : ' + (e?.message || e); err.classList.remove('hidden'); }
      }
    });

    const GLOSSARY_ITEMS = [
      ['OTS', 'Vente signée dès le 1er rendez-vous avec le client.'],
      ['Suivi', 'Vente signée après une relance (un 2e contact ou plus).'],
      ['Date réelle de la vente', 'Le jour où le contrat a été signé.'],
      ['Mois de commission', 'Le mois où la commission est comptée et payée (souvent le mois d\'installation).'],
      ['Dépôt client', '« À confirmer » : l\'acompte n\'est pas encore validé. « Dépôt reçu » : confirmé par l\'administration. Joindre la photo du dépôt Interac accélère la confirmation.'],
      ['Statut du paiement', 'Indique si votre commission vous a été versée. Géré par l\'administration.'],
      ['Total vendu / Total confirmé', 'Vendu = tous vos contrats signés. Confirmé = seulement ceux dont le dépôt est reçu.'],
      ['HIT', 'Remboursement ou reprise de commission (par exemple si un contrat est annulé).'],
      ['Bonus Target', 'Pourcentage ajouté au taux de base de la grille de commissions.'],
      ['KPI', 'Indicateurs de performance : vos ventes OTS / Suivi et la progression vers vos objectifs.'],
      ['Korpo Top Earner 🏆', 'Le canal de discussion de l\'équipe. Écrivez @ suivi d\'un nom pour mentionner quelqu\'un : il sera notifié.']
    ];
    window.openGlossaryModal = () => {
      let m = document.getElementById('glossary-modal');
      if (!m) {
        m = document.createElement('div'); m.id = 'glossary-modal';
        m.className = 'fixed inset-0 z-[60] bg-slate-950/60 backdrop-blur-sm flex items-center justify-center p-4 no-print';
        m.addEventListener('click', (e) => { if (e.target === m) m.classList.add('hidden'); });
        document.body.appendChild(m);
      }
      m.classList.remove('hidden');
      m.innerHTML = `
        <div class="bg-white rounded-3xl border border-slate-200 shadow-2xl w-full max-w-md max-h-[85vh] flex flex-col text-slate-900">
          <div class="flex items-center justify-between px-5 pt-5 pb-3"><h3 class="text-base font-extrabold">📖 Glossaire</h3>
            <button type="button" id="glossary-close" class="w-8 h-8 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-600">✕</button></div>
          <div class="px-5 pb-5 overflow-y-auto space-y-3">
            ${GLOSSARY_ITEMS.map(([t, d]) => `<div><div class="text-xs font-extrabold text-indigo-700">${esc(t)}</div><div class="text-xs text-slate-600 leading-relaxed">${esc(d)}</div></div>`).join('')}
          </div>
        </div>`;
      m.querySelector('#glossary-close').onclick = () => m.classList.add('hidden');
    };

