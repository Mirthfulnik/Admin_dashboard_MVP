/* Lotus Music · VK Digital-отчёт (MVP)
   - Без географии
   - Детализация = список объявлений
   - Стримы = колонка vk
   - Blocked/статусы не фильтруем
*/
(function(){
  const $ = (s, root=document) => root.querySelector(s);
  const $$ = (s, root=document) => Array.from(root.querySelectorAll(s));
  const CLOUD_API = "https://functions.yandexcloud.net/d4eflrgebtkcansjrlkj";
  const StoreKey = "lm_vk_report_mvp_v1";

  /** -----------------------------
   *  State / Storage
   * -----------------------------*/
  const state = {
    db: loadDB_(),
    currentReleaseId: null,
    temp: {
      streamsParsed: null,
      streamsPreview: null
    }
  };

/** -----------------------------
   *  Cloud (Yandex Object Storage via Cloud Function)
   *  - single endpoint: POST {action:...}
   * -----------------------------*/
  async function cloudCall_(payload){
    const res = await fetch(CLOUD_API, {
      method: "POST",
      headers: {"Content-Type":"application/json"},
      body: JSON.stringify(payload || {})
    });
    const txt = await res.text();
    let data = {};
    try { data = JSON.parse(txt || "{}"); } catch(e){}
    if (!res.ok || data.ok === false){
      const msg = (data && (data.error || data.errorMessage)) ? (data.error || data.errorMessage) : `Cloud API error (${res.status})`;
      throw new Error(msg);
    }
    return data;
  }

  async function cloudPresignPut_({ key, contentType }){
    const data = await cloudCall_({ action: "presignPut", key, contentType });
    // keep backward-friendly shape: { url }
    return { url: data.uploadUrl };
  }

  async function cloudPresignGet_({ key }){
    const data = await cloudCall_({ action: "presignGet", key });
    return { url: data.downloadUrl };
  }

  async function cloudPut_({ url, blob, contentType }){
    const res = await fetch(url, {
      method: "PUT",
      headers: { "Content-Type": contentType },
      body: blob
    });
    if (!res.ok) throw new Error(`upload failed (${res.status})`);
  }

  async function cloudSaveJson_({ key, obj }){
    const blob = new Blob([JSON.stringify(obj)], {type:"application/json"});
    const { url } = await cloudPresignPut_({ key, contentType:"application/json" });
    await cloudPut_({ url, blob, contentType:"application/json" });
  }

  // ===== Cloud index (to show releases/versions in any browser) =====
  const CLOUD_INDEX_KEY = "index/releases.json";
  const CLOUD_RELEASE_INDEX_PREFIX = "index/releases/"; // + {releaseId}.json

  const _cloudUrlCache = new Map(); // key -> { url, t } (presigned URL cache)
  const CLOUD_URL_TTL_MS = 240 * 1000; // refresh before 300s expiry

  async function cloudObjectUrl_(key){
    if (!key) return "";
    const cached = _cloudUrlCache.get(key);
    if (cached && cached.url && (Date.now() - cached.t) < CLOUD_URL_TTL_MS) return cached.url;

    const { url } = await cloudPresignGet_({ key });
    _cloudUrlCache.set(key, { url, t: Date.now() });
    return url;
  }

  async function fetchJsonOrNull_(url){
    const res = await fetch(url, { cache: "no-store" });
    if (res.status === 404) return null;
    if (!res.ok) throw new Error(`GET failed (${res.status})`);
    const txt = await res.text();
    try { return JSON.parse(txt || "{}"); } catch(e){ return null; }
  }

  async function cloudReadJsonKeyOrNull_(key){
    try{
      const { url } = await cloudPresignGet_({ key });
      return await fetchJsonOrNull_(url);
    }catch(e){
      // if file doesn't exist yet, treat as empty
      return null;
    }
  }

  async function cloudWriteJsonKey_(key, obj){
    await cloudSaveJson_({ key, obj });
  }

  function cloudMakeLocalStubFromIndex_(meta){
    return {
      releaseId: meta.releaseId,
      artists: meta.artists || "",
      track: meta.track || "",
      title: meta.title || releaseTitle_(meta.artists||"", meta.track||""),
      createdAt: meta.createdAt || meta.updatedAt || nowIso_(),
      updatedAt: meta.updatedAt || nowIso_(),
      coverDataUrl: null,
      coverKey: meta.coverKey || null,
      streams: null,
      communities: [],
      cloudOnly: true,
      lastCloudVersion: meta.lastTs || null,
      cloudVersions: meta.lastTs && meta.lastDataKey ? [{ ts: meta.lastTs, dataKey: meta.lastDataKey }] : []
    };
  }

  async function cloudSyncFromIndex_(){
    const idx = await cloudReadJsonKeyOrNull_(CLOUD_INDEX_KEY);
    const releases = Array.isArray(idx && idx.releases) ? idx.releases : [];
    let changed = false;

    for (const meta of releases){
      if (!meta || !meta.releaseId) continue;
      if (!state.db.releases[meta.releaseId]){
        state.db.releases[meta.releaseId] = cloudMakeLocalStubFromIndex_(meta);
        changed = true;
      }else{
        // update basic meta if cloud has newer
        const local = state.db.releases[meta.releaseId];
        if ((meta.updatedAt||"") > (local.updatedAt||"")){
          local.artists = meta.artists || local.artists;
          local.track = meta.track || local.track;
          local.title = meta.title || local.title;
          local.updatedAt = meta.updatedAt || local.updatedAt;
          local.coverKey = meta.coverKey || local.coverKey;
          local.lastCloudVersion = meta.lastTs || local.lastCloudVersion;
          if (meta.lastTs && meta.lastDataKey){
            local.cloudVersions = Array.isArray(local.cloudVersions) ? local.cloudVersions : [];
            // ensure latest at front
            const exists = local.cloudVersions.find(v=>v.ts===meta.lastTs);
            if (!exists) local.cloudVersions.unshift({ ts: meta.lastTs, dataKey: meta.lastDataKey });
          }
          changed = true;
        }
      }
    }

    if (changed) saveDB_();
    syncSelectors_();
    renderReleases_();
    renderHistory_();
    // report/editor will be rendered on navigation
  }

  async function cloudUpsertIndexOnSave_(release, { ts, dataKey, prefix }){
    const now = nowIso_();

    // 1) Global index
    const idx = await cloudReadJsonKeyOrNull_(CLOUD_INDEX_KEY) || { schema: 1, updatedAt: now, releases: [] };
    idx.schema = idx.schema || 1;
    idx.updatedAt = now;
    idx.releases = Array.isArray(idx.releases) ? idx.releases : [];

    const meta = {
      releaseId: release.releaseId,
      title: release.title,
      artists: release.artists,
      track: release.track,
      createdAt: release.createdAt || now,
      updatedAt: release.updatedAt || now,
      coverKey: release.coverKey || null,
      lastTs: ts,
      lastDataKey: dataKey
    };

    const i = idx.releases.findIndex(x => x.releaseId === release.releaseId);
    if (i >= 0) idx.releases[i] = meta; else idx.releases.push(meta);

    // keep newest first
    idx.releases.sort((a,b)=> (b.updatedAt||"").localeCompare(a.updatedAt||""));

    await cloudWriteJsonKey_(CLOUD_INDEX_KEY, idx);

    // 2) Per-release versions file
    const relIndexKey = CLOUD_RELEASE_INDEX_PREFIX + release.releaseId + ".json";
    const relIdx = await cloudReadJsonKeyOrNull_(relIndexKey) || { schema: 1, releaseId: release.releaseId, title: release.title, versions: [] };
    relIdx.schema = relIdx.schema || 1;
    relIdx.title = release.title;
    relIdx.versions = Array.isArray(relIdx.versions) ? relIdx.versions : [];
    // add new version at front
    relIdx.versions = relIdx.versions.filter(v => v && v.ts !== ts);
    relIdx.versions.unshift({ ts, dataKey, prefix, createdAt: now });
    await cloudWriteJsonKey_(relIndexKey, relIdx);
  }

  async function ensureReleaseHydrated_(releaseId){
    const r = state.db.releases[releaseId];
    if (!r) return null;

    // already hydrated or fully local
    if (r._cloudHydrated || (!r.cloudOnly && (r.communities && r.communities.length || r.streams))) {
      // still ensure image urls exist when only keys are present
      if (r.coverKey && (!r.coverDataUrl || String(r.coverDataUrl).includes("X-Amz-Algorithm="))) r.coverDataUrl = await cloudObjectUrl_(r.coverKey);
      for (const c of (r.communities || [])){
        for (const cr of (c.creatives || [])){
          if (cr.objectKey && (!cr.dataUrl || String(cr.dataUrl).includes("X-Amz-Algorithm="))) cr.dataUrl = await cloudObjectUrl_(cr.objectKey);
        }
      }
      return r;
    }

    const dataKey = (Array.isArray(r.cloudVersions) && r.cloudVersions[0] && r.cloudVersions[0].dataKey)
      ? r.cloudVersions[0].dataKey
      : null;

    if (!dataKey) return r; // nothing to hydrate from

    const { url } = await cloudPresignGet_({ key: dataKey });
    const snap = await fetchJsonOrNull_(url);
    if (!snap) throw new Error("Не удалось загрузить snapshot из облака");

    // Apply snapshot
    r.title = snap.title || r.title;
    r.artists = snap.artists || r.artists;
    r.track = snap.track || r.track;
    r.createdAt = snap.createdAt || r.createdAt;
    r.updatedAt = snap.updatedAt || r.updatedAt;
    r.coverKey = snap.coverKey || r.coverKey || null;
    r.streams = Array.isArray(snap.streams) ? snap.streams : null;
    r.communities = Array.isArray(snap.communities) ? snap.communities : [];

    // Backward compatibility: older snapshots could store compact keys (impr/clicks/etc).
    for (const c of (r.communities || [])){
      if (Array.isArray(c.adsRows)){
        c.adsRows = c.adsRows.map(row0 => {
          const row = row0 || {};
          if (row && (row.impr != null || row.clicks != null || row.spent != null || row.adds != null)){
            // upgrade compact shape -> expected VK export headers
            if (row.impr != null && row["Показы"] == null) row["Показы"] = row.impr;
            if (row.clicks != null && row["Клики"] == null) row["Клики"] = row.clicks;
            if (row.spent != null && row["Потрачено всего, ₽"] == null) row["Потрачено всего, ₽"] = row.spent;
            if (row.adds != null && row["Добавления"] == null) row["Добавления"] = row.adds;
            if (row.listens != null && row["Прослушивания"] == null) row["Прослушивания"] = row.listens;
            if (row.groupId != null && row["ID группы"] == null) row["ID группы"] = row.groupId;
            if (row.name != null && row["Название объявления"] == null) row["Название объявления"] = row.name;
          }
          return normalizeRowKeys_(row);
        });
      }
      if (Array.isArray(c.demoFiles)){
        for (const df of c.demoFiles){
          if (!df || !Array.isArray(df.rows)) continue;
          df.rows = df.rows.map(row0 => {
            const row = row0 || {};
            if (row && (row.impr != null || row.clicks != null || row.age != null || row.sex != null)){
              if (row.impr != null && row["Показы"] == null) row["Показы"] = row.impr;
              if (row.clicks != null && row["Клики"] == null) row["Клики"] = row.clicks;
              // legacy fields from older snapshots (cloud schema <=0)
              if (row.spent != null && row["Потрачено всего, ₽"] == null && row["Потрачено всего, Р"] == null && row["Потрачено всего"] == null) row["Потрачено всего, ₽"] = row.spent;
              if (row.adds != null && row["Добавили аудио"] == null && row["Добавления"] == null) row["Добавили аудио"] = row.adds;
              if (row.listens != null && row["Начали прослушивание"] == null && row["Прослушивания"] == null) row["Начали прослушивание"] = row.listens;
              if (row.streams != null && row["Начали прослушивание"] == null && row["Прослушивания"] == null) row["Начали прослушивание"] = row.streams;

              // legacy fields from older snapshots
              if (row.spent != null && row["Потрачено всего, ₽"] == null && row["Потрачено всего, Р"] == null) row["Потрачено всего, ₽"] = row.spent;
              if (row.cost != null && row["Потрачено всего, ₽"] == null && row["Потрачено всего, Р"] == null) row["Потрачено всего, ₽"] = row.cost;
              if (row.adds != null && row["Добавили аудио"] == null && row["Добавления"] == null) row["Добавили аудио"] = row.adds;
              if (row.additions != null && row["Добавили аудио"] == null && row["Добавления"] == null) row["Добавили аудио"] = row.additions;
              if (row.listens != null && row["Начали прослушивание"] == null && row["Прослушивания"] == null) row["Начали прослушивание"] = row.listens;
              if (row.streams != null && row["Начали прослушивание"] == null && row["Прослушивания"] == null) row["Начали прослушивание"] = row.streams;
              if (row.age != null && row["Возраст"] == null) row["Возраст"] = row.age;
              if (row.sex != null && row["Пол"] == null) row["Пол"] = row.sex;
            }
            return normalizeRowKeys_(row);
          });
        }
      }
    }


    // hydrate image urls
    if (r.coverKey) r.coverDataUrl = await cloudObjectUrl_(r.coverKey);
    for (const c of (r.communities || [])){
      for (const cr of (c.creatives || [])){
        if (cr.objectKey) cr.dataUrl = await cloudObjectUrl_(cr.objectKey);
      }
    }

    r.cloudOnly = false;
    r._cloudHydrated = true;
    saveDB_();
    return r;
  }

  // --- Image convert helpers (Blob/File -> WebP) ---
  async function blobToWebpBlob_(blob, {maxSide=2000, quality=0.86} = {}){
    const bmp = await createImageBitmap(blob);
    const scale = Math.min(1, maxSide / Math.max(bmp.width, bmp.height));
    const w = Math.max(1, Math.round(bmp.width * scale));
    const h = Math.max(1, Math.round(bmp.height * scale));

    const canvas = document.createElement("canvas");
    canvas.width = w; canvas.height = h;
    const ctx = canvas.getContext("2d");
    ctx.drawImage(bmp, 0, 0, w, h);

    const webp = await new Promise(res => canvas.toBlob(res, "image/webp", quality));
    if (!webp) throw new Error("WEBP encode failed");
    return { blob: webp, w, h };
  }

  function safeKeyPart_(s){
    return String(s||"").trim().replace(/[^a-zA-Z0-9_-]+/g, "_").slice(0,80) || "x";
  }
  function tsKey_(){
    return new Date().toISOString().replace(/[:.]/g,"-");
  }

  // Build minimal snapshot for cloud (only fields needed to re-render/print)
  function buildSnapshot_(r, versionPrefix){
    const snap = {
      schema: 1,
      releaseId: r.releaseId,
      title: r.title,
      artists: r.artists,
      track: r.track,
      createdAt: r.createdAt,
      updatedAt: r.updatedAt,
      versionPrefix,
      coverKey: r.coverKey || null,
      streams: Array.isArray(r.streams) ? r.streams : [],
      communities: (r.communities || []).map(c => ({
        communityId: c.communityId,
        name: c.name,
        vkId: c.vkId,
        // keep full table, but only columns used by report
        adsRows: (c.adsRows || []).map(row0 => normalizeRowKeys_(row0 || {})),
        demoFiles: (c.demoFiles || []).map(df => ({
          id: df.id,
          name: df.name,
          totals: df.totals || null,
          rows: (df.rows || []).map(row0 => normalizeRowKeys_(row0 || {}))
        })),
        creatives: (c.creatives || []).map(cr => ({
          id: cr.id,
          name: cr.name,
          objectKey: cr.objectKey || null
        }))
      }))
    };
    return snap;
  }

  // Save current release version to Yandex Object Storage:
  // - upload cover.webp (if present)
  // - upload creatives/*.webp (if present)
  // - upload data.json snapshot
  async function cloudSaveCurrentReleaseVersion_(){
    const r = getCurrentRelease_();
    if (!r) throw new Error("Нет выбранного релиза");

    const ts = tsKey_();
    const prefix = `releases/${r.releaseId}/versions/${ts}/`;

    // Cover (always re-upload if we have source)
if (r.coverDataUrl){
  const src = await (await fetch(r.coverDataUrl)).blob();
  const { blob: webp } = await blobToWebpBlob_(src, { maxSide: 1400, quality: 0.86 });
  const key = prefix + "cover.webp";
  const { url } = await cloudPresignPut_({ key, contentType: "image/webp" });
  await cloudPut_({ url, blob: webp, contentType: "image/webp" });
  r.coverKey = key;
}

    // Creatives (always re-upload if we have source)
for (const c of (r.communities || [])){
  for (const cr of (c.creatives || [])){
    if (!cr.dataUrl) continue;

    const src = await (await fetch(cr.dataUrl)).blob();
    const { blob: webp } = await blobToWebpBlob_(src, { maxSide: 2000, quality: 0.86 });

    const idPart = safeKeyPart_(cr.id || cr.name || "creative");
    const key = prefix + `creatives/${idPart}.webp`;

    const { url } = await cloudPresignPut_({ key, contentType: "image/webp" });
    await cloudPut_({ url, blob: webp, contentType: "image/webp" });

    cr.objectKey = key;
  }
}

    // Snapshot
    const snap = buildSnapshot_(r, prefix);
    const dataKey = prefix + "data.json";
    const { url } = await cloudPresignPut_({ key: dataKey, contentType: "application/json" });
    await cloudPut_({ url, blob: JSON.stringify(snap), contentType: "application/json" });

    r.cloudVersions = Array.isArray(r.cloudVersions) ? r.cloudVersions : [];
    r.cloudVersions.unshift({ ts, dataKey });
    r.lastCloudVersion = ts;
    saveDB_();

    // Update global cloud index so releases are visible in other browsers
    await cloudUpsertIndexOnSave_(r, { ts, dataKey, prefix });

    return { ts, dataKey };
  }

  function loadDB_(){
    try{
      const raw = localStorage.getItem(StoreKey);
      if (!raw) return { releases: {} };
      const parsed = JSON.parse(raw);
      if (!parsed.releases) parsed.releases = {};
      return parsed;
    }catch(e){
      console.warn("DB reset:", e);
      return { releases: {} };
    }
  }
  function saveDB_(){
    // Do not persist short-lived presigned URLs (they expire ~5 minutes)
    const db = JSON.parse(JSON.stringify(state.db || {}));

    try{
      const releases = db.releases || {};
      for (const rid in releases){
        const r = releases[rid];
        if (!r) continue;

        if (typeof r.coverDataUrl === "string" && r.coverDataUrl.includes("X-Amz-Algorithm=AWS4-HMAC-SHA256")){
          r.coverDataUrl = null;
          r._cloudHydrated = false;
        }
        if (Array.isArray(r.communities)){
          for (const c of r.communities){
            if (!c || !Array.isArray(c.creatives)) continue;
            for (const cr of c.creatives){
              if (!cr) continue;
              if (typeof cr.dataUrl === "string" && cr.dataUrl.includes("X-Amz-Algorithm=AWS4-HMAC-SHA256")){
                cr.dataUrl = null;
                r._cloudHydrated = false;
              }
            }
          }
        }
      }
    }catch(e){
      // ignore sanitize errors; fall back to storing as-is
    }

    localStorage.setItem(StoreKey, JSON.stringify(db));
  }
  function uid_(){
    // короткий ID для читаемости, но достаточно уникальный для MVP
    return "rel_" + Math.random().toString(16).slice(2,10) + "_" + Date.now().toString(16).slice(-6);
  }
  function nowIso_(){ return new Date().toISOString(); }

  /** -----------------------------
   *  Domain helpers
   * -----------------------------*/
  function releaseTitle_(artists, track){
    const a = (artists || "").trim();
    const t = (track || "").trim();
    if (!a && !t) return "—";
    if (a && !t) return a;
    if (!a && t) return t;
    return `${a} - ${t}`;
  }

  // ===== Filters helpers (Releases / History) =====
  function uniqSorted_(arr){
    return Array.from(new Set((arr || []).filter(Boolean).map(s=>String(s).trim()).filter(Boolean))).sort((a,b)=>a.localeCompare(b,"ru"));
  }

  function splitArtistsTokens_(s){
    const raw = String(s||"").split(",").map(x=>x.trim()).filter(Boolean);
    // also keep original full string as a token (for exact match on combined artists)
    const full = String(s||"").trim();
    const out = [];
    if (full) out.push(full);
    out.push(...raw);
    return uniqSorted_(out);
  }

  function collectReleaseFilterOptions_(releases){
    const artists = [];
    const tracks = [];
    const commNames = [];
    const commIds = [];
    for (const r of (releases || [])){
      if (!r) continue;
      artists.push(...splitArtistsTokens_(r.artists));
      if (r.track) tracks.push(String(r.track).trim());
      for (const c of (r.communities || [])){
        if (!c) continue;
        if (c.name) commNames.push(String(c.name).trim());
        if (c.vkId) commIds.push(String(c.vkId).trim());
      }
    }
    return {
      artists: uniqSorted_(artists),
      tracks: uniqSorted_(tracks),
      commNames: uniqSorted_(commNames),
      commIds: uniqSorted_(commIds)
    };
  }
  // Build pairing maps for community name <-> vkId (to keep filters consistent)
  function buildCommunityPairs_(releases){
    const nameToIds = new Map(); // name -> Set(ids)
    const idToNames = new Map(); // id -> Set(names)

    for (const r of (releases || [])){
      for (const c of (r.communities || [])){
        const name = String(c?.name || "").trim();
        const id = String(c?.vkId || "").trim();
        if (!name || !id) continue;

        if (!nameToIds.has(name)) nameToIds.set(name, new Set());
        nameToIds.get(name).add(id);

        if (!idToNames.has(id)) idToNames.set(id, new Set());
        idToNames.get(id).add(name);
      }
    }
    return { nameToIds, idToNames };
  }

  function sortedRu_(arr){
    return Array.from(arr || []).map(x=>String(x).trim()).filter(Boolean).sort((a,b)=>a.localeCompare(b,"ru"));
  }

  const _pairSyncGuard = { releases:false, history:false };


  function populateSelect_(sel, values, { keepValue="" } = {}){
    if (!sel) return;
    const cur = keepValue || sel.value || "";
    sel.innerHTML = "";
    const optAll = document.createElement("option");
    optAll.value = "";
    optAll.textContent = "Все";
    sel.appendChild(optAll);

    for (const v of (values || [])){
      const o = document.createElement("option");
      o.value = v;
      o.textContent = v;
      sel.appendChild(o);
    }
    // restore selection if still exists
    if (cur && (values || []).includes(cur)) sel.value = cur;
    else sel.value = "";
  }

  function getFiltersFromDom_(scope){
    const p = scope + "-filter-";
    const artist = ($("#"+p+"artist")?.value || "").trim();
    const track = ($("#"+p+"track")?.value || "").trim();
    const community = ($("#"+p+"community")?.value || "").trim();
    const communityId = ($("#"+p+"communityId")?.value || "").trim();
    return { artist, track, community, communityId };
  }

  function dateIsoToYmd_(iso){
    if (!iso) return "";
    // ISO expected: 2026-02-05T... -> 2026-02-05
    return String(iso).slice(0,10);
  }

  function betweenYmd_(ymd, from, to){
    if (!ymd) return false;
    if (from && ymd < from) return false;
    if (to && ymd > to) return false;
    return true;
  }

  function releasePassesFilters_(r, f){
    if (!r) return false;
    if (f.artist){
      const tokens = splitArtistsTokens_(r.artists);
      if (!tokens.includes(f.artist)) return false;
    }
    if (f.track){
      if (String(r.track||"").trim() !== f.track) return false;
    }
    if (f.community){
      const ok = (r.communities || []).some(c => String(c?.name||"").trim() === f.community);
      if (!ok) return false;
    }
    if (f.communityId){
      const ok = (r.communities || []).some(c => String(c?.vkId||"").trim() === f.communityId);
      if (!ok) return false;
    }
    return true;
  }

   function releasePassesFiltersExcluding_(r, f, excludeKey){
  const ff = f || {};
  const f2 = {
    artist: excludeKey === "artist" ? "" : (ff.artist || ""),
    track: excludeKey === "track" ? "" : (ff.track || ""),
    community: excludeKey === "community" ? "" : (ff.community || ""),
    communityId: excludeKey === "communityId" ? "" : (ff.communityId || "")
  };
  return releasePassesFilters_(r, f2);
}

function collectOptionsByFacet_(allReleases, scope){
  const f = getFiltersFromDom_(scope);

  const by = (excludeKey) =>
    (allReleases || []).filter(r => releasePassesFiltersExcluding_(r, f, excludeKey));

  return {
    artist: collectReleaseFilterOptions_(by("artist")).artists,
    track: collectReleaseFilterOptions_(by("track")).tracks,
    community: collectReleaseFilterOptions_(by("community")).commNames,
    communityId: collectReleaseFilterOptions_(by("communityId")).commIds
  };
}

  function updateFiltersUi_(scope, releases){
  const p = scope + "-filter-";

  const cur = getFiltersFromDom_(scope);
  const facets = collectOptionsByFacet_(releases, scope);

  // Build community name<->id pairing under OTHER selected filters (artist/track etc.)
  const baseForPairs = (releases || []).filter(r => {
    const f2 = { ...cur, community: "", communityId: "" };
    return releasePassesFilters_(r, f2);
  });
  const { nameToIds, idToNames } = buildCommunityPairs_(baseForPairs);

  const selArtist = $("#"+p+"artist");
  const selTrack = $("#"+p+"track");
  const selComm = $("#"+p+"community");
  const selCommId = $("#"+p+"communityId");

  // Faceted options
  populateSelect_(selArtist, facets.artist, { keepValue: cur.artist });
  populateSelect_(selTrack, facets.track, { keepValue: cur.track });

  // Community + CommunityId are linked (pairing)
  let commOptions = facets.community;
  let commIdOptions = facets.communityId;

  if (cur.community){
    commIdOptions = sortedRu_(nameToIds.get(cur.community) || []);
  }
  if (cur.communityId){
    commOptions = sortedRu_(idToNames.get(cur.communityId) || []);
  }

  populateSelect_(selComm, commOptions, { keepValue: cur.community });
  populateSelect_(selCommId, commIdOptions, { keepValue: cur.communityId });

  // Rebuild searchable UI (combobox) after options update
  if (typeof enhanceSearchableSelects_ === "function") enhanceSearchableSelects_();

  // Auto-sync paired fields when mapping is unambiguous
  if (_pairSyncGuard[scope]) return;
  try{
    _pairSyncGuard[scope] = true;

    if (selComm && selCommId){
      const name = (selComm.value || "").trim();
      const id = (selCommId.value || "").trim();

      if (name && !id){
        const ids = sortedRu_(nameToIds.get(name) || []);
        if (ids.length === 1){
          selCommId.value = ids[0];
          selCommId.dispatchEvent(new Event("change", { bubbles: true }));
        }
      } else if (id && !name){
        const names = sortedRu_(idToNames.get(id) || []);
        if (names.length === 1){
          selComm.value = names[0];
          selComm.dispatchEvent(new Event("change", { bubbles: true }));
        }
      }
    }
  }finally{
    _pairSyncGuard[scope] = false;
  }
}

  function releaseSearchHaystack_(r){
    const parts = [];
    if (!r) return "";
    parts.push(r.releaseId, r.title, r.artists, r.track, r.createdAt, r.updatedAt, r.lastCloudVersion);
    try{
      parts.push(fmtDate_(r.createdAt), fmtDate_(r.updatedAt));
    }catch(e){}
    for (const c of (r.communities || [])){
      if (!c) continue;
      parts.push(c.name, c.vkId, c.communityId);
    }
    return String(parts.filter(Boolean).join(" ")).toLowerCase();
  }

  function releaseMatchesQuery_(r, q){
    const qq = String(q||"").trim().toLowerCase();
    if (!qq) return true;
    return releaseSearchHaystack_(r).includes(qq);
  }



  function formatMoney_(n){
    if (!isFinite(n)) return "—";
    const v = Math.round(n);
    return v.toLocaleString("ru-RU") + " ₽";
  }
   function formatMoney2_(n){
  if (!isFinite(n)) return "—";
  return new Intl.NumberFormat("ru-RU", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }).format(Number(n)) + " ₽";
  }
  function formatInt_(n){
    if (!isFinite(n)) return "—";
    const v = Math.round(n);
    return v.toLocaleString("ru-RU");
  }
  function safeDiv_(a,b){
    const A = Number(a||0), B = Number(b||0);
    if (!B) return null;
    return A/B;
  }


/** -----------------------------
 *  Demo files (multi-upload) helpers
 * -----------------------------*/
function ensureDemoStorage_(c){
  // Migration/backward-compat: previously we stored a single demoRows array.
  if (Array.isArray(c.demoRows) && c.demoRows.length){
    c.demoFiles = Array.isArray(c.demoFiles) ? c.demoFiles : [];
    if (!c.demoFiles.length){
      c.demoFiles.push({
        id: "demo_" + Math.random().toString(16).slice(2,10),
        name: "demo.xlsx",
        rows: c.demoRows,
        totals: null
      });
    }
    c.demoRows = null;
  } else {
    c.demoFiles = Array.isArray(c.demoFiles) ? c.demoFiles : [];
  }
}

function getDemoRows_(c){
  // If some old state still has demoRows, use it.
  if (Array.isArray(c.demoRows) && c.demoRows.length) return c.demoRows;
  const files = Array.isArray(c.demoFiles) ? c.demoFiles : [];
  const out = [];
  for (const f of files){
    if (f && Array.isArray(f.rows) && f.rows.length) out.push(...f.rows);
  }
  return out;
}

function getDemoTotals_(c){
  // Sums totals (Итого) across all uploaded demography files.
  ensureDemoStorage_(c);
  const files = Array.isArray(c.demoFiles) ? c.demoFiles : [];
  let impr = 0, clicks = 0, has = false;
  for (const f of files){
    const t = f && f.totals;
    if (!t) continue;
    const ti = Number(t.impr || 0);
    const tc = Number(t.clicks || 0);
    if (ti || tc) has = true;
    impr += ti;
    clicks += tc;
  }
  return has ? { impr, clicks } : null;
}


  /** -----------------------------
   *  Router
   * -----------------------------*/
  const routes = ["releases","editor","history","report"];
  function go_(route){
    routes.forEach(r=>{
      const page = $("#page-"+r);
      if (page) page.classList.toggle("active", r===route);
      const btn = $(`.tabBtn[data-route="${r}"]`);
      if (btn) btn.classList.toggle("active", r===route);
    });
    if (route==="releases") renderReleases_();
    if (route==="editor") renderEditor_();
    if (route==="history") renderHistory_();
    if (route==="report") renderReport_();
  }

  /** -----------------------------
   *  Init nav
   * -----------------------------*/
  $$(".tabBtn").forEach(b=>b.addEventListener("click", ()=>go_(b.dataset.route)));

  // Search inputs (Releases / History) — filter controls
  function bindFilters_(scope){
    const p = scope + "-filter-";
    const ids = [
      p + "artist",
      p + "track",
      p + "community",
      p + "communityId"];
    ids.forEach(id=>{
      const el = $("#"+id);
      if (!el) return;
      el.addEventListener("change", ()=>{
        if (scope === "releases") renderReleases_();
        if (scope === "history") renderHistory_();
      });
      el.addEventListener("input", ()=>{
        if (scope === "releases") renderReleases_();
        if (scope === "history") renderHistory_();
      });
    });

    const clearBtn = $("#"+scope+"-filters-clear");
    if (clearBtn){
      clearBtn.addEventListener("click", ()=>{
  const page = $("#page-"+scope) || document;
  const prefix = scope + "-filter-";

  // Важно: не спамим render'ами при очистке
  const prevGuard = _pairSyncGuard[scope];
  _pairSyncGuard[scope] = true;
  try{
    // 1) Clear native controls that have ids with the prefix
    page.querySelectorAll(`[id^="${prefix}"]`).forEach(el=>{
      if (!el) return;
      const tag = (el.tagName || "").toUpperCase();

      if (tag === "SELECT"){
        el.value = "";
      } else if (tag === "INPUT" || tag === "TEXTAREA"){
        el.value = "";
      }
    });

    // 2) Clear combobox inputs (visible searchable UI), if any
    page.querySelectorAll(".comboSelect .comboInput").forEach(inp=>{
      inp.value = "";
      inp.dispatchEvent(new Event("input", { bubbles: true }));
    });
  } finally {
    _pairSyncGuard[scope] = prevGuard;
  }

  // Один финальный ререндер в “чистом” состоянии
  if (scope === "releases") renderReleases_();
  if (scope === "history") renderHistory_();
});
    }
  }

  bindFilters_("releases");
  bindFilters_("history");
  // --- Searchable dropdowns for filters (combobox) ---
  function enhanceSearchableSelect_(sel){
    if (!sel || sel.dataset.searchableEnhanced === "1") return;
    sel.dataset.searchableEnhanced = "1";

    const wrap = document.createElement("div");
    wrap.className = "comboSelect";

    const input = document.createElement("input");
    input.type = "text";
    input.className = "comboInput";
    input.autocomplete = "off";
    input.spellcheck = false;
    input.placeholder = "Все";

    // Menu is rendered in <body> to avoid clipping by overflow/stacking contexts
    const menu = document.createElement("div");
    menu.className = "comboMenu";
    menu.hidden = true;
    menu.style.position = "fixed";
    menu.style.zIndex = "9999";
    menu.dataset.ownerSelectId = sel.id || "";

    // move into wrapper
    const parent = sel.parentElement;
    parent.insertBefore(wrap, sel);
    wrap.appendChild(input);
    wrap.appendChild(sel);
    document.body.appendChild(menu);

    sel.style.display = "none";

    function getOptions_(){
      return Array.from(sel.options || []).map(o => ({ value: o.value, text: o.textContent || "" }));
    }

    function syncInputFromSelect_(){
      const opt = sel.selectedOptions && sel.selectedOptions[0] ? sel.selectedOptions[0] : null;
      const txt = opt ? (opt.textContent || "") : "";
      // показываем выбранное значение; для "Все" оставляем поле пустым (placeholder = "Все")
      input.value = (sel.value ? txt : "");
    }

    function positionMenu_(){
      if (menu.hidden) return;
      const r = input.getBoundingClientRect();
      menu.style.left = r.left + "px";
      menu.style.top = (r.bottom + 4) + "px";
      menu.style.width = r.width + "px";
      // cap height to viewport
      const maxH = Math.max(120, window.innerHeight - (r.bottom + 12));
      menu.style.maxHeight = maxH + "px";
      menu.style.overflowY = "auto";
    }

    function open_(){
      render_(input.value);
      menu.hidden = false;
      positionMenu_();
    }
    function close_(){
      menu.hidden = true;
    }

    function render_(q){
      const qq = String(q || "").trim().toLowerCase();
      const opts = getOptions_();
      const filtered = qq ? opts.filter(x => x.text.toLowerCase().includes(qq)) : opts;

      menu.innerHTML = "";
      if (!filtered.length){
        const empty = document.createElement("div");
        empty.className = "comboEmpty";
        empty.textContent = "Ничего не найдено";
        menu.appendChild(empty);
        positionMenu_();
        return;
      }

      for (const o of filtered){
        const div = document.createElement("div");
        div.className = "comboOpt";
        div.textContent = o.text || "—";
        div.addEventListener("mousedown", (ev)=>{
          ev.preventDefault(); // не терять фокус до выбора
          sel.value = o.value;
          sel.dispatchEvent(new Event("change", { bubbles: true }));
          syncInputFromSelect_();
          close_();
        });
        menu.appendChild(div);
      }
      positionMenu_();
    }

    input.addEventListener("focus", open_);
    input.addEventListener("click", open_);
    input.addEventListener("input", ()=>{ if (menu.hidden) menu.hidden = false; render_(input.value); });

    // If user clears input -> treat as "Все"
    input.addEventListener("keydown", (e)=>{
      if (e.key === "Escape"){ close_(); input.blur(); }
      if (e.key === "Enter"){
        // if nothing typed -> set "Все"
        if (!String(input.value||"").trim()){
          sel.value = "";
          sel.dispatchEvent(new Event("change", { bubbles: true }));
          syncInputFromSelect_();
          close_();
        }
      }
    });

    sel.addEventListener("change", syncInputFromSelect_);

    // Close on outside click (covers both wrap + menu in body)
    document.addEventListener("mousedown", (e)=>{
      if (wrap.contains(e.target)) return;
      if (menu.contains(e.target)) return;
      close_();
    });

    // Keep menu aligned
    window.addEventListener("scroll", positionMenu_, true);
    window.addEventListener("resize", positionMenu_);

    // initial sync
    syncInputFromSelect_();
  }

  function enhanceSearchableSelects_(){
  $$('select[data-searchable="1"]').forEach(enhanceSearchableSelect_);
}

  // run once at startup
  enhanceSearchableSelects_();



  /** -----------------------------
   *  Releases page
   * -----------------------------*/
  const newReleaseCard = $("#new-release-card");
  $("#btn-new-release").addEventListener("click", ()=>{
    newReleaseCard.hidden = false;
    $("#nr-artists").value = "";
    $("#nr-track").value = "";
    $("#nr-title-preview").textContent = "—";
    $("#nr-artists").focus();
  });
  $("#btn-cancel-new-release").addEventListener("click", ()=>{
    newReleaseCard.hidden = true;
  });

  const btnCloudRefresh = $("#btn-cloud-refresh");
  if (btnCloudRefresh){
    btnCloudRefresh.addEventListener("click", async ()=>{
      try{
        setLocked_(btnCloudRefresh, true);
        const t0 = btnCloudRefresh.textContent;
        btnCloudRefresh.textContent = "Обновление…";
        await cloudSyncFromIndex_();
        btnCloudRefresh.textContent = t0 || "Обновить из облака";
      }catch(e){
        console.error(e);
        alert("Не удалось обновить из облака: " + (e && e.message ? e.message : e));
        btnCloudRefresh.textContent = "Обновить из облака";
      }finally{
        setLocked_(btnCloudRefresh, false);
      }
    });
  }


  function updateNewReleasePreview_(){
    const title = releaseTitle_($("#nr-artists").value, $("#nr-track").value);
    $("#nr-title-preview").textContent = title;
  }
  $("#nr-artists").addEventListener("input", updateNewReleasePreview_);
  $("#nr-track").addEventListener("input", updateNewReleasePreview_);

  $("#btn-create-release").addEventListener("click", ()=>{
    const artists = ($("#nr-artists").value||"").trim();
    const track = ($("#nr-track").value||"").trim();
    const title = releaseTitle_(artists, track);
    if (!artists || !track){
      alert("Заполни «Артисты» и «Трек»");
      return;
    }
    const id = uid_();
    state.db.releases[id] = {
      releaseId: id,
      artists,
      track,
      title,
      createdAt: nowIso_(),
      updatedAt: nowIso_(),
      coverDataUrl: null,
      streams: null, // [{dateISO, vk}]
      communities: [] // {communityId, name, vkId, adsRows, demoRows, creatives:[]}
    };
    saveDB_();
    state.currentReleaseId = id;
    newReleaseCard.hidden = true;
    renderReleases_();
    syncSelectors_();
    go_("editor");
  });

   function buildCloudPayload_(r){
  const out = {
    releaseId: r.releaseId,
    artists: r.artists,
    track: r.track,
    title: r.title,
    createdAt: r.createdAt,
    updatedAt: r.updatedAt,
    streams: Array.isArray(r.streams) ? r.streams : [],
    communities: []
  };

  for (const c of (r.communities || [])){
    // ВАЖНО: тут ты сам задашь, какие колонки реально нужны
    const adsMin = (c.adsRows || []).map(x => ({
      adId: x.adId ?? x["ID объявления"] ?? null,
      name: x.name ?? x["Название объявления"] ?? null,
      impr: x.impr ?? x["Показы"] ?? 0,
      clicks: x.clicks ?? x["Клики"] ?? 0,
      adds: x.adds ?? x["Добавления"] ?? 0,
      spent: x.spent ?? x["Потрачено"] ?? 0,
      cpa: x.cpa ?? x["Ср. стоимость добавления"] ?? null
    }));

    // demoFiles: храним минимальную форму — rows+totals
    const demoFiles = (c.demoFiles || []).map(df => ({
      id: df.id,
      name: df.name,
      totals: df.totals || null,
      rows: (df.rows || []).map(rr => ({
        age: rr.age ?? rr["Возраст"] ?? null,
        sex: rr.sex ?? rr["Пол"] ?? null,
        impr: rr.impr ?? rr["Показы"] ?? 0,
        clicks: rr.clicks ?? rr["Клики"] ?? 0
      }))
    }));

    out.communities.push({
      communityId: c.communityId,
      name: c.name,
      vkId: c.vkId,
      adsRows: adsMin,
      demoFiles,
      // creatives будем хранить как objectKey (после загрузки)
      creatives: (c.creatives || []).map(cr => ({
        id: cr.id,
        name: cr.name,
        objectKey: cr.objectKey || null
      }))
    });
  }
  return out;
}

  function renderReleases_(){
    const list = $("#releases-list");

    const all = Object.values(state.db.releases || {});
    updateFiltersUi_("releases", all);


    const f = getFiltersFromDom_("releases");
    const releases = all
      .filter(r => releasePassesFilters_(r, f))
      .sort((a,b)=> (b.updatedAt||"").localeCompare(a.updatedAt||""));

    // показываем количество найденных (и общий объём)
    $("#releases-count").textContent = `${releases.length}/${all.length}`;

    list.innerHTML = "";
    if (!releases.length){
      list.innerHTML = `<div class="muted">Ничего не найдено по выбранным фильтрам.</div>`;
      return;
    }

    for (const r of releases){
      const item = document.createElement("div");
      item.className = "item";
      const counts = summarizeRelease_(r);
      item.innerHTML = `
        <div class="itemHead">
          <div>
            <div class="itemTitle">${escapeHtml_(r.title)}</div>
            <div class="itemMeta">ID: ${escapeHtml_(r.releaseId)} · Сообществ: ${counts.communities} · Креативов: ${counts.creatives} · Обновлено: ${fmtDate_(r.updatedAt)}</div>
          </div>
          <div class="itemActions">
            <button class="btn inline" data-act="open">Открыть</button>
            <button class="btn inline" data-act="build">Сформировать отчёт</button>
          </div>
        </div>
      `;
      item.querySelector('[data-act="open"]').addEventListener("click", async ()=>{
        state.currentReleaseId = r.releaseId;
        syncSelectors_();
        try{ await ensureReleaseHydrated_(r.releaseId); }catch(e){ console.error(e); }
        go_("editor");
      });
      item.querySelector('[data-act="build"]').addEventListener("click", async ()=>{
        state.currentReleaseId = r.releaseId;
        syncSelectors_();
        try{ await ensureReleaseHydrated_(r.releaseId); }catch(e){ console.error(e); }
        go_("report");
      });
      list.appendChild(item);
    }
  }



  /** -----------------------------
   *  Editor
   * -----------------------------*/
  const editorSelect = $("#editor-release-select");
  const reportSelect = $("#report-release-select");

  function syncSelectors_(){
    const releases = Object.values(state.db.releases || {})
      .sort((a,b)=> (b.updatedAt||"").localeCompare(a.updatedAt||""));

    function fillSelect_(sel){
      if (!sel) return;
      const prev = sel.value || "";
      sel.innerHTML = "";
      for (const r of releases){
        const o = document.createElement("option");
        o.value = r.releaseId;
        o.textContent = r.title || r.releaseId;
        sel.appendChild(o);
      }
      // Choose current release
      let choose = state.currentReleaseId;
      if (!choose || !state.db.releases[choose]){
        choose = releases[0] ? releases[0].releaseId : "";
      }
      state.currentReleaseId = choose || null;

      if (choose && (prev === choose || releases.find(x=>x.releaseId===choose))){
        sel.value = choose;
      } else if (releases[0]) {
        sel.value = releases[0].releaseId;
      }
    }

    fillSelect_(editorSelect);
    fillSelect_(reportSelect);
  }

  editorSelect.addEventListener("change", async ()=>{
    state.currentReleaseId = editorSelect.value || null;
    if (state.currentReleaseId) {
      try{ await ensureReleaseHydrated_(state.currentReleaseId); }catch(e){ console.error(e); }
    }
    renderEditor_();
  });
  reportSelect.addEventListener("change", async ()=>{
    state.currentReleaseId = reportSelect.value || null;
    if (state.currentReleaseId) {
      try{ await ensureReleaseHydrated_(state.currentReleaseId); }catch(e){ console.error(e); }
    }
    renderReport_();
  });

  function getCurrentRelease_(){
    const id = state.currentReleaseId;
    if (!id) return null;
    return state.db.releases[id] || null;
  }

  function setLocked_(el, locked){
    el.classList.toggle("locked", !!locked);
  }

  function renderEditor_(){
    syncSelectors_();

    const r = getCurrentRelease_();
    if (!r){
      // if none
      $("#pill-release-id").textContent = "—";
      $("#ed-artists").value = "";
      $("#ed-track").value = "";
      $("#ed-title").textContent = "—";
      return;
    }

    $("#pill-release-id").textContent = r.releaseId;

    // meta
    $("#ed-artists").value = r.artists || "";
    $("#ed-track").value = r.track || "";
    $("#ed-title").textContent = r.title || "—";

    // cover
    const hasCover = !!r.coverDataUrl;
    $("#pill-cover").textContent = hasCover ? "загружена" : "не загружена";
    $("#pill-cover").className = "pill " + (hasCover ? "badgeOk" : "badgeWarn");
    $("#cover-preview-wrap").hidden = !hasCover;
    $("#btn-delete-cover").hidden = !hasCover;
    if (hasCover) $("#cover-preview").src = r.coverDataUrl;

    // step locks
    setLocked_($("#step-streams"), !hasCover);
    $("#btn-add-streams").disabled = !hasCover;

    const hasStreams = Array.isArray(r.streams) && r.streams.length > 0;
    $("#pill-streams").textContent = hasStreams ? "загружено" : "не загружено";
    $("#pill-streams").className = "pill " + (hasStreams ? "badgeOk" : "badgeWarn");
    $("#btn-delete-streams").hidden = !hasStreams;

    setLocked_($("#step-communities"), !hasStreams);
    $("#btn-add-community").disabled = !hasStreams;

    // communities
    $("#pill-communities").textContent = String((r.communities||[]).length);
    renderCommunities_(r);

    // ready?
    const ready = isReleaseReady_(r);
    $("#pill-ready").textContent = ready ? "готов" : "не готов";
    $("#pill-ready").className = "pill " + (ready ? "badgeOk" : "badgeWarn");
    $("#btn-build-report").disabled = !ready;
    $("#btn-open-report").disabled = !ready;
    setLocked_($("#step-build"), false);
  }

  $("#btn-save-meta").addEventListener("click", ()=>{
    const r = getCurrentRelease_();
    if (!r) return;
    r.artists = ($("#ed-artists").value||"").trim();
    r.track = ($("#ed-track").value||"").trim();
    r.title = releaseTitle_(r.artists, r.track);
    r.updatedAt = nowIso_();
    saveDB_();
    syncSelectors_();
    $("#ed-title").textContent = r.title;
    renderReleases_();
    renderHistory_();
  });

  $("#cover-file").addEventListener("change", async (e)=>{
    const r = getCurrentRelease_();
    if (!r) return;
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const dataUrl = await readAsDataURL_(file);
    r.coverDataUrl = dataUrl;
    r.updatedAt = nowIso_();
    saveDB_();
    e.target.value = "";
    renderEditor_();
  });

  $("#btn-delete-cover").addEventListener("click", ()=>{
    const r = getCurrentRelease_();
    if (!r) return;
    if (!confirm("Удалить обложку?")) return;
    r.coverDataUrl = null;
    r.updatedAt = nowIso_();
    saveDB_();
    renderEditor_();
  });

   function closeAllModals_(){
  const ids = ["streams-modal","community-modal","delete-modal"];
  ids.forEach(id=>{
    const el = document.getElementById(id);
    if (el) el.hidden = true;
  });
}
   
  // Streams modal
  const streamsModal = $("#streams-modal");
  $("#btn-add-streams").addEventListener("click", ()=>{
    state.temp.streamsParsed = null;
    $("#streams-preview-wrap").hidden = true;
    $("#btn-save-streams").disabled = true;
    $("#streams-file").value = "";
    closeAllModals_();
    streamsModal.hidden = false;
  });
  $("#btn-close-streams-modal").addEventListener("click", ()=> streamsModal.hidden = true);

  $("#streams-file").addEventListener("change", async (e)=>{
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    try{
      const rows = await parseStreamsXlsx_(file);
      state.temp.streamsParsed = rows;
      renderMiniPreview_($("#streams-preview"), rows.slice(0,5), ["dateISO","vk"]);
      $("#streams-preview-wrap").hidden = false;
      $("#btn-save-streams").disabled = !(rows && rows.length);
    }catch(err){
      console.error(err);
      alert("Не удалось распарсить файл стримов. Проверь колонки «Дата» и «vk».");
      state.temp.streamsParsed = null;
      $("#btn-save-streams").disabled = true;
    }
  });

  $("#btn-save-streams").addEventListener("click", ()=>{
    const r = getCurrentRelease_();
    if (!r) return;
    const rows = state.temp.streamsParsed;
    if (!rows || !rows.length) return;
    r.streams = rows;
    r.updatedAt = nowIso_();
    saveDB_();
    streamsModal.hidden = true;
    renderEditor_();
  });

  $("#btn-delete-streams").addEventListener("click", ()=>{
    const r = getCurrentRelease_();
    if (!r) return;
    if (!confirm("Удалить стримы?")) return;
    r.streams = null;
    r.updatedAt = nowIso_();
    saveDB_();
    renderEditor_();
  });
  // Communities
  const commModal = $("#community-modal");
  $("#btn-add-community").addEventListener("click", ()=>{
    $("#cm-name").value = "";
    $("#cm-vkId").value = "";
    closeAllModals_();
    commModal.hidden = false;
    $("#cm-name").focus();
  });
  $("#btn-close-community-modal").addEventListener("click", ()=> commModal.hidden = true);

  $("#btn-save-community").addEventListener("click", ()=>{
    const r = getCurrentRelease_();
    if (!r) return;
    const name = ($("#cm-name").value||"").trim();
    const vkId = ($("#cm-vkId").value||"").trim();
    if (!name || !vkId){
      alert("Заполни название и ID сообщества");
      return;
    }
    const communityId = "com_" + Math.random().toString(16).slice(2,10);
    r.communities = r.communities || [];
    r.communities.push({
      communityId, name, vkId,
      adsRows: null,
      demoFiles: [], // [{id, name, rows}]
      demoRows: null, // backward-compat
      creatives: [] // {id, dataUrl, name}
    });
    r.updatedAt = nowIso_();
    saveDB_();
    commModal.hidden = true;
    renderEditor_();
  });

  function renderCommunities_(release){
    const root = $("#communities-list");
    root.innerHTML = "";
    const comms = release.communities || [];
    if (!comms.length){
      root.innerHTML = `<div class="muted">Пока нет сообществ. Нажми «+ Добавить сообщество».</div>`;
      return;
    }

    for (const c of comms){
      const adsOk = Array.isArray(c.adsRows) && c.adsRows.length>0;
      ensureDemoStorage_(c);
      const demoOk = getDemoRows_(c).length>0;
      const crCount = (c.creatives||[]).length;

      const card = document.createElement("div");
      card.className = "commCard";
      card.innerHTML = `
        <div class="commTop">
          <div>
            <div class="commName">${escapeHtml_(c.name)}</div>
            <div class="commSub">ID сообщества: <b>${escapeHtml_(c.vkId)}</b> · Внутр. ID: ${escapeHtml_(c.communityId)}</div>
          </div>
          <div class="itemActions">
            <button class="btn inline danger" data-act="del">Удалить сообщество</button>
          </div>
        </div>

        <div class="grid cols2" style="margin-top:12px">
          <div class="card" style="margin:0">
            <div class="cardTitle">
              <h3>Статистика объявлений</h3>
              <span class="pill ${adsOk ? "badgeOk":"badgeWarn"}">${adsOk ? "загружено":"не загружено"}</span>
            </div>
            <div class="row">
              <input type="file" accept=".xlsx" data-up="ads" />
              <button class="btn danger inline" data-del="ads" ${adsOk ? "":"disabled"}>Удалить</button>
            </div>
            <div class="miniTableWrap" ${adsOk ? "":"hidden"} data-prevwrap="ads">
              <div class="miniTableTitle">Превью (первые 5 строк)</div>
              <table class="miniTable" data-prev="ads"></table>
            </div>
          </div>

          <div class="card" style="margin:0">
            <div class="cardTitle">
              <h3>Демография</h3>
              <span class="pill ${demoOk ? "badgeOk":"badgeWarn"}">${demoOk ? "загружено":"не загружено"}</span>
            </div>
            <div class="row">
              <input type="file" accept=".xlsx" data-up="demo" multiple />
              <button class="btn danger inline" data-del="demo" ${demoOk ? "":"disabled"}>Удалить</button>
            </div>
<div class="demoFilesWrap" data-demo-files-wrap ${demoOk ? "":"hidden"}>
  <div class="miniTableTitle">Файлы демографии</div>
  <div class="demoFilesList" data-demo-files></div>
  <div class="row" style="margin-top:8px">
    <button class="btn danger inline" data-act="demo-clear" ${demoOk ? "":"disabled"}>Очистить все</button>
  </div>
</div>

            <div class="miniTableWrap" ${demoOk ? "":"hidden"} data-prevwrap="demo">
              <div class="miniTableTitle">Превью (первые 5 строк)</div>
              <table class="miniTable" data-prev="demo"></table>
            </div>
          </div>
        </div>

        <div class="card" style="margin:12px 0 0">
          <div class="cardTitle">
            <h3>Креативы (изображения)</h3>
            <span class="pill ${crCount ? "badgeOk":"badgeWarn"}">${crCount} шт</span>
          </div>
          <div class="row">
            <input type="file" accept="image/png,image/jpeg,image/webp" multiple data-up="creatives" />
          </div>
          <div class="row" style="margin-top:10px; overflow:auto; padding-bottom:4px" data-creatives-row></div>
        </div>
      `;

      // delete community
      card.querySelector('[data-act="del"]').addEventListener("click", ()=>{
        if (!confirm("Удалить сообщество и все его данные?")) return;
        const r = getCurrentRelease_();
        r.communities = (r.communities||[]).filter(x=>x.communityId!==c.communityId);
        r.updatedAt = nowIso_();
        saveDB_();
        renderEditor_();
      });

      // uploads
      card.querySelectorAll('input[type="file"][data-up]').forEach(inp=>{
        inp.addEventListener("change", async (e)=>{
  const kind = e.target.dataset.up;
  try{
    if (kind==="ads"){
      const file = e.target.files && e.target.files[0];
      if (!file) return;
      const rows = await parseAdsXlsx_(file);
      c.adsRows = rows;
      renderCommunityPreviews_(card, c);
    } else if (kind==="demo"){
      const files = Array.from(e.target.files || []);
      if (!files.length) return;
      ensureDemoStorage_(c);
      for (const f of files){
        const parsed = await parseDemoXlsx_(f);
        c.demoFiles.push({
          id: "demo_" + Math.random().toString(16).slice(2,10),
          name: f.name,
          rows: parsed.rows,
          totals: parsed.totals
        });
      }
      renderCommunityPreviews_(card, c);
    } else if (kind==="creatives"){
      const files = Array.from(e.target.files||[]);
      if (!files.length) return;
      for (const f of files){
        const dataUrl = await readAsDataURL_(f);
        c.creatives = c.creatives || [];
        c.creatives.push({ id:"cr_"+Math.random().toString(16).slice(2,10), name:f.name, dataUrl });
      }
      renderCommunityCreativesRow_(card, c);
    }
    const r = getCurrentRelease_();
    r.updatedAt = nowIso_();
    saveDB_();
    renderEditor_();
  }catch(err){
            console.error(err);
            alert("Ошибка парсинга файла. Проверь формат/колонки.");
          }finally{
            e.target.value = "";
          }
        });
      });

      // deletes
      card.querySelectorAll('button[data-del]').forEach(btn=>{
        btn.addEventListener("click", ()=>{
          const kind = btn.dataset.del;
          if (kind==="ads"){
            if (!confirm("Удалить файл объявлений и данные?")) return;
            c.adsRows = null;
          }
          if (kind==="demo"){
            if (!confirm("Удалить файлы демографии и данные?")) return;
            ensureDemoStorage_(c);
            c.demoFiles = [];
            c.demoRows = null;
          }
          const r = getCurrentRelease_();
          r.updatedAt = nowIso_();
          saveDB_();
          renderEditor_();
        });
      });

      // previews
      renderCommunityPreviews_(card, c);
      renderCommunityCreativesRow_(card, c);

      root.appendChild(card);
    }
  }

  function renderCommunityPreviews_(card, c){
    // ads
    const adsWrap = card.querySelector('[data-prevwrap="ads"]');
    const adsTbl = card.querySelector('[data-prev="ads"]');
    if (Array.isArray(c.adsRows) && c.adsRows.length){
      adsWrap.hidden = false;
      renderMiniPreview_(adsTbl, c.adsRows.slice(0,5), ["Название объявления","Потрачено всего, ₽","Показы","Начали прослушивание","Добавили аудио"]);
    } else {
      adsWrap.hidden = true;
    }
// demo
ensureDemoStorage_(c);
const dWrap = card.querySelector('[data-prevwrap="demo"]');
const dTbl = card.querySelector('[data-prev="demo"]');
const dfWrap = card.querySelector('[data-demo-files-wrap]');
const dfList = card.querySelector('[data-demo-files]');
const dfClearBtn = card.querySelector('[data-act="demo-clear"]');

const demoRows = getDemoRows_(c);

if (demoRows.length){
  dWrap.hidden = false;
  if (dfWrap) dfWrap.hidden = false;

  // Preview (first 5 rows)
  const demoCostKey =
    (demoRows?.[0] && ("Цена за результат, ₽" in demoRows[0])) ? "Цена за результат, ₽"
    : (demoRows?.[0] && ("Цена за результат, Р" in demoRows[0])) ? "Цена за результат, Р"
    : "Цена за результат, ₽";

  renderMiniPreview_(dTbl, demoRows.slice(0,5), ["Возраст","Пол","Показы","Клики", demoCostKey]);

  // Files list
  if (dfList){
    dfList.innerHTML = "";
    const files = c.demoFiles || [];
    for (const f of files){
      const row = document.createElement("div");
      row.className = "demoFileItem";
      row.innerHTML = `
        <div class="demoFileName">${escapeHtml_(f.name || "demo.xlsx")}</div>
        <button class="btn inline danger" data-demo-del="${escapeHtml_(f.id)}">Удалить</button>
      `;
      row.querySelector('[data-demo-del]').addEventListener("click", ()=>{
        if (!confirm("Удалить этот файл демографии?")) return;
        ensureDemoStorage_(c);
        c.demoFiles = (c.demoFiles||[]).filter(x=>x.id !== f.id);
        const r = getCurrentRelease_();
        r.updatedAt = nowIso_();
        saveDB_();
        renderEditor_();
      });
      dfList.appendChild(row);
    }
  }

  // Clear all
  if (dfClearBtn){
    dfClearBtn.disabled = !(c.demoFiles && c.demoFiles.length);
    dfClearBtn.onclick = ()=>{
      if (!confirm("Удалить все файлы демографии?")) return;
      ensureDemoStorage_(c);
      c.demoFiles = [];
      c.demoRows = null;
      const r = getCurrentRelease_();
      r.updatedAt = nowIso_();
      saveDB_();
      renderEditor_();
    };
  }
} else {
  dWrap.hidden = true;
  if (dfWrap) dfWrap.hidden = true;
  if (dfList) dfList.innerHTML = "";
}
  }

  function renderCommunityCreativesRow_(card, c){
    const row = card.querySelector('[data-creatives-row]');
    if (!row) return;
    row.innerHTML = "";
    const list = c.creatives || [];
    if (!list.length){
      row.innerHTML = `<div class="muted">Нет креативов</div>`;
      return;
    }
    for (const cr of list){
      const wrap = document.createElement("div");
      wrap.style.display="grid";
      wrap.style.gap="6px";
      wrap.style.marginRight="10px";
      const img = document.createElement("img");
      img.src = cr.dataUrl;
      img.alt = cr.name || "creative";
      img.style.width = "140px";
      img.style.height = "110px";
      img.style.objectFit = "cover";
      img.style.borderRadius = "12px";
      img.style.border = "1px solid rgba(255,255,255,.10)";
      const cap = document.createElement("div");
      cap.className = "muted";
      cap.style.fontSize="11px";
      cap.style.width="140px";
      cap.style.whiteSpace="nowrap";
      cap.style.overflow="hidden";
      cap.style.textOverflow="ellipsis";
      cap.textContent = cr.name || "creative";
      const del = document.createElement("button");
      del.className = "btn inline danger";
      del.textContent = "Удалить";
      del.addEventListener("click", ()=>{
        if (!confirm("Удалить креатив?")) return;
        c.creatives = (c.creatives||[]).filter(x=>x.id!==cr.id);
        const r = getCurrentRelease_();
        r.updatedAt = nowIso_();
        saveDB_();
        renderEditor_();
      });
      wrap.appendChild(img);
      wrap.appendChild(cap);
      wrap.appendChild(del);
      row.appendChild(wrap);
    }
  }

  function isReleaseReady_(r){
    if (!r) return false;
    if (!r.coverDataUrl) return false;
    if (!Array.isArray(r.streams) || !r.streams.length) return false;
    const comms = r.communities || [];
    if (!comms.length) return false;
    for (const c of comms){
      if (!Array.isArray(c.adsRows) || !c.adsRows.length) return false;
      ensureDemoStorage_(c);
      if (!getDemoRows_(c).length) return false;
      if (!Array.isArray(c.creatives) || !c.creatives.length) return false; // можно сделать optional, но в MVP фиксируем
    }
    return true;
  }

  $("#btn-build-report").addEventListener("click", ()=>{
    const r = getCurrentRelease_();
    if (!r) return;
    if (!isReleaseReady_(r)){
      alert("Отчёт ещё не готов: проверь обложку, стримы и данные сообществ.");
      return;
    }
    // В MVP ничего “не билдим” серверно: просто переходим на отчёт.
    go_("report");
  });
  $("#btn-open-report").addEventListener("click", ()=> go_("report"));

  $("#btn-delete-release").addEventListener("click", ()=>{
    const r = getCurrentRelease_();
    if (!r) return;
    openDeleteModal_(r.releaseId);
  });

  /** -----------------------------
   *  History + deletion
   * -----------------------------*/
  const deleteModal = $("#delete-modal");
  let deleteTargetId = null;

  function openDeleteModal_(releaseId){
    const r = state.db.releases[releaseId];
    if (!r) return;
    deleteTargetId = releaseId;
    $("#delete-title").textContent = r.title;
    $("#delete-confirm").value = "";
    $("#btn-confirm-delete").disabled = true;
    closeAllModals_();
    deleteModal.hidden = false;
    $("#delete-confirm").focus();
  }

  $("#btn-close-delete-modal").addEventListener("click", ()=>{
    deleteModal.hidden = true;
    deleteTargetId = null;
  });

  $("#delete-confirm").addEventListener("input", ()=>{
    if (!deleteTargetId) return;
    const r = state.db.releases[deleteTargetId];
    const ok = ($("#delete-confirm").value||"").trim() === (r.title||"").trim();
    $("#btn-confirm-delete").disabled = !ok;
  });

  $("#btn-confirm-delete").addEventListener("click", ()=>{
    if (!deleteTargetId) return;
    const r = state.db.releases[deleteTargetId];
    const confirmText = ($("#delete-confirm").value||"").trim();
    if (confirmText !== (r.title||"").trim()){
      alert("Название не совпадает.");
      return;
    }
    delete state.db.releases[deleteTargetId];
    saveDB_();
    deleteModal.hidden = true;

    // update selection
    const ids = Object.keys(state.db.releases);
    state.currentReleaseId = ids.length ? ids[0] : null;
    syncSelectors_();
    renderReleases_();
    renderHistory_();
    renderReport_();
    go_("releases");
  });

  function renderHistory_(){
    const list = $("#history-list");

    const all = Object.values(state.db.releases || {});
    updateFiltersUi_("history", all);

    const f = getFiltersFromDom_("history");
    const releases = all
      .filter(r => releasePassesFilters_(r, f))
      .sort((a,b)=> (b.updatedAt||"").localeCompare(a.updatedAt||""));

    // count (found/total) if element exists
    const pill = $("#history-count");
    if (pill) pill.textContent = `${releases.length}/${all.length}`;

    list.innerHTML = "";
    if (!releases.length){
      list.innerHTML = `<div class="muted">Ничего не найдено по выбранным фильтрам.</div>`;
      return;
    }

    for (const r of releases){
      const s = summarizeRelease_(r);
      const item = document.createElement("div");
      item.className = "item";
      item.innerHTML = `
        <div class="itemHead">
          <div>
            <div class="itemTitle">${escapeHtml_(r.title)}</div>
            <div class="itemMeta">
              Сообществ: ${s.communities} · Файлы ads: ${s.adsFiles} · demo: ${s.demoFiles} · Креативы: ${s.creatives} · Стримов: ${s.streamRows}
            </div>
          </div>
          <div class="itemActions">
            <button class="btn inline" data-act="open">Открыть</button>
            <button class="btn inline" data-act="report">Отчёт</button>
            <button class="btn inline danger" data-act="delete">Удалить…</button>
          </div>
        </div>
      `;
      item.querySelector('[data-act="open"]').addEventListener("click", async ()=>{
        state.currentReleaseId = r.releaseId;
        syncSelectors_();
        try{ await ensureReleaseHydrated_(r.releaseId); }catch(e){ console.error(e); }
        go_("editor");
      });
      item.querySelector('[data-act="report"]').addEventListener("click", async ()=>{
        state.currentReleaseId = r.releaseId;
        syncSelectors_();
        try{ await ensureReleaseHydrated_(r.releaseId); }catch(e){ console.error(e); }
        go_("report");
      });
      item.querySelector('[data-act="delete"]').addEventListener("click", ()=> openDeleteModal_(r.releaseId));
      list.appendChild(item);
    }
  }

  function summarizeRelease_(r){
    const comms = r.communities || [];
    let adsFiles = 0, demoFiles = 0, creatives = 0;
    for (const c of comms){
      if (Array.isArray(c.adsRows) && c.adsRows.length) adsFiles++;
      ensureDemoStorage_(c);
      if ((c.demoFiles||[]).length) demoFiles++;
      creatives += (c.creatives||[]).length;
    }
    return {
      communities: comms.length,
      adsFiles,
      demoFiles,
      creatives,
      streamRows: Array.isArray(r.streams) ? r.streams.length : 0
    };
  }

  /** -----------------------------
   *  Report render
   * -----------------------------*/
  $("#btn-print").addEventListener("click", ()=>{
  go_("report");

  const r = getCurrentRelease_();
  const originalTitle = document.title;

  const safe = (s)=> String(s||"")
    .replace(/[\\/:*?"<>|]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const fileName = safe(`VK DIGITAL-ОТЧЁТ: ${r?.title || ""}`);

  document.title = fileName;

  setTimeout(()=>window.print(), 50);

  window.onafterprint = () => {
    document.title = originalTitle;
    window.onafterprint = null;
  };
});




// Cloud save button (рядом с "Печать / PDF")
const _btnCloudSave = $("#btn-cloud-save");
if (_btnCloudSave){
  _btnCloudSave.addEventListener("click", async ()=>{
    try{
      const r = getCurrentRelease_();
      if (!r) { alert("Нет выбранного релиза"); return; }
      if (!isReleaseReady_(r)) { alert("Отчёт не готов. Сначала догрузи данные в «Конструкторе»."); return; }

      setLocked_(_btnCloudSave, true);
      const originalText = _btnCloudSave.textContent;
      _btnCloudSave.textContent = "Сохранение…";

      const { ts } = await cloudSaveCurrentReleaseVersion_();

      _btnCloudSave.textContent = originalText || "Сохранить";
      alert("Сохранено в облако. Версия: " + ts);
    }catch(e){
      console.error(e);
      _btnCloudSave.textContent = "Сохранить";
      alert("Не удалось сохранить в облако: " + (e && e.message ? e.message : e));
    }finally{
      setLocked_(_btnCloudSave, false);
    }
  });
}

  function renderReport_(){
    syncSelectors_();
    const root = $("#report-root");
    root.innerHTML = "";
    const r = getCurrentRelease_();
    if (!r){
      root.innerHTML = `<div class="card"><div class="muted">Нет релиза для отчёта.</div></div>`;
      return;
    }
    if (!isReleaseReady_(r)){
      root.innerHTML = `<div class="card"><div class="muted">Отчёт не готов. Перейди в «Конструктор» и догрузи данные.</div></div>`;
      return;
    }

    const pages = buildReportPages_(r);
    for (const p of pages){
      root.appendChild(p);
    }
  }

  function buildReportPages_(r){
    const pages = [];

    // Cover
    pages.push(renderCoverPage_(r));

    const comms = r.communities || [];
    if (comms.length === 1){
      const c = comms[0];
      pages.push(renderResultsPage_(r, [{label: c.name, community: c}], false));
      const totalCr = (c.creatives||[]).length;
      for (let i=0; i<totalCr; i+=4){
        pages.push(renderCreativesPage_(r, c, i));
      }
      pages.push(...renderDetailPages_(r, c));
      pages.push(renderDemoPage_(r, c));
      pages.push(renderStreamsPage_(r));
    } else {
      // per community
      pages.push(renderResultsPage_(r, comms.map(c=>({label:c.name, community:c})), true)); // общие + по сообществам
      for (const c of comms){
        pages.push(renderResultsPage_(r, [{label: c.name, community: c}], false, true));
        pages.push(renderCreativesPage_(r, c));
        pages.push(...renderDetailPages_(r, c));
        pages.push(renderDemoPage_(r, c));
      }
      pages.push(renderStreamsPage_(r));
    }

    return pages.flat();
  }

  function sheet_(title){
    const el = document.createElement("div");
    el.className = "pageSheet";
    el.innerHTML = `<div class="sheetInner"></div>`;
    if (title){
      const inner = el.firstElementChild;
      const head = document.createElement("div");
      head.className = "sheetHeader";
      head.innerHTML = `<img class="lotusLogo" src="assets/logo_black.png" alt="Lotus Music"/><img class="lotusLogo" src="assets/logo_black.png" alt="Lotus Music"/>`;
      inner.appendChild(head);
    }
    return el;
  }

  function renderCoverPage_(r){
  const el = sheet_(false);
  const inner = el.querySelector(".sheetInner");

  inner.insertAdjacentHTML("beforeend", `
    <div class="coverHeader">
      <img class="lotusLogo" src="assets/logo_black.png" />
      <div class="coverTitles">
        <div class="coverTitleMain">DIGITAL-ОТЧЁТ</div>
        <div class="coverTitleSub">ТАРГЕТИРОВАННАЯ РЕКЛАМА ВКОНТАКТЕ</div>
      </div>
      <img class="lotusLogo" src="assets/logo_black.png" />
    </div>

    <div class="coverCenter">
      <div class="coverBlock">
        <img class="coverImg" src="${r.coverDataUrl}" alt="cover" />
        <div class="coverName">${escapeHtml_((r.title || "").toUpperCase())}</div>
      </div>
    </div>
  `);
  return el;
}
  function calcKpisFromAdsRows_(rows){
    const col_ = (row, keys)=>{
      for (const k of keys){
        if (row && row[k] != null && row[k] !== "") return row[k];
      }
      return "";
    };

    let spent=0, shows=0, listens=0, adds=0;
    const groupIds = new Set();
    for (const row0 of rows){
      const row = normalizeRowKeys_(row0 || {});
      spent   += num_(col_(row, ["Потрачено всего, ₽","Потрачено всего, Р","Потрачено всего","Расход, ₽","Расход, Р"]));
      shows   += num_(col_(row, ["Показы","Показов"]));
      listens += num_(col_(row, ["Начали прослушивание","Прослушивания","Начали прослушивание аудио"]));
      adds    += num_(col_(row, ["Добавили аудио","Добавления аудио","Добавили"]));

      const gid = (col_(row, ["ID группы","ID группы объявления","ID группы (объявление)"]) ?? "").toString().trim();
      if (gid) groupIds.add(gid);
    }
    const avgCost = safeDiv_(spent, adds);
    return { spent, shows, listens, adds, segments: groupIds.size, avgCost };
  }

  function calcImprClicksFromAdsRows_(rows){
    let impr = 0, clicks = 0;
    for (const row0 of (rows || [])){
      const row = normalizeRowKeys_(row0 || {});
      impr += num_(row["Показы"]);
      // Different exports may name clicks differently; keep a few variants.
      clicks += num_(row["Клики"] ?? row["Переходы"] ?? row["Переходы по ссылке"] ?? row["Клики по ссылке"] ?? 0);
    }
    return { impr, clicks };
  }

  function renderResultsPage_(r, commItems, isMultiSummary, addLabel=false){
    // if isMultiSummary: show "Общие показатели" + mini list? In MVP: first KPI = totals across all communities
    const el = sheet_(false);
    const inner = el.querySelector(".sheetInner");

    const pageTitle = isMultiSummary
    ? "Результаты продвижения (общие)"
    : (addLabel ? `Результаты: ${commItems[0].label}` : "Результаты продвижения");

     inner.insertAdjacentHTML("beforeend", `
    <div class="slideHeader">
      <img class="lotusLogo" src="assets/logo_black.png" alt="Lotus Music"/>
      <div class="slideHeaderTitle">${escapeHtml_(pageTitle)}</div>
      <img class="lotusLogo" src="assets/logo_black.png" alt="Lotus Music"/>
    </div>
  `);

    inner.insertAdjacentHTML("beforeend", `<div class="slideSpacer"></div>`);

      let rowsAll = [];
  for (const it of commItems) rowsAll = rowsAll.concat(it.community.adsRows || []);
  const k = calcKpisFromAdsRows_(rowsAll);

  // 1 вертикальная колонка (сверху вниз)
  inner.insertAdjacentHTML("beforeend", `
    <div class="kpiCols">
      <div class="kpiCol">
        <div class="kpiRow"><div class="kpiLabel">Потрачено всего</div><div class="kpiValue">${formatMoney2_(k.spent)}</div></div>
        <div class="kpiRow"><div class="kpiLabel">Показы</div><div class="kpiValue">${formatInt_(k.shows)}</div></div>
        <div class="kpiRow"><div class="kpiLabel">Прослушивания</div><div class="kpiValue">${formatInt_(k.listens)}</div></div>
        <div class="kpiRow"><div class="kpiLabel">Добавления</div><div class="kpiValue">${formatInt_(k.adds)}</div></div>
        <div class="kpiRow"><div class="kpiLabel">Сегменты</div><div class="kpiValue">${formatInt_(k.segments)}</div></div>
        <div class="kpiRow"><div class="kpiLabel">Ср. стоимость добавления</div><div class="kpiValue">${k.avgCost==null ? "—" : formatMoney2_(k.avgCost)}</div></div>
      </div>
    </div>
  `);

    return el;
  }

  function renderCreativesPage_(r, c, start=0){
  const el = sheet_(false);
  const inner = el.querySelector(".sheetInner");

  const total = (c.creatives||[]).length;
  const pageNo = Math.floor(start/4) + 1;
  const pages = Math.max(1, Math.ceil(total/4));
  const pageTitle = pages > 1
    ? `Креативы: ${c.name} (${pageNo}/${pages})`
    : `Креативы: ${c.name}`;

  inner.insertAdjacentHTML("beforeend", `
    <div class="slideHeader">
      <img class="lotusLogo" src="assets/logo_black.png" alt="Lotus Music"/>
      <div class="slideHeaderTitle">${escapeHtml_(pageTitle)}</div>
      <img class="lotusLogo" src="assets/logo_black.png" alt="Lotus Music"/>
    </div>
  `);
  inner.insertAdjacentHTML("beforeend", `<div class="slideSpacer"></div>`);

  const row = document.createElement("div");
  row.className = "creativesRow";

  const imgs = (c.creatives||[]).slice(start, start+4);
  for (const cr of imgs){
    const img = document.createElement("img");
    img.className = "creativeImg";
    img.src = cr.dataUrl;
    img.alt = cr.name || "creative";
    row.appendChild(img);
  }

  inner.appendChild(row);
  inner.insertAdjacentHTML("beforeend",
    `<div class="creativesNote">Показаны креативы ${start+1}–${start+imgs.length} из ${total}.</div>`
  );

  return el;
}

   function renderDetailPages_(r, c){
  const rows = (c.adsRows || []);
   rows.sort((a, b) => {
    const A = normalizeRowKeys_(a || {});
    const B = normalizeRowKeys_(b || {});
    const addsA = num_(A["Добавили аудио"] ?? A["Добавления аудио"] ?? A["Добавили"] ?? 0);
    const addsB = num_(B["Добавили аудио"] ?? B["Добавления аудио"] ?? B["Добавили"] ?? 0);
    return addsB - addsA;
  });
  const total = rows.length;
  const perPage = 16;
  const pagesTotal = Math.max(1, Math.ceil(total / perPage));
  const out = [];

  // даже если total=0 — делаем 1 страницу (с пустой таблицей и шапкой)
  for (let start = 0; start < Math.max(total, 1); start += perPage){
    out.push(renderDetailPageChunk_(r, c, start, perPage, pagesTotal));
    if (total === 0) break;
  }
  return out;
}

function renderDetailPageChunk_(r, c, start, perPage, pagesTotal){
  const el = sheet_(false);
  const inner = el.querySelector(".sheetInner");

  const total = (c.adsRows || []).length;
  const pageNo = Math.floor(start / perPage) + 1;

  const pageTitle = pagesTotal > 1
    ? `Детализация: ${c.name} (${pageNo}/${pagesTotal})`
    : `Детализация: ${c.name}`;

  inner.insertAdjacentHTML("beforeend", `
    <div class="slideHeader">
      <img class="lotusLogo" src="assets/logo_black.png" alt="Lotus Music"/>
      <div class="slideHeaderTitle">${escapeHtml_(pageTitle)}</div>
      <img class="lotusLogo" src="assets/logo_black.png" alt="Lotus Music"/>
    </div>
  `);

  inner.insertAdjacentHTML("beforeend", `<div class="slideSpacer"></div>`);

  const slice = (c.adsRows || []).slice(start, start + perPage);

  const table = document.createElement("table");
  table.className = "table detailTable";
  table.innerHTML = `
    <thead>
      <tr>
        <th>Объявление</th>
        <th>Потрачено</th>
        <th>Показы</th>
        <th>Прослуш.</th>
        <th>Добавл.</th>
        <th>Ср. цена добавл.</th>
      </tr>
    </thead>
    <tbody></tbody>
  `;

  const tb = table.querySelector("tbody");
  for (const row of slice){
    const r2 = normalizeRowKeys_(row || {});
    const spent = num_(r2["Потрачено всего, ₽"] ?? r2["Потрачено всего, Р"] ?? r2["Потрачено всего"] ?? "");
    const adds  = num_(r2["Добавили аудио"] ?? r2["Добавления аудио"] ?? r2["Добавили"] ?? "");
    const avg   = safeDiv_(spent, adds);

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml_(String(r2["Название объявления"] || r2["ID объявления"] || "—"))}</td>
      <td>${formatMoney2_(spent)}</td>
      <td>${formatInt_(num_(r2["Показы"]))}</td>
      <td>${formatInt_(num_(r2["Начали прослушивание"]))}</td>
      <td>${formatInt_(adds)}</td>
      <td>${avg == null ? "—" : formatMoney2_(avg)}</td>
    `;
    tb.appendChild(tr);
  }

  inner.appendChild(table);

  const shownFrom = total ? (start + 1) : 0;
  const shownTo   = total ? (start + slice.length) : 0;
  inner.insertAdjacentHTML("beforeend",
    `<div class="tableNote">Показаны строки ${shownFrom}–${shownTo} из ${total}.</div>`
  );

  return el;
}


  function renderDetailPage_(r, c){
    const el = sheet_(true);
    const inner = el.querySelector(".sheetInner");
    inner.insertAdjacentHTML("beforeend", `<div class="sheetTitle" style="font-size:28px">Детализация: ${escapeHtml_(c.name)}</div>`);
    const rows = (c.adsRows||[]);
    const table = document.createElement("table");
    table.className = "table";
    table.innerHTML = `
      <thead>
        <tr>
          <th>Объявление</th>
          <th>Потрачено</th>
          <th>Показы</th>
          <th>Прослуш.</th>
          <th>Добавл.</th>
          <th>Ср. цена добавл.</th>
        </tr>
      </thead>
      <tbody></tbody>
    `;
    const tb = table.querySelector("tbody");
    for (const row of rows){
      const r2 = normalizeRowKeys_(row||{});
      const spent = num_(r2["Потрачено всего, ₽"] ?? r2["Потрачено всего, Р"] ?? r2["Потрачено всего"] ?? "");
      const adds = num_(r2["Добавили аудио"] ?? r2["Добавления аудио"] ?? r2["Добавили"] ?? "");
      const avg = safeDiv_(spent, adds);
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${escapeHtml_(String(r2["Название объявления"]||r2["ID объявления"]||r2["ID объявления"]||"—"))}</td>
        <td>${formatMoney2_(spent)}</td>
        <td>${formatInt_(num_(r2["Показы"]))}</td>
        <td>${formatInt_(num_(r2["Начали прослушивание"]))}</td>
        <td>${formatInt_(adds)}</td>
        <td>${avg==null?"—":formatMoney2_(avg)}</td>
      `;
      tb.appendChild(tr);
    }
    inner.appendChild(table);
    inner.insertAdjacentHTML("beforeend", `<div style="margin-top:10px; font-size:12px; opacity:.75">Показаны первые ${rows.length} строк (MVP).</div>`);
    return el;
  }

  
  function renderDemoPage_(r, c){
    const el = sheet_(false);
    const inner = el.querySelector(".sheetInner");

    const title = `Демография: ${escapeHtml_(c.name)}`;
    inner.insertAdjacentHTML("beforeend", `
      <div class="slideHeader">
        <img class="lotusLogo" src="assets/logo_black.png" alt="Lotus Music"/>
        <div class="slideHeaderTitle">${title}</div>
        <img class="lotusLogo" src="assets/logo_black.png" alt="Lotus Music"/>
      </div>
      <div class="slideSpacer"></div>
    `);

    ensureDemoStorage_(c);
    const demo = getDemoRows_(c);

    // Aggregates
    const agg = aggregateDemo_(demo);
    const genderImpr   = agg.genderImpr || {};
    const genderClicks = agg.genderClicks || {};
    const ageGender    = agg.ageGender || {};

    // Spend is needed for "Цена за результат" as CPC (spent / clicks)
    const genderSpent = {};
    for (const row0 of demo){
      const row = normalizeRowKeys_(row0 || {});
      const g = normalizeGender_(row["Пол"]);
      const spent =
        numOrNull_(row["Расход"] ?? row["Расход, ₽"] ?? row["Потрачено всего, ₽"] ?? row["Потрачено, ₽"] ?? row["Потрачено всего"] ?? row["Потрачено"] ?? "") ?? 0;
      genderSpent[g] = (genderSpent[g]||0) + spent;
    }

    const genderCats = ["Мужчины","Женщины","Пол не указан"];
    const ageCats = ["12-17","18-24","25-34","35-44","45-54","55-64","65+"];

    const COLORS = {
      men:   "rgb(59,130,246)",   // blue
      women: "rgb(236,72,153)",   // pink
      none:  "rgb(109,40,217)",   // purple
      click: "rgb(132,204,22)"    // green
    };


    // --- Totals methodology (to match "Результаты продвижения") ---
    // M / Ж берём по строкам, а "Всего" стараемся брать из выгрузки "Объявления" (как на слайде результатов),
    // иначе — из строки "Итого" демографии. Остаток интерпретируем как "Пол не указан".
    const menI_raw = genderImpr["Мужчины"]||0;
    const womI_raw = genderImpr["Женщины"]||0;
    const menC_raw = genderClicks["Мужчины"]||0;
    const womC_raw = genderClicks["Женщины"]||0;

    const adsTot = (Array.isArray(c.adsRows) && c.adsRows.length) ? calcImprClicksFromAdsRows_(c.adsRows) : null;
    const demoTot = getDemoTotals_(c); // сумма "Итого" по всем демо-файлам

    let totalImpr = (adsTot && adsTot.impr) ? adsTot.impr : (demoTot && demoTot.impr) ? demoTot.impr : (menI_raw + womI_raw);
    let totalClicks = (adsTot && adsTot.clicks) ? adsTot.clicks : (demoTot && demoTot.clicks) ? demoTot.clicks : (menC_raw + womC_raw);

    const sexSumImpr = menI_raw + womI_raw;
    const sexSumClicks = menC_raw + womC_raw;

    // Если общий итог меньше М+Ж — не добавляем "Пол не указан" и считаем "Всего" = М+Ж.
    if (totalImpr < sexSumImpr) totalImpr = sexSumImpr;
    if (totalClicks < sexSumClicks) totalClicks = sexSumClicks;

    const noneI = Math.max(0, totalImpr - sexSumImpr);
    const noneC = Math.max(0, totalClicks - sexSumClicks);

    // Принудительно подставляем рассчитанные значения "Пол не указан" (важно для диаграммы пола и процентов)
    genderImpr["Пол не указан"] = noneI;
    genderClicks["Пол не указан"] = noneC;

    // Эффективные значения для таблицы
    const menI = menI_raw, womI = womI_raw;
    const menC = menC_raw, womC = womC_raw;
     const capFirst_ = (s)=>{
  const str = String(s || "");
  if (!str) return str;
  if (str.length === 1) return str.toUpperCase(); // М / Ж
  return str.slice(0,1).toUpperCase() + str.slice(1).toLowerCase();
};

// Для строк вида "Нет пола - показы" или "М - показы"
const capDashParts_ = (s)=>{
  const str = String(s || "");
  return str
    .split(/\s*-\s*/g)
    .map(part => capFirst_(part.trim()))
    .join(" - ");
};

    // Capitalize words inside demoGrid (первая буква заглавная, остальные строчные).
    // Однобуквенные токены (М/Ж) и токены с цифрами (25-34) не трогаем.
    const capWords_ = (s)=>{
      return String(s||"").replace(/[A-Za-zА-Яа-яЁё]+/g, (w)=>{
        if (w.length === 1) return w.toUpperCase();
        return w.slice(0,1).toUpperCase() + w.slice(1).toLowerCase();
      });
    };

    // --- Layout root
    const grid = document.createElement("div");
    grid.className = "demoGrid";
    inner.appendChild(grid);

    const card = (cls)=>{
      const d = document.createElement("div");
      d.className = "demoCard " + (cls||"");
      return d;
    };

    // 1) Legend card
    const legendCard = card("demoLegendCard");
    legendCard.innerHTML = `
      <div class="demoLegendList" role="list">
        <div class="demoLegendItem" role="listitem"><span class="demoSw" style="background:${COLORS.men}"></span><span>${capDashParts_("М • показы")}</span></div>
        <div class="demoLegendItem" role="listitem"><span class="demoSw" style="background:${COLORS.women}"></span><span>${capDashParts_("Ж • показы")}</span></div>
        <div class="demoLegendItem" role="listitem"><span class="demoSw" style="background:${COLORS.none}"></span><span>${capDashParts_("Нет пола • показы")}</span></div>
        <div class="demoLegendItem" role="listitem"><span class="demoSw" style="background:${COLORS.click}"></span><span>${capDashParts_("Клики")}</span></div>
      </div>
    `;
    grid.appendChild(legendCard);

    // 2) Gender chart card
    const genderCard = card("demoGenderCard");
    const genderChart = svgGroupedBarChart_(
      capFirst_("Распределение по полу"),
      genderCats,
      [
         { name: capFirst_("Показы"), data: Object.fromEntries(genderCats.map(k=>[k, genderImpr[k]||0])) },
         { name: capFirst_("Клики"),  data: Object.fromEntries(genderCats.map(k=>[k, genderClicks[k]||0])) }
      ],
      {
        showXAxisLabels: false,
        overlayPairs: true,
        pairSize: 2,
        showValues: false,
        showLegend: false,
        stretch: true,
        xLabelFontSize: 18,
        wrapXLabels: false,
        xLabelLineHeight: 16,
        xLabelPadBottom: 6,
        xLabelLetterSpacing: 6,
        yPadRatio: 1.25,

        categoryColors: {
          "Мужчины": COLORS.men,
          "Женщины": COLORS.women,
          "Пол не указан": COLORS.none
        },
        overlayColor: COLORS.click,
        overlayWidthRatio: 1,
        seriesColors: [COLORS.men, COLORS.click], // legacy / fallback
        pairGap: 18,
        rx: 10,
        dualAxis: true,
        stackOverlayOnBase: true,
        axisTransparent: true
      }
    );
    genderCard.appendChild(genderChart);
    grid.appendChild(genderCard);

    // 3) Table card (М / Ж / Всего)
    const tableCard = card("demoTableCard");

        const valPct = (v, total)=>{
      if (!total) return `${formatInt_(v)} (0%)`;
      const p = Math.round((v/total)*100);
      return `${formatInt_(v)} (${p}%)`;
    };

    const cost = (spent, clicks)=>{
      if (!clicks) return "—";
      const v = spent / clicks;
      return `${formatMoney2_(v)}`;
    };

    const menSpent = genderSpent["Мужчины"]||0;
    const womSpent = genderSpent["Женщины"]||0;
    const noneSpent = genderSpent["Пол не указан"]||0;
    tableCard.innerHTML = `
      <table class="demoMiniTable">
        <thead>
          <tr>
            <th></th>
            <th>М</th>
            <th>Ж</th>
            <th>Всего</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>Показы</td>
            <td>${escapeHtml_(valPct(menI, totalImpr))}</td>
            <td>${escapeHtml_(valPct(womI, totalImpr))}</td>
            <td>${escapeHtml_(valPct(totalImpr, totalImpr))}</td>
          </tr>
          <tr>
            <td>Клики</td>
            <td>${escapeHtml_(valPct(menC, totalClicks))}</td>
            <td>${escapeHtml_(valPct(womC, totalClicks))}</td>
            <td>${escapeHtml_(valPct(totalClicks, totalClicks))}</td>
          </tr>
          <tr>
            <td>Цена за<br/>результат, ₽</td>
            <td>${escapeHtml_(cost(menSpent, menC))}</td>
            <td>${escapeHtml_(cost(womSpent, womC))}</td>
            <td>${escapeHtml_(cost(menSpent+womSpent+noneSpent, totalClicks))}</td>
          </tr>
        </tbody>
      </table>
    `;
    grid.appendChild(tableCard);

    // 4) Top-3 segments by clicks (exclude "Пол не указан")
    const topCard = card("demoTopCard");

    const segments = [];
    let totalKnownClicks = 0;
    for (const a of ageCats){
      const gObj = ageGender[a] || {};
      const mc = (gObj["Мужчины"]?.clicks||0);
      const wc = (gObj["Женщины"]?.clicks||0);
      totalKnownClicks += (mc + wc);
      segments.push({g:"М", age:a, clicks:mc});
      segments.push({g:"Ж", age:a, clicks:wc});
    }
    const top = segments
      .filter(x=>x.clicks>0)
      .sort((a,b)=>b.clicks-a.clicks)
      .slice(0,3);

    const pctSeg = (v)=> totalKnownClicks ? Math.round((v/totalKnownClicks)*100) : 0;

    topCard.innerHTML = `
      <div class="demoTopTitle">${capFirst_("ТОП-3 сегмента")}<br/><span class="demoTopSub">• ${capFirst_("клики")}</span></div>
      <div class="demoTopList">
        ${[0,1,2].map(i=>{
          const it = top[i];
          if (!it) return `
            <div class="demoTopRow">
              <span class="demoRank">${i+1}</span>
              <span class="demoTopText">—</span>
            </div>`;
          return `
            <div class="demoTopRow">
              <span class="demoRank">${i+1}</span>
              <span class="demoTopText">${capWords_(it.g)} • ${escapeHtml_(it.age)} (${pctSeg(it.clicks)}%)</span>
            </div>`;
        }).join("")}
      </div>
    `;
    grid.appendChild(topCard);

    // 5) Age chart (pairs: men, women; overlay = clicks)
    const ageCard = card("demoAgeCard");

    const ageMenImpr = {}, ageMenClicks = {}, ageWomenImpr = {}, ageWomenClicks = {};
    for (const a of ageCats){
      const gObj = ageGender[a] || {};
      ageMenImpr[a] = (gObj["Мужчины"]?.impr||0);
      ageMenClicks[a] = (gObj["Мужчины"]?.clicks||0);
      ageWomenImpr[a] = (gObj["Женщины"]?.impr||0);
      ageWomenClicks[a] = (gObj["Женщины"]?.clicks||0);
    }

    const ageChart = svgGroupedBarChart_(
      capFirst_("Распределение по возрасту"),
      ageCats,
      [
         { name: capFirst_("М • показы"), data: ageMenImpr },
         { name: capFirst_("М • клики"),  data: ageMenClicks },
         { name: capFirst_("Ж • показы"), data: ageWomenImpr },
         { name: capFirst_("Ж • клики"),  data: ageWomenClicks }
      ],
      {
        overlayPairs: true,
        pairSize: 2,
        showValues: false,
        showLegend: false,
        stretch: true,
        seriesColors: [COLORS.men, COLORS.click, COLORS.women, COLORS.click],
        overlayWidthRatio: 1,
        pairGap: 1,
        rx: 10,
        dualAxis: true,
        stackOverlayOnBase: true,
        axisTransparent: true,
        yPadRatio: 1.18,
      }
    );

    ageCard.appendChild(ageChart);
    grid.appendChild(ageCard);

    return el;
  }


  function renderStreamsPage_(r){
  // Как и на остальных слайдах: шапка = [лого] ЗАГОЛОВОК [лого]
  const el = sheet_(false);
  const inner = el.querySelector(".sheetInner");

  inner.insertAdjacentHTML("beforeend", `
    <div class="slideHeader">
      <img class="lotusLogo" src="assets/logo_black.png" alt="Lotus Music"/>
      <div class="slideHeaderTitle">Прослушивания VK + BOOM</div>
      <img class="lotusLogo" src="assets/logo_black.png" alt="Lotus Music"/>
    </div>
  `);

  inner.insertAdjacentHTML("beforeend", `<div class="slideSpacer"></div>`);

  const rows = (r.streams || []);
  const chart = svgLineChart_("VK + BOOM стримы по дням", rows.map(x=>({ x:x.dateISO, y:x.vk })));
  inner.appendChild(chart);

  // ===== Таблица (транспонированная): колонки = даты, строки = прослушивания =====
  const tab = document.createElement("table");
  tab.className = "table streamsTable";

  const dates = rows.map(s => escapeHtml_(fmtDateShort_(s.dateISO)));
  const vals  = rows.map(s => formatInt_(s.vk));

  // THEAD: первая ячейка — подпись строки, дальше даты
  tab.innerHTML = `
    <thead>
      <tr>
        <th>Дата</th>
        ${dates.map(d => `<th>${d}</th>`).join("")}
      </tr>
    </thead>
    <tbody>
      <tr>
        <th>VK + BOOM</th>
        ${vals.map(v => `<td>${v}</td>`).join("")}
      </tr>
    </tbody>
  `;

  inner.appendChild(tab);
  return el;
}


  function renderSimpleTable_(headers, rows){
    const table = document.createElement("table");
    table.className = "table";
    const thead = document.createElement("thead");
    const trh = document.createElement("tr");
    for (const h of headers){
      const th = document.createElement("th");
      th.textContent = h;
      trh.appendChild(th);
    }
    thead.appendChild(trh);
    const tbody = document.createElement("tbody");
    for (const row of rows){
      const tr = document.createElement("tr");
      for (const cell of row){
        const td = document.createElement("td");
        td.textContent = cell;
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    }
    table.appendChild(thead);
    table.appendChild(tbody);
    return table;
  }

  /** -----------------------------
   *  Parsers (xlsx)
   * -----------------------------*/
// --- Helpers: normalize column names (trim, NBSP) ---
function normKey_(s){
  return String(s ?? "")
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}
function normalizeRowKeys_(row){
  const out = {};
  for (const k in (row || {})){
    out[normKey_(k)] = row[k];
  }

  // ---- Canonicalize common VK export column variants (so cloud snapshots re-render correctly) ----
  // spent
  if (out["Потрачено, ₽"] != null && out["Потрачено всего, ₽"] == null) out["Потрачено всего, ₽"] = out["Потрачено, ₽"];
  if (out["Потрачено, Р"] != null && out["Потрачено всего, Р"] == null) out["Потрачено всего, Р"] = out["Потрачено, Р"];
  if (out["Потрачено"] != null && out["Потрачено всего"] == null) out["Потрачено всего"] = out["Потрачено"];

  // adds
  const addsCandidates = ["Добавили", "Добавили в аудио", "Добавили аудио", "Добавили в библиотеку", "Добавили в медиатеку", "Добавления в библиотеку", "Добавления аудио"];
  if (out["Добавления"] == null){
    for (const k of addsCandidates){
      if (out[k] != null) { out["Добавления"] = out[k]; break; }
    }
  }

  // listens from ads (not streams chart)
  const listensCandidates = ["Прослушивания", "Прослушивания VK+BOOM", "Прослушивания VK", "Прослушивания BOOM", "Прослушивания (VK+BOOM)"];
  if (out["Прослушивания"] == null){
    for (const k of listensCandidates){
      if (out[k] != null) { out["Прослушивания"] = out[k]; break; }
    }
  }

  return out;
}
async function readXlsx_(file){
  const buf = await file.arrayBuffer();
  return XLSX.read(buf, {type:"array"});
}

  async function parseStreamsXlsx_(file){
    const wb = await readXlsx_(file);
    // Prefer sheet "Chart data" but fallback to first
    const name = wb.SheetNames.includes("Chart data") ? "Chart data" : wb.SheetNames[0];
    const ws = wb.Sheets[name];
    const json = XLSX.utils.sheet_to_json(ws, {defval:""});
    // normalize: need columns "Дата" and "vk"
    // allow "date" variants
    const out = [];
    for (const row of json){
      const dateRaw = row["Дата"] ?? row["date"] ?? row["Date"] ?? row["DATE"];
      const vkRaw = row["vk"] ?? row["VK"] ?? row["Vk"];
      if (!dateRaw) continue;
      const dateISO = toISODate_(dateRaw);
      const vk = num_(vkRaw);
      if (!dateISO) continue;
      out.push({dateISO, vk});
    }
    // sort by date
    out.sort((a,b)=> (a.dateISO||"").localeCompare(b.dateISO||""));
    if (!out.length) throw new Error("No streams rows");
    return out;
  }

  async function parseAdsXlsx_(file){
    const wb = await readXlsx_(file);
    const sheetName = wb.SheetNames.includes("Объявления") ? "Объявления" : wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const jsonRaw = XLSX.utils.sheet_to_json(ws, {defval:""});
    if (!jsonRaw.length) throw new Error("empty ads");

    // Normalize column names to avoid issues with NBSP / trailing spaces / "₽ " etc.
    const json = jsonRaw.map(normalizeRowKeys_).map(r=>{
      // aliases for common header variants
      if (r["Потрачено всего, ₽"]==null && r["Потрачено всего, Р"]!=null) r["Потрачено всего, ₽"] = r["Потрачено всего, Р"];
      if (r["Цена за результат, ₽"]==null && r["Цена за результат, Р"]!=null) r["Цена за результат, ₽"] = r["Цена за результат, Р"];
      return r;
    });

    // minimal sanity check
    const sample = json[0] || {};
    if (!("Показы" in sample) && !("Потрачено всего, ₽" in sample) && !("Потрачено всего, Р" in sample)){
      console.warn("Ads columns unexpected:", Object.keys(sample));
    }
    return json;
  }

  async function parseDemoXlsx_(file){
  const wb = await readXlsx_(file);
  const ws = wb.Sheets[wb.SheetNames[0]];

  // Пропускаем первые 5 строк (служебные), 6-я строка становится заголовками
  const jsonRaw = XLSX.utils.sheet_to_json(ws, { defval:"", range: 5 });
  const json = jsonRaw.map(normalizeRowKeys_).map(r=>{
      // aliases for common header variants
      if (r["Потрачено всего, ₽"]==null && r["Потрачено всего, Р"]!=null) r["Потрачено всего, ₽"] = r["Потрачено всего, Р"];
      if (r["Цена за результат, ₽"]==null && r["Цена за результат, Р"]!=null) r["Цена за результат, ₽"] = r["Цена за результат, Р"];
      return r;
    });

  if (!json.length) throw new Error("empty demo");

  // 1) Extract totals from the "Итого" row (needed to match "Результаты продвижения")
  let totals = null;
  for (const row0 of json){
    const row = normalizeRowKeys_(row0 || {});
    const age = String(row["Возраст"] ?? row["Age"] ?? "").trim().toLowerCase();
    const gender = String(row["Пол"] ?? row["Gender"] ?? "").trim().toLowerCase();
    if (age.includes("итого") || gender.includes("итого")){
      totals = {
        impr: num_(row["Показы"]),
        clicks: num_(row["Клики"])
      };
      break;
    }
  }

  // 2) Keep only rows having Age+Gender (normal buckets)
  const out = [];
  for (const row0 of json){
    const row = normalizeRowKeys_(row0 || {});
    const age = row["Возраст"] ?? row["Age"] ?? "";
    const gender = row["Пол"] ?? row["Gender"] ?? "";
    const aLow = String(age||"").trim().toLowerCase();
    const gLow = String(gender||"").trim().toLowerCase();
    // skip totals row
    if (aLow.includes("итого") || gLow.includes("итого")) continue;
    if (!age || !gender) continue;
    out.push(row);
  }

  return { rows: (out.length ? out : json), totals };
}

  /** -----------------------------
   *  Demo aggregation (MVP) (MVP)
   * -----------------------------*/
  function normalizeGender_(g){
    const s = String(g||"").trim();
    const low = s.toLowerCase();

    // Standard buckets used across charts/tables
    if (!s) return "Пол не указан";
    if (low.includes("муж")) return "Мужчины";
    if (low.includes("жен")) return "Женщины";

    // Normalize all "unknown/unspecified" variants into one bucket
    // Examples from exports: "Пол не указан", "Пол не определен", "Не указан", "Нет пола"
    if (
      low.includes("не указан") ||
      low.includes("не указ") ||
      low.includes("не определ") ||
      low.includes("нет пола") ||
      (low.includes("пол") && low.includes("не"))
    ) return "Пол не указан";

    return s;
  }
  function normalizeAge_(a){
    const s = String(a||"").trim();
    // keep as is to match VK buckets
    return s || "—";
  }
  function aggregateDemo_(rows){
    const genderImpr = {};
    const genderClicks = {};
    const genderCost = {}; // avg cost per gender (simple mean of non-empty)
    const genderCostAgg = {}; // {sum, n}

    // age -> gender -> {impr, clicks}
    const ageGender = {};

    for (const row0 of rows){
      const row = normalizeRowKeys_(row0 || {});
      const g = normalizeGender_(row["Пол"]);
      const a = normalizeAge_(row["Возраст"]);
      const impr = num_(row["Показы"]);
      const clicks = num_(row["Клики"]);
      const cost = numOrNull_(row["Цена за результат, ₽"] ?? row["Цена за результат, Р"] ?? row["Цена за результат"] ?? "");

      genderImpr[g] = (genderImpr[g]||0) + impr;
      genderClicks[g] = (genderClicks[g]||0) + clicks;

      ageGender[a] = ageGender[a] || {};
      ageGender[a][g] = ageGender[a][g] || {impr:0, clicks:0};
      ageGender[a][g].impr += impr;
      ageGender[a][g].clicks += clicks;

      if (cost!=null){
        genderCostAgg[g] = genderCostAgg[g] || {sum:0,n:0};
        genderCostAgg[g].sum += cost;
        genderCostAgg[g].n += 1;
      }
    }

    for (const k of Object.keys(genderCostAgg)){
      const it = genderCostAgg[k];
      genderCost[k] = it.n ? (it.sum/it.n) : null;
    }

    return { genderImpr, genderClicks, genderCost, ageGender };
  }

  /** -----------------------------
   *  Simple SVG charts for print
   * -----------------------------*/
  function svgBarChart_(title, obj){
    const keys = Object.keys(obj);
    const values = keys.map(k=>obj[k]||0);
    const max = Math.max(1, ...values);
    const w = 1000, h = 180, pad = 30, barGap = 10;
    const barW = (w - pad*2 - barGap*(keys.length-1)) / Math.max(1, keys.length);

    const svg = document.createElement("div");
    svg.className = "chartBox";
    const header = document.createElement("div");
    header.className = "chartTitle";
    header.textContent = title;
    svg.appendChild(header);

    const svgEl = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svgEl.setAttribute("viewBox", `0 0 ${w} ${h}`);
    svgEl.classList.add("svgChart");

    // axes baseline
    const baseY = h - pad;
    const axis = document.createElementNS(svgEl.namespaceURI, "line");
    axis.setAttribute("x1", pad);
    axis.setAttribute("y1", baseY);
    axis.setAttribute("x2", w-pad);
    axis.setAttribute("y2", baseY);
    axis.setAttribute("stroke", "rgba(0,0,0,.18)");
    axis.setAttribute("stroke-width", "2");
    svgEl.appendChild(axis);

    keys.forEach((k, i)=>{
      const v = obj[k]||0;
      const bh = (h - pad*2) * (v/max);
      const x = pad + i*(barW+barGap);
      const y = baseY - bh;

      const rect = document.createElementNS(svgEl.namespaceURI, "rect");
      rect.setAttribute("x", x);
      rect.setAttribute("y", y);
      rect.setAttribute("width", barW);
      rect.setAttribute("height", bh);
      rect.setAttribute("rx", "10");
      rect.setAttribute("fill", "rgba(124,92,255,.55)");
      svgEl.appendChild(rect);

      // x label (bigger + optional wrap for long labels like "Пол не указан")
const xFs = (opts.xLabelFontSize ?? 16);
const xPadBottom = (opts.xLabelPadBottom ?? 6);
const lineH = (opts.xLabelLineHeight ?? Math.round(xFs * 0.95));
const letterSp = (opts.xLabelLetterSpacing ?? 0);

const label = document.createElementNS(svgEl.namespaceURI, "text");
label.setAttribute("x", gx + groupW / 2);
label.setAttribute("y", h - xPadBottom);
label.setAttribute("text-anchor", "middle");
label.setAttribute("font-size", String(xFs));
label.setAttribute("fill", "rgba(0,0,0,.75)");
if (letterSp) label.setAttribute("letter-spacing", String(letterSp));

// Если строка длинная и включён wrap — переносим на 2 строки, чтобы не наезжало на соседей
const wrap = (opts.wrapXLabels === true);
if (wrap && typeof cat === "string" && cat.includes(" ")) {
  const parts = cat.split(/\s+/);
  // 2 строки: первая часть/слово, вторая — остальное
  const line1 = parts[0];
  const line2 = parts.slice(1).join(" ");

  // поднимаем первую строку выше, вторую оставляем у baseline
  const t1 = document.createElementNS(svgEl.namespaceURI, "tspan");
  t1.setAttribute("x", gx + groupW / 2);
  t1.setAttribute("dy", String(-Math.round(lineH * 0.55)));
  t1.textContent = line1;

  const t2 = document.createElementNS(svgEl.namespaceURI, "tspan");
  t2.setAttribute("x", gx + groupW / 2);
  t2.setAttribute("dy", String(lineH));
  t2.textContent = line2;

  label.appendChild(t1);
  label.appendChild(t2);
} else {
  label.textContent = cat;
}

svgEl.appendChild(label);


      const val = document.createElementNS(svgEl.namespaceURI, "text");
      val.setAttribute("x", x + barW/2);
      val.setAttribute("y", y - 8);
      val.setAttribute("text-anchor", "middle");
      val.setAttribute("font-size", "12");
      val.setAttribute("font-weight", "900");
      val.setAttribute("fill", "rgba(0,0,0,.75)");
      val.textContent = Math.round(v).toLocaleString("ru-RU");
      svgEl.appendChild(val);
    });

    svg.appendChild(svgEl);
    return svg;
  }


  function svgGroupedBarChart_(title, categories, series, opts = {}) {
  const w = (opts.width ?? 1000);
  const h = (opts.height ?? 220);
  const pad = (opts.pad ?? 30);
  const svgBox = document.createElement("div");
  svgBox.className = "chartBox";

  const showLegend = (opts.showLegend !== false);

  const legend = showLegend ? series.map((s, i) => {
      const col = (opts.seriesColors && opts.seriesColors[i]) ? opts.seriesColors[i] : `rgba(122,101,255,${0.25 + i * 0.18})`;
      return `
        <span style="display:inline-flex; align-items:center; gap:6px; margin-right:14px">
          <span style="width:10px; height:10px; border-radius:3px; background:${col}; display:inline-block"></span>
          <span>${escapeHtml_(s.name)}</span>
        </span>
      `;
    }).join("") : "";

  svgBox.innerHTML = `
    <div class="chartTitle" style="display:flex; align-items:flex-end; justify-content:space-between; gap:12px">
      <span>${escapeHtml_(title)}</span>
      ${showLegend ? `<span style="font-size:12px; opacity:.75; white-space:nowrap">${legend}</span>` : ``}
    </div>
  `;

  const svgEl = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svgEl.setAttribute("viewBox", `0 0 ${w} ${h}`);
  if (opts.stretch) svgEl.setAttribute("preserveAspectRatio", "none");
  svgEl.classList.add("svgChart");

  const showValues = (opts.showValues !== false);
  const overlayPairs = !!opts.overlayPairs;
  const pairSize = overlayPairs ? (opts.pairSize || 2) : 1;
  const pairCount = overlayPairs ? Math.ceil(series.length / pairSize) : series.length;

  // max (with optional dual-axis scaling for overlay pairs)
  let max = 1;
  let maxBase = 1;
  let maxOver = 1;

  if (overlayPairs && opts.dualAxis) {
    const valsBase = [];
    const valsOver = [];
    for (const cat of categories) {
      for (let p = 0; p < pairCount; p++) {
        const baseSeries = series[p * pairSize + 0];
        const overSeries = series[p * pairSize + 1];
        valsBase.push(Number((baseSeries?.data && baseSeries.data[cat]) || 0));
        valsOver.push(Number((overSeries?.data && overSeries.data[cat]) || 0));
      }
    }
    maxBase = Math.max(1, ...valsBase);
    maxOver = Math.max(1, ...valsOver);
    max = Math.max(maxBase, maxOver);
  } else {
    const values = [];
    for (const cat of categories) {
      for (const s of series) values.push(Number((s.data && s.data[cat]) || 0));
    }
    max = Math.max(1, ...values);
    maxBase = maxOver = max;
  }

   // === Headroom по Y (чтобы столбцы не упирались в верх svg) ===
const yPadRatio = (opts.yPadRatio ?? 1.15); // дефолт 15%

max = Math.max(1, max * yPadRatio);

if (overlayPairs && opts.dualAxis) {
  maxBase = Math.max(1, maxBase * yPadRatio);
  maxOver = Math.max(1, maxOver * yPadRatio);
} else {
  maxBase = maxOver = max;
}

     // Если мы именно СТЭКАЕМ overlay на base — нужен общий максимум суммы,
// иначе bh1 и bh2 по разным шкалам дадут переполнение вверх.
let maxStack = null;

if (overlayPairs && opts.stackOverlayOnBase) {
  const stackVals = [];
  for (const cat of categories) {
    for (let p = 0; p < pairCount; p++) {
      const baseSeries = series[p * pairSize + 0];
      const overSeries = series[p * pairSize + 1];
      const v1 = Number((baseSeries?.data && baseSeries.data[cat]) || 0);
      const v2 = Number((overSeries?.data && overSeries.data[cat]) || 0);
      stackVals.push(v1 + v2);
    }
  }
  maxStack = Math.max(1, ...stackVals);
  maxStack = Math.max(1, maxStack * yPadRatio); // тот же headroom
}


  const baseY = h - pad;

  // baseline
  const base = document.createElementNS(svgEl.namespaceURI, "line");
  base.setAttribute("x1", pad);
  base.setAttribute("x2", w - pad);
  base.setAttribute("y1", baseY);
  base.setAttribute("y2", baseY);
  base.setAttribute("stroke", (opts.axisTransparent ? "rgba(0,0,0,0)" : "rgba(0,0,0,.18)"));
  base.setAttribute("stroke-width", "2");
  svgEl.appendChild(base);

  // sizes
  const groupW = (w - pad * 2) / Math.max(1, categories.length);
  const pairGap = opts.pairGap ?? 14;                // больше воздуха => “уже” визуально
  const pairW = (groupW - pairGap * (pairCount - 1)) / Math.max(1, pairCount);
  const barWidthRatio = opts.barWidthRatio ?? 0.8;  // столбцы уже
  const overlayWidthRatio = opts.overlayWidthRatio ?? 0.55;

  const rx = opts.rx ?? 10;

  const toRGBA_ = (col, a)=>{
    if (!col) return `rgba(0,0,0,${a})`;
    if (col.startsWith("rgba(")) {
      // rgba(r,g,b,alpha)
      return col.replace(/rgba\(([^,]+),([^,]+),([^,]+),[^\)]+\)/, (m,r,g,b)=>`rgba(${r.trim()},${g.trim()},${b.trim()},${a})`);
    }
    if (col.startsWith("rgb(")) {
      return col.replace(/rgb\(([^\)]+)\)/, (m,inside)=>`rgba(${inside},${a})`);
    }
    if (col[0]==="#" && col.length===7){
      const r=parseInt(col.slice(1,3),16), g=parseInt(col.slice(3,5),16), b=parseInt(col.slice(5,7),16);
      return `rgba(${r},${g},${b},${a})`;
    }
    return col; // fallback
  };

  const colorFor_ = (si, cat, isOverlay)=>{
    // overlay color override (e.g., clicks always green)
    if (isOverlay && opts.overlayColor){
      return toRGBA_(opts.overlayColor, 0.88);
    }
    // 1) категорные цвета (для графика по полу, где цвет = пол)
    if (opts.categoryColors && opts.categoryColors[cat]){
      const base = opts.categoryColors[cat];
      return toRGBA_(base, isOverlay ? 0.55 : 0.90);
    }
    // 2) цвета по сериям (для графика по возрасту: М/Ж + показы/клики)
    if (opts.seriesColors && opts.seriesColors[si]){
      const base = opts.seriesColors[si];
      // если это уже rgba/rgb/hex — пусть будет, но при overlay чуть “темнее”
      return isOverlay ? toRGBA_(base, 0.78) : base;
    }
    // 3) дефолт
    const alpha = isOverlay ? (0.78 - si * 0.06) : (0.38 - si * 0.04);
    return `rgba(122,101,255,${Math.max(0.18, alpha)})`;
  };

  categories.forEach((cat, ci) => {
    const gx = pad + ci * groupW;

    if (opts.showXAxisLabels !== false) {
    // x label
    const label = document.createElementNS(svgEl.namespaceURI, "text");
    label.setAttribute("x", gx + groupW / 2);
    label.setAttribute("y", h - 8);
    label.setAttribute("text-anchor", "middle");
    label.setAttribute("font-size", "12");
    label.setAttribute("fill", "rgba(0,0,0,.75)");
    label.textContent = cat;
    svgEl.appendChild(label);
   } 

    for (let p = 0; p < pairCount; p++) {
      const pxSlot = gx + p * (pairW + pairGap);
      const slotCenter = pxSlot + pairW/2;

      if (overlayPairs) {
        // пара: series[p*pairSize + 0] = “основание” (показы), series[+1] = overlay (клики)
        const baseSeries = series[p * pairSize + 0];
        const overSeries = series[p * pairSize + 1];

        const v1 = Number((baseSeries?.data && baseSeries.data[cat]) || 0);
        const v2 = Number((overSeries?.data && overSeries.data[cat]) || 0);

        const denom = (maxStack != null)
           ? maxStack
           : (opts.dualAxis ? maxBase : max);

        const denom2 = (maxStack != null)
           ? maxStack
           : (opts.dualAxis ? maxOver : max);
        const bh1 = (h - pad * 2) * (v1 / denom);
        const bh2 = (h - pad * 2) * (v2 / denom2);


        const mainW = pairW * barWidthRatio;
        const overW = mainW * overlayWidthRatio;

        const x1 = slotCenter - mainW/2;
        const x2 = slotCenter - overW/2;

        const y1 = baseY - bh1;
        const y2 = (opts.stackOverlayOnBase ? (y1 - bh2) : (baseY - bh2));

        const rect1 = document.createElementNS(svgEl.namespaceURI, "rect");
        rect1.setAttribute("x", x1);
        rect1.setAttribute("y", y1);
        rect1.setAttribute("width", mainW);
        rect1.setAttribute("height", bh1);
        rect1.setAttribute("rx", String(rx));
        rect1.setAttribute("fill", colorFor_(p*pairSize+0, cat, false));
        svgEl.appendChild(rect1);

        const rect2 = document.createElementNS(svgEl.namespaceURI, "rect");
        rect2.setAttribute("x", x2);
        rect2.setAttribute("y", y2);
        rect2.setAttribute("width", overW);
        rect2.setAttribute("height", bh2);
        rect2.setAttribute("rx", String(rx));
        rect2.setAttribute("fill", colorFor_(p*pairSize+1, cat, true));
        svgEl.appendChild(rect2);

        if (showValues && v1 > 0) {
          const val = document.createElementNS(svgEl.namespaceURI, "text");
          val.setAttribute("x", slotCenter);
          val.setAttribute("y", y1 - 8);
          val.setAttribute("text-anchor", "middle");
          val.setAttribute("font-size", "12");
          val.setAttribute("fill", "rgba(0,0,0,.75)");
          val.textContent = Math.round(v1).toLocaleString("ru-RU");
          svgEl.appendChild(val);
        }
      } else {
        const s = series[p];
        const v = Number((s.data && s.data[cat]) || 0);
        const bh = (h - pad * 2) * (v / max);
        const y = baseY - bh;

        const mainW = pairW * barWidthRatio;
        const x = slotCenter - mainW/2;

        const rect = document.createElementNS(svgEl.namespaceURI, "rect");
        rect.setAttribute("x", x);
        rect.setAttribute("y", y);
        rect.setAttribute("width", mainW);
        rect.setAttribute("height", bh);
        rect.setAttribute("rx", String(rx));
        rect.setAttribute("fill", colorFor_(p, cat, false));
        svgEl.appendChild(rect);

        if (showValues && v > 0) {
          const val = document.createElementNS(svgEl.namespaceURI, "text");
          val.setAttribute("x", slotCenter);
          val.setAttribute("y", y - 8);
          val.setAttribute("text-anchor", "middle");
          val.setAttribute("font-size", "12");
          val.setAttribute("fill", "rgba(0,0,0,.75)");
          val.textContent = Math.round(v).toLocaleString("ru-RU");
          svgEl.appendChild(val);
        }
      }
    }
  });

  svgBox.appendChild(svgEl);
  return svgBox;
}

  function svgLineChart_(title, points){
    const w = 1000, h = 220, pad = 30;
    const svg = document.createElement("div");
    svg.className = "chartBox";
    svg.innerHTML = `<div class="chartTitle">${escapeHtml_(title)}</div>`;
    const svgEl = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svgEl.setAttribute("viewBox", `0 0 ${w} ${h}`);
    svgEl.classList.add("svgChart");

    const vals = points.map(p=>Number(p.y||0));
    const max = Math.max(1, ...vals);
    const min = 0;

    const n = points.length || 1;
    const xStep = (w - pad*2) / Math.max(1, n-1);

    const toX = (i)=> pad + i*xStep;
    const toY = (v)=> (h-pad) - ((v-min)/(max-min||1))*(h-pad*2);

    // baseline
    const base = document.createElementNS(svgEl.namespaceURI, "line");
    base.setAttribute("x1", pad);
    base.setAttribute("y1", h-pad);
    base.setAttribute("x2", w-pad);
    base.setAttribute("y2", h-pad);
    base.setAttribute("stroke", "rgba(0,0,0,.18)");
    base.setAttribute("stroke-width", "2");
    svgEl.appendChild(base);

    // path
    let d = "";
    points.forEach((p,i)=>{
      const x = toX(i);
      const y = toY(Number(p.y||0));
      d += (i===0 ? `M ${x} ${y}` : ` L ${x} ${y}`);
    });
     
    const path = document.createElementNS(svgEl.namespaceURI, "path");
    path.setAttribute("d", d);
    path.setAttribute("fill", "none");
    path.setAttribute("stroke", "rgba(45,212,191,.85)");
    path.setAttribute("stroke-width", "4");
    path.setAttribute("stroke-linecap","round");
    path.setAttribute("stroke-linejoin","round");
    svgEl.appendChild(path);

    // dots (limit)
    points.slice(0, Math.min(30, points.length)).forEach((p,i)=>{
      const x = toX(i);
      const y = toY(Number(p.y||0));
      const dot = document.createElementNS(svgEl.namespaceURI, "circle");
      dot.setAttribute("cx", x);
      dot.setAttribute("cy", y);
      dot.setAttribute("r", "4");
      dot.setAttribute("fill", "rgba(45,212,191,.95)");
      svgEl.appendChild(dot);
    });

    
    // X-axis labels (dd.mm)
    const step = Math.max(1, Math.ceil((n-1)/10));
    for (let i=0;i<n;i++){
      if (i % step !== 0 && i !== n-1) continue;
      const px = pad + i*xStep;
      const tick = document.createElementNS("http://www.w3.org/2000/svg","line");
       tick.setAttribute("x1", px);
       tick.setAttribute("x2", px);
       tick.setAttribute("y1", h-pad);
       tick.setAttribute("y2", h-pad+6);
       tick.setAttribute("stroke", "rgba(0,0,0,.25)");
       tick.setAttribute("stroke-width", "2");
       svgEl.appendChild(tick);
      const t = document.createElementNS("http://www.w3.org/2000/svg","text");
      t.setAttribute("x", px);
      t.setAttribute("y", h - 8);
      t.setAttribute("text-anchor", "middle");
      t.setAttribute("font-size", "12");
      t.setAttribute("fill", "rgba(0,0,0,.65)");
      t.textContent = fmtDateShort_(points[i].x);
      svgEl.appendChild(t);
    }
svg.appendChild(svgEl);
    return svg;
  }

  /** -----------------------------
   *  UI helpers
   * -----------------------------*/
  function renderMiniPreview_(tableEl, rows, cols){
    if (!tableEl) return;
    const keys = cols || (rows[0] ? Object.keys(rows[0]) : []);
    const thead = document.createElement("thead");
    const trh = document.createElement("tr");
    for (const k of keys){
      const th = document.createElement("th");
      th.textContent = k;
      trh.appendChild(th);
    }
    thead.appendChild(trh);
    const tbody = document.createElement("tbody");
    for (const row of rows){
      const tr = document.createElement("tr");
      for (const k of keys){
        const td = document.createElement("td");
        td.textContent = (row[k] ?? "").toString();
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    }
    tableEl.innerHTML = "";
    tableEl.appendChild(thead);
    tableEl.appendChild(tbody);
  }

  function escapeHtml_(s){
    return String(s||"")
      .replaceAll("&","&amp;")
      .replaceAll("<","&lt;")
      .replaceAll(">","&gt;")
      .replaceAll('"',"&quot;")
      .replaceAll("'","&#039;");
  }

  function readAsDataURL_(file){
    return new Promise((resolve, reject)=>{
      const fr = new FileReader();
      fr.onload = ()=>resolve(fr.result);
      fr.onerror = reject;
      fr.readAsDataURL(file);
    });
  }

  function num_(v){
    // Robust number parser for VK exports (handles spaces, NBSP, ₽, %, etc.)
    if (v == null || v === "") return 0;
    if (typeof v === "number") return isFinite(v) ? v : 0;
    let s = String(v)
      .replace(/\u00A0/g, " ")
      .trim();
    // keep digits, minus, dot, comma
    s = s.replace(/[^0-9,\.\-]/g, "");
    // if both comma and dot exist, assume comma is thousands separator -> remove commas
    if (s.includes(",") && s.includes(".")){
      s = s.replace(/,/g, "");
    }else{
      // otherwise treat comma as decimal separator
      s = s.replace(/,/g, ".");
    }
    const n = Number(s);
    return isFinite(n) ? n : 0;
  }
  function numOrNull_(v){
    const n = num_(v);
    return n ? n : null;
  }

  function toISODate_(raw){
    // XLSX can output Date objects or numbers or strings
    if (raw instanceof Date){
      return raw.toISOString().slice(0,10);
    }
    if (typeof raw === "number"){
      // Excel date serial
      const d = XLSX.SSF.parse_date_code(raw);
      if (d && d.y && d.m && d.d){
        const mm = String(d.m).padStart(2,"0");
        const dd = String(d.d).padStart(2,"0");
        return `${d.y}-${mm}-${dd}`;
      }
    }
    const s = String(raw||"").trim();
    // try DD.MM.YYYY
    const m1 = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})/);
    if (m1){
      const dd = String(m1[1]).padStart(2,"0");
      const mm = String(m1[2]).padStart(2,"0");
      return `${m1[3]}-${mm}-${dd}`;
    }
    // try YYYY-MM-DD
    const m2 = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m2){
      const mm = String(m2[2]).padStart(2,"0");
      const dd = String(m2[3]).padStart(2,"0");
      return `${m2[1]}-${mm}-${dd}`;
    }
    return "";
  }

  function fmtDate_(iso){
    if (!iso) return "—";
    try{
      const d = new Date(iso);
      return d.toLocaleString("ru-RU");
    }catch(e){ return iso; }
  }
  function fmtDateShort_(isoDate){
    if (!isoDate) return "—";
    const m = isoDate.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return isoDate;
    return `${m[3]}.${m[2]}`;
  }
  function pct_(a, total){
    const A = Number(a||0), T = Number(total||0);
    if (!T) return "0%";
    return Math.round((A/T)*100) + "%";
  }
  function sumObj_(o){
    let s = 0;
    for (const k in o) s += Number(o[k]||0);
    return s;
  }

  /** -----------------------------
   *  Boot
   * -----------------------------*/
  // 1) show local releases immediately
  syncSelectors_();
  go_("releases");

  // 2) then pull cloud index (so releases appear in any browser)
  //    (errors are non-fatal; UI will still work with local storage)
  cloudSyncFromIndex_().catch(e=>console.warn("cloudSyncFromIndex failed:", e));
})();
