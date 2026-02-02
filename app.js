/* Lotus Music · VK Digital-отчёт (MVP)
   - Без географии
   - Детализация = список объявлений
   - Стримы = колонка vk
   - Blocked/статусы не фильтруем
*/
(function(){
  const $ = (s, root=document) => root.querySelector(s);
  const $$ = (s, root=document) => Array.from(root.querySelectorAll(s));

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
    localStorage.setItem(StoreKey, JSON.stringify(state.db));
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

  function renderReleases_(){
    const list = $("#releases-list");
    const releases = Object.values(state.db.releases).sort((a,b)=> (b.updatedAt||"").localeCompare(a.updatedAt||""));
    $("#releases-count").textContent = String(releases.length);
    list.innerHTML = "";
    if (!releases.length){
      list.innerHTML = `<div class="muted">Пока нет релизов. Нажми «+ Добавить релиз».</div>`;
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
            <button class="btn inline danger" data-act="delete">Удалить…</button>
          </div>
        </div>
      `;
      item.querySelector('[data-act="open"]').addEventListener("click", ()=>{
        state.currentReleaseId = r.releaseId;
        syncSelectors_();
        go_("editor");
      });
      item.querySelector('[data-act="build"]').addEventListener("click", ()=>{
        state.currentReleaseId = r.releaseId;
        syncSelectors_();
        go_("report");
      });
      item.querySelector('[data-act="delete"]').addEventListener("click", ()=>{
        openDeleteModal_(r.releaseId);
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
    const releases = Object.values(state.db.releases).sort((a,b)=> (b.updatedAt||"").localeCompare(a.updatedAt||""));
    const makeOptions = (sel)=>{
      sel.innerHTML = "";
      if (!releases.length){
        const o = document.createElement("option");
        o.value = "";
        o.textContent = "Нет релизов";
        sel.appendChild(o);
        sel.disabled = true;
        return;
      }
      sel.disabled = false;
      for (const r of releases){
        const o = document.createElement("option");
        o.value = r.releaseId;
        o.textContent = r.title;
        sel.appendChild(o);
      }
      const choose = state.currentReleaseId && state.db.releases[state.currentReleaseId] ? state.currentReleaseId : releases[0].releaseId;
      state.currentReleaseId = choose;
      sel.value = choose;
    };
    makeOptions(editorSelect);
    makeOptions(reportSelect);
  }

  editorSelect.addEventListener("change", ()=>{
    state.currentReleaseId = editorSelect.value || null;
    renderEditor_();
  });
  reportSelect.addEventListener("change", ()=>{
    state.currentReleaseId = reportSelect.value || null;
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
      demoRows: null,
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
      const demoOk = Array.isArray(c.demoRows) && c.demoRows.length>0;
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
            <div class="hint">Берём лист «Объявления» (если есть), иначе первый лист. Нужные колонки: Потрачено всего, ₽ / Показы / Начали прослушивание / Добавили аудио / ID группы.</div>
          </div>

          <div class="card" style="margin:0">
            <div class="cardTitle">
              <h3>Демография</h3>
              <span class="pill ${demoOk ? "badgeOk":"badgeWarn"}">${demoOk ? "загружено":"не загружено"}</span>
            </div>
            <div class="row">
              <input type="file" accept=".xlsx" data-up="demo" />
              <button class="btn danger inline" data-del="demo" ${demoOk ? "":"disabled"}>Удалить</button>
            </div>
            <div class="miniTableWrap" ${demoOk ? "":"hidden"} data-prevwrap="demo">
              <div class="miniTableTitle">Превью (первые 5 строк)</div>
              <table class="miniTable" data-prev="demo"></table>
            </div>
            <div class="hint">Ожидаем колонки: Возраст / Пол / Показы / Клики / Цена за результат, ₽ (и т.п.).</div>
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
          <div class="hint">MVP: просто загружаем изображения и показываем горизонтально в отчёте.</div>
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
          const file = e.target.files && e.target.files[0];
          const kind = e.target.dataset.up;
          if (!file) return;
          try{
            if (kind==="ads"){
              const rows = await parseAdsXlsx_(file);
              c.adsRows = rows;
              renderCommunityPreviews_(card, c);
            } else if (kind==="demo"){
              const rows = await parseDemoXlsx_(file);
              c.demoRows = rows;
              renderCommunityPreviews_(card, c);
            } else if (kind==="creatives"){
              const files = Array.from(e.target.files||[]);
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
            if (!confirm("Удалить файл демографии и данные?")) return;
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
    const dWrap = card.querySelector('[data-prevwrap="demo"]');
    const dTbl = card.querySelector('[data-prev="demo"]');
    if (Array.isArray(c.demoRows) && c.demoRows.length){
      dWrap.hidden = false;
      const demoCostKey =
  (c.demoRows?.[0] && ("Цена за результат, ₽" in c.demoRows[0])) ? "Цена за результат, ₽"
  : (c.demoRows?.[0] && ("Цена за результат, Р" in c.demoRows[0])) ? "Цена за результат, Р"
  : "Цена за результат, ₽";

renderMiniPreview_(dTbl, c.demoRows.slice(0,5), ["Возраст","Пол","Показы","Клики", demoCostKey]);

    } else {
      dWrap.hidden = true;
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
      if (!Array.isArray(c.demoRows) || !c.demoRows.length) return false;
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
    const releases = Object.values(state.db.releases).sort((a,b)=> (b.updatedAt||"").localeCompare(a.updatedAt||""));
    $("#history-count").textContent = String(releases.length);
    list.innerHTML = "";
    if (!releases.length){
      list.innerHTML = `<div class="muted">История пуста.</div>`;
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
      item.querySelector('[data-act="open"]').addEventListener("click", ()=>{
        state.currentReleaseId = r.releaseId;
        syncSelectors_();
        go_("editor");
      });
      item.querySelector('[data-act="report"]').addEventListener("click", ()=>{
        state.currentReleaseId = r.releaseId;
        syncSelectors_();
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
      if (Array.isArray(c.demoRows) && c.demoRows.length) demoFiles++;
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
    // Печать только на странице отчёта.
    go_("report");
    setTimeout(()=>window.print(), 50);
  });

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

    return pages;
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
        <div class="coverName">${escapeHtml_(r.title)}</div>
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

    if (isMultiSummary){
      const rows = commItems.map(it=>{
        const kk = calcKpisFromAdsRows_(it.community.adsRows||[]);
        return { name: it.label, ...kk };
      });
      inner.insertAdjacentHTML("beforeend", `<div style="margin-top:14px; font-weight:900; opacity:.8">Сводка по сообществам</div>`);
      inner.appendChild(renderSimpleTable_(
        ["Сообщество","Потрачено","Показы","Прослуш.","Добавл.","Сегм.","Ср.цена добав."],
        rows.map(x=>[
          x.name,
          formatMoney2_(x.spent),
          formatInt_(x.shows),
          formatInt_(x.listens),
          formatInt_(x.adds),
          formatInt_(x.segments),
          x.avgCost==null ? "—" : formatMoney2_(x.avgCost)
        ])
      ));
    }

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

    const demo = (c.demoRows || []);

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

    const sumImpr = genderCats.reduce((s,k)=>s + (genderImpr[k]||0), 0);
    const sumClicks = genderCats.reduce((s,k)=>s + (genderClicks[k]||0), 0);

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

    const menI = genderImpr["Мужчины"]||0, womI = genderImpr["Женщины"]||0, noneI = genderImpr["Пол не указан"]||0;
    const menC = genderClicks["Мужчины"]||0, womC = genderClicks["Женщины"]||0, noneC = genderClicks["Пол не указан"]||0;

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
            <td>${escapeHtml_(valPct(menI, sumImpr))}</td>
            <td>${escapeHtml_(valPct(womI, sumImpr))}</td>
            <td>${escapeHtml_(valPct(sumImpr, sumImpr))}</td>
          </tr>
          <tr>
            <td>Клики</td>
            <td>${escapeHtml_(valPct(menC, sumClicks))}</td>
            <td>${escapeHtml_(valPct(womC, sumClicks))}</td>
            <td>${escapeHtml_(valPct(sumClicks, sumClicks))}</td>
          </tr>
          <tr>
            <td>Цена за<br/>результат, ₽</td>
            <td>${escapeHtml_(cost(menSpent, menC))}</td>
            <td>${escapeHtml_(cost(womSpent, womC))}</td>
            <td>${escapeHtml_(cost(menSpent+womSpent+noneSpent, sumClicks))}</td>
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
        axisTransparent: true
      }
    );

    ageCard.appendChild(ageChart);
    grid.appendChild(ageCard);

    return el;
  }


  function renderStreamsPage_(r){
    const el = sheet_(true);
    const inner = el.querySelector(".sheetInner");
    inner.insertAdjacentHTML("beforeend", `<div class="sheetTitle" style="font-size:28px">График прослушиваний VK</div>`);
    const rows = (r.streams||[]);
    const chart = svgLineChart_("VK стримы по дням", rows.map(x=>({x:x.dateISO, y:x.vk})));
    inner.appendChild(chart);

    // table
    const tab = document.createElement("table");
    tab.className = "table";
    tab.innerHTML = `<thead><tr><th>Дата</th><th>VK</th></tr></thead><tbody></tbody>`;
    const tb = tab.querySelector("tbody");
    for (const s of rows){
      const tr = document.createElement("tr");
      tr.innerHTML = `<td>${escapeHtml_(fmtDateShort_(s.dateISO))}</td><td>${formatInt_(s.vk)}</td>`;
      tb.appendChild(tr);
    }
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
  for (const k in row){
    out[normKey_(k)] = row[k];
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

  // Keep only rows having Age+Gender
  const out = [];
  for (const row of json){
    const age = row["Возраст"] ?? row["Age"] ?? "";
    const gender = row["Пол"] ?? row["Gender"] ?? "";
    if (!age || !gender) continue;
    out.push(row);
  }
  return out.length ? out : json; // fallback
}


  /** -----------------------------
   *  Demo aggregation (MVP)
   * -----------------------------*/
  function normalizeGender_(g){
    const s = String(g||"").trim();
    if (!s) return "Пол не указан";
    if (s.toLowerCase().includes("муж")) return "Мужчины";
    if (s.toLowerCase().includes("жен")) return "Женщины";
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
  const barWidthRatio = opts.barWidthRatio ?? 0.78;  // столбцы уже
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

        const bh1 = (h - pad * 2) * (v1 / (opts.dualAxis ? maxBase : max));
        const bh2 = (h - pad * 2) * (v2 / (opts.dualAxis ? maxOver : max));

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
  syncSelectors_();
  go_("releases");
})();
