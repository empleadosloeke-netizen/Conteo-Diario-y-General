// app.js (SIN el mapeo). Requiere planimetria.js con: window.PLANIMETRIA_MAP = {...}

(function () {
  const APP_SCRIPT_URL =
    "https://script.google.com/macros/s/AKfycbzksNcQMoLHl_fDpA5LJXRL7Cm6sgWpoyRXt4xASVefhVgXiSX2AJ_7mNU5UJ0ufdcI/exec";

  const MAP = window.PLANIMETRIA_MAP || {};
  const $ = (id) => document.getElementById(id);

  function normalizeCode(v) {
    return String(v ?? "").trim().toUpperCase().replace(/\s+/g, "");
  }
  function onlyDigits(v) {
    return String(v ?? "").replace(/[^\d]/g, "");
  }

  function showOverlay(on, title = "Enviando…", sub = "Guardando en Google Sheet") {
    const o = $("overlay");
    if (!o) return;
    const t1 = $("overlayTitle");
    const t2 = $("overlaySub");
    if (t1) t1.textContent = title;
    if (t2) t2.textContent = sub;
    o.style.display = on ? "flex" : "none";
  }

  function hideAll() {
    ["home", "tab1", "tab2", "tab3", "gTracker", "gSheet"].forEach((id) => {
      const el = $(id);
      if (el) el.style.display = "none";
    });
  }
  function show(id) {
    hideAll();
    const el = $(id);
    if (el) el.style.display = "";
  }

  // ---------- HOME ----------
  if ($("goDaily")) $("goDaily").addEventListener("click", () => show("tab1"));
  if ($("goGeneral")) $("goGeneral").addEventListener("click", () => show("gTracker"));

  // =========================
  // DIARIO
  // =========================
  function setStatus() {
    const leg = onlyDigits($("legajo")?.value).slice(0, 3);
    const c1 = normalizeCode($("cod1")?.value).slice(0, 4);
    const c2 = normalizeCode($("cod2")?.value).slice(0, 4);
    const cnt = [c1, c2].filter(Boolean).length;
    const pill = $("statusPill");
    if (pill) pill.textContent = leg ? `Legajo ${leg} · ${cnt} código(s)` : "Sin datos";
  }

  const STATE = { legajo: "", codes: [], rows: [], fills: {}, resumen: {} };

  function rowKey(r) {
    return `${r.orden}|${r.sector}|${r.codArt}`;
  }

  function buildSelect(min, max, selectedValue) {
    const sel = document.createElement("select");
    const opt0 = document.createElement("option");
    opt0.value = "";
    opt0.textContent = "—";
    sel.appendChild(opt0);
    for (let i = min; i <= max; i++) {
      const opt = document.createElement("option");
      opt.value = String(i);
      opt.textContent = String(i);
      if (String(i) === String(selectedValue)) opt.selected = true;
      sel.appendChild(opt);
    }
    return sel;
  }

  function calcTotal(pilas, cxp, sueltas) {
    const P = parseInt(pilas || "0", 10);
    const C = parseInt(cxp || "0", 10);
    const S = parseInt(sueltas || "0", 10);
    return P * C + S;
  }

  function parseTab1() {
    const legajo = onlyDigits($("legajo")?.value).slice(0, 3);
    const c1 = normalizeCode($("cod1")?.value).slice(0, 4);
    const c2 = normalizeCode($("cod2")?.value).slice(0, 4);

    const msg1 = $("msg1");
    const msg1b = $("msg1b");
    if (msg1) {
      msg1.textContent = "";
      msg1.className = "msg";
    }
    if (msg1b) {
      msg1b.textContent = "";
      msg1b.className = "msg";
    }

    if (!legajo) {
      if (msg1) {
        msg1.textContent = "Ingresá el legajo.";
        msg1.classList.add("error");
      }
      return null;
    }

    const codes = [c1, c2].filter(Boolean);
    if (!codes.length) {
      if (msg1) {
        msg1.textContent = "Ingresá al menos 1 código.";
        msg1.classList.add("error");
      }
      return null;
    }

    const uniq = Array.from(new Set(codes));
    const rows = [];

    for (const code of uniq) {
      const locs = MAP[code];
      if (Array.isArray(locs) && locs.length) {
        for (const loc of locs) {
          rows.push({
            orden: Number(loc.orden),
            sector: String(loc.sector),
            codArt: String(loc.codArt ?? code),
          });
        }
      } else {
        // si no está mapeado, lo mostramos al final
        rows.push({ orden: 999999, sector: "SIN MAPEO", codArt: code });
      }
    }

    rows.sort(
      (a, b) =>
        a.orden - b.orden ||
        a.sector.localeCompare(b.sector) ||
        a.codArt.localeCompare(b.codArt)
    );

    return { legajo, codes: uniq, rows };
  }

  function renderTab2() {
    const line = $("legajoLine");
    const pill = $("countPill");
    if (line) line.textContent = "Legajo: " + STATE.legajo;
    if (pill) pill.textContent = `${STATE.rows.length} filas`;

    const tbody = $("tbody");
    if (!tbody) return;
    tbody.innerHTML = "";

    for (const r of STATE.rows) {
      const key = rowKey(r);
      if (!STATE.fills[key]) STATE.fills[key] = { pilas: "", cjasXPila: "", cjasSueltas: "" };

      const tr = document.createElement("tr");

      const tdSec = document.createElement("td");
      tdSec.className = "colSector";
      tdSec.textContent = r.sector;

      const tdCod = document.createElement("td");
      tdCod.className = "colCod";
      tdCod.textContent = r.codArt;

      const tdP = document.createElement("td");
      tdP.className = "colPilas compact";
      const selP = buildSelect(1, 30, STATE.fills[key].pilas);
      selP.addEventListener("change", () => (STATE.fills[key].pilas = selP.value));
      tdP.appendChild(selP);

      const tdC = document.createElement("td");
      tdC.className = "colCxp compact";
      const selC = buildSelect(1, 10, STATE.fills[key].cjasXPila);
      selC.addEventListener("change", () => (STATE.fills[key].cjasXPila = selC.value));
      tdC.appendChild(selC);

      const tdS = document.createElement("td");
      tdS.className = "colSueltas compact";
      const inpS = document.createElement("input");
      inpS.type = "tel";
      inpS.inputMode = "numeric";
      inpS.maxLength = 2;
      inpS.placeholder = "0";
      inpS.value = STATE.fills[key].cjasSueltas || "";
      inpS.addEventListener("input", () => {
        inpS.value = onlyDigits(inpS.value).slice(0, 2);
        STATE.fills[key].cjasSueltas = inpS.value;
      });
      tdS.appendChild(inpS);

      tr.appendChild(tdSec);
      tr.appendChild(tdCod);
      tr.appendChild(tdP);
      tr.appendChild(tdC);
      tr.appendChild(tdS);

      tbody.appendChild(tr);
    }
  }

  function renderTab3() {
    const line = $("legajoLine3");
    const pill = $("countPill3");
    if (line) line.textContent = "Legajo: " + STATE.legajo;
    if (pill) pill.textContent = `${STATE.codes.length} códigos`;

    const tbody = $("tbody3");
    if (!tbody) return;
    tbody.innerHTML = "";

    for (const code of STATE.codes) {
      if (!STATE.resumen[code]) STATE.resumen[code] = { pickings: "", pfc: "", transito: "" };

      const tr = document.createElement("tr");

      const tdCod = document.createElement("td");
      tdCod.className = "colCod3";
      tdCod.textContent = code;

      const mkInput3 = (val) => {
        const inp = document.createElement("input");
        inp.type = "tel";
        inp.inputMode = "numeric";
        inp.maxLength = 3; // 3 dígitos
        inp.placeholder = "0";
        inp.value = val || "";
        return inp;
      };

      const tdPA = document.createElement("td");
      tdPA.className = "colPA compact";
      const inpPA = mkInput3(STATE.resumen[code].pickings);
      inpPA.addEventListener("input", () => {
        inpPA.value = onlyDigits(inpPA.value).slice(0, 3);
        STATE.resumen[code].pickings = inpPA.value;
      });
      tdPA.appendChild(inpPA);

      const tdPFC = document.createElement("td");
      tdPFC.className = "colPFC compact";
      const inpPFC = mkInput3(STATE.resumen[code].pfc);
      inpPFC.addEventListener("input", () => {
        inpPFC.value = onlyDigits(inpPFC.value).slice(0, 3);
        STATE.resumen[code].pfc = inpPFC.value;
      });
      tdPFC.appendChild(inpPFC);

      const tdMT = document.createElement("td");
      tdMT.className = "colMT compact";
      const inpMT = mkInput3(STATE.resumen[code].transito);
      inpMT.addEventListener("input", () => {
        inpMT.value = onlyDigits(inpMT.value).slice(0, 3);
        STATE.resumen[code].transito = inpMT.value;
      });
      tdMT.appendChild(inpMT);

      tr.appendChild(tdCod);
      tr.appendChild(tdPA);
      tr.appendChild(tdPFC);
      tr.appendChild(tdMT);

      tbody.appendChild(tr);
    }
  }

  function buildDailyPayload() {
    const paso2Rows = STATE.rows.map((r) => {
      const key = rowKey(r);
      const f = STATE.fills[key] || {};
      const total = calcTotal(f.pilas, f.cjasXPila, f.cjasSueltas);
      return {
        orden: r.orden === 999999 ? "" : r.orden,
        sector: r.sector,
        codArt: r.codArt,
        pilas: f.pilas || "",
        cjasXPila: f.cjasXPila || "",
        cjasSueltas: f.cjasSueltas || "",
        totalCjas: String(total),
        key,
      };
    });

    const paso3Rows = STATE.codes.map((code) => ({
      codArt: code,
      pickingsArmados: STATE.resumen[code]?.pickings ?? "",
      pedidosParaFC: STATE.resumen[code]?.pfc ?? "",
      mercaderiaEnTransito: STATE.resumen[code]?.transito ?? "",
    }));

    return {
      kind: "daily_final",
      ts: new Date().toISOString(),
      legajo: STATE.legajo,
      paso2Rows,
      paso3Rows,
    };
  }

  async function sendDailyFinal() {
    const msg = $("msg3");
    if (msg) {
      msg.textContent = "";
      msg.className = "msg";
    }

    try {
      showOverlay(true, "Enviando…", "Guardando todo (Paso 2 + Paso 3)");
      await fetch(APP_SCRIPT_URL, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify(buildDailyPayload()),
      });
      await new Promise((r) => setTimeout(r, 700));
      showOverlay(false);
      resetAll();
      show("home");
    } catch (e) {
      showOverlay(false);
      if (msg) {
        msg.textContent = "Error enviando: " + (e?.message || e);
        msg.classList.add("error");
      }
    }
  }

  function resetAll() {
    if ($("legajo")) $("legajo").value = "";
    if ($("cod1")) $("cod1").value = "";
    if ($("cod2")) $("cod2").value = "";

    ["msg1", "msg1b", "msgSend", "msg3"].forEach((id) => {
      const el = $(id);
      if (el) {
        el.textContent = "";
        el.className = "msg";
      }
    });

    STATE.legajo = "";
    STATE.codes = [];
    STATE.rows = [];
    STATE.fills = {};
    STATE.resumen = {};

    setStatus();
    show("tab1");
    $("legajo")?.focus?.();
  }

  // Eventos Diario
  $("legajo")?.addEventListener("input", () => {
    $("legajo").value = onlyDigits($("legajo").value).slice(0, 3);
    setStatus();
  });
  $("cod1")?.addEventListener("input", () => {
    $("cod1").value = normalizeCode($("cod1").value).slice(0, 4);
    setStatus();
  });
  $("cod2")?.addEventListener("input", () => {
    $("cod2").value = normalizeCode($("cod2").value).slice(0, 4);
    setStatus();
  });

  $("limpiarBtn")?.addEventListener("click", resetAll);

  $("siguienteBtn")?.addEventListener("click", () => {
    const parsed = parseTab1();
    if (!parsed) return;
    STATE.legajo = parsed.legajo;
    STATE.codes = parsed.codes;
    STATE.rows = parsed.rows;
    STATE.fills = {};
    STATE.resumen = {};
    renderTab2();
    show("tab2");
  });

  $("volverBtn")?.addEventListener("click", () => show("tab1"));

  $("resetFillsBtn")?.addEventListener("click", () => {
    for (const k of Object.keys(STATE.fills)) {
      STATE.fills[k] = { pilas: "", cjasXPila: "", cjasSueltas: "" };
    }
    renderTab2();
    const m = $("msgSend");
    if (m) {
      m.textContent = "Reseteado.";
      m.className = "msg ok";
    }
  });

  $("irPaso3Btn")?.addEventListener("click", () => {
    renderTab3();
    show("tab3");
  });

  $("volverBtn3")?.addEventListener("click", () => show("tab2"));

  $("reset3Btn")?.addEventListener("click", () => {
    for (const c of STATE.codes) STATE.resumen[c] = { pickings: "", pfc: "", transito: "" };
    renderTab3();
    const m = $("msg3");
    if (m) {
      m.textContent = "Reseteado.";
      m.className = "msg ok";
    }
  });

  $("enviarFinalBtn")?.addEventListener("click", sendDailyFinal);

  // =========================
  // CONTEO GENERAL (SOLO usa sectores reales de la planimetría)
  // =========================
  const GEN = {
    pageSize: 30,
    conteoId: "",
    legajo: "",
    lists: { sin: [], con: [] },  // pages de sectores (strings)
    status: { sin: [], con: [] }, // pendiente/en_curso/terminado
    current: { rol: "", idx: -1, page: [] },
  };

  // Regla escalera sobre SECTOR real (A tiene 5, resto 4)
  function isConEscaleraFromSector(sectorStr) {
    const s = String(sectorStr || "").trim().toUpperCase();
    const m = s.match(/^([A-Z])\s*(\d+)$/);
    if (!m) return false;
    const letter = m[1];
    const n = parseInt(m[2], 10);
    if (!Number.isFinite(n)) return false;

    if (letter === "A") {
      const r = (n - 1) % 5;
      return r === 3 || r === 4; // 4 y 5 con escalera
    }
    const r = (n - 1) % 4;
    return r === 3; // 4 con escalera
  }

  function getAllSectorsFromPlanimetria() {
    const set = new Set();
    for (const cod of Object.keys(MAP)) {
      const arr = MAP[cod];
      if (!Array.isArray(arr)) continue;
      for (const it of arr) {
        const sec = String(it?.sector || "").trim();
        if (sec) set.add(sec);
      }
    }
    return Array.from(set);
  }

  function sortSectorsNatural(a, b) {
    const A = String(a).toUpperCase();
    const B = String(b).toUpperCase();
    const ma = A.match(/^([A-Z])\s*(\d+)$/);
    const mb = B.match(/^([A-Z])\s*(\d+)$/);
    if (ma && mb) {
      if (ma[1] !== mb[1]) return ma[1].localeCompare(mb[1]);
      return parseInt(ma[2], 10) - parseInt(mb[2], 10);
    }
    if (ma && !mb) return -1;
    if (!ma && mb) return 1;
    return A.localeCompare(B);
  }

  function chunk(arr, size) {
    const out = [];
    for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
    return out;
  }

  function ensureGeneralLists() {
    const sectors = getAllSectorsFromPlanimetria().sort(sortSectorsNatural);

    const sin = [];
    const con = [];
    for (const sec of sectors) {
      if (isConEscaleraFromSector(sec)) con.push(sec);
      else sin.push(sec);
    }

    GEN.lists.sin = chunk(sin, GEN.pageSize);
    GEN.lists.con = chunk(con, GEN.pageSize);

    // inicializa si estaba vacío
    if (!GEN.status.sin.length) GEN.status.sin = Array(GEN.lists.sin.length).fill("pendiente");
    if (!GEN.status.con.length) GEN.status.con = Array(GEN.lists.con.length).fill("pendiente");
  }

  function badge(st) {
    if (st === "terminado") return "✅";
    if (st === "en_curso") return "⏳";
    return "⬜";
  }

  async function postJSONRead(payload) {
    // Para leer estado/reserva necesitamos CORS en Apps Script
    const res = await fetch(APP_SCRIPT_URL, {
      method: "POST",
      mode: "cors",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload),
    });
    const txt = await res.text();
    try { return JSON.parse(txt); }
    catch { return { ok: false, error: "Respuesta no JSON", raw: txt }; }
  }

  function renderTracker() {
    const sinBox = $("g_list_sin");
    const conBox = $("g_list_con");
    if (!sinBox || !conBox) return;

    sinBox.innerHTML = "";
    conBox.innerHTML = "";

    const countDone = (arr) => arr.filter((s) => s === "terminado").length;
    if ($("g_sinPill")) $("g_sinPill").textContent = `${countDone(GEN.status.sin)}/${GEN.status.sin.length}`;
    if ($("g_conPill")) $("g_conPill").textContent = `${countDone(GEN.status.con)}/${GEN.status.con.length}`;

    const renderList = (rol, box) => {
      GEN.lists[rol].forEach((page, idx) => {
        const st = GEN.status[rol][idx] || "pendiente";
        const div = document.createElement("div");
        div.className = "trackerRow" + (st !== "pendiente" ? " disabled" : "");

        const first = page[0] || "—";
        const last = page[page.length - 1] || "—";

        div.innerHTML = `
          <div class="left">
            <div class="name">Hoja ${idx + 1}</div>
            <div class="sub">${first} … ${last}</div>
          </div>
          <div class="badge">${badge(st)}</div>
        `;

        if (st === "pendiente") div.addEventListener("click", () => openSheet(rol, idx));
        box.appendChild(div);
      });
    };

    renderList("sin", sinBox);
    renderList("con", conBox);
  }

  async function loadStatus() {
    const m = $("g_msg");
    if (m) {
      m.textContent = "";
      m.className = "msg";
    }

    GEN.legajo = onlyDigits($("g_legajo")?.value).slice(0, 3);
    GEN.conteoId = normalizeCode($("g_conteoId")?.value).slice(0, 16);

    if (!GEN.legajo) {
      if (m) { m.textContent = "Ingresá legajo."; m.classList.add("error"); }
      return;
    }
    if (!GEN.conteoId) {
      if (m) { m.textContent = "Ingresá ID Conteo."; m.classList.add("error"); }
      return;
    }

    ensureGeneralLists();

    try {
      showOverlay(true, "Cargando…", "Leyendo estado del conteo");
      const resp = await postJSONRead({ kind: "general_status", conteoId: GEN.conteoId });
      showOverlay(false);

      if (!resp || !resp.ok) {
        if (m) { m.textContent = resp?.error || "No pude leer el estado (CORS)."; m.classList.add("error"); }
        renderTracker(); // igual muestra pendiente
        return;
      }

      // Si el Apps Script devuelve arrays, los uso.
      if (Array.isArray(resp.sin)) GEN.status.sin = resp.sin;
      if (Array.isArray(resp.con)) GEN.status.con = resp.con;

      renderTracker();
      if (m) { m.textContent = "Listo."; m.className = "msg ok"; }
    } catch (e) {
      showOverlay(false);
      if (m) {
        m.textContent = "No pude leer estado (revisar CORS en Apps Script).";
        m.classList.add("error");
      }
      renderTracker();
    }
  }

  async function openSheet(rol, idx) {
    const m = $("g_msg");
    if (m) { m.textContent = ""; m.className = "msg"; }

    try {
      showOverlay(true, "Reservando…", "Marcando hoja en curso");
      const r = await postJSONRead({
        kind: "general_reserve",
        conteoId: GEN.conteoId,
        legajo: GEN.legajo,
        rol,
        sheetIndex: idx,
      });
      showOverlay(false);

      if (!r || !r.ok) {
        if (m) { m.textContent = r?.error || "No se pudo reservar."; m.classList.add("error"); }
        return;
      }

      GEN.status[rol][idx] = "en_curso";
      renderTracker();

      GEN.current = { rol, idx, page: GEN.lists[rol][idx] };
      renderGSheet();
      show("gSheet");
    } catch (e) {
      showOverlay(false);
      if (m) {
        m.textContent = "No se pudo reservar (revisar CORS en Apps Script).";
        m.classList.add("error");
      }
    }
  }

  function renderGSheet() {
    const { rol, idx, page } = GEN.current;

    if ($("g_title")) $("g_title").textContent = `${rol === "sin" ? "Sin escalera" : "Con escalera"} · Hoja ${idx + 1}`;
    if ($("g_sub")) $("g_sub").textContent = `Legajo ${GEN.legajo} · Conteo ${GEN.conteoId}`;
    if ($("g_progress")) $("g_progress").textContent = `${idx + 1}/${GEN.lists[rol].length}`;

    const tb = $("g_tbody");
    if (!tb) return;
    tb.innerHTML = "";

    page.forEach((sec) => {
      const tr = document.createElement("tr");

      const tdU = document.createElement("td");
      tdU.textContent = sec;

      const tdC = document.createElement("td");
      tdC.className = "compact";
      const inp = document.createElement("input");
      inp.type = "tel";
      inp.inputMode = "numeric";
      inp.maxLength = 4;
      inp.placeholder = "0";
      inp.id = "g_" + sec;
      inp.addEventListener("input", () => (inp.value = onlyDigits(inp.value).slice(0, 4)));
      tdC.appendChild(inp);

      tr.appendChild(tdU);
      tr.appendChild(tdC);
      tb.appendChild(tr);
    });
  }

  function resetGSheetInputs() {
    GEN.current.page.forEach((sec) => {
      const el = document.getElementById("g_" + sec);
      if (el) el.value = "";
    });
    const m = $("g_msg2");
    if (m) { m.textContent = "Reseteado."; m.className = "msg ok"; }
  }

  async function sendGSheet() {
    const m = $("g_msg2");
    if (m) { m.textContent = ""; m.className = "msg"; }

    const { rol, idx, page } = GEN.current;
    const rows = page.map((sec) => ({
      ubicacion: sec,
      cantidad: onlyDigits(document.getElementById("g_" + sec)?.value || "0"),
    }));

    try {
      showOverlay(true, "Enviando…", "Guardando hoja del conteo general");
      await fetch(APP_SCRIPT_URL, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({
          kind: "general_page",
          ts: new Date().toISOString(),
          conteoId: GEN.conteoId,
          legajo: GEN.legajo,
          rol,
          pageIndex: idx,
          pageSize: GEN.pageSize,
          rows,
        }),
      });

      await new Promise((r) => setTimeout(r, 700));
      showOverlay(false);

      GEN.status[rol][idx] = "terminado";
      renderTracker();
      show("gTracker");
    } catch (e) {
      showOverlay(false);
      if (m) { m.textContent = "Error enviando."; m.classList.add("error"); }
    }
  }

  $("g_backHome")?.addEventListener("click", () => show("home"));
  $("g_load")?.addEventListener("click", loadStatus);
  $("g_backTracker")?.addEventListener("click", () => show("gTracker"));
  $("g_reset")?.addEventListener("click", resetGSheetInputs);
  $("g_send")?.addEventListener("click", sendGSheet);

  // init
  setStatus();
  show("home");
})();
