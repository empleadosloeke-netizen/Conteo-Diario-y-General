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

  function showOverlay(on) {
    $("overlay").style.display = on ? "flex" : "none";
  }

  function show(section) {
    const ids = ["home", "daily1", "daily2", "daily3", "general", "generalPage"];
    ids.forEach((id) => {
      const el = $(id);
      if (!el) return;
      el.style.display = id === section ? "" : "none";
    });
  }

  async function postJSON(payload) {
    // OJO: usamos no-cors para que funcione desde GitHub Pages sin CORS.
    // Esto impide leer respuesta real en muchos navegadores.
    // Si necesitás leer respuesta (para status/reserve), lo ideal es que el Apps Script
    // devuelva CORS y usar mode:"cors". Por ahora lo hacemos "best effort":
    try {
      const res = await fetch(APP_SCRIPT_URL, {
        method: "POST",
        mode: "cors",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify(payload),
      });
      // Si CORS está habilitado en tu Apps Script, esto funciona:
      const txt = await res.text();
      try {
        return JSON.parse(txt);
      } catch {
        return { ok: true, raw: txt };
      }
    } catch (e) {
      // fallback: igual enviamos no-cors (no podremos leer respuesta)
      await fetch(APP_SCRIPT_URL, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify(payload),
      });
      return { ok: true, note: "no-cors (sin respuesta)" };
    }
  }

  // ========= HOME =========
  $("goDaily").addEventListener("click", () => show("daily1"));
  $("goGeneral").addEventListener("click", () => show("general"));
  $("homeFromDaily1").addEventListener("click", () => show("home"));
  $("homeFromGeneral").addEventListener("click", () => show("home"));

  // ========= DIARIO =========
  const DAILY = { legajo: "", codes: [], rows: [], fills: {}, resumen: {} };
  function rowKey(r) {
    return `${r.orden}|${r.sector}|${r.codArt}`;
  }

  $("legajo").addEventListener("input", () => {
    $("legajo").value = onlyDigits($("legajo").value).slice(0, 3);
  });
  $("cod1").addEventListener("input", () => {
    $("cod1").value = normalizeCode($("cod1").value).slice(0, 4);
  });
  $("cod2").addEventListener("input", () => {
    $("cod2").value = normalizeCode($("cod2").value).slice(0, 4);
  });

  function parseDailyTab1() {
    const legajo = onlyDigits($("legajo").value).slice(0, 3);
    const cod1 = normalizeCode($("cod1").value).slice(0, 4);
    const cod2 = normalizeCode($("cod2").value).slice(0, 4);

    $("msg1").textContent = "";
    $("msg1").className = "msg";

    if (!legajo) {
      $("msg1").textContent = "Ingresá el legajo.";
      $("msg1").classList.add("error");
      return null;
    }

    const codes = [cod1, cod2].filter(Boolean);
    if (!codes.length) {
      $("msg1").textContent = "Ingresá al menos 1 código.";
      $("msg1").classList.add("error");
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

  function buildSelect(min, max, selected) {
    const sel = document.createElement("select");
    const o0 = document.createElement("option");
    o0.value = "";
    o0.textContent = "—";
    sel.appendChild(o0);

    for (let i = min; i <= max; i++) {
      const o = document.createElement("option");
      o.value = String(i);
      o.textContent = String(i);
      if (String(i) === String(selected)) o.selected = true;
      sel.appendChild(o);
    }
    return sel;
  }

  function calcTotal(pilas, cxp, sueltas) {
    const P = parseInt(pilas || "0", 10);
    const C = parseInt(cxp || "0", 10);
    const S = parseInt(sueltas || "0", 10);
    return P * C + S;
  }

  function renderDailyTab2() {
    show("daily2");
    $("legajoLine").textContent = "Legajo: " + DAILY.legajo;

    const tbody = $("tbody");
    tbody.innerHTML = "";

    for (const r of DAILY.rows) {
      const key = rowKey(r);
      if (!DAILY.fills[key])
        DAILY.fills[key] = { pilas: "", cjasXPila: "", cjasSueltas: "" };

      const tr = document.createElement("tr");

      const tdSec = document.createElement("td");
      tdSec.textContent = r.sector;

      const tdCod = document.createElement("td");
      tdCod.textContent = r.codArt;

      const tdP = document.createElement("td");
      const selP = buildSelect(1, 30, DAILY.fills[key].pilas);
      selP.addEventListener("change", () => (DAILY.fills[key].pilas = selP.value));

      const tdC = document.createElement("td");
      const selC = buildSelect(1, 10, DAILY.fills[key].cjasXPila);
      selC.addEventListener("change", () => (DAILY.fills[key].cjasXPila = selC.value));

      const tdS = document.createElement("td");
      const inpS = document.createElement("input");
      inpS.type = "tel";
      inpS.inputMode = "numeric";
      inpS.maxLength = 2;
      inpS.placeholder = "0";
      inpS.value = DAILY.fills[key].cjasSueltas || "";
      inpS.addEventListener("input", () => {
        inpS.value = onlyDigits(inpS.value).slice(0, 2);
        DAILY.fills[key].cjasSueltas = inpS.value;
      });

      tdP.appendChild(selP);
      tdC.appendChild(selC);
      tdS.appendChild(inpS);

      tr.appendChild(tdSec);
      tr.appendChild(tdCod);
      tr.appendChild(tdP);
      tr.appendChild(tdC);
      tr.appendChild(tdS);

      tbody.appendChild(tr);
    }
  }

  function renderDailyTab3() {
    show("daily3");
    $("legajoLine3").textContent = "Legajo: " + DAILY.legajo;

    const tbody = $("tbody3");
    tbody.innerHTML = "";

    for (const code of DAILY.codes) {
      if (!DAILY.resumen[code])
        DAILY.resumen[code] = { pickings: "", pfc: "", transito: "" };

      const tr = document.createElement("tr");

      const tdCod = document.createElement("td");
      tdCod.textContent = code;

      const mkInput = (val) => {
        const inp = document.createElement("input");
        inp.type = "tel";
        inp.inputMode = "numeric";
        inp.maxLength = 3; // 3 dígitos
        inp.placeholder = "0";
        inp.value = val || "";
        return inp;
      };

      const tdPA = document.createElement("td");
      const inpPA = mkInput(DAILY.resumen[code].pickings);
      inpPA.addEventListener("input", () => {
        inpPA.value = onlyDigits(inpPA.value).slice(0, 3);
        DAILY.resumen[code].pickings = inpPA.value;
      });

      const tdPFC = document.createElement("td");
      const inpPFC = mkInput(DAILY.resumen[code].pfc);
      inpPFC.addEventListener("input", () => {
        inpPFC.value = onlyDigits(inpPFC.value).slice(0, 3);
        DAILY.resumen[code].pfc = inpPFC.value;
      });

      const tdMT = document.createElement("td");
      const inpMT = mkInput(DAILY.resumen[code].transito);
      inpMT.addEventListener("input", () => {
        inpMT.value = onlyDigits(inpMT.value).slice(0, 3);
        DAILY.resumen[code].transito = inpMT.value;
      });

      tdPA.appendChild(inpPA);
      tdPFC.appendChild(inpPFC);
      tdMT.appendChild(inpMT);

      tr.appendChild(tdCod);
      tr.appendChild(tdPA);
      tr.appendChild(tdPFC);
      tr.appendChild(tdMT);

      tbody.appendChild(tr);
    }
  }

  function buildDailyPayload() {
    const paso2Rows = DAILY.rows.map((r) => {
      const key = rowKey(r);
      const f = DAILY.fills[key] || {};
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

    const paso3Rows = DAILY.codes.map((code) => ({
      codArt: code,
      pickingsArmados: DAILY.resumen[code]?.pickings ?? "",
      pedidosParaFC: DAILY.resumen[code]?.pfc ?? "",
      mercaderiaEnTransito: DAILY.resumen[code]?.transito ?? "",
    }));

    return {
      kind: "daily_final",
      ts: new Date().toISOString(),
      legajo: DAILY.legajo,
      paso2Rows,
      paso3Rows,
    };
  }

  async function sendDaily() {
    $("msg3").textContent = "";
    $("msg3").className = "msg";
    showOverlay(true);

    await fetch(APP_SCRIPT_URL, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(buildDailyPayload()),
    });

    setTimeout(() => {
      showOverlay(false);
      resetDaily();
      show("home");
    }, 700);
  }

  function resetDaily() {
    $("legajo").value = "";
    $("cod1").value = "";
    $("cod2").value = "";
    $("msg1").textContent = "";
    $("msg3").textContent = "";
    DAILY.legajo = "";
    DAILY.codes = [];
    DAILY.rows = [];
    DAILY.fills = {};
    DAILY.resumen = {};
  }

  $("siguienteBtn").addEventListener("click", () => {
    const parsed = parseDailyTab1();
    if (!parsed) return;

    DAILY.legajo = parsed.legajo;
    DAILY.codes = parsed.codes;
    DAILY.rows = parsed.rows;
    DAILY.fills = {};
    DAILY.resumen = {};
    renderDailyTab2();
  });

  $("volverBtn").addEventListener("click", () => show("daily1"));
  $("irPaso3Btn").addEventListener("click", () => renderDailyTab3());
  $("volverBtn3").addEventListener("click", () => show("daily2"));
  $("resetFillsBtn").addEventListener("click", () => renderDailyTab2());
  $("enviarFinalBtn").addEventListener("click", () => sendDaily());

  // ========= GENERAL (tracker + hojas) =========
  const GEN = {
    conteoId: "",
    legajo: "",
    pageSize: 30,
    lists: { sin: [], con: [] },
    current: { rol: "", sheetIndex: -1, ubicaciones: [] },
    status: { sin: [], con: [] }, // "pendiente" | "en_curso" | "terminado"
  };

  function isConEscalera(L, n) {
    if (L === "A") {
      const r = (n - 1) % 5;
      return r === 3 || r === 4;
    }
    const r = (n - 1) % 4;
    return r === 3;
  }

  function genAllUbicaciones() {
    const letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("");
    const maxN = 100; // AJUSTAR a tu depósito real
    const out = [];
    for (const L of letters) {
      for (let n = 1; n <= maxN; n++) {
        out.push({
          ubicacion: `${L}${n}`,
          rol: isConEscalera(L, n) ? "con" : "sin",
        });
      }
    }
    return out;
  }

  function chunk(arr, size) {
    const out = [];
    for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
    return out;
  }

  function ensureGeneralListsBuilt() {
    const all = genAllUbicaciones();
    GEN.lists.sin = chunk(all.filter((x) => x.rol === "sin"), GEN.pageSize);
    GEN.lists.con = chunk(all.filter((x) => x.rol === "con"), GEN.pageSize);

    // default statuses si todavía no tenemos del server
    if (!GEN.status.sin.length) GEN.status.sin = Array(GEN.lists.sin.length).fill("pendiente");
    if (!GEN.status.con.length) GEN.status.con = Array(GEN.lists.con.length).fill("pendiente");
  }

  function badgeFor(st) {
    if (st === "terminado") return "✅";
    if (st === "en_curso") return "⏳";
    return "⬜";
  }

  function renderTracker() {
    const sinBox = $("g_list_sin");
    const conBox = $("g_list_con");
    sinBox.innerHTML = "";
    conBox.innerHTML = "";

    const renderList = (rol, box) => {
      const list = GEN.lists[rol];
      const stArr = GEN.status[rol] || [];

      list.forEach((page, idx) => {
        const st = stArr[idx] || "pendiente";
        const div = document.createElement("div");
        div.className = "trackerRow";
        div.innerHTML = `
          <div class="trackerLeft">
            <div class="trackerTitle">Hoja ${idx + 1}</div>
            <div class="trackerSub">${page[0].ubicacion} … ${page[page.length - 1].ubicacion}</div>
          </div>
          <div class="trackerRight">${badgeFor(st)}</div>
        `;

        // solo permite abrir si NO está terminada.
        // si está en_curso, lo dejamos bloqueado (para evitar doble conteo).
        if (st === "pendiente") {
          div.addEventListener("click", () => openGeneralSheet(rol, idx));
        } else {
          div.style.opacity = "0.65";
        }

        box.appendChild(div);
      });
    };

    renderList("sin", sinBox);
    renderList("con", conBox);
  }

  async function loadStatusFromServer() {
    // Requiere Apps Script con kind="general_status"
    // Respuesta esperada:
    // { ok:true, sin:["pendiente"...], con:["pendiente"...] }
    const resp = await postJSON({
      kind: "general_status",
      conteoId: GEN.conteoId,
    });

    if (resp && resp.ok && Array.isArray(resp.sin) && Array.isArray(resp.con)) {
      GEN.status.sin = resp.sin;
      GEN.status.con = resp.con;
    }
  }

  async function reserveSheetOnServer(rol, sheetIndex) {
    // Requiere Apps Script con kind="general_reserve"
    // Respuesta esperada:
    // { ok:true, status:"en_curso" } o { ok:false, error:"..." }
    const resp = await postJSON({
      kind: "general_reserve",
      conteoId: GEN.conteoId,
      legajo: GEN.legajo,
      rol,
      sheetIndex,
    });

    // si no pudimos leer respuesta (no-cors), igual avanzamos "optimista"
    if (!resp) return { ok: true };
    return resp;
  }

  async function openGeneralSheet(rol, sheetIndex) {
    $("g_msg").textContent = "";
    $("g_msg").className = "msg";

    // reservar (para poner relojito y bloquear)
    showOverlay(true);
    const r = await reserveSheetOnServer(rol, sheetIndex);
    showOverlay(false);

    if (r && r.ok === false) {
      $("g_msg").textContent = r.error || "No se puede abrir esa hoja.";
      $("g_msg").className = "msg error";
      return;
    }

    // marcar local como en curso
    GEN.status[rol][sheetIndex] = "en_curso";
    renderTracker();

    GEN.current.rol = rol;
    GEN.current.sheetIndex = sheetIndex;
    GEN.current.ubicaciones = GEN.lists[rol][sheetIndex];

    renderGeneralPage();
    show("generalPage");
  }

  function renderGeneralPage() {
    const rol = GEN.current.rol;
    const idx = GEN.current.sheetIndex;
    const page = GEN.current.ubicaciones;

    $("g_title").textContent = `${rol === "sin" ? "Sin escalera" : "Con escalera"} · Hoja ${idx + 1}`;
    $("g_progress").textContent = `${idx + 1}/${GEN.lists[rol].length}`;

    const tbody = $("g_tbody");
    tbody.innerHTML = "";

    page.forEach((u) => {
      const tr = document.createElement("tr");

      const tdU = document.createElement("td");
      tdU.textContent = u.ubicacion;

      const tdC = document.createElement("td");
      const inp = document.createElement("input");
      inp.type = "tel";
      inp.inputMode = "numeric";
      inp.maxLength = 4;
      inp.placeholder = "0";
      inp.id = "g_" + u.ubicacion;
      inp.addEventListener("input", () => {
        inp.value = onlyDigits(inp.value).slice(0, 4);
      });

      tdC.appendChild(inp);
      tr.appendChild(tdU);
      tr.appendChild(tdC);
      tbody.appendChild(tr);
    });
  }

  async function sendGeneralSheet() {
    $("g_msg2").textContent = "";
    $("g_msg2").className = "msg";

    const rol = GEN.current.rol;
    const sheetIndex = GEN.current.sheetIndex;
    const page = GEN.current.ubicaciones;

    const rows = page.map((u) => ({
      ubicacion: u.ubicacion,
      cantidad: onlyDigits((document.getElementById("g_" + u.ubicacion)?.value || "0")),
    }));

    const payload = {
      kind: "general_page",
      ts: new Date().toISOString(),
      conteoId: GEN.conteoId,
      legajo: GEN.legajo,
      rol,
      pageIndex: sheetIndex,
      pageSize: GEN.pageSize,
      rows,
    };

    showOverlay(true);
    await fetch(APP_SCRIPT_URL, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload),
    });
    showOverlay(false);

    // marcar terminada localmente y volver al tracker
    GEN.status[rol][sheetIndex] = "terminado";
    renderTracker();
    show("general");
  }

  function initGeneralFromInputs() {
    GEN.legajo = onlyDigits($("g_legajo").value).slice(0, 3);
    GEN.conteoId = String($("g_conteoId").value || "").trim();
  }

  async function loadGeneralTracker() {
    initGeneralFromInputs();

    if (!GEN.legajo) {
      $("g_msg").textContent = "Ingresá legajo.";
      $("g_msg").className = "msg error";
      return;
    }
    if (!GEN.conteoId) {
      $("g_msg").textContent = "Ingresá ID Conteo.";
      $("g_msg").className = "msg error";
      return;
    }

    ensureGeneralListsBuilt();

    // intentar traer estado real del server
    showOverlay(true);
    await loadStatusFromServer();
    showOverlay(false);

    renderTracker();
    $("g_msg").textContent = "Listo.";
    $("g_msg").className = "msg ok";
  }

  $("g_load").addEventListener("click", loadGeneralTracker);

  $("g_back").addEventListener("click", () => {
    // volver al tracker
    show("general");
  });

  $("g_send").addEventListener("click", sendGeneralSheet);

  // init
  show("home");
})();
