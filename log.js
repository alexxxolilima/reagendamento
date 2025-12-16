(() => {
  const DIAG_REAGENDADO_PELO_CLIENTE = "REAGENDADO PELO CLIENTE";
  const DIAG_REAGENDADO_DEVIDO_HORARIO = "REAGENDADO DEVIDO HORARIO";
  const DIAG_CLIENTE_AUSENTE = "CLIENTE AUSENTE";
  const DIAG_ENDERECO_NAO_LOCALIZADO = "ENDEREÇO NÃO LOCALIZADO";
  const DIAGS = [
    DIAG_REAGENDADO_PELO_CLIENTE,
    DIAG_REAGENDADO_DEVIDO_HORARIO,
    DIAG_CLIENTE_AUSENTE,
    DIAG_ENDERECO_NAO_LOCALIZADO
  ];
  const REGION_MAP = {
    SBC: ["SBC", "REDES SBC"],
    "GRAJAÚ": ["GRAJA", "GRAJAÚ", "GRAJAU"],
    FRANCO: ["FRANCO", "FRANCO DA ROCHA"],
    SP: ["SP", "21 -", "23 -"]
  };
  const KEYWORDS = {};
  KEYWORDS[DIAG_ENDERECO_NAO_LOCALIZADO] = ["ENDERECO NAO LOCALIZADO", "ENDERECO NAO ENCONTRADO", "NAO LOCALIZADO", "NAO ENCONTRADO", "NÃO LOCALIZADO", "NÃO ENCONTRADO"];
  KEYWORDS[DIAG_CLIENTE_AUSENTE] = ["CLIENTE AUSENTE", "AUSENTE", "NAO ATENDEU", "NÃO ATENDEU"];
  KEYWORDS[DIAG_REAGENDADO_DEVIDO_HORARIO] = ["REAGENDADO DEVIDO HORARIO", "REAGENDADO DEVIDO HORÁRIO", "MUDOU HORARIO", "MUDOU HORÁRIO", "TROCA DE HORARIO", "TROCA DE HORÁRIO"];
  KEYWORDS[DIAG_REAGENDADO_PELO_CLIENTE] = ["REAGENDADO PELO CLIENTE", "PELO CLIENTE", "PEDIDO DO CLIENTE", "CLIENTE SOLICITOU"];
  const FIELD_MAP = {
    cliente: ["CLIENTE", "NOME", "NAME"],
    assunto: ["ASSUNTO", "SUBJECT"],
    setor: ["SETOR"],
    mensagem: ["MENSAGEM", "MESSAGE"],
    diagnostico: ["DIAGNOSTICO", "DIAGNÓSTICO", "DIAG"]
  };
  const MAX_FILE_MB = 12;
  const MAX_RENDER_PER_REGION = 300;
  let loadedData = [];
  let currentRegion = "SBC";
  const dropzone = document.getElementById("dropzone");
  const fileInput = document.getElementById("fileInput");
  const resultsDiv = document.getElementById("results");
  const actionsBar = document.getElementById("actionsBar");
  const regionBtns = document.querySelectorAll(".pill-btn");
  const btnCopyAll = document.getElementById("btnCopyAll");
  const btnReset = document.getElementById("btnReset");
  const loaderEl = document.getElementById("loader");
  const rowCountEl = document.getElementById("rowCount");
  function normalizeStr(s) {
    if (s === null || s === undefined) return "";
    return String(s).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().trim();
  }
  function findFieldKey(rowKeys, logical) {
    const candidates = FIELD_MAP[logical];
    if (!candidates) return null;
    for (const k of rowKeys) {
      const nk = normalizeStr(k);
      for (const c of candidates) {
        if (nk.includes(normalizeStr(c))) return k;
      }
    }
    return null;
  }
  function setLoading(on) {
    if (on) {
      loaderEl.hidden = false;
      rowCountEl.hidden = true;
    } else {
      loaderEl.hidden = true;
    }
  }
  function setRowCount(n) {
    rowCountEl.hidden = false;
    rowCountEl.textContent = `${n} registro(s) válidos encontrados`;
  }
  function isValidFile(file) {
    if (!file) return false;
    if (file.size > MAX_FILE_MB * 1024 * 1024) {
      alert("Arquivo muito grande (> " + MAX_FILE_MB + "MB).");
      return false;
    }
    const allowedTypes = [
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "application/vnd.ms-excel"
    ];
    if (!allowedTypes.includes(file.type) && !file.name.match(/\.(xlsx|xls)$/i)) {
      alert("Selecione um arquivo .xlsx ou .xls");
      return false;
    }
    return true;
  }
  function detectStrictDiagnosis(row) {
    const keys = Object.keys(row);
    const diagKey = findFieldKey(keys, "diagnostico");
    if (diagKey) {
      const val = normalizeStr(row[diagKey] || "");
      if (val) {
        if (val.includes(normalizeStr(DIAG_REAGENDADO_PELO_CLIENTE))) return DIAG_REAGENDADO_PELO_CLIENTE;
        if (val.includes(normalizeStr(DIAG_REAGENDADO_DEVIDO_HORARIO))) return DIAG_REAGENDADO_DEVIDO_HORARIO;
        if (val.includes(normalizeStr(DIAG_CLIENTE_AUSENTE))) return DIAG_CLIENTE_AUSENTE;
        if (val.includes(normalizeStr(DIAG_ENDERECO_NAO_LOCALIZADO))) return DIAG_ENDERECO_NAO_LOCALIZADO;
        return null;
      }
    }
    const assuntoKey = findFieldKey(keys, "assunto");
    const msgKey = findFieldKey(keys, "mensagem");
    const assunto = assuntoKey ? normalizeStr(row[assuntoKey] || "") : "";
    const mensagem = msgKey ? normalizeStr(row[msgKey] || "") : "";
    const source = (assunto + " " + mensagem).trim();
    for (const kw of KEYWORDS[DIAG_ENDERECO_NAO_LOCALIZADO]) if (source.includes(normalizeStr(kw))) return DIAG_ENDERECO_NAO_LOCALIZADO;
    for (const kw of KEYWORDS[DIAG_CLIENTE_AUSENTE]) if (source.includes(normalizeStr(kw))) return DIAG_CLIENTE_AUSENTE;
    for (const kw of KEYWORDS[DIAG_REAGENDADO_DEVIDO_HORARIO]) if (source.includes(normalizeStr(kw))) return DIAG_REAGENDADO_DEVIDO_HORARIO;
    for (const kw of KEYWORDS[DIAG_REAGENDADO_PELO_CLIENTE]) if (source.includes(normalizeStr(kw))) return DIAG_REAGENDADO_PELO_CLIENTE;
    return null;
  }
  function detectRegion(row) {
    const keys = Object.keys(row);
    const setorKey = findFieldKey(keys, "setor");
    const assuntoKey = findFieldKey(keys, "assunto");
    const setor = setorKey ? normalizeStr(row[setorKey] || "") : "";
    if (setor) {
      for (const r of Object.keys(REGION_MAP)) {
        for (const term of REGION_MAP[r]) {
          if (setor.includes(normalizeStr(term))) return r;
        }
      }
    }
    const assunto = assuntoKey ? String(row[assuntoKey] || "").trim() : "";
    const m = assunto.match(/^(\d+)\s*-/);
    if (m) {
      const n = parseInt(m[1], 10);
      if (n === 1) return "SBC";
      if (n === 21 || n === 23) return "SP";
    }
    return "OUTROS";
  }
  function extractFields(row) {
    const keys = Object.keys(row);
    const clienteKey = findFieldKey(keys, "cliente");
    const assuntoKey = findFieldKey(keys, "assunto");
    const cliente = clienteKey ? String(row[clienteKey] || "").trim() : "Cliente Desconhecido";
    let assunto = assuntoKey ? String(row[assuntoKey] || "").trim() : "";
    assunto = assunto.replace(/^\s*\d+\s*-\s*/, "");
    return { cliente, assunto };
  }
  function analyzeRow(row) {
    const diag = detectStrictDiagnosis(row);
    if (!diag) return null;
    const { cliente, assunto } = extractFields(row);
    const region = detectRegion(row);
    const isWarning = diag === DIAG_ENDERECO_NAO_LOCALIZADO;
    return { cliente, assunto, region, diag, isWarning };
  }
  function statsLog() {
    const s = {};
    DIAGS.forEach(d => (s[d] = loadedData.filter(i => i.diag === d).length));
    const byRegion = {};
    loadedData.forEach(i => {
      byRegion[i.region] = (byRegion[i.region] || 0) + 1;
    });
    console.info("Resumo por diagnóstico:", s);
    console.info("Resumo por região:", byRegion);
  }
  function selectSheetName(names) {
    if (!names || !names.length) return null;
    if (names.length === 1) return names[0];
    const choice = prompt("Planilhas: \n" + names.join("\n") + "\n\nCole o nome da planilha ou deixe vazio para a primeira:", names[0]);
    return choice && names.includes(choice) ? choice : names[0];
  }
  async function processFile(file) {
    if (!isValidFile(file)) return;
    setLoading(true);
    try {
      const data = new Uint8Array(await file.arrayBuffer());
      const workbook = XLSX.read(data, { type: "array" });
      const sheetName = selectSheetName(workbook.SheetNames);
      if (!sheetName) {
        alert("");
        setLoading(false);
        return;
      }
      const sheet = workbook.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      if (!json || !json.length) {
        alert("Arquivo vazio.");
        setLoading(false);
        return;
      }
      loadedData = json.map(analyzeRow).filter(Boolean);
      actionsBar.hidden = !loadedData.length;
      setRowCount(loadedData.length);
      render();
      statsLog();
    } catch (err) {
      console.error("Erro processando arquivo:", err);
      alert("Erro ao ler arquivo. Veja console para detalhes.");
    } finally {
      setLoading(false);
    }
  }
  function render() {
    resultsDiv.innerHTML = "";
    const filtered = loadedData.filter(item => {
      if (currentRegion === "SBC") return item.region === "SBC";
      return item.region === currentRegion;
    });
    if (!filtered.length) {
      resultsDiv.innerHTML = `<div class="empty-msg">Nenhum reagendamento em ${currentRegion}.</div>`;
      return;
    }
    const order = [
      DIAG_REAGENDADO_PELO_CLIENTE,
      DIAG_REAGENDADO_DEVIDO_HORARIO,
      DIAG_CLIENTE_AUSENTE,
      DIAG_ENDERECO_NAO_LOCALIZADO
    ];
    order.forEach(cat => {
      const items = filtered.filter(i => i.diag === cat);
      if (!items.length) return;
      const block = document.createElement("div");
      block.className = "group-block";
      const header = document.createElement("div");
      header.className = "group-header";
      header.textContent = cat;
      block.appendChild(header);
      const toRender = items.slice(0, MAX_RENDER_PER_REGION);
      toRender.forEach(it => {
        const el = document.createElement("div");
        el.className = "list-item" + (it.isWarning ? " warn-item" : "");
        const nameEl = document.createElement("div");
        nameEl.className = "client-name";
        nameEl.textContent = it.cliente;
        const subEl = document.createElement("div");
        subEl.className = "client-subject";
        subEl.textContent = it.assunto;
        el.appendChild(nameEl);
        el.appendChild(subEl);
        block.appendChild(el);
      });
      if (items.length > MAX_RENDER_PER_REGION) {
        const moreBtn = document.createElement("button");
        moreBtn.className = "show-more";
        moreBtn.textContent = `Mostrar mais (${items.length - MAX_RENDER_PER_REGION})`;
        moreBtn.addEventListener("click", () => {
          items.slice(MAX_RENDER_PER_REGION).forEach(it => {
            const el = document.createElement("div");
            el.className = "list-item" + (it.isWarning ? " warn-item" : "");
            const nameEl = document.createElement("div");
            nameEl.className = "client-name";
            nameEl.textContent = it.cliente;
            const subEl = document.createElement("div");
            subEl.className = "client-subject";
            subEl.textContent = it.assunto;
            el.appendChild(nameEl);
            el.appendChild(subEl);
            block.appendChild(el);
          });
          moreBtn.remove();
        });
        block.appendChild(moreBtn);
      }
      resultsDiv.appendChild(block);
    });
  }
  function buildCopyAllText() {
    if (!loadedData.length) return "";
    const preferred = ["SBC", "GRAJAÚ", "FRANCO", "SP", "OUTROS"];
    const groups = {};
    loadedData.forEach(r => {
      groups[r.region] = groups[r.region] || [];
      groups[r.region].push(r);
    });
    const presentRegions = Object.keys(groups).filter(k => groups[k] && groups[k].length);
    presentRegions.sort((a, b) => {
      const ia = preferred.indexOf(a),
        ib = preferred.indexOf(b);
      if (ia === -1 && ib === -1) return a.localeCompare(b);
      if (ia === -1) return 1;
      if (ib === -1) return -1;
      return ia - ib;
    });
    const order = [
      DIAG_REAGENDADO_PELO_CLIENTE,
      DIAG_REAGENDADO_DEVIDO_HORARIO,
      DIAG_CLIENTE_AUSENTE,
      DIAG_ENDERECO_NAO_LOCALIZADO
    ];
    let out = "";
    presentRegions.forEach((reg, idx) => {
      out += `REAGENDAMENTO ${reg}\n\n`;
      order.forEach(cat => {
        const items = groups[reg].filter(i => i.diag === cat);
        if (!items.length) return;
        out += `${cat}:\n`;
        items.forEach(it => {
          out += `- ${it.cliente} - ${it.assunto}\n`;
        });
        out += `\n`;
      });
      if (idx !== presentRegions.length - 1) out += `-------------------\n\n`;
    });
    return out.trim();
  }
  async function writeClipboard(text) {
    if (!text) return false;
    if (navigator.clipboard && navigator.clipboard.writeText) return navigator.clipboard.writeText(text);
    const ta = document.createElement("textarea");
    ta.value = text;
    ta.style.position = "fixed";
    ta.style.left = "-9999px";
    document.body.appendChild(ta);
    ta.select();
    try {
      document.execCommand("copy");
    } finally {
      ta.remove();
    }
    return Promise.resolve();
  }
  async function copyAllHandler() {
    const text = buildCopyAllText();
    if (!text) return alert("Nada para copiar.");
    try {
      await writeClipboard(text);
      alert("Copiado!");
    } catch (e) {
      alert("Erro ao copiar.");
    }
  }
  function downloadTxt(content, filename) {
    if (!content) return;
    const blob = new Blob([content], { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename || "reagendamento.txt";
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }
  dropzone.addEventListener("click", () => fileInput.click());
  dropzone.addEventListener("keydown", e => {
    if (e.key === "Enter" || e.key === " ") fileInput.click();
  });
  fileInput.addEventListener("change", e => {
    if (e.target.files[0]) processFile(e.target.files[0]);
    fileInput.value = "";
  });
  dropzone.addEventListener("dragover", e => {
    e.preventDefault();
    dropzone.classList.add("drag-active");
  });
  dropzone.addEventListener("dragleave", () => dropzone.classList.remove("drag-active"));
  dropzone.addEventListener("drop", e => {
    e.preventDefault();
    dropzone.classList.remove("drag-active");
    if (e.dataTransfer.files[0]) processFile(e.dataTransfer.files[0]);
  });
  regionBtns.forEach(btn =>
    btn.addEventListener("click", () => {
      regionBtns.forEach(b => b.classList.remove("active"));
      btn.classList.add("active");
      currentRegion = btn.dataset.target;
      render();
    })
  );
  btnCopyAll.addEventListener("click", copyAllHandler);

  if (btnReset) {
    btnReset.addEventListener("click", () => {
      loadedData = [];
      resultsDiv.innerHTML = `<div class="empty-msg">Aguardando arquivo...</div>`;
      actionsBar.hidden = true;
      rowCountEl.hidden = true;
    });
  }
  window.__reagendamento = {
    getData: () => loadedData,
    stats: statsLog
  };
})();
