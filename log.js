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
        "SBC": ["SBC", "REDES SBC"],
        "GRAJAÚ": ["GRAJA", "GRAJAÚ", "GRAJAU"],
        "FRANCO": ["FRANCO", "FRANCO DA ROCHA"],
        "SP": ["SP", "21 -", "23 -"]
      };

      const KEYWORDS = {
        [DIAG_ENDERECO_NAO_LOCALIZADO]: ["ENDERECO NAO LOCALIZADO", "ENDERECO NAO ENCONTRADO", "NAO LOCALIZADO", "NAO ENCONTRADO", "NÃO LOCALIZADO", "NÃO ENCONTRADO"],
        [DIAG_CLIENTE_AUSENTE]: ["CLIENTE AUSENTE", "AUSENTE", "NAO ATENDEU", "NÃO ATENDEU"],
        [DIAG_REAGENDADO_DEVIDO_HORARIO]: ["REAGENDADO DEVIDO HORARIO", "REAGENDADO DEVIDO HORÁRIO", "MUDOU HORARIO", "MUDOU HORÁRIO", "TROCA DE HORARIO", "TROCA DE HORÁRIO"],
        [DIAG_REAGENDADO_PELO_CLIENTE]: ["REAGENDADO PELO CLIENTE", "PELO CLIENTE", "PEDIDO DO CLIENTE", "CLIENTE SOLICITOU"]
      };

      let loadedData = [];
      let currentRegion = "SBC";

      const dropzone = document.getElementById('dropzone');
      const fileInput = document.getElementById('fileInput');
      const resultsDiv = document.getElementById('results');
      const actionsBar = document.getElementById('actionsBar');
      const regionBtns = document.querySelectorAll('.pill-btn');
      const btnCopyAll = document.getElementById('btnCopyAll');
      const btnReset = document.getElementById('btnReset');

      function normalizeStr(s) {
        if (s === null || s === undefined) return "";
        return String(s).normalize('NFD').replace(/[\u0300-\u036f]/g, "").toUpperCase().trim();
      }

      function findKey(keys, substr) {
        const ns = normalizeStr(substr);
        return keys.find(k => normalizeStr(k).includes(ns));
      }

      function detectStrictDiagnosis(row) {
        const keys = Object.keys(row);
        const diagKey = keys.find(k => normalizeStr(k).includes("DIAGNOSTICO") || normalizeStr(k).includes("DIAGNÓSTICO"));
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

        const assuntoKey = findKey(keys, "ASSUNTO");
        const msgKey = findKey(keys, "MENSAGEM") || findKey(keys, "MESSAGE");
        const assunto = assuntoKey ? normalizeStr(row[assuntoKey] || "") : "";
        const mensagem = msgKey ? normalizeStr(row[msgKey] || "") : "";
        const source = (assunto + " " + mensagem).trim();

        for (const kw of KEYWORDS[DIAG_ENDERECO_NAO_LOCALIZADO]) {
          if (source.includes(normalizeStr(kw))) return DIAG_ENDERECO_NAO_LOCALIZADO;
        }
        for (const kw of KEYWORDS[DIAG_CLIENTE_AUSENTE]) {
          if (source.includes(normalizeStr(kw))) return DIAG_CLIENTE_AUSENTE;
        }
        for (const kw of KEYWORDS[DIAG_REAGENDADO_DEVIDO_HORARIO]) {
          if (source.includes(normalizeStr(kw))) return DIAG_REAGENDADO_DEVIDO_HORARIO;
        }
        for (const kw of KEYWORDS[DIAG_REAGENDADO_PELO_CLIENTE]) {
          if (source.includes(normalizeStr(kw))) return DIAG_REAGENDADO_PELO_CLIENTE;
        }

        return null;
      }

      function detectRegion(row) {
        const keys = Object.keys(row);
        const setorKey = findKey(keys, "SETOR");
        const assuntoKey = findKey(keys, "ASSUNTO") || findKey(keys, "SUBJECT");
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
        const clienteKey = keys.find(k => normalizeStr(k).includes("CLIENTE") || normalizeStr(k).includes("NOME") || normalizeStr(k).includes("NAME"));
        const assuntoKey = keys.find(k => normalizeStr(k).includes("ASSUNTO") || normalizeStr(k).includes("SUBJECT"));
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

      function processFile(file) {
        const reader = new FileReader();
        reader.onload = (e) => {
          try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
            if (!json || !json.length) { alert("Arquivo vazio."); return; }
            loadedData = json.map(analyzeRow).filter(Boolean);
            actionsBar.style.display = loadedData.length ? "flex" : "none";
            render();
          } catch (err) {
            console.error(err);
            alert("Erro ao ler arquivo. Verifique se é um Excel (.xlsx) válido.");
          }
        };
        reader.readAsArrayBuffer(file);
      }

      function render() {
        resultsDiv.innerHTML = "";
        const filtered = loadedData.filter(item => {
          if (currentRegion === "SBC") return item.region === "SBC";
          return item.region === currentRegion;
        });
        if (!filtered.length) {
          resultsDiv.innerHTML = `<div class="empty-msg">Nenhum diagnóstico (dos 4) encontrado para ${currentRegion}.</div>`;
          return;
        }

        const order = [DIAG_REAGENDADO_PELO_CLIENTE, DIAG_REAGENDADO_DEVIDO_HORARIO, DIAG_CLIENTE_AUSENTE, DIAG_ENDERECO_NAO_LOCALIZADO];

        order.forEach(cat => {
          const items = filtered.filter(i => i.diag === cat);
          if (!items.length) return;
          const block = document.createElement('div');
          block.className = 'group-block';
          const header = document.createElement('div');
          header.className = 'group-header';
          header.textContent = cat;
          block.appendChild(header);
          items.forEach(it => {
            const el = document.createElement('div');
            el.className = 'list-item' + (it.isWarning ? ' warn-item' : '');
            el.innerHTML = `<span class="client-name">${escapeHtml(it.cliente)}</span><span class="client-subject">${escapeHtml(it.assunto)}</span>`;
            block.appendChild(el);
          });
          resultsDiv.appendChild(block);
        });
      }

      function escapeHtml(s) { return String(s).replace(/[&<>"']/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m])); }

      function buildCopyAllText() {
        if (!loadedData.length) return "";
        const preferred = ["SBC", "GRAJAÚ", "FRANCO", "SP", "OUTROS"];
        const groups = {};
        loadedData.forEach(r => { groups[r.region] = groups[r.region] || []; groups[r.region].push(r); });
        const presentRegions = Object.keys(groups).filter(k => groups[k] && groups[k].length);
        presentRegions.sort((a, b) => {
          const ia = preferred.indexOf(a), ib = preferred.indexOf(b);
          if (ia === -1 && ib === -1) return a.localeCompare(b);
          if (ia === -1) return 1;
          if (ib === -1) return -1;
          return ia - ib;
        });

        const order = [DIAG_REAGENDADO_PELO_CLIENTE, DIAG_REAGENDADO_DEVIDO_HORARIO, DIAG_CLIENTE_AUSENTE, DIAG_ENDERECO_NAO_LOCALIZADO];
        let out = "";
        presentRegions.forEach((reg, idx) => {
          out += `REAGENDAMENTO ${reg}\n\n`;
          order.forEach(cat => {
            const items = groups[reg].filter(i => i.diag === cat);
            if (!items.length) return; // skip empty category
            out += `${cat}:\n`;
            items.forEach(it => { out += `- ${it.cliente} - ${it.assunto}\n`; });
            out += `\n`;
          });
          if (idx !== presentRegions.length - 1) out += `-------------------\n\n`;
        });
        return out.trim();
      }

      async function copyAllHandler() {
        const text = buildCopyAllText();
        if (!text) return alert("Nada para copiar.");
        try {
          await navigator.clipboard.writeText(text);
          alert("Copiado!");
        } catch (e) {
          alert("Erro ao copiar.");
        }
      }

      dropzone.addEventListener('click', () => fileInput.click());
      fileInput.addEventListener('change', (e) => { if (e.target.files[0]) processFile(e.target.files[0]); fileInput.value = ""; });
      dropzone.addEventListener('dragover', (e) => { e.preventDefault(); dropzone.classList.add('drag-active'); });
      dropzone.addEventListener('dragleave', () => dropzone.classList.remove('drag-active'));
      dropzone.addEventListener('drop', (e) => { e.preventDefault(); dropzone.classList.remove('drag-active'); if (e.dataTransfer.files[0]) processFile(e.dataTransfer.files[0]); });

      regionBtns.forEach(btn => btn.addEventListener('click', () => {
        regionBtns.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        currentRegion = btn.dataset.target;
        render();
      }));

      btnCopyAll.addEventListener('click', copyAllHandler);

      btnReset.addEventListener('click', () => {
        loadedData = [];
        resultsDiv.innerHTML = `<div class="empty-msg">Aguardando arquivo...</div>`;
        actionsBar.style.display = 'none';
      });

    })();