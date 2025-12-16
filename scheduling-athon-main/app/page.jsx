"use client"

import Image from "next/image";
import { useState, useCallback, useMemo } from "react";
import { useDropzone } from "react-dropzone";
import * as XLSX from "xlsx";
import { FileSpreadsheet, Copy, CheckCircle2 } from "lucide-react";

const DIAG_KEYS = [
  "CLIENTE AUSENTE",
  "ENDEREÇO NÃO LOCALIZADO",
  "REAGENDADO PELO CLIENTE",
  "REAGENDADO DEVIDO HORARIO",
  "TÉCNICO NÃO FOI ATENDIDO",
  "TÉCNICO NÃO ENCONTROU A RESIDÊNCIA",
];

function normalize(str) {
  return (str || "").toString().trim().toUpperCase();
}

// heurística para detectar diagnóstico: procura por substrings em Assunto/Mensagem
function detectDiagnosis(row) {
  const assunto = normalize(row.Assunto);
  const mensagem = normalize(row.Mensagem);
  for (const key of DIAG_KEYS) {
    if (assunto.includes(key) || mensagem.includes(key)) return key;
  }
  // heurísticas adicionais (palavras-chave)
  if (assunto.includes("CLIENTE AUSENTE") || mensagem.includes("CLIENTE AUSENTE")) return "CLIENTE AUSENTE";
  if (assunto.includes("ENDEREÇO NÃO LOCALIZADO") || mensagem.includes("ENDEREÇO NÃO LOCALIZADO")) return "ENDEREÇO NÃO LOCALIZADO";
  // fallback: se contém "REAGEND" assume reagendamento
  if (assunto.includes("REAGEND") || mensagem.includes("REAGEND")) {
    if (assunto.includes("PELO CLIENTE") || mensagem.includes("PELO CLIENTE")) return "REAGENDADO PELO CLIENTE";
    if (assunto.includes("HORARIO") || mensagem.includes("HORÁRIO")) return "REAGENDADO DEVIDO HORARIO";
    return "REAGENDADO PELO CLIENTE";
  }
  return null;
}

// extrai prefixo numérico do Assunto; útil para mapear setores numéricos (ex: "1-" => 1)
function extractAssuntoPrefix(row) {
  const assunto = (row.Assunto || "").toString().trim();
  const m = assunto.match(/^(\d+)\s*-/);
  return m ? parseInt(m[1], 10) : null;
}

// mapeamento por setor numérico para agrupamentos SBC / SP (ajuste conforme regra real)
const NUMERIC_SECTOR_TO_REGION = {
  1: "SBC",
  21: "SP",
  23: "SP",
  // adicione outros mapeamentos se necessário
};

export default function Page() {
  const [data, setData] = useState([]);
  const [region, setRegion] = useState("SBC");
  const [copyMessage, setCopyMessage] = useState("");

  // processa raw rows: padroniza campos, detecta diagnóstico e região baseada em Setor / prefixo
  const processedData = useMemo(() => {
    return data.map((r) => {
      const row = { ...r };
      row._Cliente = row.Cliente ? row.Cliente.toString().trim() : "";
      row._Assunto = row.Assunto ? row.Assunto.toString().trim() : "";
      row._Setor = row.Setor ? row.Setor.toString().trim() : "";
      row._diagnosis = detectDiagnosis(row) || "";
      const prefix = extractAssuntoPrefix(row);
      row._assuntoPrefix = prefix;
      // derive regionCandidate: from Setor text or numeric prefix map
      let regionCandidate = "";
      const s = row._Setor.toUpperCase();
      if (s.includes("SBC")) regionCandidate = "SBC";
      else if (s.includes("GRAJA")) regionCandidate = "GRAJAÚ";
      else if (s.includes("FRANCO")) regionCandidate = "FRANCO";
      // numeric mapping fallback
      if (!regionCandidate && prefix && NUMERIC_SECTOR_TO_REGION[prefix]) {
        regionCandidate = NUMERIC_SECTOR_TO_REGION[prefix];
      }
      row._region = regionCandidate || row._Setor || "";
      return row;
    });
  }, [data]);

  const filterDataByRegion = useCallback(
    (rows, regionName) => {
      const target = (regionName || "").toUpperCase();
      return rows.filter((r) => {
        const rRegion = (r._region || "").toString().toUpperCase();
        // match by inclusion (handles 'REDES SBC' etc.)
        if (rRegion.includes(target)) return true;
        // also check Setor direct include
        if ((r._Setor || "").toString().toUpperCase().includes(target)) return true;
        return false;
      });
    },
    []
  );

  const groupByRegionAndDiagnosis = useCallback((rows) => {
    // produce structure:
    // { regionName: { 'Técnico não foi atendido': [rows], 'REAGENDADO PELO CLIENTE': [rows], ... } }
    const out = {};
    rows.forEach((r) => {
      const regionName = r._region || "OUTROS";
      out[regionName] = out[regionName] || {};
      // normalize diagnosis to the set we want to present
      const diag = r._diagnosis || detectDiagnosis(r) || "OUTROS";
      // map some diag labels to presentation groups expected
      let groupKey = diag;
      if (groupKey === "CLIENTE AUSENTE") groupKey = "Técnico não foi atendido";
      if (groupKey === "ENDEREÇO NÃO LOCALIZADO") groupKey = "Técnico não encontrou a residência";
      out[regionName][groupKey] = out[regionName][groupKey] || [];
      out[regionName][groupKey].push(r);
    });
    return out;
  }, []);

  // dropzone + leitura Excel (usando ArrayBuffer por compatibilidade)
  const onDrop = useCallback(
    (acceptedFiles) => {
      const file = acceptedFiles[0];
      if (!file) return;
      if (!file.name.endsWith(".xlsx") && !file.name.endsWith(".xls")) {
        alert("Por favor, selecione um arquivo Excel.");
        return;
      }
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const ab = e.target.result;
          const workBook = XLSX.read(ab, { type: "array" });
          const worksheet = workBook.Sheets[workBook.SheetNames[0]];
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
          setData(jsonData);
        } catch (err) {
          console.error("Erro ao processar Excel:", err);
          alert("Erro ao processar o arquivo Excel.");
        }
      };
      reader.readAsArrayBuffer(file);
    },
    []
  );

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    multiple: false,
    accept: {
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
      "application/vnd.ms-excel": [".xls"],
    },
  });

  const filteredData = useMemo(() => filterDataByRegion(processedData, region), [processedData, filterDataByRegion, region]);

  const handleCopyFiltered = () => {
    const text = filteredData
      .map((row) => `- ${row._Cliente} - ${row._Assunto?.slice(2) || row._Assunto}\n`)
      .join("");
    navigator.clipboard.writeText(text);
    setCopyMessage("Copiado!");
    setTimeout(() => setCopyMessage(""), 2000);
  };

  const handleCopyAllGrouped = () => {
    const grouped = groupByRegionAndDiagnosis(processedData);
    // build text in requested format
    let text = "";
    // order: SBC then SP per your example
    const order = ["SBC", "SP", "GRAJAÚ", "FRANCO"];
    const keys = order.concat(Object.keys(grouped).filter(k => !order.includes(k)));
    keys.forEach((reg) => {
      if (!grouped[reg]) return;
      text += `- REAGENDAMENTO ${reg}-\n\n`;
      // expected blocks: Técnico não foi atendido, REAGENDADO PELO CLIENTE, REAGENDADO DEVIDO HORARIO, Técnico não encontrou a residência
      const blocks = [
        "Técnico não foi atendido",
        "REAGENDADO PELO CLIENTE",
        "REAGENDADO DEVIDO HORARIO",
        "Técnico não encontrou a residência",
      ];
      blocks.forEach((blk) => {
        const arr = grouped[reg][blk] || [];
        text += `${blk}:\n`;
        if (arr.length === 0) {
          text += "(sem clientes)\n\n";
        } else {
          arr.forEach((r) => {
            text += `- ${r._Cliente} - ${r._Assunto?.slice(2) || r._Assunto}\n`;
          });
          text += "\n";
        }
      });
      text += "\n";
    });
    navigator.clipboard.writeText(text);
    setCopyMessage("Todos copiados!");
    setTimeout(() => setCopyMessage(""), 2000);
  };

  const handleRegionChange = useCallback((selectedRegion) => {
    setCopyMessage("");
    setRegion(selectedRegion);
  }, []);

  const handleReset = () => {
    setData([]);
    setRegion("SBC");
    setCopyMessage("");
  };

  return (
    <div className="min-h-screen text-white">
      <div className="flex items-center justify-center relative">
        <div className="flex items-center gap-3 justify-center mt-10">
          <Image src={"/athonfav.png"} width={30} height={30} alt="Icon Athon Telecom" />
          <h1 className="text-xl">Reagendamento Ausentes</h1>
        </div>
      </div>

      <div
        {...getRootProps()}
        className={`mt-8 mb-10 border-2 border-dashed rounded-lg p-8 text-center cursor-pointer transition-all duration-200 ease-in-out ${isDragActive ? "border-cyan-400 bg-blue-800/50" : "border-white/30 hover:border-white/50"}`}
      >
        <input {...getInputProps()} />
        <FileSpreadsheet className="w-10 h-10 mx-auto mb-4" />
        <p className="text-base">{isDragActive ? "Solte o arquivo aqui..." : "Arraste e solte o arquivo Excel aqui ou clique para selecionar"}</p>
      </div>

      <div className="flex space-x-4 items-center justify-center mb-8">
        {["SBC", "GRAJAÚ", "FRANCO"].map((regionName) => (
          <button key={regionName} onClick={() => handleRegionChange(regionName)} className={`px-6 py-2 rounded-full transition-all cursor-pointer duration-200 text-sm ${region === regionName ? "bg-zinc-800" : "bg-zinc-900/50"}`}>
            {regionName}
          </button>
        ))}
        <button onClick={handleReset} className="px-4 py-2 rounded-full bg-red-700/60 text-sm">Reset</button>
      </div>

      <div className="max-w-3xl mx-auto">
        {/* Actions */}
        <div className="flex items-center justify-center gap-4 mb-6">
          <button onClick={handleCopyFiltered} className="flex items-center gap-2 px-6 py-2 bg-zinc-900 hover:bg-zinc-700 rounded-full transition-colors cursor-pointer">
            <Copy className="w-4 h-4" />
            Copiar Lista Filtrada
          </button>
          <button onClick={handleCopyAllGrouped} className="flex items-center gap-2 px-6 py-2 bg-zinc-900 hover:bg-zinc-700 rounded-full transition-colors cursor-pointer">
            <Copy className="w-4 h-4" />
            Copiar Todos (Agrupado)
          </button>
          {copyMessage && (
            <span className="flex items-center gap-2 text-green-300">
              <CheckCircle2 className="w-4 h-4" />
              {copyMessage}
            </span>
          )}
        </div>

        {/* Render agrupado por region e diagnóstico */}
        {Object.keys(groupByRegionAndDiagnosis(processedData)).length > 0 ? (
          Object.entries(groupByRegionAndDiagnosis(processedData)).map(([reg, groups]) => (
            <div key={reg} className="rounded-lg p-6 mb-6 bg-zinc-900/20">
              <h2 className="text-lg font-semibold mb-3">REAGENDAMENTO {reg}</h2>

              {["Técnico não foi atendido", "REAGENDADO PELO CLIENTE", "REAGENDADO DEVIDO HORARIO", "Técnico não encontrou a residência"].map((blk) => (
                <div key={blk} className="mb-4">
                  <h3 className="font-medium">{blk}:</h3>
                  <ul className="list-disc pl-6 mt-1">
                    {(groups[blk] || []).length > 0 ? (
                      (groups[blk] || []).map((r, i) => (
                        <li key={i} className="text-sm">
                          <strong>{r._Cliente}</strong> — {r._Assunto?.slice(2) || r._Assunto}
                        </li>
                      ))
                    ) : (
                      <li className="text-sm text-zinc-400">Nenhum cliente</li>
                    )}
                  </ul>
                </div>
              ))}
            </div>
          ))
        ) : (
          <p className="text-center text-base text-orange-300">Nenhum dado para exibir.</p>
        )}
      </div>

      <div className="fixed bottom-8 right-8 p-4 rounded-lg">
        <div className="flex items-center gap-3 mb-2">
          <span className="w-4 h-4 bg-blue-800/30 border border-white/30 rounded bg-white"></span>
          <p>Cliente Ausente</p>
        </div>
        <div className="flex items-center gap-3">
          <span className="w-4 h-4 text-orange-400 border border-orange-500/30 rounded bg-orange-500"></span>
          <p>Endereço não localizado</p>
        </div>
      </div>
    </div>
  );
}
