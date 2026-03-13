import { useState, useRef } from "react";

// ─── Parsers ────────────────────────────────────────────────────────────────
function parseCSV(text) {
  const lines = text.trim().split(/\r?\n/);
  if (!lines.length) return [];
  const headers = lines[0].split(",").map(h => h.trim().replace(/^"|"$/g, ""));
  return lines.slice(1).map(line => {
    const vals = line.match(/(".*?"|[^,]+|(?<=,)(?=,)|(?<=,)$|^(?=,))/g) || [];
    const row = {};
    headers.forEach((h, i) => { row[h] = (vals[i] || "").replace(/^"|"$/g, "").trim(); });
    return row;
  });
}
function parseTSV(text) {
  const lines = text.trim().split(/\r?\n/);
  if (!lines.length) return [];
  const headers = lines[0].split("\t").map(h => h.trim());
  return lines.slice(1).map(line => {
    const vals = line.split("\t");
    const row = {};
    headers.forEach((h, i) => { row[h] = (vals[i] || "").trim(); });
    return row;
  });
}
function parsePipeDel(text) {
  const lines = text.trim().split(/\r?\n/);
  if (!lines.length) return [];
  const headers = lines[0].split("|").map(h => h.trim());
  return lines.slice(1).map(line => {
    const vals = line.split("|");
    const row = {};
    headers.forEach((h, i) => { row[h] = (vals[i] || "").trim(); });
    return row;
  });
}
function parsePlainText(text) {
  const firstLine = text.split(/\r?\n/)[0];
  if (firstLine.includes("|")) return parsePipeDel(text);
  if (firstLine.includes("\t")) return parseTSV(text);
  return parseCSV(text);
}
async function parseFile(file) {
  return new Promise((resolve, reject) => {
    const ext = file.name.split(".").pop().toLowerCase();
    const reader = new FileReader();
    if (["xlsx", "xls"].includes(ext)) {
      reader.onload = async (e) => {
        try {
          const { read, utils } = await import("https://cdn.sheetjs.com/xlsx-0.20.1/package/xlsx.mjs");
          const wb = read(new Uint8Array(e.target.result), { type: "array" });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const data = utils.sheet_to_json(ws, { defval: "" });
          resolve(data.map(r => Object.fromEntries(Object.entries(r).map(([k, v]) => [String(k).trim(), String(v).trim()]))));
        } catch (err) { reject(err); }
      };
      reader.readAsArrayBuffer(file);
    } else if (ext === "pdf") {
      reader.onload = async (e) => {
        try {
          const pdfjsLib = await import("https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js");
          pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
          const pdfDoc = await pdfjsLib.getDocument({ data: new Uint8Array(e.target.result) }).promise;
          let allText = "";
          for (let i = 1; i <= pdfDoc.numPages; i++) {
            const page = await pdfDoc.getPage(i);
            const content = await page.getTextContent();
            allText += content.items.map(it => it.str).join(" ") + "\n";
          }
          resolve(parsePlainText(allText));
        } catch { resolve([]); }
      };
      reader.readAsArrayBuffer(file);
    } else {
      reader.onload = (e) => {
        const text = e.target.result;
        if (ext === "tsv") resolve(parseTSV(text));
        else if (ext === "txt" || ext === "tex") resolve(parsePlainText(text));
        else resolve(parseCSV(text));
      };
      reader.readAsText(file);
    }
  });
}

function suggestMappings(headersA, headersB) {
  const normalize = s => s.toLowerCase().replace(/[^a-z0-9]/g, "");
  const suggestions = [];
  const usedB = new Set();
  for (const ha of headersA) {
    const nha = normalize(ha);
    let best = null, bestScore = 0;
    for (const hb of headersB) {
      if (usedB.has(hb)) continue;
      const nhb = normalize(hb);
      const score = nha === nhb ? 100 : (nha.includes(nhb) || nhb.includes(nha)) ? 80 : 0;
      if (score > bestScore) { bestScore = score; best = hb; }
    }
    if (best && bestScore >= 80) { suggestions.push({ from: ha, to: best }); usedB.add(best); }
  }
  return suggestions;
}

function compareDatasets(dataA, dataB, keyFields, currentFieldA, currentFieldB, tolerance) {
  const tolPct = parseFloat(tolerance) / 100 || 0;
  const indexB = {};
  for (const row of dataB) {
    const key = keyFields.map(k => (row[k] || "").toString().toLowerCase().trim()).join("||");
    if (!indexB[key]) indexB[key] = [];
    indexB[key].push(row);
  }
  const results = [];
  const matchedKeysB = new Set();

  for (const rowA of dataA) {
    const key = keyFields.map(k => (rowA[k] || "").toString().toLowerCase().trim()).join("||");
    const matchesB = indexB[key] || [];
    const keyVals = Object.fromEntries(keyFields.map(k => [k, rowA[k] ?? ""]));
    if (!matchesB.length) {
      results.push({ key, rowA, rowB: null, status: "Only in A", details: [], keyVals });
    } else {
      const rowB = matchesB[0];
      matchedKeysB.add(key);
      const details = [];
      if (currentFieldA && currentFieldB) {
        const valA = rowA[currentFieldA] ?? "";
        const valB = rowB[currentFieldB] ?? "";
        const numA = parseFloat(valA), numB = parseFloat(valB);
        const isNum = !isNaN(numA) && !isNaN(numB);
        let diff = "", status = "Matched";
        if (isNum) {
          const pct = numA !== 0 ? Math.abs(numA - numB) / Math.abs(numA) : (numB !== 0 ? 1 : 0);
          diff = (numB - numA).toFixed(2);
          if (pct > tolPct) status = "Mismatched";
        } else if (valA.toLowerCase().trim() !== valB.toLowerCase().trim()) {
          diff = `${valA} → ${valB}`; status = "Mismatched";
        }
        details.push({ fieldA: currentFieldA, fieldB: currentFieldB, valA, valB, diff, status });
      }
      const overallStatus = details.some(d => d.status === "Mismatched") ? "Mismatched" : "Matched";
      results.push({ key, rowA, rowB, status: overallStatus, details, keyVals });
    }
  }
  for (const row of dataB) {
    const key = keyFields.map(k => (row[k] || "").toString().toLowerCase().trim()).join("||");
    if (!matchedKeysB.has(key)) {
      results.push({ key, rowA: null, rowB: row, status: "Only in B", details: [], keyVals: Object.fromEntries(keyFields.map(k => [k, row[k] ?? ""])) });
      matchedKeysB.add(key);
    }
  }
  const keyCounts = {};
  for (const row of dataA) {
    const key = keyFields.map(k => (row[k] || "").toString().toLowerCase().trim()).join("||");
    keyCounts[key] = (keyCounts[key] || 0) + 1;
  }
  return { results, duplicatesA: Object.values(keyCounts).filter(c => c > 1).reduce((s, c) => s + c - 1, 0) };
}

async function exportToExcel(results, keyFields, currentFieldA, currentFieldB, fileAName, fileBName, dataALen, dataBLen) {
  const { utils, writeFile } = await import("https://cdn.sheetjs.com/xlsx-0.20.1/package/xlsx.mjs");
  const matched = results.filter(r => r.status === "Matched").length;
  const mismatched = results.filter(r => r.status === "Mismatched").length;
  const onlyA = results.filter(r => r.status === "Only in A").length;
  const onlyB = results.filter(r => r.status === "Only in B").length;

  const summaryData = [
    ["CompareIQ — Comparison Summary", "", ""],
    [],
    ["File A", fileAName], ["File B", fileBName],
    ["Key Fields", keyFields.join(", ")],
    ["Comparison Field (A)", currentFieldA || "N/A"],
    ["Comparison Field (B)", currentFieldB || "N/A"],
    [],
    ["Metric", "Count", ""],
    ["Total Records in A", dataALen],
    ["Total Records in B", dataBLen],
    ["Matched", matched],
    ["Mismatched", mismatched],
    ["Only in A", onlyA],
    ["Only in B", onlyB],
    ["Match Rate", `${((matched / (matched + mismatched || 1)) * 100).toFixed(1)}%`],
  ];

  const detailHeaders = [...keyFields, `${currentFieldA} (A)`, `${currentFieldB} (B)`, "Difference", "Status"];
  const detailRows = results.map(r => {
    const d = r.details[0];
    return [
      ...keyFields.map(k => r.keyVals[k] ?? ""),
      d ? d.valA : (r.rowA ? r.rowA[currentFieldA] ?? "" : ""),
      d ? d.valB : (r.rowB ? r.rowB[currentFieldB] ?? "" : ""),
      d ? d.diff : "",
      r.status,
    ];
  });

  const allHeadersA = results.find(r => r.rowA) ? Object.keys(results.find(r => r.rowA).rowA) : [];
  const allHeadersB = results.find(r => r.rowB) ? Object.keys(results.find(r => r.rowB).rowB) : [];
  const sheetHeaders = [...allHeadersA.map(h => `A: ${h}`), ...allHeadersB.map(h => `B: ${h}`), "Status"];
  const sheetRows = results.map(r => [
    ...allHeadersA.map(h => r.rowA ? r.rowA[h] ?? "" : ""),
    ...allHeadersB.map(h => r.rowB ? r.rowB[h] ?? "" : ""),
    r.status,
  ]);

  const wb = utils.book_new();
  utils.book_append_sheet(wb, utils.aoa_to_sheet(summaryData), "Summary");
  utils.book_append_sheet(wb, utils.aoa_to_sheet([detailHeaders, ...detailRows]), "Detail Report");
  utils.book_append_sheet(wb, utils.aoa_to_sheet([sheetHeaders, ...sheetRows]), "Comparison Sheet");
  writeFile(wb, "CompareIQ_Results.xlsx");
}

const STEPS = ["Upload", "Configure", "Results"];

export default function CompareIQ() {
  const [step, setStep] = useState(0);
  const [fileA, setFileA] = useState(null);
  const [fileB, setFileB] = useState(null);
  const [dataA, setDataA] = useState([]);
  const [dataB, setDataB] = useState([]);
  const [headersA, setHeadersA] = useState([]);
  const [headersB, setHeadersB] = useState([]);
  const [keyFields, setKeyFields] = useState([]);
  const [currentFieldA, setCurrentFieldA] = useState("");
  const [currentFieldB, setCurrentFieldB] = useState("");
  const [tolerance, setTolerance] = useState("1");
  const [results, setResults] = useState(null);
  const [loading, setLoading] = useState(false);
  const [activeTab, setActiveTab] = useState("dashboard");
  const [filterStatus, setFilterStatus] = useState("All");
  const [searchKey, setSearchKey] = useState("");
  const [error, setError] = useState("");

  const PREFERRED_KEYS = ["person number", "personnumber", "balance name", "balancename", "area1", "area2", "area3"];

  const handleFileUpload = async (file, which) => {
    setError(""); setLoading(true);
    try {
      const data = await parseFile(file);
      if (!data.length) throw new Error("No data found in file");
      const headers = Object.keys(data[0]);
      if (which === "A") { setFileA(file); setDataA(data); setHeadersA(headers); }
      else { setFileB(file); setDataB(data); setHeadersB(headers); }
    } catch (e) { setError(`Error parsing file: ${e.message}`); }
    setLoading(false);
  };

  const proceedToConfigure = () => {
    if (!dataA.length || !dataB.length) { setError("Please upload both files."); return; }
    const norm = s => s.toLowerCase().replace(/[^a-z0-9]/g, "");
    const autoKeys = headersA.filter(h => PREFERRED_KEYS.some(p => norm(h).includes(norm(p))));
    setKeyFields(autoKeys.length ? autoKeys : [headersA[0]]);
    const autoCurrent = headersA.find(h => norm(h).includes("current")) || "";
    setCurrentFieldA(autoCurrent);
    const mappings = suggestMappings(headersA, headersB);
    const mappedCurrent = autoCurrent ? (mappings.find(m => m.from === autoCurrent)?.to || headersB.find(h => norm(h).includes("current")) || "") : "";
    setCurrentFieldB(mappedCurrent);
    setStep(1);
  };

  const runComparison = () => {
    if (!keyFields.length) { setError("Select at least one key field."); return; }
    setLoading(true);
    setTimeout(() => {
      const res = compareDatasets(dataA, dataB, keyFields, currentFieldA, currentFieldB, tolerance);
      setResults(res); setStep(2); setLoading(false);
    }, 400);
  };

  const stats = results ? {
    totalA: dataA.length, totalB: dataB.length,
    matched: results.results.filter(r => r.status === "Matched").length,
    mismatched: results.results.filter(r => r.status === "Mismatched").length,
    onlyA: results.results.filter(r => r.status === "Only in A").length,
    onlyB: results.results.filter(r => r.status === "Only in B").length,
    duplicates: results.duplicatesA,
  } : null;

  const filteredResults = results ? results.results.filter(r =>
    (filterStatus === "All" || r.status === filterStatus) &&
    (!searchKey || r.key.toLowerCase().includes(searchKey.toLowerCase()))
  ) : [];

  const DropZone = ({ label, file, onFile }) => {
    const ref = useRef();
    const [drag, setDrag] = useState(false);
    return (
      <div onClick={() => ref.current.click()}
        onDragOver={e => { e.preventDefault(); setDrag(true); }}
        onDragLeave={() => setDrag(false)}
        onDrop={e => { e.preventDefault(); setDrag(false); if (e.dataTransfer.files[0]) onFile(e.dataTransfer.files[0]); }}
        style={{ border: `2px dashed ${drag ? "#38bdf8" : file ? "#22c55e" : "#1e3a5f"}`, borderRadius: 14, padding: "26px 20px", textAlign: "center", cursor: "pointer", background: drag ? "#0c1e38" : file ? "#0a1f12" : "#060d1a", transition: "all .2s" }}>
        <input ref={ref} type="file" accept=".csv,.xlsx,.xls,.txt,.tsv,.pdf,.tex" style={{ display: "none" }} onChange={e => { if (e.target.files[0]) onFile(e.target.files[0]); }} />
        <div style={{ fontSize: 26, marginBottom: 8 }}>{file ? "✅" : "📂"}</div>
        <div style={{ fontWeight: 700, color: file ? "#22c55e" : "#475569", fontSize: 13 }}>{file ? file.name : label}</div>
        {file
          ? <div style={{ color: "#334155", fontSize: 11, marginTop: 4 }}>{(file.size / 1024).toFixed(1)} KB · click to replace</div>
          : <div style={{ color: "#1e3a5f", fontSize: 11, marginTop: 6 }}>CSV · Excel · TXT · TSV · PDF · TeX · Pipe</div>}
      </div>
    );
  };

  const Checkbox = ({ label, checked, onChange, color = "#38bdf8" }) => (
    <label onClick={onChange} style={{ display: "flex", alignItems: "center", gap: 10, padding: "7px 12px", borderRadius: 8, cursor: "pointer", background: checked ? color + "14" : "transparent", border: `1px solid ${checked ? color + "55" : "#0f2040"}`, transition: "all .15s", userSelect: "none" }}>
      <div style={{ width: 15, height: 15, borderRadius: 4, border: `2px solid ${checked ? color : "#1e3a5f"}`, background: checked ? color : "transparent", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0, transition: "all .15s" }}>
        {checked && <span style={{ color: "#020817", fontSize: 10, fontWeight: 900, lineHeight: 1 }}>✓</span>}
      </div>
      <span style={{ fontSize: 12, color: checked ? "#e2e8f0" : "#475569", fontWeight: checked ? 600 : 400 }}>{label}</span>
    </label>
  );

  const StatusBadge = ({ status }) => {
    const cfg = { Matched: "#22c55e", Mismatched: "#ef4444", "Only in A": "#f59e0b", "Only in B": "#a855f7" };
    const c = cfg[status] || "#64748b";
    return <span style={{ background: c + "18", color: c, border: `1px solid ${c}40`, borderRadius: 6, padding: "2px 10px", fontSize: 11, fontWeight: 700 }}>{status}</span>;
  };

  return (
    <div style={{ minHeight: "100vh", background: "#04080f", color: "#e2e8f0", fontFamily: "'DM Mono','Courier New',monospace" }}>
      <link href="https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=Outfit:wght@700;800&display=swap" rel="stylesheet" />

      {/* Header */}
      <div style={{ background: "#060d1a", borderBottom: "1px solid #0f2040", padding: "13px 28px", display: "flex", alignItems: "center", justifyContent: "space-between", position: "sticky", top: 0, zIndex: 50 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 11 }}>
          <div style={{ width: 32, height: 32, borderRadius: 8, background: "linear-gradient(135deg,#0ea5e9,#6366f1)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16 }}>⚖️</div>
          <div>
            <div style={{ fontFamily: "'Outfit',sans-serif", fontWeight: 800, fontSize: 16, background: "linear-gradient(90deg,#38bdf8,#818cf8)", WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent" }}>CompareIQ</div>
            <div style={{ fontSize: 9, color: "#1e3a5f", letterSpacing: 1.5 }}>INTELLIGENT DATASET COMPARISON</div>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
          {STEPS.map((s, i) => (
            <div key={s} style={{ display: "flex", alignItems: "center", gap: 4 }}>
              <div onClick={() => step > i && setStep(i)} style={{ display: "flex", alignItems: "center", gap: 7, padding: "5px 13px", borderRadius: 20, background: step === i ? "#0c2040" : "transparent", border: `1px solid ${step === i ? "#0ea5e9" : step > i ? "#1e3a5f" : "#0a1628"}`, color: step === i ? "#38bdf8" : step > i ? "#1e4060" : "#1e3a5f", fontSize: 11, fontWeight: 700, cursor: step > i ? "pointer" : "default" }}>
                <div style={{ width: 15, height: 15, borderRadius: "50%", background: step > i ? "#22c55e" : step === i ? "#0ea5e9" : "#0a1628", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 8, color: "#fff", fontWeight: 900 }}>{step > i ? "✓" : i + 1}</div>
                {s}
              </div>
              {i < STEPS.length - 1 && <div style={{ width: 16, height: 1, background: "#0a1628" }} />}
            </div>
          ))}
        </div>
      </div>

      <div style={{ maxWidth: 1080, margin: "0 auto", padding: "28px 20px" }}>
        {error && <div style={{ background: "#140505", border: "1px solid #7f1d1d", color: "#fca5a5", borderRadius: 10, padding: "10px 16px", marginBottom: 16, fontSize: 12, display: "flex", justifyContent: "space-between" }}>⚠️ {error}<button onClick={() => setError("")} style={{ background: "none", border: "none", color: "#fca5a5", cursor: "pointer" }}>✕</button></div>}
        {loading && <div style={{ background: "#050f1f", border: "1px solid #0ea5e9", borderRadius: 10, padding: "10px 16px", marginBottom: 16, fontSize: 12, color: "#38bdf8" }}>⏳ Processing…</div>}

        {/* STEP 0 — UPLOAD */}
        {step === 0 && (
          <div>
            <div style={{ textAlign: "center", marginBottom: 32 }}>
              <h1 style={{ fontFamily: "'Outfit',sans-serif", fontWeight: 800, fontSize: 30, margin: "0 0 8px", background: "linear-gradient(90deg,#38bdf8,#818cf8)", WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent" }}>Upload Your Datasets</h1>
              <p style={{ color: "#334155", margin: 0, fontSize: 13 }}>CSV · Excel · TXT · TSV · PDF · TeX · Pipe-delimited</p>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 18, marginBottom: 28 }}>
              {[{ label: "Drop Dataset A — Source / HR / Expected", which: "A", file: fileA, data: dataA, headers: headersA, color: "#0ea5e9" },
                { label: "Drop Dataset B — Target / Payroll / Actual", which: "B", file: fileB, data: dataB, headers: headersB, color: "#818cf8" }
              ].map(({ label, which, file, data, headers, color }) => (
                <div key={which}>
                  <div style={{ color, fontSize: 11, fontWeight: 700, marginBottom: 8, letterSpacing: 1.5 }}>DATASET {which}</div>
                  <DropZone label={label} file={file} onFile={f => handleFileUpload(f, which)} />
                  {data.length > 0 && (
                    <div style={{ marginTop: 10, background: "#060d1a", borderRadius: 10, padding: 12, border: `1px solid ${color}22` }}>
                      <div style={{ color, fontSize: 11, fontWeight: 700, marginBottom: 7 }}>{data.length} rows · {headers.length} columns</div>
                      <div style={{ overflowX: "auto" }}>
                        <table style={{ borderCollapse: "collapse", fontSize: 11, width: "100%" }}>
                          <thead><tr>{headers.slice(0, 6).map(h => <th key={h} style={{ padding: "3px 9px", background: "#04080f", color: "#334155", textAlign: "left", whiteSpace: "nowrap" }}>{h}</th>)}{headers.length > 6 && <th style={{ color: "#1e3a5f", padding: "3px 7px" }}>+{headers.length - 6}</th>}</tr></thead>
                          <tbody>{data.slice(0, 3).map((r, i) => <tr key={i}>{headers.slice(0, 6).map(h => <td key={h} style={{ padding: "3px 9px", color: "#475569", borderBottom: "1px solid #04080f", whiteSpace: "nowrap" }}>{r[h]}</td>)}</tr>)}</tbody>
                        </table>
                      </div>
                    </div>
                  )}
                </div>
              ))}
            </div>
            <div style={{ textAlign: "center" }}>
              <button onClick={proceedToConfigure} disabled={!dataA.length || !dataB.length}
                style={{ background: dataA.length && dataB.length ? "linear-gradient(135deg,#0ea5e9,#6366f1)" : "#0a1628", color: dataA.length && dataB.length ? "#fff" : "#1e3a5f", border: "none", borderRadius: 10, padding: "12px 36px", fontSize: 14, fontWeight: 700, cursor: dataA.length && dataB.length ? "pointer" : "not-allowed", fontFamily: "inherit" }}>
                Continue to Configure →
              </button>
            </div>
          </div>
        )}

        {/* STEP 1 — CONFIGURE */}
        {step === 1 && (
          <div>
            <div style={{ marginBottom: 24 }}>
              <h1 style={{ fontFamily: "'Outfit',sans-serif", fontWeight: 800, fontSize: 24, margin: "0 0 5px", background: "linear-gradient(90deg,#38bdf8,#818cf8)", WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent" }}>Configure Comparison</h1>
              <p style={{ color: "#334155", margin: 0, fontSize: 12 }}>Select key fields to match rows, then pick the value field to compare</p>
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "1.1fr 1fr", gap: 18, marginBottom: 18 }}>

              {/* Key Fields */}
              <div style={{ background: "#060d1a", borderRadius: 14, padding: 20, border: "1px solid #0f2040" }}>
                <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: 14 }}>
                  <div>
                    <div style={{ fontWeight: 700, color: "#f59e0b", fontSize: 12, letterSpacing: 1 }}>🔑 KEY FIELDS</div>
                    <div style={{ color: "#1e3a5f", fontSize: 11, marginTop: 3 }}>Used to match rows between datasets (from Dataset A)</div>
                  </div>
                  <div style={{ display: "flex", gap: 5 }}>
                    <button onClick={() => setKeyFields([...headersA])} style={{ background: "#1a110a", border: "1px solid #f59e0b33", color: "#f59e0b", borderRadius: 6, padding: "3px 9px", fontSize: 10, cursor: "pointer", fontFamily: "inherit", fontWeight: 700 }}>All</button>
                    <button onClick={() => setKeyFields([])} style={{ background: "#04080f", border: "1px solid #0f2040", color: "#334155", borderRadius: 6, padding: "3px 9px", fontSize: 10, cursor: "pointer", fontFamily: "inherit" }}>None</button>
                  </div>
                </div>
                <div style={{ display: "flex", flexDirection: "column", gap: 5, maxHeight: 300, overflowY: "auto" }}>
                  {headersA.map(h => (
                    <Checkbox key={h} label={h} color="#f59e0b"
                      checked={keyFields.includes(h)}
                      onChange={() => setKeyFields(prev => prev.includes(h) ? prev.filter(k => k !== h) : [...prev, h])} />
                  ))}
                </div>
                <div style={{ marginTop: 10, padding: "7px 11px", background: "#04080f", borderRadius: 7, fontSize: 11 }}>
                  {keyFields.length > 0
                    ? <span><span style={{ color: "#f59e0b", fontWeight: 700 }}>{keyFields.length} selected:</span> <span style={{ color: "#475569" }}>{keyFields.join(" · ")}</span></span>
                    : <span style={{ color: "#1e3a5f" }}>No key fields selected</span>}
                </div>
              </div>

              {/* Right: Comparison field + tolerance */}
              <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
                <div style={{ background: "#060d1a", borderRadius: 14, padding: 20, border: "1px solid #0f2040", flex: 1 }}>
                  <div style={{ fontWeight: 700, color: "#38bdf8", fontSize: 12, letterSpacing: 1, marginBottom: 4 }}>📊 COMPARISON FIELD</div>
                  <div style={{ color: "#1e3a5f", fontSize: 11, marginBottom: 14 }}>Select "Current" field from each dataset to compare</div>

                  <div style={{ marginBottom: 12 }}>
                    <div style={{ fontSize: 10, color: "#0ea5e9", marginBottom: 7, fontWeight: 700, letterSpacing: 1 }}>DATASET A — {fileA?.name}</div>
                    <div style={{ display: "flex", flexDirection: "column", gap: 4, maxHeight: 150, overflowY: "auto" }}>
                      {headersA.map(h => (
                        <Checkbox key={h} label={h} color="#38bdf8"
                          checked={currentFieldA === h}
                          onChange={() => setCurrentFieldA(prev => prev === h ? "" : h)} />
                      ))}
                    </div>
                  </div>

                  <div style={{ height: 1, background: "#0f2040", margin: "10px 0" }} />

                  <div>
                    <div style={{ fontSize: 10, color: "#818cf8", marginBottom: 7, fontWeight: 700, letterSpacing: 1 }}>DATASET B — {fileB?.name}</div>
                    <div style={{ display: "flex", flexDirection: "column", gap: 4, maxHeight: 150, overflowY: "auto" }}>
                      {headersB.map(h => (
                        <Checkbox key={h} label={h} color="#818cf8"
                          checked={currentFieldB === h}
                          onChange={() => setCurrentFieldB(prev => prev === h ? "" : h)} />
                      ))}
                    </div>
                  </div>

                  {currentFieldA && currentFieldB && (
                    <div style={{ marginTop: 10, padding: "7px 11px", background: "#04080f", borderRadius: 7, fontSize: 11, display: "flex", alignItems: "center", gap: 8 }}>
                      <span style={{ color: "#38bdf8", fontWeight: 700 }}>{currentFieldA}</span>
                      <span style={{ color: "#1e3a5f" }}>↔</span>
                      <span style={{ color: "#818cf8", fontWeight: 700 }}>{currentFieldB}</span>
                    </div>
                  )}
                </div>

                {/* Tolerance */}
                <div style={{ background: "#060d1a", borderRadius: 14, padding: 16, border: "1px solid #0f2040", display: "flex", alignItems: "center", gap: 14 }}>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontWeight: 700, color: "#22c55e", fontSize: 12, letterSpacing: 1 }}>⚙️ TOLERANCE</div>
                    <div style={{ color: "#1e3a5f", fontSize: 11, marginTop: 3 }}>Acceptable % variance for numeric values</div>
                  </div>
                  <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                    <input type="number" min="0" max="100" step="0.1" value={tolerance} onChange={e => setTolerance(e.target.value)}
                      style={{ width: 60, background: "#04080f", border: "1px solid #22c55e44", borderRadius: 8, padding: "7px 8px", color: "#22c55e", fontSize: 17, fontWeight: 700, textAlign: "center", fontFamily: "inherit" }} />
                    <span style={{ color: "#22c55e", fontWeight: 700 }}>%</span>
                  </div>
                </div>
              </div>
            </div>

            <div style={{ display: "flex", gap: 10, justifyContent: "center" }}>
              <button onClick={() => setStep(0)} style={{ background: "transparent", border: "1px solid #0f2040", color: "#334155", borderRadius: 10, padding: "11px 24px", fontSize: 13, cursor: "pointer", fontFamily: "inherit" }}>← Back</button>
              <button onClick={runComparison} style={{ background: "linear-gradient(135deg,#0ea5e9,#6366f1)", color: "#fff", border: "none", borderRadius: 10, padding: "11px 34px", fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>🚀 Run Comparison</button>
            </div>
          </div>
        )}

        {/* STEP 2 — RESULTS */}
        {step === 2 && stats && (
          <div>
            {/* Cards */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(7,1fr)", gap: 9, marginBottom: 18 }}>
              {[
                { label: "Total A", val: stats.totalA, color: "#38bdf8" },
                { label: "Total B", val: stats.totalB, color: "#818cf8" },
                { label: "Matched", val: stats.matched, color: "#22c55e" },
                { label: "Mismatched", val: stats.mismatched, color: "#ef4444" },
                { label: "Only in A", val: stats.onlyA, color: "#f59e0b" },
                { label: "Only in B", val: stats.onlyB, color: "#a855f7" },
                { label: "Duplicates", val: stats.duplicates, color: "#475569" },
              ].map(s => (
                <div key={s.label} style={{ background: "#060d1a", border: `1px solid ${s.color}25`, borderTop: `3px solid ${s.color}`, borderRadius: 11, padding: "12px 8px", textAlign: "center" }}>
                  <div style={{ fontSize: 20, fontWeight: 800, color: s.color, fontFamily: "'Outfit'" }}>{s.val}</div>
                  <div style={{ fontSize: 9, color: "#1e3a5f", marginTop: 3, letterSpacing: .5 }}>{s.label.toUpperCase()}</div>
                </div>
              ))}
            </div>

            {/* Match bar */}
            <div style={{ background: "#060d1a", borderRadius: 11, padding: "12px 16px", marginBottom: 18, border: "1px solid #0f2040" }}>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11, marginBottom: 7 }}>
                <span style={{ color: "#334155" }}>Match Rate</span>
                <span style={{ color: "#22c55e", fontWeight: 700 }}>{((stats.matched / (stats.matched + stats.mismatched || 1)) * 100).toFixed(1)}%</span>
              </div>
              <div style={{ height: 7, background: "#04080f", borderRadius: 99, display: "flex", overflow: "hidden" }}>
                {[{ w: stats.matched, c: "#22c55e" }, { w: stats.mismatched, c: "#ef4444" }, { w: stats.onlyA, c: "#f59e0b" }, { w: stats.onlyB, c: "#a855f7" }].map((s, i) => {
                  const t = stats.matched + stats.mismatched + stats.onlyA + stats.onlyB || 1;
                  return <div key={i} style={{ width: `${(s.w / t) * 100}%`, background: s.c }} />;
                })}
              </div>
            </div>

            {/* Tab bar */}
            <div style={{ display: "flex", gap: 5, marginBottom: 16, alignItems: "center" }}>
              {[["dashboard", "📊 Dashboard"], ["detail", "🔍 Detail Report"], ["sheet", "📋 Comparison Sheet"]].map(([t, label]) => (
                <button key={t} onClick={() => setActiveTab(t)} style={{ background: activeTab === t ? "#0c2040" : "transparent", border: `1px solid ${activeTab === t ? "#0ea5e9" : "#0f2040"}`, color: activeTab === t ? "#38bdf8" : "#1e3a5f", borderRadius: 7, padding: "7px 16px", fontSize: 12, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>{label}</button>
              ))}
              <div style={{ flex: 1 }} />
              <button onClick={() => exportToExcel(results.results, keyFields, currentFieldA, currentFieldB, fileA?.name, fileB?.name, dataA.length, dataB.length)}
                style={{ background: "#0a1f0a", border: "1px solid #22c55e44", color: "#22c55e", borderRadius: 7, padding: "7px 16px", fontSize: 12, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>⬇ Export Excel</button>
              <button onClick={() => setStep(1)} style={{ background: "transparent", border: "1px solid #0f2040", color: "#1e3a5f", borderRadius: 7, padding: "7px 12px", fontSize: 11, cursor: "pointer", fontFamily: "inherit" }}>← Reconfigure</button>
            </div>

            {/* Dashboard */}
            {activeTab === "dashboard" && (
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
                <div style={{ background: "#060d1a", borderRadius: 13, padding: 22, border: "1px solid #0f2040" }}>
                  <div style={{ fontWeight: 700, color: "#94a3b8", fontSize: 12, marginBottom: 16, letterSpacing: .5 }}>CLASSIFICATION BREAKDOWN</div>
                  <div style={{ display: "flex", alignItems: "center", gap: 20 }}>
                    <svg width="120" height="120" viewBox="0 0 120 120">
                      {(() => {
                        const segs = [{ val: stats.matched, c: "#22c55e" }, { val: stats.mismatched, c: "#ef4444" }, { val: stats.onlyA, c: "#f59e0b" }, { val: stats.onlyB, c: "#a855f7" }];
                        const total = segs.reduce((s, x) => s + x.val, 0) || 1;
                        const r = 46, cx = 60, cy = 60; let angle = -Math.PI / 2;
                        return segs.map((s, i) => {
                          if (!s.val) return null;
                          const sweep = (s.val / total) * 2 * Math.PI;
                          const x1 = cx + r * Math.cos(angle), y1 = cy + r * Math.sin(angle);
                          angle += sweep;
                          return <path key={i} d={`M${cx},${cy} L${x1},${y1} A${r},${r} 0 ${sweep > Math.PI ? 1 : 0} 1 ${cx + r * Math.cos(angle)},${cy + r * Math.sin(angle)} Z`} fill={s.c} opacity={0.88} />;
                        });
                      })()}
                      <circle cx="60" cy="60" r="28" fill="#060d1a" />
                      <text x="60" y="64" textAnchor="middle" fill="#e2e8f0" fontSize="12" fontWeight="bold" fontFamily="Outfit">{((stats.matched / (stats.matched + stats.mismatched || 1)) * 100).toFixed(0)}%</text>
                    </svg>
                    <div style={{ flex: 1 }}>
                      {[["Matched", stats.matched, "#22c55e"], ["Mismatched", stats.mismatched, "#ef4444"], ["Only A", stats.onlyA, "#f59e0b"], ["Only B", stats.onlyB, "#a855f7"]].map(([l, v, c]) => (
                        <div key={l} style={{ display: "flex", alignItems: "center", gap: 9, marginBottom: 9 }}>
                          <div style={{ width: 7, height: 7, borderRadius: 2, background: c, flexShrink: 0 }} />
                          <span style={{ flex: 1, fontSize: 12, color: "#334155" }}>{l}</span>
                          <span style={{ fontWeight: 700, color: c }}>{v}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>

                <div style={{ background: "#060d1a", borderRadius: 13, padding: 22, border: "1px solid #0f2040" }}>
                  <div style={{ fontWeight: 700, color: "#94a3b8", fontSize: 12, marginBottom: 16, letterSpacing: .5 }}>COMPARISON FIELD ANALYSIS</div>
                  {currentFieldA && currentFieldB ? (() => {
                    const total = results.results.filter(r => r.rowA && r.rowB).length || 1;
                    const mismatches = results.results.filter(r => r.details.some(d => d.status === "Mismatched")).length;
                    const matchedCt = results.results.filter(r => r.details.length && r.details[0].status === "Matched").length;
                    const pct = (mismatches / total) * 100;
                    return (
                      <div>
                        <div style={{ padding: "10px 14px", background: "#04080f", borderRadius: 9, marginBottom: 14, display: "flex", alignItems: "center", gap: 10, fontSize: 12 }}>
                          <span style={{ color: "#38bdf8", fontWeight: 700 }}>{currentFieldA}</span>
                          <span style={{ color: "#1e3a5f" }}>↔</span>
                          <span style={{ color: "#818cf8", fontWeight: 700 }}>{currentFieldB}</span>
                        </div>
                        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 14 }}>
                          {[["Matched", matchedCt, "#22c55e", "#0a1f12"], ["Mismatched", mismatches, "#ef4444", "#140505"]].map(([l, v, c, bg]) => (
                            <div key={l} style={{ background: bg, border: `1px solid ${c}25`, borderRadius: 9, padding: "10px 14px", textAlign: "center" }}>
                              <div style={{ fontSize: 20, fontWeight: 800, color: c, fontFamily: "'Outfit'" }}>{v}</div>
                              <div style={{ fontSize: 10, color: "#334155", marginTop: 2 }}>{l}</div>
                            </div>
                          ))}
                        </div>
                        <div style={{ fontSize: 11, color: "#334155", marginBottom: 6 }}>Mismatch Rate: <span style={{ color: pct > 10 ? "#ef4444" : "#22c55e", fontWeight: 700 }}>{pct.toFixed(1)}%</span></div>
                        <div style={{ height: 5, background: "#04080f", borderRadius: 99 }}>
                          <div style={{ width: `${pct}%`, height: "100%", background: pct > 10 ? "#ef4444" : "#22c55e", borderRadius: 99 }} />
                        </div>
                      </div>
                    );
                  })() : <div style={{ color: "#1e3a5f", fontSize: 12 }}>No comparison field configured</div>}
                </div>
              </div>
            )}

            {/* Detail report */}
            {activeTab === "detail" && (
              <div style={{ background: "#060d1a", borderRadius: 13, border: "1px solid #0f2040", overflow: "hidden" }}>
                <div style={{ padding: "12px 16px", borderBottom: "1px solid #0f2040", display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                  <input placeholder="Search by key values..." value={searchKey} onChange={e => setSearchKey(e.target.value)}
                    style={{ background: "#04080f", border: "1px solid #0f2040", borderRadius: 7, padding: "6px 12px", color: "#e2e8f0", fontSize: 11, fontFamily: "inherit", flex: 1, minWidth: 140 }} />
                  {["All", "Matched", "Mismatched", "Only in A", "Only in B"].map(s => (
                    <button key={s} onClick={() => setFilterStatus(s)} style={{ background: filterStatus === s ? "#0c2040" : "transparent", border: `1px solid ${filterStatus === s ? "#0ea5e9" : "#0f2040"}`, color: filterStatus === s ? "#38bdf8" : "#1e3a5f", borderRadius: 6, padding: "5px 11px", cursor: "pointer", fontSize: 10, fontWeight: 700, fontFamily: "inherit" }}>{s}</button>
                  ))}
                  <span style={{ fontSize: 10, color: "#1e3a5f" }}>{filteredResults.length} rows</span>
                </div>
                <div style={{ overflowX: "auto", maxHeight: 440, overflowY: "auto" }}>
                  <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 11 }}>
                    <thead style={{ position: "sticky", top: 0, background: "#060d1a", zIndex: 1 }}>
                      <tr>
                        {keyFields.map(k => <th key={k} style={{ padding: "9px 13px", textAlign: "left", color: "#f59e0b", borderBottom: "1px solid #0f2040", whiteSpace: "nowrap", fontSize: 10, letterSpacing: .5 }}>{k.toUpperCase()}</th>)}
                        {currentFieldA && <th style={{ padding: "9px 13px", color: "#38bdf8", borderBottom: "1px solid #0f2040", borderLeft: "1px solid #0f2040" }}>{currentFieldA} (A)</th>}
                        {currentFieldB && <th style={{ padding: "9px 13px", color: "#818cf8", borderBottom: "1px solid #0f2040" }}>{currentFieldB} (B)</th>}
                        <th style={{ padding: "9px 13px", color: "#f59e0b", borderBottom: "1px solid #0f2040" }}>Diff</th>
                        <th style={{ padding: "9px 13px", color: "#475569", borderBottom: "1px solid #0f2040" }}>Status</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredResults.slice(0, 200).map((r, i) => {
                        const d = r.details[0];
                        return (
                          <tr key={i} style={{ background: i % 2 ? "#04080f" : "#060d1a" }}>
                            {keyFields.map(k => <td key={k} style={{ padding: "8px 13px", color: "#64748b", borderBottom: "1px solid #0f204018", whiteSpace: "nowrap" }}>{r.keyVals[k] ?? "—"}</td>)}
                            {currentFieldA && <td style={{ padding: "8px 13px", borderBottom: "1px solid #0f204018", color: "#38bdf8", borderLeft: "1px solid #0f2040" }}>{d ? d.valA : (r.rowA ? r.rowA[currentFieldA] ?? "—" : "—")}</td>}
                            {currentFieldB && <td style={{ padding: "8px 13px", borderBottom: "1px solid #0f204018", color: "#818cf8" }}>{d ? d.valB : (r.rowB ? r.rowB[currentFieldB] ?? "—" : "—")}</td>}
                            <td style={{ padding: "8px 13px", borderBottom: "1px solid #0f204018", color: d?.diff && d.diff !== "0.00" ? "#f59e0b" : "#1e3a5f" }}>{d?.diff || "—"}</td>
                            <td style={{ padding: "8px 13px", borderBottom: "1px solid #0f204018" }}><StatusBadge status={r.status} /></td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                  {filteredResults.length === 0 && <div style={{ padding: 36, textAlign: "center", color: "#1e3a5f" }}>No records match the filter.</div>}
                  {filteredResults.length > 200 && <div style={{ padding: 10, textAlign: "center", color: "#1e3a5f", fontSize: 10 }}>Showing 200 of {filteredResults.length} — Export Excel for all.</div>}
                </div>
              </div>
            )}

            {/* Comparison Sheet */}
            {activeTab === "sheet" && (
              <div style={{ background: "#060d1a", borderRadius: 13, border: "1px solid #0f2040", overflow: "hidden" }}>
                <div style={{ padding: "11px 16px", borderBottom: "1px solid #0f2040", fontSize: 11, color: "#1e3a5f", display: "flex", justifyContent: "space-between" }}>
                  <span>Side-by-side view of all records</span>
                  <span>{results.results.length} records</span>
                </div>
                <div style={{ overflowX: "auto", maxHeight: 480, overflowY: "auto" }}>
                  <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 11 }}>
                    <thead style={{ position: "sticky", top: 0, zIndex: 1 }}>
                      <tr>
                        <th colSpan={headersA.length} style={{ padding: "7px 12px", background: "#041428", color: "#1e4060", borderBottom: "1px solid #0f2040", textAlign: "center", fontSize: 10 }}>📋 {fileA?.name}</th>
                        <th colSpan={headersB.length} style={{ padding: "7px 12px", background: "#0d0a20", color: "#2a1a4a", borderBottom: "1px solid #0f2040", borderLeft: "2px solid #6366f120", textAlign: "center", fontSize: 10 }}>📋 {fileB?.name}</th>
                        <th style={{ padding: "7px 12px", background: "#060d1a", borderBottom: "1px solid #0f2040" }}></th>
                      </tr>
                      <tr>
                        {headersA.map(h => <th key={h} style={{ padding: "6px 11px", background: "#041428", color: "#1e4060", textAlign: "left", borderBottom: "1px solid #0f2040", whiteSpace: "nowrap" }}>{h}</th>)}
                        {headersB.map(h => <th key={h} style={{ padding: "6px 11px", background: "#0d0a20", color: "#2a1a4a", textAlign: "left", borderBottom: "1px solid #0f2040", borderLeft: h === headersB[0] ? "2px solid #6366f115" : "none", whiteSpace: "nowrap" }}>{h}</th>)}
                        <th style={{ padding: "6px 11px", background: "#060d1a", color: "#1e3a5f", borderBottom: "1px solid #0f2040", borderLeft: "1px solid #0f2040" }}>Status</th>
                      </tr>
                    </thead>
                    <tbody>
                      {results.results.slice(0, 200).map((r, i) => (
                        <tr key={i} style={{ background: r.status === "Matched" ? (i % 2 ? "#04080f" : "#060d1a") : r.status === "Mismatched" ? "#110404" : r.status === "Only in A" ? "#110d03" : "#0b0414" }}>
                          {headersA.map(h => <td key={h} style={{ padding: "6px 11px", color: "#1e4060", borderBottom: "1px solid #0f204012", whiteSpace: "nowrap" }}>{r.rowA ? r.rowA[h] ?? "" : <span style={{ color: "#0a1628" }}>—</span>}</td>)}
                          {headersB.map(h => <td key={h} style={{ padding: "6px 11px", color: "#2a1a4a", borderBottom: "1px solid #0f204012", borderLeft: h === headersB[0] ? "2px solid #6366f112" : "none", whiteSpace: "nowrap" }}>{r.rowB ? r.rowB[h] ?? "" : <span style={{ color: "#0a1628" }}>—</span>}</td>)}
                          <td style={{ padding: "6px 11px", borderBottom: "1px solid #0f204012", borderLeft: "1px solid #0f2040" }}><StatusBadge status={r.status} /></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {results.results.length > 200 && <div style={{ padding: 10, textAlign: "center", color: "#1e3a5f", fontSize: 10 }}>Showing 200 of {results.results.length} — Export Excel for full data.</div>}
                </div>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
