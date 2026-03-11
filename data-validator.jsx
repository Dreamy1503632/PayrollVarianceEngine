import { useState, useCallback, useRef } from "react";

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
  // Try auto-detect delimiter
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
          // For PDF, extract text via pdf.js
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
        if (["tsv"].includes(ext)) resolve(parseTSV(text));
        else if (ext === "txt" || ext === "tex") resolve(parsePlainText(text));
        else resolve(parseCSV(text));
      };
      reader.readAsText(file);
    }
  });
}

// ─── AI-style field mapping suggestion ──────────────────────────────────────
function suggestMappings(headersA, headersB) {
  const normalize = s => s.toLowerCase().replace(/[^a-z0-9]/g, "");
  const aliases = {
    id: ["id", "empid", "employeeid", "emp_id", "employee_id", "personnumber", "person_number", "empno", "staffid"],
    name: ["name", "fullname", "employeename", "empname", "person_name"],
    salary: ["salary", "basepay", "base_pay", "basesalary", "pay", "wage", "compensation"],
    overtime: ["overtime", "otpay", "ot_pay", "ovtpay", "extratime"],
    date: ["date", "paydate", "pay_date", "payperiod", "period"],
    department: ["department", "dept", "division", "team"],
    email: ["email", "emailaddress", "mail"],
  };

  const suggestions = [];
  const usedB = new Set();

  for (const ha of headersA) {
    const nha = normalize(ha);
    let best = null, bestScore = 0;

    for (const hb of headersB) {
      if (usedB.has(hb)) continue;
      const nhb = normalize(hb);
      let score = 0;

      if (nha === nhb) { score = 100; }
      else if (nha.includes(nhb) || nhb.includes(nha)) { score = 80; }
      else {
        for (const group of Object.values(aliases)) {
          if (group.some(a => nha.includes(a)) && group.some(a => nhb.includes(a))) {
            score = 70; break;
          }
        }
      }

      if (score > bestScore) { bestScore = score; best = hb; }
    }

    if (best && bestScore >= 70) {
      suggestions.push({ from: ha, to: best, confidence: bestScore });
      usedB.add(best);
    }
  }
  return suggestions;
}

// ─── Compare engine ──────────────────────────────────────────────────────────
function compareDatasets(dataA, dataB, keyFields, compareFields, tolerance) {
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

    if (matchesB.length === 0) {
      results.push({ key, rowA, rowB: null, status: "Only in A", details: [] });
    } else {
      const rowB = matchesB[0];
      matchedKeysB.add(key);

      const details = compareFields.map(({ fieldA, fieldB }) => {
        const valA = rowA[fieldA] ?? "";
        const valB = rowB[fieldB] ?? "";
        const numA = parseFloat(valA), numB = parseFloat(valB);
        const isNum = !isNaN(numA) && !isNaN(numB);
        let diff = "", status = "Matched";

        if (isNum) {
          const diffVal = Math.abs(numA - numB);
          const pct = numA !== 0 ? diffVal / Math.abs(numA) : (numB !== 0 ? 1 : 0);
          diff = (numB - numA).toFixed(2);
          if (pct > tolPct) status = "Mismatched";
        } else {
          if (valA.toLowerCase().trim() !== valB.toLowerCase().trim()) {
            diff = `"${valA}" vs "${valB}"`; status = "Mismatched";
          }
        }

        return { fieldA, fieldB, valA, valB, diff, status };
      });

      const overallStatus = details.some(d => d.status === "Mismatched") ? "Mismatched" : "Matched";
      results.push({ key, rowA, rowB, status: overallStatus, details });
    }
  }

  // Only in B
  for (const row of dataB) {
    const key = keyFields.map(k => (row[k] || "").toString().toLowerCase().trim()).join("||");
    if (!matchedKeysB.has(key)) {
      results.push({ key, rowA: null, rowB: row, status: "Only in B", details: [] });
      matchedKeysB.add(key);
    }
  }

  // Duplicates in A
  const keyCounts = {};
  for (const row of dataA) {
    const key = keyFields.map(k => (row[k] || "").toString().toLowerCase().trim()).join("||");
    keyCounts[key] = (keyCounts[key] || 0) + 1;
  }

  return { results, duplicatesA: Object.values(keyCounts).filter(c => c > 1).reduce((s, c) => s + c - 1, 0) };
}

// ─── Main App ────────────────────────────────────────────────────────────────
const STEPS = ["Upload", "Configure", "Compare", "Results"];

export default function DataValidator() {
  const [step, setStep] = useState(0);
  const [fileA, setFileA] = useState(null);
  const [fileB, setFileB] = useState(null);
  const [dataA, setDataA] = useState([]);
  const [dataB, setDataB] = useState([]);
  const [headersA, setHeadersA] = useState([]);
  const [headersB, setHeadersB] = useState([]);
  const [mappings, setMappings] = useState([]);
  const [keyFields, setKeyFields] = useState([]);
  const [compareFields, setCompareFields] = useState([]);
  const [tolerance, setTolerance] = useState("1");
  const [results, setResults] = useState(null);
  const [loading, setLoading] = useState(false);
  const [activeTab, setActiveTab] = useState("dashboard");
  const [filterStatus, setFilterStatus] = useState("All");
  const [searchKey, setSearchKey] = useState("");
  const [error, setError] = useState("");

  const handleFileUpload = async (file, which) => {
    setError("");
    setLoading(true);
    try {
      const data = await parseFile(file);
      if (!data.length) throw new Error("No data found in file");
      const headers = Object.keys(data[0]);
      if (which === "A") { setFileA(file); setDataA(data); setHeadersA(headers); }
      else { setFileB(file); setDataB(data); setHeadersB(headers); }
    } catch (e) {
      setError(`Error parsing file: ${e.message}`);
    }
    setLoading(false);
  };

  const proceedToConfigure = () => {
    if (!dataA.length || !dataB.length) { setError("Please upload both files."); return; }
    const suggested = suggestMappings(headersA, headersB);
    setMappings(suggested.length ? suggested : [{ from: headersA[0], to: headersB[0], confidence: 50 }]);
    setKeyFields([headersA[0]]);
    setCompareFields(suggested.slice(1).map(m => ({ fieldA: m.from, fieldB: m.to })));
    setStep(1);
  };

  const runComparison = () => {
    setLoading(true);
    setTimeout(() => {
      const cf = compareFields.filter(f => f.fieldA && f.fieldB);
      const kf = keyFields.filter(Boolean);
      if (!kf.length) { setError("Select at least one key field."); setLoading(false); return; }
      const res = compareDatasets(dataA, dataB, kf, cf, tolerance);
      setResults(res);
      setStep(3);
      setLoading(false);
    }, 500);
  };

  // Stats
  const stats = results ? {
    totalA: dataA.length,
    totalB: dataB.length,
    matched: results.results.filter(r => r.status === "Matched").length,
    mismatched: results.results.filter(r => r.status === "Mismatched").length,
    onlyA: results.results.filter(r => r.status === "Only in A").length,
    onlyB: results.results.filter(r => r.status === "Only in B").length,
    duplicates: results.duplicatesA,
  } : null;

  const filteredResults = results ? results.results.filter(r => {
    const matchStatus = filterStatus === "All" || r.status === filterStatus;
    const key = r.key.toLowerCase();
    const matchSearch = !searchKey || key.includes(searchKey.toLowerCase());
    return matchStatus && matchSearch;
  }) : [];

  const exportCSV = () => {
    if (!results) return;
    const rows = [["Key", "Status", ...compareFields.flatMap(f => [`${f.fieldA} (A)`, `${f.fieldB} (B)`, "Diff", "Status"])]];
    for (const r of results.results) {
      const detailCols = r.details.flatMap(d => [d.valA, d.valB, d.diff, d.status]);
      rows.push([r.key, r.status, ...detailCols]);
    }
    const csv = rows.map(r => r.map(v => `"${v}"`).join(",")).join("\n");
    const blob = new Blob([csv], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a"); a.href = url; a.download = "comparison_results.csv"; a.click();
  };

  // ─── UI ────────────────────────────────────────────────────────────────────
  const DropZone = ({ label, file, onFile }) => {
    const ref = useRef();
    const [drag, setDrag] = useState(false);
    const accept = ".csv,.xlsx,.xls,.txt,.tsv,.pdf,.tex,.pipe";
    return (
      <div
        onClick={() => ref.current.click()}
        onDragOver={e => { e.preventDefault(); setDrag(true); }}
        onDragLeave={() => setDrag(false)}
        onDrop={e => { e.preventDefault(); setDrag(false); const f = e.dataTransfer.files[0]; if (f) onFile(f); }}
        style={{
          border: `2px dashed ${drag ? "#38bdf8" : file ? "#22c55e" : "#334155"}`,
          borderRadius: 12, padding: "28px 20px", textAlign: "center",
          cursor: "pointer", background: drag ? "#0f2540" : file ? "#0d2a1a" : "#0f172a",
          transition: "all .2s"
        }}
      >
        <input ref={ref} type="file" accept={accept} style={{ display: "none" }} onChange={e => { if (e.target.files[0]) onFile(e.target.files[0]); }} />
        <div style={{ fontSize: 32, marginBottom: 8 }}>{file ? "✅" : "📂"}</div>
        <div style={{ fontWeight: 700, color: file ? "#22c55e" : "#94a3b8", fontSize: 14 }}>{file ? file.name : label}</div>
        {file && <div style={{ color: "#64748b", fontSize: 12, marginTop: 4 }}>{(file.size / 1024).toFixed(1)} KB · click to change</div>}
        {!file && <div style={{ color: "#475569", fontSize: 12, marginTop: 6 }}>CSV · Excel · TXT · PDF · TSV · TeX · Pipe-delimited</div>}
      </div>
    );
  };

  const StatusBadge = ({ status }) => {
    const colors = { Matched: "#22c55e", Mismatched: "#ef4444", "Only in A": "#f59e0b", "Only in B": "#8b5cf6" };
    return <span style={{ background: colors[status] + "22", color: colors[status], border: `1px solid ${colors[status]}44`, borderRadius: 6, padding: "2px 10px", fontSize: 12, fontWeight: 700 }}>{status}</span>;
  };

  return (
    <div style={{ minHeight: "100vh", background: "#020817", color: "#e2e8f0", fontFamily: "'IBM Plex Mono', monospace" }}>
      <link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;700&family=Syne:wght@700;800&display=swap" rel="stylesheet" />

      {/* Header */}
      <div style={{ background: "#0a1628", borderBottom: "1px solid #1e293b", padding: "16px 32px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{ width: 36, height: 36, borderRadius: 8, background: "linear-gradient(135deg,#0ea5e9,#6366f1)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 20 }}>⚖️</div>
          <div>
            <div style={{ fontFamily: "'Syne', sans-serif", fontWeight: 800, fontSize: 18, color: "#f1f5f9", letterSpacing: 1 }}>DataSync Validator</div>
            <div style={{ fontSize: 11, color: "#475569" }}>AI-Powered Dataset Comparison Engine</div>
          </div>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          {STEPS.map((s, i) => (
            <div key={s} style={{ display: "flex", alignItems: "center", gap: 6, padding: "6px 14px", borderRadius: 20, background: step === i ? "#0ea5e9" : step > i ? "#1e3a5f" : "#0f172a", color: step === i ? "#fff" : step > i ? "#38bdf8" : "#475569", fontSize: 12, fontWeight: 700, border: step > i ? "1px solid #1e4d8c" : "1px solid transparent", cursor: step > i ? "pointer" : "default", transition: "all .2s" }} onClick={() => step > i && setStep(i)}>
              <span style={{ width: 18, height: 18, borderRadius: "50%", background: step > i ? "#38bdf8" : step === i ? "#fff" : "#1e293b", color: step > i ? "#020817" : step === i ? "#0ea5e9" : "#475569", display: "inline-flex", alignItems: "center", justifyContent: "center", fontSize: 10, fontWeight: 800 }}>{step > i ? "✓" : i + 1}</span>
              {s}
            </div>
          ))}
        </div>
      </div>

      <div style={{ maxWidth: 1200, margin: "0 auto", padding: 32 }}>
        {error && <div style={{ background: "#450a0a", border: "1px solid #ef4444", color: "#fca5a5", borderRadius: 10, padding: "12px 20px", marginBottom: 20, fontSize: 13 }}>⚠️ {error}</div>}
        {loading && <div style={{ background: "#0c1a2e", border: "1px solid #0ea5e9", borderRadius: 10, padding: "12px 20px", marginBottom: 20, fontSize: 13, color: "#38bdf8" }}>⏳ Processing file...</div>}

        {/* STEP 0: Upload */}
        {step === 0 && (
          <div>
            <div style={{ textAlign: "center", marginBottom: 40 }}>
              <h1 style={{ fontFamily: "'Syne'", fontSize: 36, fontWeight: 800, margin: 0, background: "linear-gradient(90deg,#38bdf8,#818cf8)", WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent" }}>Upload Your Datasets</h1>
              <p style={{ color: "#64748b", marginTop: 10 }}>Supports CSV, Excel, TXT, TSV, PDF, TeX, Pipe-delimited, and more</p>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 24 }}>
              <div>
                <div style={{ color: "#94a3b8", fontWeight: 700, marginBottom: 12, fontSize: 13 }}>📋 DATASET A (Source / HR / Expected)</div>
                <DropZone label="Drop Dataset A here or click to browse" file={fileA} onFile={f => handleFileUpload(f, "A")} />
                {dataA.length > 0 && <div style={{ marginTop: 12, background: "#0f172a", borderRadius: 10, padding: 16, border: "1px solid #1e293b" }}>
                  <div style={{ color: "#38bdf8", fontSize: 12, fontWeight: 700, marginBottom: 8 }}>PREVIEW — {dataA.length} rows · {headersA.length} columns</div>
                  <div style={{ overflowX: "auto" }}>
                    <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 11 }}>
                      <thead><tr>{headersA.slice(0, 5).map(h => <th key={h} style={{ padding: "4px 10px", background: "#1e293b", color: "#94a3b8", textAlign: "left", borderRadius: 4 }}>{h}</th>)}{headersA.length > 5 && <th style={{ color: "#475569" }}>+{headersA.length - 5} more</th>}</tr></thead>
                      <tbody>{dataA.slice(0, 3).map((r, i) => <tr key={i}>{headersA.slice(0, 5).map(h => <td key={h} style={{ padding: "4px 10px", borderBottom: "1px solid #1e293b", color: "#cbd5e1" }}>{r[h]}</td>)}</tr>)}</tbody>
                    </table>
                  </div>
                </div>}
              </div>
              <div>
                <div style={{ color: "#94a3b8", fontWeight: 700, marginBottom: 12, fontSize: 13 }}>📋 DATASET B (Target / Payroll / Actual)</div>
                <DropZone label="Drop Dataset B here or click to browse" file={fileB} onFile={f => handleFileUpload(f, "B")} />
                {dataB.length > 0 && <div style={{ marginTop: 12, background: "#0f172a", borderRadius: 10, padding: 16, border: "1px solid #1e293b" }}>
                  <div style={{ color: "#818cf8", fontSize: 12, fontWeight: 700, marginBottom: 8 }}>PREVIEW — {dataB.length} rows · {headersB.length} columns</div>
                  <div style={{ overflowX: "auto" }}>
                    <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 11 }}>
                      <thead><tr>{headersB.slice(0, 5).map(h => <th key={h} style={{ padding: "4px 10px", background: "#1e293b", color: "#94a3b8", textAlign: "left" }}>{h}</th>)}{headersB.length > 5 && <th style={{ color: "#475569" }}>+{headersB.length - 5} more</th>}</tr></thead>
                      <tbody>{dataB.slice(0, 3).map((r, i) => <tr key={i}>{headersB.slice(0, 5).map(h => <td key={h} style={{ padding: "4px 10px", borderBottom: "1px solid #1e293b", color: "#cbd5e1" }}>{r[h]}</td>)}</tr>)}</tbody>
                    </table>
                  </div>
                </div>}
              </div>
            </div>
            <div style={{ textAlign: "center", marginTop: 32 }}>
              <button onClick={proceedToConfigure} disabled={!dataA.length || !dataB.length} style={{ background: dataA.length && dataB.length ? "linear-gradient(135deg,#0ea5e9,#6366f1)" : "#1e293b", color: dataA.length && dataB.length ? "#fff" : "#475569", border: "none", borderRadius: 10, padding: "14px 40px", fontSize: 15, fontWeight: 700, cursor: dataA.length && dataB.length ? "pointer" : "not-allowed", transition: "all .2s", fontFamily: "inherit" }}>
                Continue to Configuration →
              </button>
            </div>
          </div>
        )}

        {/* STEP 1: Configure */}
        {step === 1 && (
          <div>
            <h1 style={{ fontFamily: "'Syne'", fontSize: 30, fontWeight: 800, marginBottom: 4, background: "linear-gradient(90deg,#38bdf8,#818cf8)", WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent" }}>Configure Comparison</h1>
            <p style={{ color: "#64748b", marginBottom: 28 }}>AI has auto-suggested field mappings. Review and adjust as needed.</p>

            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 24, marginBottom: 24 }}>
              {/* AI Mapping */}
              <div style={{ background: "#0a1628", borderRadius: 14, padding: 24, border: "1px solid #1e293b" }}>
                <div style={{ fontWeight: 700, color: "#38bdf8", marginBottom: 16, fontSize: 14 }}>🤖 AI FIELD MAPPING SUGGESTIONS</div>
                {mappings.map((m, i) => (
                  <div key={i} style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10 }}>
                    <select value={m.from} onChange={e => { const nm = [...mappings]; nm[i].from = e.target.value; setMappings(nm); }} style={{ flex: 1, background: "#0f172a", border: "1px solid #334155", borderRadius: 8, padding: "8px 12px", color: "#e2e8f0", fontSize: 13, fontFamily: "inherit" }}>
                      {headersA.map(h => <option key={h}>{h}</option>)}
                    </select>
                    <div style={{ color: "#0ea5e9", fontSize: 18 }}>→</div>
                    <select value={m.to} onChange={e => { const nm = [...mappings]; nm[i].to = e.target.value; setMappings(nm); }} style={{ flex: 1, background: "#0f172a", border: "1px solid #334155", borderRadius: 8, padding: "8px 12px", color: "#e2e8f0", fontSize: 13, fontFamily: "inherit" }}>
                      {headersB.map(h => <option key={h}>{h}</option>)}
                    </select>
                    <div style={{ fontSize: 11, color: m.confidence >= 80 ? "#22c55e" : "#f59e0b", minWidth: 40 }}>{m.confidence}%</div>
                    <button onClick={() => setMappings(mappings.filter((_, j) => j !== i))} style={{ background: "#450a0a", border: "none", color: "#ef4444", borderRadius: 6, padding: "6px 10px", cursor: "pointer" }}>✕</button>
                  </div>
                ))}
                <button onClick={() => setMappings([...mappings, { from: headersA[0], to: headersB[0], confidence: 50 }])} style={{ background: "#0f172a", border: "1px dashed #334155", color: "#64748b", borderRadius: 8, padding: "8px 16px", cursor: "pointer", fontSize: 13, fontFamily: "inherit", marginTop: 8 }}>+ Add Mapping</button>
              </div>

              {/* Key + Compare + Tolerance */}
              <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
                <div style={{ background: "#0a1628", borderRadius: 14, padding: 20, border: "1px solid #1e293b" }}>
                  <div style={{ fontWeight: 700, color: "#f59e0b", marginBottom: 12, fontSize: 14 }}>🔑 KEY FIELDS (from Dataset A)</div>
                  <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                    {headersA.map(h => (
                      <button key={h} onClick={() => setKeyFields(prev => prev.includes(h) ? prev.filter(k => k !== h) : [...prev, h])} style={{ background: keyFields.includes(h) ? "#b45309" : "#0f172a", border: `1px solid ${keyFields.includes(h) ? "#f59e0b" : "#334155"}`, color: keyFields.includes(h) ? "#fef3c7" : "#94a3b8", borderRadius: 8, padding: "6px 14px", cursor: "pointer", fontSize: 12, fontFamily: "inherit" }}>{h}</button>
                    ))}
                  </div>
                </div>

                <div style={{ background: "#0a1628", borderRadius: 14, padding: 20, border: "1px solid #1e293b" }}>
                  <div style={{ fontWeight: 700, color: "#818cf8", marginBottom: 12, fontSize: 14 }}>📊 FIELDS TO COMPARE</div>
                  {compareFields.map((cf, i) => (
                    <div key={i} style={{ display: "flex", gap: 8, marginBottom: 8, alignItems: "center" }}>
                      <select value={cf.fieldA} onChange={e => { const nf = [...compareFields]; nf[i].fieldA = e.target.value; setCompareFields(nf); }} style={{ flex: 1, background: "#0f172a", border: "1px solid #334155", borderRadius: 8, padding: "6px 10px", color: "#e2e8f0", fontSize: 12, fontFamily: "inherit" }}>
                        <option value="">-- A field --</option>
                        {headersA.map(h => <option key={h}>{h}</option>)}
                      </select>
                      <span style={{ color: "#818cf8" }}>↔</span>
                      <select value={cf.fieldB} onChange={e => { const nf = [...compareFields]; nf[i].fieldB = e.target.value; setCompareFields(nf); }} style={{ flex: 1, background: "#0f172a", border: "1px solid #334155", borderRadius: 8, padding: "6px 10px", color: "#e2e8f0", fontSize: 12, fontFamily: "inherit" }}>
                        <option value="">-- B field --</option>
                        {headersB.map(h => <option key={h}>{h}</option>)}
                      </select>
                      <button onClick={() => setCompareFields(compareFields.filter((_, j) => j !== i))} style={{ background: "#450a0a", border: "none", color: "#ef4444", borderRadius: 6, padding: "5px 9px", cursor: "pointer" }}>✕</button>
                    </div>
                  ))}
                  <button onClick={() => setCompareFields([...compareFields, { fieldA: "", fieldB: "" }])} style={{ background: "#0f172a", border: "1px dashed #334155", color: "#64748b", borderRadius: 8, padding: "7px 14px", cursor: "pointer", fontSize: 12, fontFamily: "inherit", marginTop: 4 }}>+ Add Field</button>
                </div>

                <div style={{ background: "#0a1628", borderRadius: 14, padding: 20, border: "1px solid #1e293b", display: "flex", alignItems: "center", gap: 16 }}>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontWeight: 700, color: "#22c55e", marginBottom: 6, fontSize: 14 }}>⚙️ TOLERANCE LEVEL</div>
                    <div style={{ fontSize: 12, color: "#64748b" }}>Acceptable % difference for numeric fields</div>
                  </div>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <input type="number" min="0" max="100" step="0.1" value={tolerance} onChange={e => setTolerance(e.target.value)} style={{ width: 70, background: "#0f172a", border: "1px solid #22c55e", borderRadius: 8, padding: "8px 12px", color: "#22c55e", fontSize: 16, fontWeight: 700, textAlign: "center", fontFamily: "inherit" }} />
                    <span style={{ color: "#22c55e", fontWeight: 700 }}>%</span>
                  </div>
                </div>
              </div>
            </div>

            <div style={{ display: "flex", gap: 12, justifyContent: "center" }}>
              <button onClick={() => setStep(0)} style={{ background: "#0f172a", border: "1px solid #334155", color: "#94a3b8", borderRadius: 10, padding: "12px 28px", fontSize: 14, cursor: "pointer", fontFamily: "inherit" }}>← Back</button>
              <button onClick={runComparison} style={{ background: "linear-gradient(135deg,#0ea5e9,#6366f1)", color: "#fff", border: "none", borderRadius: 10, padding: "12px 36px", fontSize: 15, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>
                🚀 Run Comparison
              </button>
            </div>
          </div>
        )}

        {/* STEP 3: Results */}
        {step === 3 && stats && (
          <div>
            {/* Summary Cards */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(7,1fr)", gap: 12, marginBottom: 24 }}>
              {[
                { label: "Total A", val: stats.totalA, color: "#38bdf8" },
                { label: "Total B", val: stats.totalB, color: "#818cf8" },
                { label: "Matched", val: stats.matched, color: "#22c55e" },
                { label: "Mismatched", val: stats.mismatched, color: "#ef4444" },
                { label: "Only in A", val: stats.onlyA, color: "#f59e0b" },
                { label: "Only in B", val: stats.onlyB, color: "#a855f7" },
                { label: "Duplicates", val: stats.duplicates, color: "#64748b" },
              ].map(s => (
                <div key={s.label} style={{ background: "#0a1628", border: `1px solid ${s.color}44`, borderRadius: 12, padding: "16px 12px", textAlign: "center", borderTop: `3px solid ${s.color}` }}>
                  <div style={{ fontSize: 26, fontWeight: 800, color: s.color, fontFamily: "'Syne'" }}>{s.val}</div>
                  <div style={{ fontSize: 11, color: "#64748b", marginTop: 4 }}>{s.label}</div>
                </div>
              ))}
            </div>

            {/* Progress bar */}
            <div style={{ background: "#0a1628", borderRadius: 12, padding: 16, marginBottom: 24, border: "1px solid #1e293b" }}>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: "#64748b", marginBottom: 8 }}>
                <span>Match Rate</span>
                <span style={{ color: "#22c55e", fontWeight: 700 }}>{((stats.matched / (stats.matched + stats.mismatched) || 0) * 100).toFixed(1)}%</span>
              </div>
              <div style={{ height: 10, background: "#1e293b", borderRadius: 99, overflow: "hidden", display: "flex" }}>
                {[
                  { w: stats.matched, c: "#22c55e" },
                  { w: stats.mismatched, c: "#ef4444" },
                  { w: stats.onlyA, c: "#f59e0b" },
                  { w: stats.onlyB, c: "#a855f7" },
                ].map((s, i) => {
                  const total = stats.matched + stats.mismatched + stats.onlyA + stats.onlyB || 1;
                  return <div key={i} style={{ width: `${(s.w / total) * 100}%`, background: s.c, transition: "width .5s" }} />;
                })}
              </div>
            </div>

            {/* Tabs */}
            <div style={{ display: "flex", gap: 4, marginBottom: 20 }}>
              {["dashboard", "detail", "sheet"].map(t => (
                <button key={t} onClick={() => setActiveTab(t)} style={{ background: activeTab === t ? "#0ea5e9" : "#0a1628", border: `1px solid ${activeTab === t ? "#0ea5e9" : "#1e293b"}`, color: activeTab === t ? "#fff" : "#64748b", borderRadius: 8, padding: "9px 20px", fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "inherit", textTransform: "capitalize" }}>
                  {t === "dashboard" ? "📊 Dashboard" : t === "detail" ? "🔍 Detail Report" : "📋 Comparison Sheet"}
                </button>
              ))}
              <div style={{ flex: 1 }} />
              <button onClick={exportCSV} style={{ background: "#14532d", border: "1px solid #22c55e", color: "#22c55e", borderRadius: 8, padding: "9px 20px", fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>⬇ Export CSV</button>
              <button onClick={() => setStep(1)} style={{ background: "#0f172a", border: "1px solid #334155", color: "#94a3b8", borderRadius: 8, padding: "9px 16px", fontSize: 13, cursor: "pointer", fontFamily: "inherit" }}>← Reconfigure</button>
            </div>

            {/* DASHBOARD TAB */}
            {activeTab === "dashboard" && (
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20 }}>
                {/* Donut-like chart via SVG */}
                <div style={{ background: "#0a1628", borderRadius: 14, padding: 24, border: "1px solid #1e293b" }}>
                  <div style={{ fontWeight: 700, color: "#e2e8f0", marginBottom: 16 }}>Classification Breakdown</div>
                  <div style={{ display: "flex", alignItems: "center", gap: 24 }}>
                    <svg width="140" height="140" viewBox="0 0 140 140">
                      {(() => {
                        const segs = [
                          { val: stats.matched, c: "#22c55e" },
                          { val: stats.mismatched, c: "#ef4444" },
                          { val: stats.onlyA, c: "#f59e0b" },
                          { val: stats.onlyB, c: "#a855f7" },
                        ];
                        const total = segs.reduce((s, x) => s + x.val, 0) || 1;
                        const r = 55, cx = 70, cy = 70;
                        let angle = -Math.PI / 2;
                        return segs.map((s, i) => {
                          const pct = s.val / total;
                          const sweep = pct * 2 * Math.PI;
                          const x1 = cx + r * Math.cos(angle), y1 = cy + r * Math.sin(angle);
                          angle += sweep;
                          const x2 = cx + r * Math.cos(angle), y2 = cy + r * Math.sin(angle);
                          if (s.val === 0) return null;
                          return <path key={i} d={`M${cx},${cy} L${x1},${y1} A${r},${r} 0 ${sweep > Math.PI ? 1 : 0} 1 ${x2},${y2} Z`} fill={s.c} opacity={0.85} />;
                        });
                      })()}
                      <circle cx="70" cy="70" r="32" fill="#0a1628" />
                      <text x="70" y="74" textAnchor="middle" fill="#e2e8f0" fontSize="14" fontWeight="bold" fontFamily="IBM Plex Mono">{((stats.matched / (stats.matched + stats.mismatched || 1)) * 100).toFixed(0)}%</text>
                    </svg>
                    <div style={{ flex: 1 }}>
                      {[["Matched", stats.matched, "#22c55e"], ["Mismatched", stats.mismatched, "#ef4444"], ["Only A", stats.onlyA, "#f59e0b"], ["Only B", stats.onlyB, "#a855f7"]].map(([l, v, c]) => (
                        <div key={l} style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10 }}>
                          <div style={{ width: 10, height: 10, borderRadius: 2, background: c }} />
                          <span style={{ flex: 1, fontSize: 13, color: "#94a3b8" }}>{l}</span>
                          <span style={{ fontWeight: 700, color: c }}>{v}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>

                {/* Field mismatch breakdown */}
                <div style={{ background: "#0a1628", borderRadius: 14, padding: 24, border: "1px solid #1e293b" }}>
                  <div style={{ fontWeight: 700, color: "#e2e8f0", marginBottom: 16 }}>Mismatches by Field</div>
                  {compareFields.filter(f => f.fieldA && f.fieldB).map(cf => {
                    const mismatches = results.results.filter(r => r.details.some(d => d.fieldA === cf.fieldA && d.status === "Mismatched")).length;
                    const total = results.results.filter(r => r.status !== "Only in A" && r.status !== "Only in B").length || 1;
                    const pct = (mismatches / total) * 100;
                    return (
                      <div key={cf.fieldA} style={{ marginBottom: 14 }}>
                        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, marginBottom: 4 }}>
                          <span style={{ color: "#94a3b8" }}>{cf.fieldA} ↔ {cf.fieldB}</span>
                          <span style={{ color: pct > 10 ? "#ef4444" : "#22c55e", fontWeight: 700 }}>{mismatches} issues ({pct.toFixed(1)}%)</span>
                        </div>
                        <div style={{ height: 6, background: "#1e293b", borderRadius: 99 }}>
                          <div style={{ width: `${pct}%`, height: "100%", background: pct > 10 ? "#ef4444" : "#22c55e", borderRadius: 99, transition: "width .5s" }} />
                        </div>
                      </div>
                    );
                  })}
                  {compareFields.filter(f => f.fieldA && f.fieldB).length === 0 && <div style={{ color: "#475569", fontSize: 13 }}>No comparison fields configured</div>}
                </div>
              </div>
            )}

            {/* DETAIL TAB */}
            {activeTab === "detail" && (
              <div style={{ background: "#0a1628", borderRadius: 14, border: "1px solid #1e293b", overflow: "hidden" }}>
                <div style={{ padding: "16px 20px", borderBottom: "1px solid #1e293b", display: "flex", gap: 12, alignItems: "center" }}>
                  <input placeholder="Search by key..." value={searchKey} onChange={e => setSearchKey(e.target.value)} style={{ background: "#0f172a", border: "1px solid #334155", borderRadius: 8, padding: "8px 14px", color: "#e2e8f0", fontSize: 13, fontFamily: "inherit", flex: 1 }} />
                  {["All", "Matched", "Mismatched", "Only in A", "Only in B"].map(s => (
                    <button key={s} onClick={() => setFilterStatus(s)} style={{ background: filterStatus === s ? "#0ea5e9" : "#0f172a", border: `1px solid ${filterStatus === s ? "#0ea5e9" : "#334155"}`, color: filterStatus === s ? "#fff" : "#64748b", borderRadius: 8, padding: "7px 14px", cursor: "pointer", fontSize: 12, fontFamily: "inherit" }}>{s}</button>
                  ))}
                </div>
                <div style={{ overflowX: "auto", maxHeight: 460, overflowY: "auto" }}>
                  <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 12 }}>
                    <thead style={{ position: "sticky", top: 0, background: "#0a1628", zIndex: 1 }}>
                      <tr>
                        <th style={{ padding: "12px 16px", textAlign: "left", color: "#64748b", borderBottom: "1px solid #1e293b" }}>Key</th>
                        <th style={{ padding: "12px 16px", textAlign: "left", color: "#64748b", borderBottom: "1px solid #1e293b" }}>Status</th>
                        {compareFields.filter(f => f.fieldA).map(f => (
                          <th key={f.fieldA} colSpan={4} style={{ padding: "12px 16px", textAlign: "center", color: "#64748b", borderBottom: "1px solid #1e293b", borderLeft: "1px solid #1e293b" }}>{f.fieldA} ↔ {f.fieldB}</th>
                        ))}
                      </tr>
                      {compareFields.filter(f => f.fieldA).length > 0 && (
                        <tr>
                          <th colSpan={2} />
                          {compareFields.filter(f => f.fieldA).map(f => (
                            <>
                              <th key={f.fieldA + "a"} style={{ padding: "6px 10px", color: "#38bdf8", fontSize: 11, borderLeft: "1px solid #1e293b", background: "#0a1628" }}>A Value</th>
                              <th key={f.fieldA + "b"} style={{ padding: "6px 10px", color: "#818cf8", fontSize: 11, background: "#0a1628" }}>B Value</th>
                              <th key={f.fieldA + "d"} style={{ padding: "6px 10px", color: "#f59e0b", fontSize: 11, background: "#0a1628" }}>Diff</th>
                              <th key={f.fieldA + "s"} style={{ padding: "6px 10px", color: "#64748b", fontSize: 11, background: "#0a1628" }}>Status</th>
                            </>
                          ))}
                        </tr>
                      )}
                    </thead>
                    <tbody>
                      {filteredResults.slice(0, 200).map((r, i) => (
                        <tr key={i} style={{ background: i % 2 === 0 ? "#0a1628" : "#080e1a" }}>
                          <td style={{ padding: "10px 16px", color: "#cbd5e1", borderBottom: "1px solid #1e293b10", fontWeight: 500 }}>{r.key}</td>
                          <td style={{ padding: "10px 16px", borderBottom: "1px solid #1e293b10" }}><StatusBadge status={r.status} /></td>
                          {compareFields.filter(f => f.fieldA).map(f => {
                            const d = r.details.find(d => d.fieldA === f.fieldA);
                            return d ? (
                              <>
                                <td key={f.fieldA + "a"} style={{ padding: "10px 12px", borderBottom: "1px solid #1e293b10", color: "#94a3b8", borderLeft: "1px solid #1e293b10" }}>{d.valA}</td>
                                <td key={f.fieldA + "b"} style={{ padding: "10px 12px", borderBottom: "1px solid #1e293b10", color: "#94a3b8" }}>{d.valB}</td>
                                <td key={f.fieldA + "d"} style={{ padding: "10px 12px", borderBottom: "1px solid #1e293b10", color: d.diff && d.diff !== "0.00" ? "#f59e0b" : "#475569" }}>{d.diff || "—"}</td>
                                <td key={f.fieldA + "s"} style={{ padding: "10px 12px", borderBottom: "1px solid #1e293b10" }}>{d.status !== "Matched" ? <StatusBadge status={d.status} /> : <span style={{ color: "#475569" }}>✓</span>}</td>
                              </>
                            ) : (
                              <><td colSpan={4} style={{ color: "#334155", fontSize: 11, padding: "10px 12px" }}>—</td></>
                            );
                          })}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {filteredResults.length === 0 && <div style={{ padding: 40, textAlign: "center", color: "#475569" }}>No records match the current filter.</div>}
                  {filteredResults.length > 200 && <div style={{ padding: 12, textAlign: "center", color: "#475569", fontSize: 12 }}>Showing 200 of {filteredResults.length} records. Export CSV for full results.</div>}
                </div>
              </div>
            )}

            {/* COMPARISON SHEET TAB */}
            {activeTab === "sheet" && (
              <div style={{ background: "#0a1628", borderRadius: 14, border: "1px solid #1e293b", overflow: "hidden" }}>
                <div style={{ padding: "14px 20px", borderBottom: "1px solid #1e293b", fontSize: 13, color: "#64748b" }}>
                  Side-by-side comparison of all records from both datasets
                </div>
                <div style={{ overflowX: "auto", maxHeight: 500, overflowY: "auto" }}>
                  <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 11 }}>
                    <thead style={{ position: "sticky", top: 0, background: "#0d1f3c", zIndex: 1 }}>
                      <tr>
                        <th style={{ padding: "10px 14px", textAlign: "left", color: "#64748b", borderBottom: "1px solid #1e293b", minWidth: 100 }}>Key</th>
                        <th style={{ padding: "10px 14px", color: "#38bdf8", borderBottom: "1px solid #1e293b", borderLeft: "2px solid #0ea5e9" }} colSpan={headersA.length}>📋 Dataset A — {fileA?.name}</th>
                        <th style={{ padding: "10px 14px", color: "#818cf8", borderBottom: "1px solid #1e293b", borderLeft: "2px solid #6366f1" }} colSpan={headersB.length}>📋 Dataset B — {fileB?.name}</th>
                        <th style={{ padding: "10px 14px", color: "#64748b", borderBottom: "1px solid #1e293b", borderLeft: "1px solid #1e293b" }}>Status</th>
                      </tr>
                      <tr>
                        <th style={{ background: "#0a1628" }} />
                        {headersA.map(h => <th key={h} style={{ padding: "6px 12px", background: "#0d1f2e", color: "#64748b", textAlign: "left", borderLeft: h === headersA[0] ? "2px solid #0ea5e9" : "none", borderBottom: "1px solid #1e293b", whiteSpace: "nowrap" }}>{h}</th>)}
                        {headersB.map(h => <th key={h} style={{ padding: "6px 12px", background: "#100d2e", color: "#64748b", textAlign: "left", borderLeft: h === headersB[0] ? "2px solid #6366f1" : "none", borderBottom: "1px solid #1e293b", whiteSpace: "nowrap" }}>{h}</th>)}
                        <th style={{ background: "#0a1628", borderBottom: "1px solid #1e293b" }} />
                      </tr>
                    </thead>
                    <tbody>
                      {results.results.slice(0, 150).map((r, i) => (
                        <tr key={i} style={{ background: r.status === "Matched" ? (i % 2 ? "#0a1628" : "#080e1a") : r.status === "Mismatched" ? "#1a0a0a" : r.status === "Only in A" ? "#1a110a" : "#130a1a" }}>
                          <td style={{ padding: "8px 14px", color: "#94a3b8", fontWeight: 600, borderBottom: "1px solid #1e293b10", whiteSpace: "nowrap" }}>{r.key}</td>
                          {headersA.map(h => <td key={h} style={{ padding: "8px 12px", borderBottom: "1px solid #1e293b10", color: "#7dd3fc", borderLeft: h === headersA[0] ? "2px solid #0ea5e9" : "none", whiteSpace: "nowrap" }}>{r.rowA ? r.rowA[h] ?? "—" : <span style={{ color: "#334155" }}>—</span>}</td>)}
                          {headersB.map(h => <td key={h} style={{ padding: "8px 12px", borderBottom: "1px solid #1e293b10", color: "#c4b5fd", borderLeft: h === headersB[0] ? "2px solid #6366f1" : "none", whiteSpace: "nowrap" }}>{r.rowB ? r.rowB[h] ?? "—" : <span style={{ color: "#334155" }}>—</span>}</td>)}
                          <td style={{ padding: "8px 12px", borderBottom: "1px solid #1e293b10", borderLeft: "1px solid #1e293b" }}><StatusBadge status={r.status} /></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {results.results.length > 150 && <div style={{ padding: 14, textAlign: "center", color: "#475569", fontSize: 12 }}>Showing 150 of {results.results.length} records. Export CSV for full sheet.</div>}
                </div>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
