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

// ─── Auto mapping ────────────────────────────────────────────────────────────
function autoMapHeaders(headersA, headersB) {
  const norm = s => s.toLowerCase().replace(/[^a-z0-9]/g, "");
  // Build initial mappings: for each header in A, find best match in B
  const mappings = [];
  const usedB = new Set();
  for (const ha of headersA) {
    const nha = norm(ha);
    let best = null, bestScore = 0;
    for (const hb of headersB) {
      if (usedB.has(hb)) continue;
      const nhb = norm(hb);
      const score = nha === nhb ? 100 : (nha.includes(nhb) || nhb.includes(nha)) ? 75 : 0;
      if (score > bestScore) { bestScore = score; best = hb; }
    }
    mappings.push({
      colA: ha,
      colB: best && bestScore >= 75 ? best : "",
      isKey: false,
      compare: best && bestScore >= 75,
      ignoreCase: false,
    });
    if (best && bestScore >= 75) usedB.add(best);
  }
  return mappings;
}

// ─── Compare engine ──────────────────────────────────────────────────────────
function compareDatasets(dataA, dataB, mappings, tolerance) {
  const tolPct = parseFloat(tolerance) / 100 || 0;
  const keyMappings = mappings.filter(m => m.isKey && m.colA && m.colB);
  const compareMappings = mappings.filter(m => m.compare && !m.isKey && m.colA && m.colB);

  const makeKey = (row, maps, fromA) =>
    maps.map(m => ((fromA ? row[m.colA] : row[m.colB]) || "").toString().toLowerCase().trim()).join("||");

  const indexB = {};
  for (const row of dataB) {
    const key = makeKey(row, keyMappings, false);
    if (!indexB[key]) indexB[key] = [];
    indexB[key].push(row);
  }

  const results = [];
  const matchedKeysB = new Set();

  for (const rowA of dataA) {
    const key = makeKey(rowA, keyMappings, true);
    const matchesB = indexB[key] || [];
    const keyVals = Object.fromEntries(keyMappings.map(m => [m.colA, rowA[m.colA] ?? ""]));

    if (!matchesB.length) {
      results.push({ key, rowA, rowB: null, status: "Only in A", details: [], keyVals });
    } else {
      const rowB = matchesB[0];
      matchedKeysB.add(key);
      const details = compareMappings.map(m => {
        const valA = rowA[m.colA] ?? "";
        const valB = rowB[m.colB] ?? "";
        const cmpA = m.ignoreCase ? valA.toString().toLowerCase() : valA.toString();
        const cmpB = m.ignoreCase ? valB.toString().toLowerCase() : valB.toString();
        const numA = parseFloat(valA), numB = parseFloat(valB);
        const isNum = !isNaN(numA) && !isNaN(numB);
        let diff = "", status = "Matched";
        if (isNum) {
          const pct = numA !== 0 ? Math.abs(numA - numB) / Math.abs(numA) : (numB !== 0 ? 1 : 0);
          diff = (numB - numA).toFixed(2);
          if (pct > tolPct) status = "Mismatched";
        } else if (cmpA.trim() !== cmpB.trim()) {
          diff = `${valA} → ${valB}`; status = "Mismatched";
        }
        return { colA: m.colA, colB: m.colB, valA, valB, diff, status };
      });
      const overallStatus = details.some(d => d.status === "Mismatched") ? "Mismatched" : "Matched";
      results.push({ key, rowA, rowB, status: overallStatus, details, keyVals });
    }
  }
  for (const row of dataB) {
    const key = makeKey(row, keyMappings, false);
    if (!matchedKeysB.has(key)) {
      results.push({ key, rowA: null, rowB: row, status: "Only in B", details: [], keyVals: Object.fromEntries(keyMappings.map(m => [m.colA, row[m.colB] ?? ""])) });
      matchedKeysB.add(key);
    }
  }
  const keyCounts = {};
  for (const row of dataA) {
    const key = makeKey(row, keyMappings, true);
    keyCounts[key] = (keyCounts[key] || 0) + 1;
  }
  return { results, duplicatesA: Object.values(keyCounts).filter(c => c > 1).reduce((s, c) => s + c - 1, 0) };
}

// ─── Build flat comparison sheet rows (1 row per file) ───────────────────────
// Structure: Session# | Comparison# | File | Status | Composite Key | [ALL colA cols] | [ALL colB cols] | [Diff per compare col]
// Each comparison adds exactly 2 rows (File A row + File B row).
// Columns are the UNION of all columns from all sessions so the sheet grows consistently.
function buildCompSheetRows(sessions) {
  if (!sessions.length) return { headers: [], rows: [] };

  // Collect ALL unique column names from every session's files
  const allColsA = []; // ordered union of all colA from all sessions
  const allColsB = []; // ordered union of all colB from all sessions
  const allComparePairs = []; // { colA, colB } pairs for diff columns

  for (const s of sessions) {
    const firstA = s.results.find(r => r.rowA)?.rowA || {};
    const firstB = s.results.find(r => r.rowB)?.rowB || {};
    for (const col of Object.keys(firstA)) { if (!allColsA.includes(col)) allColsA.push(col); }
    for (const col of Object.keys(firstB)) { if (!allColsB.includes(col)) allColsB.push(col); }
    const cMaps = s.mappings.filter(m => m.compare && !m.isKey && m.colA && m.colB);
    for (const m of cMaps) {
      if (!allComparePairs.find(p => p.colA === m.colA && p.colB === m.colB)) {
        allComparePairs.push({ colA: m.colA, colB: m.colB });
      }
    }
  }

  // Header: fixed meta cols + all File A cols + all File B cols + diff cols
  const headers = [
    "Session #", "Comparison", "File", "File Name", "Status", "Composite Key",
    ...allColsA.map(c => `[A] ${c}`),
    ...allColsB.map(c => `[B] ${c}`),
    ...allComparePairs.map(p => `Diff: ${p.colA}`),
  ];

  const rows = [];

  sessions.forEach((s, si) => {
    const keyMaps = s.mappings.filter(m => m.isKey && m.colA && m.colB);
    const cMaps = s.mappings.filter(m => m.compare && !m.isKey && m.colA && m.colB);
    const compLabel = `${s.fileAName} vs ${s.fileBName}`;

    // Row A — one row for the entire File A (summary row with all its columns)
    // We flatten: for each column in allColsA, pick value from the FIRST matching row in rowA
    // Actually: we write one row per FILE (not per record). Since the user's intent is
    // "1 row = 1 file in a comparison", each row carries aggregate info + all distinct values
    // concatenated. But looking at the screenshots again: it's 1 row per file, so we write
    // ALL columns of that file — meaning each comparison run produces exactly 2 rows total,
    // and the row carries the file-level metadata. Individual record values go to Detail sheet.
    //
    // For the comparison sheet: each row = one file's SUMMARY for this comparison run.
    // All column values are comma-joined if there are multiple records (so you can see all data).

    const aVals = {}; // colName -> all values joined
    const bVals = {};
    for (const r of s.results) {
      if (r.rowA) { for (const [k, v] of Object.entries(r.rowA)) { aVals[k] = aVals[k] ? aVals[k] + ", " + v : v; } }
      if (r.rowB) { for (const [k, v] of Object.entries(r.rowB)) { bVals[k] = bVals[k] ? bVals[k] + ", " + v : v; } }
    }

    const compositeKeySummary = `${s.matched} matched / ${s.mismatched} mismatched / ${s.onlyA} only-A / ${s.onlyB} only-B`;
    const overallStatus = s.mismatched === 0 && s.onlyA === 0 && s.onlyB === 0 ? "Matched" : "Has Differences";

    // Diff summary per compare pair
    const diffSummary = allComparePairs.map(p => {
      const mis = s.results.filter(r => r.details?.some(d => d.colA === p.colA && d.status === "Mismatched")).length;
      return mis > 0 ? `${mis} mismatch(es)` : "OK";
    });

    // File A row
    rows.push([
      si + 1,
      compLabel,
      "A",
      s.fileAName,
      overallStatus,
      compositeKeySummary,
      ...allColsA.map(c => aVals[c] ?? ""),
      ...allColsB.map(() => ""),    // B cols blank on A row
      ...diffSummary,
    ]);

    // File B row
    rows.push([
      si + 1,
      compLabel,
      "B",
      s.fileBName,
      overallStatus,
      compositeKeySummary,
      ...allColsA.map(() => ""),    // A cols blank on B row
      ...allColsB.map(c => bVals[c] ?? ""),
      ...diffSummary,
    ]);
  });

  return { headers, rows };
}

// ─── Excel Export ─────────────────────────────────────────────────────────────
async function exportToExcel(sessions, allMappings) {
  const { utils, writeFile } = await import("https://cdn.sheetjs.com/xlsx-0.20.1/package/xlsx.mjs");

  const keyMaps = allMappings.filter(m => m.isKey && m.colA && m.colB);
  const compareMaps = allMappings.filter(m => m.compare && !m.isKey && m.colA && m.colB);

  // ── 1. Comparison Sheet — 1 row per file, appended per session ──
  const { headers: compHeaders, rows: compDataRows } = buildCompSheetRows(sessions);
  const compSheet = utils.aoa_to_sheet([compHeaders, ...compDataRows]);

  // ── 2. Summary Sheet ──
  const summaryRows = [
    ["CompareIQ — Comparison Summary"],
    [],
    ["Session #", "File A", "File B", "Total A", "Total B", "Matched", "Mismatched", "Only in A", "Only in B", "Duplicates", "Match Rate"],
    ...sessions.map((s, i) => [
      i + 1,
      s.fileAName,
      s.fileBName,
      s.totalA,
      s.totalB,
      s.matched,
      s.mismatched,
      s.onlyA,
      s.onlyB,
      s.duplicates,
      `${((s.matched / (s.matched + s.mismatched || 1)) * 100).toFixed(1)}%`,
    ])
  ];

  // ── 3. Detail Report — 1 row per record pair, all sessions ──
  const detailHeader = [
    "Session #", "File A", "File B", "Composite Key", "Status",
    ...keyMaps.map(m => m.colA),
    ...compareMaps.map(m => `${m.colA} [A]`),
    ...compareMaps.map(m => `${m.colB} [B]`),
    ...compareMaps.map(m => `Diff: ${m.colA}`),
  ];
  const detailRows = [detailHeader];
  sessions.forEach((s, si) => {
    const sKeyMaps = s.mappings.filter(m => m.isKey && m.colA && m.colB);
    const sCmpMaps = s.mappings.filter(m => m.compare && !m.isKey && m.colA && m.colB);
    for (const r of s.results) {
      const compositeKey = sKeyMaps.map(m => r.keyVals[m.colA] ?? "").join(" | ");
      detailRows.push([
        si + 1,
        s.fileAName,
        s.fileBName,
        compositeKey,
        r.status,
        ...keyMaps.map(m => r.keyVals[m.colA] ?? ""),
        ...compareMaps.map(m => {
          const sm = sCmpMaps.find(x => x.colA === m.colA);
          return sm && r.rowA ? (r.rowA[sm.colA] ?? "") : "";
        }),
        ...compareMaps.map(m => {
          const sm = sCmpMaps.find(x => x.colB === m.colB);
          return sm && r.rowB ? (r.rowB[sm.colB] ?? "") : "";
        }),
        ...compareMaps.map(m => {
          const d = r.details?.find(d => d.colA === m.colA);
          return d?.diff || "";
        }),
      ]);
    }
  });

  const wb = utils.book_new();
  utils.book_append_sheet(wb, utils.aoa_to_sheet(summaryRows), "Summary");
  utils.book_append_sheet(wb, compSheet, "Comparison Sheet");
  utils.book_append_sheet(wb, utils.aoa_to_sheet(detailRows), "Detail Report");
  writeFile(wb, "CompareIQ_Results.xlsx");
}

// ─── Styles ───────────────────────────────────────────────────────────────────
const S = {
  card: { background: "#060d1a", borderRadius: 13, border: "1px solid #0f2040", overflow: "hidden" },
  th: { padding: "9px 13px", textAlign: "left", borderBottom: "1px solid #0f2040", whiteSpace: "nowrap", fontSize: 11, fontWeight: 700 },
  td: { padding: "7px 12px", borderBottom: "1px solid #0f204018", whiteSpace: "nowrap", fontSize: 11 },
};

const StatusBadge = ({ status }) => {
  const cfg = { Matched: "#22c55e", Mismatched: "#ef4444", "Only in A": "#f59e0b", "Only in B": "#a855f7" };
  const c = cfg[status] || "#64748b";
  return <span style={{ background: c + "18", color: c, border: `1px solid ${c}40`, borderRadius: 5, padding: "2px 9px", fontSize: 10, fontWeight: 700 }}>{status}</span>;
};

const STEPS = ["Upload", "Map Columns", "Results"];

export default function CompareIQ() {
  const [step, setStep] = useState(0);
  const [fileA, setFileA] = useState(null);
  const [fileB, setFileB] = useState(null);
  const [dataA, setDataA] = useState([]);
  const [dataB, setDataB] = useState([]);
  const [headersA, setHeadersA] = useState([]);
  const [headersB, setHeadersB] = useState([]);
  // Column mapping rows: [{ colA, colB, isKey, compare, ignoreCase }]
  const [mappings, setMappings] = useState([]);
  const [tolerance, setTolerance] = useState("1");
  const [results, setResults] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [activeTab, setActiveTab] = useState("dashboard");
  const [filterStatus, setFilterStatus] = useState("All");
  const [searchKey, setSearchKey] = useState("");
  // Accumulated sessions for export
  const [sessions, setSessions] = useState([]);

  const handleFileUpload = async (file, which) => {
    setError(""); setLoading(true);
    try {
      const data = await parseFile(file);
      if (!data.length) throw new Error("No data found in file");
      const headers = Object.keys(data[0]);
      if (which === "A") { setFileA(file); setDataA(data); setHeadersA(headers); }
      else { setFileB(file); setDataB(data); setHeadersB(headers); }
    } catch (e) { setError(`Error parsing: ${e.message}`); }
    setLoading(false);
  };

  const proceedToMap = () => {
    if (!dataA.length || !dataB.length) { setError("Upload both files first."); return; }
    const auto = autoMapHeaders(headersA, headersB);
    // Auto-mark preferred key fields
    const norm = s => s.toLowerCase().replace(/[^a-z0-9]/g, "");
    const prefKeys = ["personnumber", "person number", "balancename", "balance name", "area1", "area2", "area3"];
    const updated = auto.map(m => ({
      ...m,
      isKey: prefKeys.some(p => norm(m.colA).includes(p.replace(/ /g, ""))),
      compare: !prefKeys.some(p => norm(m.colA).includes(p.replace(/ /g, ""))),
    }));
    setMappings(updated);
    setStep(1);
  };

  const updateMapping = (i, field, value) => {
    setMappings(prev => prev.map((m, idx) => idx === i ? { ...m, [field]: value } : m));
  };

  const addMappingRow = () => {
    setMappings(prev => [...prev, { colA: headersA[0] || "", colB: headersB[0] || "", isKey: false, compare: true, ignoreCase: false }]);
  };

  const removeMappingRow = (i) => {
    setMappings(prev => prev.filter((_, idx) => idx !== i));
  };

  const runComparison = () => {
    const keyMaps = mappings.filter(m => m.isKey && m.colA && m.colB);
    if (!keyMaps.length) { setError("Mark at least one KEY column."); return; }
    setLoading(true);
    setTimeout(() => {
      const res = compareDatasets(dataA, dataB, mappings, tolerance);
      setResults(res);
      const matched = res.results.filter(r => r.status === "Matched").length;
      const mismatched = res.results.filter(r => r.status === "Mismatched").length;
      const onlyA = res.results.filter(r => r.status === "Only in A").length;
      const onlyB = res.results.filter(r => r.status === "Only in B").length;
      setSessions(prev => [...prev, {
        fileAName: fileA?.name || "File A",
        fileBName: fileB?.name || "File B",
        results: res.results,
        totalA: dataA.length,
        totalB: dataB.length,
        matched, mismatched, onlyA, onlyB,
        duplicates: res.duplicatesA,
        mappings: [...mappings],
      }]);
      setStep(2);
      setLoading(false);
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

  const keyMappings = mappings.filter(m => m.isKey && m.colA && m.colB);
  const compareMappings = mappings.filter(m => m.compare && !m.isKey && m.colA && m.colB);

  const filteredResults = results ? results.results.filter(r =>
    (filterStatus === "All" || r.status === filterStatus) &&
    (!searchKey || r.key.toLowerCase().includes(searchKey.toLowerCase()))
  ) : [];

  // ── Subcomponents ─────────────────────────────────────────────────────────
  const DropZone = ({ label, file, onFile, color }) => {
    const ref = useRef();
    const [drag, setDrag] = useState(false);
    return (
      <div onClick={() => ref.current.click()}
        onDragOver={e => { e.preventDefault(); setDrag(true); }}
        onDragLeave={() => setDrag(false)}
        onDrop={e => { e.preventDefault(); setDrag(false); if (e.dataTransfer.files[0]) onFile(e.dataTransfer.files[0]); }}
        style={{ border: `2px dashed ${drag ? color : file ? "#22c55e" : "#1e3a5f"}`, borderRadius: 13, padding: "22px 16px", textAlign: "center", cursor: "pointer", background: drag ? "#0c1e38" : file ? "#0a1f12" : "#060d1a", transition: "all .2s" }}>
        <input ref={ref} type="file" accept=".csv,.xlsx,.xls,.txt,.tsv,.pdf,.tex" style={{ display: "none" }} onChange={e => { if (e.target.files[0]) onFile(e.target.files[0]); }} />
        <div style={{ fontSize: 24, marginBottom: 6 }}>{file ? "✅" : "📂"}</div>
        <div style={{ fontWeight: 700, color: file ? "#22c55e" : "#475569", fontSize: 13 }}>{file ? file.name : label}</div>
        {file ? <div style={{ color: "#334155", fontSize: 11, marginTop: 3 }}>{(file.size / 1024).toFixed(1)} KB · click to replace</div>
          : <div style={{ color: "#1e3a5f", fontSize: 11, marginTop: 5 }}>CSV · Excel · TXT · TSV · PDF · TeX</div>}
      </div>
    );
  };

  const Radio = ({ checked, onChange, color = "#0ea5e9" }) => (
    <div onClick={onChange} style={{ width: 18, height: 18, borderRadius: "50%", border: `2px solid ${checked ? color : "#1e3a5f"}`, background: checked ? color : "transparent", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0, transition: "all .15s", margin: "auto" }}>
      {checked && <div style={{ width: 7, height: 7, borderRadius: "50%", background: "#fff" }} />}
    </div>
  );

  const Toggle = ({ checked, onChange, color = "#0ea5e9" }) => (
    <div onClick={onChange} style={{ width: 18, height: 18, borderRadius: 4, border: `2px solid ${checked ? color : "#1e3a5f"}`, background: checked ? color : "transparent", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0, transition: "all .15s", margin: "auto" }}>
      {checked && <span style={{ color: "#fff", fontSize: 11, fontWeight: 900, lineHeight: 1 }}>✓</span>}
    </div>
  );

  const ColSelect = ({ value, options, onChange, color }) => (
    <div style={{ position: "relative", flex: 1 }}>
      <select value={value} onChange={e => onChange(e.target.value)}
        style={{ width: "100%", background: "#04080f", border: `1px solid ${color}33`, borderRadius: 7, padding: "6px 28px 6px 10px", color: value ? "#e2e8f0" : "#334155", fontSize: 12, fontFamily: "inherit", appearance: "none", cursor: "pointer" }}>
        <option value="">— not mapped —</option>
        {options.map(o => <option key={o} value={o}>{o}</option>)}
      </select>
      <span style={{ position: "absolute", right: 8, top: "50%", transform: "translateY(-50%)", color: "#334155", pointerEvents: "none", fontSize: 10 }}>▾</span>
    </div>
  );

  // ── Render ─────────────────────────────────────────────────────────────────
  return (
    <div style={{ minHeight: "100vh", background: "#04080f", color: "#e2e8f0", fontFamily: "'DM Mono','Courier New',monospace" }}>
      <link href="https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=Outfit:wght@700;800&display=swap" rel="stylesheet" />

      {/* Header */}
      <div style={{ background: "#060d1a", borderBottom: "1px solid #0f2040", padding: "12px 24px", display: "flex", alignItems: "center", justifyContent: "space-between", position: "sticky", top: 0, zIndex: 50 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ width: 30, height: 30, borderRadius: 7, background: "linear-gradient(135deg,#0ea5e9,#6366f1)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 15 }}>⚖️</div>
          <div>
            <div style={{ fontFamily: "'Outfit',sans-serif", fontWeight: 800, fontSize: 15, background: "linear-gradient(90deg,#38bdf8,#818cf8)", WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent" }}>CompareIQ</div>
            <div style={{ fontSize: 9, color: "#1e3a5f", letterSpacing: 1.5 }}>INTELLIGENT DATASET COMPARISON</div>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
          {STEPS.map((s, i) => (
            <div key={s} style={{ display: "flex", alignItems: "center", gap: 3 }}>
              <div onClick={() => step > i && setStep(i)} style={{ display: "flex", alignItems: "center", gap: 6, padding: "5px 12px", borderRadius: 20, background: step === i ? "#0c2040" : "transparent", border: `1px solid ${step === i ? "#0ea5e9" : step > i ? "#1e3a5f" : "#0a1628"}`, color: step === i ? "#38bdf8" : step > i ? "#1e4060" : "#1e3a5f", fontSize: 11, fontWeight: 700, cursor: step > i ? "pointer" : "default" }}>
                <div style={{ width: 14, height: 14, borderRadius: "50%", background: step > i ? "#22c55e" : step === i ? "#0ea5e9" : "#0a1628", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 8, color: "#fff", fontWeight: 900 }}>{step > i ? "✓" : i + 1}</div>
                {s}
              </div>
              {i < STEPS.length - 1 && <div style={{ width: 14, height: 1, background: "#0a1628" }} />}
            </div>
          ))}
          {sessions.length > 0 && (
            <div style={{ marginLeft: 12, padding: "4px 12px", background: "#0a2010", border: "1px solid #22c55e33", borderRadius: 20, fontSize: 10, color: "#22c55e" }}>
              {sessions.length} session{sessions.length > 1 ? "s" : ""} accumulated
            </div>
          )}
        </div>
      </div>

      <div style={{ maxWidth: 1100, margin: "0 auto", padding: "24px 20px" }}>
        {error && <div style={{ background: "#140505", border: "1px solid #7f1d1d", color: "#fca5a5", borderRadius: 9, padding: "9px 14px", marginBottom: 14, fontSize: 12, display: "flex", justifyContent: "space-between" }}>⚠️ {error}<button onClick={() => setError("")} style={{ background: "none", border: "none", color: "#fca5a5", cursor: "pointer" }}>✕</button></div>}
        {loading && <div style={{ background: "#050f1f", border: "1px solid #0ea5e9", borderRadius: 9, padding: "9px 14px", marginBottom: 14, fontSize: 12, color: "#38bdf8" }}>⏳ Processing…</div>}

        {/* ══ STEP 0 — UPLOAD ══ */}
        {step === 0 && (
          <div>
            <div style={{ textAlign: "center", marginBottom: 28 }}>
              <h1 style={{ fontFamily: "'Outfit',sans-serif", fontWeight: 800, fontSize: 28, margin: "0 0 6px", background: "linear-gradient(90deg,#38bdf8,#818cf8)", WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent" }}>Upload Your Datasets</h1>
              <p style={{ color: "#334155", margin: 0, fontSize: 13 }}>CSV · Excel · TXT · TSV · PDF · TeX · Pipe-delimited</p>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 24 }}>
              {[
                { label: "Drop Dataset A — Source / HR / Expected", which: "A", file: fileA, data: dataA, headers: headersA, color: "#0ea5e9" },
                { label: "Drop Dataset B — Target / Payroll / Actual", which: "B", file: fileB, data: dataB, headers: headersB, color: "#818cf8" }
              ].map(({ label, which, file, data, headers, color }) => (
                <div key={which}>
                  <div style={{ color, fontSize: 10, fontWeight: 700, marginBottom: 7, letterSpacing: 1.5 }}>● FILE {which}</div>
                  <DropZone label={label} file={file} color={color} onFile={f => handleFileUpload(f, which)} />
                  {data.length > 0 && (
                    <div style={{ marginTop: 8, background: "#060d1a", borderRadius: 9, padding: 11, border: `1px solid ${color}22` }}>
                      <div style={{ color, fontSize: 10, fontWeight: 700, marginBottom: 6 }}>{data.length} rows · {headers.length} columns</div>
                      <div style={{ overflowX: "auto" }}>
                        <table style={{ borderCollapse: "collapse", fontSize: 10 }}>
                          <thead><tr>{headers.slice(0, 7).map(h => <th key={h} style={{ padding: "3px 8px", background: "#04080f", color: "#334155", textAlign: "left", whiteSpace: "nowrap" }}>{h}</th>)}{headers.length > 7 && <th style={{ color: "#1e3a5f", padding: "3px 6px" }}>+{headers.length - 7}</th>}</tr></thead>
                          <tbody>{data.slice(0, 2).map((r, i) => <tr key={i}>{headers.slice(0, 7).map(h => <td key={h} style={{ padding: "3px 8px", color: "#475569", borderBottom: "1px solid #04080f", whiteSpace: "nowrap" }}>{r[h]}</td>)}</tr>)}</tbody>
                        </table>
                      </div>
                    </div>
                  )}
                </div>
              ))}
            </div>
            <div style={{ textAlign: "center" }}>
              <button onClick={proceedToMap} disabled={!dataA.length || !dataB.length}
                style={{ background: dataA.length && dataB.length ? "linear-gradient(135deg,#0ea5e9,#6366f1)" : "#0a1628", color: dataA.length && dataB.length ? "#fff" : "#1e3a5f", border: "none", borderRadius: 9, padding: "11px 32px", fontSize: 14, fontWeight: 700, cursor: dataA.length && dataB.length ? "pointer" : "not-allowed", fontFamily: "inherit" }}>
                Continue to Column Mapping →
              </button>
            </div>
          </div>
        )}

        {/* ══ STEP 1 — COLUMN MAPPING ══ */}
        {step === 1 && (
          <div>
            <div style={{ marginBottom: 18 }}>
              <h1 style={{ fontFamily: "'Outfit',sans-serif", fontWeight: 800, fontSize: 22, margin: "0 0 4px", background: "linear-gradient(90deg,#38bdf8,#818cf8)", WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent" }}>Column Mapping Configuration</h1>
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginTop: 8 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 6, padding: "5px 12px", background: "#05131f", border: "1px solid #0ea5e933", borderRadius: 20, fontSize: 11 }}>
                  <span style={{ color: "#0ea5e9" }}>●</span> <span style={{ color: "#38bdf8" }}>{fileA?.name}</span>
                </div>
                <span style={{ color: "#334155", fontSize: 16 }}>⇌</span>
                <div style={{ display: "flex", alignItems: "center", gap: 6, padding: "5px 12px", background: "#0d0a1f", border: "1px solid #818cf833", borderRadius: 20, fontSize: 11 }}>
                  <span style={{ color: "#818cf8" }}>●</span> <span style={{ color: "#818cf8" }}>{fileB?.name}</span>
                </div>
                <div style={{ flex: 1 }} />
                <div style={{ display: "flex", gap: 6, fontSize: 11, color: "#334155" }}>
                  <span>Total: <b style={{ color: "#38bdf8" }}>{mappings.length}</b></span>
                  <span>Keys: <b style={{ color: "#f59e0b" }}>{mappings.filter(m => m.isKey).length}</b></span>
                  <span>Compare: <b style={{ color: "#22c55e" }}>{mappings.filter(m => m.compare && !m.isKey).length}</b></span>
                </div>
              </div>
            </div>

            {/* Info banner */}
            <div style={{ background: "#050f1f", border: "1px solid #0ea5e922", borderRadius: 9, padding: "10px 14px", marginBottom: 14, fontSize: 12, color: "#475569", display: "flex", alignItems: "center", gap: 10 }}>
              <span style={{ color: "#0ea5e9", fontSize: 14 }}>ℹ</span>
              Map columns from both datasets, set <span style={{ color: "#f59e0b", fontWeight: 700 }}>key columns</span> for row matching, toggle <span style={{ color: "#22c55e", fontWeight: 700 }}>Compare</span> for value comparison, and set Ignore Case or Tolerance as needed.
            </div>

            {/* Tolerance row */}
            <div style={{ display: "flex", alignItems: "center", gap: 16, marginBottom: 12, padding: "10px 14px", background: "#060d1a", borderRadius: 9, border: "1px solid #0f2040" }}>
              <span style={{ fontSize: 11, color: "#475569", fontWeight: 700 }}>⚙ GLOBAL TOLERANCE</span>
              <input type="number" min="0" max="100" step="0.1" value={tolerance} onChange={e => setTolerance(e.target.value)}
                style={{ width: 55, background: "#04080f", border: "1px solid #22c55e44", borderRadius: 6, padding: "5px 8px", color: "#22c55e", fontSize: 14, fontWeight: 700, textAlign: "center", fontFamily: "inherit" }} />
              <span style={{ color: "#22c55e", fontWeight: 700 }}>%</span>
              <span style={{ fontSize: 11, color: "#1e3a5f" }}>Acceptable numeric variance for all Compare columns</span>
              <div style={{ flex: 1 }} />
              <button onClick={addMappingRow} style={{ background: "#0c2040", border: "1px solid #0ea5e944", color: "#38bdf8", borderRadius: 7, padding: "6px 14px", cursor: "pointer", fontSize: 11, fontFamily: "inherit", fontWeight: 700 }}>+ Add Row</button>
            </div>

            {/* Mapping table */}
            <div style={{ ...S.card, marginBottom: 16 }}>
              {/* Table header */}
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 70px 80px 90px 32px", gap: 0, background: "#04080f", padding: "9px 14px", borderBottom: "1px solid #0f2040" }}>
                {[
                  { label: `DATASET 1 — ${fileA?.name}`, color: "#38bdf8" },
                  { label: `DATASET 2 — ${fileB?.name}`, color: "#818cf8" },
                  { label: "KEY 🔑", color: "#f59e0b" },
                  { label: "COMPARE ✓", color: "#22c55e" },
                  { label: "IGNORE CASE", color: "#94a3b8" },
                  { label: "", color: "" },
                ].map((h, i) => (
                  <div key={i} style={{ fontSize: 10, fontWeight: 700, color: h.color, letterSpacing: .5, textAlign: i >= 2 ? "center" : "left", padding: "0 4px" }}>{h.label}</div>
                ))}
              </div>

              {/* Mapping rows */}
              <div style={{ maxHeight: 400, overflowY: "auto" }}>
                {mappings.map((m, i) => (
                  <div key={i} style={{ display: "grid", gridTemplateColumns: "1fr 1fr 70px 80px 90px 32px", gap: 0, alignItems: "center", padding: "7px 14px", borderBottom: "1px solid #0f204020", background: m.isKey ? "#0f1a0a" : i % 2 ? "#04080f" : "#060d1a", transition: "background .15s" }}>

                    {/* Dataset A column */}
                    <div style={{ padding: "0 6px 0 0" }}>
                      <ColSelect value={m.colA} options={headersA} color="#0ea5e9" onChange={v => updateMapping(i, "colA", v)} />
                    </div>

                    {/* Dataset B column */}
                    <div style={{ padding: "0 6px" }}>
                      <ColSelect value={m.colB} options={headersB} color="#818cf8" onChange={v => updateMapping(i, "colB", v)} />
                    </div>

                    {/* KEY radio */}
                    <div style={{ textAlign: "center" }}>
                      <Radio checked={m.isKey} color="#f59e0b"
                        onChange={() => updateMapping(i, "isKey", !m.isKey)} />
                    </div>

                    {/* COMPARE toggle */}
                    <div style={{ textAlign: "center" }}>
                      <Toggle checked={m.compare && !m.isKey} color="#22c55e"
                        onChange={() => { if (!m.isKey) updateMapping(i, "compare", !m.compare); }} />
                    </div>

                    {/* IGNORE CASE toggle */}
                    <div style={{ textAlign: "center" }}>
                      <Toggle checked={m.ignoreCase} color="#94a3b8"
                        onChange={() => updateMapping(i, "ignoreCase", !m.ignoreCase)} />
                    </div>

                    {/* Remove */}
                    <div style={{ textAlign: "center" }}>
                      <button onClick={() => removeMappingRow(i)} style={{ background: "none", border: "none", color: "#334155", cursor: "pointer", fontSize: 14, lineHeight: 1 }}>✕</button>
                    </div>
                  </div>
                ))}
              </div>

              {/* Footer summary */}
              <div style={{ padding: "8px 14px", background: "#04080f", borderTop: "1px solid #0f2040", display: "flex", gap: 16, fontSize: 11 }}>
                <span style={{ color: "#475569" }}>Total: <b style={{ color: "#38bdf8" }}>{mappings.length}</b></span>
                <span style={{ color: "#475569" }}>Keys: <b style={{ color: "#f59e0b" }}>{mappings.filter(m => m.isKey).length}</b></span>
                <span style={{ color: "#475569" }}>Compare: <b style={{ color: "#22c55e" }}>{mappings.filter(m => m.compare && !m.isKey).length}</b></span>
                <span style={{ color: "#475569" }}>Ignore Case: <b style={{ color: "#94a3b8" }}>{mappings.filter(m => m.ignoreCase).length}</b></span>
              </div>
            </div>

            <div style={{ display: "flex", gap: 10, justifyContent: "center" }}>
              <button onClick={() => setStep(0)} style={{ background: "transparent", border: "1px solid #0f2040", color: "#334155", borderRadius: 9, padding: "10px 22px", fontSize: 13, cursor: "pointer", fontFamily: "inherit" }}>← Back</button>
              <button onClick={runComparison} style={{ background: "linear-gradient(135deg,#0ea5e9,#6366f1)", color: "#fff", border: "none", borderRadius: 9, padding: "10px 32px", fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>🚀 Run Comparison</button>
            </div>
          </div>
        )}

        {/* ══ STEP 2 — RESULTS ══ */}
        {step === 2 && stats && (
          <div>
            {/* Summary cards */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(7,1fr)", gap: 8, marginBottom: 16 }}>
              {[
                { label: "Total A", val: stats.totalA, color: "#38bdf8" },
                { label: "Total B", val: stats.totalB, color: "#818cf8" },
                { label: "Matched", val: stats.matched, color: "#22c55e" },
                { label: "Mismatched", val: stats.mismatched, color: "#ef4444" },
                { label: "Only in A", val: stats.onlyA, color: "#f59e0b" },
                { label: "Only in B", val: stats.onlyB, color: "#a855f7" },
                { label: "Duplicates", val: stats.duplicates, color: "#475569" },
              ].map(s => (
                <div key={s.label} style={{ background: "#060d1a", border: `1px solid ${s.color}25`, borderTop: `3px solid ${s.color}`, borderRadius: 10, padding: "10px 6px", textAlign: "center" }}>
                  <div style={{ fontSize: 19, fontWeight: 800, color: s.color, fontFamily: "'Outfit'" }}>{s.val}</div>
                  <div style={{ fontSize: 9, color: "#1e3a5f", marginTop: 2, letterSpacing: .5 }}>{s.label.toUpperCase()}</div>
                </div>
              ))}
            </div>

            {/* Match bar */}
            <div style={{ background: "#060d1a", borderRadius: 10, padding: "10px 14px", marginBottom: 14, border: "1px solid #0f2040" }}>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11, marginBottom: 6 }}>
                <span style={{ color: "#334155" }}>Match Rate — {fileA?.name} vs {fileB?.name}</span>
                <span style={{ color: "#22c55e", fontWeight: 700 }}>{((stats.matched / (stats.matched + stats.mismatched || 1)) * 100).toFixed(1)}%</span>
              </div>
              <div style={{ height: 6, background: "#04080f", borderRadius: 99, display: "flex", overflow: "hidden" }}>
                {[{ w: stats.matched, c: "#22c55e" }, { w: stats.mismatched, c: "#ef4444" }, { w: stats.onlyA, c: "#f59e0b" }, { w: stats.onlyB, c: "#a855f7" }].map((s, i) => {
                  const t = stats.matched + stats.mismatched + stats.onlyA + stats.onlyB || 1;
                  return <div key={i} style={{ width: `${(s.w / t) * 100}%`, background: s.c }} />;
                })}
              </div>
            </div>

            {/* Sessions info */}
            {sessions.length > 1 && (
              <div style={{ background: "#060d1a", borderRadius: 10, padding: "9px 14px", marginBottom: 14, border: "1px solid #22c55e22", fontSize: 11, color: "#22c55e", display: "flex", alignItems: "center", gap: 10 }}>
                <span>📚</span>
                <span><b>{sessions.length} sessions</b> accumulated — export will include all sessions in the Comparison Sheet (2 rows per record)</span>
              </div>
            )}

            {/* Tab bar */}
            <div style={{ display: "flex", gap: 5, marginBottom: 14, alignItems: "center", flexWrap: "wrap" }}>
              {[["dashboard", "📊 Dashboard"], ["detail", "🔍 Detail Report"], ["sheet", "📋 Comparison Sheet"]].map(([t, label]) => (
                <button key={t} onClick={() => setActiveTab(t)} style={{ background: activeTab === t ? "#0c2040" : "transparent", border: `1px solid ${activeTab === t ? "#0ea5e9" : "#0f2040"}`, color: activeTab === t ? "#38bdf8" : "#1e3a5f", borderRadius: 7, padding: "7px 15px", fontSize: 12, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>{label}</button>
              ))}
              <div style={{ flex: 1 }} />
              <button onClick={() => exportToExcel(sessions, mappings)}
                style={{ background: "#0a1f0a", border: "1px solid #22c55e44", color: "#22c55e", borderRadius: 7, padding: "7px 15px", fontSize: 12, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>⬇ Export Excel ({sessions.length} session{sessions.length > 1 ? "s" : ""})</button>
              <button onClick={() => { setStep(0); setResults(null); }}
                style={{ background: "#0a1f0a", border: "1px solid #22c55e44", color: "#22c55e", borderRadius: 7, padding: "7px 15px", fontSize: 12, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>+ New Comparison</button>
              <button onClick={() => setStep(1)} style={{ background: "transparent", border: "1px solid #0f2040", color: "#1e3a5f", borderRadius: 7, padding: "7px 11px", fontSize: 11, cursor: "pointer", fontFamily: "inherit" }}>← Remap</button>
            </div>

            {/* DASHBOARD */}
            {activeTab === "dashboard" && (
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
                <div style={{ ...S.card, padding: 20 }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: "#475569", letterSpacing: .5, marginBottom: 14 }}>CLASSIFICATION BREAKDOWN</div>
                  <div style={{ display: "flex", alignItems: "center", gap: 18 }}>
                    <svg width="110" height="110" viewBox="0 0 110 110">
                      {(() => {
                        const segs = [{ val: stats.matched, c: "#22c55e" }, { val: stats.mismatched, c: "#ef4444" }, { val: stats.onlyA, c: "#f59e0b" }, { val: stats.onlyB, c: "#a855f7" }];
                        const total = segs.reduce((s, x) => s + x.val, 0) || 1;
                        const r = 42, cx = 55, cy = 55; let angle = -Math.PI / 2;
                        return segs.map((s, i) => {
                          if (!s.val) return null;
                          const sweep = (s.val / total) * 2 * Math.PI;
                          const x1 = cx + r * Math.cos(angle), y1 = cy + r * Math.sin(angle);
                          angle += sweep;
                          return <path key={i} d={`M${cx},${cy} L${x1},${y1} A${r},${r} 0 ${sweep > Math.PI ? 1 : 0} 1 ${cx + r * Math.cos(angle)},${cy + r * Math.sin(angle)} Z`} fill={s.c} opacity={0.88} />;
                        });
                      })()}
                      <circle cx="55" cy="55" r="26" fill="#060d1a" />
                      <text x="55" y="59" textAnchor="middle" fill="#e2e8f0" fontSize="11" fontWeight="bold" fontFamily="Outfit">{((stats.matched / (stats.matched + stats.mismatched || 1)) * 100).toFixed(0)}%</text>
                    </svg>
                    <div style={{ flex: 1 }}>
                      {[["Matched", stats.matched, "#22c55e"], ["Mismatched", stats.mismatched, "#ef4444"], ["Only A", stats.onlyA, "#f59e0b"], ["Only B", stats.onlyB, "#a855f7"]].map(([l, v, c]) => (
                        <div key={l} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                          <div style={{ width: 7, height: 7, borderRadius: 2, background: c, flexShrink: 0 }} />
                          <span style={{ flex: 1, fontSize: 11, color: "#334155" }}>{l}</span>
                          <span style={{ fontWeight: 700, color: c }}>{v}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>

                <div style={{ ...S.card, padding: 20 }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: "#475569", letterSpacing: .5, marginBottom: 14 }}>MISMATCHES BY FIELD</div>
                  {compareMappings.length === 0 && <div style={{ color: "#1e3a5f", fontSize: 12 }}>No compare columns configured</div>}
                  {compareMappings.slice(0, 8).map(m => {
                    const total = results.results.filter(r => r.rowA && r.rowB).length || 1;
                    const mis = results.results.filter(r => r.details.some(d => d.colA === m.colA && d.status === "Mismatched")).length;
                    const pct = (mis / total) * 100;
                    return (
                      <div key={m.colA} style={{ marginBottom: 11 }}>
                        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11, marginBottom: 4 }}>
                          <span style={{ color: "#475569" }}>{m.colA}</span>
                          <span style={{ color: pct > 10 ? "#ef4444" : "#22c55e", fontWeight: 700 }}>{mis} ({pct.toFixed(1)}%)</span>
                        </div>
                        <div style={{ height: 5, background: "#04080f", borderRadius: 99 }}>
                          <div style={{ width: `${pct}%`, height: "100%", background: pct > 10 ? "#ef4444" : "#22c55e", borderRadius: 99 }} />
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}

            {/* DETAIL REPORT */}
            {activeTab === "detail" && (
              <div style={S.card}>
                <div style={{ padding: "11px 14px", borderBottom: "1px solid #0f2040", display: "flex", gap: 7, alignItems: "center", flexWrap: "wrap" }}>
                  <input placeholder="Search composite key..." value={searchKey} onChange={e => setSearchKey(e.target.value)}
                    style={{ background: "#04080f", border: "1px solid #0f2040", borderRadius: 6, padding: "5px 11px", color: "#e2e8f0", fontSize: 11, fontFamily: "inherit", flex: 1, minWidth: 130 }} />
                  {["All", "Matched", "Mismatched", "Only in A", "Only in B"].map(s => (
                    <button key={s} onClick={() => setFilterStatus(s)} style={{ background: filterStatus === s ? "#0c2040" : "transparent", border: `1px solid ${filterStatus === s ? "#0ea5e9" : "#0f2040"}`, color: filterStatus === s ? "#38bdf8" : "#1e3a5f", borderRadius: 6, padding: "4px 10px", cursor: "pointer", fontSize: 10, fontWeight: 700, fontFamily: "inherit" }}>{s}</button>
                  ))}
                  <span style={{ fontSize: 10, color: "#1e3a5f" }}>{filteredResults.length} rows</span>
                </div>
                <div style={{ overflowX: "auto", maxHeight: 420, overflowY: "auto" }}>
                  <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 11 }}>
                    <thead style={{ position: "sticky", top: 0, background: "#060d1a", zIndex: 1 }}>
                      <tr>
                        <th style={{ ...S.th, color: "#94a3b8" }}>Composite Key</th>
                        {keyMappings.map(m => <th key={m.colA} style={{ ...S.th, color: "#f59e0b" }}>{m.colA}</th>)}
                        {compareMappings.map(m => (
                          <>
                            <th key={m.colA + "a"} style={{ ...S.th, color: "#38bdf8", borderLeft: "1px solid #0f2040" }}>{m.colA} (A)</th>
                            <th key={m.colA + "b"} style={{ ...S.th, color: "#818cf8" }}>{m.colB} (B)</th>
                            <th key={m.colA + "d"} style={{ ...S.th, color: "#f59e0b" }}>Diff</th>
                          </>
                        ))}
                        <th style={{ ...S.th, color: "#475569" }}>Status</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredResults.slice(0, 200).map((r, i) => (
                        <tr key={i} style={{ background: i % 2 ? "#04080f" : "#060d1a" }}>
                          <td style={{ ...S.td, color: "#64748b", maxWidth: 200, overflow: "hidden", textOverflow: "ellipsis" }}>{r.key}</td>
                          {keyMappings.map(m => <td key={m.colA} style={{ ...S.td, color: "#64748b" }}>{r.keyVals[m.colA] ?? "—"}</td>)}
                          {compareMappings.map(m => {
                            const d = r.details.find(d => d.colA === m.colA);
                            return (
                              <>
                                <td key={m.colA + "a"} style={{ ...S.td, color: "#38bdf8", borderLeft: "1px solid #0f2040" }}>{d ? d.valA : (r.rowA ? r.rowA[m.colA] ?? "—" : "—")}</td>
                                <td key={m.colA + "b"} style={{ ...S.td, color: "#818cf8" }}>{d ? d.valB : (r.rowB ? r.rowB[m.colB] ?? "—" : "—")}</td>
                                <td key={m.colA + "d"} style={{ ...S.td, color: d?.diff && d.diff !== "0.00" ? "#f59e0b" : "#1e3a5f" }}>{d?.diff || "—"}</td>
                              </>
                            );
                          })}
                          <td style={S.td}><StatusBadge status={r.status} /></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {filteredResults.length === 0 && <div style={{ padding: 32, textAlign: "center", color: "#1e3a5f" }}>No records match filter.</div>}
                  {filteredResults.length > 200 && <div style={{ padding: 8, textAlign: "center", color: "#1e3a5f", fontSize: 10 }}>Showing 200 of {filteredResults.length} — Export Excel for all.</div>}
                </div>
              </div>
            )}

            {/* COMPARISON SHEET — 1 row per file, all columns, all sessions */}
            {activeTab === "sheet" && (() => {
              const { headers: sheetHeaders, rows: sheetRows } = buildCompSheetRows(sessions);
              // Determine which column indices are diff cols (last N cols)
              const cMaps = sessions[sessions.length - 1]?.mappings.filter(m => m.compare && !m.isKey && m.colA && m.colB) || [];
              return (
                <div>
                  {/* Legend */}
                  <div style={{ background: "#060d1a", borderRadius: 9, padding: "10px 14px", marginBottom: 10, border: "1px solid #22c55e22", fontSize: 11, display: "flex", gap: 20, alignItems: "center", flexWrap: "wrap" }}>
                    <span style={{ color: "#475569" }}>
                      📋 <b style={{ color: "#e2e8f0" }}>1 row per file</b> — every comparison appends 2 rows (File A + File B). All file columns included.
                    </span>
                    <span style={{ color: "#22c55e", fontWeight: 700 }}>{sessions.length} comparison{sessions.length !== 1 ? "s" : ""} → {sheetRows.length} rows total</span>
                    <div style={{ flex: 1 }} />
                    <div style={{ display: "flex", gap: 10, fontSize: 10 }}>
                      <span style={{ display: "flex", alignItems: "center", gap: 5 }}><span style={{ width: 8, height: 8, borderRadius: 2, background: "#0ea5e9", display: "inline-block" }} />File A row</span>
                      <span style={{ display: "flex", alignItems: "center", gap: 5 }}><span style={{ width: 8, height: 8, borderRadius: 2, background: "#818cf8", display: "inline-block" }} />File B row</span>
                      <span style={{ display: "flex", alignItems: "center", gap: 5 }}><span style={{ width: 8, height: 8, borderRadius: 2, background: "#ef4444", display: "inline-block" }} />Has Differences</span>
                    </div>
                  </div>

                  <div style={S.card}>
                    <div style={{ overflowX: "auto", maxHeight: 540, overflowY: "auto" }}>
                      {sheetRows.length === 0 ? (
                        <div style={{ padding: 40, textAlign: "center", color: "#1e3a5f" }}>No data yet.</div>
                      ) : (
                        <table style={{ borderCollapse: "collapse", fontSize: 11 }}>
                          <thead style={{ position: "sticky", top: 0, zIndex: 2 }}>
                            <tr>
                              {sheetHeaders.map((h, hi) => {
                                // Color coding: meta cols, [A] cols, [B] cols, Diff cols
                                const color = hi < 6 ? "#94a3b8"
                                  : h.startsWith("[A]") ? "#38bdf8"
                                  : h.startsWith("[B]") ? "#818cf8"
                                  : h.startsWith("Diff:") ? "#f59e0b"
                                  : "#64748b";
                                const bg = h.startsWith("[A]") ? "#04101e" : h.startsWith("[B]") ? "#0d0a1f" : h.startsWith("Diff:") ? "#130f03" : "#04080f";
                                const borderL = (h.startsWith("[A]") && !sheetHeaders[hi - 1]?.startsWith("[A]")) || (h.startsWith("[B]") && !sheetHeaders[hi - 1]?.startsWith("[B]")) || (h.startsWith("Diff:") && !sheetHeaders[hi - 1]?.startsWith("Diff:")) ? "2px solid #1e3a5f" : "none";
                                return (
                                  <th key={hi} style={{ ...S.th, color, background: bg, borderLeft: borderL, minWidth: h.length > 20 ? 140 : 90, maxWidth: 200 }}>
                                    {h}
                                  </th>
                                );
                              })}
                            </tr>
                          </thead>
                          <tbody>
                            {sheetRows.map((row, ri) => {
                              const isFileA = row[2] === "A";
                              const hasDiff = row[4] !== "Matched";
                              const isNewSession = ri === 0 || row[0] !== sheetRows[ri - 2]?.[0];
                              return (
                                <>
                                  {isFileA && isNewSession && ri > 0 && (
                                    <tr key={`sep-${ri}`}>
                                      <td colSpan={sheetHeaders.length} style={{ padding: "4px 14px", background: "#0a1628", borderTop: "2px solid #1e3a5f", borderBottom: "1px solid #0f2040" }}>
                                        <span style={{ fontSize: 10, color: "#334155", fontWeight: 700 }}>▶ Session {row[0]}: {row[1]}</span>
                                      </td>
                                    </tr>
                                  )}
                                  <tr key={ri} style={{ background: hasDiff ? (isFileA ? "#140404" : "#100c03") : (isFileA ? "#04101e" : "#0d0a1f"), borderBottom: !isFileA ? "2px solid #1e3a5f" : "none" }}>
                                    {row.map((cell, ci) => {
                                      const h = sheetHeaders[ci];
                                      const isACol = h?.startsWith("[A]");
                                      const isBCol = h?.startsWith("[B]");
                                      const isDiffCol = h?.startsWith("Diff:");
                                      const isEmpty = cell === "" || cell === undefined;
                                      // On A row, B cols are dimmed; on B row, A cols are dimmed
                                      const dimmed = (isFileA && isBCol) || (!isFileA && isACol);
                                      const diffBad = isDiffCol && cell && cell !== "OK" && cell !== "";
                                      const borderL = (isACol && !sheetHeaders[ci - 1]?.startsWith("[A]")) || (isBCol && !sheetHeaders[ci - 1]?.startsWith("[B]")) || (isDiffCol && !sheetHeaders[ci - 1]?.startsWith("Diff:")) ? "2px solid #1e3a5f" : "none";
                                      let color = "#64748b";
                                      if (ci === 2) color = isFileA ? "#38bdf8" : "#818cf8"; // File col
                                      else if (ci === 3) color = isFileA ? "#38bdf8" : "#818cf8"; // File Name
                                      else if (ci === 4) color = hasDiff ? "#ef4444" : "#22c55e"; // Status
                                      else if (isACol && !dimmed) color = "#38bdf8";
                                      else if (isBCol && !dimmed) color = "#818cf8";
                                      else if (isDiffCol) color = diffBad ? "#f59e0b" : "#22c55e";
                                      else if (dimmed) color = "#1e3a5f";
                                      return (
                                        <td key={ci} style={{ ...S.td, color, borderLeft: borderL, background: diffBad ? "#1a1000" : "transparent", fontWeight: (ci === 2 || ci === 4 || diffBad) ? 700 : 400, whiteSpace: "nowrap", maxWidth: 200, overflow: "hidden", textOverflow: "ellipsis" }}>
                                          {isEmpty && dimmed ? <span style={{ color: "#0f2040" }}>—</span> : (cell ?? "")}
                                        </td>
                                      );
                                    })}
                                  </tr>
                                </>
                              );
                            })}
                          </tbody>
                        </table>
                      )}
                    </div>
                  </div>
                  <div style={{ marginTop: 8, fontSize: 10, color: "#1e3a5f", textAlign: "center" }}>
                    Columns: <b style={{ color: "#38bdf8" }}>[A] prefix</b> = File A values &nbsp;·&nbsp; <b style={{ color: "#818cf8" }}>[B] prefix</b> = File B values &nbsp;·&nbsp; <b style={{ color: "#f59e0b" }}>Diff:</b> = mismatch count per field &nbsp;·&nbsp; Export Excel for full data
                  </div>
                </div>
              );
            })()}
          </div>
        )}
      </div>
    </div>
  );
}
