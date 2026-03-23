import { useState, useRef } from "react";

// ─── Parsers ─────────────────────────────────────────────────────────────────
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
    const vals = line.split("\t"); const row = {};
    headers.forEach((h, i) => { row[h] = (vals[i] || "").trim(); }); return row;
  });
}
function parsePipe(text) {
  const lines = text.trim().split(/\r?\n/);
  if (!lines.length) return [];
  const headers = lines[0].split("|").map(h => h.trim());
  return lines.slice(1).map(line => {
    const vals = line.split("|"); const row = {};
    headers.forEach((h, i) => { row[h] = (vals[i] || "").trim(); }); return row;
  });
}
function autoDetect(text) {
  const f = text.split(/\r?\n/)[0];
  if (f.includes("|")) return parsePipe(text);
  if (f.includes("\t")) return parseTSV(text);
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
          let txt = "";
          for (let i = 1; i <= pdfDoc.numPages; i++) {
            const page = await pdfDoc.getPage(i);
            txt += (await page.getTextContent()).items.map(it => it.str).join(" ") + "\n";
          }
          resolve(autoDetect(txt));
        } catch { resolve([]); }
      };
      reader.readAsArrayBuffer(file);
    } else {
      reader.onload = (e) => {
        const text = e.target.result;
        if (ext === "tsv") resolve(parseTSV(text));
        else resolve(autoDetect(text));
      };
      reader.readAsText(file);
    }
  });
}

// ─── Auto map headers ────────────────────────────────────────────────────────
function autoMapHeaders(headersA, headersB) {
  const norm = s => s.toLowerCase().replace(/[^a-z0-9]/g, "");
  const usedB = new Set();
  const prefKeys = ["personnumber", "balancename", "area1", "area2", "area3"];
  return headersA.map(ha => {
    const nha = norm(ha);
    let best = "", bestScore = 0;
    for (const hb of headersB) {
      if (usedB.has(hb)) continue;
      const nhb = norm(hb);
      const score = nha === nhb ? 100 : (nha.includes(nhb) || nhb.includes(nha)) ? 75 : 0;
      if (score > bestScore) { bestScore = score; best = hb; }
    }
    if (best && bestScore >= 75) usedB.add(best);
    const isKey = prefKeys.some(p => nha.includes(p));
    return { colA: ha, colB: best && bestScore >= 75 ? best : "", isKey, compare: !isKey && !!(best && bestScore >= 75), ignoreCase: false };
  });
}

// ─── Compare engine ──────────────────────────────────────────────────────────
function compareDatasets(dataA, dataB, mappings, tolerance) {
  const tolPct = parseFloat(tolerance) / 100 || 0;
  const keyMaps = mappings.filter(m => m.isKey && m.colA && m.colB);
  const cMaps = mappings.filter(m => m.compare && !m.isKey && m.colA && m.colB);
  const makeKey = (row, fromA) => keyMaps.map(m => ((fromA ? row[m.colA] : row[m.colB]) || "").toString().toLowerCase().trim()).join("||");
  const indexB = {};
  for (const row of dataB) {
    const k = makeKey(row, false);
    if (!indexB[k]) indexB[k] = [];
    indexB[k].push(row);
  }
  const results = []; const matchedB = new Set();
  for (const rowA of dataA) {
    const key = makeKey(rowA, true);
    const keyVals = Object.fromEntries(keyMaps.map(m => [m.colA, rowA[m.colA] ?? ""]));
    const matchesB = indexB[key] || [];
    if (!matchesB.length) { results.push({ key, rowA, rowB: null, status: "Only in A", details: [], keyVals }); continue; }
    const rowB = matchesB[0]; matchedB.add(key);
    const details = cMaps.map(m => {
      const valA = rowA[m.colA] ?? "", valB = rowB[m.colB] ?? "";
      const cmpA = m.ignoreCase ? valA.toLowerCase() : valA;
      const cmpB = m.ignoreCase ? valB.toLowerCase() : valB;
      const numA = parseFloat(valA), numB = parseFloat(valB);
      const isNum = !isNaN(numA) && !isNaN(numB);
      let diff = "", status = "Matched";
      if (isNum) {
        const pct = numA !== 0 ? Math.abs(numA - numB) / Math.abs(numA) : (numB !== 0 ? 1 : 0);
        diff = (numB - numA).toFixed(4);
        if (pct > tolPct) status = "Mismatched";
      } else if (cmpA.trim() !== cmpB.trim()) { diff = `${valA}→${valB}`; status = "Mismatched"; }
      return { colA: m.colA, colB: m.colB, valA, valB, diff, status };
    });
    results.push({ key, rowA, rowB, status: details.some(d => d.status === "Mismatched") ? "Mismatched" : "Matched", details, keyVals });
  }
  for (const row of dataB) {
    const key = makeKey(row, false);
    if (!matchedB.has(key)) { results.push({ key, rowA: null, rowB: row, status: "Only in B", details: [], keyVals: Object.fromEntries(keyMaps.map(m => [m.colA, row[m.colB] ?? ""])) }); matchedB.add(key); }
  }
  const kc = {};
  for (const row of dataA) { const k = makeKey(row, true); kc[k] = (kc[k] || 0) + 1; }
  return { results, duplicatesA: Object.values(kc).filter(c => c > 1).reduce((s, c) => s + c - 1, 0) };
}

// ─── Comment helper ──────────────────────────────────────────────────────────
function makeComment(diff, valA, valB, status) {
  if (status === "Only in A") return "Only in File A";
  if (status === "Only in B") return "Only in File B";
  if (!diff || diff === "0.0000") return "";
  const num = parseFloat(diff);
  if (!isNaN(num)) return Math.abs(num) < 1 ? "Less than $1 difference" : `Difference: ${diff}`;
  return "Value mismatch";
}

// ─── Excel Export — single workbook, 3 sheets matching reference ──────────────
async function exportToExcel(sessions, allMappings) {
  const { utils, writeFile } = await import("https://cdn.sheetjs.com/xlsx-0.20.1/package/xlsx.mjs");

  // Gather column layout from first session
  const s0 = sessions[0];
  const keyMaps0 = s0.mappings.filter(m => m.isKey && m.colA && m.colB);
  const cMaps0 = s0.mappings.filter(m => m.compare && !m.isKey && m.colA && m.colB);
  const firstRowA = s0.results.find(r => r.rowA)?.rowA || {};
  const allColsA = Object.keys(firstRowA);
  const compareColsSet = new Set(cMaps0.map(m => m.colA));
  const sharedCols = allColsA.filter(c => !compareColsSet.has(c));

  // ── SHEET 1: Comparison Results ──
  // Columns: Status | Source | Composite Key | <sharedCols> | Current | Year-to-Date | … (all colA cols)
  // Each record = 2 rows (File A then File B), alternating, all sessions appended
  const compCols = ["Status", "Source", "Composite Key", ...allColsA];
  const compRows = [compCols];

  // Track which col indices are compare cols (for red highlight)
  const compareColIdxSet = new Set();
  allColsA.forEach((c, i) => { if (compareColsSet.has(c)) compareColIdxSet.add(i + 3); }); // +3 for Status,Source,CompositeKey

  for (const s of sessions) {
    const kMaps = s.mappings.filter(m => m.isKey && m.colA && m.colB);
    for (const r of s.results) {
      const compositeKey = kMaps.map(m => r.keyVals[m.colA] ?? "").join("|");
      // Row A
      const rowA = [r.status, s.fileAName, compositeKey, ...allColsA.map(c => r.rowA ? (r.rowA[c] ?? "") : "")];
      // Row B — map colB values back to colA column positions
      const bMap = Object.fromEntries(s.mappings.filter(m => m.colA && m.colB).map(m => [m.colA, m.colB]));
      const rowB = [r.status, s.fileBName, compositeKey, ...allColsA.map(c => { const colB = bMap[c]; return r.rowB && colB ? (r.rowB[colB] ?? "") : ""; })];
      compRows.push(rowA);
      compRows.push(rowB);
    }
  }

  // ── SHEET 2: Difference Mismatch ──
  // Columns: Status | Source | Composite Key | <sharedCols> | Current | Current USOPTE | Difference | SK Comment | rest of cols
  // 1 row per mismatched record (not paired), File B row (USOPTE = source of truth)
  const diffExtraCols = cMaps0.flatMap(m => [m.colA, `${m.colB} (USOPTE)`, "Difference", "SK Comment"]);
  const diffNonCompareCols = allColsA.filter(c => !compareColsSet.has(c));
  const diffCols = ["Status", "Source", "Composite Key", ...diffNonCompareCols, ...diffExtraCols];
  const diffRows = [diffCols];

  for (const s of sessions) {
    const kMaps = s.mappings.filter(m => m.isKey && m.colA && m.colB);
    const sCmaps = s.mappings.filter(m => m.compare && !m.isKey && m.colA && m.colB);
    const bMapFull = Object.fromEntries(s.mappings.filter(m => m.colA && m.colB).map(m => [m.colA, m.colB]));
    for (const r of s.results.filter(r => r.status !== "Matched")) {
      const compositeKey = kMaps.map(m => r.keyVals[m.colA] ?? "").join("|");
      const sharedVals = diffNonCompareCols.map(c => { const colB = bMapFull[c]; return r.rowB && colB ? (r.rowB[colB] ?? "") : (r.rowA ? (r.rowA[c] ?? "") : ""); });
      const compareVals = cMaps0.flatMap(m => {
        const sCm = sCmaps.find(x => x.colA === m.colA);
        const valA = sCm && r.rowB ? (r.rowB[sCm.colB] ?? "") : ""; // USOPTE = File B
        const valB = sCm && r.rowA ? (r.rowA[sCm.colA] ?? "") : ""; // File A = Current
        const d = r.details?.find(d => d.colA === m.colA);
        const diff = d?.diff || "";
        const comment = makeComment(diff, valB, valA, r.status);
        return [valB, valA, diff ? parseFloat(diff) || diff : "", comment];
      });
      diffRows.push([r.status, s.fileBName, compositeKey, ...sharedVals, ...compareVals]);
    }
  }

  // ── SHEET 3: Summary ──
  const summaryData = [
    ["CompareIQ — Comparison Summary"],
    [],
    ["Session #", "File A", "File B", "Total A", "Total B", "Matched", "Mismatched", "Only in A", "Only in B", "Duplicates", "Match Rate"],
    ...sessions.map((s, i) => [i + 1, s.fileAName, s.fileBName, s.totalA, s.totalB, s.matched, s.mismatched, s.onlyA, s.onlyB, s.duplicates, `${((s.matched / (s.matched + s.mismatched || 1)) * 100).toFixed(1)}%`])
  ];

  // ── Build worksheet — 100% manual, zero aoa_to_sheet calls ──────────────
  // aoa_to_sheet (and sheet_add_aoa internally) enumerates ALL worksheet
  // properties at write time → RangeError on 100k+ row sheets.
  // We write each cell directly as {v, t} — no style objects on data rows.
  const buildSheet = (rows, redColSet) => {
    const ws = Object.create(null); // plain null-prototype obj, no inherited props
    if (!rows.length) { ws["!ref"] = "A1"; return ws; }
    const nCols = rows[0].length;
    const nRows = rows.length;
    const colLetters = [];
    // Pre-compute column letter strings to avoid encode_cell overhead in inner loop
    for (let c = 0; c < nCols; c++) {
      let n = c, s = "";
      do { s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26) - 1; } while (n >= 0);
      colLetters.push(s);
    }
    for (let r = 0; r < nRows; r++) {
      const rowArr = rows[r];
      const isHeader = r === 0;
      const rowNum = r + 1; // Excel 1-indexed
      for (let c = 0; c < nCols; c++) {
        const raw = rowArr[c];
        const strRaw = raw == null ? "" : String(raw);
        const numVal = strRaw !== "" && !isNaN(strRaw) ? Number(strRaw) : null;
        const cell = numVal !== null ? { v: numVal, t: "n" } : { v: strRaw, t: "s" };
        if (isHeader) {
          const isRed = redColSet?.has(c);
          cell.s = {
            fill: { patternType: "solid", fgColor: { rgb: isRed ? "C00000" : "1F3864" } },
            font: { bold: true, color: { rgb: "FFFFFF" }, sz: 9, name: "Arial" },
            alignment: { horizontal: "center", vertical: "center" },
          };
        }
        ws[colLetters[c] + rowNum] = cell;
      }
    }
    ws["!ref"] = `A1:${colLetters[nCols - 1]}${nRows}`;
    ws["!cols"] = rows[0].map((h, i) => ({
      wch: redColSet?.has(i) ? 16 : Math.min(Math.max(String(h ?? "").length + 3, 10), 32),
    }));
    ws["!rows"] = [{ hpt: 24 }];
    return ws;
  };

  const diffRedColSet = new Set(
    Array.from({ length: diffExtraCols.length }, (_, i) => i + 3 + diffNonCompareCols.length)
  );

  const wb = utils.book_new();
  utils.book_append_sheet(wb, buildSheet(summaryData, null), "Summary");
  utils.book_append_sheet(wb, buildSheet(compRows, compareColIdxSet), "Comparison Results");
  utils.book_append_sheet(wb, buildSheet(diffRows, diffRedColSet), "Difference Mismatch");
  // Use writeFile with no compression for faster large-file output
  writeFile(wb, "CompareIQ_Results.xlsx", { compression: false });
}

// ─── Styles (LIGHT THEME) ────────────────────────────────────────────────────
const C = {
  bg: "#F7F9FC", surface: "#FFFFFF", border: "#D1D9E6", borderStrong: "#A0B0C8",
  text: "#1A2332", textMid: "#4A5568", textLight: "#718096",
  blue: "#1A56DB", blueLight: "#EBF5FF", blueMid: "#3B82F6",
  purple: "#7C3AED", purpleLight: "#F5F3FF",
  green: "#059669", greenLight: "#ECFDF5",
  red: "#DC2626", redLight: "#FEF2F2",
  amber: "#D97706", amberLight: "#FFFBEB",
  headerBg: "#1A2332", headerText: "#FFFFFF",
};
const S = {
  card: { background: C.surface, borderRadius: 12, border: `1px solid ${C.border}`, overflow: "hidden", boxShadow: "0 1px 4px rgba(0,0,0,0.06)" },
  th: { padding: "10px 13px", textAlign: "left", borderBottom: `1px solid ${C.border}`, whiteSpace: "nowrap", fontSize: 11, fontWeight: 700, color: C.textMid, background: "#F0F4FA" },
  td: { padding: "7px 12px", borderBottom: `1px solid ${C.border}20`, whiteSpace: "nowrap", fontSize: 11, color: C.text },
  input: { background: C.surface, border: `1px solid ${C.border}`, borderRadius: 7, padding: "7px 12px", color: C.text, fontSize: 12, fontFamily: "inherit", outline: "none" },
  btn: (active, color = C.blue) => ({ background: active ? color : C.surface, border: `1px solid ${active ? color : C.border}`, color: active ? "#fff" : C.textMid, borderRadius: 7, padding: "7px 14px", cursor: "pointer", fontSize: 12, fontWeight: 600, fontFamily: "inherit", transition: "all .15s" }),
  btnPrimary: { background: `linear-gradient(135deg,${C.blue},${C.purple})`, color: "#fff", border: "none", borderRadius: 9, padding: "11px 32px", fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" },
  label: { fontSize: 11, fontWeight: 700, color: C.textLight, letterSpacing: "0.05em" },
};

const StatusBadge = ({ status }) => {
  const cfg = { Matched: [C.green, C.greenLight], Mismatched: [C.red, C.redLight], "Only in A": [C.amber, C.amberLight], "Only in B": [C.purple, C.purpleLight] };
  const [c, bg] = cfg[status] || [C.textLight, C.bg];
  return <span style={{ background: bg, color: c, border: `1px solid ${c}33`, borderRadius: 5, padding: "2px 9px", fontSize: 10, fontWeight: 700 }}>{status}</span>;
};

const STEPS = ["Upload", "Map Columns", "Results"];

export default function CompareIQ() {
  const [step, setStep] = useState(0);
  const [fileA, setFileA] = useState(null); const [fileB, setFileB] = useState(null);
  const [dataA, setDataA] = useState([]); const [dataB, setDataB] = useState([]);
  const [headersA, setHeadersA] = useState([]); const [headersB, setHeadersB] = useState([]);
  const [mappings, setMappings] = useState([]);
  const [tolerance, setTolerance] = useState("1");
  const [results, setResults] = useState(null);
  const [loading, setLoading] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [error, setError] = useState("");
  const [activeTab, setActiveTab] = useState("dashboard");
  const [filterStatus, setFilterStatus] = useState("All");
  const [searchKey, setSearchKey] = useState("");
  const [sessions, setSessions] = useState([]);

  const handleExport = async () => {
    setExporting(true);
    setError("");
    try {
      await exportToExcel(sessions, mappings);
    } catch (e) {
      setError(`Export failed: ${e.message}`);
    }
    setExporting(false);
  };

  const handleFile = async (file, which) => {
    setError(""); setLoading(true);
    try {
      const data = await parseFile(file);
      if (!data.length) throw new Error("No data found");
      const headers = Object.keys(data[0]);
      if (which === "A") { setFileA(file); setDataA(data); setHeadersA(headers); }
      else { setFileB(file); setDataB(data); setHeadersB(headers); }
    } catch (e) { setError(`Error: ${e.message}`); }
    setLoading(false);
  };

  const proceedToMap = () => {
    if (!dataA.length || !dataB.length) { setError("Upload both files first."); return; }
    setMappings(autoMapHeaders(headersA, headersB));
    setStep(1);
  };

  const updateMapping = (i, field, val) => setMappings(p => p.map((m, idx) => idx === i ? { ...m, [field]: val } : m));

  // Select/deselect all compare
  const setAllCompare = (val) => setMappings(p => p.map(m => m.isKey ? m : { ...m, compare: val }));
  const allCompareOn = mappings.filter(m => !m.isKey).every(m => m.compare);
  const anyCompareOn = mappings.filter(m => !m.isKey).some(m => m.compare);

  const runComparison = () => {
    if (!mappings.filter(m => m.isKey).length) { setError("Mark at least one KEY column."); return; }
    setLoading(true);
    setTimeout(() => {
      const res = compareDatasets(dataA, dataB, mappings, tolerance);
      setResults(res);
      const matched = res.results.filter(r => r.status === "Matched").length;
      const mismatched = res.results.filter(r => r.status === "Mismatched").length;
      const onlyA = res.results.filter(r => r.status === "Only in A").length;
      const onlyB = res.results.filter(r => r.status === "Only in B").length;
      setSessions(prev => [...prev, { fileAName: fileA?.name || "File A", fileBName: fileB?.name || "File B", results: res.results, totalA: dataA.length, totalB: dataB.length, matched, mismatched, onlyA, onlyB, duplicates: res.duplicatesA, mappings: [...mappings] }]);
      setStep(2); setLoading(false);
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

  // ── Sub-components ───────────────────────────────────────────────────────
  const DropZone = ({ label, file, onFile, color }) => {
    const ref = useRef(); const [drag, setDrag] = useState(false);
    return (
      <div onClick={() => ref.current.click()}
        onDragOver={e => { e.preventDefault(); setDrag(true); }}
        onDragLeave={() => setDrag(false)}
        onDrop={e => { e.preventDefault(); setDrag(false); if (e.dataTransfer.files[0]) onFile(e.dataTransfer.files[0]); }}
        style={{ border: `2px dashed ${drag ? color : file ? C.green : C.border}`, borderRadius: 12, padding: "22px 16px", textAlign: "center", cursor: "pointer", background: drag ? `${color}08` : file ? C.greenLight : C.bg, transition: "all .2s" }}>
        <input ref={ref} type="file" accept=".csv,.xlsx,.xls,.txt,.tsv,.pdf,.tex" style={{ display: "none" }} onChange={e => { if (e.target.files[0]) onFile(e.target.files[0]); }} />
        <div style={{ fontSize: 26, marginBottom: 6 }}>{file ? "✅" : "📂"}</div>
        <div style={{ fontWeight: 700, color: file ? C.green : C.textMid, fontSize: 13 }}>{file ? file.name : label}</div>
        {file ? <div style={{ color: C.textLight, fontSize: 11, marginTop: 3 }}>{(file.size / 1024).toFixed(1)} KB · click to replace</div>
          : <div style={{ color: C.textLight, fontSize: 11, marginTop: 5 }}>CSV · Excel · TXT · TSV · PDF · TeX</div>}
      </div>
    );
  };

  const Radio = ({ checked, onChange, color = C.amber }) => (
    <div onClick={onChange} style={{ width: 17, height: 17, borderRadius: "50%", border: `2px solid ${checked ? color : C.borderStrong}`, background: checked ? color : "#fff", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", margin: "auto", transition: "all .15s", flexShrink: 0 }}>
      {checked && <div style={{ width: 6, height: 6, borderRadius: "50%", background: "#fff" }} />}
    </div>
  );

  const Checkbox = ({ checked, onChange, color = C.blue }) => (
    <div onClick={onChange} style={{ width: 17, height: 17, borderRadius: 4, border: `2px solid ${checked ? color : C.borderStrong}`, background: checked ? color : "#fff", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", margin: "auto", transition: "all .15s", flexShrink: 0 }}>
      {checked && <span style={{ color: "#fff", fontSize: 11, fontWeight: 900, lineHeight: 1 }}>✓</span>}
    </div>
  );

  const ColSelect = ({ value, options, onChange, color }) => (
    <div style={{ position: "relative", flex: 1 }}>
      <select value={value} onChange={e => onChange(e.target.value)}
        style={{ ...S.input, width: "100%", paddingRight: 24, appearance: "none", cursor: "pointer", borderColor: value ? `${color}55` : C.border }}>
        <option value="">— not mapped —</option>
        {options.map(o => <option key={o}>{o}</option>)}
      </select>
      <span style={{ position: "absolute", right: 8, top: "50%", transform: "translateY(-50%)", color: C.textLight, pointerEvents: "none", fontSize: 10 }}>▾</span>
    </div>
  );

  return (
    <div style={{ minHeight: "100vh", background: C.bg, color: C.text, fontFamily: "'Inter','Segoe UI',sans-serif" }}>
      <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet" />

      {/* ── Header ── */}
      <div style={{ background: C.headerBg, padding: "12px 28px", display: "flex", alignItems: "center", justifyContent: "space-between", position: "sticky", top: 0, zIndex: 50, boxShadow: "0 2px 8px rgba(0,0,0,0.15)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ width: 32, height: 32, borderRadius: 8, background: "linear-gradient(135deg,#3B82F6,#7C3AED)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16 }}>⚖️</div>
          <div>
            <div style={{ fontWeight: 800, fontSize: 16, color: "#fff", letterSpacing: 0.3 }}>CompareIQ</div>
            <div style={{ fontSize: 9, color: "#94A3B8", letterSpacing: 1.5 }}>INTELLIGENT DATASET COMPARISON</div>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
          {STEPS.map((s, i) => (
            <div key={s} style={{ display: "flex", alignItems: "center", gap: 3 }}>
              <div onClick={() => step > i && setStep(i)} style={{ display: "flex", alignItems: "center", gap: 6, padding: "5px 12px", borderRadius: 20, background: step === i ? "#3B82F6" : step > i ? "#1E40AF" : "transparent", border: `1px solid ${step === i ? "#3B82F6" : step > i ? "#3B82F6" : "#475569"}`, color: step >= i ? "#fff" : "#94A3B8", fontSize: 11, fontWeight: 600, cursor: step > i ? "pointer" : "default" }}>
                <div style={{ width: 14, height: 14, borderRadius: "50%", background: step > i ? "#22C55E" : step === i ? "#fff" : "#475569", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 8, color: step === i ? "#3B82F6" : "#fff", fontWeight: 900 }}>{step > i ? "✓" : i + 1}</div>
                {s}
              </div>
              {i < STEPS.length - 1 && <div style={{ width: 14, height: 1, background: "#475569" }} />}
            </div>
          ))}
          {sessions.length > 0 && <div style={{ marginLeft: 12, padding: "4px 10px", background: "#14532D", border: "1px solid #22C55E55", borderRadius: 20, fontSize: 10, color: "#22C55E" }}>{sessions.length} session{sessions.length > 1 ? "s" : ""}</div>}
        </div>
      </div>

      <div style={{ maxWidth: 1100, margin: "0 auto", padding: "28px 20px" }}>
        {error && <div style={{ background: C.redLight, border: `1px solid ${C.red}33`, color: C.red, borderRadius: 9, padding: "9px 14px", marginBottom: 14, fontSize: 12, display: "flex", justifyContent: "space-between", alignItems: "center" }}>⚠️ {error}<button onClick={() => setError("")} style={{ background: "none", border: "none", color: C.red, cursor: "pointer", fontSize: 16, lineHeight: 1 }}>×</button></div>}
        {loading && <div style={{ background: C.blueLight, border: `1px solid ${C.blue}33`, borderRadius: 9, padding: "9px 14px", marginBottom: 14, fontSize: 12, color: C.blue }}>⏳ Processing…</div>}
        {exporting && <div style={{ background: C.amberLight, border: `1px solid ${C.amber}33`, borderRadius: 9, padding: "9px 14px", marginBottom: 14, fontSize: 12, color: C.amber, fontWeight: 600 }}>⏳ Building Excel file — this may take a moment for large datasets ({sessions.reduce((a, s) => a + s.results.length, 0).toLocaleString()} records)…</div>}

        {/* ══ STEP 0 — UPLOAD ══ */}
        {step === 0 && (
          <div>
            <div style={{ textAlign: "center", marginBottom: 28 }}>
              <h1 style={{ fontWeight: 800, fontSize: 28, margin: "0 0 6px", color: C.text }}>Upload Your Datasets</h1>
              <p style={{ color: C.textLight, margin: 0, fontSize: 13 }}>CSV · Excel · TXT · TSV · PDF · TeX · Pipe-delimited</p>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 24 }}>
              {[
                { label: "Drop Dataset A — Source / HR / Expected", which: "A", file: fileA, data: dataA, headers: headersA, color: C.blue },
                { label: "Drop Dataset B — Target / Payroll / Actual", which: "B", file: fileB, data: dataB, headers: headersB, color: C.purple }
              ].map(({ label, which, file, data, headers, color }) => (
                <div key={which}>
                  <div style={{ ...S.label, marginBottom: 8, color }}>● DATASET {which}</div>
                  <DropZone label={label} file={file} color={color} onFile={f => handleFile(f, which)} />
                  {data.length > 0 && (
                    <div style={{ marginTop: 8, ...S.card, padding: 12 }}>
                      <div style={{ color, fontSize: 11, fontWeight: 700, marginBottom: 6 }}>{data.length} rows · {headers.length} columns</div>
                      <div style={{ overflowX: "auto" }}>
                        <table style={{ borderCollapse: "collapse", fontSize: 10 }}>
                          <thead><tr>{headers.slice(0, 7).map(h => <th key={h} style={{ padding: "3px 8px", background: "#F0F4FA", color: C.textMid, textAlign: "left", whiteSpace: "nowrap", fontWeight: 600 }}>{h}</th>)}{headers.length > 7 && <th style={{ color: C.textLight, padding: "3px 6px" }}>+{headers.length - 7}</th>}</tr></thead>
                          <tbody>{data.slice(0, 2).map((r, i) => <tr key={i}>{headers.slice(0, 7).map(h => <td key={h} style={{ padding: "3px 8px", color: C.textMid, whiteSpace: "nowrap" }}>{r[h]}</td>)}</tr>)}</tbody>
                        </table>
                      </div>
                    </div>
                  )}
                </div>
              ))}
            </div>
            <div style={{ textAlign: "center" }}>
              <button onClick={proceedToMap} disabled={!dataA.length || !dataB.length} style={{ ...S.btnPrimary, opacity: dataA.length && dataB.length ? 1 : 0.4 }}>
                Continue to Column Mapping →
              </button>
            </div>
          </div>
        )}

        {/* ══ STEP 1 — COLUMN MAPPING ══ */}
        {step === 1 && (
          <div>
            <div style={{ marginBottom: 18 }}>
              <h1 style={{ fontWeight: 800, fontSize: 22, margin: "0 0 4px" }}>Column Mapping Configuration</h1>
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginTop: 8, flexWrap: "wrap" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 6, padding: "4px 12px", background: C.blueLight, border: `1px solid ${C.blue}33`, borderRadius: 20, fontSize: 11 }}>
                  <span style={{ color: C.blue }}>●</span><span style={{ color: C.blue, fontWeight: 600 }}>{fileA?.name}</span>
                </div>
                <span style={{ color: C.textLight, fontSize: 16 }}>⇌</span>
                <div style={{ display: "flex", alignItems: "center", gap: 6, padding: "4px 12px", background: C.purpleLight, border: `1px solid ${C.purple}33`, borderRadius: 20, fontSize: 11 }}>
                  <span style={{ color: C.purple }}>●</span><span style={{ color: C.purple, fontWeight: 600 }}>{fileB?.name}</span>
                </div>
                <div style={{ flex: 1 }} />
                <span style={{ fontSize: 11, color: C.textLight }}>Total: <b style={{ color: C.blue }}>{mappings.length}</b></span>
                <span style={{ fontSize: 11, color: C.textLight }}>Keys: <b style={{ color: C.amber }}>{mappings.filter(m => m.isKey).length}</b></span>
                <span style={{ fontSize: 11, color: C.textLight }}>Compare: <b style={{ color: C.green }}>{mappings.filter(m => m.compare && !m.isKey).length}</b></span>
              </div>
            </div>

            {/* Info banner */}
            <div style={{ background: C.blueLight, border: `1px solid ${C.blue}22`, borderRadius: 9, padding: "10px 14px", marginBottom: 14, fontSize: 12, color: C.textMid, display: "flex", alignItems: "center", gap: 8 }}>
              <span style={{ color: C.blue, fontSize: 16 }}>ℹ</span>
              Map columns, set <b style={{ color: C.amber }}>Key</b> columns for row matching, toggle <b style={{ color: C.green }}>Compare</b> for value comparison.
            </div>

            {/* Tolerance + actions */}
            <div style={{ display: "flex", alignItems: "center", gap: 14, marginBottom: 10, padding: "10px 14px", ...S.card, boxShadow: "none" }}>
              <span style={{ ...S.label }}>⚙ GLOBAL TOLERANCE</span>
              <input type="number" min="0" max="100" step="0.1" value={tolerance} onChange={e => setTolerance(e.target.value)}
                style={{ ...S.input, width: 60, textAlign: "center", borderColor: C.green, color: C.green, fontWeight: 700, fontSize: 14 }} />
              <span style={{ color: C.green, fontWeight: 700 }}>%</span>
              <span style={{ fontSize: 11, color: C.textLight }}>Acceptable numeric variance</span>
              <div style={{ flex: 1 }} />
              <button onClick={() => setMappings(p => [...p, { colA: headersA[0] || "", colB: headersB[0] || "", isKey: false, compare: true, ignoreCase: false }])}
                style={{ ...S.btn(false), fontSize: 12, fontWeight: 700 }}>+ Add Row</button>
            </div>

            {/* Mapping table */}
            <div style={{ ...S.card, marginBottom: 16 }}>
              {/* Table header */}
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 64px 80px 96px 32px", background: "#F0F4FA", padding: "8px 14px", borderBottom: `1px solid ${C.border}` }}>
                {[
                  { label: `DATASET 1 — ${fileA?.name}`, color: C.blue },
                  { label: `DATASET 2 — ${fileB?.name}`, color: C.purple },
                  { label: "KEY 🔑", color: C.amber },
                  { label: "", color: C.green },
                  { label: "IGNORE CASE", color: C.textLight },
                  { label: "", color: "" },
                ].map((h, i) => (
                  <div key={i} style={{ ...S.label, color: h.color, textAlign: i >= 2 ? "center" : "left", display: i === 3 ? "flex" : "block", alignItems: "center", gap: 4, flexDirection: "column" }}>
                    {i === 3 ? (
                      <>
                        <span style={{ color: C.green }}>COMPARE ✓</span>
                        {/* Select All / Deselect All dropdown */}
                        <div style={{ display: "flex", gap: 4, marginTop: 2 }}>
                          <button onClick={() => setAllCompare(true)} style={{ ...S.btn(allCompareOn, C.green), padding: "2px 7px", fontSize: 9, fontWeight: 700 }}>All</button>
                          <button onClick={() => setAllCompare(false)} style={{ ...S.btn(!anyCompareOn, C.red), padding: "2px 7px", fontSize: 9, fontWeight: 700 }}>None</button>
                        </div>
                      </>
                    ) : h.label}
                  </div>
                ))}
              </div>

              {/* Rows */}
              <div style={{ maxHeight: 380, overflowY: "auto" }}>
                {mappings.map((m, i) => (
                  <div key={i} style={{ display: "grid", gridTemplateColumns: "1fr 1fr 64px 80px 96px 32px", alignItems: "center", padding: "6px 14px", borderBottom: `1px solid ${C.border}20`, background: m.isKey ? "#FFFBEB" : i % 2 ? C.surface : C.bg }}>
                    <div style={{ paddingRight: 8 }}><ColSelect value={m.colA} options={headersA} color={C.blue} onChange={v => updateMapping(i, "colA", v)} /></div>
                    <div style={{ paddingRight: 8 }}><ColSelect value={m.colB} options={headersB} color={C.purple} onChange={v => updateMapping(i, "colB", v)} /></div>
                    <div style={{ textAlign: "center" }}><Radio checked={m.isKey} color={C.amber} onChange={() => updateMapping(i, "isKey", !m.isKey)} /></div>
                    <div style={{ textAlign: "center" }}><Checkbox checked={m.compare && !m.isKey} color={C.green} onChange={() => { if (!m.isKey) updateMapping(i, "compare", !m.compare); }} /></div>
                    <div style={{ textAlign: "center" }}><Checkbox checked={m.ignoreCase} color={C.textLight} onChange={() => updateMapping(i, "ignoreCase", !m.ignoreCase)} /></div>
                    <div style={{ textAlign: "center" }}>
                      <button onClick={() => setMappings(p => p.filter((_, j) => j !== i))} style={{ background: "none", border: "none", color: C.textLight, cursor: "pointer", fontSize: 16, lineHeight: 1 }}>×</button>
                    </div>
                  </div>
                ))}
              </div>

              {/* Footer */}
              <div style={{ padding: "8px 14px", background: "#F8FAFC", borderTop: `1px solid ${C.border}`, display: "flex", gap: 16, fontSize: 11 }}>
                <span style={{ color: C.textLight }}>Total: <b style={{ color: C.blue }}>{mappings.length}</b></span>
                <span style={{ color: C.textLight }}>Keys: <b style={{ color: C.amber }}>{mappings.filter(m => m.isKey).length}</b></span>
                <span style={{ color: C.textLight }}>Compare: <b style={{ color: C.green }}>{mappings.filter(m => m.compare && !m.isKey).length}</b></span>
                <span style={{ color: C.textLight }}>Ignore Case: <b>{mappings.filter(m => m.ignoreCase).length}</b></span>
              </div>
            </div>

            <div style={{ display: "flex", gap: 10, justifyContent: "center" }}>
              <button onClick={() => setStep(0)} style={{ ...S.btn(false), padding: "10px 22px" }}>← Back</button>
              <button onClick={runComparison} style={S.btnPrimary}>🚀 Run Comparison</button>
            </div>
          </div>
        )}

        {/* ══ STEP 2 — RESULTS ══ */}
        {step === 2 && stats && (
          <div>
            {/* Summary cards */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(7,1fr)", gap: 8, marginBottom: 16 }}>
              {[
                { label: "Total A", val: stats.totalA, color: C.blue, bg: C.blueLight },
                { label: "Total B", val: stats.totalB, color: C.purple, bg: C.purpleLight },
                { label: "Matched", val: stats.matched, color: C.green, bg: C.greenLight },
                { label: "Mismatched", val: stats.mismatched, color: C.red, bg: C.redLight },
                { label: "Only in A", val: stats.onlyA, color: C.amber, bg: C.amberLight },
                { label: "Only in B", val: stats.onlyB, color: C.purple, bg: C.purpleLight },
                { label: "Duplicates", val: stats.duplicates, color: C.textMid, bg: "#F1F5F9" },
              ].map(s => (
                <div key={s.label} style={{ background: s.bg, border: `1px solid ${s.color}30`, borderTop: `3px solid ${s.color}`, borderRadius: 10, padding: "12px 8px", textAlign: "center", boxShadow: "0 1px 3px rgba(0,0,0,0.04)" }}>
                  <div style={{ fontSize: 20, fontWeight: 800, color: s.color }}>{s.val}</div>
                  <div style={{ fontSize: 9, color: C.textLight, marginTop: 2, fontWeight: 600, letterSpacing: "0.05em" }}>{s.label.toUpperCase()}</div>
                </div>
              ))}
            </div>

            {/* Match rate bar */}
            <div style={{ ...S.card, padding: "12px 16px", marginBottom: 14, boxShadow: "none" }}>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11, marginBottom: 7 }}>
                <span style={{ color: C.textMid, fontWeight: 600 }}>Match Rate — {fileA?.name} vs {fileB?.name}</span>
                <span style={{ color: C.green, fontWeight: 700 }}>{((stats.matched / (stats.matched + stats.mismatched || 1)) * 100).toFixed(1)}%</span>
              </div>
              <div style={{ height: 7, background: "#E2E8F0", borderRadius: 99, display: "flex", overflow: "hidden" }}>
                {[{ w: stats.matched, c: C.green }, { w: stats.mismatched, c: C.red }, { w: stats.onlyA, c: C.amber }, { w: stats.onlyB, c: C.purple }].map((s, i) => {
                  const t = stats.matched + stats.mismatched + stats.onlyA + stats.onlyB || 1;
                  return <div key={i} style={{ width: `${(s.w / t) * 100}%`, background: s.c }} />;
                })}
              </div>
            </div>

            {sessions.length > 1 && (
              <div style={{ background: C.greenLight, border: `1px solid ${C.green}33`, borderRadius: 9, padding: "8px 14px", marginBottom: 12, fontSize: 11, color: C.green, fontWeight: 600 }}>
                📚 {sessions.length} sessions accumulated — Export Excel will include all sessions in one workbook
              </div>
            )}

            {/* Tab bar */}
            <div style={{ display: "flex", gap: 5, marginBottom: 14, alignItems: "center", flexWrap: "wrap" }}>
              {[["dashboard", "📊 Dashboard"], ["detail", "🔍 Detail Report"], ["sheet", "📋 Comparison Sheet"]].map(([t, label]) => (
                <button key={t} onClick={() => setActiveTab(t)} style={S.btn(activeTab === t)}>{label}</button>
              ))}
              <div style={{ flex: 1 }} />
              <button onClick={handleExport} disabled={exporting}
                style={{ ...S.btn(false), background: exporting ? C.textLight : C.green, color: "#fff", border: "none", fontWeight: 700, padding: "7px 16px", opacity: exporting ? 0.7 : 1, cursor: exporting ? "wait" : "pointer" }}>
                {exporting ? "⏳ Exporting…" : `⬇ Export Excel (${sessions.length} session${sessions.length > 1 ? "s" : ""})`}
              </button>
              <button onClick={() => { setStep(0); setResults(null); }}
                style={{ ...S.btn(false), color: C.blue, borderColor: C.blue, fontWeight: 700 }}>+ New Comparison</button>
              <button onClick={() => setStep(1)} style={S.btn(false)}>← Remap</button>
            </div>

            {/* DASHBOARD */}
            {activeTab === "dashboard" && (
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
                <div style={{ ...S.card, padding: 20 }}>
                  <div style={{ ...S.label, marginBottom: 14 }}>CLASSIFICATION BREAKDOWN</div>
                  <div style={{ display: "flex", alignItems: "center", gap: 20 }}>
                    <svg width="110" height="110" viewBox="0 0 110 110">
                      {(() => {
                        const segs = [{ val: stats.matched, c: C.green }, { val: stats.mismatched, c: C.red }, { val: stats.onlyA, c: C.amber }, { val: stats.onlyB, c: C.purple }];
                        const total = segs.reduce((s, x) => s + x.val, 0) || 1;
                        const r = 42, cx = 55, cy = 55; let angle = -Math.PI / 2;
                        return segs.map((s, i) => {
                          if (!s.val) return null;
                          const sweep = (s.val / total) * 2 * Math.PI;
                          const x1 = cx + r * Math.cos(angle), y1 = cy + r * Math.sin(angle);
                          angle += sweep;
                          return <path key={i} d={`M${cx},${cy} L${x1},${y1} A${r},${r} 0 ${sweep > Math.PI ? 1 : 0} 1 ${cx + r * Math.cos(angle)},${cy + r * Math.sin(angle)} Z`} fill={s.c} opacity={0.9} />;
                        });
                      })()}
                      <circle cx="55" cy="55" r="26" fill="white" />
                      <text x="55" y="59" textAnchor="middle" fill={C.text} fontSize="11" fontWeight="bold" fontFamily="Inter">{((stats.matched / (stats.matched + stats.mismatched || 1)) * 100).toFixed(0)}%</text>
                    </svg>
                    <div style={{ flex: 1 }}>
                      {[["Matched", stats.matched, C.green], ["Mismatched", stats.mismatched, C.red], ["Only A", stats.onlyA, C.amber], ["Only B", stats.onlyB, C.purple]].map(([l, v, c]) => (
                        <div key={l} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                          <div style={{ width: 8, height: 8, borderRadius: 2, background: c, flexShrink: 0 }} />
                          <span style={{ flex: 1, fontSize: 12, color: C.textMid }}>{l}</span>
                          <span style={{ fontWeight: 700, color: c }}>{v}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
                <div style={{ ...S.card, padding: 20 }}>
                  <div style={{ ...S.label, marginBottom: 14 }}>MISMATCHES BY FIELD</div>
                  {compareMappings.length === 0 && <div style={{ color: C.textLight, fontSize: 12 }}>No compare columns configured</div>}
                  {compareMappings.slice(0, 8).map(m => {
                    const total = results.results.filter(r => r.rowA && r.rowB).length || 1;
                    const mis = results.results.filter(r => r.details.some(d => d.colA === m.colA && d.status === "Mismatched")).length;
                    const pct = (mis / total) * 100;
                    return (
                      <div key={m.colA} style={{ marginBottom: 11 }}>
                        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11, marginBottom: 4 }}>
                          <span style={{ color: C.textMid }}>{m.colA}</span>
                          <span style={{ color: pct > 10 ? C.red : C.green, fontWeight: 700 }}>{mis} ({pct.toFixed(1)}%)</span>
                        </div>
                        <div style={{ height: 5, background: "#E2E8F0", borderRadius: 99 }}>
                          <div style={{ width: `${pct}%`, height: "100%", background: pct > 10 ? C.red : C.green, borderRadius: 99 }} />
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
                <div style={{ padding: "11px 14px", borderBottom: `1px solid ${C.border}`, display: "flex", gap: 7, alignItems: "center", flexWrap: "wrap", background: "#F8FAFC" }}>
                  <input placeholder="Search composite key..." value={searchKey} onChange={e => setSearchKey(e.target.value)} style={{ ...S.input, flex: 1, minWidth: 130 }} />
                  {["All", "Matched", "Mismatched", "Only in A", "Only in B"].map(s => (
                    <button key={s} onClick={() => setFilterStatus(s)} style={S.btn(filterStatus === s)}>{s}</button>
                  ))}
                  <span style={{ fontSize: 11, color: C.textLight }}>{filteredResults.length} rows</span>
                </div>
                <div style={{ overflowX: "auto", maxHeight: 420, overflowY: "auto" }}>
                  <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 11 }}>
                    <thead style={{ position: "sticky", top: 0, zIndex: 1 }}>
                      <tr>
                        <th style={S.th}>Composite Key</th>
                        {keyMappings.map(m => <th key={m.colA} style={{ ...S.th, color: C.amber }}>{m.colA}</th>)}
                        {compareMappings.map(m => (
                          <>
                            <th key={m.colA + "a"} style={{ ...S.th, color: C.blue, borderLeft: `2px solid ${C.border}` }}>{m.colA}</th>
                            <th key={m.colA + "b"} style={{ ...S.th, color: C.purple }}>{m.colB} (B)</th>
                            <th key={m.colA + "d"} style={{ ...S.th, color: C.amber }}>Diff</th>
                            <th key={m.colA + "c"} style={{ ...S.th, color: C.textLight }}>Comment</th>
                          </>
                        ))}
                        <th style={S.th}>Status</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredResults.slice(0, 200).map((r, i) => (
                        <tr key={i} style={{ background: i % 2 ? C.bg : C.surface }}>
                          <td style={{ ...S.td, color: C.textMid, maxWidth: 200, overflow: "hidden", textOverflow: "ellipsis" }}>{r.key}</td>
                          {keyMappings.map(m => <td key={m.colA} style={{ ...S.td, color: C.textMid }}>{r.keyVals[m.colA] ?? "—"}</td>)}
                          {compareMappings.map(m => {
                            const d = r.details.find(d => d.colA === m.colA);
                            const isMis = d?.status === "Mismatched";
                            const diff = d?.diff || "";
                            const comment = makeComment(diff, d?.valA || "", d?.valB || "", r.status);
                            return (
                              <>
                                <td key={m.colA + "a"} style={{ ...S.td, color: C.blue, borderLeft: `2px solid ${C.border}`, background: isMis ? C.redLight : "transparent" }}>{d ? d.valA : (r.rowA ? r.rowA[m.colA] ?? "—" : "—")}</td>
                                <td key={m.colA + "b"} style={{ ...S.td, color: C.purple, background: isMis ? C.redLight : "transparent" }}>{d ? d.valB : (r.rowB ? r.rowB[m.colB] ?? "—" : "—")}</td>
                                <td key={m.colA + "d"} style={{ ...S.td, color: isMis ? C.red : C.textLight, fontWeight: isMis ? 700 : 400 }}>{diff || "—"}</td>
                                <td key={m.colA + "c"} style={{ ...S.td, color: C.amber, fontStyle: "italic", fontSize: 10 }}>{comment}</td>
                              </>
                            );
                          })}
                          <td style={S.td}><StatusBadge status={r.status} /></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {filteredResults.length === 0 && <div style={{ padding: 32, textAlign: "center", color: C.textLight }}>No records match filter.</div>}
                  {filteredResults.length > 200 && <div style={{ padding: 8, textAlign: "center", color: C.textLight, fontSize: 10 }}>Showing 200 of {filteredResults.length} — Export Excel for all.</div>}
                </div>
              </div>
            )}

            {/* COMPARISON SHEET */}
            {activeTab === "sheet" && (() => {
              // Build the interleaved view
              const s0 = sessions[sessions.length - 1];
              if (!s0) return null;
              const cMaps0 = s0.mappings.filter(m => m.compare && !m.isKey && m.colA && m.colB);
              const kMaps0 = s0.mappings.filter(m => m.isKey && m.colA && m.colB);
              const firstRowA = s0.results.find(r => r.rowA)?.rowA || {};
              const allColsA = Object.keys(firstRowA);
              const compareColsSet = new Set(cMaps0.map(m => m.colA));
              const sharedCols = allColsA.filter(c => !compareColsSet.has(c));

              return (
                <div>
                  <div style={{ background: C.blueLight, border: `1px solid ${C.blue}22`, borderRadius: 9, padding: "10px 14px", marginBottom: 10, fontSize: 11, display: "flex", gap: 16, alignItems: "center", flexWrap: "wrap" }}>
                    <span style={{ color: C.textMid }}>📋 <b style={{ color: C.text }}>1 row per file per record</b> — shared cols shown once per row, <b style={{ color: C.red }}>compare cols highlighted in red</b> with Current | Current (USOPTE) | Difference | Comment</span>
                    <span style={{ color: C.green, fontWeight: 700, marginLeft: "auto" }}>{sessions.reduce((a, s) => a + s.results.length * 2, 0)} total rows across {sessions.length} session{sessions.length !== 1 ? "s" : ""}</span>
                  </div>

                  <div style={{ ...S.card }}>
                    <div style={{ overflowX: "auto", maxHeight: 520, overflowY: "auto" }}>
                      <table style={{ borderCollapse: "collapse", fontSize: 10 }}>
                        <thead style={{ position: "sticky", top: 0, zIndex: 2 }}>
                          <tr>
                            <th style={{ ...S.th, minWidth: 80, background: C.headerBg, color: "#fff" }}>Status</th>
                            <th style={{ ...S.th, minWidth: 100, background: C.headerBg, color: "#fff" }}>Source</th>
                            <th style={{ ...S.th, minWidth: 160, background: C.headerBg, color: "#fff" }}>Composite Key</th>
                            {sharedCols.map(c => <th key={c} style={{ ...S.th, background: C.headerBg, color: "#CBD5E1", minWidth: 90 }}>{c}</th>)}
                            {cMaps0.map(m => (
                              <>
                                <th key={m.colA} style={{ ...S.th, background: "#7F1D1D", color: "#FCA5A5", borderLeft: "2px solid #991B1B", minWidth: 100 }}>{m.colA}</th>
                                <th key={m.colB + "u"} style={{ ...S.th, background: "#7F1D1D", color: "#FCA5A5", minWidth: 120 }}>{m.colB} (USOPTE)</th>
                                <th key={m.colA + "d"} style={{ ...S.th, background: "#7F1D1D", color: "#FCA5A5", minWidth: 80 }}>Difference</th>
                                <th key={m.colA + "c"} style={{ ...S.th, background: "#7F1D1D", color: "#FCA5A5", borderRight: "2px solid #991B1B", minWidth: 130 }}>Comment</th>
                              </>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {sessions.flatMap((s, si) => {
                            const kMaps = s.mappings.filter(m => m.isKey && m.colA && m.colB);
                            const sCmaps = s.mappings.filter(m => m.compare && !m.isKey && m.colA && m.colB);
                            const bMap = Object.fromEntries(s.mappings.filter(m => m.colA && m.colB).map(m => [m.colA, m.colB]));

                            return s.results.slice(0, 100).flatMap((r, ri) => {
                              const key = kMaps.map(m => r.keyVals[m.colA] ?? "").join("|");
                              const isMismatch = r.status !== "Matched";
                              const pairBg = isMismatch ? "#FFF5F5" : ri % 2 === 0 ? C.surface : C.bg;
                              const pairBgB = isMismatch ? "#FFF0F0" : ri % 2 === 0 ? "#F5F8FF" : "#EFF4FF";

                              const sharedValsA = sharedCols.map(c => r.rowA ? (r.rowA[c] ?? "") : "");
                              const sharedValsB = sharedCols.map(c => { const cb = bMap[c]; return r.rowB && cb ? (r.rowB[cb] ?? "") : ""; });

                              const cValsA = cMaps0.flatMap(m => {
                                const sc = sCmaps.find(x => x.colA === m.colA);
                                const valA = sc && r.rowA ? (r.rowA[sc.colA] ?? "") : "";
                                const d = r.details?.find(d => d.colA === m.colA);
                                const diff = d?.diff || "";
                                const comment = makeComment(diff, d?.valA || "", d?.valB || "", r.status);
                                const isMis = d?.status === "Mismatched";
                                return [
                                  <td key={m.colA} style={{ ...S.td, color: isMis ? C.red : C.blue, background: isMis ? "#FFDDDD" : "#FFF8F8", borderLeft: "2px solid #FECACA", fontWeight: isMis ? 700 : 400 }}>{valA || "—"}</td>,
                                  <td key={m.colB + "u"} style={{ ...S.td, color: "#888", background: "#FFF8F8" }}>—</td>,
                                  <td key={m.colA + "d"} style={{ ...S.td, color: isMis ? C.red : C.textLight, fontWeight: isMis ? 700 : 400, background: isMis ? "#FFDDDD" : "#FFF8F8" }}>{isMis ? diff : "—"}</td>,
                                  <td key={m.colA + "c"} style={{ ...S.td, color: C.amber, fontStyle: "italic", background: "#FFF8F8", borderRight: "2px solid #FECACA" }}>{isMis ? comment : ""}</td>,
                                ];
                                const diff2 = d?.diff || ""; const isMis2 = d?.status === "Mismatched";
                              });

                              const cValsB = cMaps0.flatMap(m => {
                                const sc = sCmaps.find(x => x.colA === m.colA);
                                const valB = sc && r.rowB ? (r.rowB[sc.colB] ?? "") : "";
                                const d = r.details?.find(d => d.colA === m.colA);
                                const isMis = d?.status === "Mismatched";
                                return [
                                  <td key={m.colA} style={{ ...S.td, color: "#888", background: "#FFF8F8", borderLeft: "2px solid #FECACA" }}>—</td>,
                                  <td key={m.colB + "u"} style={{ ...S.td, color: isMis ? C.purple : C.purple, background: isMis ? "#FFE8FF" : "#FFF8F8", fontWeight: isMis ? 700 : 400 }}>{valB || "—"}</td>,
                                  <td key={m.colA + "d"} style={{ ...S.td, background: "#FFF8F8" }}>—</td>,
                                  <td key={m.colA + "c"} style={{ ...S.td, background: "#FFF8F8", borderRight: "2px solid #FECACA" }}>—</td>,
                                ];
                              });

                              return [
                                <tr key={`${si}-${ri}-a`} style={{ background: pairBg }}>
                                  <td style={{ ...S.td }}><StatusBadge status={r.status} /></td>
                                  <td style={{ ...S.td, color: C.blue, fontWeight: 600, fontSize: 10 }}>{s.fileAName}</td>
                                  <td style={{ ...S.td, color: C.textMid, maxWidth: 180, overflow: "hidden", textOverflow: "ellipsis" }}>{key}</td>
                                  {sharedValsA.map((v, vi) => <td key={vi} style={{ ...S.td, color: C.textMid }}>{v}</td>)}
                                  {cValsA}
                                </tr>,
                                <tr key={`${si}-${ri}-b`} style={{ background: pairBgB, borderBottom: `2px solid ${C.border}` }}>
                                  <td style={{ ...S.td, color: C.textLight }}></td>
                                  <td style={{ ...S.td, color: C.purple, fontWeight: 600, fontSize: 10 }}>{s.fileBName}</td>
                                  <td style={{ ...S.td, color: C.textLight }}>{key}</td>
                                  {sharedValsB.map((v, vi) => <td key={vi} style={{ ...S.td, color: C.textLight }}>{v}</td>)}
                                  {cValsB}
                                </tr>
                              ];
                            });
                          })}
                        </tbody>
                      </table>
                      {sessions.reduce((a, s) => a + s.results.length, 0) > 100 && (
                        <div style={{ padding: 10, textAlign: "center", color: C.textLight, fontSize: 10, borderTop: `1px solid ${C.border}` }}>Showing first 100 records per session — Export Excel for complete data</div>
                      )}
                    </div>
                  </div>
                  <div style={{ marginTop: 8, fontSize: 10, color: C.textLight, textAlign: "center" }}>
                    <b style={{ color: C.red }}>Red columns</b> = comparison fields (Current, Current USOPTE, Difference, Comment) &nbsp;·&nbsp; Thick line separates each record pair &nbsp;·&nbsp; Export Excel for full formatted workbook
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
