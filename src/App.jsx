import { useState, useRef, Fragment } from "react";

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
      } else if (cmpA.trim() !== cmpB.trim()) { diff = `${valA}\u2192${valB}`; status = "Mismatched"; }
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
  if (!isNaN(num)) {
    if (Math.abs(num) < 1) return "Less than $1 difference";
    return "More $ Difference";
  }
  return "Value mismatch";
}

// ─── Raw XLSX writer ─────────────────────────────────────────────────────────
async function exportToExcel(sessions) {
  const fflate = await import("https://cdn.jsdelivr.net/npm/fflate@0.8.2/esm/browser.js");
  const enc = new TextEncoder();
  const s0 = sessions[0];
  const cMaps0 = s0.mappings.filter(m => m.compare && !m.isKey && m.colA && m.colB);
  const firstRowA = s0.results.find(r => r.rowA)?.rowA || {};
  const allColsA = Object.keys(firstRowA);
  const compareColsSet = new Set(cMaps0.map(m => m.colA));
  const diffNonCompareCols = allColsA.filter(c => !compareColsSet.has(c));
  const esc = s => String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
  const colName = n => { let s = ""; for (; n >= 0; n = Math.floor(n / 26) - 1) s = String.fromCharCode(65 + n % 26) + s; return s; };
  const rowToBytes = (row, ri, styleIdx) => {
    let s = `<row r="${ri + 1}">`;
    for (let ci = 0; ci < row.length; ci++) {
      const val = row[ci]; const addr = colName(ci) + (ri + 1); const si = styleIdx || 0;
      const isNum = val !== "" && val !== null && val !== undefined && !isNaN(Number(val)) && String(val).trim() !== "" && typeof val !== "boolean";
      if (isNum) s += `<c r="${addr}" s="${si}" t="n"><v>${Number(val)}</v></c>`;
      else s += `<c r="${addr}" s="${si}" t="inlineStr"><is><t>${esc(val)}</t></is></c>`;
    }
    s += "</row>"; return enc.encode(s);
  };
  const concatParts = (parts) => {
    let total = 0; for (const p of parts) total += p.length;
    const out = new Uint8Array(total); let pos = 0;
    for (const p of parts) { out.set(p, pos); pos += p.length; } return out;
  };
  const buildSheetBytes = (styledRows) => {
    const parts = [enc.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>')];
    for (let ri = 0; ri < styledRows.length; ri++) parts.push(rowToBytes(styledRows[ri].cells, ri, styledRows[ri].style));
    parts.push(enc.encode("</sheetData></worksheet>")); return concatParts(parts);
  };
  const nonMatchedResults = sessions.flatMap(s => s.results.filter(r => r.status !== "Matched").map(r => ({ r, s })));
  const compCols = ["Status", "Source", "Composite Key", ...allColsA];
  const compStyledRows = [{ cells: compCols, style: 1 }];
  for (const { r, s } of nonMatchedResults) {
    const kMaps = s.mappings.filter(m => m.isKey && m.colA && m.colB);
    const bMap = Object.fromEntries(s.mappings.filter(m => m.colA && m.colB).map(m => [m.colA, m.colB]));
    const key = kMaps.map(m => r.keyVals[m.colA] ?? "").join("|");
    const st = r.status === "Mismatched" ? 2 : r.status === "Only in A" ? 3 : r.status === "Only in B" ? 4 : 0;
    compStyledRows.push({ cells: [r.status, s.fileAName, key, ...allColsA.map(c => r.rowA ? (r.rowA[c] ?? "") : "")], style: st });
    compStyledRows.push({ cells: [r.status, s.fileBName, key, ...allColsA.map(c => { const cb = bMap[c]; return r.rowB && cb ? (r.rowB[cb] ?? "") : ""; })], style: st });
  }
  const diffExtraCols = cMaps0.flatMap(m => [m.colA, `${m.colA} (USOPTE)`, "Difference", "SK Comment"]);
  const diffCols = ["Status", "Source", "Composite Key", ...diffNonCompareCols, ...diffExtraCols];
  const diffStyledRows = [{ cells: diffCols, style: 1 }];
  for (const { r, s } of nonMatchedResults) {
    const kMaps = s.mappings.filter(m => m.isKey && m.colA && m.colB);
    const sCmaps = s.mappings.filter(m => m.compare && !m.isKey && m.colA && m.colB);
    const bMapFull = Object.fromEntries(s.mappings.filter(m => m.colA && m.colB).map(m => [m.colA, m.colB]));
    const key = kMaps.map(m => r.keyVals[m.colA] ?? "").join("|");
    const st = r.status === "Mismatched" ? 2 : r.status === "Only in A" ? 3 : r.status === "Only in B" ? 4 : 0;
    const sharedValsA = diffNonCompareCols.map(c => { if (r.rowA) return r.rowA[c] ?? ""; const cb = bMapFull[c]; return r.rowB && cb ? (r.rowB[cb] ?? "") : ""; });
    const compareValsRowA = cMaps0.flatMap(m => {
      const sc = sCmaps.find(x => x.colA === m.colA); const valA = sc && r.rowA ? (r.rowA[sc.colA] ?? "") : "";
      const valB = sc && r.rowB ? (r.rowB[sc.colB] ?? "") : ""; const d = r.details?.find(dd => dd.colA === m.colA);
      const diff = d?.diff || ""; const comment = makeComment(diff, valA, valB, r.status);
      return [valA, "", diff ? (parseFloat(diff) || diff) : "", comment];
    });
    diffStyledRows.push({ cells: [r.status, s.fileAName, key, ...sharedValsA, ...compareValsRowA], style: st });
    const sharedValsB = diffNonCompareCols.map(c => { const cb = bMapFull[c]; return r.rowB && cb ? (r.rowB[cb] ?? "") : ""; });
    const compareValsRowB = cMaps0.flatMap(m => { const sc = sCmaps.find(x => x.colA === m.colA); const valB = sc && r.rowB ? (r.rowB[sc.colB] ?? "") : ""; return ["", valB, "", ""]; });
    diffStyledRows.push({ cells: [r.status, s.fileBName, key, ...sharedValsB, ...compareValsRowB], style: st });
  }
  const summaryStyledRows = [
    { cells: ["CompareIQ \u2014 Comparison Summary"], style: 1 }, { cells: [], style: 0 },
    { cells: ["Session #", "File A", "File B", "Total A", "Total B", "Matched", "Mismatched", "Only in A", "Only in B", "Duplicates", "Match Rate"], style: 1 },
    ...sessions.map((s, i) => ({ cells: [i + 1, s.fileAName, s.fileBName, s.totalA, s.totalB, s.matched, s.mismatched, s.onlyA, s.onlyB, s.duplicates, `${((s.matched / (s.matched + s.mismatched || 1)) * 100).toFixed(1)}%`], style: 0 }))
  ];
  const toBytes = s => enc.encode(s);
  const sheets = [{ name: "Summary", data: summaryStyledRows }, { name: "Comparison Results", data: compStyledRows }, { name: "Difference Mismatch", data: diffStyledRows }];
  const stylesXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="7"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1A2332"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFEF2F2"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFBEB"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF5F3FF"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFECFDF5"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/><xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1"/><xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1" applyBorder="1"/><xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0" applyFill="1" applyBorder="1"/></cellXfs></styleSheet>`;
  const workbookXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets.map((s, i) => `<sheet name="${esc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("")}</sheets></workbook>`;
  const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((s, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("")}<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((s, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}</Types>`;
  const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
  const zipFiles = { "[Content_Types].xml": toBytes(contentTypes), "_rels/.rels": toBytes(rootRels), "xl/workbook.xml": toBytes(workbookXML), "xl/_rels/workbook.xml.rels": toBytes(workbookRels), "xl/styles.xml": toBytes(stylesXML) };
  for (let i = 0; i < sheets.length; i++) zipFiles[`xl/worksheets/sheet${i + 1}.xml`] = buildSheetBytes(sheets[i].data);
  const zipped = fflate.zipSync(zipFiles, { level: 0 });
  const blob = new Blob([zipped], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob); const a = document.createElement("a");
  a.href = url; a.download = "CompareIQ_Results.xlsx"; a.click();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

// ─── Styles ──────────────────────────────────────────────────────────────────
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
  const [fileA, setFileA] = useState(null);
  const [fileB, setFileB] = useState(null);
  const [dataA, setDataA] = useState([]);
  const [dataB, setDataB] = useState([]);
  const [headersA, setHeadersA] = useState([]);
  const [headersB, setHeadersB] = useState([]);
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
    setExporting(true); setError("");
    try { await exportToExcel(sessions); } catch (e) { setError(`Export failed: ${e.message}`); }
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
    setMappings(autoMapHeaders(headersA, headersB)); setStep(1);
  };

  const updateMapping = (i, field, val) => setMappings(p => p.map((m, idx) => idx === i ? { ...m, [field]: val } : m));
  const setAllCompare = (val) => setMappings(p => p.map(m => m.isKey ? m : { ...m, compare: val }));
  const allCompareOn = mappings.filter(m => !m.isKey).every(m => m.compare);

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

  // ─── DropZone component ────────────────────────────────────────────────────
  const DropZone = ({ label, file, onFile, color }) => {
    const ref = useRef();
    const [drag, setDrag] = useState(false);
    return (
      <div
        onClick={() => ref.current.click()}
        onDragOver={e => { e.preventDefault(); setDrag(true); }}
        onDragLeave={() => setDrag(false)}
        onDrop={e => { e.preventDefault(); setDrag(false); if (e.dataTransfer.files[0]) onFile(e.dataTransfer.files[0]); }}
        style={{
          border: `2px dashed ${drag ? color : file ? C.green : C.border}`,
          borderRadius: 12, padding: "22px 16px", textAlign: "center",
          cursor: "pointer", background: drag ? `${color}08` : C.surface,
          transition: "all .2s"
        }}
      >
        <input ref={ref} type="file" accept=".csv,.tsv,.xlsx,.xls,.txt,.pdf" style={{ display: "none" }}
          onChange={e => { if (e.target.files[0]) onFile(e.target.files[0]); }} />
        <div style={{ fontSize: 28, marginBottom: 6 }}>{file ? "\u2705" : "\u{1F4C1}"}</div>
        <div style={{ fontSize: 13, fontWeight: 700, color: file ? C.green : C.text, marginBottom: 4 }}>
          {file ? file.name : label}
        </div>
        <div style={{ fontSize: 11, color: C.textLight }}>
          {file ? `${file.name} loaded` : "Drop file or click to browse"}
        </div>
      </div>
    );
  };

  // ─── New Comparison reset ──────────────────────────────────────────────────
  const resetForNew = () => {
    setStep(0); setFileA(null); setFileB(null); setDataA([]); setDataB([]);
    setHeadersA([]); setHeadersB([]); setMappings([]); setResults(null);
    setError(""); setActiveTab("dashboard"); setFilterStatus("All"); setSearchKey("");
  };

  // ─── Render ────────────────────────────────────────────────────────────────
  return (
    <div style={{ minHeight: "100vh", background: C.bg, fontFamily: "-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif", color: C.text }}>
      {/* Header */}
      <div style={{ background: `linear-gradient(135deg,${C.headerBg},#2D3748)`, padding: "18px 28px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <span style={{ fontSize: 22, fontWeight: 800, color: "#fff" }}>CompareIQ</span>
          <span style={{ fontSize: 11, color: "#94A3B8", background: "#334155", borderRadius: 5, padding: "2px 8px" }}>v2.0</span>
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          {STEPS.map((s, i) => (
            <Fragment key={s}>
              {i > 0 && <span style={{ color: "#475569", fontSize: 11 }}>{"\u203A"}</span>}
              <span style={{ fontSize: 11, fontWeight: step === i ? 700 : 400, color: step === i ? "#fff" : "#94A3B8", background: step === i ? C.blue : "transparent", borderRadius: 5, padding: "3px 10px" }}>{s}</span>
            </Fragment>
          ))}
        </div>
      </div>

      <div style={{ maxWidth: 1200, margin: "0 auto", padding: "24px 20px" }}>
        {error && (
          <div style={{ background: C.redLight, border: `1px solid ${C.red}33`, borderRadius: 8, padding: "10px 16px", marginBottom: 16, color: C.red, fontSize: 12, fontWeight: 600 }}>
            {error}
          </div>
        )}

        {loading && (
          <div style={{ textAlign: "center", padding: 40, color: C.textMid, fontSize: 14 }}>Processing...</div>
        )}

        {/* ─── STEP 0: Upload ──────────────────────────────────────────── */}
        {step === 0 && !loading && (
          <div>
            <h2 style={{ fontSize: 20, fontWeight: 700, marginBottom: 16 }}>Upload Files</h2>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20, marginBottom: 24 }}>
              <div>
                <div style={{ ...S.label, marginBottom: 8 }}>FILE A (Source)</div>
                <DropZone label="Upload File A" file={fileA} onFile={f => handleFile(f, "A")} color={C.blue} />
                {dataA.length > 0 && <div style={{ fontSize: 11, color: C.green, marginTop: 6 }}>{dataA.length} rows, {headersA.length} columns</div>}
              </div>
              <div>
                <div style={{ ...S.label, marginBottom: 8 }}>FILE B (Compare)</div>
                <DropZone label="Upload File B" file={fileB} onFile={f => handleFile(f, "B")} color={C.purple} />
                {dataB.length > 0 && <div style={{ fontSize: 11, color: C.green, marginTop: 6 }}>{dataB.length} rows, {headersB.length} columns</div>}
              </div>
            </div>
            <div style={{ textAlign: "center" }}>
              <button style={S.btnPrimary} onClick={proceedToMap}>Map Columns &rarr;</button>
            </div>
          </div>
        )}

        {/* ─── STEP 1: Map Columns ─────────────────────────────────────── */}
        {step === 1 && !loading && (
          <div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
              <h2 style={{ fontSize: 20, fontWeight: 700 }}>Map Columns</h2>
              <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
                <label style={S.label}>Tolerance %</label>
                <input style={{ ...S.input, width: 60 }} value={tolerance} onChange={e => setTolerance(e.target.value)} />
                <button style={S.btn(!allCompareOn)} onClick={() => setAllCompare(!allCompareOn)}>
                  {allCompareOn ? "Deselect All" : "Select All Compare"}
                </button>
              </div>
            </div>
            <div style={S.card}>
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead>
                  <tr>
                    <th style={S.th}>File A Column</th>
                    <th style={S.th}>File B Column</th>
                    <th style={S.th}>Key</th>
                    <th style={S.th}>Compare</th>
                    <th style={S.th}>Ignore Case</th>
                  </tr>
                </thead>
                <tbody>
                  {mappings.map((m, i) => (
                    <tr key={i} style={{ background: m.isKey ? C.blueLight : "transparent" }}>
                      <td style={S.td}><span style={{ fontWeight: 600 }}>{m.colA}</span></td>
                      <td style={S.td}>
                        <select style={{ ...S.input, minWidth: 140 }} value={m.colB} onChange={e => updateMapping(i, "colB", e.target.value)}>
                          <option value="">-- skip --</option>
                          {headersB.map(h => <option key={h} value={h}>{h}</option>)}
                        </select>
                      </td>
                      <td style={{ ...S.td, textAlign: "center" }}>
                        <input type="checkbox" checked={m.isKey} onChange={e => updateMapping(i, "isKey", e.target.checked)} />
                      </td>
                      <td style={{ ...S.td, textAlign: "center" }}>
                        <input type="checkbox" checked={m.compare} disabled={m.isKey} onChange={e => updateMapping(i, "compare", e.target.checked)} />
                      </td>
                      <td style={{ ...S.td, textAlign: "center" }}>
                        <input type="checkbox" checked={m.ignoreCase} onChange={e => updateMapping(i, "ignoreCase", e.target.checked)} />
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div style={{ textAlign: "center", marginTop: 20, display: "flex", justifyContent: "center", gap: 12 }}>
              <button style={S.btn(false)} onClick={() => setStep(0)}>&larr; Back</button>
              <button style={S.btnPrimary} onClick={runComparison}>Run Comparison &rarr;</button>
            </div>
          </div>
        )}

        {/* ─── STEP 2: Results ─────────────────────────────────────────── */}
        {step === 2 && !loading && stats && (
          <div>
            {/* Tabs */}
            <div style={{ display: "flex", gap: 6, marginBottom: 20 }}>
              {["dashboard", "details", "sessions"].map(t => (
                <button key={t} style={S.btn(activeTab === t)} onClick={() => setActiveTab(t)}>
                  {t.charAt(0).toUpperCase() + t.slice(1)}
                </button>
              ))}
              <div style={{ flex: 1 }} />
              <button style={S.btn(false, C.green)} onClick={resetForNew}>+ New Comparison</button>
              {sessions.length > 0 && (
                <button style={S.btn(false, C.purple)} onClick={handleExport} disabled={exporting}>
                  {exporting ? "Exporting..." : "Export XLSX"}
                </button>
              )}
            </div>

            {/* Dashboard Tab */}
            {activeTab === "dashboard" && (
              <div>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))", gap: 12, marginBottom: 24 }}>
                  {[
                    ["File A Rows", stats.totalA, C.blue],
                    ["File B Rows", stats.totalB, C.purple],
                    ["Matched", stats.matched, C.green],
                    ["Mismatched", stats.mismatched, C.red],
                    ["Only in A", stats.onlyA, C.amber],
                    ["Only in B", stats.onlyB, C.purple],
                    ["Duplicates", stats.duplicates, C.textMid],
                  ].map(([label, val, color]) => (
                    <div key={label} style={{ ...S.card, padding: 16, textAlign: "center" }}>
                      <div style={{ fontSize: 24, fontWeight: 800, color }}>{val}</div>
                      <div style={{ fontSize: 11, color: C.textLight, marginTop: 4 }}>{label}</div>
                    </div>
                  ))}
                </div>
                {/* Match rate bar */}
                <div style={{ ...S.card, padding: 16 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8 }}>
                    <span style={{ fontSize: 12, fontWeight: 700 }}>Match Rate</span>
                    <span style={{ fontSize: 12, fontWeight: 700, color: C.green }}>
                      {((stats.matched / (stats.matched + stats.mismatched || 1)) * 100).toFixed(1)}%
                    </span>
                  </div>
                  <div style={{ background: C.border, borderRadius: 6, height: 10, overflow: "hidden" }}>
                    <div style={{ background: `linear-gradient(90deg, ${C.green}, ${C.blueMid})`, height: "100%", width: `${(stats.matched / (stats.matched + stats.mismatched || 1)) * 100}%`, borderRadius: 6, transition: "width .5s" }} />
                  </div>
                </div>
              </div>
            )}

            {/* Details Tab */}
            {activeTab === "details" && (
              <div>
                <div style={{ display: "flex", gap: 8, marginBottom: 16, flexWrap: "wrap", alignItems: "center" }}>
                  {["All", "Matched", "Mismatched", "Only in A", "Only in B"].map(s => (
                    <button key={s} style={S.btn(filterStatus === s, s === "Matched" ? C.green : s === "Mismatched" ? C.red : s === "Only in A" ? C.amber : s === "Only in B" ? C.purple : C.blue)} onClick={() => setFilterStatus(s)}>{s}</button>
                  ))}
                  <input style={{ ...S.input, marginLeft: "auto", minWidth: 180 }} placeholder="Search by key..." value={searchKey} onChange={e => setSearchKey(e.target.value)} />
                  <span style={{ fontSize: 11, color: C.textLight }}>{filteredResults.length} results</span>
                </div>
                <div style={{ ...S.card, maxHeight: 500, overflowY: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse" }}>
                    <thead>
                      <tr>
                        <th style={S.th}>Status</th>
                        {keyMappings.map(m => <th key={m.colA} style={S.th}>{m.colA}</th>)}
                        {compareMappings.map(m => (
                          <Fragment key={m.colA}>
                            <th style={S.th}>{m.colA} (A)</th>
                            <th style={S.th}>{m.colB} (B)</th>
                            <th style={S.th}>Diff</th>
                          </Fragment>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {filteredResults.slice(0, 200).map((r, i) => (
                        <tr key={i} style={{ background: r.status === "Mismatched" ? C.redLight : r.status === "Only in A" ? C.amberLight : r.status === "Only in B" ? C.purpleLight : "transparent" }}>
                          <td style={S.td}><StatusBadge status={r.status} /></td>
                          {keyMappings.map(m => <td key={m.colA} style={{ ...S.td, fontWeight: 600 }}>{r.keyVals[m.colA] ?? ""}</td>)}
                          {compareMappings.map(m => {
                            const d = r.details?.find(dd => dd.colA === m.colA);
                            return (
                              <Fragment key={m.colA}>
                                <td style={S.td}>{d?.valA ?? (r.rowA ? r.rowA[m.colA] ?? "" : "")}</td>
                                <td style={S.td}>{d?.valB ?? (r.rowB ? r.rowB[m.colB] ?? "" : "")}</td>
                                <td style={{ ...S.td, color: d?.status === "Mismatched" ? C.red : C.green, fontWeight: 600 }}>{d?.diff || (r.status === "Matched" ? "\u2713" : "\u2014")}</td>
                              </Fragment>
                            );
                          })}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {filteredResults.length > 200 && (
                    <div style={{ padding: 12, textAlign: "center", fontSize: 11, color: C.textLight }}>Showing 200 of {filteredResults.length} results. Export to see all.</div>
                  )}
                </div>
              </div>
            )}

            {/* Sessions Tab */}
            {activeTab === "sessions" && (
              <div style={S.card}>
                <table style={{ width: "100%", borderCollapse: "collapse" }}>
                  <thead>
                    <tr>
                      <th style={S.th}>#</th>
                      <th style={S.th}>File A</th>
                      <th style={S.th}>File B</th>
                      <th style={S.th}>Matched</th>
                      <th style={S.th}>Mismatched</th>
                      <th style={S.th}>Only A</th>
                      <th style={S.th}>Only B</th>
                      <th style={S.th}>Match Rate</th>
                    </tr>
                  </thead>
                  <tbody>
                    {sessions.map((s, i) => (
                      <tr key={i}>
                        <td style={S.td}>{i + 1}</td>
                        <td style={S.td}>{s.fileAName}</td>
                        <td style={S.td}>{s.fileBName}</td>
                        <td style={{ ...S.td, color: C.green, fontWeight: 700 }}>{s.matched}</td>
                        <td style={{ ...S.td, color: C.red, fontWeight: 700 }}>{s.mismatched}</td>
                        <td style={{ ...S.td, color: C.amber }}>{s.onlyA}</td>
                        <td style={{ ...S.td, color: C.purple }}>{s.onlyB}</td>
                        <td style={S.td}>{((s.matched / (s.matched + s.mismatched || 1)) * 100).toFixed(1)}%</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
