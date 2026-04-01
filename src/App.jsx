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

// ─── Comment helper (matches SK USOPTE reference) ────────────────────────────
function makeComment(diff, valA, valB, status) {
  if (status === "Only in A") return "Only in File A";
  if (status === "Only in B") return "Only in File B";
  if (!diff || diff === "0.0000") return "";
  const num = parseFloat(diff);
  if (!isNaN(num)) {
    if (Math.abs(num) < 1) return "Less than $1 difference";
    return `More $ Difference`;
  }
  return "Value mismatch";
}

// ─── Raw XLSX writer with styling — no SheetJS, streams row-by-row to Uint8Array ──
async function exportToExcel(sessions) {
  const fflate = await import("https://cdn.jsdelivr.net/npm/fflate@0.8.2/esm/browser.js");
  const enc = new TextEncoder();

  const s0 = sessions[0];
  const cMaps0 = s0.mappings.filter(m => m.compare && !m.isKey && m.colA && m.colB);
  const firstRowA = s0.results.find(r => r.rowA)?.rowA || {};
  const allColsA = Object.keys(firstRowA);
  const compareColsSet = new Set(cMaps0.map(m => m.colA));
  const diffNonCompareCols = allColsA.filter(c => !compareColsSet.has(c));

  // XML helpers
  const esc = s => String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
  const colName = n => { let s = ""; for (; n >= 0; n = Math.floor(n / 26) - 1) s = String.fromCharCode(65 + n % 26) + s; return s; };

  // Encode one row to Uint8Array — never accumulate large strings
  // styleIdx: 0=normal, 1=header, 2=mismatch-highlight, 3=only-in-A, 4=only-in-B
  const rowToBytes = (row, ri, styleIdx) => {
    let s = `<row r="${ri + 1}">`;
    for (let ci = 0; ci < row.length; ci++) {
      const val = row[ci];
      const addr = colName(ci) + (ri + 1);
      const si = styleIdx || 0;
      const isNum = val !== "" && val !== null && val !== undefined && !isNaN(Number(val)) && String(val).trim() !== "" && typeof val !== "boolean";
      if (isNum) {
        s += `<c r="${addr}" s="${si}" t="n"><v>${Number(val)}</v></c>`;
      } else {
        s += `<c r="${addr}" s="${si}" t="inlineStr"><is><t>${esc(val)}</t></is></c>`;
      }
    }
    s += "</row>";
    return enc.encode(s);
  };

  const concatParts = (parts) => {
    let total = 0;
    for (let i = 0; i < parts.length; i++) total += parts[i].length;
    const out = new Uint8Array(total);
    let pos = 0;
    for (let i = 0; i < parts.length; i++) { out.set(parts[i], pos); pos += parts[i].length; }
    return out;
  };

  // Build sheet with per-row style indices
  // rows = array of { cells: [...], style: 0|1|2|3|4 }
  const buildSheetBytes = (styledRows) => {
    const parts = [enc.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>')];
    for (let ri = 0; ri < styledRows.length; ri++) {
      parts.push(rowToBytes(styledRows[ri].cells, ri, styledRows[ri].style));
    }
    parts.push(enc.encode("</sheetData></worksheet>"));
    return concatParts(parts);
  };

  // ── Build non-matched results ──
  const nonMatchedResults = sessions.flatMap(s =>
    s.results.filter(r => r.status !== "Matched").map(r => ({ r, s }))
  );

  // ── Sheet: Comparison Results (2 rows per record, like SK reference) ──
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

  // ── Sheet: Difference Mismatch (matches SK USOPTE reference format exactly) ──
  // Columns: Status | Source | Composite Key | <shared non-compare cols> | then for each compare col: ColName | ColName (USOPTE) | Difference | SK Comment
  const diffExtraCols = cMaps0.flatMap(m => [m.colA, `${m.colA} (USOPTE)`, "Difference", "SK Comment"]);
  const diffCols = ["Status", "Source", "Composite Key", ...diffNonCompareCols, ...diffExtraCols];
  const diffStyledRows = [{ cells: diffCols, style: 1 }];

  for (const { r, s } of nonMatchedResults) {
    const kMaps = s.mappings.filter(m => m.isKey && m.colA && m.colB);
    const sCmaps = s.mappings.filter(m => m.compare && !m.isKey && m.colA && m.colB);
    const bMapFull = Object.fromEntries(s.mappings.filter(m => m.colA && m.colB).map(m => [m.colA, m.colB]));
    const key = kMaps.map(m => r.keyVals[m.colA] ?? "").join("|");
    const st = r.status === "Mismatched" ? 2 : r.status === "Only in A" ? 3 : r.status === "Only in B" ? 4 : 0;

    // Row 1 (File A): shared cols from best available, Current filled, USOPTE blank, Diff + Comment filled
    const sharedValsA = diffNonCompareCols.map(c => {
      if (r.rowA) return r.rowA[c] ?? "";
      const cb = bMapFull[c];
      return r.rowB && cb ? (r.rowB[cb] ?? "") : "";
    });
    const compareValsRowA = cMaps0.flatMap(m => {
      const sc = sCmaps.find(x => x.colA === m.colA);
      const valA = sc && r.rowA ? (r.rowA[sc.colA] ?? "") : "";
      const valB = sc && r.rowB ? (r.rowB[sc.colB] ?? "") : "";
      const d = r.details?.find(dd => dd.colA === m.colA);
      const diff = d?.diff || "";
      const comment = makeComment(diff, valA, valB, r.status);
      // Current = valA, Current (USOPTE) = blank, Difference = diff, SK Comment = comment
      return [valA, "", diff ? (parseFloat(diff) || diff) : "", comment];
    });
    diffStyledRows.push({ cells: [r.status, s.fileAName, key, ...sharedValsA, ...compareValsRowA], style: st });

    // Row 2 (File B): shared cols from B, Current blank, USOPTE filled, Diff blank, Comment blank
    const sharedValsB = diffNonCompareCols.map(c => {
      const cb = bMapFull[c];
      return r.rowB && cb ? (r.rowB[cb] ?? "") : "";
    });
    const compareValsRowB = cMaps0.flatMap(m => {
      const sc = sCmaps.find(x => x.colA === m.colA);
      const valB = sc && r.rowB ? (r.rowB[sc.colB] ?? "") : "";
      // Current = blank, Current (USOPTE) = valB, Difference = blank, SK Comment = blank
      return ["", valB, "", ""];
    });
    diffStyledRows.push({ cells: [r.status, s.fileBName, key, ...sharedValsB, ...compareValsRowB], style: st });
  }

  // ── Sheet: Summary ──
  const summaryStyledRows = [
    { cells: ["CompareIQ — Comparison Summary"], style: 1 },
    { cells: [], style: 0 },
    { cells: ["Session #", "File A", "File B", "Total A", "Total B", "Matched", "Mismatched", "Only in A", "Only in B", "Duplicates", "Match Rate"], style: 1 },
    ...sessions.map((s, i) => ({
      cells: [i + 1, s.fileAName, s.fileBName, s.totalA, s.totalB, s.matched, s.mismatched, s.onlyA, s.onlyB, s.duplicates, `${((s.matched / (s.matched + s.mismatched || 1)) * 100).toFixed(1)}%`],
      style: 0
    }))
  ];

  const toBytes = s => enc.encode(s);

  const sheets = [
    { name: "Summary", data: summaryStyledRows },
    { name: "Comparison Results", data: compStyledRows },
    { name: "Difference Mismatch", data: diffStyledRows },
  ];

  // ── Styles XML (cell formatting with colors) ──
  const stylesXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="3">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="7">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1A2332"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFEF2F2"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFBEB"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF5F3FF"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFECFDF5"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="5">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0" applyFill="1" applyBorder="1"/>
  </cellXfs>
</styleSheet>`;

  // ── XLSX boilerplate ──
  const workbookXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets.map((s, i) => `<sheet name="${esc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("")}</sheets></workbook>`;

  const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((s, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("")}<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;

  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((s, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}</Types>`;

  const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;

  // ── Assemble ZIP ──
  const zipFiles = {
    "[Content_Types].xml": toBytes(contentTypes),
    "_rels/.rels": toBytes(rootRels),
    "xl/workbook.xml": toBytes(workbookXML),
    "xl/_rels/workbook.xml.rels": toBytes(workbookRels),
    "xl/styles.xml": toBytes(stylesXML),
  };
  for (let i = 0; i < sheets.length; i++) {
    zipFiles[`xl/worksheets/sheet${i + 1}.xml`] = buildSheetBytes(sheets[i].data);
  }

  const zipped = fflate.zipSync(zipFiles, { level: 0 });
  const blob = new Blob([zipped], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "CompareIQ_Results.xlsx";
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
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
      await exportToExcel(sessions);
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

  const DropZone = ({ label, file, onFile, color }) => {
    const ref = useRef(); const [drag, setDrag] = useState(false);
    return (
      <div onClick={() => ref.current.click()}
        onDragOver={e => { e.preventDefault(); setDrag(true); }}
        onDragLeave={() => setDrag(false)}
        onDrop={e => { e.preventDefault(); setDrag(false); if (e.dataTransfer.files[0]) onFile(e.dataTransfer.files[0]); }}
        style={{ border: `2px dashed ${drag ? color : file ? C.green : C.border}`, borderRadius: 12, padding: "22px 16px", textAlign: "center", cursor: "pointer", background: drag ? `${color}08` :
