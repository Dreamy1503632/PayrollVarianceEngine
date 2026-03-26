
import { useState, useRef } from "react";

// ─── Parsers ─────────────────────────────────────────────────────────────────
function parseCSV(text) {
  const lines = text.trim().split(/\r?\n/);
  if (!lines.length) return [];
  const headers = lines[0].split(",").map(h => h.trim().replace(/^"|"$/g, ""));
  return lines.slice(1).map(line => {
    const vals = [];
    let cur = "", inQ = false;
    for (let i = 0; i <= line.length; i++) {
      const ch = line[i];
      if (ch === '"') { inQ = !inQ; }
      else if (ch === "," && !inQ) { vals.push(cur.trim()); cur = ""; }
      else if (ch === undefined) { vals.push(cur.trim()); }
      else { cur += ch; }
    }
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
function parsePipe(text) {
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
  const prefKeys = ["personnumber", "balancename", "area1", "area2", "area3", "runtype", "balancecategory", "componentruntype"];
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
  const makeKeyA = row => keyMaps.map(m => (row[m.colA] || "").toString().trim().toLowerCase()).join("||");
  const makeKeyB = row => keyMaps.map(m => (row[m.colB] || "").toString().trim().toLowerCase()).join("||");

  const indexB = {};
  for (const row of dataB) {
    const k = makeKeyB(row);
    if (!indexB[k]) indexB[k] = [];
    indexB[k].push(row);
  }

  const results = [];
  const matchedBKeys = new Set();

  for (const rowA of dataA) {
    const key = makeKeyA(rowA);
    const keyVals = Object.fromEntries(keyMaps.map(m => [m.colA, rowA[m.colA] ?? ""]));
    const matchesB = indexB[key] || [];
    if (!matchesB.length) {
      results.push({ key, rowA, rowB: null, status: "Only in A", details: [], keyVals });
      continue;
    }
    const rowB = matchesB[0];
    matchedBKeys.add(key);
    const details = cMaps.map(m => {
      const valA = String(rowA[m.colA] ?? "").trim();
      const valB = String(rowB[m.colB] ?? "").trim();
      const cmpA = m.ignoreCase ? valA.toLowerCase() : valA;
      const cmpB = m.ignoreCase ? valB.toLowerCase() : valB;
      const numA = parseFloat(valA.replace(/,/g, "")), numB = parseFloat(valB.replace(/,/g, ""));
      const isNum = !isNaN(numA) && !isNaN(numB) && valA !== "" && valB !== "";
      let diff = "", status = "Matched";
      if (isNum) {
        const absDiff = Math.abs(numA - numB);
        const pct = numA !== 0 ? absDiff / Math.abs(numA) : (numB !== 0 ? 1 : 0);
        diff = (numB - numA).toFixed(4);
        if (pct > tolPct) status = "Mismatched";
      } else if (cmpA !== cmpB) {
        diff = valA !== valB ? `${valA}→${valB}` : "";
        if (cmpA !== cmpB) status = "Mismatched";
      }
      return { colA: m.colA, colB: m.colB, valA, valB, diff, status };
    });
    results.push({ key, rowA, rowB, status: details.some(d => d.status === "Mismatched") ? "Mismatched" : "Matched", details, keyVals });
  }

  for (const row of dataB) {
    const key = makeKeyB(row);
    if (!matchedBKeys.has(key)) {
      const keyVals = Object.fromEntries(keyMaps.map(m => [m.colA, row[m.colB] ?? ""]));
      results.push({ key, rowA: null, rowB: row, status: "Only in B", details: [], keyVals });
      matchedBKeys.add(key);
    }
  }

  const kc = {};
  for (const row of dataA) { const k = makeKeyA(row); kc[k] = (kc[k] || 0) + 1; }
  return { results, duplicatesA: Object.values(kc).filter(c => c > 1).reduce((s, c) => s + c - 1, 0) };
}

// ─── Comment helper ──────────────────────────────────────────────────────────
function makeComment(diff, status) {
  if (status === "Only in A") return "Only in File A";
  if (status === "Only in B") return "Only in File B";
  if (!diff || diff === "0.0000") return "";
  const num = parseFloat(diff);
  if (!isNaN(num)) return Math.abs(num) < 1 ? "Less than $1 difference" : `Difference: ${diff}`;
  return "Value mismatch";
}

// ─── Excel Export ─────────────────────────────────────────────────────────────
// Writes a valid XLSX with no string-length issues.
// Each sheet XML is built as encoded byte chunks — never one giant string.
async function exportToExcel(sessions) {
  if (!sessions.length) return;

  const s0 = sessions[0];
  const cMaps0 = s0.mappings.filter(m => m.compare && !m.isKey && m.colA && m.colB);
  const firstRowA = s0.results.find(r => r.rowA)?.rowA || {};
  const allColsA = Object.keys(firstRowA);
  const compareColsSet = new Set(cMaps0.map(m => m.colA));
  const sharedCols = allColsA.filter(c => !compareColsSet.has(c));

  // Header: Status | Source | Composite Key | <shared> | <ColA | ColB | Difference | SK Comment> per cmap
  const headerRow = [
    "Status", "Source", "Composite Key",
    ...sharedCols,
    ...cMaps0.flatMap(m => [m.colA, m.colB, "Difference", "SK Comment"]),
  ];

  const enc = new TextEncoder();

  const xesc = v => String(v ?? "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");

  // column letter from 0-index
  const colLetter = n => {
    let s = "";
    for (let x = n; x >= 0; x = Math.floor(x / 26) - 1)
      s = String.fromCharCode(65 + (x % 26)) + s;
    return s;
  };

  // encode one row to bytes immediately
  const encodeRow = (cells, rowIdx) => {
    let s = `<row r="${rowIdx}">`;
    for (let c = 0; c < cells.length; c++) {
      const sv = String(cells[c] ?? "");
      if (!sv) continue;
      const addr = colLetter(c) + rowIdx;
      const n = Number(sv.replace(/,/g, ""));
      if (sv !== "" && !isNaN(n) && sv.trim() !== "") {
        s += `<c r="${addr}"><v>${n}</v></c>`;
      } else {
        s += `<c r="${addr}" t="inlineStr"><is><t>${xesc(sv)}</t></is></c>`;
      }
    }
    s += "</row>";
    return enc.encode(s);
  };

  // concat array of Uint8Arrays
  const concat = parts => {
    const len = parts.reduce((n, b) => n + b.length, 0);
    const out = new Uint8Array(len);
    let pos = 0;
    for (const b of parts) { out.set(b, pos); pos += b.length; }
    return out;
  };

  // build full sheet bytes from a rows-generator function
  const buildSheet = (getRows) => {
    const parts = [];
    parts.push(enc.encode(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<sheetData>'
    ));
    let rowIdx = 1;
    for (const row of getRows()) {
      parts.push(encodeRow(row, rowIdx++));
    }
    parts.push(enc.encode("</sheetData></worksheet>"));
    return concat(parts);
  };

  // generator for comparison rows (2 per record)
  function* compRowGen(filterFn) {
    yield headerRow;
    for (const s of sessions) {
      const kMaps = s.mappings.filter(m => m.isKey && m.colA && m.colB);
      const sCmaps = s.mappings.filter(m => m.compare && !m.isKey && m.colA && m.colB);
      const bMap = Object.fromEntries(s.mappings.filter(m => m.colA && m.colB).map(m => [m.colA, m.colB]));

      for (const r of s.results) {
        if (!filterFn(r)) continue;
        const key = kMaps.map(m => r.keyVals[m.colA] ?? "").join("|");

        // row A
        const sharedA = sharedCols.map(c => r.rowA ? (r.rowA[c] ?? "") : "");
        const cmpA = cMaps0.flatMap(m => {
          const sc = sCmaps.find(x => x.colA === m.colA);
          const valA = sc && r.rowA ? (r.rowA[sc.colA] ?? "") : "";
          const valB = sc && r.rowB ? (r.rowB[sc.colB] ?? "") : "";
          const d = r.details?.find(d => d.colA === m.colA);
          const rawDiff = d?.diff ?? "";
          const numDiff = rawDiff !== "" && !isNaN(parseFloat(rawDiff)) ? parseFloat(rawDiff) : rawDiff;
          const comment = makeComment(rawDiff, r.status);
          return [valA, valB, numDiff, comment];
        });
        yield [r.status, s.fileAName, key, ...sharedA, ...cmpA];

        // row B
        const sharedB = sharedCols.map(c => {
          const cb = bMap[c] || c;
          return r.rowB ? (r.rowB[cb] ?? "") : "";
        });
        const cmpB = cMaps0.flatMap(m => {
          const sc = sCmaps.find(x => x.colA === m.colA);
          const valA = sc && r.rowA ? (r.rowA[sc.colA] ?? "") : "";
          const valB = sc && r.rowB ? (r.rowB[sc.colB] ?? "") : "";
          return [valA, valB, "", ""];
        });
        yield [r.status, s.fileBName, key, ...sharedB, ...cmpB];
      }
    }
  }

  function* summaryGen() {
    yield ["CompareIQ — Comparison Summary"];
    yield [];
    yield ["2 rows per record: Row 1 = File A values, Row 2 = File B values"];
    yield ["Columns: Status | Source | Composite Key | shared cols | ColA | ColB | Difference | SK Comment"];
    yield [];
    yield ["Session", "File A", "File B", "Total A", "Total B", "Matched", "Mismatched", "Only in A", "Only in B", "Match %"];
    for (const [i, s] of sessions.entries()) {
      yield [i + 1, s.fileAName, s.fileBName, s.totalA, s.totalB, s.matched, s.mismatched, s.onlyA, s.onlyB,
        ((s.matched / (s.matched + s.mismatched || 1)) * 100).toFixed(1) + "%"];
    }
  }

  const sheetDefs = [
    { name: "Summary",         data: buildSheet(summaryGen) },
    { name: "All Records",     data: buildSheet(() => compRowGen(() => true)) },
    { name: "Mismatches Only", data: buildSheet(() => compRowGen(r => r.status !== "Matched")) },
    { name: "Differences",     data: buildSheet(() => compRowGen(r => r.status === "Mismatched")) },
    { name: "Only in A",       data: buildSheet(() => compRowGen(r => r.status === "Only in A")) },
    { name: "Only in B",       data: buildSheet(() => compRowGen(r => r.status === "Only in B")) },
  ];

  // ── Static XML parts ──
  const wbXML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<sheets>' +
    sheetDefs.map((s, i) => `<sheet name="${xesc(s.name)}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join("") +
    '</sheets></workbook>';

  const wbRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    sheetDefs.map((s, i) =>
      `<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`
    ).join("") + '</Relationships>';

  const ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    sheetDefs.map((s, i) =>
      `<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
    ).join("") + '</Types>';

  const rootRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
    '</Relationships>';

  // ── ZIP builder ──
  // CRC32 lookup table
  const crcTable = (() => {
    const t = new Uint32Array(256);
    for (let i = 0; i < 256; i++) {
      let c = i;
      for (let j = 0; j < 8; j++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
      t[i] = c;
    }
    return t;
  })();

  const crc32 = data => {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < data.length; i++) c = crcTable[(c ^ data[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  };

  const u32le = n => { const b = new Uint8Array(4); new DataView(b.buffer).setUint32(0, n, true); return b; };
  const u16le = n => { const b = new Uint8Array(2); new DataView(b.buffer).setUint16(0, n, true); return b; };

  // DOS date/time (fixed: 2024-01-01 00:00:00)
  const dosDate = new Uint8Array([0x21, 0x00]); // 2024-01-01
  const dosTime = new Uint8Array([0x00, 0x00]);

  const zipFiles = [
    { name: "[Content_Types].xml",        bytes: enc.encode(ct) },
    { name: "_rels/.rels",                bytes: enc.encode(rootRels) },
    { name: "xl/workbook.xml",            bytes: enc.encode(wbXML) },
    { name: "xl/_rels/workbook.xml.rels", bytes: enc.encode(wbRels) },
    ...sheetDefs.map((s, i) => ({ name: `xl/worksheets/sheet${i+1}.xml`, bytes: s.data })),
  ];

  const parts = [];         // all bytes to concatenate at end
  const cdirEntries = [];   // central directory info
  let offset = 0;

  for (const f of zipFiles) {
    const nameBytes = enc.encode(f.name);
    const crc = crc32(f.bytes);
    const size = f.bytes.length;

    // Local file header (30 bytes + name)
    const lhParts = [
      new Uint8Array([0x50,0x4B,0x03,0x04]), // signature
      u16le(20),            // version needed
      u16le(0),             // general purpose bit flag
      u16le(0),             // compression method: store
      dosTime, dosDate,     // last mod time/date
      u32le(crc),           // crc-32
      u32le(size),          // compressed size
      u32le(size),          // uncompressed size
      u16le(nameBytes.length), // file name length
      u16le(0),             // extra field length
      nameBytes,
    ];

    const lhLen = lhParts.reduce((n, b) => n + b.length, 0);
    const lh = new Uint8Array(lhLen);
    let pos = 0;
    for (const b of lhParts) { lh.set(b, pos); pos += b.length; }

    parts.push(lh, f.bytes);
    cdirEntries.push({ nameBytes, crc, size, offset });
    offset += lhLen + size;
  }

  // Central directory
  const cdStart = offset;
  for (const e of cdirEntries) {
    const cdParts = [
      new Uint8Array([0x50,0x4B,0x01,0x02]), // signature
      u16le(20),            // version made by
      u16le(20),            // version needed
      u16le(0),             // flags
      u16le(0),             // compression: store
      dosTime, dosDate,
      u32le(e.crc),
      u32le(e.size),
      u32le(e.size),
      u16le(e.nameBytes.length),
      u16le(0),             // extra field length
      u16le(0),             // file comment length
      u16le(0),             // disk number start
      u16le(0),             // internal attributes
      u32le(0),             // external attributes
      u32le(e.offset),      // relative offset of local header
      e.nameBytes,
    ];
    const cd = new Uint8Array(cdParts.reduce((n, b) => n + b.length, 0));
    let pos = 0;
    for (const b of cdParts) { cd.set(b, pos); pos += b.length; }
    parts.push(cd);
    offset += cd.length;
  }

  const cdSize = offset - cdStart;

  // End of central directory
  const eocdParts = [
    new Uint8Array([0x50,0x4B,0x05,0x06]), // signature
    u16le(0),              // disk number
    u16le(0),              // disk with cd
    u16le(zipFiles.length),
    u16le(zipFiles.length),
    u32le(cdSize),
    u32le(cdStart),
    u16le(0),              // comment length
  ];
  const eocd = new Uint8Array(eocdParts.reduce((n, b) => n + b.length, 0));
  let epos = 0;
  for (const b of eocdParts) { eocd.set(b, epos); epos += b.length; }
  parts.push(eocd);

  // Trigger download
  const blob = new Blob(parts, { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "CompareIQ_Results.xlsx";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 15000);
}

// ─── Styles ───────────────────────────────────────────────────────────────────
const C = {
  bg: "#F7F9FC", surface: "#FFFFFF", border: "#D1D9E6", borderStrong: "#A0B0C8",
  text: "#1A2332", textMid: "#4A5568", textLight: "#718096",
  blue: "#1A56DB", blueLight: "#EBF5FF",
  purple: "#7C3AED", purpleLight: "#F5F3FF",
  green: "#059669", greenLight: "#ECFDF5",
  red: "#DC2626", redLight: "#FEF2F2",
  amber: "#D97706", amberLight: "#FFFBEB",
  headerBg: "#1A2332",
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
    setExporting(true);
    setError("");
    try {
      await new Promise(r => setTimeout(r, 50));
      await exportToExcel(sessions);
    } catch (e) {
      setError(`Export failed: ${e.message}`);
      console.error(e);
    }
    setExporting(false);
  };

  const handleFile = async (file, which) => {
    setError(""); setLoading(true);
    try {
      const data = await parseFile(file);
      if (!data.length) throw new Error("No data found in file");
      const headers = Object.keys(data[0]);
      if (which === "A") { setFileA(file); setDataA(data); setHeadersA(headers); }
      else { setFileB(file); setDataB(data); setHeadersB(headers); }
    } catch (e) { setError(`Error reading file: ${e.message}`); }
    setLoading(false);
  };

  const proceedToMap = () => {
    if (!dataA.length || !dataB.length) { setError("Upload both files first."); return; }
    setMappings(autoMapHeaders(headersA, headersB));
    setStep(1);
  };

  const updateMapping = (i, field, val) =>
    setMappings(p => p.map((m, idx) => idx === i ? { ...m, [field]: val } : m));

  const setAllCompare = val => setMappings(p => p.map(m => m.isKey ? m : { ...m, compare: val }));

  const runComparison = () => {
    if (!mappings.filter(m => m.isKey).length) { setError("Mark at least one KEY column."); return; }
    setLoading(true);
    setTimeout(() => {
      try {
        const res = compareDatasets(dataA, dataB, mappings, tolerance);
        setResults(res);
        const matched   = res.results.filter(r => r.status === "Matched").length;
        const mismatched= res.results.filter(r => r.status === "Mismatched").length;
        const onlyA     = res.results.filter(r => r.status === "Only in A").length;
        const onlyB     = res.results.filter(r => r.status === "Only in B").length;
        setSessions(prev => [...prev, {
          fileAName: fileA?.name || "File A", fileBName: fileB?.name || "File B",
          results: res.results, totalA: dataA.length, totalB: dataB.length,
          matched, mismatched, onlyA, onlyB, duplicates: res.duplicatesA,
          mappings: mappings.map(m => ({ ...m })),
        }]);
        setActiveTab("dashboard");
        setStep(2);
      } catch(e) { setError(`Comparison failed: ${e.message}`); }
      setLoading(false);
    }, 100);
  };

  const stats = results ? {
    totalA: dataA.length, totalB: dataB.length,
    matched:    results.results.filter(r => r.status === "Matched").length,
    mismatched: results.results.filter(r => r.status === "Mismatched").length,
    onlyA:      results.results.filter(r => r.status === "Only in A").length,
    onlyB:      results.results.filter(r => r.status === "Only in B").length,
    duplicates: results.duplicatesA,
  } : null;

  const keyMappings     = mappings.filter(m => m.isKey && m.colA && m.colB);
  const compareMappings = mappings.filter(m => m.compare && !m.isKey && m.colA && m.colB);

  const filteredResults = results ? results.results.filter(r =>
    (filterStatus === "All" || r.status === filterStatus) &&
    (!searchKey || r.key.toLowerCase().includes(searchKey.toLowerCase()))
  ) : [];

  // ── Sub-components ──────────────────────────────────────────────────────
  const DropZone = ({ label, file, onFile, color }) => {
    const ref = useRef();
    const [drag, setDrag] = useState(false);
    return (
      <div onClick={() => ref.current.click()}
        onDragOver={e => { e.preventDefault(); setDrag(true); }}
        onDragLeave={() => setDrag(false)}
        onDrop={e => { e.preventDefault(); setDrag(false); if (e.dataTransfer.files[0]) onFile(e.dataTransfer.files[0]); }}
        style={{ border: `2px dashed ${drag ? color : file ? C.green : C.border}`, borderRadius: 12, padding: "22px 16px", textAlign: "center", cursor: "pointer", background: drag ? `${color}08` : file ? C.greenLight : C.bg, transition: "all .2s" }}>
        <input ref={ref} type="file" accept=".csv,.xlsx,.xls,.txt,.tsv" style={{ display: "none" }}
          onChange={e => { if (e.target.files[0]) onFile(e.target.files[0]); }} />
        <div style={{ fontSize: 26, marginBottom: 6 }}>{file ? "✅" : "📂"}</div>
        <div style={{ fontWeight: 700, color: file ? C.green : C.textMid, fontSize: 13 }}>{file ? file.name : label}</div>
        <div style={{ color: C.textLight, fontSize: 11, marginTop: 4 }}>
          {file ? `${(file.size/1024).toFixed(1)} KB · click to replace` : "CSV · Excel · TXT · TSV"}
        </div>
      </div>
    );
  };

  const Radio = ({ checked, onChange, color = C.amber }) => (
    <div onClick={onChange} style={{ width: 17, height: 17, borderRadius: "50%", border: `2px solid ${checked ? color : C.borderStrong}`, background: checked ? color : "#fff", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", margin: "auto", flexShrink: 0 }}>
      {checked && <div style={{ width: 6, height: 6, borderRadius: "50%", background: "#fff" }} />}
    </div>
  );

  const Checkbox = ({ checked, onChange, color = C.blue }) => (
    <div onClick={onChange} style={{ width: 17, height: 17, borderRadius: 4, border: `2px solid ${checked ? color : C.borderStrong}`, background: checked ? color : "#fff", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", margin: "auto", flexShrink: 0 }}>
      {checked && <span style={{ color: "#fff", fontSize: 11, fontWeight: 900, lineHeight: 1 }}>✓</span>}
    </div>
  );

  const ColSelect = ({ value, options, onChange, color }) => (
    <div style={{ position: "relative", flex: 1 }}>
      <select value={value} onChange={e => onChange(e.target.value)}
        style={{ ...S.input, width: "100%", paddingRight: 24, appearance: "none", cursor: "pointer", borderColor: value ? `${color}55` : C.border }}>
        <option value="">— not mapped —</option>
        {options.map(o => <option key={o} value={o}>{o}</option>)}
      </select>
      <span style={{ position: "absolute", right: 8, top: "50%", transform: "translateY(-50%)", color: C.textLight, pointerEvents: "none", fontSize: 10 }}>▾</span>
    </div>
  );

  // ── Comparison Sheet tab (inline preview) ─────────────────────────────
  const ComparisonSheetTab = () => {
    if (!sessions.length) return <div style={{ padding: 32, color: C.textLight, textAlign: "center" }}>No session data</div>;
    const s = sessions[sessions.length - 1];
    const cM = s.mappings.filter(m => m.compare && !m.isKey && m.colA && m.colB);
    const kM = s.mappings.filter(m => m.isKey && m.colA && m.colB);
    const firstRowA = s.results.find(r => r.rowA)?.rowA || {};
    const allCols = Object.keys(firstRowA);
    const cSet = new Set(cM.map(m => m.colA));
    const shared = allCols.filter(c => !cSet.has(c));
    const preview = s.results.slice(0, 80);

    return (
      <div>
        <div style={{ background: C.blueLight, border: `1px solid ${C.blue}22`, borderRadius: 9, padding: "9px 14px", marginBottom: 10, fontSize: 11, color: C.textMid }}>
          📋 <b>2 rows per record</b> — Row 1 = File A &nbsp;|&nbsp; Row 2 = File B &nbsp;|&nbsp; Showing first 80 records. Export Excel for all {s.results.length} records.
        </div>
        <div style={{ ...S.card }}>
          <div style={{ overflowX: "auto", maxHeight: 520, overflowY: "auto" }}>
            <table style={{ borderCollapse: "collapse", fontSize: 10 }}>
              <thead style={{ position: "sticky", top: 0, zIndex: 2 }}>
                <tr>
                  <th style={{ ...S.th, background: C.headerBg, color: "#fff", minWidth: 90 }}>Status</th>
                  <th style={{ ...S.th, background: C.headerBg, color: "#fff", minWidth: 110 }}>Source</th>
                  <th style={{ ...S.th, background: C.headerBg, color: "#fff", minWidth: 160 }}>Composite Key</th>
                  {shared.map(c => <th key={c} style={{ ...S.th, background: C.headerBg, color: "#CBD5E1", minWidth: 90 }}>{c}</th>)}
                  {cM.map(m => (
                    <React.Fragment key={m.colA}>
                      <th style={{ ...S.th, background: "#7F1D1D", color: "#FCA5A5", borderLeft: "2px solid #991B1B", minWidth: 100 }}>{m.colA}</th>
                      <th style={{ ...S.th, background: "#7F1D1D", color: "#FCA5A5", minWidth: 100 }}>{m.colB}</th>
                      <th style={{ ...S.th, background: "#7F1D1D", color: "#FCA5A5", minWidth: 80 }}>Difference</th>
                      <th style={{ ...S.th, background: "#7F1D1D", color: "#FCA5A5", borderRight: "2px solid #991B1B", minWidth: 130 }}>SK Comment</th>
                    </React.Fragment>
                  ))}
                </tr>
              </thead>
              <tbody>
                {preview.map((r, ri) => {
                  const key = kM.map(m => r.keyVals[m.colA] ?? "").join("|");
                  const isMis = r.status !== "Matched";
                  const bg = isMis ? "#FFF5F5" : ri % 2 === 0 ? C.surface : C.bg;
                  const bgB = isMis ? "#FFF0F0" : ri % 2 === 0 ? "#F5F8FF" : "#EEF2FF";
                  const bMap = Object.fromEntries(s.mappings.filter(m => m.colA && m.colB).map(m => [m.colA, m.colB]));

                  return (
                    <React.Fragment key={ri}>
                      {/* Row A */}
                      <tr style={{ background: bg }}>
                        <td style={S.td}><StatusBadge status={r.status} /></td>
                        <td style={{ ...S.td, color: C.blue, fontWeight: 600 }}>{s.fileAName}</td>
                        <td style={{ ...S.td, color: C.textMid, maxWidth: 180, overflow: "hidden", textOverflow: "ellipsis" }}>{key}</td>
                        {shared.map(c => <td key={c} style={{ ...S.td, color: C.textMid }}>{r.rowA ? r.rowA[c] ?? "" : ""}</td>)}
                        {cM.map(m => {
                          const d = r.details?.find(d => d.colA === m.colA);
                          const isMisCol = d?.status === "Mismatched";
                          return (
                            <React.Fragment key={m.colA}>
                              <td style={{ ...S.td, borderLeft: "2px solid #FECACA", color: isMisCol ? C.red : C.text, background: isMisCol ? "#FFDDDD" : undefined, fontWeight: isMisCol ? 700 : 400 }}>{d?.valA ?? (r.rowA ? r.rowA[m.colA] ?? "—" : "—")}</td>
                              <td style={{ ...S.td, color: isMisCol ? C.purple : C.textLight, background: isMisCol ? "#FFE8FF" : undefined, fontWeight: isMisCol ? 700 : 400 }}>{d?.valB ?? (r.rowB ? r.rowB[m.colB] ?? "—" : "—")}</td>
                              <td style={{ ...S.td, color: isMisCol ? C.red : C.textLight, fontWeight: isMisCol ? 700 : 400 }}>{isMisCol ? (d?.diff || "—") : "—"}</td>
                              <td style={{ ...S.td, color: C.amber, fontStyle: "italic", borderRight: "2px solid #FECACA", fontSize: 10 }}>{isMisCol ? makeComment(d?.diff || "", r.status) : ""}</td>
                            </React.Fragment>
                          );
                        })}
                      </tr>
                      {/* Row B */}
                      <tr style={{ background: bgB, borderBottom: `2px solid ${C.border}` }}>
                        <td style={{ ...S.td, color: C.textLight, fontSize: 10 }}></td>
                        <td style={{ ...S.td, color: C.purple, fontWeight: 600 }}>{s.fileBName}</td>
                        <td style={{ ...S.td, color: C.textLight }}>{key}</td>
                        {shared.map(c => {
                          const cb = bMap[c] || c;
                          return <td key={c} style={{ ...S.td, color: C.textLight }}>{r.rowB ? r.rowB[cb] ?? "" : ""}</td>;
                        })}
                        {cM.map(m => {
                          const d = r.details?.find(d => d.colA === m.colA);
                          const isMisCol = d?.status === "Mismatched";
                          return (
                            <React.Fragment key={m.colA}>
                              <td style={{ ...S.td, borderLeft: "2px solid #FECACA", color: C.textLight }}>{d?.valA ?? (r.rowA ? r.rowA[m.colA] ?? "—" : "—")}</td>
                              <td style={{ ...S.td, color: isMisCol ? C.purple : C.textLight, fontWeight: isMisCol ? 700 : 400 }}>{d?.valB ?? (r.rowB ? r.rowB[m.colB] ?? "—" : "—")}</td>
                              <td style={{ ...S.td, color: C.textLight }}>—</td>
                              <td style={{ ...S.td, borderRight: "2px solid #FECACA" }}>—</td>
                            </React.Fragment>
                          );
                        })}
                      </tr>
                    </React.Fragment>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
        <div style={{ marginTop: 8, fontSize: 10, color: C.textLight, textAlign: "center" }}>
          <b style={{ color: C.red }}>Red columns</b> = comparison fields &nbsp;·&nbsp; Export Excel for complete data with all {s.results.length} records
        </div>
      </div>
    );
  };

  // ── Render ───────────────────────────────────────────────────────────────
  return (
    <div style={{ minHeight: "100vh", background: C.bg, color: C.text, fontFamily: "'Inter','Segoe UI',sans-serif" }}>
      <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet" />

      {/* Header */}
      <div style={{ background: C.headerBg, padding: "12px 28px", display: "flex", alignItems: "center", justifyContent: "space-between", position: "sticky", top: 0, zIndex: 50, boxShadow: "0 2px 8px rgba(0,0,0,0.15)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ width: 32, height: 32, borderRadius: 8, background: "linear-gradient(135deg,#3B82F6,#7C3AED)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16 }}>⚖️</div>
          <div>
            <div style={{ fontWeight: 800, fontSize: 16, color: "#fff" }}>CompareIQ</div>
            <div style={{ fontSize: 9, color: "#94A3B8", letterSpacing: 1.5 }}>INTELLIGENT DATASET COMPARISON</div>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
          {STEPS.map((s, i) => (
            <div key={s} style={{ display: "flex", alignItems: "center", gap: 3 }}>
              <div onClick={() => step > i && setStep(i)}
                style={{ display: "flex", alignItems: "center", gap: 6, padding: "5px 12px", borderRadius: 20, background: step === i ? "#3B82F6" : step > i ? "#1E40AF" : "transparent", border: `1px solid ${step === i ? "#3B82F6" : step > i ? "#3B82F6" : "#475569"}`, color: step >= i ? "#fff" : "#94A3B8", fontSize: 11, fontWeight: 600, cursor: step > i ? "pointer" : "default" }}>
                <div style={{ width: 14, height: 14, borderRadius: "50%", background: step > i ? "#22C55E" : step === i ? "#fff" : "#475569", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 8, color: step === i ? "#3B82F6" : "#fff", fontWeight: 900 }}>
                  {step > i ? "✓" : i + 1}
                </div>
                {s}
              </div>
              {i < STEPS.length - 1 && <div style={{ width: 14, height: 1, background: "#475569" }} />}
            </div>
          ))}
          {sessions.length > 0 && (
            <div style={{ marginLeft: 12, padding: "4px 10px", background: "#14532D", border: "1px solid #22C55E55", borderRadius: 20, fontSize: 10, color: "#22C55E" }}>
              {sessions.length} session{sessions.length > 1 ? "s" : ""}
            </div>
          )}
        </div>
      </div>

      <div style={{ maxWidth: 1200, margin: "0 auto", padding: "28px 20px" }}>
        {error && (
          <div style={{ background: C.redLight, border: `1px solid ${C.red}33`, color: C.red, borderRadius: 9, padding: "9px 14px", marginBottom: 14, fontSize: 12, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            ⚠️ {error}
            <button onClick={() => setError("")} style={{ background: "none", border: "none", color: C.red, cursor: "pointer", fontSize: 16 }}>×</button>
          </div>
        )}
        {loading && <div style={{ background: C.blueLight, border: `1px solid ${C.blue}33`, borderRadius: 9, padding: "9px 14px", marginBottom: 14, fontSize: 12, color: C.blue }}>⏳ Processing…</div>}
        {exporting && <div style={{ background: C.amberLight, border: `1px solid ${C.amber}33`, borderRadius: 9, padding: "9px 14px", marginBottom: 14, fontSize: 12, color: C.amber, fontWeight: 600 }}>⏳ Building Excel file — encoding {sessions.reduce((a, s) => a + s.results.length * 2, 0).toLocaleString()} rows…</div>}

        {/* STEP 0 — UPLOAD */}
        {step === 0 && (
          <div>
            <div style={{ textAlign: "center", marginBottom: 28 }}>
              <h1 style={{ fontWeight: 800, fontSize: 28, margin: "0 0 6px" }}>Upload Your Datasets</h1>
              <p style={{ color: C.textLight, margin: 0, fontSize: 13 }}>CSV · Excel · TXT · TSV · Pipe-delimited</p>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 24 }}>
              {[
                { label: "Dataset A — Source / Expected", which: "A", file: fileA, data: dataA, headers: headersA, color: C.blue },
                { label: "Dataset B — Target / Actual", which: "B", file: fileB, data: dataB, headers: headersB, color: C.purple },
              ].map(({ label, which, file, data, headers, color }) => (
                <div key={which}>
                  <div style={{ ...S.label, marginBottom: 8, color }}>● DATASET {which}</div>
                  <DropZone label={label} file={file} color={color} onFile={f => handleFile(f, which)} />
                  {data.length > 0 && (
                    <div style={{ marginTop: 8, ...S.card, padding: 12 }}>
                      <div style={{ color, fontSize: 11, fontWeight: 700, marginBottom: 6 }}>{data.length.toLocaleString()} rows · {headers.length} columns</div>
                      <div style={{ overflowX: "auto" }}>
                        <table style={{ borderCollapse: "collapse", fontSize: 10 }}>
                          <thead><tr>{headers.slice(0, 7).map(h => <th key={h} style={{ padding: "3px 8px", background: "#F0F4FA", color: C.textMid, textAlign: "left", whiteSpace: "nowrap" }}>{h}</th>)}{headers.length > 7 && <th style={{ color: C.textLight, padding: "3px 6px" }}>+{headers.length - 7}</th>}</tr></thead>
                          <tbody>{data.slice(0, 2).map((r, i) => <tr key={i}>{headers.slice(0, 7).map(h => <td key={h} style={{ padding: "3px 8px", color: C.textMid }}>{r[h]}</td>)}</tr>)}</tbody>
                        </table>
                      </div>
                    </div>
                  )}
                </div>
              ))}
            </div>
            <div style={{ textAlign: "center" }}>
              <button onClick={proceedToMap} disabled={!dataA.length || !dataB.length}
                style={{ ...S.btnPrimary, opacity: dataA.length && dataB.length ? 1 : 0.4 }}>
                Continue to Column Mapping →
              </button>
            </div>
          </div>
        )}

        {/* STEP 1 — MAPPING */}
        {step === 1 && (
          <div>
            <h1 style={{ fontWeight: 800, fontSize: 22, margin: "0 0 16px" }}>Column Mapping</h1>
            <div style={{ background: C.blueLight, border: `1px solid ${C.blue}22`, borderRadius: 9, padding: "10px 14px", marginBottom: 14, fontSize: 12, color: C.textMid }}>
              <b style={{ color: C.blue }}>ℹ </b>
              Mark <b style={{ color: C.amber }}>Key</b> columns for row matching. Check <b style={{ color: C.green }}>Compare</b> to include in value comparison.
              Excel output: <code style={{ background: "#E0EAFF", padding: "1px 5px", borderRadius: 3 }}>Status | Source | Composite Key | shared cols | ColA | ColB | Difference | SK Comment</code> — 2 rows per record.
            </div>
            <div style={{ display: "flex", alignItems: "center", gap: 14, marginBottom: 10, padding: "10px 14px", ...S.card, boxShadow: "none" }}>
              <span style={S.label}>TOLERANCE</span>
              <input type="number" min="0" max="100" step="0.1" value={tolerance} onChange={e => setTolerance(e.target.value)}
                style={{ ...S.input, width: 70, textAlign: "center", borderColor: C.green, color: C.green, fontWeight: 700 }} />
              <span style={{ color: C.green, fontWeight: 700 }}>%</span>
              <span style={{ fontSize: 11, color: C.textLight }}>Numeric variance allowed</span>
              <div style={{ flex: 1 }} />
              <button onClick={() => setAllCompare(true)} style={{ ...S.btn(false, C.green), fontSize: 11 }}>Select All Compare</button>
              <button onClick={() => setAllCompare(false)} style={{ ...S.btn(false, C.red), fontSize: 11 }}>Deselect All</button>
              <button onClick={() => setMappings(p => [...p, { colA: headersA[0] || "", colB: headersB[0] || "", isKey: false, compare: true, ignoreCase: false }])}
                style={S.btn(false)}>+ Add Row</button>
            </div>
            <div style={{ ...S.card, marginBottom: 16 }}>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 60px 80px 90px 32px", background: "#F0F4FA", padding: "8px 14px", borderBottom: `1px solid ${C.border}` }}>
                <div style={{ ...S.label, color: C.blue }}>DATASET A — {fileA?.name}</div>
                <div style={{ ...S.label, color: C.purple }}>DATASET B — {fileB?.name}</div>
                <div style={{ ...S.label, color: C.amber, textAlign: "center" }}>KEY 🔑</div>
                <div style={{ ...S.label, color: C.green, textAlign: "center" }}>COMPARE</div>
                <div style={{ ...S.label, textAlign: "center" }}>IGNORE CASE</div>
                <div></div>
              </div>
              <div style={{ maxHeight: 400, overflowY: "auto" }}>
                {mappings.map((m, i) => (
                  <div key={i} style={{ display: "grid", gridTemplateColumns: "1fr 1fr 60px 80px 90px 32px", alignItems: "center", padding: "6px 14px", borderBottom: `1px solid ${C.border}20`, background: m.isKey ? "#FFFBEB" : i % 2 ? C.surface : C.bg }}>
                    <div style={{ paddingRight: 8 }}><ColSelect value={m.colA} options={headersA} color={C.blue} onChange={v => updateMapping(i, "colA", v)} /></div>
                    <div style={{ paddingRight: 8 }}><ColSelect value={m.colB} options={headersB} color={C.purple} onChange={v => updateMapping(i, "colB", v)} /></div>
                    <div style={{ textAlign: "center" }}><Radio checked={m.isKey} color={C.amber} onChange={() => updateMapping(i, "isKey", !m.isKey)} /></div>
                    <div style={{ textAlign: "center" }}><Checkbox checked={!!m.compare && !m.isKey} color={C.green} onChange={() => { if (!m.isKey) updateMapping(i, "compare", !m.compare); }} /></div>
                    <div style={{ textAlign: "center" }}><Checkbox checked={!!m.ignoreCase} color={C.textLight} onChange={() => updateMapping(i, "ignoreCase", !m.ignoreCase)} /></div>
                    <div style={{ textAlign: "center" }}>
                      <button onClick={() => setMappings(p => p.filter((_, j) => j !== i))} style={{ background: "none", border: "none", color: C.textLight, cursor: "pointer", fontSize: 16 }}>×</button>
                    </div>
                  </div>
                ))}
              </div>
              <div style={{ padding: "8px 14px", background: "#F8FAFC", borderTop: `1px solid ${C.border}`, display: "flex", gap: 16, fontSize: 11 }}>
                <span style={{ color: C.textLight }}>Total: <b style={{ color: C.blue }}>{mappings.length}</b></span>
                <span style={{ color: C.textLight }}>Keys: <b style={{ color: C.amber }}>{mappings.filter(m => m.isKey).length}</b></span>
                <span style={{ color: C.textLight }}>Compare: <b style={{ color: C.green }}>{mappings.filter(m => m.compare && !m.isKey).length}</b></span>
              </div>
            </div>
            <div style={{ display: "flex", gap: 10, justifyContent: "center" }}>
              <button onClick={() => setStep(0)} style={{ ...S.btn(false), padding: "10px 22px" }}>← Back</button>
              <button onClick={runComparison} style={S.btnPrimary}>🚀 Run Comparison</button>
            </div>
          </div>
        )}

        {/* STEP 2 — RESULTS */}
        {step === 2 && stats && (
          <div>
            {/* Stat cards */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(7,1fr)", gap: 8, marginBottom: 16 }}>
              {[
                { label: "Total A",    val: stats.totalA,     color: C.blue,    bg: C.blueLight },
                { label: "Total B",    val: stats.totalB,     color: C.purple,  bg: C.purpleLight },
                { label: "Matched",    val: stats.matched,    color: C.green,   bg: C.greenLight },
                { label: "Mismatched", val: stats.mismatched, color: C.red,     bg: C.redLight },
                { label: "Only in A",  val: stats.onlyA,      color: C.amber,   bg: C.amberLight },
                { label: "Only in B",  val: stats.onlyB,      color: C.purple,  bg: C.purpleLight },
                { label: "Duplicates", val: stats.duplicates, color: C.textMid, bg: "#F1F5F9" },
              ].map(s => (
                <div key={s.label} style={{ background: s.bg, border: `1px solid ${s.color}30`, borderTop: `3px solid ${s.color}`, borderRadius: 10, padding: "12px 8px", textAlign: "center" }}>
                  <div style={{ fontSize: 20, fontWeight: 800, color: s.color }}>{s.val.toLocaleString()}</div>
                  <div style={{ fontSize: 9, color: C.textLight, marginTop: 2, fontWeight: 600 }}>{s.label.toUpperCase()}</div>
                </div>
              ))}
            </div>

            {/* Match bar */}
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

            {/* Tab bar + actions */}
            <div style={{ display: "flex", gap: 5, marginBottom: 14, alignItems: "center", flexWrap: "wrap" }}>
              {[
                ["dashboard", "📊 Dashboard"],
                ["detail", "🔍 Detail Report"],
                ["sheet", "📋 Comparison Sheet"],
              ].map(([t, label]) => (
                <button key={t} onClick={() => setActiveTab(t)} style={S.btn(activeTab === t)}>{label}</button>
              ))}
              <div style={{ flex: 1 }} />
              <button onClick={handleExport} disabled={exporting}
                style={{ ...S.btn(false), background: exporting ? C.textLight : C.green, color: "#fff", border: "none", fontWeight: 700, padding: "7px 18px", cursor: exporting ? "wait" : "pointer" }}>
                {exporting ? "⏳ Building…" : `⬇ Export Excel (${sessions.length} session${sessions.length > 1 ? "s" : ""})`}
              </button>
              <button onClick={() => { setStep(0); setResults(null); }} style={{ ...S.btn(false), color: C.blue, borderColor: C.blue }}>+ New</button>
              <button onClick={() => setStep(1)} style={S.btn(false)}>← Remap</button>
            </div>

            {/* DASHBOARD TAB */}
            {activeTab === "dashboard" && (
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
                <div style={{ ...S.card, padding: 20 }}>
                  <div style={{ ...S.label, marginBottom: 14 }}>CLASSIFICATION BREAKDOWN</div>
                  <div style={{ display: "flex", alignItems: "center", gap: 20 }}>
                    <svg width="110" height="110" viewBox="0 0 110 110">
                      {(() => {
                        const segs = [{ val: stats.matched, c: C.green }, { val: stats.mismatched, c: C.red }, { val: stats.onlyA, c: C.amber }, { val: stats.onlyB, c: C.purple }];
                        const total = segs.reduce((s, x) => s + x.val, 0) || 1;
                        const r2 = 42, cx = 55, cy = 55; let angle = -Math.PI / 2;
                        return segs.map((seg, i) => {
                          if (!seg.val) return null;
                          const sweep = (seg.val / total) * 2 * Math.PI;
                          const x1 = cx + r2 * Math.cos(angle), y1 = cy + r2 * Math.sin(angle);
                          angle += sweep;
                          const x2 = cx + r2 * Math.cos(angle), y2 = cy + r2 * Math.sin(angle);
                          return <path key={i} d={`M${cx},${cy} L${x1},${y1} A${r2},${r2} 0 ${sweep > Math.PI ? 1 : 0} 1 ${x2},${y2} Z`} fill={seg.c} opacity={0.9} />;
                        });
                      })()}
                      <circle cx="55" cy="55" r="26" fill="white" />
                      <text x="55" y="59" textAnchor="middle" fill={C.text} fontSize="11" fontWeight="bold" fontFamily="Inter">
                        {((stats.matched / (stats.matched + stats.mismatched || 1)) * 100).toFixed(0)}%
                      </text>
                    </svg>
                    <div style={{ flex: 1 }}>
                      {[["Matched", stats.matched, C.green], ["Mismatched", stats.mismatched, C.red], ["Only A", stats.onlyA, C.amber], ["Only B", stats.onlyB, C.purple]].map(([l, v, c]) => (
                        <div key={l} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                          <div style={{ width: 8, height: 8, borderRadius: 2, background: c }} />
                          <span style={{ flex: 1, fontSize: 12, color: C.textMid }}>{l}</span>
                          <span style={{ fontWeight: 700, color: c }}>{v.toLocaleString()}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
                <div style={{ ...S.card, padding: 20 }}>
                  <div style={{ ...S.label, marginBottom: 14 }}>MISMATCHES BY FIELD</div>
                  {compareMappings.length === 0 && <div style={{ color: C.textLight, fontSize: 12 }}>No compare columns configured.</div>}
                  {compareMappings.slice(0, 8).map(m => {
                    const paired = results.results.filter(r => r.rowA && r.rowB).length || 1;
                    const mis = results.results.filter(r => r.details.some(d => d.colA === m.colA && d.status === "Mismatched")).length;
                    const pct = (mis / paired) * 100;
                    return (
                      <div key={m.colA} style={{ marginBottom: 11 }}>
                        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11, marginBottom: 4 }}>
                          <span style={{ color: C.textMid }}>{m.colA}</span>
                          <span style={{ color: pct > 10 ? C.red : C.green, fontWeight: 700 }}>{mis} ({pct.toFixed(1)}%)</span>
                        </div>
                        <div style={{ height: 5, background: "#E2E8F0", borderRadius: 99 }}>
                          <div style={{ width: `${Math.min(pct, 100)}%`, height: "100%", background: pct > 10 ? C.red : C.green, borderRadius: 99 }} />
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}

            {/* DETAIL REPORT TAB */}
            {activeTab === "detail" && (
              <div style={S.card}>
                <div style={{ padding: "11px 14px", borderBottom: `1px solid ${C.border}`, display: "flex", gap: 7, alignItems: "center", flexWrap: "wrap", background: "#F8FAFC" }}>
                  <input placeholder="Search key…" value={searchKey} onChange={e => setSearchKey(e.target.value)}
                    style={{ ...S.input, flex: 1, minWidth: 140 }} />
                  {["All", "Matched", "Mismatched", "Only in A", "Only in B"].map(s => (
                    <button key={s} onClick={() => setFilterStatus(s)} style={S.btn(filterStatus === s)}>{s}</button>
                  ))}
                  <span style={{ fontSize: 11, color: C.textLight, whiteSpace: "nowrap" }}>{filteredResults.length.toLocaleString()} rows</span>
                </div>
                <div style={{ overflowX: "auto", maxHeight: 500, overflowY: "auto" }}>
                  {filteredResults.length === 0
                    ? <div style={{ padding: 40, textAlign: "center", color: C.textLight }}>No records match the current filter.</div>
                    : (
                    <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 11 }}>
                      <thead style={{ position: "sticky", top: 0, zIndex: 1 }}>
                        <tr>
                          <th style={S.th}>Status</th>
                          <th style={S.th}>Composite Key</th>
                          {keyMappings.map(m => <th key={m.colA} style={{ ...S.th, color: C.amber }}>{m.colA}</th>)}
                          {compareMappings.map(m => (
                            <React.Fragment key={m.colA}>
                              <th style={{ ...S.th, color: C.blue, borderLeft: `2px solid ${C.border}` }}>{m.colA}</th>
                              <th style={{ ...S.th, color: C.purple }}>{m.colB}</th>
                              <th style={{ ...S.th, color: C.amber }}>Diff</th>
                              <th style={{ ...S.th, color: C.textLight }}>Comment</th>
                            </React.Fragment>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {filteredResults.slice(0, 300).map((r, i) => (
                          <tr key={i} style={{ background: i % 2 ? C.bg : C.surface }}>
                            <td style={S.td}><StatusBadge status={r.status} /></td>
                            <td style={{ ...S.td, color: C.textMid, maxWidth: 220, overflow: "hidden", textOverflow: "ellipsis" }}>{r.key}</td>
                            {keyMappings.map(m => <td key={m.colA} style={{ ...S.td, color: C.textMid }}>{r.keyVals[m.colA] ?? "—"}</td>)}
                            {compareMappings.map(m => {
                              const d = r.details.find(d => d.colA === m.colA);
                              const isMis = d?.status === "Mismatched";
                              const diff = d?.diff || "";
                              return (
                                <React.Fragment key={m.colA}>
                                  <td style={{ ...S.td, borderLeft: `2px solid ${C.border}`, background: isMis ? C.redLight : undefined, color: isMis ? C.red : C.blue, fontWeight: isMis ? 700 : 400 }}>
                                    {d ? d.valA : (r.rowA ? r.rowA[m.colA] ?? "—" : "—")}
                                  </td>
                                  <td style={{ ...S.td, background: isMis ? C.redLight : undefined, color: isMis ? C.purple : C.textMid, fontWeight: isMis ? 700 : 400 }}>
                                    {d ? d.valB : (r.rowB ? r.rowB[m.colB] ?? "—" : "—")}
                                  </td>
                                  <td style={{ ...S.td, color: isMis ? C.red : C.textLight, fontWeight: isMis ? 700 : 400 }}>{diff || "—"}</td>
                                  <td style={{ ...S.td, color: C.amber, fontStyle: "italic", fontSize: 10 }}>{makeComment(diff, r.status)}</td>
                                </React.Fragment>
                              );
                            })}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )}
                  {filteredResults.length > 300 && (
                    <div style={{ padding: 10, textAlign: "center", color: C.textLight, fontSize: 10, borderTop: `1px solid ${C.border}` }}>
                      Showing 300 of {filteredResults.length.toLocaleString()} rows — Export Excel for complete data.
                    </div>
                  )}
                </div>
              </div>
            )}

            {/* COMPARISON SHEET TAB */}
            {activeTab === "sheet" && <ComparisonSheetTab />}
          </div>
        )}
      </div>
    </div>
  );
}
