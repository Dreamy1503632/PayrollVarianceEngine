import { useState, useRef, useCallback } from "react";

/* ═══ PARSERS ═══ */
function parseCSV(text) {
  const lines = text.trim().split(/\r?\n/);
  if (!lines.length) return [];
  const hdrs = lines[0].split(",").map(h => h.trim().replace(/^"|"$/g, ""));
  return lines.slice(1).map(line => {
    const vals = line.match(/(".*?"|[^,]+|(?<=,)(?=,)|(?<=,)$|^(?=,))/g) || [];
    const row = {};
    hdrs.forEach((h, i) => { row[h] = (vals[i] || "").replace(/^"|"$/g, "").trim(); });
    return row;
  });
}
function parseTSV(text) {
  const lines = text.trim().split(/\r?\n/);
  if (!lines.length) return [];
  const hdrs = lines[0].split("\t").map(h => h.trim());
  return lines.slice(1).map(line => {
    const v = line.split("\t"); const row = {};
    hdrs.forEach((h, i) => { row[h] = (v[i] || "").trim(); }); return row;
  });
}
function parsePipe(text) {
  const lines = text.trim().split(/\r?\n/);
  if (!lines.length) return [];
  const hdrs = lines[0].split("|").map(h => h.trim());
  return lines.slice(1).map(line => {
    const v = line.split("|"); const row = {};
    hdrs.forEach((h, i) => { row[h] = (v[i] || "").trim(); }); return row;
  });
}
function autoDetect(text) {
  const f = text.split(/\r?\n/)[0] || "";
  if (f.includes("|")) return parsePipe(text);
  if (f.includes("\t")) return parseTSV(text);
  return parseCSV(text);
}
async function parseFile(file) {
  return new Promise((res, rej) => {
    const ext = file.name.split(".").pop().toLowerCase();
    const rdr = new FileReader();
    if (["xlsx", "xls"].includes(ext)) {
      rdr.onload = async e => {
        try {
          const { read, utils } = await import("https://cdn.sheetjs.com/xlsx-0.20.1/package/xlsx.mjs");
          const wb = read(new Uint8Array(e.target.result), { type: "array" });
          const data = utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });
          res(data.map(r => Object.fromEntries(Object.entries(r).map(([k, v]) => [String(k).trim(), String(v).trim()]))));
        } catch (err) { rej(err); }
      };
      rdr.readAsArrayBuffer(file);
    } else {
      rdr.onload = e => {
        const t = e.target.result;
        res(ext === "tsv" ? parseTSV(t) : autoDetect(t));
      };
      rdr.readAsText(file);
    }
  });
}

/* ═══ AUTO MAP ═══ */
function autoMapHeaders(ha, hb) {
  const norm = s => s.toLowerCase().replace(/[^a-z0-9]/g, "");
  const used = new Set();
  const keys = ["personnumber", "balancename", "area1", "area2", "area3"];
  return ha.map(a => {
    const na = norm(a); let best = "", bs = 0;
    for (const b of hb) { if (used.has(b)) continue; const nb = norm(b); const sc = na === nb ? 100 : (na.includes(nb) || nb.includes(na)) ? 75 : 0; if (sc > bs) { bs = sc; best = b; } }
    if (best && bs >= 75) used.add(best);
    const isKey = keys.some(k => na.includes(k));
    return { colA: a, colB: best && bs >= 75 ? best : "", isKey, compare: !isKey && !!(best && bs >= 75), ignoreCase: false };
  });
}

/* ═══ COMPARE ═══ */
function compare(dA, dB, maps, tol) {
  const tp = parseFloat(tol) / 100 || 0;
  const km = maps.filter(m => m.isKey && m.colA && m.colB);
  const cm = maps.filter(m => m.compare && !m.isKey && m.colA && m.colB);
  const mk = (row, a) => km.map(m => ((a ? row[m.colA] : row[m.colB]) || "").toString().toLowerCase().trim()).join("||");
  const idx = {};
  for (const r of dB) { const k = mk(r, false); (idx[k] = idx[k] || []).push(r); }
  const res = [], mb = new Set();
  for (const rA of dA) {
    const k = mk(rA, true), kv = Object.fromEntries(km.map(m => [m.colA, rA[m.colA] ?? ""]));
    const bs = idx[k] || [];
    if (!bs.length) { res.push({ key: k, rowA: rA, rowB: null, status: "Only in A", details: [], keyVals: kv }); continue; }
    const rB = bs[0]; mb.add(k);
    const det = cm.map(m => {
      const vA = rA[m.colA] ?? "", vB = rB[m.colB] ?? "";
      const a = m.ignoreCase ? vA.toLowerCase() : vA, b = m.ignoreCase ? vB.toLowerCase() : vB;
      const nA = parseFloat(vA), nB = parseFloat(vB), isN = !isNaN(nA) && !isNaN(nB);
      let d = "", st = "Matched";
      if (isN) { const p = nA !== 0 ? Math.abs(nA - nB) / Math.abs(nA) : nB !== 0 ? 1 : 0; d = (nB - nA).toFixed(4); if (p > tp) st = "Mismatched"; }
      else if (a.trim() !== b.trim()) { d = vA + "\u2192" + vB; st = "Mismatched"; }
      return { colA: m.colA, colB: m.colB, valA: vA, valB: vB, diff: d, status: st };
    });
    res.push({ key: k, rowA: rA, rowB: rB, status: det.some(d => d.status === "Mismatched") ? "Mismatched" : "Matched", details: det, keyVals: kv });
  }
  for (const r of dB) { const k = mk(r, false); if (!mb.has(k)) { res.push({ key: k, rowA: null, rowB: r, status: "Only in B", details: [], keyVals: Object.fromEntries(km.map(m => [m.colA, r[m.colB] ?? ""])) }); mb.add(k); } }
  const kc = {}; for (const r of dA) { const k = mk(r, true); kc[k] = (kc[k] || 0) + 1; }
  return { results: res, dupes: Object.values(kc).filter(c => c > 1).reduce((s, c) => s + c - 1, 0) };
}

function skComment(diff, st) {
  if (st === "Only in A") return "Only in File A";
  if (st === "Only in B") return "Only in File B";
  if (!diff || diff === "0.0000") return "";
  const n = parseFloat(diff);
  if (!isNaN(n)) return Math.abs(n) < 1 ? "Less than $1 difference" : "More $ Difference";
  return "Value mismatch";
}

/* ═══ XLSX EXPORT — async chunked, no crash ═══ */
function yieldUI() { return new Promise(r => setTimeout(r, 0)); }

async function exportXLSX(sessions, onMsg) {
  onMsg("Preparing data...");
  await yieldUI();

  const te = new TextEncoder();
  const esc = s => String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const cn = n => { let s = ""; for (; n >= 0; n = Math.floor(n / 26) - 1) s = String.fromCharCode(65 + n % 26) + s; return s; };

  const s0 = sessions[0];
  const cmp0 = s0.mappings.filter(m => m.compare && !m.isKey && m.colA && m.colB);
  const fr = s0.results.find(r => r.rowA)?.rowA || {};
  const allC = Object.keys(fr);
  const cmpSet = new Set(cmp0.map(m => m.colA));
  const shC = allC.filter(c => !cmpSet.has(c));

  // Build rows as simple arrays: [cells[], styleIndex]
  // Styles: 0=normal 1=header 2=mismatch 3=onlyA 4=onlyB 5=matched
  const stIdx = st => st === "Mismatched" ? 2 : st === "Only in A" ? 3 : st === "Only in B" ? 4 : st === "Matched" ? 5 : 0;

  // Sheet 1: Summary
  const sum = [
    [["CompareIQ Summary"], 1], [[], 0],
    [["#", "File A", "File B", "Total A", "Total B", "Matched", "Mismatched", "Only A", "Only B", "Dupes", "Match%"], 1],
    ...sessions.map((s, i) => [[i + 1, s.fileAName, s.fileBName, s.totalA, s.totalB, s.matched, s.mismatched, s.onlyA, s.onlyB, s.dupes, ((s.matched / (s.matched + s.mismatched || 1)) * 100).toFixed(1) + "%"], 0])
  ];

  onMsg("Building comparison results...");
  await yieldUI();

  // Sheet 2: Comparison Results (non-matched only, 2 rows per record)
  const compH = ["Status", "Source", "Key", ...allC];
  const comp = [[compH, 1]];
  for (const s of sessions) {
    const km = s.mappings.filter(m => m.isKey && m.colA && m.colB);
    const bm = Object.fromEntries(s.mappings.filter(m => m.colA && m.colB).map(m => [m.colA, m.colB]));
    for (const r of s.results) {
      if (r.status === "Matched") continue;
      const k = km.map(m => r.keyVals[m.colA] ?? "").join("|");
      const si = stIdx(r.status);
      comp.push([[r.status, s.fileAName, k, ...allC.map(c => r.rowA ? (r.rowA[c] ?? "") : "")], si]);
      comp.push([[r.status, s.fileBName, k, ...allC.map(c => { const cb = bm[c]; return r.rowB && cb ? (r.rowB[cb] ?? "") : ""; })], si]);
    }
  }

  onMsg("Building difference mismatch...");
  await yieldUI();

  // Sheet 3: Difference Mismatch (SK USOPTE format)
  const diffH = ["Status", "Source", "Key", ...shC, ...cmp0.flatMap(m => [m.colA, m.colA + " (USOPTE)", "Difference", "SK Comment"])];
  const diff = [[diffH, 1]];
  for (const s of sessions) {
    const km = s.mappings.filter(m => m.isKey && m.colA && m.colB);
    const sc = s.mappings.filter(m => m.compare && !m.isKey && m.colA && m.colB);
    const bm = Object.fromEntries(s.mappings.filter(m => m.colA && m.colB).map(m => [m.colA, m.colB]));
    for (const r of s.results) {
      if (r.status === "Matched") continue;
      const k = km.map(m => r.keyVals[m.colA] ?? "").join("|");
      const si = stIdx(r.status);
      // Row A
      const shA = shC.map(c => r.rowA ? (r.rowA[c] ?? "") : (r.rowB && bm[c] ? (r.rowB[bm[c]] ?? "") : ""));
      const cA = cmp0.flatMap(m => {
        const f = sc.find(x => x.colA === m.colA);
        const vA = f && r.rowA ? (r.rowA[f.colA] ?? "") : "";
        const d = r.details?.find(x => x.colA === m.colA);
        const df = d?.diff || "";
        return [vA, "", df ? (parseFloat(df) || df) : "", skComment(df, r.status)];
      });
      diff.push([[r.status, s.fileAName, k, ...shA, ...cA], si]);
      // Row B
      const shB = shC.map(c => { const cb = bm[c]; return r.rowB && cb ? (r.rowB[cb] ?? "") : ""; });
      const cB = cmp0.flatMap(m => {
        const f = sc.find(x => x.colA === m.colA);
        const vB = f && r.rowB ? (r.rowB[f.colB] ?? "") : "";
        return ["", vB, "", ""];
      });
      diff.push([[r.status, s.fileBName, k, ...shB, ...cB], si]);
    }
  }

  onMsg("Encoding sheets...");
  await yieldUI();

  // Build sheet XML async — yield every 500 rows
  async function buildSheet(rows) {
    const parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'];
    for (let r = 0; r < rows.length; r++) {
      const [cells, si] = rows[r];
      let x = '<row r="' + (r + 1) + '">';
      for (let c = 0; c < cells.length; c++) {
        const v = cells[c], a = cn(c) + (r + 1), sv = String(v ?? "");
        const isN = sv !== "" && sv.trim() !== "" && !isNaN(Number(sv)) && !/[%]/.test(sv);
        x += isN ? '<c r="' + a + '" s="' + si + '"><v>' + sv + '</v></c>'
          : '<c r="' + a + '" s="' + si + '" t="inlineStr"><is><t>' + esc(v) + '</t></is></c>';
      }
      x += '</row>';
      parts.push(x);
      if (r % 500 === 0) await yieldUI();
    }
    parts.push('</sheetData></worksheet>');
    return te.encode(parts.join(""));
  }

  const sheet1 = await buildSheet(sum);
  onMsg("Encoding sheet 2/" + 3 + "...");
  const sheet2 = await buildSheet(comp);
  onMsg("Encoding sheet 3/" + 3 + "...");
  const sheet3 = await buildSheet(diff);

  onMsg("Creating ZIP file...");
  await yieldUI();

  const sheets = [
    { name: "Summary", data: sheet1 },
    { name: "Comparison Results", data: sheet2 },
    { name: "Difference Mismatch", data: sheet3 },
  ];

  const styXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts><fills count="8"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1A2332"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFEF2F2"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFBEB"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF5F3FF"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFECFDF5"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="6"><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/><xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1"/><xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1" applyBorder="1"/><xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0" applyFill="1" applyBorder="1"/><xf numFmtId="0" fontId="0" fillId="6" borderId="1" xfId="0" applyFill="1" applyBorder="1"/></cellXfs></styleSheet>';

  const wbXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + sheets.map((s, i) => '<sheet name="' + esc(s.name) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>').join("") + '</sheets></workbook>';
  const wbRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + sheets.map((s, i) => '<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>').join("") + '<Relationship Id="rId' + (sheets.length + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>';
  const ctXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' + sheets.map((s, i) => '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>').join("") + '</Types>';
  const rrXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';

  // ZIP — single pre-sized buffer, no spread, no crash
  const enc = s => te.encode(s);
  const files = [
    { n: "[Content_Types].xml", d: enc(ctXml) },
    { n: "_rels/.rels", d: enc(rrXml) },
    { n: "xl/workbook.xml", d: enc(wbXml) },
    { n: "xl/_rels/workbook.xml.rels", d: enc(wbRels) },
    { n: "xl/styles.xml", d: enc(styXml) },
    ...sheets.map((s, i) => ({ n: "xl/worksheets/sheet" + (i + 1) + ".xml", d: s.data })),
  ];

  // CRC32
  const crcT = new Uint32Array(256);
  for (let i = 0; i < 256; i++) { let c = i; for (let j = 0; j < 8; j++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; crcT[i] = c; }
  const crc32 = d => { let c = ~0; for (let i = 0; i < d.length; i++) c = crcT[(c ^ d[i]) & 0xFF] ^ (c >>> 8); return (~c) >>> 0; };

  // Pre-calculate total ZIP size
  let totalSz = 22; // EOCD
  const entries = files.map(f => {
    const nb = enc(f.n); const crc = crc32(f.d);
    totalSz += 30 + nb.length + f.d.length + 46 + nb.length;
    return { nb, d: f.d, crc };
  });

  const buf = new Uint8Array(totalSz);
  const w16 = (o, v) => { buf[o] = v & 0xFF; buf[o + 1] = (v >> 8) & 0xFF; };
  const w32 = (o, v) => { buf[o] = v & 0xFF; buf[o + 1] = (v >> 8) & 0xFF; buf[o + 2] = (v >> 16) & 0xFF; buf[o + 3] = (v >> 24) & 0xFF; };

  let lo = 0;
  const offsets = [];
  const cdStart = entries.reduce((s, e) => s + 30 + e.nb.length + e.d.length, 0);
  let co = cdStart;

  for (let i = 0; i < entries.length; i++) {
    const e = entries[i];
    offsets.push(lo);
    // Local header
    buf[lo] = 0x50; buf[lo + 1] = 0x4B; buf[lo + 2] = 3; buf[lo + 3] = 4;
    w16(lo + 4, 20); w32(lo + 14, e.crc); w32(lo + 18, e.d.length); w32(lo + 22, e.d.length);
    w16(lo + 26, e.nb.length);
    buf.set(e.nb, lo + 30);
    buf.set(e.d, lo + 30 + e.nb.length);
    lo += 30 + e.nb.length + e.d.length;
    // Central dir
    buf[co] = 0x50; buf[co + 1] = 0x4B; buf[co + 2] = 1; buf[co + 3] = 2;
    w16(co + 4, 20); w16(co + 6, 20); w32(co + 16, e.crc); w32(co + 20, e.d.length); w32(co + 24, e.d.length);
    w16(co + 28, e.nb.length); w32(co + 42, offsets[i]);
    buf.set(e.nb, co + 46);
    co += 46 + e.nb.length;
  }

  // EOCD
  buf[co] = 0x50; buf[co + 1] = 0x4B; buf[co + 2] = 5; buf[co + 3] = 6;
  w16(co + 8, entries.length); w16(co + 10, entries.length);
  w32(co + 12, co - cdStart); w32(co + 16, cdStart);

  onMsg("Downloading...");
  const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a"); a.href = url; a.download = "CompareIQ_Results.xlsx";
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

/* ═══ STYLES ═══ */
const C = {
  bg: "#F7F9FC", sf: "#FFFFFF", bd: "#D1D9E6", bs: "#A0B0C8",
  tx: "#1A2332", tm: "#4A5568", tl: "#718096",
  bl: "#1A56DB", bll: "#EBF5FF", blm: "#3B82F6",
  pu: "#7C3AED", pul: "#F5F3FF",
  gn: "#059669", gnl: "#ECFDF5",
  rd: "#DC2626", rdl: "#FEF2F2",
  am: "#D97706", aml: "#FFFBEB",
  hbg: "#1A2332",
};

const sCard = { background: C.sf, borderRadius: 12, border: `1px solid ${C.bd}`, overflow: "hidden", boxShadow: "0 1px 4px rgba(0,0,0,.06)" };
const sTh = { padding: "10px 13px", textAlign: "left", borderBottom: `1px solid ${C.bd}`, whiteSpace: "nowrap", fontSize: 11, fontWeight: 700, color: C.tm, background: "#F0F4FA" };
const sTd = { padding: "7px 12px", borderBottom: `1px solid ${C.bd}20`, whiteSpace: "nowrap", fontSize: 11, color: C.tx };
const sInp = { background: C.sf, border: `1px solid ${C.bd}`, borderRadius: 7, padding: "7px 12px", color: C.tx, fontSize: 12, fontFamily: "inherit", outline: "none" };
const sBtn = (on, cl = C.bl) => ({ background: on ? cl : C.sf, border: `1px solid ${on ? cl : C.bd}`, color: on ? "#fff" : C.tm, borderRadius: 7, padding: "7px 14px", cursor: "pointer", fontSize: 12, fontWeight: 600, fontFamily: "inherit" });
const sBtnP = { background: `linear-gradient(135deg,${C.bl},${C.pu})`, color: "#fff", border: "none", borderRadius: 9, padding: "11px 32px", fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" };
const sLbl = { fontSize: 11, fontWeight: 700, color: C.tl, letterSpacing: ".05em" };

const Badge = ({ s }) => {
  const m = { Matched: [C.gn, C.gnl], Mismatched: [C.rd, C.rdl], "Only in A": [C.am, C.aml], "Only in B": [C.pu, C.pul] };
  const [c, b] = m[s] || [C.tl, C.bg];
  return <span style={{ background: b, color: c, border: `1px solid ${c}33`, borderRadius: 5, padding: "2px 9px", fontSize: 10, fontWeight: 700 }}>{s}</span>;
};

export default function App() {
  const [step, setStep] = useState(0);
  const [fA, setFA] = useState(null), [fB, setFB] = useState(null);
  const [dA, setDA] = useState([]), [dB, setDB] = useState([]);
  const [hA, setHA] = useState([]), [hB, setHB] = useState([]);
  const [maps, setMaps] = useState([]);
  const [tol, setTol] = useState("1");
  const [res, setRes] = useState(null);
  const [loading, setLoading] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [expMsg, setExpMsg] = useState("");
  const [err, setErr] = useState("");
  const [tab, setTab] = useState("dashboard");
  const [filt, setFilt] = useState("All");
  const [search, setSearch] = useState("");
  const [sessions, setSessions] = useState([]);

  const doExport = useCallback(async () => {
    if (!sessions.length) return;
    setExporting(true); setErr("");
    try { await exportXLSX(sessions, setExpMsg); } catch (e) { setErr("Export failed: " + e.message); }
    setExporting(false); setExpMsg("");
  }, [sessions]);

  const loadFile = async (file, w) => {
    setErr(""); setLoading(true);
    try {
      const d = await parseFile(file);
      if (!d.length) throw new Error("No data");
      const h = Object.keys(d[0]);
      if (w === "A") { setFA(file); setDA(d); setHA(h); } else { setFB(file); setDB(d); setHB(h); }
    } catch (e) { setErr(e.message); }
    setLoading(false);
  };

  const goMap = () => { if (!dA.length || !dB.length) { setErr("Upload both files"); return; } setMaps(autoMapHeaders(hA, hB)); setStep(1); };
  const upMap = (i, f, v) => setMaps(p => p.map((m, j) => j === i ? { ...m, [f]: v } : m));

  const run = () => {
    if (!maps.some(m => m.isKey)) { setErr("Mark at least one KEY"); return; }
    setLoading(true);
    setTimeout(() => {
      const r = compare(dA, dB, maps, tol); setRes(r);
      const ct = s => r.results.filter(x => x.status === s).length;
      setSessions(p => [...p, { fileAName: fA?.name || "A", fileBName: fB?.name || "B", results: r.results, totalA: dA.length, totalB: dB.length, matched: ct("Matched"), mismatched: ct("Mismatched"), onlyA: ct("Only in A"), onlyB: ct("Only in B"), dupes: r.dupes, mappings: [...maps] }]);
      setStep(2); setLoading(false);
    }, 100);
  };

  const st = res ? {
    tA: dA.length, tB: dB.length,
    ma: res.results.filter(r => r.status === "Matched").length,
    mi: res.results.filter(r => r.status === "Mismatched").length,
    oA: res.results.filter(r => r.status === "Only in A").length,
    oB: res.results.filter(r => r.status === "Only in B").length,
    du: res.dupes,
  } : null;

  const kMaps = maps.filter(m => m.isKey && m.colA && m.colB);
  const cMaps = maps.filter(m => m.compare && !m.isKey && m.colA && m.colB);
  const filtered = res ? res.results.filter(r => (filt === "All" || r.status === filt) && (!search || r.key.toLowerCase().includes(search.toLowerCase()))) : [];

  const Drop = ({ label, file, onFile, color }) => {
    const ref = useRef(), [dr, setDr] = useState(false);
    return (
      <div onClick={() => ref.current.click()} onDragOver={e => { e.preventDefault(); setDr(true); }} onDragLeave={() => setDr(false)}
        onDrop={e => { e.preventDefault(); setDr(false); e.dataTransfer.files[0] && onFile(e.dataTransfer.files[0]); }}
        style={{ border: `2px dashed ${dr ? color : file ? C.gn : C.bd}`, borderRadius: 12, padding: "22px 16px", textAlign: "center", cursor: "pointer", background: dr ? color + "08" : file ? C.gnl : C.bg }}>
        <input ref={ref} type="file" accept=".csv,.xlsx,.xls,.txt,.tsv" style={{ display: "none" }} onChange={e => e.target.files[0] && onFile(e.target.files[0])} />
        <div style={{ fontSize: 26, marginBottom: 6 }}>{file ? "\u2705" : "\uD83D\uDCC1"}</div>
        <div style={{ fontWeight: 700, color: file ? C.gn : C.tm, fontSize: 13 }}>{file ? file.name : label}</div>
        <div style={{ color: C.tl, fontSize: 11, marginTop: 4 }}>{file ? (file.size / 1024).toFixed(1) + " KB" : "CSV, Excel, TXT, TSV"}</div>
      </div>
    );
  };

  return (
    <div style={{ minHeight: "100vh", background: C.bg, color: C.tx, fontFamily: "'Segoe UI',system-ui,sans-serif" }}>
      {/* HEADER */}
      <div style={{ background: C.hbg, padding: "12px 28px", display: "flex", alignItems: "center", justifyContent: "space-between", position: "sticky", top: 0, zIndex: 50 }}>
        <div style={{ fontWeight: 800, fontSize: 18, color: "#fff" }}>CompareIQ</div>
        <div style={{ display: "flex", gap: 6 }}>
          {["Upload", "Map", "Results"].map((s, i) => (
            <div key={s} style={{ display: "flex", alignItems: "center", gap: 4 }}>
              <span onClick={() => step > i && setStep(i)} style={{ padding: "4px 14px", borderRadius: 20, background: step === i ? "#3B82F6" : step > i ? "#1E40AF" : "transparent", border: `1px solid ${step >= i ? "#3B82F6" : "#475569"}`, color: step >= i ? "#fff" : "#94A3B8", fontSize: 11, fontWeight: 600, cursor: step > i ? "pointer" : "default" }}>{s}</span>
              {i < 2 && <span style={{ color: "#475569" }}>{"\u203A"}</span>}
            </div>
          ))}
        </div>
      </div>

      <div style={{ maxWidth: 1100, margin: "0 auto", padding: "24px 20px" }}>
        {err && <div style={{ background: C.rdl, border: `1px solid ${C.rd}33`, color: C.rd, borderRadius: 9, padding: "9px 14px", marginBottom: 14, fontSize: 12 }}>{err}</div>}
        {loading && <div style={{ background: C.bll, borderRadius: 9, padding: "9px 14px", marginBottom: 14, fontSize: 12, color: C.bl }}>Processing...</div>}
        {exporting && <div style={{ background: C.aml, borderRadius: 9, padding: "9px 14px", marginBottom: 14, fontSize: 12, color: C.am, fontWeight: 600 }}>{expMsg || "Exporting..."}</div>}

        {/* ═══ STEP 0 ═══ */}
        {step === 0 && (
          <div>
            <h1 style={{ fontWeight: 800, fontSize: 26, textAlign: "center", margin: "0 0 20px" }}>Upload Datasets</h1>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 24 }}>
              {[{ l: "Dataset A (Source)", w: "A", f: fA, d: dA, h: hA, c: C.bl }, { l: "Dataset B (Target)", w: "B", f: fB, d: dB, h: hB, c: C.pu }].map(x => (
                <div key={x.w}>
                  <div style={{ ...sLbl, color: x.c, marginBottom: 8 }}>{x.w}</div>
                  <Drop label={x.l} file={x.f} color={x.c} onFile={f => loadFile(f, x.w)} />
                  {x.d.length > 0 && <div style={{ marginTop: 6, fontSize: 11, color: C.gn }}>{x.d.length} rows, {x.h.length} cols</div>}
                </div>
              ))}
            </div>
            <div style={{ textAlign: "center" }}>
              <button onClick={goMap} disabled={!dA.length || !dB.length} style={{ ...sBtnP, opacity: dA.length && dB.length ? 1 : .4 }}>Map Columns {"\u2192"}</button>
            </div>
          </div>
        )}

        {/* ═══ STEP 1 ═══ */}
        {step === 1 && (
          <div>
            <h1 style={{ fontWeight: 800, fontSize: 22, margin: "0 0 12px" }}>Column Mapping</h1>
            <div style={{ display: "flex", gap: 10, marginBottom: 12, alignItems: "center" }}>
              <span style={sLbl}>Tolerance</span>
              <input type="number" value={tol} onChange={e => setTol(e.target.value)} style={{ ...sInp, width: 60, textAlign: "center" }} />
              <span style={{ color: C.gn, fontWeight: 700 }}>%</span>
              <div style={{ flex: 1 }} />
              <button onClick={() => setMaps(p => p.map(m => m.isKey ? m : { ...m, compare: true }))} style={sBtn(false, C.gn)}>All Compare</button>
              <button onClick={() => setMaps(p => p.map(m => m.isKey ? m : { ...m, compare: false }))} style={sBtn(false, C.rd)}>None</button>
            </div>
            <div style={sCard}>
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead>
                  <tr>
                    <th style={sTh}>File A</th><th style={sTh}>File B</th>
                    <th style={{ ...sTh, textAlign: "center", width: 50 }}>Key</th>
                    <th style={{ ...sTh, textAlign: "center", width: 70 }}>Compare</th>
                    <th style={{ ...sTh, textAlign: "center", width: 80 }}>Ign. Case</th>
                  </tr>
                </thead>
                <tbody>
                  {maps.map((m, i) => (
                    <tr key={i} style={{ background: m.isKey ? C.aml : i % 2 ? C.sf : C.bg }}>
                      <td style={sTd}><b>{m.colA}</b></td>
                      <td style={sTd}>
                        <select value={m.colB} onChange={e => upMap(i, "colB", e.target.value)} style={{ ...sInp, width: "100%" }}>
                          <option value="">-- skip --</option>
                          {hB.map(h => <option key={h} value={h}>{h}</option>)}
                        </select>
                      </td>
                      <td style={{ ...sTd, textAlign: "center" }}><input type="checkbox" checked={m.isKey} onChange={() => upMap(i, "isKey", !m.isKey)} /></td>
                      <td style={{ ...sTd, textAlign: "center" }}><input type="checkbox" checked={m.compare} disabled={m.isKey} onChange={() => !m.isKey && upMap(i, "compare", !m.compare)} /></td>
                      <td style={{ ...sTd, textAlign: "center" }}><input type="checkbox" checked={m.ignoreCase} onChange={() => upMap(i, "ignoreCase", !m.ignoreCase)} /></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div style={{ display: "flex", gap: 10, justifyContent: "center", marginTop: 16 }}>
              <button onClick={() => setStep(0)} style={sBtn(false)}>{"\u2190"} Back</button>
              <button onClick={run} style={sBtnP}>Run Comparison</button>
            </div>
          </div>
        )}

        {/* ═══ STEP 2 ═══ */}
        {step === 2 && st && (
          <div>
            {/* Stats */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(7,1fr)", gap: 8, marginBottom: 16 }}>
              {[["Total A", st.tA, C.bl, C.bll], ["Total B", st.tB, C.pu, C.pul], ["Matched", st.ma, C.gn, C.gnl], ["Mismatched", st.mi, C.rd, C.rdl], ["Only A", st.oA, C.am, C.aml], ["Only B", st.oB, C.pu, C.pul], ["Dupes", st.du, C.tm, "#F1F5F9"]].map(([l, v, c, b]) => (
                <div key={l} style={{ background: b, borderTop: `3px solid ${c}`, borderRadius: 10, padding: "12px 8px", textAlign: "center", border: `1px solid ${c}30` }}>
                  <div style={{ fontSize: 20, fontWeight: 800, color: c }}>{v}</div>
                  <div style={{ fontSize: 9, color: C.tl, fontWeight: 600, marginTop: 2 }}>{l.toUpperCase()}</div>
                </div>
              ))}
            </div>

            {/* Match bar */}
            <div style={{ ...sCard, padding: "12px 16px", marginBottom: 14 }}>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11, marginBottom: 6 }}>
                <span style={{ color: C.tm }}>Match Rate</span>
                <span style={{ color: C.gn, fontWeight: 700 }}>{((st.ma / (st.ma + st.mi || 1)) * 100).toFixed(1)}%</span>
              </div>
              <div style={{ height: 7, background: "#E2E8F0", borderRadius: 99, display: "flex", overflow: "hidden" }}>
                {[[st.ma, C.gn], [st.mi, C.rd], [st.oA, C.am], [st.oB, C.pu]].map(([w, c], i) => {
                  const t = st.ma + st.mi + st.oA + st.oB || 1;
                  return <div key={i} style={{ width: `${(w / t) * 100}%`, background: c }} />;
                })}
              </div>
            </div>

            {/* Tab bar */}
            <div style={{ display: "flex", gap: 5, marginBottom: 14, flexWrap: "wrap", alignItems: "center" }}>
              {[["dashboard", "Dashboard"], ["detail", "Detail Report"], ["sheet", "Comparison Sheet"]].map(([k, l]) => (
                <button key={k} onClick={() => setTab(k)} style={sBtn(tab === k)}>{l}</button>
              ))}
              <div style={{ flex: 1 }} />
              <button onClick={doExport} disabled={exporting} style={{ ...sBtn(false), background: C.gn, color: "#fff", border: "none", fontWeight: 700 }}>{exporting ? "Exporting..." : "Export Excel"}</button>
              <button onClick={() => { setStep(0); setRes(null); }} style={{ ...sBtn(false), color: C.bl, borderColor: C.bl }}>+ New</button>
              <button onClick={() => setStep(1)} style={sBtn(false)}>{"\u2190"} Remap</button>
            </div>

            {/* ─── DASHBOARD ─── */}
            {tab === "dashboard" && (
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
                <div style={{ ...sCard, padding: 20 }}>
                  <div style={{ ...sLbl, marginBottom: 14 }}>BREAKDOWN</div>
                  <div style={{ display: "flex", alignItems: "center", gap: 20 }}>
                    <svg width="110" height="110" viewBox="0 0 110 110">
                      {(() => {
                        const sg = [{ v: st.ma, c: C.gn }, { v: st.mi, c: C.rd }, { v: st.oA, c: C.am }, { v: st.oB, c: C.pu }];
                        const tot = sg.reduce((s, x) => s + x.v, 0) || 1;
                        let a = -Math.PI / 2;
                        return sg.map((s, i) => {
                          if (!s.v) return null;
                          const sw = (s.v / tot) * 2 * Math.PI;
                          const x1 = 55 + 42 * Math.cos(a), y1 = 55 + 42 * Math.sin(a);
                          a += sw;
                          return <path key={i} d={`M55,55 L${x1},${y1} A42,42 0 ${sw > Math.PI ? 1 : 0} 1 ${55 + 42 * Math.cos(a)},${55 + 42 * Math.sin(a)} Z`} fill={s.c} opacity={.9} />;
                        });
                      })()}
                      <circle cx="55" cy="55" r="26" fill="white" />
                      <text x="55" y="59" textAnchor="middle" fill={C.tx} fontSize="11" fontWeight="bold">{((st.ma / (st.ma + st.mi || 1)) * 100).toFixed(0)}%</text>
                    </svg>
                    <div style={{ flex: 1 }}>
                      {[["Matched", st.ma, C.gn], ["Mismatched", st.mi, C.rd], ["Only A", st.oA, C.am], ["Only B", st.oB, C.pu]].map(([l, v, c]) => (
                        <div key={l} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                          <div style={{ width: 8, height: 8, borderRadius: 2, background: c }} />
                          <span style={{ flex: 1, fontSize: 12, color: C.tm }}>{l}</span>
                          <span style={{ fontWeight: 700, color: c }}>{v}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
                <div style={{ ...sCard, padding: 20 }}>
                  <div style={{ ...sLbl, marginBottom: 14 }}>MISMATCHES BY FIELD</div>
                  {cMaps.slice(0, 8).map(m => {
                    const tot = res.results.filter(r => r.rowA && r.rowB).length || 1;
                    const mis = res.results.filter(r => r.details.some(d => d.colA === m.colA && d.status === "Mismatched")).length;
                    const pct = (mis / tot) * 100;
                    return (
                      <div key={m.colA} style={{ marginBottom: 10 }}>
                        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11, marginBottom: 3 }}>
                          <span>{m.colA}</span>
                          <span style={{ color: pct > 10 ? C.rd : C.gn, fontWeight: 700 }}>{mis} ({pct.toFixed(1)}%)</span>
                        </div>
                        <div style={{ height: 5, background: "#E2E8F0", borderRadius: 99 }}>
                          <div style={{ width: `${Math.min(pct, 100)}%`, height: "100%", background: pct > 10 ? C.rd : C.gn, borderRadius: 99 }} />
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}

            {/* ─── DETAIL REPORT ─── */}
            {tab === "detail" && (
              <div style={sCard}>
                <div style={{ padding: "10px 14px", borderBottom: `1px solid ${C.bd}`, display: "flex", gap: 7, flexWrap: "wrap", alignItems: "center", background: "#F8FAFC" }}>
                  <input placeholder="Search key..." value={search} onChange={e => setSearch(e.target.value)} style={{ ...sInp, flex: 1, minWidth: 120 }} />
                  {["All", "Matched", "Mismatched", "Only in A", "Only in B"].map(s => (
                    <button key={s} onClick={() => setFilt(s)} style={sBtn(filt === s)}>{s}</button>
                  ))}
                  <span style={{ fontSize: 11, color: C.tl }}>{filtered.length}</span>
                </div>
                <div style={{ overflowX: "auto", maxHeight: 450, overflowY: "auto" }}>
                  <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 11 }}>
                    <thead style={{ position: "sticky", top: 0, zIndex: 1 }}>
                      <tr>
                        <th style={sTh}>Status</th>
                        {kMaps.map(m => <th key={m.colA} style={{ ...sTh, color: C.am }}>{m.colA}</th>)}
                        {cMaps.map(m => (
                          [<th key={m.colA + "a"} style={{ ...sTh, color: C.bl, borderLeft: `2px solid ${C.bd}` }}>Current ({m.colA})</th>,
                          <th key={m.colA + "b"} style={{ ...sTh, color: C.pu }}>USOPTE ({m.colB})</th>,
                          <th key={m.colA + "d"} style={{ ...sTh, color: C.am }}>Diff</th>,
                          <th key={m.colA + "c"} style={{ ...sTh, color: C.tl }}>SK Comment</th>]
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {filtered.slice(0, 300).map((r, i) => (
                        <tr key={i} style={{ background: r.status === "Mismatched" ? C.rdl : r.status === "Only in A" ? C.aml : r.status === "Only in B" ? C.pul : i % 2 ? C.bg : C.sf }}>
                          <td style={sTd}><Badge s={r.status} /></td>
                          {kMaps.map(m => <td key={m.colA} style={sTd}>{r.keyVals[m.colA] ?? ""}</td>)}
                          {cMaps.map(m => {
                            const d = r.details.find(x => x.colA === m.colA);
                            const mis = d?.status === "Mismatched";
                            const vA = d?.valA ?? (r.rowA?.[m.colA] ?? "");
                            const vB = d?.valB ?? (r.rowB?.[m.colB] ?? "");
                            const df = d?.diff || "";
                            const cm = skComment(df, r.status);
                            return [
                              <td key={m.colA + "a"} style={{ ...sTd, color: C.bl, borderLeft: `2px solid ${C.bd}`, background: mis ? C.rdl : "transparent" }}>{vA || "\u2014"}</td>,
                              <td key={m.colA + "b"} style={{ ...sTd, color: C.pu, background: mis ? C.rdl : "transparent" }}>{vB || "\u2014"}</td>,
                              <td key={m.colA + "d"} style={{ ...sTd, color: mis ? C.rd : C.tl, fontWeight: mis ? 700 : 400 }}>{df || "\u2014"}</td>,
                              <td key={m.colA + "c"} style={{ ...sTd, color: C.am, fontStyle: "italic", fontSize: 10 }}>{cm || "\u2014"}</td>,
                            ];
                          })}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {filtered.length === 0 && <div style={{ padding: 30, textAlign: "center", color: C.tl }}>No records match.</div>}
                  {filtered.length > 300 && <div style={{ padding: 8, textAlign: "center", color: C.tl, fontSize: 10 }}>Showing 300 of {filtered.length}. Export for all.</div>}
                </div>
              </div>
            )}

            {/* ─── COMPARISON SHEET ─── */}
            {tab === "sheet" && (() => {
              const ls = sessions[sessions.length - 1];
              if (!ls) return <div style={{ color: C.tl }}>No session data.</div>;
              const scm = ls.mappings.filter(m => m.compare && !m.isKey && m.colA && m.colB);
              const fr = ls.results.find(r => r.rowA)?.rowA || {};
              const ac = Object.keys(fr);
              const cs = new Set(scm.map(m => m.colA));
              const sc = ac.filter(c => !cs.has(c));

              return (
                <div>
                  <div style={{ background: C.bll, borderRadius: 9, padding: "10px 14px", marginBottom: 10, fontSize: 11, color: C.tm }}>
                    2 rows per record: Row 1 = File A (Current), Row 2 = File B (USOPTE). Showing first 200 records.
                  </div>
                  <div style={sCard}>
                    <div style={{ overflowX: "auto", maxHeight: 520, overflowY: "auto" }}>
                      <table style={{ borderCollapse: "collapse", fontSize: 10 }}>
                        <thead style={{ position: "sticky", top: 0, zIndex: 2 }}>
                          <tr>
                            <th style={{ ...sTh, background: C.hbg, color: "#fff" }}>Status</th>
                            <th style={{ ...sTh, background: C.hbg, color: "#fff" }}>Source</th>
                            <th style={{ ...sTh, background: C.hbg, color: "#fff" }}>Key</th>
                            {sc.map(c => <th key={c} style={{ ...sTh, background: C.hbg, color: "#CBD5E1" }}>{c}</th>)}
                            {scm.map(m => [
                              <th key={m.colA + "c"} style={{ ...sTh, background: "#7F1D1D", color: "#FCA5A5", borderLeft: "2px solid #991B1B" }}>{m.colA}</th>,
                              <th key={m.colA + "u"} style={{ ...sTh, background: "#7F1D1D", color: "#FCA5A5" }}>{m.colB} (USOPTE)</th>,
                              <th key={m.colA + "d"} style={{ ...sTh, background: "#7F1D1D", color: "#FCA5A5" }}>Difference</th>,
                              <th key={m.colA + "k"} style={{ ...sTh, background: "#7F1D1D", color: "#FCA5A5", borderRight: "2px solid #991B1B" }}>SK Comment</th>,
                            ])}
                          </tr>
                        </thead>
                        <tbody>
                          {ls.results.slice(0, 200).flatMap((r, ri) => {
                            const km = ls.mappings.filter(m => m.isKey && m.colA && m.colB);
                            const sm = ls.mappings.filter(m => m.compare && !m.isKey && m.colA && m.colB);
                            const bm = Object.fromEntries(ls.mappings.filter(m => m.colA && m.colB).map(m => [m.colA, m.colB]));
                            const key = km.map(m => r.keyVals[m.colA] ?? "").join("|");
                            const isMis = r.status !== "Matched";
                            const bgA = isMis ? "#FFF5F5" : ri % 2 === 0 ? C.sf : C.bg;
                            const bgB = isMis ? "#FFF0F0" : ri % 2 === 0 ? "#F5F8FF" : "#EFF4FF";

                            const rowA = (
                              <tr key={ri + "a"} style={{ background: bgA }}>
                                <td style={sTd}><Badge s={r.status} /></td>
                                <td style={{ ...sTd, color: C.bl, fontWeight: 600 }}>{ls.fileAName}</td>
                                <td style={{ ...sTd, color: C.tm }}>{key}</td>
                                {sc.map(c => <td key={c} style={{ ...sTd, color: C.tm }}>{r.rowA ? (r.rowA[c] ?? "") : "\u2014"}</td>)}
                                {scm.map(m => {
                                  const f = sm.find(x => x.colA === m.colA);
                                  const vA = f && r.rowA ? (r.rowA[f.colA] ?? "") : "";
                                  const vB = f && r.rowB ? (r.rowB[f.colB] ?? "") : "";
                                  const d = r.details?.find(x => x.colA === m.colA);
                                  const df = d?.diff || "";
                                  const mi = d?.status === "Mismatched";
                                  const cm = skComment(df, r.status);
                                  return [
                                    <td key={m.colA + "c"} style={{ ...sTd, color: mi ? C.rd : C.bl, borderLeft: "2px solid #FECACA", fontWeight: mi ? 700 : 400, background: "#FFF8F8" }}>{vA || "\u2014"}</td>,
                                    <td key={m.colA + "u"} style={{ ...sTd, color: C.tl, background: "#FFF8F8" }}>{"\u2014"}</td>,
                                    <td key={m.colA + "d"} style={{ ...sTd, color: mi ? C.rd : C.tl, background: "#FFF8F8" }}>{mi ? df : "\u2014"}</td>,
                                    <td key={m.colA + "k"} style={{ ...sTd, color: C.am, fontStyle: "italic", background: "#FFF8F8", borderRight: "2px solid #FECACA" }}>{cm || "\u2014"}</td>,
                                  ];
                                })}
                              </tr>
                            );

                            const rowB = (
                              <tr key={ri + "b"} style={{ background: bgB, borderBottom: `2px solid ${C.bd}` }}>
                                <td style={{ ...sTd, color: C.tl }} />
                                <td style={{ ...sTd, color: C.pu, fontWeight: 600 }}>{ls.fileBName}</td>
                                <td style={{ ...sTd, color: C.tl }}>{key}</td>
                                {sc.map(c => { const cb = bm[c] || c; return <td key={c} style={{ ...sTd, color: C.tl }}>{r.rowB ? (r.rowB[cb] ?? "") : "\u2014"}</td>; })}
                                {scm.map(m => {
                                  const f = sm.find(x => x.colA === m.colA);
                                  const vB = f && r.rowB ? (r.rowB[f.colB] ?? "") : "";
                                  const d = r.details?.find(x => x.colA === m.colA);
                                  const mi = d?.status === "Mismatched";
                                  return [
                                    <td key={m.colA + "c"} style={{ ...sTd, color: C.tl, background: "#FFF8F8", borderLeft: "2px solid #FECACA" }}>{"\u2014"}</td>,
                                    <td key={m.colA + "u"} style={{ ...sTd, color: mi ? C.pu : C.pu, fontWeight: mi ? 700 : 400, background: mi ? "#F5E8FF" : "#FFF8F8" }}>{vB || "\u2014"}</td>,
                                    <td key={m.colA + "d"} style={{ ...sTd, color: C.tl, background: "#FFF8F8" }}>{"\u2014"}</td>,
                                    <td key={m.colA + "k"} style={{ ...sTd, color: C.tl, background: "#FFF8F8", borderRight: "2px solid #FECACA" }}>{"\u2014"}</td>,
                                  ];
                                })}
                              </tr>
                            );
                            return [rowA, rowB];
                          })}
                        </tbody>
                      </table>
                      {ls.results.length > 200 && <div style={{ padding: 10, textAlign: "center", color: C.tl, fontSize: 10 }}>Showing 200 of {ls.results.length}. Export for all.</div>}
                    </div>
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
