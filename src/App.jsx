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
    if (["xlsx","xls"].includes(ext)) {
      reader.onload = async (e) => {
        try {
          const { read, utils } = await import("https://cdn.sheetjs.com/xlsx-0.20.1/package/xlsx.mjs");
          const wb = read(new Uint8Array(e.target.result), { type: "array" });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const data = utils.sheet_to_json(ws, { defval: "" });
          resolve(data.map(r => Object.fromEntries(Object.entries(r).map(([k,v]) => [String(k).trim(), String(v).trim()]))));
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
          for (let i = 1; i <= pdfDoc.numPages; i++) { const p = await pdfDoc.getPage(i); txt += (await p.getTextContent()).items.map(it => it.str).join(" ") + "\n"; }
          resolve(autoDetect(txt));
        } catch { resolve([]); }
      };
      reader.readAsArrayBuffer(file);
    } else {
      reader.onload = (e) => {
        const text = e.target.result;
        if (ext === "tsv") resolve(parseTSV(text)); else resolve(autoDetect(text));
      };
      reader.readAsText(file);
    }
  });
}

// ─── Auto map headers ────────────────────────────────────────────────────────
function autoMapHeaders(headersA, headersB) {
  const norm = s => s.toLowerCase().replace(/[^a-z0-9]/g, "");
  const usedB = new Set();
  const prefKeys = ["personnumber","balancename","area1","area2","area3"];
  return headersA.map(ha => {
    const nha = norm(ha); let best = "", bestScore = 0;
    for (const hb of headersB) { if (usedB.has(hb)) continue; const nhb = norm(hb); const score = nha === nhb ? 100 : (nha.includes(nhb)||nhb.includes(nha)) ? 75 : 0; if (score > bestScore) { bestScore = score; best = hb; } }
    if (best && bestScore >= 75) usedB.add(best);
    const isKey = prefKeys.some(p => nha.includes(p));
    return { colA: ha, colB: best && bestScore >= 75 ? best : "", isKey, compare: !isKey && !!(best && bestScore >= 75), ignoreCase: false };
  });
}

// ─── Compare engine ──────────────────────────────────────────────────────────
function compareDatasets(dataA, dataB, mappings, tolerance) {
  const tolPct = parseFloat(tolerance)/100||0;
  const keyMaps = mappings.filter(m => m.isKey && m.colA && m.colB);
  const cMaps = mappings.filter(m => m.compare && !m.isKey && m.colA && m.colB);
  const makeKey = (row, fromA) => keyMaps.map(m => ((fromA ? row[m.colA] : row[m.colB])||"").toString().toLowerCase().trim()).join("||");
  const indexB = {};
  for (const row of dataB) { const k = makeKey(row,false); if (!indexB[k]) indexB[k]=[]; indexB[k].push(row); }
  const results = []; const matchedB = new Set();
  for (const rowA of dataA) {
    const key = makeKey(rowA,true);
    const keyVals = Object.fromEntries(keyMaps.map(m => [m.colA, rowA[m.colA]??""]));
    const matchesB = indexB[key]||[];
    if (!matchesB.length) { results.push({key,rowA,rowB:null,status:"Only in A",details:[],keyVals}); continue; }
    const rowB = matchesB[0]; matchedB.add(key);
    const details = cMaps.map(m => {
      const valA = rowA[m.colA]??"", valB = rowB[m.colB]??"";
      const cmpA = m.ignoreCase ? valA.toLowerCase() : valA, cmpB = m.ignoreCase ? valB.toLowerCase() : valB;
      const numA = parseFloat(valA), numB = parseFloat(valB);
      const isNum = !isNaN(numA)&&!isNaN(numB); let diff = "", status = "Matched";
      if (isNum) { const pct = numA!==0 ? Math.abs(numA-numB)/Math.abs(numA) : (numB!==0?1:0); diff = (numB-numA).toFixed(4); if (pct > tolPct) status = "Mismatched"; }
      else if (cmpA.trim()!==cmpB.trim()) { diff = `${valA}\u2192${valB}`; status = "Mismatched"; }
      return {colA:m.colA,colB:m.colB,valA,valB,diff,status};
    });
    results.push({key,rowA,rowB,status:details.some(d=>d.status==="Mismatched")?"Mismatched":"Matched",details,keyVals});
  }
  for (const row of dataB) { const key = makeKey(row,false); if (!matchedB.has(key)) { results.push({key,rowA:null,rowB:row,status:"Only in B",details:[],keyVals:Object.fromEntries(keyMaps.map(m=>[m.colA,row[m.colB]??""]))}); matchedB.add(key); } }
  const kc = {};
  for (const row of dataA) { const k = makeKey(row,true); kc[k]=(kc[k]||0)+1; }
  return { results, duplicatesA: Object.values(kc).filter(c=>c>1).reduce((s,c)=>s+c-1,0) };
}

// ─── SK Comment ──────────────────────────────────────────────────────────────
function makeComment(diff, valA, valB, status) {
  if (status === "Only in A") return "Only in File A";
  if (status === "Only in B") return "Only in File B";
  if (!diff || diff === "0.0000") return "";
  const num = parseFloat(diff);
  if (!isNaN(num)) return Math.abs(num) < 1 ? "Less than $1 difference" : "More $ Difference";
  return "Value mismatch";
}

// ─── ZIP builder — buffer-safe, no spread on data ────────────────────────────
function makeZip(files) {
  const te = new TextEncoder();
  const crcTable = new Uint32Array(256);
  for (let i = 0; i < 256; i++) { let c = i; for (let j = 0; j < 8; j++) c = (c&1)?(0xEDB88320^(c>>>1)):(c>>>1); crcTable[i]=c; }
  function crc32(data) { let crc = 0xFFFFFFFF; for (let i = 0; i < data.length; i++) crc = crcTable[(crc^data[i])&0xFF]^(crc>>>8); return (crc^0xFFFFFFFF)>>>0; }

  function writeU16(buf, off, v) { buf[off]=v&0xFF; buf[off+1]=(v>>>8)&0xFF; }
  function writeU32(buf, off, v) { buf[off]=v&0xFF; buf[off+1]=(v>>>8)&0xFF; buf[off+2]=(v>>>16)&0xFF; buf[off+3]=(v>>>24)&0xFF; }

  // Pre-calculate total size to allocate ONE buffer
  let totalSize = 0;
  const entries = files.map(f => {
    const nameBytes = te.encode(f.name);
    const data = f.data;
    const crc = crc32(data);
    const localHeaderSize = 30 + nameBytes.length;
    const centralDirSize = 46 + nameBytes.length;
    totalSize += localHeaderSize + data.length + centralDirSize;
    return { nameBytes, data, crc, localHeaderSize, centralDirSize };
  });
  totalSize += 22; // EOCD

  const buf = new Uint8Array(totalSize);
  let localOff = 0;
  const localOffsets = [];

  // Total local section size (for central directory offset)
  const localSectionSize = entries.reduce((s, e) => s + e.localHeaderSize + e.data.length, 0);
  let centralOff = localSectionSize;

  // Write local file headers + data
  for (const e of entries) {
    localOffsets.push(localOff);
    // Local file header signature
    buf[localOff]=0x50; buf[localOff+1]=0x4B; buf[localOff+2]=0x03; buf[localOff+3]=0x04;
    writeU16(buf, localOff+4, 20);  // version needed
    writeU16(buf, localOff+6, 0);   // flags
    writeU16(buf, localOff+8, 0);   // compression: store
    writeU16(buf, localOff+10, 0);  // mod time
    writeU16(buf, localOff+12, 0);  // mod date
    writeU32(buf, localOff+14, e.crc);
    writeU32(buf, localOff+18, e.data.length); // compressed
    writeU32(buf, localOff+22, e.data.length); // uncompressed
    writeU16(buf, localOff+26, e.nameBytes.length);
    writeU16(buf, localOff+28, 0);  // extra field length
    buf.set(e.nameBytes, localOff+30);
    buf.set(e.data, localOff+30+e.nameBytes.length);
    localOff += e.localHeaderSize + e.data.length;
  }

  // Write central directory
  const centralStart = centralOff;
  for (let i = 0; i < entries.length; i++) {
    const e = entries[i];
    buf[centralOff]=0x50; buf[centralOff+1]=0x4B; buf[centralOff+2]=0x01; buf[centralOff+3]=0x02;
    writeU16(buf, centralOff+4, 20);  // version made by
    writeU16(buf, centralOff+6, 20);  // version needed
    writeU16(buf, centralOff+8, 0);   // flags
    writeU16(buf, centralOff+10, 0);  // compression
    writeU16(buf, centralOff+12, 0);  // mod time
    writeU16(buf, centralOff+14, 0);  // mod date
    writeU32(buf, centralOff+16, e.crc);
    writeU32(buf, centralOff+20, e.data.length);
    writeU32(buf, centralOff+24, e.data.length);
    writeU16(buf, centralOff+28, e.nameBytes.length);
    writeU16(buf, centralOff+30, 0);  // extra
    writeU16(buf, centralOff+32, 0);  // comment
    writeU16(buf, centralOff+34, 0);  // disk start
    writeU16(buf, centralOff+36, 0);  // internal attr
    writeU32(buf, centralOff+38, 0);  // external attr
    writeU32(buf, centralOff+42, localOffsets[i]);
    buf.set(e.nameBytes, centralOff+46);
    centralOff += e.centralDirSize;
  }

  // EOCD
  const cdSize = centralOff - centralStart;
  buf[centralOff]=0x50; buf[centralOff+1]=0x4B; buf[centralOff+2]=0x05; buf[centralOff+3]=0x06;
  writeU16(buf, centralOff+4, 0);
  writeU16(buf, centralOff+6, 0);
  writeU16(buf, centralOff+8, entries.length);
  writeU16(buf, centralOff+10, entries.length);
  writeU32(buf, centralOff+12, cdSize);
  writeU32(buf, centralOff+16, centralStart);
  writeU16(buf, centralOff+20, 0);

  return buf;
}

// ─── XLSX Export with styling ────────────────────────────────────────────────
function exportToExcel(sessions, progressCb) {
  return new Promise((resolve, reject) => {
    // Use setTimeout to yield to the UI between heavy steps
    setTimeout(() => {
      try {
        const te = new TextEncoder();
        const encStr = s => te.encode(s);
        const esc = s => String(s??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
        const colName = n => { let s=""; for(;n>=0;n=Math.floor(n/26)-1) s=String.fromCharCode(65+n%26)+s; return s; };

        const s0 = sessions[0];
        const cMaps0 = s0.mappings.filter(m => m.compare && !m.isKey && m.colA && m.colB);
        const firstRowA = s0.results.find(r => r.rowA)?.rowA || {};
        const allColsA = Object.keys(firstRowA);
        const compareColsSet = new Set(cMaps0.map(m => m.colA));
        const sharedCols = allColsA.filter(c => !compareColsSet.has(c));

        progressCb?.("Building sheet data...");

        // Style: 0=normal, 1=header, 2=mismatch, 3=only-in-A, 4=only-in-B, 5=matched
        const stylesXML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts><fills count="8"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1A2332"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFEF2F2"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFBEB"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF5F3FF"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFECFDF5"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="6"><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/><xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1"/><xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1" applyBorder="1"/><xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0" applyFill="1" applyBorder="1"/><xf numFmtId="0" fontId="0" fillId="6" borderId="1" xfId="0" applyFill="1" applyBorder="1"/></cellXfs></styleSheet>';

        const stStyle = st => st==="Mismatched"?2:st==="Only in A"?3:st==="Only in B"?4:st==="Matched"?5:0;

        // Build sheet XML as string chunks, encode at end
        const buildSheet = (styledRows) => {
          const chunks = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'];
          for (let r = 0; r < styledRows.length; r++) {
            const row = styledRows[r];
            let xml = `<row r="${r+1}">`;
            for (let ci = 0; ci < row.cells.length; ci++) {
              const val = row.cells[ci];
              const addr = colName(ci)+(r+1);
              const si = row.style||0;
              const sv = String(val??"");
              const isNum = sv!==""&&sv.trim()!==""&&!isNaN(Number(sv))&&!/^\d{4}-/.test(sv);
              xml += isNum ? `<c r="${addr}" s="${si}"><v>${sv}</v></c>` : `<c r="${addr}" s="${si}" t="inlineStr"><is><t>${esc(val)}</t></is></c>`;
            }
            xml += "</row>";
            chunks.push(xml);
            // Keep strings from getting too huge — flush every 2000 rows
            if (chunks.length > 2000) {
              chunks[0] = chunks.join("");
              chunks.length = 1;
            }
          }
          chunks.push("</sheetData></worksheet>");
          return encStr(chunks.join(""));
        };

        // ── Summary sheet ──
        const summaryRows = [
          {cells:["CompareIQ \u2014 Comparison Summary"],style:1},{cells:[],style:0},
          {cells:["Session #","File A","File B","Total A","Total B","Matched","Mismatched","Only in A","Only in B","Duplicates","Match Rate"],style:1},
          ...sessions.map((s,i)=>({cells:[i+1,s.fileAName,s.fileBName,s.totalA,s.totalB,s.matched,s.mismatched,s.onlyA,s.onlyB,s.duplicates,`${((s.matched/(s.matched+s.mismatched||1))*100).toFixed(1)}%`],style:0}))
        ];

        // ── Comparison Results (non-matched, 2 rows per record) ──
        progressCb?.("Building comparison results...");
        const compHeader = ["Status","Source","Composite Key",...allColsA];
        const compRows = [{cells:compHeader,style:1}];
        for (const s of sessions) {
          const kMaps = s.mappings.filter(m=>m.isKey&&m.colA&&m.colB);
          const bMap = Object.fromEntries(s.mappings.filter(m=>m.colA&&m.colB).map(m=>[m.colA,m.colB]));
          for (const r of s.results.filter(r=>r.status!=="Matched")) {
            const key = kMaps.map(m=>r.keyVals[m.colA]??"").join("|");
            const st = stStyle(r.status);
            compRows.push({cells:[r.status,s.fileAName,key,...allColsA.map(c=>r.rowA?(r.rowA[c]??""):"")],style:st});
            compRows.push({cells:[r.status,s.fileBName,key,...allColsA.map(c=>{const cb=bMap[c];return r.rowB&&cb?(r.rowB[cb]??""):"";})],style:st});
          }
        }

        // ── Difference Mismatch (SK USOPTE format) ──
        progressCb?.("Building difference mismatch...");
        const diffHeader = ["Status","Source","Composite Key",...sharedCols,...cMaps0.flatMap(m=>[m.colA,`${m.colA} (USOPTE)`,"Difference","SK Comment"])];
        const diffRows = [{cells:diffHeader,style:1}];
        for (const s of sessions) {
          const kMaps = s.mappings.filter(m=>m.isKey&&m.colA&&m.colB);
          const sCmaps = s.mappings.filter(m=>m.compare&&!m.isKey&&m.colA&&m.colB);
          const bMapFull = Object.fromEntries(s.mappings.filter(m=>m.colA&&m.colB).map(m=>[m.colA,m.colB]));
          for (const r of s.results.filter(r=>r.status!=="Matched")) {
            const key = kMaps.map(m=>r.keyVals[m.colA]??"").join("|");
            const st = stStyle(r.status);
            const shA = sharedCols.map(c=>{if(r.rowA) return r.rowA[c]??""; const cb=bMapFull[c]; return r.rowB&&cb?(r.rowB[cb]??""):"";});
            const cmpA = cMaps0.flatMap(m=>{
              const sc=sCmaps.find(x=>x.colA===m.colA); const valA=sc&&r.rowA?(r.rowA[sc.colA]??""):""; const valB=sc&&r.rowB?(r.rowB[sc.colB]??""):"";
              const d=r.details?.find(dd=>dd.colA===m.colA); const diff=d?.diff||""; const comment=makeComment(diff,valA,valB,r.status);
              return [valA,"",diff?(parseFloat(diff)||diff):"",comment];
            });
            diffRows.push({cells:[r.status,s.fileAName,key,...shA,...cmpA],style:st});
            const shB = sharedCols.map(c=>{const cb=bMapFull[c]; return r.rowB&&cb?(r.rowB[cb]??""):"";});
            const cmpB = cMaps0.flatMap(m=>{const sc=sCmaps.find(x=>x.colA===m.colA); const valB=sc&&r.rowB?(r.rowB[sc.colB]??""):""; return ["",valB,"",""];});
            diffRows.push({cells:[r.status,s.fileBName,key,...shB,...cmpB],style:st});
          }
        }

        progressCb?.("Encoding XML sheets...");

        const sheets = [
          {name:"Summary",data:buildSheet(summaryRows)},
          {name:"Comparison Results",data:buildSheet(compRows)},
          {name:"Difference Mismatch",data:buildSheet(diffRows)},
        ];

        progressCb?.("Creating ZIP...");

        const wbXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets.map((s,i)=>`<sheet name="${esc(s.name)}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join("")}</sheets></workbook>`;
        const wbRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((s,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join("")}<Relationship Id="rId${sheets.length+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
        const ctXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((s,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}</Types>`;
        const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;

        const zipFiles = [
          {name:"[Content_Types].xml",data:encStr(ctXML)},
          {name:"_rels/.rels",data:encStr(rootRels)},
          {name:"xl/workbook.xml",data:encStr(wbXML)},
          {name:"xl/_rels/workbook.xml.rels",data:encStr(wbRels)},
          {name:"xl/styles.xml",data:encStr(stylesXML)},
          ...sheets.map((s,i)=>({name:`xl/worksheets/sheet${i+1}.xml`,data:s.data})),
        ];

        const zipped = makeZip(zipFiles);
        const blob = new Blob([zipped], {type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a"); a.href = url; a.download = "CompareIQ_Results.xlsx"; a.click();
        setTimeout(()=>URL.revokeObjectURL(url),3000);
        resolve();
      } catch (err) { reject(err); }
    }, 50); // yield to UI
  });
}

// ─── Styles ──────────────────────────────────────────────────────────────────
const C = {
  bg:"#F7F9FC",surface:"#FFFFFF",border:"#D1D9E6",borderStrong:"#A0B0C8",
  text:"#1A2332",textMid:"#4A5568",textLight:"#718096",
  blue:"#1A56DB",blueLight:"#EBF5FF",blueMid:"#3B82F6",
  purple:"#7C3AED",purpleLight:"#F5F3FF",
  green:"#059669",greenLight:"#ECFDF5",
  red:"#DC2626",redLight:"#FEF2F2",
  amber:"#D97706",amberLight:"#FFFBEB",
  headerBg:"#1A2332",headerText:"#FFFFFF",
};
const S = {
  card:{background:C.surface,borderRadius:12,border:`1px solid ${C.border}`,overflow:"hidden",boxShadow:"0 1px 4px rgba(0,0,0,0.06)"},
  th:{padding:"10px 13px",textAlign:"left",borderBottom:`1px solid ${C.border}`,whiteSpace:"nowrap",fontSize:11,fontWeight:700,color:C.textMid,background:"#F0F4FA"},
  td:{padding:"7px 12px",borderBottom:`1px solid ${C.border}20`,whiteSpace:"nowrap",fontSize:11,color:C.text},
  input:{background:C.surface,border:`1px solid ${C.border}`,borderRadius:7,padding:"7px 12px",color:C.text,fontSize:12,fontFamily:"inherit",outline:"none"},
  btn:(active,color=C.blue)=>({background:active?color:C.surface,border:`1px solid ${active?color:C.border}`,color:active?"#fff":C.textMid,borderRadius:7,padding:"7px 14px",cursor:"pointer",fontSize:12,fontWeight:600,fontFamily:"inherit",transition:"all .15s"}),
  btnPrimary:{background:`linear-gradient(135deg,${C.blue},${C.purple})`,color:"#fff",border:"none",borderRadius:9,padding:"11px 32px",fontSize:14,fontWeight:700,cursor:"pointer",fontFamily:"inherit"},
  label:{fontSize:11,fontWeight:700,color:C.textLight,letterSpacing:"0.05em"},
};

const StatusBadge = ({status}) => {
  const cfg = {Matched:[C.green,C.greenLight],Mismatched:[C.red,C.redLight],"Only in A":[C.amber,C.amberLight],"Only in B":[C.purple,C.purpleLight]};
  const [c,bg] = cfg[status]||[C.textLight,C.bg];
  return <span style={{background:bg,color:c,border:`1px solid ${c}33`,borderRadius:5,padding:"2px 9px",fontSize:10,fontWeight:700}}>{status}</span>;
};

const STEPS = ["Upload","Map Columns","Results"];

export default function CompareIQ() {
  const [step,setStep] = useState(0);
  const [fileA,setFileA] = useState(null); const [fileB,setFileB] = useState(null);
  const [dataA,setDataA] = useState([]); const [dataB,setDataB] = useState([]);
  const [headersA,setHeadersA] = useState([]); const [headersB,setHeadersB] = useState([]);
  const [mappings,setMappings] = useState([]);
  const [tolerance,setTolerance] = useState("1");
  const [results,setResults] = useState(null);
  const [loading,setLoading] = useState(false);
  const [exporting,setExporting] = useState(false);
  const [exportMsg,setExportMsg] = useState("");
  const [error,setError] = useState("");
  const [activeTab,setActiveTab] = useState("dashboard");
  const [filterStatus,setFilterStatus] = useState("All");
  const [searchKey,setSearchKey] = useState("");
  const [sessions,setSessions] = useState([]);

  const handleExport = async () => {
    setExporting(true); setError(""); setExportMsg("Starting export...");
    try {
      await exportToExcel(sessions, msg => setExportMsg(msg));
      setExportMsg("");
    } catch (e) { setError(`Export failed: ${e.message}`); setExportMsg(""); }
    setExporting(false);
  };

  const handleFile = async (file, which) => {
    setError(""); setLoading(true);
    try {
      const data = await parseFile(file);
      if (!data.length) throw new Error("No data found");
      const headers = Object.keys(data[0]);
      if (which==="A"){setFileA(file);setDataA(data);setHeadersA(headers);}
      else{setFileB(file);setDataB(data);setHeadersB(headers);}
    } catch (e) { setError(`Error: ${e.message}`); }
    setLoading(false);
  };

  const proceedToMap = () => {
    if (!dataA.length||!dataB.length){setError("Upload both files first.");return;}
    setMappings(autoMapHeaders(headersA,headersB)); setStep(1);
  };

  const updateMapping = (i,field,val) => setMappings(p=>p.map((m,idx)=>idx===i?{...m,[field]:val}:m));
  const setAllCompare = val => setMappings(p=>p.map(m=>m.isKey?m:{...m,compare:val}));
  const allCompareOn = mappings.filter(m=>!m.isKey).every(m=>m.compare);
  const anyCompareOn = mappings.filter(m=>!m.isKey).some(m=>m.compare);

  const runComparison = () => {
    if (!mappings.filter(m=>m.isKey).length){setError("Mark at least one KEY column.");return;}
    setLoading(true);
    setTimeout(()=>{
      const res = compareDatasets(dataA,dataB,mappings,tolerance);
      setResults(res);
      const matched=res.results.filter(r=>r.status==="Matched").length;
      const mismatched=res.results.filter(r=>r.status==="Mismatched").length;
      const onlyA=res.results.filter(r=>r.status==="Only in A").length;
      const onlyB=res.results.filter(r=>r.status==="Only in B").length;
      setSessions(prev=>[...prev,{fileAName:fileA?.name||"File A",fileBName:fileB?.name||"File B",results:res.results,totalA:dataA.length,totalB:dataB.length,matched,mismatched,onlyA,onlyB,duplicates:res.duplicatesA,mappings:[...mappings]}]);
      setStep(2); setLoading(false);
    },400);
  };

  const stats = results ? {
    totalA:dataA.length,totalB:dataB.length,
    matched:results.results.filter(r=>r.status==="Matched").length,
    mismatched:results.results.filter(r=>r.status==="Mismatched").length,
    onlyA:results.results.filter(r=>r.status==="Only in A").length,
    onlyB:results.results.filter(r=>r.status==="Only in B").length,
    duplicates:results.duplicatesA,
  } : null;

  const keyMappings = mappings.filter(m=>m.isKey&&m.colA&&m.colB);
  const compareMappings = mappings.filter(m=>m.compare&&!m.isKey&&m.colA&&m.colB);
  const filteredResults = results ? results.results.filter(r=>
    (filterStatus==="All"||r.status===filterStatus) &&
    (!searchKey||r.key.toLowerCase().includes(searchKey.toLowerCase()))
  ) : [];

  const DropZone = ({label,file,onFile,color}) => {
    const ref = useRef(); const [drag,setDrag] = useState(false);
    return (
      <div onClick={()=>ref.current.click()} onDragOver={e=>{e.preventDefault();setDrag(true);}} onDragLeave={()=>setDrag(false)}
        onDrop={e=>{e.preventDefault();setDrag(false);if(e.dataTransfer.files[0])onFile(e.dataTransfer.files[0]);}}
        style={{border:`2px dashed ${drag?color:file?C.green:C.border}`,borderRadius:12,padding:"22px 16px",textAlign:"center",cursor:"pointer",background:drag?`${color}08`:file?C.greenLight:C.bg,transition:"all .2s"}}>
        <input ref={ref} type="file" accept=".csv,.xlsx,.xls,.txt,.tsv,.pdf" style={{display:"none"}} onChange={e=>{if(e.target.files[0])onFile(e.target.files[0]);}} />
        <div style={{fontSize:26,marginBottom:6}}>{file?"\u2705":"\uD83D\uDCC1"}</div>
        <div style={{fontWeight:700,color:file?C.green:C.textMid,fontSize:13}}>{file?file.name:label}</div>
        {file?<div style={{color:C.textLight,fontSize:11,marginTop:3}}>{(file.size/1024).toFixed(1)} KB</div>
          :<div style={{color:C.textLight,fontSize:11,marginTop:5}}>CSV, Excel, TXT, TSV, PDF</div>}
      </div>
    );
  };

  const Radio = ({checked,onChange,color=C.amber}) => (
    <div onClick={onChange} style={{width:17,height:17,borderRadius:"50%",border:`2px solid ${checked?color:C.borderStrong}`,background:checked?color:"#fff",cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",margin:"auto"}}>
      {checked&&<div style={{width:6,height:6,borderRadius:"50%",background:"#fff"}}/>}
    </div>
  );
  const Checkbox = ({checked,onChange,color=C.blue}) => (
    <div onClick={onChange} style={{width:17,height:17,borderRadius:4,border:`2px solid ${checked?color:C.borderStrong}`,background:checked?color:"#fff",cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",margin:"auto"}}>
      {checked&&<span style={{color:"#fff",fontSize:11,fontWeight:900,lineHeight:1}}>{"\u2713"}</span>}
    </div>
  );
  const ColSelect = ({value,options,onChange,color}) => (
    <div style={{position:"relative",flex:1}}>
      <select value={value} onChange={e=>onChange(e.target.value)}
        style={{...S.input,width:"100%",paddingRight:24,appearance:"none",cursor:"pointer",borderColor:value?`${color}55`:C.border}}>
        <option value="">{"\u2014"} not mapped {"\u2014"}</option>
        {options.map(o=><option key={o}>{o}</option>)}
      </select>
    </div>
  );

  return (
    <div style={{minHeight:"100vh",background:C.bg,color:C.text,fontFamily:"'Inter','Segoe UI',sans-serif"}}>
      {/* Header */}
      <div style={{background:C.headerBg,padding:"12px 28px",display:"flex",alignItems:"center",justifyContent:"space-between",position:"sticky",top:0,zIndex:50,boxShadow:"0 2px 8px rgba(0,0,0,0.15)"}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <div style={{width:32,height:32,borderRadius:8,background:"linear-gradient(135deg,#3B82F6,#7C3AED)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:16}}>{"\u2696\uFE0F"}</div>
          <div><div style={{fontWeight:800,fontSize:16,color:"#fff"}}>CompareIQ</div><div style={{fontSize:9,color:"#94A3B8",letterSpacing:1.5}}>DATASET COMPARISON</div></div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:4}}>
          {STEPS.map((s,i) => (
            <div key={s} style={{display:"flex",alignItems:"center",gap:3}}>
              <div onClick={()=>step>i&&setStep(i)} style={{display:"flex",alignItems:"center",gap:6,padding:"5px 12px",borderRadius:20,background:step===i?"#3B82F6":step>i?"#1E40AF":"transparent",border:`1px solid ${step===i?"#3B82F6":step>i?"#3B82F6":"#475569"}`,color:step>=i?"#fff":"#94A3B8",fontSize:11,fontWeight:600,cursor:step>i?"pointer":"default"}}>
                <div style={{width:14,height:14,borderRadius:"50%",background:step>i?"#22C55E":step===i?"#fff":"#475569",display:"flex",alignItems:"center",justifyContent:"center",fontSize:8,color:step===i?"#3B82F6":"#fff",fontWeight:900}}>{step>i?"\u2713":i+1}</div>{s}
              </div>
              {i<STEPS.length-1&&<div style={{width:14,height:1,background:"#475569"}}/>}
            </div>
          ))}
        </div>
      </div>

      <div style={{maxWidth:1100,margin:"0 auto",padding:"28px 20px"}}>
        {error&&<div style={{background:C.redLight,border:`1px solid ${C.red}33`,color:C.red,borderRadius:9,padding:"9px 14px",marginBottom:14,fontSize:12,display:"flex",justifyContent:"space-between",alignItems:"center"}}>{"\u26A0\uFE0F"} {error}<button onClick={()=>setError("")} style={{background:"none",border:"none",color:C.red,cursor:"pointer",fontSize:16}}>{"\u00D7"}</button></div>}
        {loading&&<div style={{background:C.blueLight,border:`1px solid ${C.blue}33`,borderRadius:9,padding:"9px 14px",marginBottom:14,fontSize:12,color:C.blue}}>{"\u23F3"} Processing...</div>}
        {exporting&&<div style={{background:C.amberLight,border:`1px solid ${C.amber}33`,borderRadius:9,padding:"9px 14px",marginBottom:14,fontSize:12,color:C.amber,fontWeight:600}}>{"\u23F3"} {exportMsg||"Exporting..."}</div>}

        {/* STEP 0: Upload */}
        {step===0&&(
          <div>
            <div style={{textAlign:"center",marginBottom:28}}>
              <h1 style={{fontWeight:800,fontSize:28,margin:"0 0 6px"}}>Upload Your Datasets</h1>
              <p style={{color:C.textLight,margin:0,fontSize:13}}>CSV, Excel, TXT, TSV, PDF</p>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:24}}>
              {[{label:"Drop Dataset A \u2014 Source",which:"A",file:fileA,data:dataA,headers:headersA,color:C.blue},
                {label:"Drop Dataset B \u2014 Target",which:"B",file:fileB,data:dataB,headers:headersB,color:C.purple}
              ].map(({label,which,file,data,headers,color})=>(
                <div key={which}>
                  <div style={{...S.label,marginBottom:8,color}}>{"\u25CF"} DATASET {which}</div>
                  <DropZone label={label} file={file} color={color} onFile={f=>handleFile(f,which)} />
                  {data.length>0&&(
                    <div style={{marginTop:8,...S.card,padding:12}}>
                      <div style={{color,fontSize:11,fontWeight:700,marginBottom:6}}>{data.length} rows, {headers.length} columns</div>
                      <div style={{overflowX:"auto"}}>
                        <table style={{borderCollapse:"collapse",fontSize:10}}>
                          <thead><tr>{headers.slice(0,6).map(h=><th key={h} style={{padding:"3px 8px",background:"#F0F4FA",color:C.textMid,textAlign:"left",whiteSpace:"nowrap",fontWeight:600}}>{h}</th>)}{headers.length>6&&<th style={{color:C.textLight,padding:"3px 6px"}}>+{headers.length-6}</th>}</tr></thead>
                          <tbody>{data.slice(0,2).map((r,i)=><tr key={i}>{headers.slice(0,6).map(h=><td key={h} style={{padding:"3px 8px",color:C.textMid,whiteSpace:"nowrap"}}>{r[h]}</td>)}</tr>)}</tbody>
                        </table>
                      </div>
                    </div>
                  )}
                </div>
              ))}
            </div>
            <div style={{textAlign:"center"}}>
              <button onClick={proceedToMap} disabled={!dataA.length||!dataB.length} style={{...S.btnPrimary,opacity:dataA.length&&dataB.length?1:0.4}}>Continue to Column Mapping {"\u2192"}</button>
            </div>
          </div>
        )}

        {/* STEP 1: Column Mapping */}
        {step===1&&(
          <div>
            <div style={{marginBottom:18}}>
              <h1 style={{fontWeight:800,fontSize:22,margin:"0 0 4px"}}>Column Mapping</h1>
              <div style={{display:"flex",alignItems:"center",gap:10,marginTop:8,flexWrap:"wrap"}}>
                <div style={{padding:"4px 12px",background:C.blueLight,border:`1px solid ${C.blue}33`,borderRadius:20,fontSize:11,color:C.blue,fontWeight:600}}>{fileA?.name}</div>
                <span style={{color:C.textLight}}>{"\u21CC"}</span>
                <div style={{padding:"4px 12px",background:C.purpleLight,border:`1px solid ${C.purple}33`,borderRadius:20,fontSize:11,color:C.purple,fontWeight:600}}>{fileB?.name}</div>
                <div style={{flex:1}} />
                <span style={{fontSize:11,color:C.textLight}}>Keys: <b style={{color:C.amber}}>{mappings.filter(m=>m.isKey).length}</b></span>
                <span style={{fontSize:11,color:C.textLight}}>Compare: <b style={{color:C.green}}>{mappings.filter(m=>m.compare&&!m.isKey).length}</b></span>
              </div>
            </div>
            <div style={{display:"flex",alignItems:"center",gap:14,marginBottom:10,padding:"10px 14px",...S.card,boxShadow:"none"}}>
              <span style={S.label}>TOLERANCE</span>
              <input type="number" min="0" max="100" step="0.1" value={tolerance} onChange={e=>setTolerance(e.target.value)} style={{...S.input,width:60,textAlign:"center",borderColor:C.green,color:C.green,fontWeight:700,fontSize:14}} />
              <span style={{color:C.green,fontWeight:700}}>%</span>
              <div style={{flex:1}} />
              <button onClick={()=>setMappings(p=>[...p,{colA:headersA[0]||"",colB:headersB[0]||"",isKey:false,compare:true,ignoreCase:false}])} style={{...S.btn(false),fontSize:12,fontWeight:700}}>+ Add Row</button>
            </div>
            <div style={{...S.card,marginBottom:16}}>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 64px 80px 96px 32px",background:"#F0F4FA",padding:"8px 14px",borderBottom:`1px solid ${C.border}`}}>
                <div style={{...S.label,color:C.blue}}>DATASET A</div>
                <div style={{...S.label,color:C.purple}}>DATASET B</div>
                <div style={{...S.label,color:C.amber,textAlign:"center"}}>KEY</div>
                <div style={{...S.label,color:C.green,textAlign:"center",display:"flex",flexDirection:"column",alignItems:"center",gap:2}}>
                  <span>COMPARE</span>
                  <div style={{display:"flex",gap:4}}>
                    <button onClick={()=>setAllCompare(true)} style={{...S.btn(allCompareOn,C.green),padding:"2px 7px",fontSize:9}}>All</button>
                    <button onClick={()=>setAllCompare(false)} style={{...S.btn(!anyCompareOn,C.red),padding:"2px 7px",fontSize:9}}>None</button>
                  </div>
                </div>
                <div style={{...S.label,textAlign:"center"}}>IGNORE CASE</div>
                <div />
              </div>
              <div style={{maxHeight:380,overflowY:"auto"}}>
                {mappings.map((m,i)=>(
                  <div key={i} style={{display:"grid",gridTemplateColumns:"1fr 1fr 64px 80px 96px 32px",alignItems:"center",padding:"6px 14px",borderBottom:`1px solid ${C.border}20`,background:m.isKey?"#FFFBEB":i%2?C.surface:C.bg}}>
                    <div style={{paddingRight:8}}><ColSelect value={m.colA} options={headersA} color={C.blue} onChange={v=>updateMapping(i,"colA",v)} /></div>
                    <div style={{paddingRight:8}}><ColSelect value={m.colB} options={headersB} color={C.purple} onChange={v=>updateMapping(i,"colB",v)} /></div>
                    <div style={{textAlign:"center"}}><Radio checked={m.isKey} color={C.amber} onChange={()=>updateMapping(i,"isKey",!m.isKey)} /></div>
                    <div style={{textAlign:"center"}}><Checkbox checked={m.compare&&!m.isKey} color={C.green} onChange={()=>{if(!m.isKey)updateMapping(i,"compare",!m.compare);}} /></div>
                    <div style={{textAlign:"center"}}><Checkbox checked={m.ignoreCase} color={C.textLight} onChange={()=>updateMapping(i,"ignoreCase",!m.ignoreCase)} /></div>
                    <div style={{textAlign:"center"}}><button onClick={()=>setMappings(p=>p.filter((_,j)=>j!==i))} style={{background:"none",border:"none",color:C.textLight,cursor:"pointer",fontSize:16}}>{"\u00D7"}</button></div>
                  </div>
                ))}
              </div>
            </div>
            <div style={{display:"flex",gap:10,justifyContent:"center"}}>
              <button onClick={()=>setStep(0)} style={{...S.btn(false),padding:"10px 22px"}}>{"\u2190"} Back</button>
              <button onClick={runComparison} style={S.btnPrimary}>{"\uD83D\uDE80"} Run Comparison</button>
            </div>
          </div>
        )}

        {/* STEP 2: Results */}
        {step===2&&stats&&(
          <div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:8,marginBottom:16}}>
              {[{label:"Total A",val:stats.totalA,color:C.blue,bg:C.blueLight},{label:"Total B",val:stats.totalB,color:C.purple,bg:C.purpleLight},{label:"Matched",val:stats.matched,color:C.green,bg:C.greenLight},{label:"Mismatched",val:stats.mismatched,color:C.red,bg:C.redLight},{label:"Only in A",val:stats.onlyA,color:C.amber,bg:C.amberLight},{label:"Only in B",val:stats.onlyB,color:C.purple,bg:C.purpleLight},{label:"Duplicates",val:stats.duplicates,color:C.textMid,bg:"#F1F5F9"}].map(s=>(
                <div key={s.label} style={{background:s.bg,border:`1px solid ${s.color}30`,borderTop:`3px solid ${s.color}`,borderRadius:10,padding:"12px 8px",textAlign:"center"}}>
                  <div style={{fontSize:20,fontWeight:800,color:s.color}}>{s.val}</div>
                  <div style={{fontSize:9,color:C.textLight,marginTop:2,fontWeight:600}}>{s.label.toUpperCase()}</div>
                </div>
              ))}
            </div>

            <div style={{...S.card,padding:"12px 16px",marginBottom:14,boxShadow:"none"}}>
              <div style={{display:"flex",justifyContent:"space-between",fontSize:11,marginBottom:7}}>
                <span style={{color:C.textMid,fontWeight:600}}>Match Rate</span>
                <span style={{color:C.green,fontWeight:700}}>{((stats.matched/(stats.matched+stats.mismatched||1))*100).toFixed(1)}%</span>
              </div>
              <div style={{height:7,background:"#E2E8F0",borderRadius:99,display:"flex",overflow:"hidden"}}>
                {[{w:stats.matched,c:C.green},{w:stats.mismatched,c:C.red},{w:stats.onlyA,c:C.amber},{w:stats.onlyB,c:C.purple}].map((s,i)=>{
                  const t=stats.matched+stats.mismatched+stats.onlyA+stats.onlyB||1;
                  return <div key={i} style={{width:`${(s.w/t)*100}%`,background:s.c}} />;
                })}
              </div>
            </div>

            <div style={{display:"flex",gap:5,marginBottom:14,alignItems:"center",flexWrap:"wrap"}}>
              {[["dashboard","\uD83D\uDCCA Dashboard"],["detail","\uD83D\uDD0D Detail Report"],["sheet","\uD83D\uDCCB Comparison Sheet"]].map(([t,label])=>(
                <button key={t} onClick={()=>setActiveTab(t)} style={S.btn(activeTab===t)}>{label}</button>
              ))}
              <div style={{flex:1}} />
              <button onClick={handleExport} disabled={exporting} style={{...S.btn(false),background:exporting?C.textLight:C.green,color:"#fff",border:"none",fontWeight:700,padding:"7px 16px"}}>
                {exporting?"\u23F3 Exporting...":"\u2B07 Export Excel"}
              </button>
              <button onClick={()=>{setStep(0);setResults(null);}} style={{...S.btn(false),color:C.blue,borderColor:C.blue,fontWeight:700}}>+ New</button>
              <button onClick={()=>setStep(1)} style={S.btn(false)}>{"\u2190"} Remap</button>
            </div>

            {/* DASHBOARD */}
            {activeTab==="dashboard"&&(
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
                <div style={{...S.card,padding:20}}>
                  <div style={{...S.label,marginBottom:14}}>CLASSIFICATION BREAKDOWN</div>
                  <div style={{display:"flex",alignItems:"center",gap:20}}>
                    <svg width="110" height="110" viewBox="0 0 110 110">
                      {(()=>{
                        const segs=[{val:stats.matched,c:C.green},{val:stats.mismatched,c:C.red},{val:stats.onlyA,c:C.amber},{val:stats.onlyB,c:C.purple}];
                        const total=segs.reduce((s,x)=>s+x.val,0)||1;
                        const r=42,cx=55,cy=55; let angle=-Math.PI/2;
                        return segs.map((s,i)=>{
                          if(!s.val) return null;
                          const sweep=(s.val/total)*2*Math.PI;
                          const x1=cx+r*Math.cos(angle),y1=cy+r*Math.sin(angle);
                          angle+=sweep;
                          return <path key={i} d={`M${cx},${cy} L${x1},${y1} A${r},${r} 0 ${sweep>Math.PI?1:0} 1 ${cx+r*Math.cos(angle)},${cy+r*Math.sin(angle)} Z`} fill={s.c} opacity={0.9} />;
                        });
                      })()}
                      <circle cx="55" cy="55" r="26" fill="white" />
                      <text x="55" y="59" textAnchor="middle" fill={C.text} fontSize="11" fontWeight="bold">{((stats.matched/(stats.matched+stats.mismatched||1))*100).toFixed(0)}%</text>
                    </svg>
                    <div style={{flex:1}}>
                      {[["Matched",stats.matched,C.green],["Mismatched",stats.mismatched,C.red],["Only A",stats.onlyA,C.amber],["Only B",stats.onlyB,C.purple]].map(([l,v,c])=>(
                        <div key={l} style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
                          <div style={{width:8,height:8,borderRadius:2,background:c,flexShrink:0}} />
                          <span style={{flex:1,fontSize:12,color:C.textMid}}>{l}</span>
                          <span style={{fontWeight:700,color:c}}>{v}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
                <div style={{...S.card,padding:20}}>
                  <div style={{...S.label,marginBottom:14}}>MISMATCHES BY FIELD</div>
                  {compareMappings.length===0&&<div style={{color:C.textLight,fontSize:12}}>No compare columns</div>}
                  {compareMappings.slice(0,8).map(m=>{
                    const total=results.results.filter(r=>r.rowA&&r.rowB).length||1;
                    const mis=results.results.filter(r=>r.details.some(d=>d.colA===m.colA&&d.status==="Mismatched")).length;
                    const pct=(mis/total)*100;
                    return (
                      <div key={m.colA} style={{marginBottom:11}}>
                        <div style={{display:"flex",justifyContent:"space-between",fontSize:11,marginBottom:4}}>
                          <span style={{color:C.textMid}}>{m.colA}</span>
                          <span style={{color:pct>10?C.red:C.green,fontWeight:700}}>{mis} ({pct.toFixed(1)}%)</span>
                        </div>
                        <div style={{height:5,background:"#E2E8F0",borderRadius:99}}>
                          <div style={{width:`${Math.min(pct,100)}%`,height:"100%",background:pct>10?C.red:C.green,borderRadius:99}} />
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}

            {/* DETAIL REPORT */}
            {activeTab==="detail"&&(
              <div style={S.card}>
                <div style={{padding:"11px 14px",borderBottom:`1px solid ${C.border}`,display:"flex",gap:7,alignItems:"center",flexWrap:"wrap",background:"#F8FAFC"}}>
                  <input placeholder="Search key..." value={searchKey} onChange={e=>setSearchKey(e.target.value)} style={{...S.input,flex:1,minWidth:130}} />
                  {["All","Matched","Mismatched","Only in A","Only in B"].map(s=>(
                    <button key={s} onClick={()=>setFilterStatus(s)} style={S.btn(filterStatus===s)}>{s}</button>
                  ))}
                  <span style={{fontSize:11,color:C.textLight}}>{filteredResults.length} rows</span>
                </div>
                <div style={{overflowX:"auto",maxHeight:420,overflowY:"auto"}}>
                  <table style={{borderCollapse:"collapse",width:"100%",fontSize:11}}>
                    <thead style={{position:"sticky",top:0,zIndex:1}}>
                      <tr>
                        <th style={S.th}>Key</th>
                        {keyMappings.map(m=><th key={m.colA} style={{...S.th,color:C.amber}}>{m.colA}</th>)}
                        {compareMappings.map(m=>[
                          <th key={m.colA+"c"} style={{...S.th,color:C.blue,borderLeft:`2px solid ${C.border}`}}>Current ({m.colA})</th>,
                          <th key={m.colA+"u"} style={{...S.th,color:C.purple}}>USOPTE ({m.colB})</th>,
                          <th key={m.colA+"d"} style={{...S.th,color:C.amber}}>Diff</th>,
                          <th key={m.colA+"sk"} style={{...S.th,color:C.textLight}}>SK Comment</th>,
                        ])}
                        <th style={S.th}>Status</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredResults.slice(0,200).map((r,i)=>(
                        <tr key={i} style={{background:r.status==="Mismatched"?C.redLight:r.status==="Only in A"?C.amberLight:r.status==="Only in B"?C.purpleLight:i%2?C.bg:C.surface}}>
                          <td style={{...S.td,color:C.textMid,maxWidth:200,overflow:"hidden",textOverflow:"ellipsis"}}>{r.key}</td>
                          {keyMappings.map(m=><td key={m.colA} style={{...S.td,color:C.textMid}}>{r.keyVals[m.colA]??"\u2014"}</td>)}
                          {compareMappings.map(m=>{
                            const d=r.details.find(dd=>dd.colA===m.colA);
                            const isMis=d?.status==="Mismatched";
                            const diff=d?.diff||"";
                            const valA=d?d.valA:(r.rowA?r.rowA[m.colA]??"":"");
                            const valB=d?d.valB:(r.rowB?r.rowB[m.colB]??"":"");
                            const comment=makeComment(diff,valA,valB,r.status);
                            return [
                              <td key={m.colA+"c"} style={{...S.td,color:C.blue,borderLeft:`2px solid ${C.border}`,background:isMis?C.redLight:"transparent"}}>{valA||"\u2014"}</td>,
                              <td key={m.colA+"u"} style={{...S.td,color:C.purple,background:isMis?C.redLight:"transparent"}}>{valB||"\u2014"}</td>,
                              <td key={m.colA+"d"} style={{...S.td,color:isMis?C.red:C.textLight,fontWeight:isMis?700:400}}>{diff||"\u2014"}</td>,
                              <td key={m.colA+"sk"} style={{...S.td,color:C.amber,fontStyle:"italic",fontSize:10}}>{comment||"\u2014"}</td>,
                            ];
                          })}
                          <td style={S.td}><StatusBadge status={r.status} /></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {filteredResults.length===0&&<div style={{padding:32,textAlign:"center",color:C.textLight}}>No records match filter.</div>}
                  {filteredResults.length>200&&<div style={{padding:8,textAlign:"center",color:C.textLight,fontSize:10}}>Showing 200 of {filteredResults.length}. Export for all.</div>}
                </div>
              </div>
            )}

            {/* COMPARISON SHEET */}
            {activeTab==="sheet"&&(()=>{
              const lastS = sessions[sessions.length-1];
              if (!lastS) return null;
              const scm = lastS.mappings.filter(m=>m.compare&&!m.isKey&&m.colA&&m.colB);
              const fR = lastS.results.find(r=>r.rowA)?.rowA||{};
              const aC = Object.keys(fR);
              const cS = new Set(scm.map(m=>m.colA));
              const sC = aC.filter(c=>!cS.has(c));

              return (
                <div>
                  <div style={{background:C.blueLight,border:`1px solid ${C.blue}22`,borderRadius:9,padding:"10px 14px",marginBottom:10,fontSize:11,color:C.textMid}}>
                    2 rows per record: Row 1 = File A (Current), Row 2 = File B (USOPTE)
                  </div>
                  <div style={S.card}>
                    <div style={{overflowX:"auto",maxHeight:520,overflowY:"auto"}}>
                      <table style={{borderCollapse:"collapse",fontSize:10}}>
                        <thead style={{position:"sticky",top:0,zIndex:2}}>
                          <tr>
                            <th style={{...S.th,background:C.headerBg,color:"#fff",minWidth:80}}>Status</th>
                            <th style={{...S.th,background:C.headerBg,color:"#fff",minWidth:120}}>Source</th>
                            <th style={{...S.th,background:C.headerBg,color:"#fff",minWidth:160}}>Composite Key</th>
                            {sC.map(c=><th key={c} style={{...S.th,background:C.headerBg,color:"#CBD5E1",minWidth:90}}>{c}</th>)}
                            {scm.map(m=>[
                              <th key={m.colA+"c"} style={{...S.th,background:"#7F1D1D",color:"#FCA5A5",borderLeft:"2px solid #991B1B",minWidth:100}}>{m.colA}</th>,
                              <th key={m.colA+"u"} style={{...S.th,background:"#7F1D1D",color:"#FCA5A5",minWidth:120}}>{m.colB} (USOPTE)</th>,
                              <th key={m.colA+"d"} style={{...S.th,background:"#7F1D1D",color:"#FCA5A5",minWidth:80}}>Difference</th>,
                              <th key={m.colA+"sk"} style={{...S.th,background:"#7F1D1D",color:"#FCA5A5",borderRight:"2px solid #991B1B",minWidth:140}}>SK Comment</th>,
                            ])}
                          </tr>
                        </thead>
                        <tbody>
                          {sessions.flatMap((s,si)=>{
                            const kM=s.mappings.filter(m=>m.isKey&&m.colA&&m.colB);
                            const sm=s.mappings.filter(m=>m.compare&&!m.isKey&&m.colA&&m.colB);
                            const bM=Object.fromEntries(s.mappings.filter(m=>m.colA&&m.colB).map(m=>[m.colA,m.colB]));
                            return s.results.slice(0,150).flatMap((r,ri)=>{
                              const key=kM.map(m=>r.keyVals[m.colA]??"").join("|");
                              const isMis=r.status!=="Matched";
                              const bgA=isMis?"#FFF5F5":ri%2===0?C.surface:C.bg;
                              const bgB=isMis?"#FFF0F0":ri%2===0?"#F5F8FF":"#EFF4FF";
                              return [
                                <tr key={`${si}-${ri}-a`} style={{background:bgA}}>
                                  <td style={S.td}><StatusBadge status={r.status} /></td>
                                  <td style={{...S.td,color:C.blue,fontWeight:600}}>{s.fileAName}</td>
                                  <td style={{...S.td,color:C.textMid}}>{key}</td>
                                  {sC.map(c=><td key={c} style={{...S.td,color:C.textMid}}>{r.rowA?(r.rowA[c]??"\u2014"):"\u2014"}</td>)}
                                  {scm.map(m=>{
                                    const sc=sm.find(x=>x.colA===m.colA);
                                    const valA=sc&&r.rowA?(r.rowA[sc.colA]??""):"";
                                    const d=r.details?.find(dd=>dd.colA===m.colA);
                                    const diff=d?.diff||"";
                                    const mi=d?.status==="Mismatched";
                                    const valB=sc&&r.rowB?(r.rowB[sc.colB]??""):"";
                                    const comment=makeComment(diff,valA,valB,r.status);
                                    return [
                                      <td key={m.colA+"c"} style={{...S.td,color:mi?C.red:C.blue,background:mi?"#FFF8F8":"transparent",borderLeft:"2px solid #FECACA",fontWeight:mi?700:400}}>{valA||"\u2014"}</td>,
                                      <td key={m.colA+"u"} style={{...S.td,color:C.textLight,background:"#FFF8F8"}}>{"\u2014"}</td>,
                                      <td key={m.colA+"d"} style={{...S.td,color:mi?C.red:C.textLight,fontWeight:mi?700:400,background:"#FFF8F8"}}>{mi?diff:"\u2014"}</td>,
                                      <td key={m.colA+"sk"} style={{...S.td,color:C.amber,fontStyle:"italic",background:"#FFF8F8",borderRight:"2px solid #FECACA"}}>{comment||"\u2014"}</td>,
                                    ];
                                  })}
                                </tr>,
                                <tr key={`${si}-${ri}-b`} style={{background:bgB,borderBottom:`2px solid ${C.border}`}}>
                                  <td style={{...S.td,color:C.textLight}} />
                                  <td style={{...S.td,color:C.purple,fontWeight:600}}>{s.fileBName}</td>
                                  <td style={{...S.td,color:C.textLight}}>{key}</td>
                                  {sC.map(c=>{const cb=bM[c]||c;return <td key={c} style={{...S.td,color:C.textLight}}>{r.rowB?(r.rowB[cb]??"\u2014"):"\u2014"}</td>;})}
                                  {scm.map(m=>{
                                    const sc=sm.find(x=>x.colA===m.colA);
                                    const valB=sc&&r.rowB?(r.rowB[sc.colB]??""):"";
                                    const d=r.details?.find(dd=>dd.colA===m.colA);
                                    const mi=d?.status==="Mismatched";
                                    return [
                                      <td key={m.colA+"c"} style={{...S.td,color:C.textLight,background:"#FFF8F8",borderLeft:"2px solid #FECACA"}}>{"\u2014"}</td>,
                                      <td key={m.colA+"u"} style={{...S.td,color:mi?C.purple:C.purple,fontWeight:mi?700:400,background:mi?"#F5E8FF":"#FFF8F8"}}>{valB||"\u2014"}</td>,
                                      <td key={m.colA+"d"} style={{...S.td,color:C.textLight,background:"#FFF8F8"}}>{"\u2014"}</td>,
                                      <td key={m.colA+"sk"} style={{...S.td,color:C.textLight,background:"#FFF8F8",borderRight:"2px solid #FECACA"}}>{"\u2014"}</td>,
                                    ];
                                  })}
                                </tr>
                              ];
                            });
                          })}
                        </tbody>
                      </table>
                      {sessions.reduce((a,s)=>a+s.results.length,0)>150&&(
                        <div style={{padding:10,textAlign:"center",color:C.textLight,fontSize:10,borderTop:`1px solid ${C.border}`}}>
                          Showing first 150 records. Export for all {sessions.reduce((a,s)=>a+s.results.length,0).toLocaleString()}.
                        </div>
                      )}
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
