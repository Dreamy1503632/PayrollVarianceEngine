import { useState, useRef, useMemo } from "react";
import * as Papa from "papaparse";
import _ from "lodash";

const fmt = (n) => (typeof n === "number" ? n.toLocaleString() : n);
const pct = (n, d) => (d === 0 ? "0.0%" : ((n / d) * 100).toFixed(1) + "%");
const cleanNum = (s) => {
  if (s === null || s === undefined || String(s).trim() === "") return 0;
  return parseFloat(String(s).replace(/,/g, "")) || 0;
};

const parseFile = (file) =>
  new Promise((resolve, reject) => {
    Papa.parse(file, {
      header: true, skipEmptyLines: true, dynamicTyping: false,
      complete: (r) => {
        const cols = r.meta.fields.filter((f) => f && f.trim() !== "");
        const data = r.data.map((row) => { const c = {}; cols.forEach((cl) => (c[cl] = row[cl] ?? "")); return c; });
        resolve({ columns: cols, data });
      },
      error: reject,
    });
  });

const parseCSVString = (csvStr) => {
  const r = Papa.parse(csvStr.trim(), { header: true, skipEmptyLines: true, dynamicTyping: false });
  const cols = r.meta.fields.filter((f) => f && f.trim() !== "");
  const data = r.data.map((row) => { const c = {}; cols.forEach((cl) => (c[cl] = row[cl] ?? "")); return c; });
  return { columns: cols, data };
};

const runComparison = (data1, data2, mappings) => {
  const keyMappings = mappings.filter((m) => m.isKey);
  const compareMappings = mappings.filter((m) => m.compare && !m.isKey);
  const allActiveMappings = mappings.filter((m) => m.compare || m.isKey);
  const makeKey = (row, side) =>
    keyMappings.map((m) => {
      const col = side === 1 ? m.col1 : m.col2;
      let val = String(row[col] ?? "").trim();
      if (m.ignoreCase) val = val.toLowerCase();
      return val;
    }).join("||");
  const map1 = new Map(); const map2 = new Map();
  data1.forEach((row) => { const k = makeKey(row, 1); if (!map1.has(k)) map1.set(k, []); map1.get(k).push(row); });
  data2.forEach((row) => { const k = makeKey(row, 2); if (!map2.has(k)) map2.set(k, []); map2.get(k).push(row); });
  const allKeys = new Set([...map1.keys(), ...map2.keys()]);
  const matched = []; const mismatched = []; const onlyIn1 = []; const onlyIn2 = [];
  allKeys.forEach((k) => {
    const rows1 = map1.get(k); const rows2 = map2.get(k);
    if (!rows1) { rows2.forEach((r) => onlyIn2.push({ key: k, row: r })); return; }
    if (!rows2) { rows1.forEach((r) => onlyIn1.push({ key: k, row: r })); return; }
    const r1 = rows1[0]; const r2 = rows2[0]; const diffs = {}; let hasDiff = false;
    compareMappings.forEach((m) => {
      let v1 = String(r1[m.col1] ?? "").trim(); let v2 = String(r2[m.col2] ?? "").trim();
      if (m.ignoreCase) { v1 = v1.toLowerCase(); v2 = v2.toLowerCase(); }
      const n1 = cleanNum(v1); const n2 = cleanNum(v2);
      const isNum = !isNaN(n1) && !isNaN(n2) && v1 !== "" && v2 !== "";
      if (isNum) {
        const tol = parseFloat(m.tolerance) || 0;
        if (Math.abs(n1 - n2) > tol) { diffs[m.col1] = { v1: n1, v2: n2, diff: n2 - n1 }; hasDiff = true; }
      } else { if (v1 !== v2) { diffs[m.col1] = { v1, v2, diff: null }; hasDiff = true; } }
    });
    if (hasDiff) mismatched.push({ key: k, row1: r1, row2: r2, diffs });
    else matched.push({ key: k, row1: r1, row2: r2 });
  });
  const catCol = keyMappings.find((m) => { const l = m.col1.toLowerCase(); return l.includes("category") || l.includes("type"); })?.col1 || keyMappings[1]?.col1 || keyMappings[0]?.col1;
  const catStats = {};
  const addCat = (row, type) => { const cat = row[catCol] || "(blank)"; if (!catStats[cat]) catStats[cat] = { matched: 0, mismatched: 0, onlyIn1: 0, onlyIn2: 0 }; catStats[cat][type]++; };
  matched.forEach((m) => addCat(m.row1, "matched")); mismatched.forEach((m) => addCat(m.row1, "mismatched"));
  onlyIn1.forEach((m) => addCat(m.row, "onlyIn1")); onlyIn2.forEach((m) => addCat(m.row, "onlyIn2"));
  return { matched, mismatched, onlyIn1, onlyIn2, catStats, catCol, keyMappings, compareMappings, allActiveMappings };
};

/* ── INTERACTIVE EXCEL EXPORT ── */
const exportExcel = (results, mappings, file1Name, file2Name) => {
  const { mismatched, onlyIn1, onlyIn2, matched, keyMappings, compareMappings, catStats, catCol } = results;
  const total = matched.length + mismatched.length + onlyIn1.length + onlyIn2.length;
  const sheets = {};

  // Sheet 1: Summary
  let s = "Category,Count,Percentage\n";
  s += `Total Unique Keys,${total},100.0%\n`;
  s += `Exact Matches,${matched.length},${pct(matched.length,total)}\n`;
  s += `Value Mismatches,${mismatched.length},${pct(mismatched.length,total)}\n`;
  s += `Only in ${file1Name},${onlyIn1.length},${pct(onlyIn1.length,total)}\n`;
  s += `Only in ${file2Name},${onlyIn2.length},${pct(onlyIn2.length,total)}\n`;
  s += `\nPrimary Key,${keyMappings.map(m=>m.col1).join(" + ")}\n`;
  s += `Compare Columns,${compareMappings.map(m=>m.col1).join("; ")}\n`;
  s += `Dataset 1,${file1Name}\nDataset 2,${file2Name}\n`;
  sheets["Summary"] = s;

  // Sheet 2: Category Breakdown
  let cb = `${catCol},Matched,Mismatched,Only ${file1Name},Only ${file2Name},Total,Match Rate\n`;
  Object.entries(catStats).sort((a,b)=>b[1].mismatched-a[1].mismatched).forEach(([cat,st])=>{
    const ct=st.matched+st.mismatched+st.onlyIn1+st.onlyIn2;
    cb+=`"${cat}",${st.matched},${st.mismatched},${st.onlyIn1},${st.onlyIn2},${ct},${ct>0?((st.matched/ct)*100).toFixed(1):"100.0"}%\n`;
  });
  sheets["By Category"] = cb;

  // Sheet 3: Mismatches
  const mH = [...keyMappings.map(m=>m.col1),"Status",...compareMappings.flatMap(m=>[`${m.col1} (${file1Name})`,`${m.col1} (${file2Name})`,`${m.col1} (Diff)`,`${m.col1} (Diff %)`])];
  let mc = mH.join(",")+"\n";
  mismatched.forEach(m=>{
    const diffCols = Object.keys(m.diffs);
    const status = diffCols.length > 3 ? "HIGH" : diffCols.length > 1 ? "MEDIUM" : "LOW";
    mc += [...keyMappings.map(km=>`"${String(m.row1[km.col1]||"").replace(/"/g,'""')}"`), `"${status}"`,
      ...compareMappings.flatMap(cm=>{
        const d=m.diffs[cm.col1];
        if(d&&d.diff!==null){const pctD=d.v1!==0?((d.diff/Math.abs(d.v1))*100).toFixed(2)+"%":"NEW";return[d.v1,d.v2,d.diff.toFixed(2),pctD];}
        if(d)return[`"${d.v1}"`,`"${d.v2}"`,"TEXT DIFF",""];
        const v=cleanNum(m.row1[cm.col1]);return[v,v,"0","0%"];
      })].join(",")+"\n";
  });
  sheets["Mismatches"] = mc;

  // Sheet 4: Only DS1
  const allCols=[...keyMappings,...compareMappings];
  let o1=allCols.map(m=>m.col1).join(",")+"\n";
  onlyIn1.forEach(m=>{o1+=allCols.map(c=>`"${String(m.row[c.col1]||"").replace(/"/g,'""')}"`).join(",")+"\n";});
  sheets[`Only ${file1Name.slice(0,20)}`]=o1;

  // Sheet 5: Only DS2
  let o2=allCols.map(m=>m.col2).join(",")+"\n";
  onlyIn2.forEach(m=>{o2+=allCols.map(c=>`"${String(m.row[c.col2]||"").replace(/"/g,'""')}"`).join(",")+"\n";});
  sheets[`Only ${file2Name.slice(0,20)}`]=o2;

  // Sheet 6: All Matched
  let am=[...keyMappings.map(m=>m.col1),...compareMappings.map(m=>m.col1)].join(",")+"\n";
  matched.slice(0,5000).forEach(m=>{am+=[...keyMappings,...compareMappings].map(cm=>`"${String(m.row1[cm.col1]||"").replace(/"/g,'""')}"`).join(",")+"\n";});
  sheets["Matched"]=am;

  const dl=(c,n)=>{const b=new Blob(["\uFEFF"+c],{type:"text/csv;charset=utf-8;"});const a=document.createElement("a");a.href=URL.createObjectURL(b);a.download=n;a.click();};
  Object.entries(sheets).forEach(([name,content],i)=>{
    setTimeout(()=>dl(content,`CompareIQ_${name.replace(/[^a-zA-Z0-9]/g,"_")}.csv`),i*350);
  });
};

// ── DEMO DATA ──
const DEMO_CSV_1 = `Person Number,Person Name,Balance Category,Balance Name,Component Run Type,Gross Pay,Net Pay,Current,Year-to-Date,Primary Department
121018,Aamina Khattak,Imputed Earnings,GTL Taxable,Regular Normal,594.00,218.02,1.70,8.50,USU_PSY-CSTS
121018,Aamina Khattak,Standard Earnings,Regular Pay,Regular Normal,594.00,218.02,594.00,"4,459.29",USU_PSY-CSTS
121018,Aamina Khattak,Employee Tax Deductions,FIT Withheld,Regular Normal,594.00,218.02,0.00,345.30,USU_PSY-CSTS
121018,Aamina Khattak,Employee Tax Deductions,SS Withheld,Regular Normal,594.00,218.02,36.83,276.47,USU_PSY-CSTS
121018,Aamina Khattak,Employee Tax Deductions,Medicare Withheld,Regular Normal,594.00,218.02,8.61,64.66,USU_PSY-CSTS
121018,Aamina Khattak,Employer Taxes,SS ER,Regular Normal,594.00,218.02,36.83,276.47,USU_PSY-CSTS
121018,Aamina Khattak,Employer Taxes,Medicare ER,Regular Normal,594.00,218.02,8.61,64.66,USU_PSY-CSTS
121018,Aamina Khattak,Pretax Deductions,Medical Pre-Tax,Regular Normal,594.00,218.02,125.40,627.00,USU_PSY-CSTS
121018,Aamina Khattak,Voluntary Deductions,Roth 403(b),Regular Normal,594.00,218.02,59.40,445.50,USU_PSY-CSTS
119538,Ciera Price,Imputed Earnings,GTL Taxable,Regular Normal,"2,150.00","1,412.85",3.20,16.00,NMRC-ADMIN
119538,Ciera Price,Standard Earnings,Regular Pay,Regular Normal,"2,150.00","1,412.85","2,150.00","10,750.00",NMRC-ADMIN
119538,Ciera Price,Employee Tax Deductions,FIT Withheld,Regular Normal,"2,150.00","1,412.85",215.00,"1,075.00",NMRC-ADMIN
119538,Ciera Price,Employee Tax Deductions,SS Withheld,Regular Normal,"2,150.00","1,412.85",133.30,666.50,NMRC-ADMIN
119538,Ciera Price,Employee Tax Deductions,Medicare Withheld,Regular Normal,"2,150.00","1,412.85",31.18,155.88,NMRC-ADMIN
119538,Ciera Price,Employee Tax Deductions,VA SIT Withheld,Regular Normal,"2,150.00","1,412.85",96.75,483.75,NMRC-ADMIN
119538,Ciera Price,Employer Taxes,SS ER,Regular Normal,"2,150.00","1,412.85",133.30,666.50,NMRC-ADMIN
119538,Ciera Price,Employer Taxes,Medicare ER,Regular Normal,"2,150.00","1,412.85",31.18,155.88,NMRC-ADMIN
119538,Ciera Price,Employer Taxes,FUTA,Regular Normal,"2,150.00","1,412.85",0.00,42.00,NMRC-ADMIN
119538,Ciera Price,Pretax Deductions,Dental Pre-Tax,Regular Normal,"2,150.00","1,412.85",18.50,92.50,NMRC-ADMIN
119538,Ciera Price,Voluntary Deductions,Traditional 403(b),Regular Normal,"2,150.00","1,412.85",107.50,537.50,NMRC-ADMIN
117721,Vassiliy Tsytsarev,Standard Earnings,Regular Pay,Regular Normal,"3,800.00","2,610.42","3,800.00","19,000.00",BIO-RESEARCH
117721,Vassiliy Tsytsarev,Employee Tax Deductions,FIT Withheld,Regular Normal,"3,800.00","2,610.42",456.00,"2,280.00",BIO-RESEARCH
117721,Vassiliy Tsytsarev,Employee Tax Deductions,SS Withheld,Regular Normal,"3,800.00","2,610.42",235.60,"1,178.00",BIO-RESEARCH
117721,Vassiliy Tsytsarev,Employee Tax Deductions,Medicare Withheld,Regular Normal,"3,800.00","2,610.42",55.10,275.50,BIO-RESEARCH
117721,Vassiliy Tsytsarev,Employee Tax Deductions,MD SIT Withheld,Regular Normal,"3,800.00","2,610.42",171.00,855.00,BIO-RESEARCH
117721,Vassiliy Tsytsarev,Employer Taxes,SS ER,Regular Normal,"3,800.00","2,610.42",235.60,"1,178.00",BIO-RESEARCH
117721,Vassiliy Tsytsarev,Employer Taxes,Medicare ER,Regular Normal,"3,800.00","2,610.42",55.10,275.50,BIO-RESEARCH
117721,Vassiliy Tsytsarev,Pretax Deductions,Medical Pre-Tax,Regular Normal,"3,800.00","2,610.42",245.00,"1,225.00",BIO-RESEARCH
900531,Fiona Kiprop,Standard Earnings,Regular Pay,Regular Normal,"1,923.08","1,345.16","1,923.08","9,615.38",EMER-MED
900531,Fiona Kiprop,Employee Tax Deductions,FIT Withheld,Regular Normal,"1,923.08","1,345.16",192.31,961.54,EMER-MED
900531,Fiona Kiprop,Employee Tax Deductions,SS Withheld,Regular Normal,"1,923.08","1,345.16",119.23,596.15,EMER-MED
900531,Fiona Kiprop,Employer Taxes,SS ER,Regular Normal,"1,923.08","1,345.16",119.23,596.15,EMER-MED
900531,Fiona Kiprop,Pretax Deductions,Vision Pre-Tax,Regular Normal,"1,923.08","1,345.16",12.00,60.00,EMER-MED
120755,Franklin Morgan,Standard Earnings,Regular Pay,Regular Normal,"2,750.00","1,890.25","2,750.00","13,750.00",IT-SECURITY
120755,Franklin Morgan,Employee Tax Deductions,FIT Withheld,Regular Normal,"2,750.00","1,890.25",302.50,"1,512.50",IT-SECURITY
120755,Franklin Morgan,Employee Tax Deductions,SS Withheld,Regular Normal,"2,750.00","1,890.25",170.50,852.50,IT-SECURITY
120755,Franklin Morgan,Employee Tax Deductions,Medicare Withheld,Regular Normal,"2,750.00","1,890.25",39.88,199.38,IT-SECURITY
120755,Franklin Morgan,Employer Taxes,SS ER,Regular Normal,"2,750.00","1,890.25",170.50,852.50,IT-SECURITY
120755,Franklin Morgan,Employer Taxes,Medicare ER,Regular Normal,"2,750.00","1,890.25",39.88,199.38,IT-SECURITY`;

const DEMO_CSV_2 = `Person Number,Person Name,Balance Category,Balance Name,Component Run Type,Gross Pay,Net Pay,Current,Year-to-Date,Primary Department
121018,Aamina Khattak,Imputed Earnings,GTL Taxable,Regular Normal,594.00,218.02,1.70,8.50,USU_PSY-CSTS
121018,Aamina Khattak,Standard Earnings,Regular Pay,Regular Normal,594.00,218.02,594.00,"4,459.29",USU_PSY-CSTS
121018,Aamina Khattak,Employee Tax Deductions,FIT Withheld,Regular Normal,594.00,218.02,0.00,345.30,USU_PSY-CSTS
121018,Aamina Khattak,Employee Tax Deductions,SS Withheld,Regular Normal,594.00,218.02,36.83,276.47,USU_PSY-CSTS
121018,Aamina Khattak,Employee Tax Deductions,Medicare Withheld,Regular Normal,594.00,218.02,8.61,64.66,USU_PSY-CSTS
121018,Aamina Khattak,Employer Taxes,SS ER,Regular Normal,594.00,218.02,36.83,276.47,USU_PSY-CSTS
121018,Aamina Khattak,Employer Taxes,Medicare ER,Regular Normal,594.00,218.02,8.61,64.66,USU_PSY-CSTS
121018,Aamina Khattak,Pretax Deductions,Medical Pre-Tax,Regular Normal,594.00,218.02,125.40,627.00,USU_PSY-CSTS
121018,Aamina Khattak,Voluntary Deductions,Roth 403(b),Regular Normal,594.00,218.02,59.40,445.50,USU_PSY-CSTS
119538,Ciera Price,Imputed Earnings,GTL Taxable,Regular Normal,"2,150.00","1,412.85",3.20,16.00,NMRC-ADMIN
119538,Ciera Price,Standard Earnings,Regular Pay,Regular Normal,"2,150.00","1,412.85","2,150.00","10,750.00",NMRC-ADMIN
119538,Ciera Price,Employee Tax Deductions,FIT Withheld,Regular Normal,"2,150.00","1,412.85",228.50,"1,142.50",NMRC-ADMIN
119538,Ciera Price,Employee Tax Deductions,SS Withheld,Regular Normal,"2,150.00","1,412.85",139.75,698.75,NMRC-ADMIN
119538,Ciera Price,Employee Tax Deductions,Medicare Withheld,Regular Normal,"2,150.00","1,412.85",33.25,166.25,NMRC-ADMIN
119538,Ciera Price,Employee Tax Deductions,VA SIT Withheld,Regular Normal,"2,150.00","1,412.85",102.15,510.75,NMRC-ADMIN
119538,Ciera Price,Employee Tax Deductions,VA LIT Withheld,Regular Normal,"2,150.00","1,412.85",15.05,75.25,NMRC-ADMIN
119538,Ciera Price,Employer Taxes,SS ER,Regular Normal,"2,150.00","1,412.85",139.75,698.75,NMRC-ADMIN
119538,Ciera Price,Employer Taxes,Medicare ER,Regular Normal,"2,150.00","1,412.85",33.25,166.25,NMRC-ADMIN
119538,Ciera Price,Employer Taxes,FUTA,Regular Normal,"2,150.00","1,412.85",0.00,42.00,NMRC-ADMIN
119538,Ciera Price,Pretax Deductions,Dental Pre-Tax,Regular Normal,"2,150.00","1,412.85",18.50,92.50,NMRC-ADMIN
119538,Ciera Price,Voluntary Deductions,Traditional 403(b),Regular Normal,"2,150.00","1,412.85",107.50,537.50,NMRC-ADMIN
117721,Vassiliy Tsytsarev,Standard Earnings,Regular Pay,Regular Normal,"3,800.00","2,610.42","3,800.00","19,000.00",BIO-RESEARCH
117721,Vassiliy Tsytsarev,Employee Tax Deductions,FIT Withheld,Regular Normal,"3,800.00","2,610.42",480.00,"2,400.00",BIO-RESEARCH
117721,Vassiliy Tsytsarev,Employee Tax Deductions,SS Withheld,Regular Normal,"3,800.00","2,610.42",247.00,"1,235.00",BIO-RESEARCH
117721,Vassiliy Tsytsarev,Employee Tax Deductions,Medicare Withheld,Regular Normal,"3,800.00","2,610.42",57.80,289.00,BIO-RESEARCH
117721,Vassiliy Tsytsarev,Employee Tax Deductions,MD SIT Withheld,Regular Normal,"3,800.00","2,610.42",180.50,902.50,BIO-RESEARCH
117721,Vassiliy Tsytsarev,Employer Taxes,SS ER,Regular Normal,"3,800.00","2,610.42",247.00,"1,235.00",BIO-RESEARCH
117721,Vassiliy Tsytsarev,Employer Taxes,Medicare ER,Regular Normal,"3,800.00","2,610.42",57.80,289.00,BIO-RESEARCH
117721,Vassiliy Tsytsarev,Pretax Deductions,Medical Pre-Tax,Regular Normal,"3,800.00","2,610.42",245.00,"1,225.00",BIO-RESEARCH
900531,Fiona Kiprop,Standard Earnings,Regular Pay,Regular Normal,"1,923.08","1,345.16","1,923.08","9,615.38",EMER-MED
900531,Fiona Kiprop,Employee Tax Deductions,FIT Withheld,Regular Normal,"1,923.08","1,345.16",200.00,"1,000.00",EMER-MED
900531,Fiona Kiprop,Employee Tax Deductions,SS Withheld,Regular Normal,"1,923.08","1,345.16",125.00,625.00,EMER-MED
900531,Fiona Kiprop,Employer Taxes,SS ER,Regular Normal,"1,923.08","1,345.16",125.00,625.00,EMER-MED
900531,Fiona Kiprop,Pretax Deductions,Vision Pre-Tax,Regular Normal,"1,923.08","1,345.16",12.00,60.00,EMER-MED
120755,Franklin Morgan,Standard Earnings,Regular Pay,Regular Normal,"2,750.00","1,890.25","2,750.00","13,750.00",IT-SECURITY
120755,Franklin Morgan,Employee Tax Deductions,FIT Withheld,Regular Normal,"2,750.00","1,890.25",318.00,"1,590.00",IT-SECURITY
120755,Franklin Morgan,Employee Tax Deductions,SS Withheld,Regular Normal,"2,750.00","1,890.25",178.75,893.75,IT-SECURITY
120755,Franklin Morgan,Employee Tax Deductions,Medicare Withheld,Regular Normal,"2,750.00","1,890.25",41.80,209.00,IT-SECURITY
120755,Franklin Morgan,Employer Taxes,SS ER,Regular Normal,"2,750.00","1,890.25",178.75,893.75,IT-SECURITY
120755,Franklin Morgan,Employer Taxes,Medicare ER,Regular Normal,"2,750.00","1,890.25",41.80,209.00,IT-SECURITY`;

const Ico = {
  Upload: () => <svg width="28" height="28" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="1.5"><path strokeLinecap="round" strokeLinejoin="round" d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5m-13.5-9L12 3m0 0l4.5 4.5M12 3v13.5"/></svg>,
  Swap: () => <svg width="18" height="18" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="2"><path strokeLinecap="round" strokeLinejoin="round" d="M7.5 21L3 16.5m0 0L7.5 12M3 16.5h13.5m0-13.5L21 7.5m0 0L16.5 12M21 7.5H7.5"/></svg>,
  Key: () => <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor"><path d="M12.65 10C11.83 7.67 9.61 6 7 6c-3.31 0-6 2.69-6 6s2.69 6 6 6c2.61 0 4.83-1.67 5.65-4H17v4h4v-4h2v-4H12.65zM7 14c-1.1 0-2-.9-2-2s.9-2 2-2 2 .9 2 2-.9 2-2 2z"/></svg>,
  Download: () => <svg width="16" height="16" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="2"><path strokeLinecap="round" strokeLinejoin="round" d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5M16.5 12L12 16.5m0 0L7.5 12m4.5 4.5V3"/></svg>,
  Search: () => <svg width="15" height="15" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="2"><path strokeLinecap="round" strokeLinejoin="round" d="M21 21l-5.197-5.197m0 0A7.5 7.5 0 105.196 5.196a7.5 7.5 0 0010.607 10.607z"/></svg>,
  Info: () => <svg width="16" height="16" viewBox="0 0 24 24" fill="#3B82F6"><circle cx="12" cy="12" r="10"/><path d="M12 16v-4M12 8h.01" stroke="#fff" strokeWidth="2" strokeLinecap="round"/></svg>,
  Play: () => <svg width="18" height="18" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="2"><path strokeLinecap="round" strokeLinejoin="round" d="M5.25 5.653c0-.856.917-1.398 1.667-.986l11.54 6.348a1.125 1.125 0 010 1.971l-11.54 6.347a1.125 1.125 0 01-1.667-.985V5.653z"/></svg>,
};

export default function CompareIQPro() {
  const [step, setStep] = useState("upload");
  const [file1, setFile1] = useState(null);
  const [file2, setFile2] = useState(null);
  const [data1, setData1] = useState(null);
  const [data2, setData2] = useState(null);
  const [mappings, setMappings] = useState([]);
  const [results, setResults] = useState(null);
  const [activeTab, setActiveTab] = useState("overview");
  const [searchTerm, setSearchTerm] = useState("");
  const [loading, setLoading] = useState(false);
  const [sortConfig, setSortConfig] = useState({ key: null, dir: "asc" });
  const [demoMode, setDemoMode] = useState(false);
  const ref1 = useRef(); const ref2 = useRef();

  const buildMappings = (p1, p2) => p1.columns.map((col) => ({
    col1: col, col2: p2.columns.includes(col) ? col : p2.columns.find((c) => c.toLowerCase() === col.toLowerCase()) || "",
    isKey: false, compare: true, ignoreCase: false, tolerance: "0",
  }));

  const handleFiles = async (f1, f2) => {
    setLoading(true);
    try {
      const [p1, p2] = await Promise.all([parseFile(f1), parseFile(f2)]);
      setData1(p1); setData2(p2); setMappings(buildMappings(p1, p2)); setDemoMode(false); setStep("config");
    } catch (e) { alert("Error: " + e.message); }
    setLoading(false);
  };

  const loadDemo = () => {
    const p1 = parseCSVString(DEMO_CSV_1); const p2 = parseCSVString(DEMO_CSV_2);
    setData1(p1); setData2(p2);
    setFile1({ name: "Payroll_Vertex_2_27.csv", size: 4200 }); setFile2({ name: "Payroll_USOPTE.csv", size: 4350 });
    setDemoMode(true);
    const maps = buildMappings(p1, p2);
    ["Person Number","Balance Category","Balance Name","Component Run Type"].forEach(k=>{const m=maps.find(x=>x.col1===k);if(m)m.isKey=true;});
    setMappings(maps); setStep("config");
  };

  const updateMapping = (i, f, v) => setMappings(p => { const n=[...p]; n[i]={...n[i],[f]:v}; return n; });
  const keyCount = mappings.filter(m => m.isKey).length;
  const compareCount = mappings.filter(m => m.compare).length;

  const runCompare = () => {
    if (keyCount === 0) return alert("Select at least one Key column");
    setLoading(true);
    setTimeout(() => { setResults(runComparison(data1.data, data2.data, mappings)); setStep("results"); setActiveTab("overview"); setLoading(false); }, 50);
  };

  const filteredMismatches = useMemo(() => {
    if (!results) return [];
    let items = results.mismatched;
    if (searchTerm) { const s = searchTerm.toLowerCase(); items = items.filter(m => results.keyMappings.some(km => String(m.row1[km.col1]).toLowerCase().includes(s))); }
    if (sortConfig.key) items = [...items].sort((a, b) => { const ad = Math.abs(a.diffs[sortConfig.key]?.diff ?? 0); const bd = Math.abs(b.diffs[sortConfig.key]?.diff ?? 0); return sortConfig.dir === "asc" ? ad - bd : bd - ad; });
    return items;
  }, [results, searchTerm, sortConfig]);

  const total = results ? results.matched.length + results.mismatched.length + results.onlyIn1.length + results.onlyIn2.length : 0;

  const C = { bg:"#F5F6F8",surface:"#FFFFFF",border:"#E2E5EA",borderLight:"#F0F1F4",text:"#1A1D26",textDim:"#6C7281",textLight:"#A0A6B4",primary:"#1A7F64",primaryBg:"#E8F5F0",key:"#D97706",keyBg:"#FFF8EB",keyBorder:"#FDE68A",green:"#059669",greenBg:"#ECFDF5",red:"#DC2626",redBg:"#FEF2F2",orange:"#D97706",orangeBg:"#FFFBEB",blue:"#2563EB",blueBg:"#EFF6FF",purple:"#7C3AED",purpleBg:"#F5F3FF" };

  // ═══ UPLOAD ═══
  if (step === "upload") return (
    <div style={{ minHeight:"100vh",background:C.bg,fontFamily:"'Segoe UI',-apple-system,sans-serif",color:C.text }}>
      <div style={{ maxWidth:860,margin:"0 auto",padding:"40px 24px" }}>
        <div style={{ display:"flex",alignItems:"center",gap:10,marginBottom:48 }}>
          <div style={{ width:36,height:36,borderRadius:9,background:`linear-gradient(135deg,${C.primary},#15A07A)`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,fontWeight:800,color:"#fff" }}>IQ</div>
          <span style={{ fontSize:21,fontWeight:700,letterSpacing:-0.5 }}>CompareIQ Pro</span>
        </div>
        <div style={{ textAlign:"center",marginBottom:36 }}>
          <h1 style={{ fontSize:26,fontWeight:700,marginBottom:8 }}>Compare Any Two Datasets</h1>
          <p style={{ fontSize:14,color:C.textDim,marginBottom:20 }}>Upload CSV, TXT or TSV — map columns, set keys, configure tolerance, get detailed results + Excel export</p>
          <button onClick={loadDemo} style={{ padding:"10px 28px",borderRadius:10,border:`2px solid ${C.primary}`,background:C.primaryBg,color:C.primary,fontWeight:700,fontSize:14,cursor:"pointer",display:"inline-flex",alignItems:"center",gap:8 }}><Ico.Play /> Load Demo</button>
        </div>
        <div style={{ display:"flex",alignItems:"center",gap:16,justifyContent:"center",margin:"24px 0",color:C.textLight,fontSize:13 }}><div style={{height:1,flex:1,maxWidth:120,background:C.border}}/>or upload your files<div style={{height:1,flex:1,maxWidth:120,background:C.border}}/></div>
        <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:20,marginBottom:32 }}>
          {[{r:ref1,f:file1,s:setFile1,l:"Dataset 1"},{r:ref2,f:file2,s:setFile2,l:"Dataset 2"}].map(({r,f,s,l})=>(
            <div key={l} onClick={()=>r.current?.click()} style={{ border:`2px dashed ${f?C.green:C.border}`,borderRadius:14,padding:"40px 20px",textAlign:"center",cursor:"pointer",background:f?C.greenBg:"#FAFBFC" }}>
              <input ref={r} type="file" accept=".csv,.txt,.tsv" style={{display:"none"}} onChange={e=>{if(e.target.files[0])s(e.target.files[0]);}}/>
              <div style={{marginBottom:10,color:f?C.green:C.textLight}}><Ico.Upload/></div>
              <div style={{fontSize:14,fontWeight:600,marginBottom:4}}>{l}</div>
              {f?<div style={{color:C.primary,fontSize:13,fontWeight:600}}>{f.name}</div>:<div style={{color:C.textLight,fontSize:13}}>Drop file or click</div>}
            </div>
          ))}
        </div>
        <div style={{textAlign:"center"}}><button disabled={!file1||!file2||loading} onClick={()=>handleFiles(file1,file2)} style={{padding:"13px 40px",borderRadius:10,border:"none",fontWeight:700,fontSize:15,cursor:"pointer",background:C.primary,color:"#fff",opacity:!file1||!file2?0.4:1}}>{loading?"Parsing...":"Continue to Column Mapping →"}</button></div>
      </div>
    </div>
  );

  // ═══ CONFIG (no Data Type column) ═══
  if (step === "config") return (
    <div style={{ minHeight:"100vh",background:C.bg,fontFamily:"'Segoe UI',-apple-system,sans-serif",color:C.text }}>
      <div style={{ maxWidth:1000,margin:"0 auto",padding:"24px 24px" }}>
        <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20,flexWrap:"wrap",gap:12 }}>
          <div style={{ display:"flex",alignItems:"center",gap:14,flexWrap:"wrap" }}>
            <span style={{fontSize:18,fontWeight:700}}>Column Mapping Configuration</span>
            <span style={{display:"inline-flex",alignItems:"center",gap:5,padding:"4px 12px",borderRadius:14,background:C.greenBg,fontSize:12,fontWeight:600,color:C.green}}><span style={{width:7,height:7,borderRadius:"50%",background:C.green}}/>{file1?.name}</span>
            <span style={{color:C.textLight}}><Ico.Swap/></span>
            <span style={{display:"inline-flex",alignItems:"center",gap:5,padding:"4px 12px",borderRadius:14,background:C.orangeBg,fontSize:12,fontWeight:600,color:C.orange}}><span style={{width:7,height:7,borderRadius:"50%",background:C.orange}}/>{file2?.name}</span>
            {demoMode&&<span style={{padding:"3px 10px",borderRadius:12,background:"#DBEAFE",color:"#1D4ED8",fontSize:11,fontWeight:700}}>DEMO</span>}
          </div>
          <div style={{display:"flex",gap:10}}>
            <button onClick={()=>{setStep("upload");setData1(null);setData2(null);setFile1(null);setFile2(null);setDemoMode(false);}} style={{padding:"9px 20px",borderRadius:8,border:`1px solid ${C.border}`,background:"#fff",fontWeight:600,fontSize:13,cursor:"pointer",display:"flex",alignItems:"center",gap:6,color:C.text}}><Ico.Swap/> Change</button>
            <button onClick={runCompare} disabled={loading||keyCount===0} style={{padding:"9px 24px",borderRadius:8,border:"none",background:C.primary,color:"#fff",fontWeight:700,fontSize:13,cursor:"pointer",opacity:keyCount===0?0.5:1}}>Compare</button>
          </div>
        </div>
        <div style={{background:"#fff",border:`1px solid ${C.border}`,borderRadius:12,overflow:"hidden"}}>
          <div style={{display:"flex",alignItems:"center",gap:10,padding:"12px 20px",background:"#F0F7FF",borderBottom:"1px solid #BFDBFE",fontSize:13,color:"#1D4ED8"}}><Ico.Info/><span>Map columns, set <strong style={{color:C.key}}>key columns</strong> for matching, and specify tolerance for numeric comparisons.</span></div>
          <div style={{display:"grid",gridTemplateColumns:"1.3fr 1.3fr 62px 74px 74px 88px",gap:0,padding:"11px 20px",background:"#F9FAFB",borderBottom:`1px solid ${C.border}`}}>
            {["DATASET 1","DATASET 2","KEY 🔑","COMPARE ✓","IGNORE CASE","TOLERANCE"].map((h,i)=>(
              <div key={i} style={{fontSize:10.5,fontWeight:700,color:C.textDim,textTransform:"uppercase",letterSpacing:0.7,textAlign:i>1?"center":"left"}}>{h}</div>
            ))}
          </div>
          <div style={{maxHeight:420,overflowY:"auto"}}>
            {mappings.map((m,i)=>(
              <div key={i} style={{display:"grid",gridTemplateColumns:"1.3fr 1.3fr 62px 74px 74px 88px",gap:0,padding:"7px 20px",borderBottom:`1px solid ${C.borderLight}`,alignItems:"center",background:i%2===0?"#fff":"#FAFBFC"}}>
                <div style={{paddingRight:10}}>
                  <div style={{display:"inline-flex",alignItems:"center",gap:6,padding:"7px 14px",borderRadius:8,background:m.isKey?C.keyBg:"#F3F4F6",border:`1px solid ${m.isKey?C.keyBorder:"#E5E7EB"}`,fontSize:13,fontWeight:500,maxWidth:"100%",overflow:"hidden"}}>
                    {m.isKey&&<span style={{color:C.key,flexShrink:0}}><Ico.Key/></span>}
                    <span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{m.col1}</span>
                    <span style={{color:C.textLight,fontSize:10,marginLeft:4}}>✕ ▾</span>
                  </div>
                </div>
                <div style={{paddingRight:10}}>
                  <div style={{display:"inline-flex",alignItems:"center",padding:"4px 6px",borderRadius:8,background:"#F3F4F6",border:"1px solid #E5E7EB",maxWidth:"100%"}}>
                    <select value={m.col2} onChange={e=>updateMapping(i,"col2",e.target.value)} style={{border:"none",background:"transparent",fontSize:13,fontWeight:500,color:C.text,outline:"none",width:"100%",cursor:"pointer",padding:"3px 4px"}}>
                      <option value="">— None —</option>
                      {data2.columns.map(c=><option key={c} value={c}>{c}</option>)}
                    </select>
                    <span style={{color:C.textLight,fontSize:10}}>✕ ▾</span>
                  </div>
                </div>
                <div style={{textAlign:"center"}}><div onClick={()=>updateMapping(i,"isKey",!m.isKey)} style={{width:22,height:22,borderRadius:"50%",border:`2px solid ${m.isKey?C.key:"#D1D5DB"}`,background:m.isKey?C.key:"transparent",display:"inline-flex",alignItems:"center",justifyContent:"center",cursor:"pointer"}}>{m.isKey&&<div style={{width:10,height:10,borderRadius:"50%",background:"#fff"}}/>}</div></div>
                <div style={{textAlign:"center"}}><div onClick={()=>updateMapping(i,"compare",!m.compare)} style={{width:40,height:22,borderRadius:11,padding:2,cursor:"pointer",background:m.compare?C.primary:"#D1D5DB",display:"inline-flex",alignItems:"center",justifyContent:m.compare?"flex-end":"flex-start"}}><div style={{width:18,height:18,borderRadius:"50%",background:"#fff",boxShadow:"0 1px 3px rgba(0,0,0,0.18)"}}/></div></div>
                <div style={{textAlign:"center"}}><div onClick={()=>updateMapping(i,"ignoreCase",!m.ignoreCase)} style={{width:22,height:22,borderRadius:"50%",border:`2px solid ${m.ignoreCase?C.blue:"#D1D5DB"}`,background:m.ignoreCase?C.blue:"transparent",display:"inline-flex",alignItems:"center",justifyContent:"center",cursor:"pointer"}}>{m.ignoreCase&&<div style={{width:10,height:10,borderRadius:"50%",background:"#fff"}}/>}</div></div>
                <div style={{textAlign:"center"}}><input type="text" value={m.tolerance} onChange={e=>updateMapping(i,"tolerance",e.target.value)} placeholder="0" style={{width:72,padding:"5px 8px",borderRadius:6,border:`1px solid ${C.border}`,fontSize:12,textAlign:"center",background:"#fff",outline:"none"}}/></div>
              </div>
            ))}
          </div>
          <div style={{display:"flex",alignItems:"center",gap:28,padding:"12px 20px",background:"#F9FAFB",borderTop:`1px solid ${C.border}`,fontSize:13}}>
            <span>Total: <strong style={{color:C.blue,fontSize:15}}>{mappings.length}</strong></span>
            <span>Keys: <strong style={{color:keyCount>0?C.key:C.red,fontSize:15}}>{keyCount}</strong></span>
            <span>Compare: <strong style={{color:C.primary,fontSize:15}}>{compareCount}</strong></span>
            {keyCount===0&&<span style={{color:C.red,fontSize:12,fontWeight:600,marginLeft:"auto"}}>⚠ Select at least one key column</span>}
          </div>
        </div>
      </div>
    </div>
  );

  // ═══ RESULTS ═══
  const tabs=[{id:"overview",label:"Overview"},{id:"mismatches",label:`Mismatches (${fmt(results.mismatched.length)})`},{id:"only1",label:`Only DS1 (${fmt(results.onlyIn1.length)})`},{id:"only2",label:`Only DS2 (${fmt(results.onlyIn2.length)})`},{id:"categories",label:"By Category"},{id:"matched",label:`Matched (${fmt(results.matched.length)})`}];
  const kpis=[{label:"Exact Matches",value:results.matched.length,color:C.green,bg:C.greenBg},{label:"Value Mismatches",value:results.mismatched.length,color:C.red,bg:C.redBg},{label:"Only in Dataset 1",value:results.onlyIn1.length,color:C.orange,bg:C.orangeBg},{label:"Only in Dataset 2",value:results.onlyIn2.length,color:C.blue,bg:C.blueBg},{label:"Match Rate",value:pct(results.matched.length,total),color:C.purple,bg:C.purpleBg,isText:true}];
  const topCat=Object.entries(results.catStats).sort((a,b)=>b[1].mismatched-a[1].mismatched)[0];
  const keyLabels=results.keyMappings.map(m=>m.col1);const valLabels=results.compareMappings.map(m=>m.col1);

  return (
    <div style={{minHeight:"100vh",background:C.bg,fontFamily:"'Segoe UI',-apple-system,sans-serif",color:C.text}}>
      <div style={{maxWidth:1400,margin:"0 auto",padding:"24px 24px"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20,flexWrap:"wrap",gap:12}}>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <div style={{width:34,height:34,borderRadius:8,background:`linear-gradient(135deg,${C.primary},#15A07A)`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,fontWeight:800,color:"#fff"}}>IQ</div>
            <span style={{fontSize:20,fontWeight:700}}>CompareIQ Pro</span>
            <span style={{fontSize:12,color:C.textDim,marginLeft:8}}>Key: <strong style={{color:C.key}}>{keyLabels.join(" + ")}</strong></span>
            {demoMode&&<span style={{padding:"3px 10px",borderRadius:12,background:"#DBEAFE",color:"#1D4ED8",fontSize:11,fontWeight:700}}>DEMO</span>}
          </div>
          <div style={{display:"flex",gap:10}}>
            <button onClick={()=>{setStep("config");setResults(null);}} style={{padding:"8px 18px",borderRadius:8,border:`1px solid ${C.border}`,background:"#fff",fontWeight:600,fontSize:13,cursor:"pointer",color:C.text}}>← Reconfigure</button>
            <button onClick={()=>exportExcel(results,mappings,file1?.name||"DS1",file2?.name||"DS2")} style={{padding:"8px 20px",borderRadius:8,border:"none",background:C.primary,color:"#fff",fontWeight:600,fontSize:13,cursor:"pointer",display:"flex",alignItems:"center",gap:6}}><Ico.Download/> Export Excel (CSVs)</button>
          </div>
        </div>

        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(185px,1fr))",gap:14,marginBottom:22}}>
          {kpis.map((k,i)=>(<div key={i} style={{padding:"20px 18px",borderRadius:12,background:k.bg,border:`1px solid ${k.color}22`,position:"relative"}}><div style={{position:"absolute",top:0,left:0,right:0,height:3,background:k.color,borderRadius:"12px 12px 0 0"}}/><div style={{fontSize:11,color:C.textDim,textTransform:"uppercase",letterSpacing:0.8,marginBottom:6}}>{k.label}</div><div style={{fontSize:26,fontWeight:700,color:k.color,fontFamily:"monospace"}}>{k.isText?k.value:fmt(k.value)}</div></div>))}
        </div>

        <div style={{display:"flex",gap:4,marginBottom:18,background:"#fff",borderRadius:10,padding:4,width:"fit-content",border:`1px solid ${C.border}`,flexWrap:"wrap"}}>
          {tabs.map(t=><button key={t.id} onClick={()=>setActiveTab(t.id)} style={{padding:"9px 16px",borderRadius:8,border:"none",fontWeight:500,fontSize:13,cursor:"pointer",color:activeTab===t.id?"#fff":C.textDim,background:activeTab===t.id?C.primary:"transparent"}}>{t.label}</button>)}
        </div>

        {activeTab==="overview"&&(<div>
          <div style={{background:"#fff",border:`1px solid ${C.border}`,borderRadius:12,padding:22,marginBottom:14}}>
            <div style={{fontSize:14,fontWeight:600,marginBottom:14}}>Comparison Breakdown</div>
            <div style={{display:"flex",height:28,borderRadius:6,overflow:"hidden",marginBottom:12}}>
              {[{v:results.matched.length,c:C.green},{v:results.mismatched.length,c:C.red},{v:results.onlyIn1.length,c:C.orange},{v:results.onlyIn2.length,c:C.blue}].map((s,i)=>{const w=(s.v/total)*100;return w>0?<div key={i} style={{width:`${Math.max(w,1.5)}%`,background:s.c}}/>:null;})}
            </div>
            <div style={{display:"flex",gap:24,flexWrap:"wrap"}}>{[{c:C.green,l:"Matched",v:results.matched.length},{c:C.red,l:"Mismatched",v:results.mismatched.length},{c:C.orange,l:"Only DS1",v:results.onlyIn1.length},{c:C.blue,l:"Only DS2",v:results.onlyIn2.length}].map((x,i)=>(<div key={i} style={{display:"flex",alignItems:"center",gap:6,fontSize:13,color:C.textDim}}><div style={{width:10,height:10,borderRadius:3,background:x.c}}/>{x.l}: <strong style={{color:C.text}}>{fmt(x.v)}</strong></div>))}</div>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginBottom:14}}>
            <div style={{background:"#fff",border:`1px solid ${C.border}`,borderRadius:12,padding:20}}>
              <div style={{fontSize:13,fontWeight:600,marginBottom:10,color:C.primary}}>Configuration</div>
              <div style={{fontSize:12,color:C.textDim,marginBottom:5}}>Key: <strong style={{color:C.key}}>{keyLabels.join(" + ")}</strong></div>
              <div style={{fontSize:12,color:C.textDim,marginBottom:5}}>Values: <strong style={{color:C.text}}>{valLabels.join(", ")||"(all compare)"}</strong></div>
              <div style={{fontSize:12,color:C.textDim}}>DS1: {fmt(data1.data.length)} rows | DS2: {fmt(data2.data.length)} rows</div>
            </div>
            {topCat&&topCat[1].mismatched>0&&(<div style={{background:"#fff",border:`1px solid ${C.red}33`,borderRadius:12,padding:20}}>
              <div style={{fontSize:13,fontWeight:600,marginBottom:6,color:C.red}}>Top Problem Area</div>
              <div style={{fontSize:17,fontWeight:700,color:C.red}}>{topCat[0]}</div>
              <div style={{fontSize:12,color:C.textDim,marginTop:4}}>{fmt(topCat[1].mismatched)} of {fmt(results.mismatched.length)} mismatches</div>
            </div>)}
          </div>
          <div style={{background:"#fff",border:`1px solid ${C.border}`,borderRadius:12,overflow:"hidden"}}>
            <table style={{width:"100%",borderCollapse:"collapse"}}>
              <thead><tr>{[results.catCol,"Matched","Mismatched","Only DS1","Only DS2","Match Rate"].map(h=><th key={h} style={{padding:"10px 14px",textAlign:"left",fontSize:11,fontWeight:600,textTransform:"uppercase",color:C.textDim,borderBottom:`2px solid ${C.border}`,background:"#F9FAFB"}}>{h}</th>)}</tr></thead>
              <tbody>{Object.entries(results.catStats).sort((a,b)=>b[1].mismatched-a[1].mismatched).map(([cat,s])=>{const ct=s.matched+s.mismatched+s.onlyIn1+s.onlyIn2;const r=ct>0?(s.matched/ct)*100:100;return(<tr key={cat}><td style={{padding:"9px 14px",borderBottom:`1px solid ${C.borderLight}`,fontSize:13,fontWeight:500}}>{cat}</td><td style={{padding:"9px 14px",borderBottom:`1px solid ${C.borderLight}`,fontSize:13,fontFamily:"monospace",color:C.green}}>{fmt(s.matched)}</td><td style={{padding:"9px 14px",borderBottom:`1px solid ${C.borderLight}`,fontSize:13,fontFamily:"monospace",color:s.mismatched>0?C.red:C.textLight,fontWeight:s.mismatched>0?700:400}}>{fmt(s.mismatched)}</td><td style={{padding:"9px 14px",borderBottom:`1px solid ${C.borderLight}`,fontSize:13,fontFamily:"monospace",color:s.onlyIn1>0?C.orange:C.textLight}}>{fmt(s.onlyIn1)}</td><td style={{padding:"9px 14px",borderBottom:`1px solid ${C.borderLight}`,fontSize:13,fontFamily:"monospace",color:s.onlyIn2>0?C.blue:C.textLight}}>{fmt(s.onlyIn2)}</td><td style={{padding:"9px 14px",borderBottom:`1px solid ${C.borderLight}`}}><div style={{display:"flex",alignItems:"center",gap:8}}><div style={{flex:1,height:6,background:"#F3F4F6",borderRadius:3,overflow:"hidden"}}><div style={{width:`${r}%`,height:"100%",borderRadius:3,background:r===100?C.green:r>95?C.orange:C.red}}/></div><span style={{fontSize:12,color:C.textDim,minWidth:42,textAlign:"right"}}>{r.toFixed(1)}%</span></div></td></tr>);})}</tbody>
            </table>
          </div>
        </div>)}

        {activeTab==="mismatches"&&(<div>
          <div style={{display:"flex",gap:12,alignItems:"center",marginBottom:14}}>
            <div style={{position:"relative",flex:1,maxWidth:360}}><div style={{position:"absolute",left:11,top:"50%",transform:"translateY(-50%)",color:C.textLight}}><Ico.Search/></div><input placeholder="Search..." value={searchTerm} onChange={e=>setSearchTerm(e.target.value)} style={{width:"100%",padding:"9px 14px 9px 36px",borderRadius:8,border:`1px solid ${C.border}`,background:"#fff",fontSize:13,outline:"none",color:C.text}}/></div>
            <span style={{fontSize:12,color:C.textDim}}>{fmt(filteredMismatches.length)} mismatches</span>
          </div>
          <div style={{background:"#fff",border:`1px solid ${C.border}`,borderRadius:12,overflow:"hidden",maxHeight:500,overflowY:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse"}}>
              <thead><tr>{keyLabels.map(k=><th key={k} style={{padding:"10px 12px",textAlign:"left",fontSize:11,fontWeight:600,textTransform:"uppercase",color:C.textDim,borderBottom:`2px solid ${C.border}`,background:"#F9FAFB",position:"sticky",top:0,zIndex:2}}>{k}</th>)}{valLabels.map(v=><th key={v} colSpan={3} onClick={()=>setSortConfig(p=>({key:v,dir:p.key===v&&p.dir==="asc"?"desc":"asc"}))} style={{padding:"10px 8px",textAlign:"center",fontSize:11,fontWeight:600,textTransform:"uppercase",color:C.textDim,borderBottom:`2px solid ${C.border}`,background:"#F9FAFB",borderLeft:`2px solid ${C.border}`,cursor:"pointer",position:"sticky",top:0,zIndex:2}}>{v} {sortConfig.key===v?(sortConfig.dir==="asc"?"↑":"↓"):""}</th>)}</tr>
                <tr>{keyLabels.map(k=><th key={`s${k}`} style={{padding:"3px 12px",fontSize:9,fontWeight:600,color:C.textLight,background:"#F9FAFB",borderBottom:`1px solid ${C.border}`,position:"sticky",top:34,zIndex:2}}/>)}{valLabels.flatMap(v=>["DS1","DS2","DIFF"].map((l,li)=><th key={`${v}${l}`} style={{padding:"3px 8px",fontSize:9,fontWeight:600,color:C.textLight,background:"#F9FAFB",borderBottom:`1px solid ${C.border}`,borderLeft:li===0?`2px solid ${C.border}`:"none",position:"sticky",top:34,zIndex:2}}>{l}</th>))}</tr>
              </thead>
              <tbody>{filteredMismatches.slice(0,500).map((m,i)=>(<tr key={i} style={{background:i%2===0?"#fff":"#FAFBFC"}}>{results.keyMappings.map(km=><td key={km.col1} style={{padding:"8px 12px",borderBottom:`1px solid ${C.borderLight}`,fontSize:13}}>{m.row1[km.col1]}</td>)}{results.compareMappings.map(cm=>{const d=m.diffs[cm.col1];const v1=d?d.v1:cleanNum(m.row1[cm.col1]);const v2=d?d.v2:cleanNum(m.row2[cm.col2]);const diff=d?.diff;const hasDiff=d!=null;const isNum=typeof v1==="number";return[<td key={`${cm.col1}-1`} style={{padding:"8px 8px",borderBottom:`1px solid ${C.borderLight}`,borderLeft:`2px solid ${C.borderLight}`,fontSize:12,fontFamily:"monospace"}}>{isNum?v1.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2}):v1}</td>,<td key={`${cm.col1}-2`} style={{padding:"8px 8px",borderBottom:`1px solid ${C.borderLight}`,fontSize:12,fontFamily:"monospace"}}>{isNum?v2.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2}):v2}</td>,<td key={`${cm.col1}-d`} style={{padding:"8px 8px",borderBottom:`1px solid ${C.borderLight}`,fontSize:12,fontFamily:"monospace",color:hasDiff?(diff>0?C.green:C.red):C.textLight,fontWeight:hasDiff?700:400,background:hasDiff?(diff>0?"#F0FDF4":"#FEF2F2"):"transparent"}}>{diff!==null&&diff!==undefined?(diff>0?"+":"")+diff.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2}):hasDiff?"≠":"—"}</td>];})}</tr>))}</tbody>
            </table>
          </div>
        </div>)}

        {(activeTab==="only1"||activeTab==="only2")&&(()=>{const items=activeTab==="only1"?results.onlyIn1:results.onlyIn2;const label=activeTab==="only1"?"Dataset 1":"Dataset 2";const side=activeTab==="only1"?"col1":"col2";const allM=results.allActiveMappings;return(<div><div style={{fontSize:13,color:C.textDim,marginBottom:12}}>{fmt(items.length)} records only in {label}</div><div style={{background:"#fff",border:`1px solid ${C.border}`,borderRadius:12,overflow:"hidden",maxHeight:500,overflowY:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}><thead><tr>{allM.map(m=><th key={m[side]} style={{padding:"10px 14px",textAlign:"left",fontSize:11,fontWeight:600,textTransform:"uppercase",color:C.textDim,borderBottom:`2px solid ${C.border}`,background:"#F9FAFB",position:"sticky",top:0}}>{m[side]}</th>)}</tr></thead><tbody>{items.slice(0,500).map((m,i)=>(<tr key={i} style={{background:i%2===0?"#fff":"#FAFBFC"}}>{allM.map(cm=><td key={cm[side]} style={{padding:"9px 14px",borderBottom:`1px solid ${C.borderLight}`,fontSize:13,fontFamily:"monospace"}}>{m.row[cm[side]]}</td>)}</tr>))}{items.length===0&&<tr><td colSpan={allM.length} style={{padding:48,textAlign:"center",color:C.textLight}}>No orphan records</td></tr>}</tbody></table></div></div>);})()}

        {activeTab==="categories"&&(<div style={{display:"grid",gap:12}}>{Object.entries(results.catStats).sort((a,b)=>b[1].mismatched-a[1].mismatched).map(([cat,s])=>{const ct=s.matched+s.mismatched+s.onlyIn1+s.onlyIn2;const r=ct>0?(s.matched/ct)*100:100;return(<div key={cat} style={{background:"#fff",border:`1px solid ${C.border}`,borderRadius:12,padding:18}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}><span style={{fontSize:14,fontWeight:600}}>{cat}</span><span style={{padding:"3px 12px",borderRadius:20,fontSize:11,fontWeight:600,background:r===100?C.greenBg:r>95?C.orangeBg:C.redBg,color:r===100?C.green:r>95?C.orange:C.red}}>{r.toFixed(1)}%</span></div><div style={{display:"flex",height:8,borderRadius:4,overflow:"hidden",marginBottom:10}}><div style={{width:`${(s.matched/ct)*100}%`,background:C.green}}/><div style={{width:`${(s.mismatched/ct)*100}%`,background:C.red}}/><div style={{width:`${(s.onlyIn1/ct)*100}%`,background:C.orange}}/><div style={{width:`${(s.onlyIn2/ct)*100}%`,background:C.blue}}/></div><div style={{display:"flex",gap:20,fontSize:12,color:C.textDim}}><span>Matched: <strong style={{color:C.green}}>{fmt(s.matched)}</strong></span><span>Mismatched: <strong style={{color:C.red}}>{fmt(s.mismatched)}</strong></span><span>Only DS1: <strong style={{color:C.orange}}>{fmt(s.onlyIn1)}</strong></span><span>Only DS2: <strong style={{color:C.blue}}>{fmt(s.onlyIn2)}</strong></span></div></div>);})}</div>)}

        {activeTab==="matched"&&(<div><div style={{fontSize:13,color:C.textDim,marginBottom:12}}>Showing {fmt(Math.min(results.matched.length,200))} of {fmt(results.matched.length)}</div><div style={{background:"#fff",border:`1px solid ${C.border}`,borderRadius:12,overflow:"hidden",maxHeight:500,overflowY:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}><thead><tr>{results.allActiveMappings.map(m=><th key={m.col1} style={{padding:"10px 14px",textAlign:"left",fontSize:11,fontWeight:600,textTransform:"uppercase",color:C.textDim,borderBottom:`2px solid ${C.border}`,background:"#F9FAFB",position:"sticky",top:0}}>{m.col1}</th>)}</tr></thead><tbody>{results.matched.slice(0,200).map((m,i)=>(<tr key={i} style={{background:i%2===0?"#fff":"#FAFBFC"}}>{results.allActiveMappings.map(cm=><td key={cm.col1} style={{padding:"9px 14px",borderBottom:`1px solid ${C.borderLight}`,fontSize:13,fontFamily:"monospace"}}>{m.row1[cm.col1]}</td>)}</tr>))}</tbody></table></div></div>)}

        <div style={{textAlign:"center",padding:"28px 0 12px",fontSize:11,color:C.textLight}}>CompareIQ Pro — Precision Data Comparison Engine</div>
      </div>
    </div>
  );
}
