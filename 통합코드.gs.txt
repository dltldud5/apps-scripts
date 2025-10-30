/* =====================================================
   통합코드.gs — scalable, low-overhead version (hardened)
   ===================================================== */

const SHEET_CONFIG = {
  list:   ["1-1-1","1-1-2","1-2-1","2-1-1","2-1-2","2-1-3","2-1-4","2-6-1","2-6-2","2-6-3","2-9-2","2-10-2","3-1-2","3-1-3","4-1-1","4-1-2","4-2-1","4-2-2","4-2-3","4-3-1","4-3-2","4-4-1","4-4-2","4-4-3","4-5-1","4-5-2","4-6-1","5-1-1","5-1-2","5-1-3","6-1-3"],
  status: ["1-3-1","1-4-1","6-2-2"],
  number: ["2-2-1","2-3-1","2-4-1","2-4-2","2-4-3","2-4-4","2-7-1","2-8-1"],
  monthly:["2-5-1","2-5-2","2-9-1","2-10-1","2-11-1","3-1-1","4-6-2","4-6-3","4-6-4","6-1-1","6-1-2","6-2-1","6-3-1"],
  regional:["6-3-2"] // 새 유형: 지역별 숫자(월곶/목감/배곧1/배곧2)
};

function ensureSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}
function clearBodyRows_(sh) {
  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr > 1 && lc > 0) sh.getRange(2, 1, lr - 1, lc).clearContent();
}

function _norm(s) {
  return String(s || "").replace(/\s+/g, "").replace(/[\",']/g, "").replace(/[.,]/g, "").trim();
}
function _mdText(dateLike) {
  if (!dateLike) return "";
  const d = new Date(dateLike);
  if (!isNaN(d)) return `${d.getMonth()+1}월 ${d.getDate()}일`;
  const t = String(dateLike).replace(/\s+/g, " ").trim();
  const m = t.match(/^(\d{1,2})\s*월\s*(\d{1,2})\s*일$/);
  return m ? `${Number(m[1])}월 ${Number(m[2])}일` : "";
}
function _extractArrayFromD(raw) {
  const text = String(raw || "").trim();
  if (!text) return [];
  const asArray = text.match(/^\s*\[([\s\S]*)\]\s*$/);
  if (asArray) {
    try { return JSON.parse(text); } catch(e) {
      const out = []; let q, re = /"([^\"]+)"/g;
      while ((q = re.exec(asArray[1])) !== null) out.push(String(q[1]));
      return out;
    }
  }
  const asObj = text.match(/^\s*\{[\s\S]*\}\s*$/);
  if (asObj) { try { return [JSON.parse(text)]; } catch(e) {} }
  const objSub = text.match(/\{[\s\S]*\}/);
  if (objSub) { try { return [JSON.parse(objSub[0])]; } catch(e) {} }
  return [text];
}

function populateTabFromSheet1(ss, tabName, type) {
  const sheet1 = ss.getSheetByName("시트1");
  const target = ss.getSheetByName(tabName);
  if (!sheet1 || !target) return;
  const [maj, min, met] = target.getRange("H1:J1").getDisplayValues()[0];
  const nMaj = _norm(maj), nMin = _norm(min), nMet = _norm(met);
  const last = sheet1.getLastRow();
  if (last < 2) return;
  const keys = sheet1.getRange(2, 1, last - 1, 3).getDisplayValues();
  const vals = sheet1.getRange(2, 4, last - 1, 1).getDisplayValues();
  const idx = [];
  for (let i = 0; i < keys.length; i++) {
    if (_norm(keys[i][0]) === nMaj && _norm(keys[i][1]) === nMin && _norm(keys[i][2]) === nMet) idx.push(i);
  }
  if (type === "list") {
    const out = [];
    idx.forEach(i => {
      const arr = _extractArrayFromD(vals[i][0]);
      arr.forEach(item => {
        if (typeof item === "string") {
          let text = item.trim();
          let datePart = "";
          let contentPart = text;
          const mIdx = text.indexOf("월");
          const dIdx = text.indexOf("일", mIdx + 1);
          if (mIdx > 0 && dIdx > mIdx) {
            datePart = text.slice(0, dIdx + 1).trim();
            contentPart = text.slice(dIdx + 1).trim();
          }
          out.push([datePart, contentPart, "", "", "", ""]);
          return;
        }
        if (item && typeof item === "object") {
          const dt   = _mdText(item.date);
          const text = (item.content || item.issue || "").toString().trim();
          const cat  = (item.category || "").toString().trim();
          const tgt  = (item.target || "").toString().trim();
          const md   = (item.medal || "").toString().trim();
          const aff  = (item.affiliation || "").toString().trim();
          out.push([dt, text, cat, tgt, md, aff]);
        }
      });
    });

    if (tabName === "6-1-3") {
      target.getRange("A6:F").clearContent();
      if (out.length) target.getRange(6, 1, out.length, 6).setValues(out);
    } else {
      const slim = out.map(r => [r[0], r[1]]);
      target.getRange("A6:B").clearContent();
      if (slim.length) target.getRange(6, 1, slim.length, 2).setValues(slim);
    }
    return;
  }
  if (type === "status") {
    const map = new Map();
    const toKoStatus = (v) => {
      const s = String(v || "").toLowerCase().trim();
      if (["completed","accept","accepted","true","yes","y","반영","채택","이행"].includes(s)) return "이행";
      if (["not_completed","reject","rejected","false","no","n","미반영","미채택","미이행"].includes(s)) return "미이행";
      return v ? String(v) : "";
    };
    idx.forEach(i => {
      try {
        const arr = JSON.parse(vals[i][0]);
        (Array.isArray(arr) ? arr : [arr]).forEach(item => {
          if (!item) return;
          const issue = item.issue || item.content || item.suggestion || item.proposal || "";
          const status = item.status || item.accepted || item.result || "";
          const dateStr = item.date || item.submittedAt || item.createdAt || "";
          const md = dateStr ? _mdText(dateStr) : "";
          if (String(issue).trim()) {
            map.set(String(issue), [toKoStatus(status), md, String(issue)]);
          }
        });
      } catch(e) {}
    });
    const out = Array.from(map.values());
    target.getRange("A6:C").clearContent();
    if (out.length) target.getRange(6, 1, out.length, 3).setValues(out);
    return;
  }
  if (type === "number") {
    let val = null;
    idx.forEach(i => { const n = parseFloat(vals[i][0]); if (!isNaN(n)) val = n; });
    if (val !== null && val !== 0) target.getRange("A15").setValue(val);
    return;
  }
  if (type === "regional") {
    // 6-3-2 : 지역별 숫자 (월곶/목감/배곧1/배곧2), 단일 제출 단위는 {affiliation:"월곶", value:95.5}
    const ORDER = ["월곶","목감","배곧1","배곧2"]; // A15:D15 에 머리글, A16:D16 에 값
    // 머리글이 없다면 채워 둠
    const heads = target.getRange(15,1,1,ORDER.length).getDisplayValues()[0];
    if (!heads.some(h=>String(h).trim())) target.getRange(15,1,1,ORDER.length).setValues([ORDER]);
    // 기존 값 가져오기 -> 덮어쓰기 방식
    const existing = target.getRange(16,1,1,ORDER.length).getDisplayValues()[0];
    const store = {}; ORDER.forEach((k,i)=> store[k] = existing[i] || "");
    idx.forEach(i=>{
      try {
        const obj = JSON.parse(String(vals[i][0]));
        const aff = obj && obj.affiliation ? String(obj.affiliation).trim() : "";
        const v = obj && (obj.value ?? obj.score ?? obj.satisfaction);
        const n = parseFloat(v);
        if (aff && !isNaN(n) && ORDER.includes(aff)) store[aff] = n;
      } catch(e) {}
    });
    const row = ORDER.map(k=> store[k] || "");
    target.getRange(16,1,1,ORDER.length).setValues([row]);
    // 평균(유효 숫자만) -> G4 셀에 표기 (원하는 위치로 변경 가능)
    const nums = row.map(x=>parseFloat(x)).filter(x=>!isNaN(x));
    const avg = nums.length ? (nums.reduce((a,b)=>a+b,0)/nums.length) : "";
    target.getRange("G4").setValue(avg);
    return;
  }
  if (type === "monthly") {
    if (tabName === "6-2-1") {
      // A15:L15 = 회의, A16:L16 = 활동 (표시는 B열~M열)
      const existMeet = target.getRange("A15:L15").getDisplayValues()[0];
      const existAct  = target.getRange("A16:L16").getDisplayValues()[0];
      const finalMeet = {...Object.fromEntries(existMeet.map((v,i)=>[String(i+1), v||""]))};
      const finalAct  = {...Object.fromEntries(existAct .map((v,i)=>[String(i+1), v||""]))};

      const mergePayload = (payload) => {
        if (!payload) return;
        const mergeObj = (obj) => {
          const m = obj && obj.meetings ? obj.meetings : (obj && obj["회의"]) || null;
          const a = obj && obj.activities ? obj.activities : (obj && obj["활동"]) || null;
          if (m) Object.keys(m).forEach(k=>{ const n=k.replace(/[^0-9]/g,""); if (m[k]!=="" && m[k]!==null && n) finalMeet[n]=m[k]; });
          if (a) Object.keys(a).forEach(k=>{ const n=k.replace(/[^0-9]/g,""); if (a[k]!=="" && a[k]!==null && n) finalAct[n]=a[k]; });
        };
        try {
          const parsed = JSON.parse(String(payload));
          if (Array.isArray(parsed)) parsed.forEach(mergeObj); else mergeObj(parsed);
        } catch(e) { /* ignore malformed */ }
      };

      idx.forEach(i => mergePayload(vals[i][0]));

      const meetRow = Array.from({length:12},(_,i)=> finalMeet[String(i+1)] || "");
      const actRow  = Array.from({length:12},(_,i)=> finalAct [String(i+1)] || "");
      // Clean leftovers like literal labels (e.g., "회의"/"활동") and overwrite safely
      const clean = v => (v === "회의" || v === "활동") ? "" : v;
      const meetOut = meetRow.map(clean);
      const actOut  = actRow.map(clean);
      target.getRange("B15:M16").clearContent();
      target.getRange("B15:M15").setValues([meetOut]);
      target.getRange("B16:M16").setValues([actOut]);
      return;
    }

    if (tabName === "6-3-1") {
      // Multiple series: 회의, 간담회, 워크숍, 협약, 공동사업, 지원사업 (latest year block)
      const findYearRow = (year) => {
        const lr = target.getLastRow();
        const colA = target.getRange(1,1,lr,1).getDisplayValues().map(r=>String(r[0]).trim());
        for (let r=1;r<=lr;r++){ if (colA[r-1] === `${year}년`) return r; }
        return null;
      };
      const year = 2025; // latest block 우선 지원
      const baseRow = findYearRow(year) || findYearRow(year-1) || findYearRow(year-2) || 0;
      if (baseRow) {
        // Build current state maps for merge
        const series = [
          {label:"회의",       keyAliases:["meetings","회의"],       row: baseRow+1},
          {label:"간담회",     keyAliases:["roundtables","간담회"], row: baseRow+2},
          {label:"워크숍",     keyAliases:["workshops","워크숍","워크샵"], row: baseRow+3},
          {label:"협약",       keyAliases:["agreements","협약"],   row: baseRow+4},
          {label:"공동사업",   keyAliases:["joint_projects","jointProjects","공동사업"], row: baseRow+5},
          {label:"지원사업",   keyAliases:["support_projects","supportProjects","지원사업"], row: baseRow+6}
        ];
        const finalMap = {};
        series.forEach(s=>{
          const exist = target.getRange(s.row,2,1,12).getDisplayValues()[0];
          finalMap[s.label] = {...Object.fromEntries(exist.map((v,i)=>[String(i+1), v||""]))};
        });

        const mergeObj = (obj) => {
          series.forEach(s=>{
            let container = null;
            s.keyAliases.some(k=>{ if (obj && obj[k]!==undefined) { container=obj[k]; return true; } return false; });
            if (container) Object.keys(container).forEach(k=>{ const n=k.replace(/[^0-9]/g,""); if (container[k]!=="" && container[k]!==null && n) finalMap[s.label][n]=container[k]; });
          });
        };
        idx.forEach(i=>{ try{ const parsed=JSON.parse(String(vals[i][0])); if(Array.isArray(parsed)) parsed.forEach(mergeObj); else mergeObj(parsed);}catch(e){} });

        // Write back
        series.forEach(s=>{
          const rowVals = Array.from({length:12},(_,i)=> finalMap[s.label][String(i+1)] || "");
          target.getRange(s.row,2,1,12).clearContent();
          target.getRange(s.row,2,1,12).setValues([rowVals]);
        });
        return;
      }
    }

    const exist = target.getRange("A16:L16").getDisplayValues()[0];
    const final = {};
    for (let i = 0; i < 12; i++) if (exist[i]) final[String(i+1)] = exist[i];
    idx.forEach(i => {
      try {
        const obj = JSON.parse(vals[i][0]);
        if (obj && obj.type === "monthly" && obj.values) {
          Object.keys(obj.values).forEach(k => { if (obj.values[k]) final[String(k)] = obj.values[k]; });
        }
      } catch(e) {}
    });
    const row = []; for (let i = 1; i <= 12; i++) row.push(final[String(i)] || "");
    target.getRange("A16:L16").setValues([row]);
    return;
  }
}

function updateDashboards() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ensureSheet_(ss, "시트2");
  clearBodyRows_(dash);
  const rows = [];
  Object.entries(SHEET_CONFIG).forEach(([type, tabs]) => {
    tabs.forEach(name => {
      const t = ss.getSheetByName(name); if (!t) return;
      const [maj,min,met] = t.getRange("H1:J1").getDisplayValues()[0];
      const cat=t.getRange("D2").getDisplayValue();
      const status=t.getRange("E2").getDisplayValue();
      const targetVal=t.getRange("F2").getDisplayValue();
      const score=t.getRange("D4").getDisplayValue();
      const rate=t.getRange("E4").getDisplayValue();
      const finalScore=t.getRange("F4").getDisplayValue();
      const achRate=t.getRange("G4").getDisplayValue();
      const common=[maj,min,met,cat,status,targetVal,score,rate,finalScore,achRate];
      if (type==="list") {
        const cnt=Math.max(0,t.getLastRow()-5);
        if (cnt) t.getRange(6,1,cnt,2).getDisplayValues().forEach(r=>rows.push([...common,`${r[0]} ${r[1]}`.trim(),"","",...Array(12).fill("")]));
      } else if (type==="status") {
        const cnt = Math.max(0, t.getLastRow() - 5);
        if (cnt) {
          t.getRange(6,1,cnt,3).getDisplayValues().forEach(r => {
            rows.push([...common, "", r.join(" ").trim(), "", ...Array(12).fill("")]);
          });
        }
      } else if (type==="number") {
        const v=t.getRange("A15").getDisplayValue();
        rows.push([...common,"","",v,...Array(12).fill("")]);
      } else if (type==="regional") {
        // 6-3-2: 지역별 수치 -> 시트2에 기관/지점별로 한 줄씩 기록
        const ORDER = ["월곶","목감","배곧1","배곧2"];
        const heads = t.getRange(15,1,1,ORDER.length).getDisplayValues()[0];
        const vals  = t.getRange(16,1,1,ORDER.length).getDisplayValues()[0];
        for (let i=0;i<ORDER.length;i++) {
          const aff = heads[i] || ORDER[i];
          const v = vals[i] || "";
          rows.push([...common, aff, "", v, ...Array(12).fill("")]);
        }
      } else if (type==="monthly") {
        if (name === "6-2-1") {
      // Build two dashboard rows that match the standard 25-column shape
      // [...common, "", "", "", ...12 months]
      const meetVals = t.getRange("B15:M15").getDisplayValues()[0];
      const actVals  = t.getRange("B16:M16").getDisplayValues()[0];
      rows.push([...common, "", "", "", ...meetVals]);
      rows.push([...common, "", "", "", ...actVals]);
    } else if (name === "6-3-1") {
      // Create six rows (회의, 간담회, 워크숍, 협약, 공동사업, 지원사업) in the same 25-col format
      const findYearRow = (year) => {
        const lr = t.getLastRow();
        const colA = t.getRange(1,1,lr,1).getDisplayValues().map(r=>String(r[0]).trim());
        for (let r=1;r<=lr;r++){ if (colA[r-1] === `${year}년`) return r; }
        return null;
      };
      const baseRow = findYearRow(2025) || findYearRow(2024) || findYearRow(2023) || 0;
      if (baseRow) {
        const seriesRows=[baseRow+1,baseRow+2,baseRow+3,baseRow+4,baseRow+5,baseRow+6];
        seriesRows.forEach(ridx => {
          const vals = t.getRange(ridx,2,1,12).getDisplayValues()[0];
          rows.push([...common, "", "", "", ...vals]);
        });
      }
    } else {
          const vals=t.getRange("A16:L16").getDisplayValues()[0];
          rows.push([...common,"","","",...vals]);
        }
      }
    });
  });
  if (rows.length) dash.getRange(2,1,rows.length,rows[0].length).setValues(rows);
}

function updateMonthlyRollup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const roll = ensureSheet_(ss, "시트3");
  clearBodyRows_(roll);
  const rows = [];
  SHEET_CONFIG.monthly.forEach(name => {
    const t=ss.getSheetByName(name); if (!t) return;
    const [maj,min,met]=t.getRange("H1:J1").getDisplayValues()[0];
    const rate=t.getRange("E4").getDisplayValue();
    if (name === "6-2-1") {
      const months=t.getRange("B14:M14").getDisplayValues()[0];
      const meetVals=t.getRange("B15:M15").getDisplayValues()[0];
      const actVals =t.getRange("B16:M16").getDisplayValues()[0];
      for (let i=0;i<12;i++) rows.push([months[i],meetVals[i],maj,min,met,rate]);
      for (let i=0;i<12;i++) rows.push([months[i],actVals[i],maj,min,met,rate]);
    } else {
      const months=t.getRange("A15:L15").getDisplayValues()[0];
      const vals=t.getRange("A16:L16").getDisplayValues()[0];
      for (let i=0;i<12;i++) rows.push([months[i],vals[i],maj,min,met,rate]);
    }
  });
  if (rows.length) roll.getRange(2,1,rows.length,rows[0].length).setValues(rows);
}

function updateAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.entries(SHEET_CONFIG).forEach(([type,tabs])=>{
    tabs.forEach(name=>populateTabFromSheet1(ss,name,type));
  });
  updateDashboards();
  updateMonthlyRollup();
}

/** 자동 반영: 시트1 편집 시 전체 갱신 */
function onEdit(e) {
  const sh = e && e.range && e.range.getSheet();
  if (sh && sh.getName() === "시트1") updateAllSheets();
}

/** 메뉴 추가 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('자동화 스크립트')
    .addItem('모든 시트 업데이트', 'updateAllSheets')
    .addToUi();
}
