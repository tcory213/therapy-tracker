// ═══════════════════════════════════════════════════════════════════
// 治療人次統計系統 - Firebase 雲端版
// ═══════════════════════════════════════════════════════════════════

import { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import { initializeApp } from "firebase/app";
import {
  getFirestore, doc, getDoc, setDoc,
  collection, getDocs, onSnapshot, deleteDoc
} from "firebase/firestore";

// ──────────────────────────────────────────────
// 🔥 Firebase 設定
// ──────────────────────────────────────────────
const FIREBASE_CONFIG = {
  apiKey:            "AIzaSyDYDD2YmqgmgzgUVni0Saf7aQM7ctwFcno",
  authDomain:        "therapy-amount-calculate.firebaseapp.com",
  projectId:         "therapy-amount-calculate",
  storageBucket:     "therapy-amount-calculate.firebasestorage.app",
  messagingSenderId: "671504021434",
  appId:             "1:671504021434:web:66365ac6c0e49e76d18f83",
};

// ──────────────────────────────────────────────
// Firebase 初始化（靜態 import，最穩定）
// ──────────────────────────────────────────────
const _app = initializeApp(FIREBASE_CONFIG);
const db   = getFirestore(_app);

// Firestore CRUD helpers
async function fsGet(col, docId) {
  const snap = await getDoc(doc(db, col, docId));
  return snap.exists() ? snap.data() : null;
}
async function fsSet(col, docId, data) {
  await setDoc(doc(db, col, docId), data, { merge: true });
}
async function fsGetAll(col) {
  const snap = await getDocs(collection(db, col));
  const result = {};
  snap.forEach(d => { result[d.id] = d.data(); });
  return result;
}
function fsOnSnapshot(col, callback) {
  return onSnapshot(collection(db, col), snap => {
    const result = {};
    snap.forEach(d => { result[d.id] = d.data(); });
    callback(result);
  });
}

// ──────────────────────────────────────────────
// 常數
// ──────────────────────────────────────────────
const STAFF_OPTIONS  = ["6","5+1","5","4+1","4"];
const SESSIONS       = ["A","B","C"];
const ASSISTANT_CODES = ["R","U","W"];   // 助理代碼

function sessionType(sess) { return sess === "C" ? "C" : "AB"; }

function dateKey(y, m, d) {
  return `${y}-${String(m+1).padStart(2,"0")}-${String(d).padStart(2,"0")}`;
}
function getDaysInMonth(y, m) { return new Date(y, m+1, 0).getDate(); }
function getFirstDayOfWeek(y, m) { return new Date(y, m, 1).getDay(); }

function inRange(val, range) {
  if (!range || range[0]==="" || range[1]==="" || range[0]==null || range[1]==null) return null;
  const n = Number(val);
  if (isNaN(n) || val==="" || val===null) return null;
  return n >= Number(range[0]) && n <= Number(range[1]);
}
function aboveBonus(val, threshold) {
  if (threshold==="" || threshold==null) return false;
  const n = Number(val), t = Number(threshold);
  if (isNaN(n)||isNaN(t)||val===""||val===null) return false;
  return n > t;
}
function getPrev3WeeksDates(y, m, d) {
  const base = new Date(y, m, d);
  return [1,2,3].map(w => {
    const p = new Date(base);
    p.setDate(base.getDate() - w*7);
    return dateKey(p.getFullYear(), p.getMonth(), p.getDate());
  });
}

// ──────────────────────────────────────────────
// 班表 Excel 解析
// ──────────────────────────────────────────────
function parseScheduleExcel(arrayBuffer) {
  // cellDates:true → 日期 cell 直接輸出 JS Date 物件（SheetJS 0.18 相容）
  const wb = XLSX.read(arrayBuffer, { type:"array", cellDates: true });
  const ws = wb.Sheets["PT"];
  if (!ws) throw new Error("找不到 PT 工作表");

  const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:null });
  const result = {}; // { "YYYY-MM-DD": { A:[codes], B:[codes], C:[codes] } }

  for (let i = 0; i < rows.length - 3; i++) {
    const row = rows[i];

    // 找含 Date 物件的欄位
    const dateCells = row.map((v, ci) => {
      if (!v) return null;
      if (v instanceof Date && !isNaN(v.getTime())) {
        return { col: ci, date: v };
      }
      return null;
    }).filter(Boolean);

    if (dateCells.length === 0) continue;

    // 確認下一行是 A 班（0800-1200）
    const rowA = rows[i+1] || [];
    const rowB = rows[i+2] || [];
    const rowC = rows[i+3] || [];
    const firstCell = String(rowA[0] || "");
    if (!firstCell.includes("0800") && !firstCell.includes("08:")) continue;

    // 計算欄寬（相鄰日期的間距，通常是 6）
    const sorted = [...dateCells].sort((a, b) => a.col - b.col);
    const colWidth = sorted.length >= 2 ? (sorted[1].col - sorted[0].col) : 6;

    for (const { col, date } of sorted) {
      const dk = dateKey(date.getFullYear(), date.getMonth(), date.getDate());
      const slice = (r) => (r || []).slice(col, col + colWidth);
      const codes = (r) => slice(r)
        .filter(v => v && typeof v === "string" && /^[A-Z]$/.test(v.trim()))
        .map(v => v.trim());

      const codesA = codes(rowA);
      const codesB = codes(rowB);
      const codesC = codes(rowC);

      if (codesA.length || codesB.length || codesC.length) {
        result[dk] = { A: codesA, B: codesB, C: codesC };
      }
    }
  }

  if (Object.keys(result).length === 0) {
    throw new Error("未解析到任何排班資料，請確認 Excel 格式是否正確（需有 PT 工作表）");
  }
  return result;
}

// 從人員代碼陣列推算人員量選項
function codestoStaffOption(codes) {
  if (!codes || codes.length === 0) return null;
  const assistants = codes.filter(c => ASSISTANT_CODES.includes(c)).length;
  const therapists = codes.length - assistants;
  if (assistants > 0) return `${therapists}+1`;
  return String(therapists);
}

// ──────────────────────────────────────────────
// Toast
// ──────────────────────────────────────────────
// ══════════════════════════════════════════════════════════════
// WORKING BAR（頂部進度條，不遮蓋畫面）
// ══════════════════════════════════════════════════════════════
function WorkingBar() {
  return (
    <div style={{
      position:"fixed", top:0, left:0, right:0, height:4, zIndex:9998,
      background:"linear-gradient(90deg,#0ea5e9,#10b981,#0ea5e9)",
      backgroundSize:"200% 100%",
      animation:"workingBar 1.2s linear infinite",
    }}>
      <style>{`@keyframes workingBar{0%{background-position:0% 0}100%{background-position:200% 0}}`}</style>
    </div>
  );
}

function Toast({ message, type, onDone }) {
  useEffect(() => { const t = setTimeout(onDone, 2200); return ()=>clearTimeout(t); }, [onDone]);
  const bg = type==="success"?"#10b981":type==="error"?"#ef4444":"#0ea5e9";
  return (
    <div style={{ position:"fixed",top:24,left:"50%",transform:"translateX(-50%)",
      background:bg,color:"#fff",borderRadius:12,padding:"12px 28px",
      fontSize:16,fontWeight:700,boxShadow:"0 4px 20px rgba(0,0,0,0.18)",
      zIndex:9999,whiteSpace:"nowrap",animation:"toastIn .25s ease" }}>
      {message}
      <style>{`@keyframes toastIn{from{opacity:0;transform:translateX(-50%) translateY(-12px)}to{opacity:1;transform:translateX(-50%) translateY(0)}}`}</style>
    </div>
  );
}

// ──────────────────────────────────────────────
// Loading overlay
// ──────────────────────────────────────────────
function Loader({ text }) {
  return (
    <div style={{ position:"fixed",inset:0,background:"rgba(255,255,255,0.85)",
      display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",zIndex:8888 }}>
      <div style={{ width:48,height:48,border:"5px solid #bae6fd",
        borderTopColor:"#0ea5e9",borderRadius:"50%",animation:"spin 0.8s linear infinite" }}/>
      <div style={{ marginTop:16,color:"#0369a1",fontWeight:600,fontSize:16 }}>{text}</div>
      <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
    </div>
  );
}

// ══════════════════════════════════════════════
// MAIN APP
// ══════════════════════════════════════════════
export default function App() {
  const [fbReady,  setFbReady]  = useState(false);
  const [fbError,  setFbError]  = useState(false);
  const [loading,  setLoading]  = useState(true);   // 只用於初始化
  const [loadText, setLoadText] = useState("載入資料中…");
  const [working,  setWorking]  = useState(false);  // 設定/班表上傳用

  const [page, setPage] = useState("calendar");
  const [ranges,     setRanges]     = useState({ AB:{}, C:{} });
  const [bonusThres, setBonusThres] = useState({ AB:{}, C:{} }); // 超次獎金門檻
  const [sessions,   setSessions]   = useState({});   // { dateKey: { A:{count,staff,note,codes}, ... } }
  const [schedule,   setSchedule]   = useState({});   // { dateKey: { A:[codes], B:[codes], C:[codes] } }

  const today = new Date();
  const [viewY, setViewY] = useState(today.getFullYear());
  const [viewM, setViewM] = useState(today.getMonth());

  const [inputDate,  setInputDate]  = useState(dateKey(today.getFullYear(),today.getMonth(),today.getDate()));
  const [inputSess,  setInputSess]  = useState("A");
  const [inputCount, setInputCount] = useState("");
  const [inputStaff, setInputStaff] = useState("6");
  const [inputNote,  setInputNote]  = useState("");
  const [toast,      setToast]      = useState(null);
  const [saving,     setSaving]     = useState(false);

  const unsubRef = useRef(null);

  // ── Firebase 初始化 ──────────────────────────
  useEffect(() => {
    (async () => {
      setLoadText("載入設定…");
      try {
        const [cfg, sched] = await Promise.all([
          fsGet("config", "settings"),
          fsGetAll("schedule"),
        ]);
        if (cfg) {
          if (cfg.ranges)     setRanges(cfg.ranges);
          if (cfg.bonusThres) setBonusThres(cfg.bonusThres);
        }
        if (sched) setSchedule(sched);
      } catch(e) {
        console.error(e);
        setFbError(true);
        setLoading(false);
        return;
      }

      setLoadText("即時同步中…");
      // 即時監聽 sessions
      unsubRef.current = fsOnSnapshot("sessions", data => {
        setSessions(data);
        setLoading(false);
        setFbReady(true);
      });
    })();
    return () => { if (unsubRef.current) unsubRef.current(); };
  }, []);

  // 切換日期/節次時帶出已存資料，或從班表帶入人員量
  useEffect(() => {
    const entry = sessions[inputDate]?.[inputSess];
    if (entry) {
      setInputCount(String(entry.count));
      setInputNote(entry.note || "");
      setInputStaff(entry.staff || staffFromSchedule(inputDate, inputSess));
    } else {
      setInputCount("");
      setInputNote("");
      setInputStaff(staffFromSchedule(inputDate, inputSess) || "6");
    }
  }, [inputDate, inputSess, sessions, schedule]);

  function staffFromSchedule(dk, sess) {
    const codes = schedule[dk]?.[sess];
    if (!codes || codes.length === 0) return null;
    return codestoStaffOption(codes);
  }

  function showToast(msg, type="success") { setToast({ message:msg, type }); }

  // ── 儲存人次 ───────────────────────────────
  async function saveEntry() {
    if (inputCount === "") { showToast("請先輸入人次！","error"); return; }
    setSaving(true);
    try {
      const entry = { count:Number(inputCount), staff:inputStaff, note:inputNote };
      const existing = sessions[inputDate] || {};
      await fsSet("sessions", inputDate, { ...existing, [inputSess]: entry });
      showToast("✓ 儲存成功！");
    } catch(e) {
      showToast("儲存失敗："+e.message,"error");
    }
    setSaving(false);
  }

  // ── 分析某節 ───────────────────────────────
  const analyzeEntry = useCallback((dk, sess) => {
    const entry = sessions[dk]?.[sess];
    if (!entry) return null;
    const { count, staff } = entry;
    const st = sessionType(sess);
    const range  = ranges[st]?.[staff];
    const thres  = bonusThres[st]?.[staff];
    const ok     = inRange(count, range);
    const isBonus = aboveBonus(count, thres);

    const [y,m,d] = dk.split("-").map(Number);
    const prev3 = getPrev3WeeksDates(y, m-1, d);
    const prevCounts = prev3.map(pk=>sessions[pk]?.[sess]?.count).filter(v=>v!==undefined);
    const avg = prevCounts.length===3 ? prevCounts.reduce((a,b)=>a+b,0)/3 : null;
    const avgOk = avg!==null ? inRange(avg, range) : null;
    const avgWarn = avgOk === false;

    return { count, staff, range, thres, amt, ok, isBonus, avg, avgOk, avgWarn };
  }, [sessions, ranges, bonusThres]);

  // ── chip 顏色狀態 ───────────────────────────
  function chipStatus(dk, sess) {
    const info = sessions[dk]?.[sess];
    if (!info) return "empty";
    const a = analyzeEntry(dk, sess);
    if (!a) return "empty";
    if (a.avgWarn) return "avg_warn";
    if (a.isBonus) return "bonus";
    if (a.ok === true)  return "ok";
    if (a.ok === false) return "avg_warn";
    return "neutral";
  }

  // ── 月超次統計 ──────────────────────────────
  function calcMonthlyOvertime(y, m) {
    const days = getDaysInMonth(y, m);
    const overtimeMap = {}; // { therapistCode: count }

    for (let d = 1; d <= days; d++) {
      const dk = dateKey(y, m, d);
      const sched = schedule[dk];
      if (!sched) continue;

      for (const sess of SESSIONS) {
        const entry = sessions[dk]?.[sess];
        if (!entry) continue;
        const st = sessionType(sess);
        const thres = bonusThres[st]?.[entry.staff];
        if (!aboveBonus(entry.count, thres)) continue;

        const codes = sched[sess] || [];
        const therapists = codes.filter(c => !ASSISTANT_CODES.includes(c));
        for (const code of therapists) {
          overtimeMap[code] = (overtimeMap[code] || 0) + 1;
        }
      }
    }
    return overtimeMap;
  }

  // ── 儲存設定 ───────────────────────────────
  async function saveSettings(newRanges, newBonusThres) {
    setWorking(true);
    try {
      await fsSet("config","settings",{ ranges:newRanges, bonusThres:newBonusThres });
      setRanges(newRanges);
      setBonusThres(newBonusThres);
      showToast("✓ 設定儲存成功！");
    } catch(e) { showToast("儲存失敗","error"); }
    setWorking(false);
  }

  // ── 上傳班表（清空該月再重寫）─────────────
  async function uploadSchedule(file) {
    setWorking(true);
    try {
      const buf = await file.arrayBuffer();
      const parsed = parseScheduleExcel(buf);
      if (Object.keys(parsed).length === 0) throw new Error("未解析到任何資料");

      // 找出這次上傳涵蓋的年月
      const months = new Set(Object.keys(parsed).map(dk => dk.slice(0,7))); // "YYYY-MM"


      // 刪除 Firestore 中同月份的舊資料
      const snap = await getDocs(collection(db, "schedule"));
      const toDelete = [];
      snap.forEach(d => {
        if (months.has(d.id.slice(0,7))) toDelete.push(d.id);
      });
      for (const id of toDelete) {
        await deleteDoc(doc(db, "schedule", id));
      }


      for (const [dk, data] of Object.entries(parsed)) {
        await fsSet("schedule", dk, data);
      }

      // 更新本地 schedule state：移除舊月份，加入新資料
      setSchedule(prev => {
        const next = { ...prev };
        for (const id of toDelete) delete next[id];
        return { ...next, ...parsed };
      });
      showToast(`✓ 班表上傳成功！共 ${Object.keys(parsed).length} 天`);
    } catch(e) {
      showToast("班表解析失敗："+e.message, "error");
    }
    setWorking(false);
  }

  // ──────────────────────────────────────────────
  if (fbError) return (
    <div style={{ minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",
      fontFamily:"'Noto Sans TC',sans-serif",background:"#fff0f0",flexDirection:"column",gap:16,padding:24 }}>
      <div style={{ fontSize:48 }}>⚠️</div>
      <div style={{ fontSize:20,fontWeight:700,color:"#991b1b" }}>Firebase 連線失敗</div>
      <div style={{ color:"#7f1d1d",fontSize:14,maxWidth:400,textAlign:"center" }}>
        請確認已在程式碼頂部填入正確的 FIREBASE_CONFIG，並在 Firebase Console 啟用 Firestore。
      </div>
    </div>
  );

  return (
    <div style={S.root}>
      {loading && <Loader text={loadText} />}
      {working && <WorkingBar />}
      {toast && <Toast message={toast.message} type={toast.type} onDone={()=>setToast(null)} />}

      <nav style={S.nav}>
        <span style={S.navBrand}>🏥 治療人次統計</span>
        <div style={S.navTabs}>
          {[["calendar","📅 月曆"],["input","✏️ 輸入"],["stats","📊 超次統計"],["settings","⚙️ 設定"]].map(([p,l])=>(
            <button key={p} onClick={()=>setPage(p)}
              style={{ ...S.navBtn, ...(page===p?S.navBtnActive:{}) }}>{l}</button>
          ))}
        </div>
      </nav>

      <main style={S.main}>
        {page==="calendar" && (
          <CalendarPage viewY={viewY} viewM={viewM}
            setViewY={setViewY} setViewM={setViewM}
            sessions={sessions} schedule={schedule}
            analyzeEntry={analyzeEntry} chipStatus={chipStatus}
            onClickCell={(dk,sess)=>{ setInputDate(dk); setInputSess(sess); setPage("input"); }}
          />
        )}
        {page==="input" && (
          <InputPage
            inputDate={inputDate} setInputDate={setInputDate}
            inputSess={inputSess} setInputSess={setInputSess}
            inputCount={inputCount} setInputCount={setInputCount}
            inputStaff={inputStaff} setInputStaff={setInputStaff}
            inputNote={inputNote} setInputNote={setInputNote}
            ranges={ranges} bonusThres={bonusThres}
            analyzeEntry={analyzeEntry} onSave={saveEntry}
            schedule={schedule} saving={saving}
          />
        )}
        {page==="stats" && (
          <StatsPage viewY={viewY} viewM={viewM}
            setViewY={setViewY} setViewM={setViewM}
            sessions={sessions} schedule={schedule}
            bonusThres={bonusThres}
          />
        )}
        {page==="settings" && (
          <SettingsPage
            ranges={ranges} bonusThres={bonusThres}
            onSave={saveSettings} uploadSchedule={uploadSchedule}
            showToast={showToast}
          />
        )}
      </main>
    </div>
  );
}

// ══════════════════════════════════════════════
// CALENDAR PAGE
// ══════════════════════════════════════════════
function CalendarPage({ viewY,viewM,setViewY,setViewM,sessions,schedule,analyzeEntry,chipStatus,onClickCell }) {
  const days = getDaysInMonth(viewY, viewM);
  // 週一為起始（0=週一, 5=週六，週日不顯示）
  const rawFirst = getFirstDayOfWeek(viewY, viewM); // 0=日,1=一,...,6=六
  const firstDay = rawFirst === 0 ? 6 : rawFirst - 1; // 轉換為週一起始
  const cells = Array(firstDay).fill(null).concat(Array.from({length:days},(_,i)=>i+1));
  // 過濾掉週日（原始weekday===0）
  const filteredCells = cells.filter((d, i) => {
    if (d === null) {
      // 空格：計算對應的weekday
      const weekday = i % 7; // 0=週一,1=週二,...,5=週六
      return weekday < 6; // 只保留週一到週六
    }
    // 有日期：計算該日是週幾
    const date = new Date(viewY, viewM, d);
    return date.getDay() !== 0; // 0=週日，排除
  });
  while (filteredCells.length % 6 !== 0) filteredCells.push(null);
  const cells2 = filteredCells;

  function prevMonth() { if(viewM===0){setViewY(y=>y-1);setViewM(11);}else setViewM(m=>m-1); }
  function nextMonth() { if(viewM===11){setViewY(y=>y+1);setViewM(0);}else setViewM(m=>m+1); }

  const today = new Date();
  const isToday = d => d && viewY===today.getFullYear() && viewM===today.getMonth() && d===today.getDate();

  return (
    <div>
      <div style={S.calHeader}>
        <button onClick={prevMonth} style={S.arrowBtn}>‹</button>
        <span style={S.monthTitle}>{viewY} 年 {viewM+1} 月</span>
        <button onClick={nextMonth} style={S.arrowBtn}>›</button>
      </div>
      <div style={S.weekRow}>
        {["一","二","三","四","五","六"].map(w=>(
          <div key={w} style={S.weekLabel}>{w}</div>
        ))}
      </div>
      <div style={S.calGrid}>
        {cells2.map((d,i) => {
          if (!d) return <div key={`e${i}`} style={S.emptyCell}/>;
          const dk = dateKey(viewY, viewM, d);
          const dayData = sessions[dk] || {};
          const hasSched = !!schedule[dk];

          return (
            <div key={dk} style={{ ...S.dayCell, ...(isToday(d)?S.dayCellToday:{}) }}>
              <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:3 }}>
                <span style={S.dayNum}>{d}</span>
                {hasSched && <span style={S.schedBadge} title="已上傳班表">📋</span>}
              </div>
              <div style={S.sessRow}>
                {SESSIONS.map(sess => {
                  const info = dayData[sess];
                  const status = chipStatus(dk, sess);
                  const a = analyzeEntry(dk, sess);
                  // 從班表取人員量
                  const schedCodes = schedule[dk]?.[sess];
                  const schedStaff = schedCodes ? codestoStaffOption(schedCodes) : null;
                  const displayStaff = info?.staff || schedStaff;
                  return (
                    <div key={sess} onClick={()=>onClickCell(dk,sess)}
                      style={{ ...S.sessChip, ...chipColor(status) }}
                      title={info?.note||""}>
                      <span style={S.sessLabel}>{sess}</span>
                      <span style={S.sessCount}>{info ? info.count : (schedStaff?"—":"·")}</span>
                      {displayStaff && <span style={S.staffPill}>{displayStaff}</span>}
                      {info?.note && <span style={{fontSize:9}}>📝</span>}
                    </div>
                  );
                })}
              </div>
              {SESSIONS.map(sess => {
                const a = analyzeEntry(dk, sess);
                if (!a||!a.avgWarn||a.avg===null) return null;
                return <div key={`av${sess}`} style={S.avgWarning}>{sess}均{a.avg.toFixed(1)}⚠️</div>;
              })}
            </div>
          );
        })}
      </div>
      <div style={S.legend}>
        {[["ok","正常"],["bonus","超獎金"],["avg_warn","超出範圍"],["neutral","未設範圍"],["empty","未輸入"]].map(([s,l])=>(
          <div key={s} style={S.legendItem}><div style={{ ...S.legendDot,...chipColor(s) }}/>{l}</div>
        ))}
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════
// INPUT PAGE
// ══════════════════════════════════════════════
function InputPage({ inputDate,setInputDate,inputSess,setInputSess,inputCount,setInputCount,
  inputStaff,setInputStaff,inputNote,setInputNote,ranges,bonusThres,analyzeEntry,onSave,schedule,saving }) {

  const st = sessionType(inputSess);
  const range = ranges[st]?.[inputStaff];
  const thres = bonusThres[st]?.[inputStaff];
  const countNum = inputCount!=="" ? Number(inputCount) : null;
  const ok = countNum!==null ? inRange(countNum, range) : null;
  const isBonus = countNum!==null ? aboveBonus(countNum, thres) : false;
  const analysis = analyzeEntry(inputDate, inputSess);

  // 班表提示
  const schedCodes = schedule[inputDate]?.[inputSess];
  const schedStaff = schedCodes ? codestoStaffOption(schedCodes) : null;
  const therapists = schedCodes ? schedCodes.filter(c=>!ASSISTANT_CODES.includes(c)) : [];
  const assistants = schedCodes ? schedCodes.filter(c=>ASSISTANT_CODES.includes(c)) : [];

  function dispStyle() {
    if (countNum===null) return {};
    if (isBonus) return { background:"#fef9c3",color:"#854d0e",borderColor:"#fbbf24" };
    if (ok===false) return { background:"#fee2e2",color:"#991b1b",borderColor:"#fca5a5" };
    if (ok===true)  return { background:"#d1fae5",color:"#065f46",borderColor:"#6ee7b7" };
    return {};
  }

  function handleNum(n) {
    setInputCount(prev => {
      if (n==="⌫") return prev.slice(0,-1);
      if (n==="C") return "";
      const next = prev+n;
      return next.length<=4 ? next : prev;
    });
  }

  return (
    <div style={S.inputWrap}>
      {/* 日期 + 節次 */}
      <div style={S.inputTopRow}>
        <input type="date" value={inputDate} onChange={e=>setInputDate(e.target.value)} style={S.dateInput}/>
        <div style={S.sessSelect}>
          {SESSIONS.map(s=>(
            <button key={s} onClick={()=>setInputSess(s)}
              style={{ ...S.sessBtn, ...(inputSess===s?S.sessBtnActive:{}) }}>{s}班</button>
          ))}
        </div>
      </div>

      {/* 班表人員提示 */}
      {schedCodes && (
        <div style={S.schedInfo}>
          <span style={S.schedInfoTitle}>📋 班表：</span>
          {therapists.length>0 && <span>治療師 {therapists.join(" ")}　</span>}
          {assistants.length>0 && <span style={{color:"#7c3aed"}}>助理 {assistants.join(" ")}</span>}
          <span style={S.schedInfoStaff}>→ {schedStaff}</span>
        </div>
      )}

      {/* 治療人員量 */}
      <div style={S.staffRow}>
        <span style={S.staffLabel}>治療人員量</span>
        <div style={S.staffBtns}>
          {STAFF_OPTIONS.map(s=>(
            <button key={s} onClick={()=>setInputStaff(s)}
              style={{ ...S.staffBtn, ...(inputStaff===s?S.staffBtnActive:{}),
                       ...(s===schedStaff?{boxShadow:"0 0 0 2px #7c3aed"}:{}) }}>
              {s}{s===schedStaff?" ✓":""}
            </button>
          ))}
        </div>
      </div>

      {/* 緩衝範圍 + 獎金門檻 */}
      <div style={S.rangeBox}>
        <div>
          <span style={S.rangeLabel}>{inputSess==="C"?"C班":"AB班"} × {inputStaff}　緩衝區間：</span>
          <span style={S.rangeVal}>
            {(range&&range[0]!==""&&range[1]!=="") ? `${range[0]} ～ ${range[1]}` : "未設定"}
          </span>
        </div>
        <div style={{marginTop:4}}>
          <span style={S.rangeLabel}>超次獎金門檻：</span>
          <span style={{...S.rangeVal,color:"#b45309"}}>
            {(thres!==""&&thres!=null)?`> ${thres}`:"未設定"}
          </span>
        </div>
      </div>

      {/* 大數字 */}
      <div style={{ ...S.bigDisplay,...dispStyle(),border:`2px solid ${dispStyle().borderColor||"#e2e8f0"}` }}>
        <span style={S.bigNum}>{inputCount||"—"}</span>
        {isBonus && <span style={S.bonusBadge}>🏅 超次獎金！</span>}
        {!isBonus && ok===true  && <span style={S.okBadge}>✓ 正常範圍</span>}
        {!isBonus && ok===false && <span style={S.warnBadge}>⚠ 超出範圍</span>}
      </div>

      {/* 鍵盤 */}
      <div style={S.numpad}>
        {["7","8","9","4","5","6","1","2","3","C","0","⌫"].map(n=>(
          <button key={n} onClick={()=>handleNum(n)}
            style={{ ...S.numKey,
              ...(n==="C"?{background:"#fee2e2",color:"#991b1b"}:
                  n==="⌫"?{background:"#fef3c7",color:"#92400e"}:{}) }}>
            {n}
          </button>
        ))}
      </div>

      {/* 備註 */}
      <div style={S.noteRow}>
        <span style={S.noteLabel}>📝 單節備註</span>
        <textarea value={inputNote} onChange={e=>setInputNote(e.target.value)}
          placeholder="選填：輸入本節備註…" style={S.noteInput} rows={2}/>
      </div>

      <button onClick={onSave} style={{...S.saveBtn, opacity:saving?0.7:1}} disabled={inputCount===" "||saving}>{saving?"儲存中…":"💾 儲存此節"}</button>

      {/* 前三週分析 */}
      {analysis && (
        <div style={S.analysisBox}>
          <div style={S.analysisTitle}>📊 前三週同節平均分析</div>
          {analysis.avg!==null ? (
            <div style={{...S.analysisRow,color:analysis.avgWarn?"#dc2626":"#065f46",fontWeight:600}}>
              前三週平均：{analysis.avg.toFixed(1)}
              {analysis.avgWarn && <span style={S.avgAlert}>　⚠️ 連續三週超出緩衝範圍！</span>}
              {!analysis.avgWarn && analysis.avgOk===true && <span style={{color:"#059669"}}>　✓ 正常</span>}
            </div>
          ) : (
            <div style={S.analysisRow}>前三週資料不足（需完整三週）</div>
          )}
          {range&&range[0]!==""&&<div style={{...S.analysisRow,color:"#64748b"}}>緩衝區間：{range[0]}～{range[1]}</div>}
          {thres!==""&&thres!=null&&<div style={{...S.analysisRow,color:"#b45309"}}>超次門檻：＞{thres}</div>}
        </div>
      )}
    </div>
  );
}

// ══════════════════════════════════════════════
// STATS PAGE - 超次統計
// ══════════════════════════════════════════════
function StatsPage({ viewY,viewM,setViewY,setViewM,sessions,schedule,bonusThres }) {
  function prev() { if(viewM===0){setViewY(y=>y-1);setViewM(11);}else setViewM(m=>m-1); }
  function next() { if(viewM===11){setViewY(y=>y+1);setViewM(0);}else setViewM(m=>m+1); }

  // 計算當月每位治療師的超次節次明細
  const days = getDaysInMonth(viewY, viewM);
  const overtimeMap = {}; // { code: [ {dk, sess, count, thres} ] }
  const details = [];

  for (let d = 1; d <= days; d++) {
    const dk = dateKey(viewY, viewM, d);
    const sched = schedule[dk];
    if (!sched) continue;
    for (const sess of SESSIONS) {
      const entry = sessions[dk]?.[sess];
      if (!entry) continue;
      const st = sessionType(sess);
      const thres = bonusThres[st]?.[entry.staff];
      if (!aboveBonus(entry.count, thres)) continue;
      const codes = sched[sess] || [];
      const therapists = codes.filter(c => !ASSISTANT_CODES.includes(c));
      details.push({ dk, sess, count: entry.count, staff: entry.staff, thres, therapists });
      for (const code of therapists) {
        if (!overtimeMap[code]) overtimeMap[code] = [];
        overtimeMap[code].push({ dk, sess, count: entry.count });
      }
    }
  }

  const sortedTherapists = Object.entries(overtimeMap).sort((a,b) => b[1].length - a[1].length);

  return (
    <div style={{ maxWidth:680, margin:"0 auto" }}>
      <div style={S.calHeader}>
        <button onClick={prev} style={S.arrowBtn}>‹</button>
        <span style={S.monthTitle}>{viewY} 年 {viewM+1} 月　超次統計</span>
        <button onClick={next} style={S.arrowBtn}>›</button>
      </div>

      {/* 每人超次次數 */}
      <div style={S.bonusSection}>
        <div style={S.bonusSectionTitle}>📊 每位治療師本月超次節數</div>
        {sortedTherapists.length === 0 ? (
          <div style={{ color:"#94a3b8", padding:"20px 0", textAlign:"center" }}>
            本月尚無超次記錄
          </div>
        ) : (
          <div style={S.bonusGrid}>
            {sortedTherapists.map(([code, list]) => (
              <div key={code} style={S.bonusCard}>
                <div style={S.bonusCode}>{code}</div>
                <div style={S.bonusAmount}>
                  <span style={{ fontSize:32, fontWeight:900 }}>{list.length}</span>
                  <span style={{ fontSize:14, marginLeft:4 }}>節</span>
                </div>
                <div style={{ fontSize:11, color:"#92400e", marginTop:4 }}>
                  {list.map(r => `${r.dk.slice(5)} ${r.sess}班`).join("、")}
                </div>
              </div>
            ))}
          </div>
        )}
        {sortedTherapists.length > 0 && (
          <div style={{ ...S.bonusTotal, color:"#854d0e" }}>
            本月超次總節數：<strong>{details.length} 節</strong>
          </div>
        )}
      </div>

      {/* 超次明細 */}
      {details.length > 0 && (
        <div style={S.bonusSection}>
          <div style={S.bonusSectionTitle}>📋 超次明細</div>
          <table style={S.table}>
            <thead>
              <tr style={S.thead}>
                <th style={S.th}>日期</th>
                <th style={S.th}>節次</th>
                <th style={S.th}>人員量</th>
                <th style={S.th}>人次</th>
                <th style={S.th}>門檻</th>
                <th style={S.th}>治療師</th>
              </tr>
            </thead>
            <tbody>
              {details.map((r,i) => (
                <tr key={i} style={{ background:i%2===0?"#f8fafc":"#fff" }}>
                  <td style={S.td}>{r.dk.slice(5)}</td>
                  <td style={S.td}>{r.sess}班</td>
                  <td style={S.td}>{r.staff}</td>
                  <td style={{ ...S.td, fontWeight:700, color:"#854d0e" }}>{r.count}</td>
                  <td style={S.td}>{r.thres}</td>
                  <td style={S.td}>{r.therapists.join(" ") || "—"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

// ══════════════════════════════════════════════
// SETTINGS PAGE
// ══════════════════════════════════════════════
function SettingsPage({ ranges,bonusThres,onSave,uploadSchedule,showToast }) {
  const [lRanges, setLRanges] = useState(()=>JSON.parse(JSON.stringify(ranges)));
  const [lThres,  setLThres]  = useState(()=>JSON.parse(JSON.stringify(bonusThres)));
  const fileRef = useRef();

  // 同步 props 變更（Firebase 載入後）
  useEffect(()=>{ setLRanges(JSON.parse(JSON.stringify(ranges))); },[ranges]);
  useEffect(()=>{ setLThres(JSON.parse(JSON.stringify(bonusThres))); },[bonusThres]);

  function setVal(setter, type, staff, key, val) {
    setter(prev => {
      const next = JSON.parse(JSON.stringify(prev));
      if (!next[type]) next[type] = {};
      if (key==="range0"||key==="range1") {
        if (!next[type][staff]) next[type][staff] = ["",""];
        next[type][staff][key==="range0"?0:1] = val===""?"":Number(val);
      } else {
        next[type][staff] = val===""?"":Number(val);
      }
      return next;
    });
  }

  async function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    await uploadSchedule(file);
    e.target.value = "";
  }

  return (
    <div style={{ maxWidth:720,margin:"0 auto" }}>
      <h2 style={S.settingsTitle}>⚙️ 系統設定</h2>

      {/* 班表上傳 */}
      <div style={S.settingsSection}>
        <div style={S.sectionTitle}>📤 上傳班表 Excel</div>
        <p style={{ color:"#64748b",fontSize:14,marginBottom:12 }}>
          上傳 PT 工作表的班表 Excel，系統會自動解析每節治療人員量。<br/>
          助理代碼：<strong>R、U、W</strong>（出現即計為 +1）
        </p>
        <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleFile} style={{display:"none"}}/>
        <button onClick={()=>fileRef.current.click()} style={S.uploadBtn}>
          📂 選擇 Excel 班表
        </button>
      </div>

      {/* 各班設定 */}
      {["AB","C"].map(type => (
        <div key={type} style={S.settingsSection}>
          <div style={S.sectionTitle}>{type==="AB"?"🌅 AB班":"🌙 C班"}</div>
          <div style={{ overflowX:"auto" }}>
            <table style={{ ...S.table,marginTop:0 }}>
              <thead>
                <tr style={S.thead}>
                  <th style={S.th}>人員量</th>
                  <th style={S.th}>緩衝下限</th>
                  <th style={S.th}>緩衝上限</th>
                  <th style={{ ...S.th,color:"#b45309",background:"#fffbeb" }}>超次門檻</th>
                </tr>
              </thead>
              <tbody>
                {STAFF_OPTIONS.map((staff,i) => (
                  <tr key={staff} style={{ background:i%2===0?"#f8fafc":"#fff" }}>
                    <td style={{ ...S.td,fontWeight:700,color:"#0369a1" }}>{staff}</td>
                    <td style={S.td}>
                      <input type="number" min="0" max="999"
                        value={lRanges[type]?.[staff]?.[0]??""}
                        onChange={e=>setVal(setLRanges,type,staff,"range0",e.target.value)}
                        style={S.settingInput} placeholder="—"/>
                    </td>
                    <td style={S.td}>
                      <input type="number" min="0" max="999"
                        value={lRanges[type]?.[staff]?.[1]??""}
                        onChange={e=>setVal(setLRanges,type,staff,"range1",e.target.value)}
                        style={S.settingInput} placeholder="—"/>
                    </td>
                    <td style={S.td}>
                      <input type="number" min="0" max="999"
                        value={lThres[type]?.[staff]??""}
                        onChange={e=>setVal(setLThres,type,staff,"val",e.target.value)}
                        style={{ ...S.settingInput,borderColor:"#fbbf24",background:"#fffbeb" }} placeholder="—"/>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      ))}

      <div style={S.settingsBtnRow}>
        <button onClick={()=>{ setLRanges({AB:{},C:{}}); setLThres({AB:{},C:{}}); }}
          style={S.resetBtn}>↩ 清空所有</button>
        <button onClick={()=>onSave(lRanges,lThres)} style={S.saveSettingsBtn}>✓ 儲存設定</button>
      </div>
    </div>
  );
}

// ──────────────────────────────────────────────
// chipColor helper
// ──────────────────────────────────────────────
function chipColor(status) {
  switch(status) {
    case "bonus":    return { background:"#fef9c3",color:"#854d0e",border:"1.5px solid #fbbf24" };
    case "avg_warn": return { background:"#fee2e2",color:"#991b1b",border:"1.5px solid #f87171" };
    case "ok":       return { background:"#d1fae5",color:"#065f46",border:"1.5px solid #6ee7b7" };
    case "neutral":  return { background:"#e0f2fe",color:"#0c4a6e",border:"1.5px solid #7dd3fc" };
    default:         return { background:"#f1f5f9",color:"#94a3b8",border:"1px solid #e2e8f0" };
  }
}

// ──────────────────────────────────────────────
// STYLES
// ──────────────────────────────────────────────
const S = {
  root:{ fontFamily:"'Noto Sans TC','PingFang TC',sans-serif",minHeight:"100vh",
    background:"linear-gradient(135deg,#f0f9ff 0%,#e0f2fe 50%,#f0fdf4 100%)",color:"#1e293b" },
  nav:{ background:"linear-gradient(90deg,#0ea5e9,#10b981)",padding:"12px 20px",
    display:"flex",alignItems:"center",justifyContent:"space-between",
    boxShadow:"0 2px 12px rgba(0,0,0,0.15)",flexWrap:"wrap",gap:8 },
  navBrand:{ color:"#fff",fontWeight:700,fontSize:18 },
  navTabs:{ display:"flex",gap:6,flexWrap:"wrap" },
  navBtn:{ background:"rgba(255,255,255,0.2)",color:"#fff",border:"none",
    borderRadius:20,padding:"6px 12px",cursor:"pointer",fontSize:13,fontWeight:500 },
  navBtnActive:{ background:"rgba(255,255,255,0.95)",color:"#0369a1",fontWeight:700 },
  main:{ padding:"16px",maxWidth:960,margin:"0 auto" },

  // Calendar
  calHeader:{ display:"flex",alignItems:"center",justifyContent:"center",gap:20,marginBottom:12 },
  arrowBtn:{ background:"#0ea5e9",color:"#fff",border:"none",borderRadius:8,
    width:36,height:36,fontSize:22,cursor:"pointer",lineHeight:1 },
  monthTitle:{ fontSize:22,fontWeight:700,color:"#0c4a6e" },
  weekRow:{ display:"grid",gridTemplateColumns:"repeat(6,1fr)",gap:2,marginBottom:2 },
  weekLabel:{ textAlign:"center",fontWeight:600,color:"#64748b",fontSize:13,padding:"4px 0" },
  calGrid:{ display:"grid",gridTemplateColumns:"repeat(6,1fr)",gap:3 },
  emptyCell:{ minHeight:110 },
  dayCell:{ background:"#fff",borderRadius:10,padding:6,minHeight:110,
    boxShadow:"0 1px 4px rgba(0,0,0,0.06)",border:"1px solid #e2e8f0",overflow:"hidden" },
  dayCellToday:{ border:"2px solid #0ea5e9",background:"#f0f9ff" },
  dayNum:{ fontWeight:700,fontSize:14,color:"#334155" },
  schedBadge:{ fontSize:10,opacity:0.7 },
  sessRow:{ display:"flex",flexDirection:"column",gap:2 },
  sessChip:{ display:"flex",alignItems:"center",gap:2,borderRadius:5,
    padding:"2px 4px",fontSize:10,cursor:"pointer",flexWrap:"wrap" },
  sessLabel:{ fontWeight:700,minWidth:12 },
  sessCount:{ fontWeight:700,minWidth:16 },
  staffPill:{ fontSize:9,background:"rgba(0,0,0,0.08)",borderRadius:3,
    padding:"0 3px",marginLeft:"auto" },
  avgWarning:{ fontSize:9,color:"#dc2626",marginTop:2,background:"#fff7f7",
    borderRadius:4,padding:"1px 4px" },
  legend:{ display:"flex",gap:10,justifyContent:"center",marginTop:14,flexWrap:"wrap" },
  legendItem:{ display:"flex",alignItems:"center",gap:4,fontSize:11,color:"#64748b" },
  legendDot:{ width:12,height:12,borderRadius:3 },

  // Input
  inputWrap:{ display:"flex",flexDirection:"column",gap:14,maxWidth:440,margin:"0 auto" },
  inputTopRow:{ display:"flex",gap:10,alignItems:"center",flexWrap:"wrap" },
  dateInput:{ border:"2px solid #bae6fd",borderRadius:8,padding:"8px 12px",
    fontSize:16,flex:1,minWidth:140,outline:"none" },
  sessSelect:{ display:"flex",gap:6 },
  sessBtn:{ border:"2px solid #bae6fd",borderRadius:20,padding:"6px 14px",
    background:"#fff",cursor:"pointer",fontWeight:600,fontSize:15 },
  sessBtnActive:{ background:"#0ea5e9",color:"#fff",border:"2px solid #0ea5e9" },
  schedInfo:{ background:"#ede9fe",borderRadius:10,padding:"8px 12px",
    fontSize:13,color:"#4c1d95",border:"1px solid #c4b5fd" },
  schedInfoTitle:{ fontWeight:700 },
  schedInfoStaff:{ fontWeight:700,color:"#6d28d9",marginLeft:8 },
  staffRow:{ display:"flex",alignItems:"center",gap:10,flexWrap:"wrap" },
  staffLabel:{ fontWeight:600,color:"#475569",minWidth:70 },
  staffBtns:{ display:"flex",gap:6,flexWrap:"wrap" },
  staffBtn:{ border:"2px solid #d1d5db",borderRadius:8,padding:"5px 12px",
    background:"#f8fafc",cursor:"pointer",fontWeight:600,fontSize:14 },
  staffBtnActive:{ background:"#10b981",color:"#fff",border:"2px solid #10b981" },
  rangeBox:{ background:"#f0fdf4",borderRadius:10,padding:"10px 14px",
    border:"1px solid #6ee7b7",fontSize:14 },
  rangeLabel:{ color:"#065f46",fontWeight:500 },
  rangeVal:{ fontWeight:700,color:"#059669",fontSize:15 },
  bigDisplay:{ borderRadius:16,background:"#f1f5f9",minHeight:110,display:"flex",
    alignItems:"center",justifyContent:"center",flexDirection:"column",gap:4,
    border:"2px solid #e2e8f0",transition:"background .25s,color .25s" },
  bigNum:{ fontSize:76,fontWeight:900,lineHeight:1 },
  okBadge:{ fontSize:17,fontWeight:600,color:"#065f46" },
  bonusBadge:{ fontSize:17,fontWeight:700,color:"#854d0e" },
  warnBadge:{ fontSize:17,fontWeight:600,color:"#991b1b" },
  numpad:{ display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8 },
  numKey:{ height:64,fontSize:28,fontWeight:700,border:"2px solid #e2e8f0",
    borderRadius:12,background:"#fff",cursor:"pointer",boxShadow:"0 2px 4px rgba(0,0,0,0.06)" },
  noteRow:{ display:"flex",flexDirection:"column",gap:6 },
  noteLabel:{ fontWeight:600,color:"#475569",fontSize:14 },
  noteInput:{ border:"2px solid #e2e8f0",borderRadius:10,padding:"8px 12px",
    fontSize:15,outline:"none",resize:"vertical",fontFamily:"inherit",color:"#334155",background:"#fff" },
  saveBtn:{ background:"linear-gradient(90deg,#0ea5e9,#10b981)",color:"#fff",border:"none",
    borderRadius:12,padding:"16px",fontSize:18,fontWeight:700,cursor:"pointer",
    boxShadow:"0 4px 12px rgba(14,165,233,0.3)" },
  analysisBox:{ background:"#fff",borderRadius:12,padding:"14px 16px",
    border:"1px solid #e2e8f0",boxShadow:"0 1px 4px rgba(0,0,0,0.06)" },
  analysisTitle:{ fontWeight:700,color:"#334155",marginBottom:8,fontSize:15 },
  analysisRow:{ fontSize:14,color:"#475569",marginBottom:4,lineHeight:1.7 },
  avgAlert:{ color:"#dc2626",fontWeight:700 },

  // Bonus
  bonusSection:{ background:"#fff",borderRadius:14,padding:20,marginBottom:16,
    boxShadow:"0 1px 6px rgba(0,0,0,0.06)" },
  bonusSectionTitle:{ fontWeight:700,fontSize:17,color:"#0369a1",marginBottom:14 },
  bonusGrid:{ display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(130px,1fr))",gap:10 },
  bonusCard:{ background:"linear-gradient(135deg,#fef9c3,#fef3c7)",borderRadius:12,
    padding:"12px 16px",textAlign:"center",border:"1.5px solid #fbbf24" },
  bonusCode:{ fontWeight:700,fontSize:20,color:"#854d0e",marginBottom:4 },
  bonusAmount:{ fontWeight:700,fontSize:16,color:"#92400e" },
  bonusTotal:{ marginTop:14,padding:"10px 14px",background:"#f0fdf4",borderRadius:8,
    fontSize:15,color:"#065f46",border:"1px solid #6ee7b7",textAlign:"right" },

  // Settings
  settingsTitle:{ fontSize:22,fontWeight:700,color:"#0c4a6e",marginBottom:4 },
  settingsSection:{ background:"#fff",borderRadius:14,padding:18,marginBottom:16,
    boxShadow:"0 1px 6px rgba(0,0,0,0.06)" },
  sectionTitle:{ fontWeight:700,fontSize:17,color:"#0369a1",marginBottom:12 },
  uploadBtn:{ background:"linear-gradient(90deg,#7c3aed,#6d28d9)",color:"#fff",border:"none",
    borderRadius:10,padding:"10px 24px",cursor:"pointer",fontWeight:700,fontSize:15,
    boxShadow:"0 3px 10px rgba(109,40,217,0.3)" },
  settingsBtnRow:{ display:"flex",gap:12,justifyContent:"flex-end",marginTop:8 },
  resetBtn:{ border:"2px solid #e2e8f0",borderRadius:10,padding:"10px 20px",
    background:"#fff",cursor:"pointer",fontWeight:600,color:"#64748b",fontSize:15 },
  saveSettingsBtn:{ background:"linear-gradient(90deg,#0ea5e9,#10b981)",color:"#fff",border:"none",
    borderRadius:10,padding:"10px 24px",cursor:"pointer",fontWeight:700,fontSize:15,
    boxShadow:"0 3px 10px rgba(14,165,233,0.3)" },

  // Table
  table:{ width:"100%",borderCollapse:"collapse",fontSize:13 },
  thead:{ background:"#f1f5f9" },
  th:{ padding:"8px 10px",fontWeight:600,color:"#475569",textAlign:"center",
    borderBottom:"2px solid #e2e8f0" },
  td:{ padding:"7px 10px",textAlign:"center",borderBottom:"1px solid #f1f5f9" },
  settingInput:{ border:"2px solid #bae6fd",borderRadius:8,padding:"6px 4px",
    fontSize:15,textAlign:"center",outline:"none",width:"100%",boxSizing:"border-box" },
};
