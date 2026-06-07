import { useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";

// ─── CONSTANTS ────────────────────────────────────────────────────────────────
const ROOMS = ["Cyto 1", "Cyto 2", "CIVAS 1", "CIVAS 2"];

const TREATMENT_STAGES = [
  { id: "introduced", label: "Introduced to Room" },
  { id: "in_cabinet",  label: "In Cabinet" },
  { id: "prepared",   label: "Prep Complete" },
  { id: "checked",    label: "Checked / Labelled" },
  { id: "dispatched", label: "Dispatched" },
];

const STAGE_COLORS = {
  introduced: "#2563eb",
  in_cabinet: "#d97706",
  prepared:   "#7c3aed",
  checked:    "#0891b2",
  dispatched: "#059669",
};

const ROOM_ACCENT = {
  "Cyto 1":  "#2563eb",
  "Cyto 2":  "#7c3aed",
  "CIVAS 1": "#0891b2",
  "CIVAS 2": "#059669",
};

const ROOM_TYPE = room => room.startsWith("Cyto") ? "cyto" : "civas";

const PRESSURE_SPECS = {
  cyto: [
    { key:"internalPressure", label:"Internal Pressure", unit:"Pa",      min:-100, max:-70  },
    { key:"downflowVelocity", label:"Downflow Velocity",  unit:"m/s",    min:0.36, max:0.45 },
    { key:"filterChange",     label:"Filter Change",      unit:"Pa",     min:200,  max:320  },
    { key:"totalAirChanges",  label:"Total Air Changes",  unit:"TAC/hr", min:1710, max:2199 },
  ],
  civas: [
    { key:"internalPressure", label:"Internal Pressure", unit:"Pa",      min:-100, max:60   },
    { key:"downflowVelocity", label:"Downflow Velocity",  unit:"m/s",    min:0.36, max:0.45 },
    { key:"filterChange",     label:"Filter Change",      unit:"Pa",     min:60,   max:150  },
    { key:"totalAirChanges",  label:"Total Air Changes",  unit:"TAC/hr", min:624,  max:803  },
  ],
};

function inRange(val, spec) {
  const n = parseFloat(val);
  if (isNaN(n) || val === "") return null;
  return n >= spec.min && n <= spec.max;
}
function emptyReadings() {
  return { internalPressure:"", downflowVelocity:"", filterChange:"", totalAirChanges:"" };
}
function emptyPersonnel() {
  // Per-treatment cabinet personnel
  return { operative:"", checker:"", sprayedBy:"operative" };
}

function timestamp() { return new Date().toISOString(); }
function fmtTime(iso) {
  if (!iso) return "—";
  return new Date(iso).toLocaleTimeString("en-GB", { hour:"2-digit", minute:"2-digit" });
}
function fmtDate(iso) {
  if (!iso) return "—";
  return new Date(iso).toLocaleDateString("en-GB");
}
function todayKey() { return new Date().toISOString().slice(0,10); }

// ─── STORAGE ──────────────────────────────────────────────────────────────────
const STORAGE_KEY = "chemo_unit_data_v4";
async function loadData() {
  try { const r = await window.storage.get(STORAGE_KEY, true); return r ? JSON.parse(r.value) : null; }
  catch { return null; }
}
async function saveData(data) {
  try { await window.storage.set(STORAGE_KEY, JSON.stringify(data), true); }
  catch(e) { console.error("Save failed", e); }
}

function emptySession(sessionNum, date) {
  return {
    id:`${date}-S${sessionNum}`, date, sessionNum,
    startedAt:null, closedAt:null,
    rooms: Object.fromEntries(ROOMS.map(r=>[r, {
      cleaningBefore: { done:false, by:"", time:null },
      cleaningAfter:  { done:false, by:"", time:null },
      readingsStart: emptyReadings(),
      readingsEnd:   emptyReadings(),
    }])),
    treatments:[],
  };
}
function emptyAppData() {
  return { sessions:[emptySession(1,todayKey())], activeSessionIdx:0 };
}

// ─── EXCEL EXPORT ─────────────────────────────────────────────────────────────
function exportToExcel(session) {
  const wb = XLSX.utils.book_new();

  // Sheet 1 – Treatments
  const txHeaders = [
    "Patient ID","Drug / Regimen","Batch No.","Room",
    "Operative","In-Process Checker","Sprayed Into Cabinet By",
    "Introduced","In Cabinet","Prep Complete","Checked / Labelled","Dispatched",
    "Notes"
  ];
  const txRows = session.treatments.map(tx=>{
    const p = tx.personnel || emptyPersonnel();
    return [
      tx.patientId, tx.drug, tx.batch, tx.room,
      p.operative||"—", p.checker||"—",
      p.sprayedBy==="operative" ? `Operative (${p.operative||"?"})` : `Checker (${p.checker||"?"})`,
      fmtTime(tx.stageHistory.introduced),
      fmtTime(tx.stageHistory.in_cabinet),
      fmtTime(tx.stageHistory.prepared),
      fmtTime(tx.stageHistory.checked),
      fmtTime(tx.stageHistory.dispatched),
      tx.notes||"",
    ];
  });
  const ws1 = XLSX.utils.aoa_to_sheet([txHeaders,...txRows]);
  ws1["!cols"] = [14,24,16,9,16,18,24,12,12,14,16,12,22].map(w=>({wch:w}));
  XLSX.utils.book_append_sheet(wb, ws1, "Treatments");

  // Sheet 2 – Room Environmental Records
  const rows = [];
  rows.push(["ROOM ENVIRONMENTAL RECORDS"]);
  rows.push([`Session ${session.sessionNum}  —  ${fmtDate(session.date+"T00:00:00.000Z")}  —  Started: ${fmtTime(session.startedAt)}  Closed: ${fmtTime(session.closedAt)}`]);
  rows.push([]);

  ROOMS.forEach(room=>{
    const rd    = session.rooms[room];
    const type  = ROOM_TYPE(room);
    const specs = PRESSURE_SPECS[type];
    rows.push([`${room.toUpperCase()}  (${type==="cyto"?"CYTO ISOLATOR":"CIVAS ISOLATOR"})`]);
    rows.push(["Measurement","Unit","Acceptable Range","Session START Reading","In Range?","Session END Reading","In Range?"]);
    specs.forEach(spec=>{
      const sv=rd.readingsStart[spec.key], ev=rd.readingsEnd[spec.key];
      const si=inRange(sv,spec), ei=inRange(ev,spec);
      rows.push([
        spec.label, spec.unit, `${spec.min} – ${spec.max}`,
        sv||"", sv===""?"": si?"✓ IN RANGE":"✗ OUT OF RANGE",
        ev||"", ev===""?"": ei?"✓ IN RANGE":"✗ OUT OF RANGE",
      ]);
    });
    rows.push(["Cleaning Before Session",
      rd.cleaningBefore.done?"✓ Done":"Not recorded",
      rd.cleaningBefore.done?`By: ${rd.cleaningBefore.by}`:"",
      rd.cleaningBefore.done?`Time: ${fmtTime(rd.cleaningBefore.time)}`:"",
      "","","",
    ]);
    rows.push(["Cleaning After Session",
      rd.cleaningAfter.done?"✓ Done":"Not recorded",
      rd.cleaningAfter.done?`By: ${rd.cleaningAfter.by}`:"",
      rd.cleaningAfter.done?`Time: ${fmtTime(rd.cleaningAfter.time)}`:"",
      "","","",
    ]);
    rows.push([]);
  });

  const ws2 = XLSX.utils.aoa_to_sheet(rows);
  ws2["!cols"] = [26,10,18,20,14,18,14].map(w=>({wch:w}));
  XLSX.utils.book_append_sheet(wb, ws2, "Room Records");

  // Sheet 3 – Session Info
  const ws3 = XLSX.utils.aoa_to_sheet([
    ["Date",             fmtDate(session.date+"T00:00:00.000Z")],
    ["Session Number",   session.sessionNum],
    ["Started",          fmtTime(session.startedAt)],
    ["Closed",           fmtTime(session.closedAt)],
    ["Total Treatments", session.treatments.length],
  ]);
  ws3["!cols"] = [18,22].map(w=>({wch:w}));
  XLSX.utils.book_append_sheet(wb, ws3, "Session Info");

  XLSX.writeFile(wb, `AsepticUnit_${session.date}_Session${session.sessionNum}.xlsx`);
}

// ─── ICONS ────────────────────────────────────────────────────────────────────
const Icons = {
  Plus:     ()=><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>,
  Check:    ()=><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3"><polyline points="20 6 9 17 4 12"/></svg>,
  Clock:    ()=><svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>,
  Trash:    ()=><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6M14 11v6"/></svg>,
  Download: ()=><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>,
  Note:     ()=><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>,
  Person:   ()=><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>,
  Spray:    ()=><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M3 9h4l2-5h6l2 5h4"/><path d="M5 9v11a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V9"/><line x1="12" y1="9" x2="12" y2="20"/></svg>,
};

// ─── APP ──────────────────────────────────────────────────────────────────────
export default function App() {
  const [appData,   setAppData]   = useState(null);
  const [loading,   setLoading]   = useState(true);
  const [activeTab, setActiveTab] = useState("treatments");
  const [addTxForm, setAddTxForm] = useState({
    open:false, room:ROOMS[0], patientId:"", drug:"", batch:"", notes:"",
    operative:"", checker:"", sprayedBy:"operative"
  });
  const [editNotes,  setEditNotes]  = useState(null);
  const [lastSaved,  setLastSaved]  = useState(null);

  useEffect(()=>{
    (async()=>{
      const d = await loadData();
      setAppData(d || emptyAppData());
      setLoading(false);
    })();
    const id = setInterval(async()=>{ const d=await loadData(); if(d) setAppData(d); }, 8000);
    return ()=>clearInterval(id);
  },[]);

  const persist = useCallback(async(newData)=>{
    setAppData(newData);
    await saveData(newData);
    setLastSaved(new Date());
  },[]);

  if (loading||!appData) return (
    <div style={{background:"#f1f5f9",minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center"}}>
      <div style={{color:"#64748b",fontFamily:"monospace",fontSize:15}}>Loading unit data…</div>
    </div>
  );

  const session = appData.sessions[appData.activeSessionIdx];

  // ── Session ────────────────────────────────────────────────────────────────
  const newSession = async()=>{
    const date=todayKey(), num=appData.sessions.filter(s=>s.date===date).length+1;
    const s={...emptySession(num,date),startedAt:timestamp()};
    const sessions=[...appData.sessions,s];
    await persist({...appData,sessions,activeSessionIdx:sessions.length-1});
  };
  const startCurrentSession = async()=>{
    const sessions=appData.sessions.map((s,i)=>i===appData.activeSessionIdx?{...s,startedAt:timestamp()}:s);
    await persist({...appData,sessions});
  };
  const closeSession = async()=>{
    const sessions=appData.sessions.map((s,i)=>i===appData.activeSessionIdx?{...s,closedAt:timestamp()}:s);
    await persist({...appData,sessions});
  };

  // ── Treatments ────────────────────────────────────────────────────────────
  const addTreatment = async()=>{
    if(!addTxForm.patientId||!addTxForm.drug||!addTxForm.batch) return;
    const tx={
      id:Date.now().toString(), room:addTxForm.room,
      patientId:addTxForm.patientId, drug:addTxForm.drug, batch:addTxForm.batch,
      stage:"introduced", stageHistory:{introduced:timestamp()}, notes:addTxForm.notes||"",
      personnel:{ operative:addTxForm.operative, checker:addTxForm.checker, sprayedBy:addTxForm.sprayedBy },
    };
    const sessions=appData.sessions.map((s,i)=>i===appData.activeSessionIdx?{...s,treatments:[...s.treatments,tx]}:s);
    await persist({...appData,sessions});
    setAddTxForm({open:false,room:ROOMS[0],patientId:"",drug:"",batch:"",notes:"",operative:"",checker:"",sprayedBy:"operative"});
  };
  const advanceStage = async(txId)=>{
    const tx=session.treatments.find(t=>t.id===txId);
    const idx=TREATMENT_STAGES.findIndex(st=>st.id===tx?.stage);
    if(idx>=TREATMENT_STAGES.length-1) return;
    const nextStage=TREATMENT_STAGES[idx+1].id;
    const sessions=appData.sessions.map((s,i)=>i!==appData.activeSessionIdx?s:{...s,
      treatments:s.treatments.map(t=>t.id!==txId?t:{...t,stage:nextStage,stageHistory:{...t.stageHistory,[nextStage]:timestamp()}})
    });
    await persist({...appData,sessions});
  };
  const removeTreatment = async(txId)=>{
    const sessions=appData.sessions.map((s,i)=>i!==appData.activeSessionIdx?s:{...s,treatments:s.treatments.filter(t=>t.id!==txId)});
    await persist({...appData,sessions});
  };
  const saveNotes = async(txId,notes)=>{
    const sessions=appData.sessions.map((s,i)=>i!==appData.activeSessionIdx?s:{...s,
      treatments:s.treatments.map(t=>t.id!==txId?t:{...t,notes})
    });
    await persist({...appData,sessions});
    setEditNotes(null);
  };
  const updatePersonnel = async(txId,field,value)=>{
    const sessions=appData.sessions.map((s,i)=>i!==appData.activeSessionIdx?s:{...s,
      treatments:s.treatments.map(t=>t.id!==txId?t:{...t,personnel:{...(t.personnel||emptyPersonnel()),[field]:value}})
    });
    await persist({...appData,sessions});
  };

  // ── Rooms ─────────────────────────────────────────────────────────────────
  const updateReading = async(room,timing,key,value)=>{
    const field=timing==="start"?"readingsStart":"readingsEnd";
    const sessions=appData.sessions.map((s,i)=>{
      if(i!==appData.activeSessionIdx) return s;
      return {...s,rooms:{...s.rooms,[room]:{...s.rooms[room],[field]:{...s.rooms[room][field],[key]:value}}}};
    });
    await persist({...appData,sessions});
  };
  const markCleaning = async(room,when)=>{
    const initials=prompt(`Enter your initials for ${room} ${when} cleaning:`);
    if(!initials) return;
    const key=when==="before"?"cleaningBefore":"cleaningAfter";
    const sessions=appData.sessions.map((s,i)=>{
      if(i!==appData.activeSessionIdx) return s;
      return {...s,rooms:{...s.rooms,[room]:{...s.rooms[room],[key]:{done:true,by:initials.toUpperCase(),time:timestamp()}}}};
    });
    await persist({...appData,sessions});
  };

  // ─── RENDER ───────────────────────────────────────────────────────────────
  return (
    <div style={S.root}>

      {/* HEADER */}
      <header style={S.header}>
        <div style={{display:"flex",alignItems:"center",gap:12}}>
          <div style={S.logoMark}>⬡</div>
          <div>
            <div style={S.logoTitle}>ASEPTIC UNIT</div>
            <div style={S.logoSub}>Chemotherapy Production Tracker</div>
          </div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <div style={S.savedPill}>
            <span style={{color:lastSaved?"#059669":"#94a3b8",fontSize:9}}>●</span>
            {lastSaved?`Saved ${fmtTime(lastSaved.toISOString())}`:"Not saved"}
          </div>
          <div style={S.datePill}>{fmtDate(todayKey()+"T00:00:00.000Z")}</div>
        </div>
      </header>

      {/* SESSION BAR */}
      <div style={S.sessionBar}>
        <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
          {appData.sessions.map((s,i)=>(
            <button key={s.id} style={{...S.sessionTab,...(i===appData.activeSessionIdx?S.sessionTabActive:{})}}
              onClick={()=>setAppData({...appData,activeSessionIdx:i})}>
              <span style={{fontSize:10,color:"#94a3b8"}}>{fmtDate(s.date+"T00:00:00.000Z")}</span>
              <span style={{fontSize:13,fontWeight:600}}>Session {s.sessionNum}</span>
              {s.closedAt&&<span style={S.closedPill}>CLOSED</span>}
            </button>
          ))}
          <button style={S.newSessionBtn} onClick={newSession}><Icons.Plus/> New Session</button>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          {!session.startedAt?(
            <button style={S.btnGreen} onClick={startCurrentSession}>Start Session</button>
          ):!session.closedAt?(
            <>
              <span style={{fontSize:13,color:"#64748b",display:"flex",alignItems:"center",gap:6}}>
                <span style={{color:"#059669"}}>●</span> Active since {fmtTime(session.startedAt)}
              </span>
              <button style={S.btnRed} onClick={closeSession}>Close Session</button>
            </>
          ):(
            <span style={{fontSize:13,color:"#64748b",display:"flex",alignItems:"center",gap:6}}>
              <span style={{color:"#dc2626"}}>●</span> Closed {fmtTime(session.closedAt)}
            </span>
          )}
        </div>
      </div>

      {/* NAV */}
      <nav style={S.nav}>
        {[{id:"treatments",label:"Treatments"},{id:"rooms",label:"Rooms & Cleaning"},{id:"log",label:"Session Log"}].map(t=>(
          <button key={t.id} style={{...S.navTab,...(activeTab===t.id?S.navTabActive:{})}} onClick={()=>setActiveTab(t.id)}>
            {t.label}
          </button>
        ))}
      </nav>

      <main style={S.main}>

        {/* ══════ TREATMENTS ══════ */}
        {activeTab==="treatments" && (
          <div>
            {!session.closedAt&&(
              <div style={S.card}>
                {!addTxForm.open?(
                  <button style={S.btnBlue} onClick={()=>setAddTxForm({...addTxForm,open:true})}>
                    <Icons.Plus/> Add Treatment
                  </button>
                ):(
                  <div>
                    <div style={S.sectionHeading}>New Treatment</div>

                    {/* Row 1: treatment details */}
                    <div style={S.formGrid}>
                      <label style={S.label}>Room
                        <select style={S.input} value={addTxForm.room} onChange={e=>setAddTxForm({...addTxForm,room:e.target.value})}>
                          {ROOMS.map(r=><option key={r}>{r}</option>)}
                        </select>
                      </label>
                      <label style={S.label}>Patient ID / Name
                        <input style={S.input} value={addTxForm.patientId} onChange={e=>setAddTxForm({...addTxForm,patientId:e.target.value})} placeholder="e.g. PT-0042"/>
                      </label>
                      <label style={S.label}>Drug / Regimen
                        <input style={S.input} value={addTxForm.drug} onChange={e=>setAddTxForm({...addTxForm,drug:e.target.value})} placeholder="e.g. Carboplatin 450mg"/>
                      </label>
                      <label style={S.label}>Batch Number
                        <input style={S.input} value={addTxForm.batch} onChange={e=>setAddTxForm({...addTxForm,batch:e.target.value})} placeholder="e.g. BN-2024-1173"/>
                      </label>
                    </div>

                    {/* Row 2: cabinet personnel */}
                    <div style={S.personnelBox}>
                      <div style={S.personnelTitle}><Icons.Person/> Cabinet Personnel</div>
                      <div style={S.formGrid}>
                        <label style={S.label}>Operative
                          <input style={S.input} value={addTxForm.operative} onChange={e=>setAddTxForm({...addTxForm,operative:e.target.value})} placeholder="Initials / name"/>
                        </label>
                        <label style={S.label}>In-Process Checker (IPC)
                          <input style={S.input} value={addTxForm.checker} onChange={e=>setAddTxForm({...addTxForm,checker:e.target.value})} placeholder="Initials / name"/>
                        </label>
                        <div>
                          <div style={{...S.label,marginBottom:6}}>Sprayed into cabinet by</div>
                          <SprayToggle value={addTxForm.sprayedBy} onChange={v=>setAddTxForm({...addTxForm,sprayedBy:v})}/>
                        </div>
                      </div>
                    </div>

                    <label style={{...S.label,marginTop:10}}>Notes (optional)
                      <textarea style={{...S.input,resize:"vertical",minHeight:52,fontFamily:"inherit"}}
                        value={addTxForm.notes} onChange={e=>setAddTxForm({...addTxForm,notes:e.target.value})}
                        placeholder="Any preparation notes…"/>
                    </label>
                    <div style={{display:"flex",gap:8,marginTop:14}}>
                      <button style={S.btnGreen} onClick={addTreatment}>Add Treatment</button>
                      <button style={S.btnGhost} onClick={()=>setAddTxForm({...addTxForm,open:false})}>Cancel</button>
                    </div>
                  </div>
                )}
              </div>
            )}

            {ROOMS.map(room=>{
              const txs=session.treatments.filter(t=>t.room===room);
              if(!txs.length) return null;
              const accent=ROOM_ACCENT[room];
              return (
                <div key={room} style={S.roomSection}>
                  <div style={{...S.roomHeader,borderLeftColor:accent,color:accent}}>{room}</div>
                  {txs.map(tx=>{
                    const stageIdx=TREATMENT_STAGES.findIndex(s=>s.id===tx.stage);
                    const isComplete=tx.stage==="dispatched";
                    const isEditingNotes=editNotes===tx.id;
                    const p=tx.personnel||emptyPersonnel();
                    const sprayLabel=p.sprayedBy==="operative"
                      ? `Operative${p.operative?` (${p.operative})`:""}`
                      : `IPC${p.checker?` (${p.checker})` :""}`;
                    return (
                      <div key={tx.id} style={{...S.txCard,opacity:isComplete?.82:1,borderLeftColor:accent}}>
                        <div style={S.txTop}>
                          <div>
                            <div style={S.txPatient}>{tx.patientId}</div>
                            <div style={S.txDrug}>{tx.drug}</div>
                            <div style={S.txBatch}>Batch: {tx.batch}</div>
                          </div>
                          <div style={{display:"flex",gap:7,alignItems:"flex-start",flexWrap:"wrap",justifyContent:"flex-end"}}>
                            {!session.closedAt&&!isComplete&&(
                              <button style={{...S.btnBlue,fontSize:12,padding:"4px 10px"}} onClick={()=>advanceStage(tx.id)}>
                                → Advance
                              </button>
                            )}
                            {!session.closedAt&&(
                              <button style={{...S.btnGhost,fontSize:12,padding:"4px 8px",display:"flex",alignItems:"center",gap:4}}
                                onClick={()=>setEditNotes(isEditingNotes?null:tx.id)}>
                                <Icons.Note/> {tx.notes?"Edit":"Add"} Note
                              </button>
                            )}
                            {!session.closedAt&&(
                              <button style={S.btnDanger} onClick={()=>removeTreatment(tx.id)}><Icons.Trash/></button>
                            )}
                          </div>
                        </div>

                        {/* Cabinet personnel strip */}
                        <div style={S.personnelStrip}>
                          <div style={S.personnelField}>
                            <span style={S.personnelFieldLabel}>Operative</span>
                            {session.closedAt
                              ? <span style={S.personnelValue}>{p.operative||"—"}</span>
                              : <input style={S.personnelInput} value={p.operative||""} placeholder="Initials"
                                  onChange={e=>updatePersonnel(tx.id,"operative",e.target.value)}/>
                            }
                          </div>
                          <div style={S.personnelField}>
                            <span style={S.personnelFieldLabel}>IPC Checker</span>
                            {session.closedAt
                              ? <span style={S.personnelValue}>{p.checker||"—"}</span>
                              : <input style={S.personnelInput} value={p.checker||""} placeholder="Initials"
                                  onChange={e=>updatePersonnel(tx.id,"checker",e.target.value)}/>
                            }
                          </div>
                          <div style={S.personnelField}>
                            <span style={S.personnelFieldLabel}>Sprayed into cabinet by</span>
                            {session.closedAt
                              ? <span style={{...S.personnelValue,color:accent}}>{sprayLabel}</span>
                              : <SprayToggle value={p.sprayedBy||"operative"}
                                  onChange={v=>updatePersonnel(tx.id,"sprayedBy",v)}/>
                            }
                          </div>
                        </div>

                        {/* Notes */}
                        {isEditingNotes&&(
                          <NoteEditor initial={tx.notes} onSave={v=>saveNotes(tx.id,v)} onCancel={()=>setEditNotes(null)}/>
                        )}
                        {!isEditingNotes&&tx.notes&&(
                          <div style={S.noteBubble}><Icons.Note/> <span>{tx.notes}</span></div>
                        )}

                        {/* Stage track */}
                        <div style={S.stageTrack}>
                          {TREATMENT_STAGES.map((st,idx)=>{
                            const done=idx<=stageIdx, active=idx===stageIdx, color=STAGE_COLORS[st.id];
                            return (
                              <div key={st.id} style={S.stageCol}>
                                {idx>0&&<div style={{...S.stageLine,background:idx<=stageIdx?"#cbd5e1":"#e2e8f0"}}/>}
                                <div style={{...S.stageDot,background:done?color:"#f1f5f9",borderColor:done?color:"#cbd5e1",
                                  boxShadow:active?`0 0 0 3px ${color}33`:"none",color:done?"#fff":"#94a3b8"}}>
                                  {done&&<Icons.Check/>}
                                </div>
                                <div style={{...S.stageLabel,color:done?color:"#94a3b8"}}>{st.label}</div>
                                {tx.stageHistory[st.id]&&(
                                  <div style={S.stageTime}><Icons.Clock/> {fmtTime(tx.stageHistory[st.id])}</div>
                                )}
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    );
                  })}
                </div>
              );
            })}

            {session.treatments.length===0&&(
              <div style={S.emptyState}>No treatments added to this session yet.</div>
            )}
          </div>
        )}

        {/* ══════ ROOMS TAB ══════ */}
        {activeTab==="rooms" && (
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(300px,1fr))",gap:18}}>
            {ROOMS.map(room=>{
              const rd=session.rooms[room], accent=ROOM_ACCENT[room], type=ROOM_TYPE(room), specs=PRESSURE_SPECS[type];
              return (
                <div key={room} style={{...S.roomCard,borderTopColor:accent}}>
                  <div style={{...S.roomCardTitle,color:accent}}>{room}</div>
                  <div style={{fontSize:10,color:"#94a3b8",marginBottom:14,letterSpacing:.5}}>
                    {type==="cyto"?"CYTO ISOLATOR":"CIVAS ISOLATOR"}
                  </div>

                  <div style={{overflowX:"auto"}}>
                    <table style={S.specTable}>
                      <thead>
                        <tr>
                          <th style={S.th}>Measurement</th>
                          <th style={S.th}>Range</th>
                          <th style={{...S.th,color:"#2563eb"}}>Session Start</th>
                          <th style={{...S.th,color:"#7c3aed"}}>Session End</th>
                        </tr>
                      </thead>
                      <tbody>
                        {specs.map(spec=>{
                          const sv=rd.readingsStart[spec.key], ev=rd.readingsEnd[spec.key];
                          const si=inRange(sv,spec), ei=inRange(ev,spec);
                          return (
                            <tr key={spec.key}>
                              <td style={S.td}>
                                <div style={{fontWeight:600,fontSize:12,color:"#374151"}}>{spec.label}</div>
                                <div style={{fontSize:10,color:"#94a3b8"}}>{spec.unit}</div>
                              </td>
                              <td style={{...S.td,fontSize:12,color:"#64748b",whiteSpace:"nowrap"}}>
                                {spec.min} – {spec.max}
                              </td>
                              <td style={S.td}>
                                <ReadingInput val={sv} spec={spec} inR={si} disabled={!!session.closedAt}
                                  onChange={v=>updateReading(room,"start",spec.key,v)}/>
                              </td>
                              <td style={S.td}>
                                <ReadingInput val={ev} spec={spec} inR={ei} disabled={!!session.closedAt}
                                  onChange={v=>updateReading(room,"end",spec.key,v)}/>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>

                  {/* Cleaning records with full timestamp display */}
                  <div style={{...S.fieldLabel,marginTop:16}}>🧹 Cleaning Before Session</div>
                  <CleaningRecord data={rd.cleaningBefore} disabled={!!session.closedAt} onMark={()=>markCleaning(room,"before")}/>
                  <div style={S.fieldLabel}>🧹 Cleaning After Session</div>
                  <CleaningRecord data={rd.cleaningAfter} disabled={!!session.closedAt} onMark={()=>markCleaning(room,"after")}/>
                </div>
              );
            })}
          </div>
        )}

        {/* ══════ LOG TAB ══════ */}
        {activeTab==="log" && (
          <div>
            <div style={{...S.card,display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:12}}>
              <div>
                <div style={S.sectionHeading}>Session {session.sessionNum} — {fmtDate(session.date+"T00:00:00.000Z")}</div>
                <div style={{display:"flex",gap:20,fontSize:12,color:"#64748b",marginTop:4}}>
                  <span>Started: {session.startedAt?fmtTime(session.startedAt):"Not started"}</span>
                  <span>Closed: {session.closedAt?fmtTime(session.closedAt):"Still active"}</span>
                  <span>Treatments: {session.treatments.length}</span>
                </div>
              </div>
              <button style={S.btnExcel} onClick={()=>exportToExcel(session)}>
                <Icons.Download/> Export to Excel
              </button>
            </div>

            {/* Treatments table */}
            <div style={S.tableWrap}>
              <div style={{...S.tableHead,gridTemplateColumns:"1fr 1.4fr 0.7fr 1fr 1fr 1fr 0.7fr 0.7fr 0.7fr 0.7fr 0.7fr 1.4fr"}}>
                <span>Patient</span><span>Drug / Batch</span><span>Room</span>
                <span>Operative</span><span>IPC Checker</span><span>Sprayed By</span>
                <span>Intro</span><span>Cabinet</span><span>Prepared</span><span>Checked</span><span>Dispatch</span>
                <span>Notes</span>
              </div>
              {session.treatments.length===0&&<div style={{...S.emptyState,padding:"16px 14px"}}>No treatments.</div>}
              {session.treatments.map((tx,ri)=>{
                const p=tx.personnel||emptyPersonnel();
                const sprayLabel=p.sprayedBy==="operative"?`Op (${p.operative||"?"})`:`IPC (${p.checker||"?"})`;
                return (
                  <div key={tx.id} style={{...S.tableRow,gridTemplateColumns:"1fr 1.4fr 0.7fr 1fr 1fr 1fr 0.7fr 0.7fr 0.7fr 0.7fr 0.7fr 1.4fr",background:ri%2===0?"#fff":"#f8fafc"}}>
                    <span style={{fontWeight:600,color:"#1e293b"}}>{tx.patientId}</span>
                    <span style={{fontSize:12}}>{tx.drug}<br/><span style={{color:"#94a3b8"}}>{tx.batch}</span></span>
                    <span style={{color:ROOM_ACCENT[tx.room],fontWeight:600,fontSize:12}}>{tx.room}</span>
                    <span style={{fontSize:12}}>{p.operative||"—"}</span>
                    <span style={{fontSize:12}}>{p.checker||"—"}</span>
                    <span style={{fontSize:11,color:p.sprayedBy==="checker"?"#7c3aed":"#2563eb",fontWeight:600}}>{sprayLabel}</span>
                    {TREATMENT_STAGES.map(st=>(
                      <span key={st.id} style={{fontSize:11,color:tx.stageHistory[st.id]?"#059669":"#cbd5e1"}}>
                        {tx.stageHistory[st.id]?fmtTime(tx.stageHistory[st.id]):"—"}
                      </span>
                    ))}
                    <span style={{fontSize:11,color:"#64748b",fontStyle:tx.notes?"normal":"italic"}}>{tx.notes||"—"}</span>
                  </div>
                );
              })}
            </div>

            {/* Room records */}
            <div style={{marginTop:28}}>
              <div style={S.sectionHeading}>Room Environmental Records</div>
              {ROOMS.map(room=>{
                const rd=session.rooms[room], accent=ROOM_ACCENT[room], type=ROOM_TYPE(room), specs=PRESSURE_SPECS[type];
                return (
                  <div key={room} style={{...S.card,borderLeft:`4px solid ${accent}`,marginBottom:12}}>
                    <div style={{fontWeight:700,fontSize:13,color:accent,marginBottom:10,letterSpacing:1}}>{room}</div>
                    <div style={{overflowX:"auto"}}>
                      <table style={{...S.specTable,fontSize:12}}>
                        <thead>
                          <tr>
                            <th style={S.th}>Measurement</th>
                            <th style={S.th}>Unit</th>
                            <th style={S.th}>Range</th>
                            <th style={{...S.th,color:"#2563eb"}}>Start</th>
                            <th style={S.th}>✓?</th>
                            <th style={{...S.th,color:"#7c3aed"}}>End</th>
                            <th style={S.th}>✓?</th>
                          </tr>
                        </thead>
                        <tbody>
                          {specs.map(spec=>{
                            const sv=rd.readingsStart[spec.key], ev=rd.readingsEnd[spec.key];
                            const si=inRange(sv,spec), ei=inRange(ev,spec);
                            return (
                              <tr key={spec.key}>
                                <td style={S.td}>{spec.label}</td>
                                <td style={{...S.td,color:"#94a3b8"}}>{spec.unit}</td>
                                <td style={{...S.td,color:"#64748b"}}>{spec.min} – {spec.max}</td>
                                <td style={{...S.td,fontWeight:600}}>{sv||"—"}</td>
                                <td style={{...S.td,color:si===null?"#94a3b8":si?"#059669":"#dc2626",fontWeight:700}}>
                                  {si===null?"—":si?"✓":"✗"}
                                </td>
                                <td style={{...S.td,fontWeight:600}}>{ev||"—"}</td>
                                <td style={{...S.td,color:ei===null?"#94a3b8":ei?"#059669":"#dc2626",fontWeight:700}}>
                                  {ei===null?"—":ei?"✓":"✗"}
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                    <div style={{display:"flex",gap:24,marginTop:10,fontSize:12,color:"#64748b",flexWrap:"wrap"}}>
                      <span>Cleaning before: <b style={{color:rd.cleaningBefore.done?"#059669":"#dc2626"}}>
                        {rd.cleaningBefore.done?`✓ ${rd.cleaningBefore.by} at ${fmtTime(rd.cleaningBefore.time)}`:"Not recorded"}
                      </b></span>
                      <span>Cleaning after: <b style={{color:rd.cleaningAfter.done?"#059669":"#dc2626"}}>
                        {rd.cleaningAfter.done?`✓ ${rd.cleaningAfter.by} at ${fmtTime(rd.cleaningAfter.time)}`:"Not recorded"}
                      </b></span>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

      </main>
    </div>
  );
}

// ─── SPRAY TOGGLE ─────────────────────────────────────────────────────────────
function SprayToggle({value, onChange}) {
  return (
    <div style={{display:"flex",gap:0,borderRadius:7,overflow:"hidden",border:"1px solid #e2e8f0",width:"fit-content",marginTop:2}}>
      {["operative","checker"].map(opt=>{
        const active=value===opt;
        const label=opt==="operative"?"Operative":"IPC Checker";
        return (
          <button key={opt}
            style={{padding:"5px 13px",fontSize:11,fontFamily:"'DM Mono',monospace",
              cursor:"pointer",border:"none",letterSpacing:.3,transition:"all .15s",
              background:active?(opt==="operative"?"#2563eb":"#7c3aed"):"#f8fafc",
              color:active?"#fff":"#64748b",fontWeight:active?700:400}}
            onClick={()=>onChange(opt)}>
            {active&&"✓ "}{label}
          </button>
        );
      })}
    </div>
  );
}

// ─── READING INPUT ────────────────────────────────────────────────────────────
function ReadingInput({val,spec,inR,disabled,onChange}) {
  const borderColor=inR===null?"#e2e8f0":inR?"#86efac":"#fca5a5";
  const bg=inR===null?"#f8fafc":inR?"#f0fdf4":"#fef2f2";
  return (
    <div style={{display:"flex",alignItems:"center",gap:4}}>
      <input style={{...S.input,width:72,background:bg,borderColor,padding:"5px 7px",fontSize:12}}
        value={val} disabled={disabled} onChange={e=>onChange(e.target.value)} placeholder="—"/>
      {inR!==null&&(
        <span style={{fontSize:10,fontWeight:700,color:inR?"#059669":"#dc2626"}}>{inR?"✓":"✗"}</span>
      )}
    </div>
  );
}

// ─── NOTE EDITOR ──────────────────────────────────────────────────────────────
function NoteEditor({initial,onSave,onCancel}) {
  const [val,setVal]=useState(initial||"");
  return (
    <div style={{margin:"8px 0 12px",background:"#f8fafc",border:"1px solid #e2e8f0",borderRadius:7,padding:10}}>
      <textarea style={{...S.input,width:"100%",resize:"vertical",minHeight:60,boxSizing:"border-box",background:"#fff"}}
        value={val} onChange={e=>setVal(e.target.value)} placeholder="Enter notes…" autoFocus/>
      <div style={{display:"flex",gap:7,marginTop:8}}>
        <button style={{...S.btnGreen,fontSize:12,padding:"4px 12px"}} onClick={()=>onSave(val)}>Save Note</button>
        <button style={{...S.btnGhost,fontSize:12,padding:"4px 10px"}} onClick={onCancel}>Cancel</button>
      </div>
    </div>
  );
}

// ─── CLEANING RECORD ──────────────────────────────────────────────────────────
function CleaningRecord({data,disabled,onMark}) {
  if(data.done) return (
    <div style={S.cleanDone}>
      <div>
        <div style={{color:"#059669",fontWeight:700,fontSize:13}}>✓ Complete</div>
        <div style={{color:"#64748b",fontSize:11,marginTop:2}}>
          By: <b>{data.by}</b> &nbsp;·&nbsp; <Icons.Clock/> <b>{fmtTime(data.time)}</b>
        </div>
      </div>
    </div>
  );
  return (
    <div style={S.cleanPending}>
      <span style={{color:"#d97706",fontSize:12}}>Not yet recorded</span>
      {!disabled&&(
        <button style={{...S.btnGreen,fontSize:12,padding:"4px 10px"}} onClick={onMark}>
          <Icons.Check/> Mark Done
        </button>
      )}
    </div>
  );
}

// ─── STYLES ───────────────────────────────────────────────────────────────────
const S = {
  root:      {background:"#f1f5f9",minHeight:"100vh",color:"#1e293b",fontFamily:"'DM Mono','Courier New',monospace"},
  header:    {background:"#fff",borderBottom:"1px solid #e2e8f0",padding:"13px 22px",
    display:"flex",alignItems:"center",justifyContent:"space-between",
    position:"sticky",top:0,zIndex:100,boxShadow:"0 1px 3px rgba(0,0,0,.06)"},
  logoMark:  {fontSize:26,color:"#2563eb",lineHeight:1},
  logoTitle: {fontWeight:700,fontSize:14,letterSpacing:4,color:"#1e293b"},
  logoSub:   {fontSize:10,color:"#94a3b8",letterSpacing:2},
  savedPill: {fontSize:12,color:"#64748b",display:"flex",alignItems:"center",gap:5,
    background:"#f8fafc",border:"1px solid #e2e8f0",borderRadius:20,padding:"3px 10px"},
  datePill:  {background:"#eff6ff",color:"#2563eb",border:"1px solid #bfdbfe",
    borderRadius:20,padding:"3px 12px",fontSize:12,fontWeight:600},
  sessionBar:{background:"#fff",borderBottom:"1px solid #e2e8f0",padding:"8px 22px",
    display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:8},
  sessionTab:{background:"#f8fafc",border:"1px solid #e2e8f0",borderRadius:7,padding:"5px 12px",
    color:"#64748b",cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"flex-start",gap:1},
  sessionTabActive:{background:"#eff6ff",borderColor:"#2563eb",color:"#1e293b"},
  closedPill:{fontSize:9,background:"#fef2f2",color:"#dc2626",border:"1px solid #fecaca",
    borderRadius:3,padding:"1px 5px",letterSpacing:1},
  newSessionBtn:{background:"#fff",border:"1px dashed #cbd5e1",borderRadius:7,padding:"5px 12px",
    color:"#94a3b8",cursor:"pointer",display:"flex",alignItems:"center",gap:5,fontSize:12},
  nav:        {background:"#fff",borderBottom:"1px solid #e2e8f0",display:"flex",padding:"0 22px"},
  navTab:     {background:"transparent",border:"none",borderBottom:"2px solid transparent",
    color:"#94a3b8",cursor:"pointer",fontFamily:"'DM Mono',monospace",
    fontSize:13,padding:"12px 16px",letterSpacing:.5,transition:"color .15s"},
  navTabActive:{color:"#2563eb",borderBottomColor:"#2563eb"},
  main:       {padding:"22px",maxWidth:1200,margin:"0 auto"},
  card:       {background:"#fff",border:"1px solid #e2e8f0",borderRadius:10,
    padding:"16px 18px",marginBottom:16,boxShadow:"0 1px 3px rgba(0,0,0,.04)"},
  sectionHeading:{fontWeight:700,fontSize:13,letterSpacing:.5,color:"#475569",
    marginBottom:10,textTransform:"uppercase"},
  formGrid:   {display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(175px,1fr))",gap:12},
  label:      {display:"flex",flexDirection:"column",gap:4,fontSize:11,color:"#64748b",letterSpacing:.5,textTransform:"uppercase"},
  input:      {background:"#f8fafc",border:"1px solid #e2e8f0",borderRadius:6,color:"#1e293b",
    fontFamily:"'DM Mono',monospace",fontSize:13,padding:"7px 10px",outline:"none",marginTop:2},

  personnelBox:{background:"#f0f7ff",border:"1px solid #bfdbfe",borderRadius:8,
    padding:"12px 14px",marginTop:14},
  personnelTitle:{fontSize:11,fontWeight:700,color:"#2563eb",letterSpacing:1,textTransform:"uppercase",
    marginBottom:10,display:"flex",alignItems:"center",gap:6},
  personnelStrip:{display:"flex",gap:14,flexWrap:"wrap",alignItems:"flex-start",
    background:"#f8faff",border:"1px solid #dbeafe",borderRadius:7,
    padding:"10px 12px",marginBottom:10},
  personnelField:{display:"flex",flexDirection:"column",gap:4,minWidth:120},
  personnelFieldLabel:{fontSize:10,color:"#64748b",letterSpacing:.5,textTransform:"uppercase",fontWeight:600},
  personnelValue:{fontSize:13,fontWeight:600,color:"#1e293b"},
  personnelInput:{background:"#fff",border:"1px solid #dbeafe",borderRadius:5,
    color:"#1e293b",fontFamily:"'DM Mono',monospace",fontSize:12,
    padding:"4px 8px",outline:"none",width:100},

  roomSection: {marginBottom:22},
  roomHeader:  {fontSize:11,fontWeight:700,letterSpacing:3,textTransform:"uppercase",
    marginBottom:8,paddingBottom:6,borderBottom:"2px solid",borderLeftWidth:3,
    borderLeftStyle:"solid",paddingLeft:8},
  txCard:      {background:"#fff",border:"1px solid #e2e8f0",borderRadius:9,borderLeft:"4px solid",
    padding:"14px 16px",marginBottom:10,boxShadow:"0 1px 3px rgba(0,0,0,.04)"},
  txTop:       {display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10},
  txPatient:   {fontWeight:700,fontSize:15,color:"#0f172a"},
  txDrug:      {fontSize:13,color:"#475569",marginTop:2},
  txBatch:     {fontSize:11,color:"#94a3b8",marginTop:2,letterSpacing:.5},
  noteBubble:  {display:"flex",alignItems:"flex-start",gap:6,fontSize:12,color:"#475569",
    background:"#fefce8",border:"1px solid #fef08a",borderRadius:6,padding:"7px 10px",
    marginBottom:10,lineHeight:1.5},
  stageTrack:  {display:"flex",alignItems:"flex-start",overflowX:"auto",paddingBottom:4,gap:0,marginTop:10},
  stageCol:    {display:"flex",flexDirection:"column",alignItems:"center",minWidth:90,position:"relative"},
  stageLine:   {position:"absolute",top:11,right:"50%",width:"100%",height:2,zIndex:0},
  stageDot:    {width:24,height:24,borderRadius:"50%",border:"2px solid",display:"flex",
    alignItems:"center",justifyContent:"center",zIndex:1,transition:"all .2s",position:"relative"},
  stageLabel:  {fontSize:9,textAlign:"center",marginTop:5,letterSpacing:.3,lineHeight:1.3,maxWidth:82},
  stageTime:   {fontSize:9,color:"#94a3b8",display:"flex",alignItems:"center",gap:2,marginTop:2},

  roomCard:      {background:"#fff",border:"1px solid #e2e8f0",borderTop:"3px solid",
    borderRadius:10,padding:"16px",boxShadow:"0 1px 3px rgba(0,0,0,.04)"},
  roomCardTitle: {fontWeight:700,fontSize:14,letterSpacing:2,textTransform:"uppercase",marginBottom:2},
  fieldLabel:    {fontSize:10,color:"#94a3b8",letterSpacing:1,textTransform:"uppercase",
    marginBottom:7,marginTop:14,display:"flex",alignItems:"center",gap:5},
  specTable:   {width:"100%",borderCollapse:"collapse",fontSize:12},
  th:          {textAlign:"left",padding:"6px 8px",fontSize:10,color:"#94a3b8",letterSpacing:.8,
    textTransform:"uppercase",borderBottom:"1px solid #e2e8f0",whiteSpace:"nowrap"},
  td:          {padding:"7px 8px",borderBottom:"1px solid #f1f5f9",verticalAlign:"middle"},
  cleanDone:   {background:"#f0fdf4",border:"1px solid #bbf7d0",borderRadius:7,
    padding:"10px 14px",display:"flex",alignItems:"center"},
  cleanPending:{background:"#fffbeb",border:"1px solid #fde68a",borderRadius:7,
    padding:"8px 12px",display:"flex",justifyContent:"space-between",alignItems:"center"},
  tableWrap:   {background:"#fff",border:"1px solid #e2e8f0",borderRadius:10,
    overflow:"hidden",boxShadow:"0 1px 3px rgba(0,0,0,.04)",marginBottom:8},
  tableHead:   {display:"grid",background:"#f8fafc",padding:"10px 14px",fontSize:10,
    color:"#64748b",letterSpacing:.8,textTransform:"uppercase",borderBottom:"1px solid #e2e8f0",gap:6},
  tableRow:    {display:"grid",padding:"10px 14px",fontSize:13,
    borderBottom:"1px solid #f1f5f9",alignItems:"center",gap:6},
  emptyState:  {color:"#cbd5e1",fontSize:13,padding:"22px",textAlign:"center"},
  btnBlue:     {background:"#2563eb",border:"none",borderRadius:6,color:"#fff",cursor:"pointer",
    fontFamily:"'DM Mono',monospace",fontSize:13,padding:"7px 14px",display:"flex",alignItems:"center",gap:6},
  btnGreen:    {background:"#059669",border:"none",borderRadius:6,color:"#fff",cursor:"pointer",
    fontFamily:"'DM Mono',monospace",fontSize:13,padding:"7px 14px",display:"flex",alignItems:"center",gap:6},
  btnRed:      {background:"#dc2626",border:"none",borderRadius:6,color:"#fff",cursor:"pointer",
    fontFamily:"'DM Mono',monospace",fontSize:13,padding:"7px 14px"},
  btnExcel:    {background:"#166534",border:"none",borderRadius:6,color:"#fff",cursor:"pointer",
    fontFamily:"'DM Mono',monospace",fontSize:13,padding:"8px 16px",
    display:"flex",alignItems:"center",gap:7,whiteSpace:"nowrap"},
  btnDanger:   {background:"#fef2f2",border:"1px solid #fecaca",borderRadius:6,color:"#dc2626",
    cursor:"pointer",padding:"5px 8px",display:"flex",alignItems:"center"},
  btnGhost:    {background:"#fff",border:"1px solid #e2e8f0",borderRadius:6,color:"#64748b",
    cursor:"pointer",fontFamily:"'DM Mono',monospace",fontSize:13,padding:"7px 14px"},
};
