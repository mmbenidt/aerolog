import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

const C = {
  bg:"#0d0f14", panel:"#13161e", border:"#1e2330",
  accent:"#3b82f6", accentDim:"#1d3e6f",
  ok:"#22c55e", warn:"#f59e0b", danger:"#ef4444",
  text:"#e2e8f0", muted:"#64748b",
};

const CSS = `
  @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;600&family=Barlow:wght@400;500;600&display=swap');
  *,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
  body{background:#0d0f14;color:#e2e8f0;font-family:'Barlow',sans-serif;min-height:100vh;-webkit-font-smoothing:antialiased}
  input,select,textarea{width:100%;background:#0d0f14;border:1px solid #1e2330;border-radius:6px;color:#e2e8f0;font-family:'JetBrains Mono',monospace;font-size:13px;padding:8px 10px;outline:none;transition:border-color .15s}
  input:focus,select:focus,textarea:focus{border-color:#3b82f6;box-shadow:0 0 0 2px #1d3e6f}
  select option{background:#13161e}
  textarea{resize:vertical;min-height:56px}
  ::-webkit-scrollbar{width:4px}::-webkit-scrollbar-track{background:#0d0f14}::-webkit-scrollbar-thumb{background:#1e2330;border-radius:2px}
  @keyframes slideIn{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:translateY(0)}}
  @keyframes pdot{0%,100%{box-shadow:0 0 0 0 rgba(59,130,246,.5)}50%{box-shadow:0 0 0 5px rgba(59,130,246,0)}}
  @keyframes blink{0%,100%{opacity:1}50%{opacity:.3}}
  @keyframes tIn{from{opacity:0;transform:translateX(16px)}to{opacity:1;transform:translateX(0)}}
  @keyframes tOut{from{opacity:1}to{opacity:0;transform:translateX(16px)}}
`;

const KEY_LOG="acft_v3_log", KEY_COMP="acft_v3_comp", KEY_SQ="acft_v4_squawks", KEY_DISP="acft_v3_dispatch", KEY_PROF="acft_v3_profile", KEY_REQ="acft_v5_required", KEY_AC="acft_v6_aircraft";

const REQUIRED_DEFS = [
  { key:"annual",     label:"Annual Inspection",      months:12, ref:"All aircraft" },
  { key:"transponder",label:"Transponder Check",      months:24, ref:"14 CFR §91.413" },
  { key:"altimeter",  label:"Altimeter / Altitude Reporting Check", months:24, ref:"14 CFR §91.411 — controlled airspace" },
  { key:"elt",        label:"ELT Inspection",         months:12, ref:"ELT battery: 1 hr cumulative use or 50% life" },
];

const SQUAWK_SEVERITIES = [
  { value:"minor",    label:"Minor",     ref:"Track / fix at next opportunity" },
  { value:"major",    label:"Major",     ref:"Should be resolved before next flight" },
  { value:"grounding",label:"Grounding", ref:"Aircraft not airworthy — AOG" },
];

const loadLS = k => { try { return JSON.parse(localStorage.getItem(k)||"null"); } catch { return null; } };
const saveLS = (k,v) => localStorage.setItem(k, JSON.stringify(v));
const fmt1 = n => (n!=null&&!isNaN(n)) ? Number(n).toFixed(1) : "—";
const todayStr = () => new Date().toISOString().split("T")[0];

// Always picks the entry with the most recent date field, regardless of insertion order
function latestForTail(entries, tail) {
  const te = entries.filter(e => e.tail === tail);
  if (!te.length) return null;
  return te.reduce((best, e) => (e.date > best.date ? e : best), te[0]);
}

function compStatus(item, entries) {
  if (item.type==="date") {
    const now=new Date(); now.setHours(0,0,0,0);
    const diff=Math.round((new Date(item.due)-now)/86400000);
    const warn=parseInt(item.warn)||30;
    if(diff<0) return {label:`${Math.abs(diff)}d overdue`,cls:"danger"};
    if(diff<=warn) return {label:`${diff}d remaining`,cls:"warn"};
    return {label:`${diff}d remaining`,cls:"ok"};
  }
  const latest = latestForTail(entries, item.tail);
  if(!latest) return {label:"No time data",cls:"warn"};
  const cur=item.type==="hobbs"?latest.hobbs:latest.tach;
  if(cur==null) return {label:"No time data",cls:"warn"};
  const rem=parseFloat(item.due)-cur, wH=parseFloat(item.warn)||10;
  if(rem<0) return {label:`${Math.abs(rem).toFixed(1)}h overdue`,cls:"danger"};
  if(rem<=wH) return {label:`${rem.toFixed(1)}h left`,cls:"warn"};
  return {label:`${rem.toFixed(1)}h left`,cls:"ok"};
}

function requiredStatus(lastDoneDate, months) {
  if (!lastDoneDate) return { label:"Not recorded", cls:"danger", missing:true };
  const last = new Date(lastDoneDate);
  const due = new Date(last);
  due.setMonth(due.getMonth() + months);
  const now = new Date(); now.setHours(0,0,0,0);
  const diff = Math.round((due - now) / 86400000);
  if (diff < 0) return { label:`${Math.abs(diff)}d overdue`, cls:"danger", due, missing:false };
  if (diff <= 30) return { label:`${diff}d remaining`, cls:"warn", due, missing:false };
  return { label:`${diff}d remaining`, cls:"ok", due, missing:false };
}

// Returns { allCurrent, items: [{key,label,status}] } for a tail's required inspections
function requiredInspectionsStatus(requiredItems, tail) {
  const rec = requiredItems?.[tail] || {};
  const items = REQUIRED_DEFS.map(def => ({
    ...def,
    status: requiredStatus(rec[def.key], def.months),
  }));
  const allCurrent = items.every(i => i.status.cls === "ok");
  return { allCurrent, items };
}

// 100-hour inspection: optional, based on Tach hours since last 100hr
function hundredHourStatus(entries, tail, lastTach) {
  const latest = latestForTail(entries, tail);
  if (!latest || latest.tach==null) return { label:"No tach data", cls:"warn", missing: lastTach==null };
  if (lastTach==null) return { label:"Not recorded", cls:"danger", missing:true };
  const used = latest.tach - lastTach;
  const remaining = 100 - used;
  if (remaining < 0) return { label:`${Math.abs(remaining).toFixed(1)}h overdue`, cls:"danger", missing:false };
  if (remaining <= 10) return { label:`${remaining.toFixed(1)}h left`, cls:"warn", missing:false };
  return { label:`${remaining.toFixed(1)}h left`, cls:"ok", missing:false };
}

// Average quarts of oil added per flight hour over the last N flights for a tail
function oilConsumptionRate(entries, tail, lookback=5) {
  const te=[...entries].filter(e=>e.tail===tail).sort((a,b)=>new Date(b.date)-new Date(a.date)||(b.id-a.id));
  let totalOil=0, totalHrs=0;
  for(let i=0;i<te.length-1 && i<lookback;i++){
    const cur=te[i], prev=te[i+1];
    if(cur.oil) totalOil+=cur.oil;
    if(cur.hobbs!=null&&prev.hobbs!=null) totalHrs+=(cur.hobbs-prev.hobbs);
  }
  if(totalHrs<=0) return null;
  return totalOil/totalHrs; // qt per hour
}

function StatBox({label,value,barColor,bg=C.bg}) {
  return (
    <div style={{background:bg,borderRadius:6,padding:"8px 10px",borderTop:barColor?`3px solid ${barColor}`:undefined}}>
      <div style={{fontSize:11,color:C.muted,marginBottom:3}}>{label}</div>
      <div style={{fontSize:14,fontFamily:"'JetBrains Mono',monospace"}}>{value}</div>
    </div>
  );
}

function Badge({cls,children}) {
  const m={ok:{bg:"#052e16",c:C.ok,b:"#166534"},warn:{bg:"#2d1f02",c:C.warn,b:"#854d0e"},danger:{bg:"#2d0a0a",c:C.danger,b:"#991b1b"},blue:{bg:"#0f2a50",c:C.accent,b:C.accentDim}};
  const s=m[cls]||m.ok;
  return <span style={{background:s.bg,color:s.c,border:`1px solid ${s.b}`,borderRadius:4,fontSize:11,fontWeight:600,padding:"2px 8px",fontFamily:"'JetBrains Mono',monospace",whiteSpace:"nowrap"}}>{children}</span>;
}

function Btn({onClick,variant="ghost",size="md",style:sx={},children,disabled}) {
  const p=size==="sm"?"4px 10px":"8px 16px", fs=size==="sm"?12:13;
  const v={primary:{background:C.accent,borderColor:C.accent,color:"#fff"},ghost:{background:"transparent",borderColor:C.border,color:C.text},danger:{background:"transparent",borderColor:C.danger,color:C.danger},success:{background:"#052e16",borderColor:"#166534",color:C.ok},warn:{background:"#2d1f02",borderColor:"#854d0e",color:C.warn}};
  return <button disabled={disabled} onClick={onClick} style={{border:"1px solid",borderRadius:6,cursor:disabled?"not-allowed":"pointer",fontFamily:"'Barlow',sans-serif",fontWeight:500,transition:"all .15s",opacity:disabled?.45:1,display:"inline-flex",alignItems:"center",gap:6,padding:p,fontSize:fs,...v[variant],...sx}}>{children}</button>;
}

function Card({children,style:sx={}}) {
  return <div style={{background:C.panel,border:`1px solid ${C.border}`,borderRadius:10,padding:"16px 18px",marginBottom:12,animation:"slideIn .2s ease",...sx}}>{children}</div>;
}

function Field({label,children}) {
  return <div style={{display:"flex",flexDirection:"column",gap:5}}><label style={{fontSize:11,color:C.muted,textTransform:"uppercase",letterSpacing:".08em",fontWeight:500}}>{label}</label>{children}</div>;
}

function Row({cols=2,gap=10,mb=12,children}) {
  return <div style={{display:"grid",gridTemplateColumns:`repeat(${cols},1fr)`,gap,marginBottom:mb}}>{children}</div>;
}

function Toast({msg,onDone}) {
  const [out,setOut]=useState(false);
  useEffect(()=>{const a=setTimeout(()=>setOut(true),2400),b=setTimeout(onDone,2700);return()=>{clearTimeout(a);clearTimeout(b)};},[]);
  return <div style={{background:C.panel,border:`1px solid ${C.border}`,borderRadius:8,padding:"10px 16px",fontFamily:"'JetBrains Mono',monospace",fontSize:13,color:C.text,boxShadow:"0 8px 24px rgba(0,0,0,.6)",animation:out?"tOut .3s ease forwards":"tIn .3s ease",display:"flex",alignItems:"center",gap:8}}><span style={{color:C.ok}}>✓</span>{msg}</div>;
}

function Modal({title,onClose,children}) {
  return (
    <div style={{position:"fixed",inset:0,zIndex:300,display:"flex",alignItems:"center",justifyContent:"center",background:"rgba(0,0,0,.8)",padding:16}} onClick={onClose}>
      <div style={{background:C.panel,border:`1px solid ${C.border}`,borderRadius:12,padding:20,width:"100%",maxWidth:520,maxHeight:"90vh",overflowY:"auto",animation:"slideIn .2s ease"}} onClick={e=>e.stopPropagation()}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
          <span style={{fontSize:15,fontWeight:500}}>{title}</span>
          <button onClick={onClose} style={{background:"none",border:"none",cursor:"pointer",color:C.muted,fontSize:20,lineHeight:1}}>✕</button>
        </div>
        {children}
      </div>
    </div>
  );
}

// ── Modals ──────────────────────────────────────────────────────────────────

function EditFlightModal({entry,onSave,onClose}) {
  const [f,setF]=useState({date:entry.date,tail:entry.tail,hobbs:entry.hobbs??"",tach:entry.tach??"",fuel:entry.fuel??"",gallonsAdded:entry.gallonsAdded??"",oil:entry.oil??"",notes:entry.notes??""});
  const set=(k,v)=>setF(p=>({...p,[k]:v}));
  const submit=()=>{
    if(!f.date||!f.tail||f.hobbs==="") return;
    onSave({...entry,date:f.date,tail:f.tail.trim().toUpperCase(),hobbs:parseFloat(f.hobbs),tach:f.tach!==""?parseFloat(f.tach):null,fuel:f.fuel!==""?parseFloat(f.fuel):null,gallonsAdded:f.gallonsAdded!==""?parseFloat(f.gallonsAdded):null,oil:f.oil!==""?parseInt(f.oil):null,notes:f.notes.trim()});
  };
  return (
    <Modal title="Edit flight entry" onClose={onClose}>
      <Row cols={2}><Field label="Date"><input type="date" value={f.date} onChange={e=>set("date",e.target.value)}/></Field><Field label="Tail #"><input value={f.tail} onChange={e=>set("tail",e.target.value)}/></Field></Row>
      <Row cols={3}><Field label="Hobbs end"><input type="number" step="0.1" value={f.hobbs} onChange={e=>set("hobbs",e.target.value)}/></Field><Field label="Tach end"><input type="number" step="0.1" value={f.tach} onChange={e=>set("tach",e.target.value)}/></Field><Field label="Fuel end (gal)"><input type="number" step="0.5" value={f.fuel} onChange={e=>set("fuel",e.target.value)}/></Field></Row>
      <Row cols={2}><Field label="Gallons added"><input type="number" step="0.5" value={f.gallonsAdded} onChange={e=>set("gallonsAdded",e.target.value)}/></Field><Field label="Oil added (qt)"><input type="number" step="1" min="0" value={f.oil} onChange={e=>set("oil",String(Math.round(parseFloat(e.target.value)||0)))}/></Field></Row>
      <Field label="Notes"><textarea value={f.notes} onChange={e=>set("notes",e.target.value)} style={{marginBottom:14}}/></Field>
      <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}><Btn onClick={onClose}>Cancel</Btn><Btn variant="primary" onClick={submit}>Save changes</Btn></div>
    </Modal>
  );
}

function EditCompModal({item,onSave,onClose}) {
  const [f,setF]=useState({name:item.name,tail:item.tail,type:item.type,due:item.due,warn:item.warn||"",notes:item.notes||""});
  const set=(k,v)=>setF(p=>({...p,[k]:v}));
  const DL={date:"Due date",hobbs:"Due at Hobbs",tach:"Due at Tach"};
  const WL={date:"Warn (days)",hobbs:"Warn (hrs)",tach:"Warn (hrs)"};
  const submit=()=>{if(!f.name.trim()||!f.tail.trim()||!f.due)return;onSave({...item,...f,tail:f.tail.trim().toUpperCase(),name:f.name.trim()});};
  return (
    <Modal title="Edit compliance item" onClose={onClose}>
      <Row cols={2}><Field label="Item name"><input value={f.name} onChange={e=>set("name",e.target.value)}/></Field><Field label="Tail #"><input value={f.tail} onChange={e=>set("tail",e.target.value)}/></Field></Row>
      <Row cols={3}><Field label="Type"><select value={f.type} onChange={e=>set("type",e.target.value)}><option value="date">By date</option><option value="hobbs">Hobbs hrs</option><option value="tach">Tach hrs</option></select></Field><Field label={DL[f.type]}><input type={f.type==="date"?"date":"number"} step="0.1" value={f.due} onChange={e=>set("due",e.target.value)}/></Field><Field label={WL[f.type]}><input type="number" value={f.warn} onChange={e=>set("warn",e.target.value)}/></Field></Row>
      <Field label="Notes"><textarea value={f.notes} onChange={e=>set("notes",e.target.value)} style={{marginBottom:14}}/></Field>
      <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}><Btn onClick={onClose}>Cancel</Btn><Btn variant="primary" onClick={submit}>Save changes</Btn></div>
    </Modal>
  );
}

function ResolveModal({sq,onResolve,onClose}) {
  const [date,setDate]=useState(todayStr());
  const [action,setAction]=useState("");
  const submit=()=>{if(!action.trim())return;onResolve(sq.id,{resolvedDate:date,action:action.trim()});};
  return (
    <Modal title={`Resolve: ${sq.title}`} onClose={onClose}>
      <div style={{background:C.bg,borderRadius:8,padding:"10px 14px",marginBottom:14,fontSize:13,color:C.muted}}><div style={{fontWeight:500,color:C.text,marginBottom:4}}>{sq.tail} — reported {sq.reportedDate}</div>{sq.description&&<div style={{fontStyle:"italic"}}>{sq.description}</div>}</div>
      <Row cols={2} mb={14}><Field label="Date resolved"><input type="date" value={date} onChange={e=>setDate(e.target.value)}/></Field></Row>
      <Field label="What was done"><textarea placeholder="Corrective action, inspection finding, or deferral…" value={action} onChange={e=>setAction(e.target.value)} style={{marginBottom:14}}/></Field>
      <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}><Btn onClick={onClose}>Cancel</Btn><Btn variant="success" onClick={submit}>✓ Mark resolved</Btn></div>
    </Modal>
  );
}

function SquawkDetailModal({sq,onClose}) {
  return (
    <Modal title="Squawk record" onClose={onClose}>
      <div style={{display:"flex",gap:8,alignItems:"center",marginBottom:12}}><Badge cls={sq.resolved?"ok":"danger"}>{sq.resolved?"RESOLVED":"OPEN"}</Badge><span style={{fontSize:12,color:C.muted}}>{sq.tail}</span></div>
      <div style={{marginBottom:12}}><div style={{fontSize:11,color:C.muted,marginBottom:3,textTransform:"uppercase",letterSpacing:".08em"}}>Issue</div><div style={{fontSize:14,fontWeight:500,marginBottom:4}}>{sq.title}</div>{sq.description&&<div style={{fontSize:13,color:C.muted,fontStyle:"italic"}}>{sq.description}</div>}</div>
      <div style={{fontSize:11,color:C.muted,marginBottom:12}}>Reported: {sq.reportedDate}</div>
      {sq.resolved&&<div style={{background:C.bg,border:"1px solid #166534",borderRadius:8,padding:"12px 14px"}}><div style={{fontSize:11,color:C.ok,textTransform:"uppercase",letterSpacing:".08em",marginBottom:6}}>Resolution</div><div style={{fontSize:13,marginBottom:4}}>{sq.resolution.action}</div><div style={{fontSize:12,color:C.muted}}>Resolved: {sq.resolution.resolvedDate}</div></div>}
      <div style={{marginTop:14,display:"flex",justifyContent:"flex-end"}}><Btn onClick={onClose}>Close</Btn></div>
    </Modal>
  );
}

// ── Profile tab ─────────────────────────────────────────────────────────────

function ProfileTab({profile,setProfile,toast}) {
  const [editing,setEditing]=useState(!profile?.name);
  const [f,setF]=useState({name:profile?.name||"",homeAirport:profile?.homeAirport||"",role:profile?.role||"",photo:profile?.photo||null});
  const fileRef=useRef();
  const set=(k,v)=>setF(p=>({...p,[k]:v}));

  const handlePhoto=e=>{const file=e.target.files[0];if(!file)return;const r=new FileReader();r.onload=ev=>set("photo",ev.target.result);r.readAsDataURL(file);};

  const save=()=>{
    if(!f.name.trim()){toast("Name is required.");return;}
    const p={...f,name:f.name.trim(),homeAirport:f.homeAirport.trim().toUpperCase()};
    setProfile(p);saveLS(KEY_PROF,p);setEditing(false);toast("Profile saved.");
  };

  if(!editing&&profile?.name) return (
    <div>
      <Card>
        <div style={{display:"flex",gap:18,alignItems:"center",marginBottom:18}}>
          <div style={{width:72,height:72,borderRadius:"50%",overflow:"hidden",border:`2px solid ${C.border}`,flexShrink:0,background:C.bg,display:"flex",alignItems:"center",justifyContent:"center"}}>
            {profile.photo?<img src={profile.photo} alt="" style={{width:"100%",height:"100%",objectFit:"cover"}}/>:<span style={{fontSize:28,color:C.muted}}>✈</span>}
          </div>
          <div>
            <div style={{fontSize:20,fontWeight:600,marginBottom:4}}>{profile.name}</div>
            {profile.role&&<div style={{fontSize:13,color:C.muted,marginBottom:6}}>{profile.role}</div>}
            {profile.homeAirport&&<div style={{display:"inline-flex",alignItems:"center",gap:6,background:C.bg,border:`1px solid ${C.border}`,borderRadius:6,padding:"4px 10px"}}><span style={{fontSize:11,color:C.muted}}>HOME</span><span style={{fontFamily:"'JetBrains Mono',monospace",fontSize:14,fontWeight:600,color:C.accent}}>{profile.homeAirport}</span></div>}
          </div>
        </div>
        <Btn onClick={()=>setEditing(true)}>✎ Edit profile</Btn>
      </Card>
      <Card><div style={{fontSize:11,color:C.muted,textTransform:"uppercase",letterSpacing:".08em",marginBottom:8}}>App info</div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>{[["Version","v4.0"],["Storage","Local browser"],["Data","Persists between sessions"]].map(([k,v])=><div key={k} style={{background:C.bg,borderRadius:6,padding:"10px 12px"}}><div style={{fontSize:11,color:C.muted,marginBottom:3}}>{k}</div><div style={{fontSize:13,fontFamily:"'JetBrains Mono',monospace"}}>{v}</div></div>)}</div></Card>
    </div>
  );

  return (
    <Card>
      <div style={{fontSize:14,fontWeight:500,marginBottom:16}}>{profile?.name?"Edit profile":"Set up your profile"}</div>
      <div style={{display:"flex",alignItems:"center",gap:14,marginBottom:18}}>
        <div onClick={()=>fileRef.current.click()} style={{width:72,height:72,borderRadius:"50%",overflow:"hidden",border:`2px dashed ${C.border}`,background:C.bg,display:"flex",alignItems:"center",justifyContent:"center",cursor:"pointer",flexShrink:0}} onMouseEnter={e=>e.currentTarget.style.borderColor=C.accent} onMouseLeave={e=>e.currentTarget.style.borderColor=C.border}>
          {f.photo?<img src={f.photo} alt="" style={{width:"100%",height:"100%",objectFit:"cover"}}/>:<span style={{fontSize:24,color:C.muted}}>+</span>}
        </div>
        <div><div style={{fontSize:13,marginBottom:4}}>Profile photo</div><div style={{fontSize:12,color:C.muted}}>Click to upload</div></div>
        <input ref={fileRef} type="file" accept="image/*" style={{display:"none"}} onChange={handlePhoto}/>
      </div>
      <Row cols={2}><Field label="Full name"><input placeholder="Jane Smith" value={f.name} onChange={e=>set("name",e.target.value)}/></Field><Field label="Role / cert"><input placeholder="PPL, CFI, A&P…" value={f.role} onChange={e=>set("role",e.target.value)}/></Field></Row>
      <Row cols={2} mb={18}><Field label="Home airport (ICAO)"><input placeholder="KBFI" value={f.homeAirport} onChange={e=>set("homeAirport",e.target.value.toUpperCase())}/></Field></Row>
      <div style={{display:"flex",gap:8}}>{profile?.name&&<Btn onClick={()=>setEditing(false)}>Cancel</Btn>}<Btn variant="primary" onClick={save}>Save profile</Btn></div>
    </Card>
  );
}

// ── Aircraft registry tab ───────────────────────────────────────────────────

function AircraftTab({entries,aircraft,setAircraft,toast}) {
  const tails=[...new Set(entries.map(e=>e.tail))];
  const [editingTail,setEditingTail]=useState(null);
  const [f,setF]=useState({make:"",model:"",serial:"",year:"",engine:"",hundredHourEnabled:false,lastHundredHourTach:""});

  const startEdit=tail=>{
    const rec=aircraft?.[tail]||{};
    setF({make:rec.make||"",model:rec.model||"",serial:rec.serial||"",year:rec.year||"",engine:rec.engine||"",hundredHourEnabled:!!rec.hundredHourEnabled,lastHundredHourTach:rec.lastHundredHourTach??""});
    setEditingTail(tail);
  };

  const set=(k,v)=>setF(p=>({...p,[k]:v}));

  const save=()=>{
    const rec={...f, lastHundredHourTach: f.lastHundredHourTach!==""?parseFloat(f.lastHundredHourTach):null};
    const next={...aircraft, [editingTail]: rec};
    setAircraft(next); saveLS(KEY_AC,next);
    setEditingTail(null);
    toast("Aircraft details saved.");
  };

  if(!tails.length) return <div style={{textAlign:"center",padding:"2rem",color:C.muted,fontSize:14}}>Log a flight to register an aircraft.</div>;

  return (
    <div>
      {tails.map(tail=>{
        const rec=aircraft?.[tail]||{};
        const latest=latestForTail(entries,tail);
        const isEditing=editingTail===tail;
        const hh=rec.hundredHourEnabled?hundredHourStatus(entries,tail,rec.lastHundredHourTach):null;
        return (
          <Card key={tail}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:isEditing?14:8}}>
              <div>
                <div style={{fontSize:11,color:C.muted,textTransform:"uppercase",letterSpacing:".1em",marginBottom:2}}>Aircraft</div>
                <div style={{fontSize:20,fontWeight:600,fontFamily:"'JetBrains Mono',monospace",color:C.accent}}>{tail}</div>
              </div>
              {!isEditing&&<Btn size="sm" onClick={()=>startEdit(tail)}>✎ Edit details</Btn>}
            </div>

            {!isEditing&&(rec.make||rec.model||rec.serial||rec.year||rec.engine)&&(
              <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:8,marginBottom:rec.hundredHourEnabled?12:0}}>
                {rec.make&&<StatBox label="Make" value={rec.make}/>}
                {rec.model&&<StatBox label="Model" value={rec.model}/>}
                {rec.year&&<StatBox label="Year" value={rec.year}/>}
                {rec.serial&&<StatBox label="Serial #" value={rec.serial}/>}
                {rec.engine&&<StatBox label="Engine" value={rec.engine}/>}
              </div>
            )}
            {!isEditing&&!(rec.make||rec.model||rec.serial||rec.year||rec.engine)&&(
              <div style={{fontSize:13,color:C.muted,marginBottom:8}}>No aircraft details on file yet.</div>
            )}

            {!isEditing&&rec.hundredHourEnabled&&hh&&(
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginTop:8,paddingTop:12,borderTop:`1px solid ${C.border}`}}>
                <span style={{fontSize:13}}>100-hour inspection</span>
                <Badge cls={hh.cls}>{hh.missing?"Not recorded":hh.label}</Badge>
              </div>
            )}

            {isEditing&&(
              <div>
                <Row cols={2}>
                  <Field label="Make"><input placeholder="Cessna" value={f.make} onChange={e=>set("make",e.target.value)}/></Field>
                  <Field label="Model"><input placeholder="172N" value={f.model} onChange={e=>set("model",e.target.value)}/></Field>
                </Row>
                <Row cols={2}>
                  <Field label="Year"><input placeholder="1978" value={f.year} onChange={e=>set("year",e.target.value)}/></Field>
                  <Field label="Serial #"><input placeholder="17269201" value={f.serial} onChange={e=>set("serial",e.target.value)}/></Field>
                </Row>
                <Row cols={1}>
                  <Field label="Engine"><input placeholder="Lycoming O-320-H2AD" value={f.engine} onChange={e=>set("engine",e.target.value)}/></Field>
                </Row>
                <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:12,marginTop:4}}>
                  <input type="checkbox" id={`hh-${tail}`} checked={f.hundredHourEnabled} onChange={e=>set("hundredHourEnabled",e.target.checked)} style={{width:"auto"}}/>
                  <label htmlFor={`hh-${tail}`} style={{fontSize:13,cursor:"pointer"}}>Track 100-hour inspections (used for hire / flight instruction)</label>
                </div>
                {f.hundredHourEnabled&&(
                  <Row cols={2}>
                    <Field label="Tach at last 100-hr inspection"><input type="number" step="0.1" placeholder="1085.4" value={f.lastHundredHourTach} onChange={e=>set("lastHundredHourTach",e.target.value)}/></Field>
                  </Row>
                )}
                <div style={{display:"flex",gap:8}}>
                  <Btn onClick={()=>setEditingTail(null)}>Cancel</Btn>
                  <Btn variant="primary" onClick={save}>Save</Btn>
                </div>
              </div>
            )}
          </Card>
        );
      })}
    </div>
  );
}

// ── Dispatch tab ─────────────────────────────────────────────────────────────

function DispatchTab({entries,dispatch,setDispatch,compItems,squawks,requiredItems,aircraft,toast}) {
  const tails=[...new Set(entries.map(e=>e.tail))];
  const [sel,setSel]=useState(dispatch?.tail||(tails[0]??""));
  const [elapsed,setElapsed]=useState(0);
  const alerted=useRef(new Set());

  const beep=()=>{try{const ctx=new(window.AudioContext||window.webkitAudioContext)(),o=ctx.createOscillator(),g=ctx.createGain();o.connect(g);g.connect(ctx.destination);o.type="sine";o.frequency.value=880;g.gain.setValueAtTime(.4,ctx.currentTime);g.gain.exponentialRampToValueAtTime(.001,ctx.currentTime+.6);o.start();o.stop(ctx.currentTime+.6);}catch{}};

  useEffect(()=>{if(dispatch?.alertedHours)alerted.current=new Set(dispatch.alertedHours);},[]);

  useEffect(()=>{
    if(!dispatch)return;
    const iv=setInterval(()=>{
      const s=Math.floor((Date.now()-dispatch.startedAt)/1000);
      setElapsed(s);
      const h=Math.floor(s/3600);
      if(h>=3)for(let i=3;i<=h;i++){if(!alerted.current.has(i)){alerted.current.add(i);beep();toast(`⏰ ${i}h since dispatch — log totals for ${dispatch.tail}`);const u={...dispatch,alertedHours:[...alerted.current]};setDispatch(u);saveLS(KEY_DISP,u);}}
    },1000);
    return()=>clearInterval(iv);
  },[dispatch]);

  const fmtE=s=>`${String(Math.floor(s/3600)).padStart(2,"0")}:${String(Math.floor((s%3600)/60)).padStart(2,"0")}:${String(s%60).padStart(2,"0")}`;

  function isAllCurrent(tail) {
    const overdueComp = compItems.filter(c=>c.tail===tail && compStatus(c,entries).cls==="danger").length;
    const warnComp    = compItems.filter(c=>c.tail===tail && compStatus(c,entries).cls==="warn").length;
    const openSq      = squawks.filter(s=>s.tail===tail && !s.resolved).length;
    const req         = requiredInspectionsStatus(requiredItems, tail);
    const rec         = aircraft?.[tail];
    const hh          = rec?.hundredHourEnabled ? hundredHourStatus(entries,tail,rec.lastHundredHourTach) : null;
    const hhOk        = !hh || hh.cls==="ok";
    return overdueComp===0 && warnComp===0 && openSq===0 && req.allCurrent && hhOk;
  }

  function groundingSquawks(tail) {
    return squawks.filter(s=>s.tail===tail && !s.resolved && s.severity==="grounding");
  }

  if(dispatch){
    const hrs=elapsed/3600;
    const next=hrs<3?3:Math.ceil(hrs)===Math.floor(hrs)?Math.floor(hrs)+1:Math.ceil(hrs);
    const pf=latestForTail(entries, dispatch.tail);
    const current=isAllCurrent(dispatch.tail);
    return (
      <Card style={{border:`1px solid ${C.accentDim}`}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
          <div><div style={{fontSize:11,color:C.muted,textTransform:"uppercase",letterSpacing:".1em",marginBottom:4}}>Active dispatch</div><div style={{fontSize:22,fontWeight:600,fontFamily:"'JetBrains Mono',monospace",color:C.accent}}>{dispatch.tail}</div></div>
          <div style={{textAlign:"right"}}><div style={{fontSize:11,color:C.muted,marginBottom:2}}>Airborne</div><div style={{fontSize:28,fontFamily:"'JetBrains Mono',monospace",color:elapsed>=10800?C.warn:C.text,animation:elapsed>=10800?"blink 1.5s infinite":"none"}}>{fmtE(elapsed)}</div></div>
        </div>
        <div style={{fontSize:12,color:elapsed<10800?C.muted:C.warn,marginBottom:14}}>{elapsed<10800?`First alert in ${fmtE(10800-elapsed)}`:`Next alert in ${fmtE(Math.max(0,next*3600-elapsed))}`}</div>
        {pf&&<div style={{borderTop:`1px solid ${C.border}`,paddingTop:12,marginBottom:14}}>
          <div style={{fontSize:11,color:C.muted,marginBottom:8,textTransform:"uppercase",letterSpacing:".08em"}}>Prior ending totals</div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8}}>
            <StatBox label="Hobbs" value={fmt1(pf.hobbs)}/>
            <StatBox label="Tach"  value={fmt1(pf.tach)}/>
            <StatBox label="Fuel"  value={pf.fuel!=null?`${fmt1(pf.fuel)} gal`:"—"}/>
            <StatBox label="Status" value={current?"✓ Current":"⚠ Check"}/>
          </div>
        </div>}
        <Btn variant="danger" onClick={()=>{setDispatch(null);saveLS(KEY_DISP,null);alerted.current=new Set();toast("Flight closed — log ending totals.");}}>✕ Close flight</Btn>
      </Card>
    );
  }

  const pf=latestForTail(entries, sel);
  const current=sel?isAllCurrent(sel):false;
  const grounding=sel?groundingSquawks(sel):[];
  return (
    <Card>
      <div style={{fontSize:14,fontWeight:500,marginBottom:12}}>Dispatch a flight</div>
      <Row cols={2} mb={14}><Field label="Aircraft"><select value={sel} onChange={e=>setSel(e.target.value)}><option value="">Select…</option>{tails.map(t=><option key={t} value={t}>{t}</option>)}</select></Field></Row>
      {sel&&grounding.length>0&&(
        <div style={{background:"#2d0a0a",border:"1px solid #991b1b",borderRadius:8,padding:"10px 14px",marginBottom:14}}>
          <div style={{color:C.danger,fontSize:13,fontWeight:600,marginBottom:6}}>⛔ Aircraft grounded — AOG</div>
          {grounding.map(s=><div key={s.id} style={{fontSize:12,color:C.text,marginBottom:2}}>• {s.title}</div>)}
        </div>
      )}
      {sel&&pf&&<div style={{background:C.bg,border:`1px solid ${C.border}`,borderRadius:8,padding:"12px 14px",marginBottom:14}}>
        <div style={{fontSize:11,color:C.muted,textTransform:"uppercase",letterSpacing:".08em",marginBottom:10}}>Prior ending totals — {pf.date}</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8}}>
          <StatBox label="Hobbs" value={fmt1(pf.hobbs)} bg={C.panel}/>
          <StatBox label="Tach"  value={fmt1(pf.tach)}  bg={C.panel}/>
          <StatBox label="Fuel"  value={pf.fuel!=null?`${fmt1(pf.fuel)} gal`:"—"} bg={C.panel}/>
          <StatBox label="Status" value={current?"✓ Current":"⚠ Check"} bg={C.panel}/>
        </div>
        {!current&&<div style={{fontSize:12,color:C.warn,marginTop:10}}>⚠ This aircraft has open items — check the Compliance tab before flight.</div>}
      </div>}
      {sel&&!pf&&<div style={{fontSize:13,color:C.muted,marginBottom:12}}>No prior flight on record.</div>}
      <Btn variant={grounding.length>0?"danger":"primary"} disabled={!sel||grounding.length>0} onClick={()=>{const d={tail:sel,startedAt:Date.now(),alertedHours:[]};setDispatch(d);saveLS(KEY_DISP,d);toast(`Dispatched ${sel}`);}}>{grounding.length>0?"⛔ Grounded — cannot dispatch":`▶ Dispatch ${sel||"aircraft"}`}</Btn>
    </Card>
  );
}

// ── Times tab ────────────────────────────────────────────────────────────────

// Finds the entry immediately preceding (by date, then id) a given entry for the same tail
function priorEntryFor(entries, entry) {
  const te = entries.filter(e => e.tail === entry.tail && e.id !== entry.id);
  const earlier = te.filter(e => e.date < entry.date || (e.date === entry.date && e.id < entry.id));
  if (!earlier.length) return null;
  return earlier.reduce((best, e) => {
    if (e.date > best.date) return e;
    if (e.date === best.date && e.id > best.id) return e;
    return best;
  }, earlier[0]);
}

function exportToExcel(entries, profile) {
  const sorted=[...entries].sort((a,b)=>new Date(a.date)-new Date(b.date)||(a.id-b.id));
  const rows = sorted.map(e=>{
    const prior = priorEntryFor(entries, e);
    const hobbsTime = prior&&prior.hobbs!=null&&e.hobbs!=null ? +(e.hobbs - prior.hobbs).toFixed(1) : "";
    const tachTime  = prior&&prior.tach!=null&&e.tach!=null  ? +(e.tach  - prior.tach ).toFixed(1) : "";
    return {
      Date: e.date,
      Tail: e.tail,
      "Hobbs Start": prior?.hobbs ?? "",
      "Hobbs End": e.hobbs,
      "Hobbs Time": hobbsTime,
      "Tach Start": prior?.tach ?? "",
      "Tach End": e.tach ?? "",
      "Tach Time": tachTime,
      "Fuel End (gal)": e.fuel ?? "",
      "Gallons Added": e.gallonsAdded ?? "",
      "Oil Added (qt)": e.oil ?? "",
      "Logged By": e.loggedBy || "",
      Notes: e.notes || "",
    };
  });
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Flight Log");
  const fname = `flight-log-${todayStr()}.xlsx`;
  XLSX.writeFile(wb, fname);
}

function TimesTab({entries,setEntries,profile,toast}) {
  const [form,setForm]=useState({date:todayStr(),tail:"",hobbs:"",tach:"",fuel:"",gallonsAdded:"",oil:"",notes:""});
  const [filter,setFilter]=useState("");
  const [editing,setEditing]=useState(null);
  const tails=[...new Set(entries.map(e=>e.tail))];
  const set=(k,v)=>setForm(f=>({...f,[k]:v}));

  const save=()=>{
    if(!form.date||!form.tail.trim()||form.hobbs===""){toast("Date, tail #, and Hobbs end required.");return;}
    const e={id:Date.now(),date:form.date,tail:form.tail.trim().toUpperCase(),hobbs:parseFloat(form.hobbs),tach:form.tach!==""?parseFloat(form.tach):null,fuel:form.fuel!==""?parseFloat(form.fuel):null,gallonsAdded:form.gallonsAdded!==""?parseFloat(form.gallonsAdded):null,oil:form.oil!==""?parseInt(form.oil):null,notes:form.notes.trim(),loggedBy:profile?.name||""};
    const next=[e,...entries];setEntries(next);saveLS(KEY_LOG,next);setForm(f=>({...f,hobbs:"",tach:"",fuel:"",gallonsAdded:"",oil:"",notes:""}));toast("Entry saved.");
  };

  const del=id=>{if(!window.confirm("Delete this entry?"))return;const n=entries.filter(e=>e.id!==id);setEntries(n);saveLS(KEY_LOG,n);};

  const applyEdit=u=>{const n=entries.map(e=>e.id===u.id?u:e);n.sort((a,b)=>new Date(b.date)-new Date(a.date)||(b.id-a.id));setEntries(n);saveLS(KEY_LOG,n);setEditing(null);toast("Entry updated.");};

  const sorted=[...entries].sort((a,b)=>new Date(b.date)-new Date(a.date)||(b.id-a.id));
  const rows=filter?sorted.filter(e=>e.tail===filter):sorted;
  return (
    <div>
      {editing&&<EditFlightModal entry={editing} onSave={applyEdit} onClose={()=>setEditing(null)}/>}
      <Card>
        <div style={{fontSize:14,fontWeight:500,marginBottom:14}}>Log flight totals</div>
        <Row cols={2}><Field label="Date"><input type="date" value={form.date} onChange={e=>set("date",e.target.value)}/></Field><Field label="Aircraft tail #"><input placeholder="N12345" value={form.tail} onChange={e=>set("tail",e.target.value)}/></Field></Row>
        <Row cols={3}><Field label="Hobbs end"><input type="number" step="0.1" placeholder="1234.5" value={form.hobbs} onChange={e=>set("hobbs",e.target.value)}/></Field><Field label="Tach end"><input type="number" step="0.1" placeholder="1100.2" value={form.tach} onChange={e=>set("tach",e.target.value)}/></Field><Field label="Fuel end (gal)"><input type="number" step="0.5" placeholder="35.0" value={form.fuel} onChange={e=>set("fuel",e.target.value)}/></Field></Row>
        <Row cols={2}><Field label="Gallons added"><input type="number" step="0.5" placeholder="0.0" value={form.gallonsAdded} onChange={e=>set("gallonsAdded",e.target.value)}/></Field><Field label="Oil added (qt)"><input type="number" step="1" min="0" placeholder="0" value={form.oil} onChange={e=>set("oil",String(Math.round(parseFloat(e.target.value)||0)))}/></Field></Row>
        <Field label="Notes"><textarea placeholder="Destination, squawks, remarks…" value={form.notes} onChange={e=>set("notes",e.target.value)} style={{marginBottom:12}}/></Field>
        <Btn variant="primary" onClick={save}>+ Save entry</Btn>
      </Card>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10,gap:10,flexWrap:"wrap"}}>
        <span style={{fontSize:15,fontWeight:500}}>Flight log</span>
        <div style={{display:"flex",gap:8,alignItems:"center"}}>
          <select value={filter} onChange={e=>setFilter(e.target.value)} style={{width:"auto",fontSize:12}}><option value="">All aircraft</option>{tails.map(t=><option key={t} value={t}>{t}</option>)}</select>
          <Btn size="sm" onClick={()=>exportToExcel(entries, profile)} disabled={!entries.length}>⬇ Export to Excel</Btn>
        </div>
      </div>
      {!rows.length&&<div style={{textAlign:"center",padding:"2rem",color:C.muted,fontSize:14}}>No entries yet.</div>}
      {rows.map(e=>{
        const prior=priorEntryFor(entries,e);
        const hobbsTime=prior&&prior.hobbs!=null&&e.hobbs!=null?(e.hobbs-prior.hobbs):null;
        const tachTime=prior&&prior.tach!=null&&e.tach!=null?(e.tach-prior.tach):null;
        return (
        <div key={e.id} style={{background:C.panel,border:`1px solid ${C.border}`,borderRadius:8,padding:"10px 14px",marginBottom:8,animation:"slideIn .2s ease"}}>
          <div style={{display:"flex",justifyContent:"space-between",marginBottom:8}}>
            <div style={{display:"flex",gap:10,alignItems:"center"}}><span style={{fontFamily:"'JetBrains Mono',monospace",fontWeight:600,color:C.accent}}>{e.tail}</span><span style={{fontSize:12,color:C.muted}}>{e.date}</span></div>
            <div style={{display:"flex",gap:6}}><Btn variant="ghost" size="sm" onClick={()=>setEditing(e)}>✎ Edit</Btn><Btn variant="danger" size="sm" onClick={()=>del(e.id)}>✕</Btn></div>
          </div>
          <div style={{display:"flex",gap:18,flexWrap:"wrap"}}>
            {[["Hobbs",fmt1(e.hobbs)],["Hobbs time",hobbsTime!=null?fmt1(hobbsTime):"—"],["Tach",fmt1(e.tach)],["Tach time",tachTime!=null?fmt1(tachTime):"—"],["Fuel",e.fuel!=null?`${fmt1(e.fuel)} gal`:"—"],["Added",e.gallonsAdded!=null?`+${fmt1(e.gallonsAdded)} gal`:"—"],["Oil",e.oil!=null?`+${e.oil} qt`:"—"]].map(([k,v])=><div key={k}><div style={{fontSize:11,color:C.muted}}>{k}</div><div style={{fontSize:13,fontFamily:"'JetBrains Mono',monospace"}}>{v}</div></div>)}
          </div>
          {(e.notes||e.loggedBy)&&<div style={{marginTop:6,display:"flex",justifyContent:"space-between",alignItems:"flex-end",gap:10}}>
            {e.notes&&<div style={{fontSize:12,color:C.muted,fontStyle:"italic"}}>{e.notes}</div>}
            {e.loggedBy&&<div style={{fontSize:11,color:C.muted,marginLeft:"auto",whiteSpace:"nowrap"}}>Logged by {e.loggedBy}</div>}
          </div>}
        </div>
        );
      })}
    </div>
  );
}

// ── Compliance tab ────────────────────────────────────────────────────────────

function ComplianceTab({entries,compItems,setCompItems,squawks,setSquawks,requiredItems,setRequiredItems,profile,toast}) {
  const [sub,setSub]=useState("compliance");
  const [cForm,setCForm]=useState({name:"",tail:"",type:"date",due:"",warn:"",notes:""});
  const [sForm,setSForm]=useState({title:"",tail:"",reportedDate:todayStr(),description:"",severity:"minor"});
  const [editC,setEditC]=useState(null);
  const [resolving,setResolving]=useState(null);
  const [viewing,setViewing]=useState(null);
  const [reqTail,setReqTail]=useState("");
  const sc=(k,v)=>setCForm(f=>({...f,[k]:v}));
  const ss=(k,v)=>setSForm(f=>({...f,[k]:v}));

  // All tails seen across flights and compliance items
  const allTails=[...new Set([...entries.map(e=>e.tail), ...compItems.map(c=>c.tail)])];
  const activeReqTail = reqTail || allTails[0] || "";

  const setRequiredDate=(tail,key,value)=>{
    const next={...requiredItems, [tail]:{...(requiredItems[tail]||{}), [key]:value||null, [key+"_by"]:value?(profile?.name||""):null}};
    setRequiredItems(next); saveLS(KEY_REQ,next);
  };

  const saveComp=()=>{
    if(!cForm.name.trim()||!cForm.tail.trim()||!cForm.due){toast("Name, tail, and due value required.");return;}
    const item={id:Date.now(),...cForm,tail:cForm.tail.trim().toUpperCase(),name:cForm.name.trim(),addedBy:profile?.name||""};
    const n=[item,...compItems];setCompItems(n);saveLS(KEY_COMP,n);setCForm(f=>({...f,name:"",tail:"",due:"",warn:"",notes:""}));toast("Added.");
  };
  const applyComp=u=>{const n=compItems.map(c=>c.id===u.id?u:c);setCompItems(n);saveLS(KEY_COMP,n);setEditC(null);toast("Updated.");};
  const delComp=id=>{if(!window.confirm("Delete?"))return;const n=compItems.filter(c=>c.id!==id);setCompItems(n);saveLS(KEY_COMP,n);};
  const saveSq=()=>{
    if(!sForm.title.trim()||!sForm.tail.trim()){toast("Title and tail required.");return;}
    const item={id:Date.now(),...sForm,tail:sForm.tail.trim().toUpperCase(),title:sForm.title.trim(),resolved:false,resolution:null,reportedBy:profile?.name||""};
    const n=[item,...squawks];setSquawks(n);saveLS(KEY_SQ,n);setSForm(f=>({...f,title:"",tail:"",description:""}));toast("Squawk logged.");
  };
  const resolveSq=(id,res)=>{const n=squawks.map(s=>s.id===id?{...s,resolved:true,resolution:{...res,resolvedBy:profile?.name||""}}:s);setSquawks(n);saveLS(KEY_SQ,n);setResolving(null);toast("Squawk resolved.");};
  const delSq=id=>{if(!window.confirm("Delete?"))return;const n=squawks.filter(s=>s.id!==id);setSquawks(n);saveLS(KEY_SQ,n);};

  const openSq=squawks.filter(s=>!s.resolved);
  const archived=squawks.filter(s=>s.resolved);
  const DL={date:"Due date",hobbs:"Due at Hobbs",tach:"Due at Tach"};
  const WL={date:"Warn (days)",hobbs:"Warn (hrs)",tach:"Warn (hrs)"};

  const SubBtn=({id,label,badge})=>(
    <button onClick={()=>setSub(id)} style={{background:"none",border:"none",cursor:"pointer",padding:"7px 14px",fontSize:13,fontFamily:"'Barlow',sans-serif",color:sub===id?C.text:C.muted,borderBottom:`2px solid ${sub===id?C.accent:"transparent"}`,fontWeight:sub===id?500:400,display:"inline-flex",alignItems:"center",gap:5,whiteSpace:"nowrap"}}>
      {label}{badge>0&&<span style={{background:"#991b1b",color:C.danger,borderRadius:10,fontSize:10,padding:"0 5px",fontFamily:"'JetBrains Mono',monospace"}}>{badge}</span>}
    </button>
  );

  return (
    <div>
      {editC&&<EditCompModal item={editC} onSave={applyComp} onClose={()=>setEditC(null)}/>}
      {resolving&&<ResolveModal sq={resolving} onResolve={resolveSq} onClose={()=>setResolving(null)}/>}
      {viewing&&<SquawkDetailModal sq={viewing} onClose={()=>setViewing(null)}/>}

      <div style={{display:"flex",borderBottom:`1px solid ${C.border}`,marginBottom:14,overflowX:"auto"}}>
        <SubBtn id="compliance" label="Inspections & ADs"/>
        <SubBtn id="squawks" label="Squawks" badge={openSq.length}/>
        <SubBtn id="archive" label="Archive"/>
      </div>

      {sub==="compliance"&&(
        <div>
          <Card>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14,flexWrap:"wrap",gap:8}}>
              <div style={{fontSize:14,fontWeight:500}}>Required inspections</div>
              {allTails.length>1&&(
                <select value={activeReqTail} onChange={e=>setReqTail(e.target.value)} style={{width:"auto",fontSize:12}}>
                  {allTails.map(t=><option key={t} value={t}>{t}</option>)}
                </select>
              )}
            </div>
            {!activeReqTail&&<div style={{fontSize:13,color:C.muted,marginBottom:4}}>Log a flight or add a compliance item for an aircraft to track its required inspections.</div>}
            {activeReqTail&&REQUIRED_DEFS.map(def=>{
              const rec=requiredItems?.[activeReqTail]||{};
              const lastDone=rec[def.key]||"";
              const st=requiredStatus(lastDone, def.months);
              return (
                <div key={def.key} style={{display:"flex",alignItems:"flex-end",gap:10,marginBottom:10,flexWrap:"wrap"}}>
                  <div style={{flex:"1 1 220px"}}>
                    <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:4,flexWrap:"wrap"}}>
                      <span style={{fontSize:13,fontWeight:500}}>{def.label}</span>
                      <Badge cls={st.cls}>{st.missing?"Missing":st.label}</Badge>
                    </div>
                    <div style={{fontSize:11,color:C.muted}}>{def.ref} · every {def.months} months{st.due?` · next due ${st.due.toISOString().split("T")[0]}`:""}{rec[def.key+"_by"]?` · recorded by ${rec[def.key+"_by"]}`:""}</div>
                  </div>
                  <div style={{width:160}}>
                    <Field label="Last completed">
                      <input type="date" value={lastDone} onChange={e=>setRequiredDate(activeReqTail, def.key, e.target.value)}/>
                    </Field>
                  </div>
                </div>
              );
            })}
          </Card>

          <Card>
            <div style={{fontSize:14,fontWeight:500,marginBottom:14}}>Add inspection / AD</div>
            <Row cols={2}><Field label="Item name"><input placeholder="Annual, AD 2024-01-05…" value={cForm.name} onChange={e=>sc("name",e.target.value)}/></Field><Field label="Tail #"><input placeholder="N12345" value={cForm.tail} onChange={e=>sc("tail",e.target.value)}/></Field></Row>
            <Row cols={3}><Field label="Type"><select value={cForm.type} onChange={e=>sc("type",e.target.value)}><option value="date">By date</option><option value="hobbs">Hobbs hrs</option><option value="tach">Tach hrs</option></select></Field><Field label={DL[cForm.type]}><input type={cForm.type==="date"?"date":"number"} step="0.1" placeholder={cForm.type==="date"?"":"1250.0"} value={cForm.due} onChange={e=>sc("due",e.target.value)}/></Field><Field label={WL[cForm.type]}><input type="number" placeholder={cForm.type==="date"?"30":"10"} value={cForm.warn} onChange={e=>sc("warn",e.target.value)}/></Field></Row>
            <Field label="Notes / AD reference"><textarea placeholder="FAA AD #, SB reference…" value={cForm.notes} onChange={e=>sc("notes",e.target.value)} style={{marginBottom:12}}/></Field>
            <Btn variant="primary" onClick={saveComp}>+ Add item</Btn>
          </Card>
          {!compItems.length&&<div style={{textAlign:"center",padding:"2rem",color:C.muted,fontSize:14}}>No items yet.</div>}
          {compItems.map(item=>{const st=compStatus(item,entries);return(
            <div key={item.id} style={{background:C.panel,border:`1px solid ${C.border}`,borderRadius:8,padding:"10px 14px",marginBottom:8,display:"flex",alignItems:"flex-start",justifyContent:"space-between",gap:12,animation:"slideIn .2s ease"}}>
              <div style={{flex:1}}><div style={{display:"flex",alignItems:"center",gap:8,marginBottom:4,flexWrap:"wrap"}}><span style={{fontSize:14,fontWeight:500}}>{item.name}</span><Badge cls={st.cls}>{st.label}</Badge></div><div style={{fontSize:12,color:C.muted}}>{item.tail} · {item.type==="date"?"By date":item.type==="hobbs"?"Hobbs hrs":"Tach hrs"} · Due: {item.due}{item.notes&&` · ${item.notes}`}{item.addedBy&&` · Added by ${item.addedBy}`}</div></div>
              <div style={{display:"flex",gap:6,flexShrink:0}}><Btn variant="ghost" size="sm" onClick={()=>setEditC(item)}>✎</Btn><Btn variant="danger" size="sm" onClick={()=>delComp(item.id)}>✕</Btn></div>
            </div>
          );})}
        </div>
      )}

      {sub==="squawks"&&(
        <div>
          <Card>
            <div style={{fontSize:14,fontWeight:500,marginBottom:14}}>Log a squawk</div>
            <Row cols={2}><Field label="Issue / squawk title"><input placeholder="Oil pressure fluctuating…" value={sForm.title} onChange={e=>ss("title",e.target.value)}/></Field><Field label="Tail #"><input placeholder="N12345" value={sForm.tail} onChange={e=>ss("tail",e.target.value)}/></Field></Row>
            <Row cols={2} mb={4}>
              <Field label="Date reported"><input type="date" value={sForm.reportedDate} onChange={e=>ss("reportedDate",e.target.value)}/></Field>
              <Field label="Severity">
                <select value={sForm.severity} onChange={e=>ss("severity",e.target.value)}>
                  {SQUAWK_SEVERITIES.map(s=><option key={s.value} value={s.value}>{s.label}</option>)}
                </select>
              </Field>
            </Row>
            <div style={{fontSize:11,color:C.muted,marginBottom:10}}>{SQUAWK_SEVERITIES.find(s=>s.value===sForm.severity)?.ref}</div>
            <Field label="Description (optional)"><textarea placeholder="Additional details, when it occurs…" value={sForm.description} onChange={e=>ss("description",e.target.value)} style={{marginBottom:12}}/></Field>
            <Btn variant="warn" onClick={saveSq}>⚠ Log squawk</Btn>
          </Card>
          {!openSq.length&&<div style={{textAlign:"center",padding:"2rem",color:C.muted,fontSize:14}}>No open squawks — aircraft is clear.</div>}
          {openSq.map(sq=>(
            <div key={sq.id} style={{background:C.panel,border:`1px solid ${sq.severity==="grounding"?"#991b1b":C.border}`,borderRadius:8,padding:"10px 14px",marginBottom:8,display:"flex",alignItems:"flex-start",justifyContent:"space-between",gap:12,animation:"slideIn .2s ease"}}>
              <div style={{flex:1}}><div style={{display:"flex",alignItems:"center",gap:8,marginBottom:4,flexWrap:"wrap"}}><Badge cls="danger">OPEN</Badge>{sq.severity==="grounding"&&<Badge cls="danger">⛔ GROUNDING</Badge>}{sq.severity==="major"&&<Badge cls="warn">MAJOR</Badge>}{(!sq.severity||sq.severity==="minor")&&<Badge cls="blue">MINOR</Badge>}<span style={{fontSize:14,fontWeight:500}}>{sq.title}</span></div><div style={{fontSize:12,color:C.muted}}>{sq.tail} · Reported {sq.reportedDate}{sq.description&&` · ${sq.description}`}{sq.reportedBy&&` · by ${sq.reportedBy}`}</div></div>
              <div style={{display:"flex",gap:6,flexShrink:0}}><Btn variant="success" size="sm" onClick={()=>setResolving(sq)}>✓ Resolve</Btn><Btn variant="ghost" size="sm" onClick={()=>setViewing(sq)}>View</Btn><Btn variant="danger" size="sm" onClick={()=>delSq(sq.id)}>✕</Btn></div>
            </div>
          ))}
        </div>
      )}

      {sub==="archive"&&(
        <div>
          {!archived.length&&<div style={{textAlign:"center",padding:"2rem",color:C.muted,fontSize:14}}>No resolved squawks yet.</div>}
          {archived.map(sq=>(
            <div key={sq.id} style={{background:C.panel,border:"1px solid #166534",borderRadius:8,padding:"10px 14px",marginBottom:8,display:"flex",alignItems:"flex-start",justifyContent:"space-between",gap:12,animation:"slideIn .2s ease"}}>
              <div style={{flex:1}}><div style={{display:"flex",alignItems:"center",gap:8,marginBottom:4,flexWrap:"wrap"}}><Badge cls="ok">RESOLVED</Badge><span style={{fontSize:14,fontWeight:500}}>{sq.title}</span></div><div style={{fontSize:12,color:C.muted,marginBottom:4}}>{sq.tail} · Reported {sq.reportedDate}{sq.reportedBy&&` by ${sq.reportedBy}`} · Resolved {sq.resolution?.resolvedDate}{sq.resolution?.resolvedBy&&` by ${sq.resolution.resolvedBy}`}</div>{sq.resolution?.action&&<div style={{fontSize:12,color:C.muted,fontStyle:"italic",borderLeft:"2px solid #166534",paddingLeft:8}}>{sq.resolution.action}</div>}</div>
              <div style={{display:"flex",gap:6,flexShrink:0}}><Btn variant="ghost" size="sm" onClick={()=>setViewing(sq)}>View</Btn><Btn variant="danger" size="sm" onClick={()=>delSq(sq.id)}>✕</Btn></div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ── Summary tab ───────────────────────────────────────────────────────────────

function SummaryTab({entries,compItems,squawks,requiredItems,aircraft}) {
  const tails=[...new Set(entries.map(e=>e.tail))];
  if(!entries.length) return <div style={{textAlign:"center",padding:"3rem",color:C.muted,fontSize:14}}>Log some flights to see a summary.</div>;
  return (
    <div>
      {tails.map(tail=>{
        const latest=latestForTail(entries, tail);
        const te=entries.filter(e=>e.tail===tail);
        const comp=compItems.filter(c=>c.tail===tail);
        const openSq=(squawks||[]).filter(s=>s.tail===tail&&!s.resolved);
        const groundingSq=openSq.filter(s=>s.severity==="grounding");
        const sts=comp.map(c=>compStatus(c,entries));
        const overdue=sts.filter(s=>s.cls==="danger").length;
        const warning=sts.filter(s=>s.cls==="warn").length;
        const req=requiredInspectionsStatus(requiredItems, tail);
        const reqOverdue=req.items.filter(i=>i.status.cls==="danger").length;
        const reqWarn=req.items.filter(i=>i.status.cls==="warn").length;
        const acRec=aircraft?.[tail];
        const hh=acRec?.hundredHourEnabled?hundredHourStatus(entries,tail,acRec.lastHundredHourTach):null;
        const hhOverdue=hh&&hh.cls==="danger";
        const hhWarn=hh&&hh.cls==="warn";
        const oilRate=oilConsumptionRate(entries,tail);
        const anyOverdue = overdue>0 || openSq.length>0 || reqOverdue>0 || hhOverdue;
        const anyWarn = !anyOverdue && (warning>0 || reqWarn>0 || hhWarn);
        const allCurrent = !anyOverdue && !anyWarn;
        return (
          <Card key={tail} style={groundingSq.length>0?{border:"1px solid #991b1b"}:{}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:14}}>
              <div><div style={{fontSize:11,color:C.muted,textTransform:"uppercase",letterSpacing:".1em",marginBottom:2}}>Aircraft</div><div style={{fontSize:22,fontWeight:600,fontFamily:"'JetBrains Mono',monospace",color:C.accent}}>{tail}</div>{(acRec?.make||acRec?.model)&&<div style={{fontSize:12,color:C.muted,marginTop:2}}>{[acRec?.year,acRec?.make,acRec?.model].filter(Boolean).join(" ")}</div>}</div>
              <div style={{display:"flex",flexDirection:"column",gap:4,alignItems:"flex-end"}}>
                {groundingSq.length>0&&<Badge cls="danger">⛔ Grounded</Badge>}
                {overdue>0&&<Badge cls="danger">⚠ {overdue} compliance overdue</Badge>}
                {openSq.filter(s=>s.severity!=="grounding").length>0&&<Badge cls="danger">⚠ {openSq.filter(s=>s.severity!=="grounding").length} open squawk{openSq.length>1?"s":""}</Badge>}
                {reqOverdue>0&&<Badge cls="danger">⚠ {reqOverdue} required insp. overdue</Badge>}
                {hhOverdue&&<Badge cls="danger">⚠ 100-hr overdue</Badge>}
                {!anyOverdue&&warning>0&&<Badge cls="warn">{warning} compliance due soon</Badge>}
                {!anyOverdue&&reqWarn>0&&<Badge cls="warn">{reqWarn} required insp. due soon</Badge>}
                {!anyOverdue&&hhWarn&&<Badge cls="warn">100-hr due soon</Badge>}
                {allCurrent&&<Badge cls="ok">✓ All current</Badge>}
              </div>
            </div>
            <div style={{fontSize:11,color:C.muted,textTransform:"uppercase",letterSpacing:".08em",marginBottom:8}}>Current totals — last entry: {latest.date}</div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8,marginBottom:12}}>
              {[["Hobbs",fmt1(latest.hobbs)],["Tach",fmt1(latest.tach)],["Fuel on board",latest.fuel!=null?`${fmt1(latest.fuel)} gal`:"—"],["Last oil added",latest.oil!=null?`${latest.oil} qt`:"—"]].map(([k,v])=><StatBox key={k} label={k} value={v}/>)}
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8,marginBottom:12}}>
              {[["Flights",te.length],["Total gal added",te.reduce((s,e)=>s+(e.gallonsAdded||0),0).toFixed(1)],["Total oil added",`${te.reduce((s,e)=>s+(e.oil||0),0)} qt`]].map(([k,v])=><StatBox key={k} label={k} value={v}/>)}
            </div>
            {oilRate!=null&&(
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12,background:C.bg,borderRadius:6,padding:"8px 10px"}}>
                <span style={{fontSize:12,color:C.muted}}>Avg. oil consumption (last {Math.min(5,te.length-1)} flights)</span>
                <span style={{fontSize:13,fontFamily:"'JetBrains Mono',monospace",color:oilRate>=1?C.warn:C.text}}>{oilRate.toFixed(2)} qt/hr{oilRate>=1?" ⚠":""}</span>
              </div>
            )}
            {groundingSq.length>0&&<><div style={{borderTop:"1px solid #991b1b",paddingTop:12,marginBottom:8,fontSize:11,color:C.danger,textTransform:"uppercase",letterSpacing:".08em"}}>Grounding squawks</div>{groundingSq.map(s=><div key={s.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}><span style={{fontSize:13}}>{s.title}</span><Badge cls="danger">⛔ GROUNDING</Badge></div>)}</>}
            {openSq.filter(s=>s.severity!=="grounding").length>0&&<><div style={{borderTop:"1px solid #991b1b",paddingTop:12,marginBottom:8,fontSize:11,color:C.danger,textTransform:"uppercase",letterSpacing:".08em"}}>Open squawks</div>{openSq.filter(s=>s.severity!=="grounding").map(s=><div key={s.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}><span style={{fontSize:13}}>{s.title}</span><Badge cls={s.severity==="major"?"warn":"blue"}>{(s.severity||"minor").toUpperCase()}</Badge></div>)}</>}
            <div style={{borderTop:`1px solid ${C.border}`,paddingTop:12,marginBottom:8,fontSize:11,color:C.muted,textTransform:"uppercase",letterSpacing:".08em"}}>Required inspections</div>
            {req.items.map(i=><div key={i.key} style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}><span style={{fontSize:13}}>{i.label}</span><Badge cls={i.status.cls}>{i.status.missing?"Missing":i.status.label}</Badge></div>)}
            {hh&&<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}><span style={{fontSize:13}}>100-hour inspection</span><Badge cls={hh.cls}>{hh.missing?"Not recorded":hh.label}</Badge></div>}
            {comp.length>0&&<><div style={{borderTop:`1px solid ${C.border}`,paddingTop:12,marginBottom:8,fontSize:11,color:C.muted,textTransform:"uppercase",letterSpacing:".08em"}}>Compliance</div>{comp.map(c=>{const s=compStatus(c,entries);return<div key={c.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}><span style={{fontSize:13}}>{c.name}</span><Badge cls={s.cls}>{s.label}</Badge></div>;})}</>}
          </Card>
        );
      })}
    </div>
  );
}

// ── App ────────────────────────────────────────────────────────────────────────

const TABS=[{id:"dispatch",label:"Dispatch"},{id:"times",label:"Times"},{id:"compliance",label:"Compliance"},{id:"aircraft",label:"Aircraft"},{id:"summary",label:"Summary"},{id:"profile",label:"Profile"}];

export default function AircraftLog() {
  const [entries,setEntries]   = useState(()=>loadLS(KEY_LOG)||[]);
  const [compItems,setCompItems]= useState(()=>loadLS(KEY_COMP)||[]);
  const [squawks,setSquawks]   = useState(()=>loadLS(KEY_SQ)||[]);
  const [dispatch,setDispatch] = useState(()=>loadLS(KEY_DISP)||null);
  const [profile,setProfile]   = useState(()=>loadLS(KEY_PROF)||null);
  const [requiredItems,setRequiredItems] = useState(()=>loadLS(KEY_REQ)||{});
  const [aircraft,setAircraft] = useState(()=>loadLS(KEY_AC)||{});
  const [tab,setTab]           = useState(dispatch?"dispatch":"times");
  const [toasts,setToasts]     = useState([]);

  const toast = msg => setToasts(t=>[...t,{id:Date.now()+Math.random(),msg}]);
  const rmToast = id => setToasts(t=>t.filter(x=>x.id!==id));

  const allTails = [...new Set(entries.map(e=>e.tail))];
  const overdueComp = compItems.filter(c=>compStatus(c,entries).cls==="danger");
  const openSquawks = squawks.filter(s=>!s.resolved);
  const groundingSquawks = openSquawks.filter(s=>s.severity==="grounding");
  const overdueReq = allTails.flatMap(tail=>requiredInspectionsStatus(requiredItems,tail).items.filter(i=>i.status.cls==="danger"));
  const overdue100hr = allTails.filter(tail=>{
    const rec=aircraft?.[tail];
    if(!rec?.hundredHourEnabled) return false;
    const hh=hundredHourStatus(entries,tail,rec.lastHundredHourTach);
    return hh.cls==="danger";
  });
  const alert = overdueComp.length>0||openSquawks.length>0||overdueReq.length>0||overdue100hr.length>0;

  return (
    <>
      <style>{CSS}</style>

      {/* ── Header ── */}
      <div style={{background:C.panel,borderBottom:`1px solid ${C.border}`,padding:"0 16px",position:"sticky",top:0,zIndex:50}}>
        <div style={{maxWidth:700,margin:"0 auto"}}>
          <div style={{padding:"12px 0 0",display:"flex",alignItems:"center",gap:12}}>
            <span style={{fontSize:16,fontWeight:700,fontFamily:"'JetBrains Mono',monospace",color:C.accent}}>✈</span>
            <span style={{fontSize:15,fontWeight:600}}>AircraftLog</span>
            {dispatch&&<Badge cls="blue">⏱ {dispatch.tail}</Badge>}
            <div style={{marginLeft:"auto",display:"flex",alignItems:"center",gap:10}}>
              {profile?.homeAirport&&<span style={{fontFamily:"'JetBrains Mono',monospace",fontSize:12,color:C.muted}}>{profile.homeAirport}</span>}
              <div onClick={()=>setTab("profile")} style={{width:30,height:30,borderRadius:"50%",overflow:"hidden",border:`1.5px solid ${C.border}`,cursor:"pointer",background:C.bg,display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,color:C.muted,flexShrink:0}}>
                {profile?.photo?<img src={profile.photo} alt="" style={{width:"100%",height:"100%",objectFit:"cover"}}/>:profile?.name?profile.name[0].toUpperCase():"?"}
              </div>
            </div>
          </div>
          <div style={{display:"flex",marginTop:8,overflowX:"auto"}}>
            {TABS.map(t=>(
              <button key={t.id} onClick={()=>setTab(t.id)} style={{background:"none",border:"none",cursor:"pointer",padding:"8px 14px",fontSize:13,fontFamily:"'Barlow',sans-serif",color:tab===t.id?C.text:C.muted,borderBottom:`2px solid ${tab===t.id?C.accent:"transparent"}`,fontWeight:tab===t.id?500:400,position:"relative",whiteSpace:"nowrap",transition:"color .15s"}}>
                {t.label}
                {t.id==="dispatch"&&dispatch&&<span style={{position:"absolute",top:5,right:5,width:6,height:6,borderRadius:"50%",background:C.accent,animation:"pdot 2s infinite"}}/>}
                {t.id==="compliance"&&alert&&<span style={{position:"absolute",top:5,right:5,width:6,height:6,borderRadius:"50%",background:C.danger}}/>}
              </button>
            ))}
          </div>
        </div>
      </div>

      {/* ── Body ── */}
      <div style={{maxWidth:700,margin:"0 auto",padding:"16px 12px 60px"}}>
        {alert&&tab!=="compliance"&&(
          <div style={{background:"#2d0a0a",border:"1px solid #991b1b",borderRadius:8,padding:"10px 14px",marginBottom:12,display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8}}>
            <span style={{color:C.danger,fontSize:13,fontWeight:500}}>
              {[groundingSquawks.length>0&&`⛔ ${groundingSquawks.length} grounding squawk${groundingSquawks.length>1?"s":""}`,overdueComp.length>0&&`${overdueComp.length} compliance item${overdueComp.length>1?"s":""} overdue`,openSquawks.length>0&&`${openSquawks.length} open squawk${openSquawks.length>1?"s":""}`,overdueReq.length>0&&`${overdueReq.length} required inspection${overdueReq.length>1?"s":""} overdue/missing`,overdue100hr.length>0&&`100-hr inspection overdue: ${overdue100hr.join(", ")}`].filter(Boolean).join(" · ")}
            </span>
            <Btn size="sm" variant="danger" onClick={()=>setTab("compliance")}>View →</Btn>
          </div>
        )}
        {tab==="dispatch"  &&<DispatchTab   entries={entries} dispatch={dispatch} setDispatch={setDispatch} compItems={compItems} squawks={squawks} requiredItems={requiredItems} aircraft={aircraft} toast={toast}/>}
        {tab==="times"     &&<TimesTab      entries={entries} setEntries={setEntries} profile={profile} toast={toast}/>}
        {tab==="compliance"&&<ComplianceTab entries={entries} compItems={compItems} setCompItems={setCompItems} squawks={squawks} setSquawks={setSquawks} requiredItems={requiredItems} setRequiredItems={setRequiredItems} profile={profile} toast={toast}/>}
        {tab==="aircraft"  &&<AircraftTab   entries={entries} aircraft={aircraft} setAircraft={setAircraft} toast={toast}/>}
        {tab==="summary"   &&<SummaryTab    entries={entries} compItems={compItems} squawks={squawks} requiredItems={requiredItems} aircraft={aircraft}/>}
        {tab==="profile"   &&<ProfileTab    profile={profile} setProfile={setProfile} toast={toast}/>}
      </div>

      {/* ── Toasts ── */}
      <div style={{position:"fixed",bottom:16,right:16,zIndex:999,display:"flex",flexDirection:"column",gap:8}}>
        {toasts.map(t=><Toast key={t.id} msg={t.msg} onDone={()=>rmToast(t.id)}/>)}
      </div>
    </>
  );
}
