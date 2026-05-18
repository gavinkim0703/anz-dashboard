<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>ANZ Cadence Dashboard</title>
  <script src="https://unpkg.com/react@18/umd/react.development.js" crossorigin></script>
  <script src="https://unpkg.com/react-dom@18/umd/react-dom.development.js" crossorigin></script>
  <script src="https://unpkg.com/@babel/standalone/babel.min.js"></script>
  <style>
    *, *::before, *::after { box-sizing: border-box; }
    body { margin: 0; font-family: 'Segoe UI', 'SF Pro Display', system-ui, sans-serif; background: #F0F4F8; }
  </style>
</head>
<body>
  <div id="root"></div>
  <script type="text/babel">
    const { useState, useEffect, useRef } = React;

    const PILLARS = [
      { id: "corp",       label: "Corp Field",  subtitle: "Partner Driven & Next Gen Accounts", color: "#1D4ED8", light: "#DBEAFE", dark: "#1E3A8A", owners: ["Dan Casey","Sue Johnston","Audrey Sleaford","Will Santry"] },
      { id: "move",       label: "MOVE",         subtitle: "",                                   color: "#7C3AED", light: "#EDE9FE", dark: "#4C1D95", owners: ["Sally Kingston"] },
      { id: "crosssell",  label: "Cross Sell",   subtitle: "",                                   color: "#0891B2", light: "#CFFAFE", dark: "#164E63", owners: ["Linda Davis"] },
      { id: "upsell",     label: "Upsell",        subtitle: "",                                   color: "#059669", light: "#D1FAE5", dark: "#064E3B", owners: ["Kevin Angland"] },
      { id: "largedeals", label: "Large Deals",   subtitle: "",                                   color: "#D97706", light: "#FEF3C7", dark: "#78350F", owners: ["Dan Getliffe","Brian Senior"] },
    ];

    const ALL_OWNERS  = ["Dan Casey","Sue Johnston","Audrey Sleaford","Will Santry","Sally Kingston","Linda Davis","Kevin Angland","Dan Getliffe","Brian Senior"];
    const PRIORITIES  = ["Critical","High","Medium","Low"];
    const STATUSES    = ["Pending","In Progress","On Hold","Resolved"];
    const PRIORITY_META = { Critical:{bg:"#FEE2E2",text:"#991B1B",dot:"#EF4444"}, High:{bg:"#FEF3C7",text:"#92400E",dot:"#F59E0B"}, Medium:{bg:"#DBEAFE",text:"#1E40AF",dot:"#3B82F6"}, Low:{bg:"#D1FAE5",text:"#065F46",dot:"#10B981"} };
    const STATUS_META   = { Pending:{bg:"#F1F5F9",text:"#475569",dot:"#94A3B8"}, "In Progress":{bg:"#DBEAFE",text:"#1E3A8A",dot:"#3B82F6"}, "On Hold":{bg:"#FEF3C7",text:"#78350F",dot:"#F59E0B"}, Resolved:{bg:"#D1FAE5",text:"#064E3B",dot:"#10B981"} };

    const today = new Date().toISOString().slice(0,10);
    const STORAGE_KEY_TASKS          = "anz_tasks";
    const STORAGE_KEY_SESSIONS       = "anz_sessions";
    const STORAGE_KEY_NEXT_TASK_ID   = "anz_next_task_id";
    const STORAGE_KEY_NEXT_SESSION_ID = "anz_next_session_id";

    const INIT_SESSIONS = [{ id:1, date:today, title:"ANZ Fortnightly Cadence — Session 1", summary:"", keyPoints:[""], decisions:[""], actionItems:[""] }];

    // localStorage-backed persistence (works in any browser, no server needed)
    function persist(key, value) {
      try { localStorage.setItem(key, JSON.stringify(value)); } catch(e) { console.error("Storage error", e); }
    }
    function load(key) {
      try { const r = localStorage.getItem(key); return r ? JSON.parse(r) : null; } catch(e) { return null; }
    }

    // ─── SharePoint REST API ─────────────────────────────────────────────────────
    const SP_SITE    = 'https://sap.sharepoint.com/teams/ANZFortnightlyCadence';
    const SP_FILE_ID = '6fb1d590-1b0e-4866-afab-2e91918df428';

    async function spRead() {
      const r = await fetch(`${SP_SITE}/_api/web/GetFileById('${SP_FILE_ID}')/$value`, {
        credentials: 'include', headers: { Accept: 'application/json' }
      });
      if (!r.ok) throw new Error(r.status);
      return r.json();
    }
    async function spWrite(data) {
      const dr = await fetch(`${SP_SITE}/_api/contextinfo`, {
        method: 'POST', credentials: 'include',
        headers: { Accept: 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose' }
      });
      if (!dr.ok) throw new Error('digest ' + dr.status);
      const dj = await dr.json();
      const digest = dj.d.GetContextWebInformation.FormDigestValue;
      const r = await fetch(`${SP_SITE}/_api/web/GetFileById('${SP_FILE_ID}')/$value`, {
        method: 'POST', credentials: 'include',
        headers: { 'X-RequestDigest': digest, 'X-HTTP-Method': 'PUT', 'Content-Type': 'application/json' },
        body: JSON.stringify(data, null, 2)
      });
      if (!r.ok) throw new Error(r.status);
    }

    // ─── Sub-components ──────────────────────────────────────────────────────────

    function Avatar({ name, size=28 }) {
      const parts = name.split(" ");
      const initials = (parts[0][0]+(parts[1]?.[0]||"")).toUpperCase();
      const palette = ["#1D4ED8","#7C3AED","#0891B2","#059669","#D97706","#DC2626","#9333EA","#0284C7","#16A34A"];
      const idx = name.charCodeAt(0) % palette.length;
      return <div style={{ width:size, height:size, borderRadius:"50%", background:palette[idx], color:"#fff", fontSize:size*0.38, fontWeight:700, display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0, letterSpacing:"0.02em" }}>{initials}</div>;
    }

    function Chip({ label, meta }) {
      return <span style={{ display:"inline-flex", alignItems:"center", gap:5, background:meta.bg, color:meta.text, fontSize:11, fontWeight:600, padding:"2px 8px", borderRadius:20 }}>
        <span style={{ width:6, height:6, borderRadius:"50%", background:meta.dot, flexShrink:0 }} />{label}
      </span>;
    }

    function ProgressBar({ value, color }) {
      return <div style={{ background:"#E2E8F0", borderRadius:8, height:5, width:"100%", overflow:"hidden" }}>
        <div style={{ width:`${value}%`, background:color, height:"100%", borderRadius:8, transition:"width 0.4s" }} />
      </div>;
    }

    const inputStyle = { width:"100%", border:"1.5px solid #E2E8F0", borderRadius:8, padding:"8px 12px", fontSize:13, color:"#0F172A", fontFamily:"inherit", boxSizing:"border-box", outline:"none", background:"#F8FAFC" };
    const labelStyle = { display:"block", fontSize:11, fontWeight:700, color:"#64748B", marginBottom:5, textTransform:"uppercase", letterSpacing:"0.05em" };

    function TaskModal({ task, onSave, onClose }) {
      const isNew = !task;
      const [form, setForm] = useState(task ? {...task} : { pillarId:PILLARS[0].id, title:"", owner:ALL_OWNERS[0], priority:"Medium", status:"Pending", progress:0, dueDate:"", notes:"", sessionId:null });
      const set = (k,v) => setForm(f => ({...f,[k]:v}));
      const pillar = PILLARS.find(p => p.id===form.pillarId);
      return (
        <div style={{ position:"fixed", inset:0, background:"rgba(15,23,42,0.6)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:1000 }}>
          <div style={{ background:"#fff", borderRadius:18, padding:"28px 32px", width:500, maxWidth:"95vw", maxHeight:"90vh", overflowY:"auto" }}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:22 }}>
              <div style={{ display:"flex", alignItems:"center", gap:10 }}>
                <div style={{ width:4, height:24, borderRadius:2, background:pillar.color }} />
                <h3 style={{ margin:0, fontSize:17, fontWeight:800, color:"#0F172A" }}>{isNew?"New Action Item":"Edit Action Item"}</h3>
              </div>
              <button onClick={onClose} style={{ border:"none", background:"none", cursor:"pointer", fontSize:22, color:"#94A3B8", lineHeight:1 }}>×</button>
            </div>
            <div style={{ marginBottom:14 }}>
              <label style={labelStyle}>Pillar</label>
              <select value={form.pillarId} onChange={e => set("pillarId",e.target.value)} style={inputStyle}>
                {PILLARS.map(p => <option key={p.id} value={p.id}>{p.label}{p.subtitle?` — ${p.subtitle}`:""}</option>)}
              </select>
            </div>
            <div style={{ marginBottom:14 }}>
              <label style={labelStyle}>Action / Task</label>
              <input value={form.title} onChange={e => set("title",e.target.value)} placeholder="Describe the action item..." style={inputStyle} />
            </div>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14, marginBottom:14 }}>
              <div><label style={labelStyle}>Owner</label>
                <select value={form.owner} onChange={e => set("owner",e.target.value)} style={inputStyle}>{ALL_OWNERS.map(o => <option key={o}>{o}</option>)}</select>
              </div>
              <div><label style={labelStyle}>Priority</label>
                <select value={form.priority} onChange={e => set("priority",e.target.value)} style={inputStyle}>{PRIORITIES.map(p => <option key={p}>{p}</option>)}</select>
              </div>
            </div>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14, marginBottom:14 }}>
              <div><label style={labelStyle}>Status</label>
                <select value={form.status} onChange={e => set("status",e.target.value)} style={inputStyle}>{STATUSES.map(s => <option key={s}>{s}</option>)}</select>
              </div>
              <div><label style={labelStyle}>Due Date</label>
                <input type="date" value={form.dueDate} onChange={e => set("dueDate",e.target.value)} style={inputStyle} />
              </div>
            </div>
            <div style={{ marginBottom:14 }}>
              <label style={labelStyle}>Progress — {form.progress}%</label>
              <input type="range" min={0} max={100} step={1} value={form.progress} onChange={e => set("progress",+e.target.value)} style={{ width:"100%" }} />
            </div>
            <div style={{ marginBottom:20 }}>
              <label style={labelStyle}>Notes</label>
              <textarea value={form.notes} onChange={e => set("notes",e.target.value)} rows={3} placeholder="Any context or notes..." style={{ ...inputStyle, resize:"vertical" }} />
            </div>
            <div style={{ display:"flex", gap:10 }}>
              <button onClick={() => onSave(form)} style={{ flex:1, background:pillar.color, color:"#fff", border:"none", borderRadius:10, padding:"12px 0", fontWeight:800, fontSize:14, cursor:"pointer" }}>
                {isNew?"Add Action Item":"Save Changes"}
              </button>
              <button onClick={onClose} style={{ padding:"12px 20px", border:"1.5px solid #E2E8F0", borderRadius:10, background:"none", cursor:"pointer", color:"#64748B" }}>Cancel</button>
            </div>
          </div>
        </div>
      );
    }

    function SessionModal({ session, nextId, onSave, onClose }) {
      const isNew = !session;
      const [form, setForm] = useState(session
        ? {...session, keyPoints:[...(session.keyPoints||[""])], decisions:[...(session.decisions||[""])], actionItems:[...(session.actionItems||[""])]}
        : { date:today, title:`ANZ Fortnightly Cadence — Session ${nextId}`, summary:"", keyPoints:[""], decisions:[""], actionItems:[""] }
      );
      const set = (k,v) => setForm(f => ({...f,[k]:v}));
      const setList = (k,i,v) => setForm(f => { const a=[...f[k]]; a[i]=v; return {...f,[k]:a}; });
      const addList = (k) => setForm(f => ({...f,[k]:[...f[k],""]}));
      const removeList = (k,i) => setForm(f => { const a=f[k].filter((_,j)=>j!==i); return {...f,[k]:a.length?a:[""]}; });
      const ListEditor = ({ field, placeholder }) => (
        <div>
          {form[field].map((item,i) => (
            <div key={i} style={{ display:"flex", gap:8, marginBottom:6 }}>
              <input value={item} onChange={e => setList(field,i,e.target.value)} placeholder={placeholder} style={{ ...inputStyle, flex:1 }} />
              <button onClick={() => removeList(field,i)} style={{ border:"none", background:"#FEE2E2", borderRadius:6, padding:"0 10px", cursor:"pointer", color:"#EF4444", fontSize:16 }}>×</button>
            </div>
          ))}
          <button onClick={() => addList(field)} style={{ border:"1.5px dashed #CBD5E1", borderRadius:8, padding:"6px 14px", background:"none", cursor:"pointer", fontSize:12, color:"#64748B", width:"100%" }}>+ Add</button>
        </div>
      );
      return (
        <div style={{ position:"fixed", inset:0, background:"rgba(15,23,42,0.6)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:1000 }}>
          <div style={{ background:"#fff", borderRadius:18, padding:"28px 32px", width:580, maxWidth:"95vw", maxHeight:"90vh", overflowY:"auto" }}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:22 }}>
              <h3 style={{ margin:0, fontSize:17, fontWeight:800, color:"#0F172A" }}>{isNew?"New Session":"Edit Session"}</h3>
              <button onClick={onClose} style={{ border:"none", background:"none", cursor:"pointer", fontSize:22, color:"#94A3B8" }}>×</button>
            </div>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14, marginBottom:14 }}>
              <div><label style={labelStyle}>Session Date</label><input type="date" value={form.date} onChange={e => set("date",e.target.value)} style={inputStyle} /></div>
              <div><label style={labelStyle}>Session Title</label><input value={form.title} onChange={e => set("title",e.target.value)} style={inputStyle} /></div>
            </div>
            <div style={{ marginBottom:14 }}>
              <label style={labelStyle}>Executive Summary</label>
              <textarea value={form.summary} onChange={e => set("summary",e.target.value)} rows={3} placeholder="High-level summary of what was discussed..." style={{ ...inputStyle, resize:"vertical" }} />
            </div>
            <div style={{ marginBottom:14 }}><label style={labelStyle}>Key Discussion Points</label><ListEditor field="keyPoints" placeholder="e.g. Pipeline review for Corp Field..." /></div>
            <div style={{ marginBottom:14 }}><label style={labelStyle}>Decisions Made</label><ListEditor field="decisions" placeholder="e.g. Approved Large Deals threshold at $500K..." /></div>
            <div style={{ marginBottom:20 }}><label style={labelStyle}>Action Items Raised</label><ListEditor field="actionItems" placeholder="e.g. Dan Casey to follow up on Q3 targets..." /></div>
            <div style={{ display:"flex", gap:10 }}>
              <button onClick={() => onSave(form)} style={{ flex:1, background:"#1D4ED8", color:"#fff", border:"none", borderRadius:10, padding:"12px 0", fontWeight:800, fontSize:14, cursor:"pointer" }}>
                {isNew?"Create Session":"Save Session"}
              </button>
              <button onClick={onClose} style={{ padding:"12px 20px", border:"1.5px solid #E2E8F0", borderRadius:10, background:"none", cursor:"pointer", color:"#64748B" }}>Cancel</button>
            </div>
          </div>
        </div>
      );
    }

    // ─── Main App ─────────────────────────────────────────────────────────────────

    function ANZDashboard() {
      const [tasks,          setTasks]          = useState([]);
      const [sessions,       setSessions]       = useState([]);
      const [nextTaskId,     setNextTaskId]     = useState(1);
      const [nextSessionId,  setNextSessionId]  = useState(2);
      const [ready,          setReady]          = useState(false);
      const [activeView,     setActiveView]     = useState("pillars");
      const [taskModal,      setTaskModal]      = useState(null);
      const [sessionModal,   setSessionModal]   = useState(null);
      const [filterPillar,   setFilterPillar]   = useState("all");
      const [filterOwner,    setFilterOwner]    = useState("all");
      const [filterStatus,   setFilterStatus]   = useState("all");
      const [deleteTaskId,   setDeleteTaskId]   = useState(null);
      const [deleteSessionId,setDeleteSessionId]= useState(null);
      const [expandedSession,setExpandedSession]= useState(null);
      const importRef   = useRef(null);
      const [syncStatus, setSyncStatus] = useState('loading');
      const latestRef   = useRef(null);

      useEffect(() => {
        (async () => {
          let data = null;
          try { data = await spRead(); } catch(e) {}
          if (!data) {
            try {
              const res = await fetch('./data.json?t=' + Date.now());
              if (res.ok) { const j = await res.json(); if (j && (Array.isArray(j.tasks)||Array.isArray(j.sessions))) data = j; }
            } catch(e) {}
          }
          const source = data || {
            tasks: load(STORAGE_KEY_TASKS), sessions: load(STORAGE_KEY_SESSIONS),
            nextTaskId: load(STORAGE_KEY_NEXT_TASK_ID), nextSessionId: load(STORAGE_KEY_NEXT_SESSION_ID),
          };
          const loadedSessions = source.sessions ?? INIT_SESSIONS;
          setTasks(source.tasks ?? []);
          setSessions(loadedSessions);
          setNextTaskId(source.nextTaskId ?? 1);
          setNextSessionId(source.nextSessionId ?? 2);
          if (loadedSessions.length > 0) {
            const sorted = [...loadedSessions].sort((a,b) => b.date.localeCompare(a.date));
            setExpandedSession(sorted[0].id);
          }
          setReady(true);
          setSyncStatus(data ? 'synced' : 'offline');
        })();
      }, []);

      useEffect(() => { if (ready) persist(STORAGE_KEY_TASKS,           tasks);         }, [tasks,         ready]);
      useEffect(() => { if (ready) persist(STORAGE_KEY_SESSIONS,        sessions);      }, [sessions,      ready]);
      useEffect(() => { if (ready) persist(STORAGE_KEY_NEXT_TASK_ID,    nextTaskId);    }, [nextTaskId,    ready]);
      useEffect(() => { if (ready) persist(STORAGE_KEY_NEXT_SESSION_ID, nextSessionId); }, [nextSessionId, ready]);

      useEffect(() => { if (ready) latestRef.current = { tasks, sessions, nextTaskId, nextSessionId }; }, [tasks, sessions, nextTaskId, nextSessionId, ready]);

      useEffect(() => {
        if (!ready) return;
        const poll = async () => {
          try {
            const remote = await spRead();
            if (!remote || typeof remote !== 'object') return;
            const cur = JSON.stringify(latestRef.current);
            const inc = JSON.stringify({ tasks: remote.tasks, sessions: remote.sessions, nextTaskId: remote.nextTaskId, nextSessionId: remote.nextSessionId });
            if (cur !== inc) {
              if (Array.isArray(remote.tasks))    setTasks(remote.tasks);
              if (Array.isArray(remote.sessions)) setSessions(remote.sessions);
              if (remote.nextTaskId)    setNextTaskId(remote.nextTaskId);
              if (remote.nextSessionId) setNextSessionId(remote.nextSessionId);
            }
            setSyncStatus('synced');
          } catch(e) { setSyncStatus('offline'); }
        };
        const id = setInterval(poll, 30000);
        return () => clearInterval(id);
      }, [ready]);

      const writeToSP = (data) => {
        setSyncStatus('syncing');
        spWrite(data).then(() => setSyncStatus('synced')).catch(() => setSyncStatus('error'));
      };

      // ── Export / Import ──────────────────────────────────────────────────────────

      const exportJSON = () => {
        const data = { tasks, sessions, nextTaskId, nextSessionId, exportedAt: new Date().toISOString() };
        const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
        const url  = URL.createObjectURL(blob);
        const a    = document.createElement("a");
        a.href     = url;
        a.download = `anz_dashboard_${today}.json`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      };

      const exportCSV = () => {
        const header = ["ID","Pillar","Title","Owner","Priority","Status","Progress (%)","Due Date","Notes"];
        const rows   = tasks.map(t => {
          const pillarLabel = PILLARS.find(p => p.id === t.pillarId)?.label || t.pillarId;
          const esc = s => `"${(s||"").replace(/"/g,'""')}"`;
          return [t.id, pillarLabel, esc(t.title), t.owner, t.priority, t.status, t.progress, t.dueDate||"", esc(t.notes)].join(",");
        });
        const csv  = [header.join(","), ...rows].join("\n");
        const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
        const url  = URL.createObjectURL(blob);
        const a    = document.createElement("a");
        a.href     = url;
        a.download = `anz_tasks_${today}.csv`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      };

      const handleImport = (e) => {
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (ev) => {
          try {
            const data = JSON.parse(ev.target.result);
            if (Array.isArray(data.tasks))    setTasks(data.tasks);
            if (Array.isArray(data.sessions)) setSessions(data.sessions);
            if (data.nextTaskId)    setNextTaskId(data.nextTaskId);
            if (data.nextSessionId) setNextSessionId(data.nextSessionId);
            alert("Data imported successfully!");
          } catch {
            alert("Import failed: the file is not a valid ANZ dashboard JSON backup.");
          }
        };
        reader.readAsText(file);
        e.target.value = "";
      };

      // ── Task / Session CRUD ──────────────────────────────────────────────────────

      const saveTask = (form) => {
        let newTasks, newNextTaskId = nextTaskId;
        if (taskModal === "new") {
          newTasks = [...tasks, {...form, id: nextTaskId}];
          newNextTaskId = nextTaskId + 1;
          setTasks(newTasks);
          setNextTaskId(newNextTaskId);
        } else {
          newTasks = tasks.map(x => x.id===form.id ? form : x);
          setTasks(newTasks);
        }
        setTaskModal(null);
        writeToSP({ tasks: newTasks, sessions, nextTaskId: newNextTaskId, nextSessionId });
      };

      const saveSession = (form) => {
        let newSessions, newNextSessId = nextSessionId;
        if (sessionModal === "new") {
          newSessions = [...sessions, {...form, id: nextSessionId}];
          newNextSessId = nextSessionId + 1;
          setSessions(newSessions);
          setNextSessionId(newNextSessId);
          setExpandedSession(nextSessionId);
        } else {
          newSessions = sessions.map(x => x.id===form.id ? form : x);
          setSessions(newSessions);
        }
        setSessionModal(null);
        writeToSP({ tasks, sessions: newSessions, nextTaskId, nextSessionId: newNextSessId });
      };

      const deleteTask = (id) => {
        const newTasks = tasks.filter(x => x.id !== id);
        setTasks(newTasks);
        setDeleteTaskId(null);
        writeToSP({ tasks: newTasks, sessions, nextTaskId, nextSessionId });
      };

      const deleteSession = (id) => {
        const newSessions = sessions.filter(x => x.id !== id);
        setSessions(newSessions);
        setDeleteSessionId(null);
        writeToSP({ tasks, sessions: newSessions, nextTaskId, nextSessionId });
      };

      const filteredTasks = tasks.filter(t =>
        (filterPillar==="all" || t.pillarId===filterPillar) &&
        (filterOwner==="all"  || t.owner===filterOwner)     &&
        (filterStatus==="all" || t.status===filterStatus)
      );

      const pillarStats = PILLARS.map(p => {
        const pt  = tasks.filter(t => t.pillarId===p.id);
        const avg = pt.length ? Math.round(pt.reduce((s,t)=>s+t.progress,0)/pt.length) : 0;
        return {...p, count:pt.length, avg, resolved:pt.filter(t=>t.status==="Resolved").length};
      });

      const sortedSessions = [...sessions].sort((a,b)=>b.date.localeCompare(a.date));
      const latestSession  = sortedSessions[0];

      const btnSmall = (color, border) => ({
        border: `1px solid ${border||"#2D3748"}`, borderRadius:7, padding:"5px 11px",
        background:"none", cursor:"pointer", fontSize:12, color, fontWeight:600,
        display:"flex", alignItems:"center", gap:4,
      });

      if (!ready) return (
        <div style={{ fontFamily:"system-ui,sans-serif", minHeight:"100vh", background:"#F0F4F8", display:"flex", alignItems:"center", justifyContent:"center", flexDirection:"column", gap:14 }}>
          <div style={{ width:40, height:40, border:"4px solid #E2E8F0", borderTop:"4px solid #1D4ED8", borderRadius:"50%", animation:"spin 0.8s linear infinite" }} />
          <p style={{ color:"#64748B", fontSize:14, margin:0 }}>Loading…</p>
          <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
        </div>
      );

      return (
        <div style={{ fontFamily:"'Segoe UI','SF Pro Display',system-ui,sans-serif", background:"#F0F4F8", minHeight:"100vh" }}>

          {/* ── Top Nav ── */}
          <div style={{ background:"#0F172A", padding:"0 24px", display:"flex", alignItems:"center", justifyContent:"space-between", height:58, position:"sticky", top:0, zIndex:50 }}>
            <div style={{ display:"flex", alignItems:"center", gap:16 }}>
              <div style={{ display:"flex", flexDirection:"column" }}>
                <span style={{ fontWeight:800, fontSize:15, color:"#fff", letterSpacing:"-0.02em" }}>ANZ Cadence</span>
                <span style={{ fontSize:10, color:"#64748B", fontWeight:500, letterSpacing:"0.08em", textTransform:"uppercase" }}>Fortnightly Meeting Tracker</span>
              </div>
              <div style={{ width:1, height:32, background:"#1E293B" }} />
              <div style={{ display:"flex", gap:2 }}>
                {[["pillars","Pillars"],["sessions","Sessions"],["tracker","Action Items"]].map(([k,label]) => (
                  <button key={k} onClick={() => setActiveView(k)} style={{ border:"none", borderRadius:8, padding:"7px 14px", fontSize:13, fontWeight:600, cursor:"pointer", transition:"all 0.15s", background:activeView===k?"#1E293B":"none", color:activeView===k?"#fff":"#64748B" }}>{label}</button>
                ))}
              </div>
            </div>

            <div style={{ display:"flex", gap:8, alignItems:"center" }}>
              {latestSession && <span style={{ fontSize:12, color:"#475569" }}>Last: {new Date(latestSession.date).toLocaleDateString("en-AU",{day:"numeric",month:"short",year:"numeric"})}</span>}

              <span title="SharePoint sync status" style={{ fontSize:11, fontWeight:700, padding:"4px 10px", borderRadius:20, whiteSpace:"nowrap",
                background: syncStatus==='synced'?'#D1FAE5': syncStatus==='syncing'?'#FEF3C7': syncStatus==='error'?'#FEE2E2': syncStatus==='offline'?'#F1F5F9':'#1E293B',
                color:      syncStatus==='synced'?'#065F46': syncStatus==='syncing'?'#78350F': syncStatus==='error'?'#991B1B': syncStatus==='offline'?'#475569':'#94A3B8'
              }}>
                {syncStatus==='synced'?'✓ Synced': syncStatus==='syncing'?'↑ Syncing…': syncStatus==='error'?'⚠ Sync Error': syncStatus==='offline'?'◌ Offline':'◌ Loading'}
              </span>

              {/* Data buttons */}
              <button onClick={exportJSON} title="Export all data to JSON backup" style={{ background:"#1E293B", color:"#94A3B8", border:"1px solid #334155", borderRadius:7, padding:"6px 12px", cursor:"pointer", fontSize:12, fontWeight:600 }}>↓ JSON</button>
              <button onClick={exportCSV}  title="Export action items to CSV"     style={{ background:"#1E293B", color:"#94A3B8", border:"1px solid #334155", borderRadius:7, padding:"6px 12px", cursor:"pointer", fontSize:12, fontWeight:600 }}>↓ CSV</button>
              <button onClick={() => importRef.current.click()} title="Restore from JSON backup" style={{ background:"#1E293B", color:"#94A3B8", border:"1px solid #334155", borderRadius:7, padding:"6px 12px", cursor:"pointer", fontSize:12, fontWeight:600 }}>↑ Import</button>
              <input type="file" accept=".json" ref={importRef} onChange={handleImport} style={{ display:"none" }} />

              <div style={{ width:1, height:24, background:"#1E293B" }} />
              <button onClick={() => setTaskModal("new")} style={{ background:"#1D4ED8", color:"#fff", border:"none", borderRadius:8, padding:"7px 16px", cursor:"pointer", fontSize:13, fontWeight:700, display:"flex", alignItems:"center", gap:6 }}>
                <span style={{ fontSize:16, lineHeight:1 }}>+</span> Action Item
              </button>
              <button onClick={() => setSessionModal("new")} style={{ background:"#0F172A", color:"#94A3B8", border:"1.5px solid #1E293B", borderRadius:8, padding:"7px 14px", cursor:"pointer", fontSize:13, fontWeight:700, display:"flex", alignItems:"center", gap:6 }}>
                <span style={{ fontSize:16, lineHeight:1 }}>+</span> Session
              </button>
            </div>
          </div>

          <div style={{ maxWidth:1400, margin:"0 auto", padding:"28px 24px" }}>

            {/* ══ PILLARS VIEW ══ */}
            {activeView==="pillars" && (
              <>
                <div style={{ marginBottom:24 }}>
                  <h1 style={{ margin:"0 0 4px", fontSize:24, fontWeight:800, color:"#0F172A", letterSpacing:"-0.02em" }}>Pillar Overview</h1>
                  <p style={{ margin:0, fontSize:14, color:"#64748B" }}>Track progress across all pillars and owners from the fortnightly cadence.</p>
                </div>
                <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:14, marginBottom:28 }}>
                  {[["Total Actions",tasks.length,"#1D4ED8"],["In Progress",tasks.filter(t=>t.status==="In Progress").length,"#7C3AED"],["Resolved",tasks.filter(t=>t.status==="Resolved").length,"#059669"],["Critical",tasks.filter(t=>t.priority==="Critical").length,"#DC2626"]].map(([label,val,color]) => (
                    <div key={label} style={{ background:"#fff", borderRadius:14, padding:"16px 18px", border:"1px solid #E2E8F0", display:"flex", alignItems:"center", gap:14 }}>
                      <div style={{ width:42, height:42, borderRadius:12, background:color+"15", display:"flex", alignItems:"center", justifyContent:"center" }}>
                        <div style={{ width:14, height:14, borderRadius:3, background:color, opacity:0.8 }} />
                      </div>
                      <div>
                        <p style={{ margin:"0 0 2px", fontSize:22, fontWeight:800, color:"#0F172A" }}>{val}</p>
                        <p style={{ margin:0, fontSize:12, color:"#64748B", fontWeight:500 }}>{label}</p>
                      </div>
                    </div>
                  ))}
                </div>
                <div style={{ display:"flex", flexDirection:"column", gap:16 }}>
                  {PILLARS.map(p => {
                    const pl = pillarStats.find(x=>x.id===p.id);
                    const pt = tasks.filter(t=>t.pillarId===p.id);
                    return (
                      <div key={p.id} style={{ background:"#fff", borderRadius:16, border:"1px solid #E2E8F0", overflow:"hidden" }}>
                        <div style={{ background:p.color, padding:"16px 24px", display:"flex", alignItems:"center", justifyContent:"space-between" }}>
                          <div style={{ display:"flex", alignItems:"center", gap:14 }}>
                            <div style={{ background:"rgba(255,255,255,0.15)", borderRadius:10, padding:"6px 12px" }}>
                              <span style={{ fontWeight:900, fontSize:13, color:"#fff", letterSpacing:"0.04em", textTransform:"uppercase" }}>{p.label}</span>
                            </div>
                            {p.subtitle && <span style={{ fontSize:13, color:"rgba(255,255,255,0.75)", fontWeight:500 }}>{p.subtitle}</span>}
                          </div>
                          <div style={{ display:"flex", alignItems:"center", gap:16 }}>
                            <div style={{ display:"flex", gap:8 }}>
                              {p.owners.map(o => (
                                <div key={o} style={{ display:"flex", alignItems:"center", gap:6, background:"rgba(255,255,255,0.15)", borderRadius:20, padding:"4px 10px 4px 4px" }}>
                                  <Avatar name={o} size={20} />
                                  <span style={{ fontSize:12, color:"#fff", fontWeight:600 }}>{o.split(" ")[0]}</span>
                                </div>
                              ))}
                            </div>
                            <div style={{ textAlign:"right" }}>
                              <p style={{ margin:0, fontSize:22, fontWeight:900, color:"#fff" }}>{pl.avg}%</p>
                              <p style={{ margin:0, fontSize:11, color:"rgba(255,255,255,0.7)" }}>{pl.resolved}/{pl.count} resolved</p>
                            </div>
                          </div>
                        </div>
                        <div style={{ height:4, background:"#F1F5F9" }}>
                          <div style={{ width:`${pl.avg}%`, height:"100%", background:p.color, transition:"width 0.5s" }} />
                        </div>
                        <div style={{ padding:"16px 24px" }}>
                          {pt.length===0 ? (
                            <div style={{ textAlign:"center", padding:"20px 0", color:"#CBD5E1", fontSize:13 }}>
                              No action items yet.{" "}<span style={{ color:p.color, cursor:"pointer", fontWeight:600 }} onClick={() => setTaskModal("new")}>Add one +</span>
                            </div>
                          ) : (
                            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(300px,1fr))", gap:12 }}>
                              {pt.map(t => (
                                <div key={t.id} style={{ background:"#F8FAFC", borderRadius:12, padding:"14px 16px", border:"1px solid #E2E8F0", borderLeft:`3px solid ${p.color}`, cursor:"pointer" }} onClick={() => setTaskModal(t)}>
                                  <div style={{ display:"flex", justifyContent:"space-between", gap:8, marginBottom:8 }}>
                                    <p style={{ margin:0, fontWeight:700, fontSize:13, color:"#0F172A", lineHeight:1.4, flex:1 }}>{t.title||"Untitled"}</p>
                                    <Chip label={t.priority} meta={PRIORITY_META[t.priority]} />
                                  </div>
                                  <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:10 }}>
                                    <Chip label={t.status} meta={STATUS_META[t.status]} />
                                    {t.dueDate && <span style={{ fontSize:11, color:"#94A3B8" }}>Due {t.dueDate}</span>}
                                  </div>
                                  <ProgressBar value={t.progress} color={p.color} />
                                  <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginTop:10 }}>
                                    <div style={{ display:"flex", alignItems:"center", gap:6 }}>
                                      <Avatar name={t.owner} size={22} />
                                      <span style={{ fontSize:12, color:"#64748B" }}>{t.owner}</span>
                                    </div>
                                    <span style={{ fontSize:12, fontWeight:700, color:p.color }}>{t.progress}%</span>
                                  </div>
                                </div>
                              ))}
                            </div>
                          )}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </>
            )}

            {/* ══ SESSIONS VIEW ══ */}
            {activeView==="sessions" && (
              <>
                <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:24 }}>
                  <div>
                    <h1 style={{ margin:"0 0 4px", fontSize:24, fontWeight:800, color:"#0F172A", letterSpacing:"-0.02em" }}>Session Summaries</h1>
                    <p style={{ margin:0, fontSize:14, color:"#64748B" }}>Key points, decisions and actions from each fortnightly meeting.</p>
                  </div>
                  <button onClick={() => setSessionModal("new")} style={{ background:"#1D4ED8", color:"#fff", border:"none", borderRadius:10, padding:"10px 18px", cursor:"pointer", fontSize:13, fontWeight:700, display:"flex", alignItems:"center", gap:6 }}>
                    <span style={{ fontSize:16 }}>+</span> New Session
                  </button>
                </div>
                {sortedSessions.length===0 ? (
                  <div style={{ background:"#fff", borderRadius:16, padding:"60px 32px", textAlign:"center", border:"1px solid #E2E8F0" }}>
                    <p style={{ color:"#94A3B8", fontSize:15, margin:"0 0 16px" }}>No sessions recorded yet.</p>
                    <button onClick={() => setSessionModal("new")} style={{ background:"#1D4ED8", color:"#fff", border:"none", borderRadius:10, padding:"10px 20px", cursor:"pointer", fontWeight:700, fontSize:14 }}>Create First Session</button>
                  </div>
                ) : (
                  <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
                    {sortedSessions.map((s,idx) => {
                      const isExpanded = expandedSession===s.id;
                      const dateStr    = new Date(s.date).toLocaleDateString("en-AU",{weekday:"long",day:"numeric",month:"long",year:"numeric"});
                      return (
                        <div key={s.id} style={{ background:"#fff", borderRadius:16, border:"1px solid #E2E8F0", overflow:"hidden" }}>
                          <div style={{ padding:"18px 24px", display:"flex", alignItems:"center", justifyContent:"space-between", cursor:"pointer", borderBottom:isExpanded?"1px solid #F1F5F9":"none" }} onClick={() => setExpandedSession(isExpanded?null:s.id)}>
                            <div style={{ display:"flex", alignItems:"center", gap:14 }}>
                              <div style={{ background:idx===0?"#DBEAFE":"#F1F5F9", borderRadius:10, padding:"8px 14px", textAlign:"center", minWidth:56 }}>
                                <p style={{ margin:0, fontSize:18, fontWeight:900, color:idx===0?"#1D4ED8":"#64748B" }}>{new Date(s.date).getDate()}</p>
                                <p style={{ margin:0, fontSize:10, fontWeight:700, color:idx===0?"#3B82F6":"#94A3B8", textTransform:"uppercase", letterSpacing:"0.06em" }}>{new Date(s.date).toLocaleDateString("en-AU",{month:"short"})}</p>
                              </div>
                              <div>
                                <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                                  <p style={{ margin:"0 0 2px", fontWeight:800, fontSize:15, color:"#0F172A" }}>{s.title}</p>
                                  {idx===0 && <span style={{ fontSize:10, fontWeight:700, background:"#DBEAFE", color:"#1D4ED8", padding:"2px 8px", borderRadius:20, textTransform:"uppercase", letterSpacing:"0.06em" }}>Latest</span>}
                                </div>
                                <p style={{ margin:0, fontSize:13, color:"#64748B" }}>{dateStr}</p>
                              </div>
                            </div>
                            <div style={{ display:"flex", alignItems:"center", gap:14 }}>
                              <div style={{ display:"flex", gap:8, fontSize:12, color:"#94A3B8" }}>
                                <span>{(s.keyPoints||[]).filter(k=>k).length} points</span>
                                <span>{(s.decisions||[]).filter(k=>k).length} decisions</span>
                              </div>
                              <div style={{ display:"flex", gap:6 }}>
                                <button onClick={e => { e.stopPropagation(); setSessionModal(s); }} style={{ border:"1px solid #E2E8F0", borderRadius:8, padding:"6px 12px", background:"#fff", cursor:"pointer", fontSize:12, color:"#475569", fontWeight:600 }}>Edit</button>
                                <button onClick={e => { e.stopPropagation(); setDeleteSessionId(s.id); }} style={{ border:"1px solid #FEE2E2", borderRadius:8, padding:"6px 12px", background:"#fff", cursor:"pointer", fontSize:12, color:"#EF4444", fontWeight:600 }}>Delete</button>
                              </div>
                              <span style={{ fontSize:18, color:"#94A3B8" }}>{isExpanded?"▲":"▼"}</span>
                            </div>
                          </div>
                          {isExpanded && (
                            <div style={{ padding:"24px", display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:20 }}>
                              <div style={{ gridColumn:"1 / -1", background:"#F8FAFC", borderRadius:12, padding:"16px 18px", borderLeft:"3px solid #1D4ED8" }}>
                                <p style={{ margin:"0 0 6px", fontSize:11, fontWeight:800, color:"#1D4ED8", textTransform:"uppercase", letterSpacing:"0.06em" }}>Executive Summary</p>
                                <p style={{ margin:0, fontSize:14, color:"#334155", lineHeight:1.6 }}>{s.summary||<em style={{ color:"#CBD5E1" }}>No summary recorded.</em>}</p>
                              </div>
                              {[["Key Discussion Points","keyPoints"],["Decisions Made","decisions"],["Action Items Raised","actionItems"]].map(([heading,field]) => (
                                <div key={field}>
                                  <p style={{ margin:"0 0 10px", fontSize:12, fontWeight:700, color:"#475569" }}>{heading}</p>
                                  <ul style={{ margin:0, padding:"0 0 0 18px" }}>
                                    {(s[field]||[]).filter(k=>k).map((pt,i) => <li key={i} style={{ fontSize:13, color:"#334155", marginBottom:6, lineHeight:1.5 }}>{pt}</li>)}
                                    {!(s[field]||[]).filter(k=>k).length && <li style={{ fontSize:13, color:"#CBD5E1" }}>None recorded.</li>}
                                  </ul>
                                </div>
                              ))}
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}
              </>
            )}

            {/* ══ ACTION ITEMS TRACKER ══ */}
            {activeView==="tracker" && (
              <>
                <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:20 }}>
                  <div>
                    <h1 style={{ margin:"0 0 4px", fontSize:24, fontWeight:800, color:"#0F172A", letterSpacing:"-0.02em" }}>Action Items Tracker</h1>
                    <p style={{ margin:0, fontSize:14, color:"#64748B" }}>All action items across pillars and owners.</p>
                  </div>
                  <button onClick={() => setTaskModal("new")} style={{ background:"#1D4ED8", color:"#fff", border:"none", borderRadius:10, padding:"10px 18px", cursor:"pointer", fontSize:13, fontWeight:700, display:"flex", alignItems:"center", gap:6 }}>
                    <span style={{ fontSize:16 }}>+</span> New Action Item
                  </button>
                </div>
                <div style={{ display:"flex", gap:10, marginBottom:18, flexWrap:"wrap" }}>
                  {[
                    [filterPillar, setFilterPillar, [["all","All Pillars"],...PILLARS.map(p=>[p.id,p.label])]],
                    [filterOwner,  setFilterOwner,  [["all","All Owners"],...ALL_OWNERS.map(o=>[o,o])]],
                    [filterStatus, setFilterStatus, [["all","All Statuses"],...STATUSES.map(s=>[s,s])]],
                  ].map(([val,setter,opts],i) => (
                    <select key={i} value={val} onChange={e => setter(e.target.value)} style={{ border:"1.5px solid #E2E8F0", borderRadius:8, padding:"8px 12px", fontSize:13, color:"#475569", background:"#fff", cursor:"pointer" }}>
                      {opts.map(([v,l]) => <option key={v} value={v}>{l}</option>)}
                    </select>
                  ))}
                </div>
                <div style={{ background:"#fff", borderRadius:16, border:"1px solid #E2E8F0", overflow:"auto" }}>
                  <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13, minWidth:900 }}>
                    <thead>
                      <tr style={{ borderBottom:"2px solid #F1F5F9" }}>
                        {["Pillar","Action Item","Owner","Priority","Status","Progress","Due",""].map(h => (
                          <th key={h} style={{ padding:"12px 16px", textAlign:"left", fontWeight:700, color:"#475569", fontSize:11, textTransform:"uppercase", letterSpacing:"0.05em", whiteSpace:"nowrap" }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {filteredTasks.map((t,i) => {
                        const p = PILLARS.find(x=>x.id===t.pillarId);
                        return (
                          <tr key={t.id} style={{ borderBottom:"1px solid #F8FAFC", background:i%2?"#FAFBFC":"#fff" }}>
                            <td style={{ padding:"11px 16px" }}><span style={{ background:p.light, color:p.dark, fontWeight:700, fontSize:11, padding:"3px 9px", borderRadius:6, whiteSpace:"nowrap" }}>{p.label}</span></td>
                            <td style={{ padding:"11px 16px", fontWeight:600, color:"#0F172A", maxWidth:260 }}>
                              <p style={{ margin:0, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{t.title||"—"}</p>
                              {t.notes && <p style={{ margin:"2px 0 0", fontSize:11, color:"#94A3B8", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{t.notes}</p>}
                            </td>
                            <td style={{ padding:"11px 16px" }}>
                              <div style={{ display:"flex", alignItems:"center", gap:7 }}>
                                <Avatar name={t.owner} size={24} />
                                <span style={{ color:"#334155", fontWeight:500, whiteSpace:"nowrap" }}>{t.owner}</span>
                              </div>
                            </td>
                            <td style={{ padding:"11px 16px" }}><Chip label={t.priority} meta={PRIORITY_META[t.priority]} /></td>
                            <td style={{ padding:"11px 16px" }}><Chip label={t.status} meta={STATUS_META[t.status]} /></td>
                            <td style={{ padding:"11px 16px", minWidth:130 }}>
                              <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                                <div style={{ flex:1 }}><ProgressBar value={t.progress} color={p.color} /></div>
                                <span style={{ fontSize:12, fontWeight:700, color:p.color, minWidth:30 }}>{t.progress}%</span>
                              </div>
                            </td>
                            <td style={{ padding:"11px 16px", color:"#94A3B8", whiteSpace:"nowrap" }}>{t.dueDate||"—"}</td>
                            <td style={{ padding:"11px 16px" }}>
                              <div style={{ display:"flex", gap:6 }}>
                                <button onClick={() => setTaskModal(t)} style={{ border:"1px solid #E2E8F0", borderRadius:7, padding:"5px 11px", background:"#fff", cursor:"pointer", fontSize:12, color:"#1D4ED8", fontWeight:600 }}>Edit</button>
                                <button onClick={() => setDeleteTaskId(t.id)} style={{ border:"1px solid #FEE2E2", borderRadius:7, padding:"5px 11px", background:"#fff", cursor:"pointer", fontSize:12, color:"#EF4444", fontWeight:600 }}>Del</button>
                              </div>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                  {filteredTasks.length===0 && (
                    <div style={{ padding:"50px 32px", textAlign:"center" }}>
                      <p style={{ color:"#CBD5E1", fontSize:14, margin:"0 0 14px" }}>No action items match your filters.</p>
                      <button onClick={() => setTaskModal("new")} style={{ background:"#1D4ED8", color:"#fff", border:"none", borderRadius:9, padding:"9px 20px", cursor:"pointer", fontWeight:700, fontSize:13 }}>Add First Action Item</button>
                    </div>
                  )}
                </div>
                {tasks.length>0 && (
                  <div style={{ marginTop:24 }}>
                    <h2 style={{ margin:"0 0 14px", fontSize:16, fontWeight:800, color:"#0F172A" }}>Owner Summary</h2>
                    <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))", gap:12 }}>
                      {ALL_OWNERS.map(owner => {
                        const ot = tasks.filter(t=>t.owner===owner);
                        if (!ot.length) return null;
                        return (
                          <div key={owner} style={{ background:"#fff", borderRadius:12, padding:"14px 16px", border:"1px solid #E2E8F0" }}>
                            <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:10 }}>
                              <Avatar name={owner} size={32} />
                              <div>
                                <p style={{ margin:0, fontSize:13, fontWeight:700, color:"#0F172A" }}>{owner}</p>
                                <p style={{ margin:0, fontSize:11, color:"#94A3B8" }}>{ot.length} item{ot.length!==1?"s":""}</p>
                              </div>
                            </div>
                            <div style={{ display:"flex", gap:8 }}>
                              <span style={{ fontSize:11, background:"#D1FAE5", color:"#065F46", padding:"2px 7px", borderRadius:20, fontWeight:600 }}>{ot.filter(t=>t.status==="Resolved").length} resolved</span>
                              {ot.filter(t=>t.priority==="Critical").length>0 && <span style={{ fontSize:11, background:"#FEE2E2", color:"#991B1B", padding:"2px 7px", borderRadius:20, fontWeight:600 }}>{ot.filter(t=>t.priority==="Critical").length} critical</span>}
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                )}
              </>
            )}
          </div>

          {taskModal    && <TaskModal    task={taskModal==="new"?null:taskModal}       onSave={saveTask}    onClose={() => setTaskModal(null)} />}
          {sessionModal && <SessionModal session={sessionModal==="new"?null:sessionModal} nextId={nextSessionId} onSave={saveSession} onClose={() => setSessionModal(null)} />}

          {deleteTaskId && (
            <div style={{ position:"fixed", inset:0, background:"rgba(15,23,42,0.6)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:1000 }}>
              <div style={{ background:"#fff", borderRadius:14, padding:"28px 32px", maxWidth:360, textAlign:"center" }}>
                <p style={{ margin:"0 0 6px", fontSize:16, fontWeight:700, color:"#0F172A" }}>Delete this action item?</p>
                <p style={{ margin:"0 0 20px", fontSize:13, color:"#64748B" }}>This cannot be undone.</p>
                <div style={{ display:"flex", gap:10, justifyContent:"center" }}>
                  <button onClick={() => deleteTask(deleteTaskId)} style={{ background:"#EF4444", color:"#fff", border:"none", borderRadius:9, padding:"10px 24px", fontWeight:700, cursor:"pointer" }}>Delete</button>
                  <button onClick={() => setDeleteTaskId(null)} style={{ border:"1.5px solid #E2E8F0", borderRadius:9, padding:"10px 20px", background:"#fff", cursor:"pointer", color:"#475569" }}>Cancel</button>
                </div>
              </div>
            </div>
          )}
          {deleteSessionId && (
            <div style={{ position:"fixed", inset:0, background:"rgba(15,23,42,0.6)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:1000 }}>
              <div style={{ background:"#fff", borderRadius:14, padding:"28px 32px", maxWidth:360, textAlign:"center" }}>
                <p style={{ margin:"0 0 6px", fontSize:16, fontWeight:700, color:"#0F172A" }}>Delete this session?</p>
                <p style={{ margin:"0 0 20px", fontSize:13, color:"#64748B" }}>All session notes will be lost.</p>
                <div style={{ display:"flex", gap:10, justifyContent:"center" }}>
                  <button onClick={() => deleteSession(deleteSessionId)} style={{ background:"#EF4444", color:"#fff", border:"none", borderRadius:9, padding:"10px 24px", fontWeight:700, cursor:"pointer" }}>Delete</button>
                  <button onClick={() => setDeleteSessionId(null)} style={{ border:"1.5px solid #E2E8F0", borderRadius:9, padding:"10px 20px", background:"#fff", cursor:"pointer", color:"#475569" }}>Cancel</button>
                </div>
              </div>
            </div>
          )}
        </div>
      );
    }

    ReactDOM.createRoot(document.getElementById("root")).render(<ANZDashboard />);
  </script>
</body>
</html>
