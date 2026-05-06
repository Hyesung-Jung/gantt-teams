import { useState, useRef, useEffect, useMemo, useCallback } from "react";

const APP_VERSION = "1.3.0";

const SUPABASE_URL = "https://yueibaamyinfdiiylzyt.supabase.co";
const SUPABASE_KEY = "sb_publishable_PG4PDvjODQqIPF7xYaQpFw_IXOiN5cw";

async function sbGet(channelId) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/gantt_data?id=eq.${channelId}&select=data`, {
    headers: { apikey: SUPABASE_KEY, Authorization: `Bearer ${SUPABASE_KEY}` }
  });
  const rows = await res.json();
  return rows?.[0]?.data ?? null;
}

async function sbSave(channelId, data) {
  const body = JSON.stringify({ id: channelId, data, updated_at: new Date().toISOString() });
  const headers = {
    apikey: SUPABASE_KEY,
    Authorization: `Bearer ${SUPABASE_KEY}`,
    "Content-Type": "application/json",
    // Ask PostgREST to return updated rows so we can detect 0-row updates
    // (otherwise PATCH on a non-existent row returns 200 OK with no body
    //  and our code silently treats it as success — leading to data loss).
    Prefer: "return=representation",
  };
  const patch = await fetch(`${SUPABASE_URL}/rest/v1/gantt_data?id=eq.${channelId}`, {
    method: "PATCH", headers, body,
  });
  if (patch.ok) {
    const updated = await patch.json().catch(() => []);
    if (Array.isArray(updated) && updated.length > 0) return; // success
    // PATCH succeeded HTTP-wise but didn't update any row — row doesn't exist.
    // Fall through to POST below.
  }
  const post = await fetch(`${SUPABASE_URL}/rest/v1/gantt_data`, {
    method: "POST", headers, body,
  });
  if (!post.ok) {
    const err = await post.text();
    throw new Error("Save failed: " + post.status + " " + err);
  }
}

// Cached channel id from previous session — used to short-circuit Teams getContext()
// so the first paint isn't blocked on a 200–500 ms SDK round-trip.
const LAST_CH_KEY = "gantt_lastch";
function getCachedChannelId() {
  try { return localStorage.getItem(LAST_CH_KEY) || null; } catch(e) { return null; }
}
function setCachedChannelId(id) {
  try { if (id && id !== "default") localStorage.setItem(LAST_CH_KEY, id); } catch(e) {}
}

// Open Teams' built-in people picker (the same UI used for @mentions in chat).
// Resolves to an array of { objectId, displayName, email }.
// Falls back gracefully when not running inside Teams or when the SDK lacks
// the people capability — caller can then prompt for manual entry.
function pickTeamsPeople(opts = {}) {
  if (!window.microsoftTeams?.people?.selectPeople) {
    return Promise.reject(new Error("UNAVAILABLE"));
  }
  const options = {
    title: opts.title || "담당자 선택",
    setSelected: opts.setSelected || [],
    openOrgWideSearchInChatOrChannel: true,
    singleSelect: opts.singleSelect ?? false,
  };
  // teams-js v2+ returns a Promise; v1.x uses the (callback, options) signature.
  // Try the modern call first and adapt if it returned a non-thenable.
  try {
    const r = window.microsoftTeams.people.selectPeople(options);
    if (r && typeof r.then === "function") {
      return r.then(arr => Array.isArray(arr) ? arr : []);
    }
  } catch(e) {
    // Fall through to v1 callback style
  }
  return new Promise((resolve, reject) => {
    try {
      window.microsoftTeams.people.selectPeople((err, people) => {
        if (err) reject(err); else resolve(Array.isArray(people) ? people : []);
      }, options);
    } catch(e) { reject(e); }
  });
}

function getChannelId() {
  // Priority 1: URL ?ch= parameter — used by channel-tab installations
  // (config.html sets this so each channel gets its own ID).
  try {
    const params = new URLSearchParams(window.location.search);
    const ch = params.get("ch");
    if (ch && ch !== "default") return Promise.resolve(ch);
  } catch(e) {}
  // Priority 2: Teams context (works for personal apps)
  try {
    if (window.microsoftTeams) {
      return new Promise(resolve => {
        let resolved = false;
        const safeResolve = (v) => { if (!resolved) { resolved = true; resolve(v); } };
        // Hard timeout reduced from 5 s to 1.5 s — the long fallback was the
        // biggest contributor to the perceived "tab is stuck loading" delay
        // when running outside Teams or when SDK init stalls.
        setTimeout(() => safeResolve("default"), 1500);
        try {
          window.microsoftTeams.app.initialize().then(() => {
            window.microsoftTeams.app.getContext().then(ctx => {
              safeResolve(ctx.channel?.id || ctx.chat?.id || "default");
            }).catch(() => safeResolve("default"));
          }).catch(() => safeResolve("default"));
        } catch(e) { safeResolve("default"); }
      });
    }
  } catch(e) {}
  return Promise.resolve("default");
}

const Icon = ({ d, size=14, stroke="#94a3b8", fill="none", sw=1.8, ...rest }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill={fill} stroke={stroke}
    strokeWidth={sw} strokeLinecap="round" strokeLinejoin="round" style={{ flexShrink:0 }} {...rest}>
    {Array.isArray(d) ? d.map((p,i)=><path key={i} d={p}/>) : <path d={d}/>}
  </svg>
);

const EditIcon = ({ size=14, color="#64748b" }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" style={{flexShrink:0}}>
    <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/>
    <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>
  </svg>
);
const TrashIcon = ({ size=14, color="#f87171" }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" style={{flexShrink:0}}>
    <polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/>
    <path d="M10 11v6M14 11v6"/><path d="M9 6V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2"/>
  </svg>
);
const PlusIcon = ({ size=14, color="#94a3b8" }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round" style={{flexShrink:0}}>
    <line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/>
  </svg>
);
const LinkIcon = ({ size=13, color="#6366f1" }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{flexShrink:0}}>
    <path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/>
    <path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/>
  </svg>
);
const ChevronDown = ({ size=12 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="#475569" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{flexShrink:0}}>
    <polyline points="6 9 12 15 18 9"/>
  </svg>
);
const ChevronRight = ({ size=12 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="#475569" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{flexShrink:0}}>
    <polyline points="9 6 15 12 9 18"/>
  </svg>
);
const ExcelIcon = ({ size=14 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" style={{flexShrink:0}}>
    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
    <polyline points="14 2 14 8 20 8"/>
    <line x1="12" y1="11" x2="12" y2="17"/>
    <polyline points="9 14 12 17 15 14"/>
  </svg>
);
const MilestoneIcon = ({ size=13, color="#818cf8" }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill={color} stroke={color} strokeWidth="1" style={{flexShrink:0}}>
    <polygon points="12,2 22,12 12,22 2,12"/>
  </svg>
);
const MenuIcon = ({ size=18 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{flexShrink:0}}>
    <line x1="3" y1="6" x2="21" y2="6"/><line x1="3" y1="12" x2="21" y2="12"/><line x1="3" y1="18" x2="21" y2="18"/>
  </svg>
);
const CloseIcon = ({ size=16 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{flexShrink:0}}>
    <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
  </svg>
);
const ListIcon = ({ size=14 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{flexShrink:0}}>
    <line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/>
    <line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/>
  </svg>
);
const GanttIcon = ({ size=14 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{flexShrink:0}}>
    <rect x="3" y="4" width="10" height="3" rx="1"/>
    <rect x="7" y="10" width="14" height="3" rx="1"/>
    <rect x="5" y="16" width="8" height="3" rx="1"/>
  </svg>
);
const KanbanIcon = ({ size=14 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{flexShrink:0}}>
    <rect x="3" y="3" width="5" height="14" rx="1"/>
    <rect x="10" y="3" width="5" height="9" rx="1"/>
    <rect x="17" y="3" width="4" height="11" rx="1"/>
  </svg>
);

const SECTOR_COLORS = ["#6366f1","#f43f5e","#10b981","#f59e0b","#3b82f6","#8b5cf6","#06b6d4","#ec4899","#84cc16","#14b8a6"];

function hexToRgb(hex) {
  try {
    const h = (hex||"#6366f1").replace("#","");
    return { r:parseInt(h.slice(0,2),16)||0, g:parseInt(h.slice(2,4),16)||0, b:parseInt(h.slice(4,6),16)||0 };
  } catch(e) { return {r:99,g:102,b:241}; }
}
function lighten(hex, amount) {
  try {
    const {r,g,b} = hexToRgb(hex||"#6366f1");
    const m = c => Math.min(255, Math.max(0, Math.round(c+(255-c)*amount)));
    const x = c => c.toString(16).padStart(2,"0");
    return `#${x(m(r))}${x(m(g))}${x(m(b))}`;
  } catch(e) { return "#6366f1"; }
}
function hexAlpha(hex, alpha) {
  try {
    if (!hex || typeof hex !== "string" || !hex.startsWith("#")) return "#6366f1";
    return hex + Math.round(alpha*255).toString(16).padStart(2,"0");
  } catch(e) { return "#6366f1"; }
}
function childColor(sectorColor, index, total) {
  try {
    const t = total<=1 ? 0.28 : 0.14+(index/(Math.max(total-1,1)))*0.42;
    return lighten(sectorColor||"#6366f1", t);
  } catch(e) { return "#818cf8"; }
}
function effectiveColor(item, allItems) {
  try {
    if (!item) return "#6366f1";
    if (item.type==="sector") return item.color || "#6366f1";
    const sector = (allItems||[]).find(i=>i.id===item.parentId&&i.type==="sector");
    if (!sector) return "#6366f1";
    const siblings = allItems.filter(i=>i.parentId===sector.id&&i.type!=="sector");
    const idx = siblings.findIndex(i=>i.id===item.id);
    return childColor(sector.color || "#6366f1", Math.max(idx,0), siblings.length);
  } catch(e) { return "#6366f1"; }
}

const ROW_HEIGHT = 32;
const ROW_HEIGHT_MOBILE = 44;
const DETAIL_HEIGHT = 110;
const DETAIL_HEIGHT_COMPACT = 80;
const DETAIL_HEIGHT_MOBILE = 110;

const THEMES = {
  dark: {
    bg: "#0f1117",
    bgDeep: "#0b0e17",
    bgCard: "#161b27",
    bgRow: "#131820",
    bgRowHover: "#1a1f2e",
    bgHeader: "#0d111a",
    bgInput: "#0f1117",
    bgSelected: "#0f1b30",
    border: "#1e2535",
    borderLight: "#1a2030",
    text: "#e2e8f0",
    textSub: "#94a3b8",
    textMuted: "#64748b",
    textFaint: "#475569",
    textDim: "#334155",
  },
  light: {
    bg: "#ffffff",
    bgDeep: "#f8f9fb",
    bgCard: "#ffffff",
    bgRow: "#f5f6f8",
    bgRowHover: "#eef1f6",
    bgHeader: "#f0f1f5",
    bgInput: "#ffffff",
    bgSelected: "#e8f0fe",
    border: "#e0e3eb",
    borderLight: "#eceef3",
    text: "#1a2030",
    textSub: "#3a4055",
    textMuted: "#5c6378",
    textFaint: "#8891a5",
    textDim: "#b0b7c8",
  },
};

let T = THEMES.dark;
function setGlobalTheme(key) { T = THEMES[key]; }

const COL_NAME = "1fr";
const COL_START = "60px";
const COL_END = "60px";
const COL_PROG = "54px";
const GRID_COLS = `26px 6px ${COL_NAME} ${COL_START} ${COL_END} ${COL_PROG}`;
// Mobile-friendly grid: bigger drag handle, no separate dot col, larger touch targets
const GRID_COLS_MOBILE = `34px ${COL_NAME} 54px 54px 44px`;

const ZOOM_PRESETS = [
  { key:"fit", label:"전체", dw:0 },
  { key:"year", label:"연", dw:2 },
  { key:"month", label:"월", dw:10 },
  { key:"day", label:"일", dw:28 },
];

const today = new Date(); today.setHours(0,0,0,0);
const addDays = (d,n) => { const x=new Date(d); x.setDate(x.getDate()+n); return x; };
const dateDiff = (a,b) => Math.round((b-a)/86400000);
const toISO = d => d.toISOString().slice(0,10);
const parseDate = s => { if (!s || !s.trim()) return null; const d=new Date(s); if (isNaN(d.getTime())) return null; d.setHours(0,0,0,0); return d; };

let _uid=20;
const uid=()=>++_uid;
// Bumps the uid counter so new ids never collide with existing ones from
// loaded data. Must be called whenever items state is replaced via deserialize.
const syncUid = (items) => {
  if (!Array.isArray(items)) return;
  const max = Math.max(_uid, ...items.map(i => (typeof i?.id === "number" ? i.id : 0)));
  if (max > _uid) _uid = max;
};

function sortedByDate(raw) {
  const categories = raw.filter(i=>i.type==="category").slice().sort((a,b)=>(a.order??0)-(b.order??0));
  const sectors = raw.filter(i=>i.type==="sector").map((s,si)=>{
    const kids = raw.filter(c=>c.parentId===s.id&&c.start);
    const min = kids.length ? Math.min(...kids.map(c=>c.start.getTime())) : si*1e12;
    return {...s, _o:min};
  }).sort((a,b)=>a._o-b._o).map((s,i)=>{ const {_o,...r}=s; return {...r,order:i}; });
  const children = sectors.flatMap(sector=>{
    return raw.filter(i=>i.parentId===sector.id)
      .slice().sort((a,b)=>(a.start?.getTime()||0)-(b.start?.getTime()||0))
      .map((k,i)=>({...k,order:i}));
  });
  return [...categories, ...sectors, ...children];
}

const initialItems = sortedByDate([
  { id:100, type:"category", name:"기본", collapsed:false, order:0 },
  { id:1, type:"sector", parentId:100, name:"기획", color:SECTOR_COLORS[0], collapsed:false },
  { id:2, type:"task", parentId:1, name:"요구사항 분석", start:addDays(today,-5), end:addDays(today,2), progress:70 },
  { id:3, type:"task", parentId:1, name:"기술 스택 결정", start:addDays(today,1), end:addDays(today,5), progress:0 },
  { id:4, type:"milestone", parentId:1, name:"기획 완료", start:addDays(today,4), end:addDays(today,4), progress:0 },
  { id:5, type:"sector", parentId:100, name:"개발", color:SECTOR_COLORS[5], collapsed:false },
  { id:6, type:"task", parentId:5, name:"UI/UX 디자인", start:addDays(today,5), end:addDays(today,12), progress:0 },
  { id:7, type:"task", parentId:5, name:"프론트엔드 개발", start:addDays(today,10), end:addDays(today,22), progress:0 },
  { id:8, type:"task", parentId:5, name:"백엔드 API 개발", start:addDays(today,5), end:addDays(today,20), progress:0 },
  { id:9, type:"sector", parentId:100, name:"출시", color:SECTOR_COLORS[3], collapsed:false },
  { id:10, type:"task", parentId:9, name:"QA 테스트", start:addDays(today,20), end:addDays(today,27), progress:0 },
  { id:11, type:"milestone", parentId:9, name:"제품 런칭", start:addDays(today,28), end:addDays(today,28), progress:0 },
]);

function DepArrows({ flatRows, viewStart, DAY_WIDTH, getRowTop, rowH }) {
  const arrows = [];
  flatRows.forEach((row, toIdx) => {
    (row.item.deps || []).forEach(depId => {
      const fromIdx = flatRows.findIndex(r => r.item.id === depId);
      if (fromIdx < 0) return;
      const from = flatRows[fromIdx].item;
      if (!from.end || !row.item.start) return;
      const x1 = (dateDiff(viewStart, from.end) + 1) * DAY_WIDTH;
      const y1 = (getRowTop ? getRowTop(fromIdx) : fromIdx * rowH) + rowH / 2;
      const x2 = dateDiff(viewStart, row.item.start) * DAY_WIDTH;
      const y2 = (getRowTop ? getRowTop(toIdx) : toIdx * rowH) + rowH / 2;
      arrows.push({ x1, y1, x2, y2, key: `${depId}-${row.item.id}` });
    });
  });
  return (
    <svg style={{position:"absolute",inset:0,pointerEvents:"none",zIndex:8,overflow:"visible"}}>
      <defs>
        <marker id="arr" markerWidth="6" markerHeight="6" refX="5" refY="3" orient="auto">
          <path d="M0,0 L0,6 L6,3 z" fill="#818cf8" opacity="0.9"/>
        </marker>
      </defs>
      {arrows.map(a => (
        <path key={a.key}
          d={`M${a.x1},${a.y1} C${a.x1+28},${a.y1} ${a.x2-28},${a.y2} ${a.x2},${a.y2}`}
          fill="none" stroke="#818cf8" strokeWidth="1.5" strokeDasharray="5 3"
          opacity="0.75" markerEnd="url(#arr)"/>
      ))}
    </svg>
  );
}

const GripIcon = ({active}) => (
  <svg width="12" height="12" viewBox="0 0 14 14" fill="none">
    {[2,6,10].map(x=>[2,6,10].map(y=>(
      <circle key={x+'-'+y} cx={x} cy={y} r="1.3" fill={active?"#818cf8":"#475569"}/>
    )))}
  </svg>
);

const DragBtn = ({active, onActivate, onMoveUp, onMoveDown, canUp, canDown, onMouseDownDrag, isMobile}) => (
  <div style={{display:"flex", alignItems:"center", gap:1, flexShrink:0}}>
    <button
      data-dragbtn="1"
      onClick={e=>{ e.stopPropagation(); e.preventDefault(); onActivate(); }}
      onMouseDown={e=>{ if(e.button===0) onMouseDownDrag(e); }}
      onTouchStart={e=>{ onMouseDownDrag(e); }}
      title="드래그하여 이동 / 클릭하여 이동 모드"
      style={{
        width: isMobile?28:22, height:isMobile?28:22, flexShrink:0,
        display:"flex", alignItems:"center", justifyContent:"center",
        background: active ? "rgba(99,102,241,0.25)" : "transparent",
        border: active ? "1.5px solid #6366f1" : "1px solid transparent",
        borderRadius:6, cursor:"grab", touchAction:"none",
        transition:"all 0.15s",
      }}>
      <GripIcon active={active}/>
    </button>
    {active && (
      <div style={{display:"flex", flexDirection:"column", gap:1}}>
        <button data-dragbtn="1"
          onClick={e=>{ e.stopPropagation(); e.preventDefault(); onMoveUp(); }}
          disabled={!canUp}
          style={{width:isMobile?26:22,height:isMobile?15:13,display:"flex",alignItems:"center",justifyContent:"center",
            background:canUp?"rgba(99,102,241,0.3)":"rgba(255,255,255,0.04)",
            border:"none",borderRadius:"3px 3px 0 0",cursor:canUp?"pointer":"default",padding:0}}>
          <svg width="9" height="6" viewBox="0 0 9 6"><path d="M4.5 0L9 6H0z" fill={canUp?"#818cf8":"#475569"}/></svg>
        </button>
        <button data-dragbtn="1"
          onClick={e=>{ e.stopPropagation(); e.preventDefault(); onMoveDown(); }}
          disabled={!canDown}
          style={{width:isMobile?26:22,height:isMobile?15:13,display:"flex",alignItems:"center",justifyContent:"center",
            background:canDown?"rgba(99,102,241,0.3)":"rgba(255,255,255,0.04)",
            border:"none",borderRadius:"0 0 3px 3px",cursor:canDown?"pointer":"default",padding:0}}>
          <svg width="9" height="6" viewBox="0 0 9 6"><path d="M4.5 6L0 0H9z" fill={canDown?"#818cf8":"#475569"}/></svg>
        </button>
      </div>
    )}
  </div>
);

function Ib({onClick,title,children,sz=28}) {
  return (
    <button title={title} onClick={onClick}
      style={{width:sz,height:sz,background:"transparent",border:"none",cursor:"pointer",borderRadius:6,display:"flex",alignItems:"center",justifyContent:"center",color:"#94a3b8",transition:"background 0.15s"}}
      onMouseEnter={e=>e.currentTarget.style.background="rgba(255,255,255,0.07)"}
      onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
      {children}
    </button>
  );
}

function F({label,children}) {
  return <label style={{display:"flex",flexDirection:"column",gap:5,fontSize:11,color:"#64748b",fontWeight:600}}>{label}{children}</label>;
}

// ─── LeftRow component (used by Gantt left panel + List view) ───
function LeftRow({item,hovered,onHover,onEdit,onDelete,onToggle,allItems,isDragging,isSelected,
  activeDragId,onActivateDrag,onMoveUp,onMoveDown,canUp,canDown,onTapWhileDragging,onSelect,
  onMouseDownDrag,dropIndicator,themeKey="dark",isMobile=false,nameColWidth=null}) {
  const T = THEMES[themeKey];
  const isSector = item.type==="sector";
  const c = isSector ? item.color : effectiveColor(item,allItems||[]);
  const fmtD = d => d ? `${String(d.getFullYear()).slice(2)}/${String(d.getMonth()+1).padStart(2,"0")}/${String(d.getDate()).padStart(2,"0")}` : "";
  const fmtDShort = d => d ? `${String(d.getMonth()+1).padStart(2,"0")}/${String(d.getDate()).padStart(2,"0")}` : "";

  const handleClick = e => {
    if (activeDragId && activeDragId !== item.id) { onTapWhileDragging(item.id); return; }
    if (activeDragId === item.id) return;
    onSelect(item.id);
  };

  const rowBg = isDragging ? "rgba(99,102,241,0.22)"
    : isSelected ? T.bgSelected
    : activeDragId&&!isDragging ? "rgba(99,102,241,0.04)"
    : hovered ? T.bgRowHover
    : isSector ? T.bgRow : "transparent";

  if (!item) return null;

  const rowHeight = isMobile ? ROW_HEIGHT_MOBILE : ROW_HEIGHT;
  // Desktop: use custom name col width if provided, else 1fr
  const desktopNameCol = nameColWidth != null ? `${nameColWidth}px` : COL_NAME;
  const gridCols = isMobile ? GRID_COLS_MOBILE : `26px 6px ${desktopNameCol} ${COL_START} ${COL_END} ${COL_PROG}`;
  const detailHeight = isMobile ? DETAIL_HEIGHT_MOBILE : DETAIL_HEIGHT_COMPACT;

  return (
    <>
      <div onMouseEnter={()=>onHover(item.id)} onMouseLeave={()=>onHover(null)}
        onClick={handleClick}
        style={{
          display:"grid", gridTemplateColumns:gridCols,
          alignItems:"center", height:rowHeight,
          padding: isMobile ? "0 10px 0 6px" : "0 6px 0 4px",
          gap: isMobile ? 6 : 0,
          borderBottom: isSelected ? "none" : `1px solid ${T.borderLight}`,
          borderLeft: isSelected ? `3px solid ${c}` : "3px solid transparent",
          background: rowBg,
          opacity: isDragging ? 0.4 : 1,
          transition:"background 0.1s",
          cursor:"pointer", boxSizing:"border-box",
          position:"relative",
        }}>
        <DragBtn active={activeDragId===item.id} onActivate={()=>onActivateDrag(item.id)}
          onMoveUp={onMoveUp} onMoveDown={onMoveDown} canUp={canUp} canDown={canDown}
          onMouseDownDrag={onMouseDownDrag} isMobile={isMobile}/>

        {/* Desktop has separate dot column; mobile merges dot into name */}
        {!isMobile && (
          <div style={{display:"flex",justifyContent:"center"}}>
            {isSector
              ? <div style={{width:8,height:8,borderRadius:2,background:c}}/>
              : item.type==="milestone"
                ? <div style={{width:8,height:8,background:c,transform:"rotate(45deg)",borderRadius:1}}/>
                : <div style={{width:7,height:7,borderRadius:"50%",background:c}}/>
            }
          </div>
        )}

        <div style={{
            overflow:"hidden",
            paddingLeft: isSector ? (isMobile?6:4) : (isMobile?14:14),
            display:"flex",alignItems:"center",gap:isMobile?7:5,minWidth:0,
            borderLeft: isSector ? "none" : `2px solid ${T.border}`,
            marginLeft: isSector ? 0 : 4
          }}>
          {/* Mobile: dot inline before name */}
          {isMobile && (
            isSector
              ? <div style={{width:9,height:9,borderRadius:2,background:c,flexShrink:0}}/>
              : item.type==="milestone"
                ? <div style={{width:9,height:9,background:c,transform:"rotate(45deg)",borderRadius:1,flexShrink:0}}/>
                : <div style={{width:8,height:8,borderRadius:"50%",background:c,flexShrink:0}}/>
          )}
          <span style={{
            fontSize: isMobile ? 13 : 11,
            fontWeight:isSector?700:500,
            color:isSector?T.text:T.textSub,
            whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",
            minWidth:0,flex:1
          }}>
            {item.name}
          </span>
          {item.link && <LinkIcon size={isMobile?12:10} color="#6366f1"/>}
          {/* Assignee avatars — show up to 2 initials + "+N" badge for the rest.
              Sectors don't get this (assignees live on tasks/milestones). */}
          {!isSector && Array.isArray(item.assignees) && item.assignees.length > 0 && (() => {
            const list = item.assignees;
            const shown = list.slice(0, 2);
            const extra = list.length - shown.length;
            const av = isMobile ? 16 : 14;
            const tip = list.map(a => `${a.name || ""}${a.email ? " <" + a.email + ">" : ""}`.trim()).join("\n");
            return (
              <div title={tip} style={{display:"inline-flex",alignItems:"center",flexShrink:0}}>
                {shown.map((a, idx) => {
                  const initial = (a.name || a.email || "?").trim().charAt(0).toUpperCase();
                  return (
                    <span key={a.id || a.email || idx}
                      style={{
                        width:av, height:av, borderRadius:"50%",
                        background:"linear-gradient(135deg,#6366f1,#8b5cf6)",
                        color:"#fff", fontSize: isMobile?10:9, fontWeight:700,
                        display:"flex", alignItems:"center", justifyContent:"center",
                        marginLeft: idx===0 ? 0 : -4,
                        border: `1.5px solid ${T.bg}`,
                        boxSizing:"border-box",
                      }}>{initial}</span>
                  );
                })}
                {extra > 0 && (
                  <span style={{
                    marginLeft:-4, padding:"0 4px", height:av,
                    borderRadius: av,
                    background: T.border, color: T.textMuted,
                    fontSize: isMobile?9:8, fontWeight:700,
                    display:"flex", alignItems:"center", justifyContent:"center",
                    border: `1.5px solid ${T.bg}`,
                    boxSizing:"border-box",
                  }}>+{extra}</span>
                )}
              </div>
            );
          })()}
          {isSector && (
            <div style={{display:"flex",alignItems:"center",gap:2,marginLeft:"auto",flexShrink:0}}>
              <div
                onClick={e=>{e.stopPropagation();onToggle(item.id);}}
                style={{color:"#6366f1",padding: isMobile?"6px 8px":"4px 6px",cursor:"pointer",borderRadius:5,
                  display:"flex",alignItems:"center",
                  background:"transparent",transition:"background 0.12s"}}
                onMouseEnter={e=>e.currentTarget.style.background="rgba(99,102,241,0.15)"}
                onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                {item.collapsed ? <ChevronRight size={isMobile?14:12}/> : <ChevronDown size={isMobile?14:12}/>}
              </div>
            </div>
          )}
        </div>

        <div style={{textAlign:"center",fontSize:isMobile?11:11,color: isSector?T.textFaint:T.textMuted,fontVariantNumeric:"tabular-nums"}}>
          {!isSector && item.start ? (isMobile ? fmtDShort(item.start) : fmtD(item.start)) : ""}
        </div>
        <div style={{textAlign:"center",fontSize:isMobile?11:11,color: isSector?T.textFaint:T.textMuted,fontVariantNumeric:"tabular-nums"}}>
          {!isSector && item.type!=="milestone" && item.end ? (isMobile ? fmtDShort(item.end) : fmtD(item.end)) : ""}
        </div>
        <div style={{textAlign:"center",fontSize:10,fontWeight:700,display:"flex",justifyContent:"center"}}>
          {(()=>{
            if (item.type==="task") {
              const p = item.progress;
              return <span style={{
                background:p===100?"#10b981":p>0?"#6366f1":"transparent",
                color:p>0?"#fff":T.textDim,
                borderRadius:10,padding:p>0?(isMobile?"2px 7px":"1px 6px"):"0",
                fontSize:isMobile?10:9,fontWeight:700,
                border:p===0?("1px solid "+T.border):"none",
              }}>{p}%</span>;
            }
            if (isSector) {
              const kids = (allItems||[]).filter(i=>i.parentId===item.id&&i.type==="task");
              if (!kids.length) return "";
              const avg = Math.round(kids.reduce((s,k)=>s+(k.progress||0),0)/kids.length);
              return <span style={{
                background:avg===100?"#10b981":avg>0?hexAlpha(item.color,0.8):"transparent",
                color:avg>0?"#fff":T.textDim,
                borderRadius:10,padding:avg>0?(isMobile?"2px 7px":"1px 6px"):"0",
                fontSize:isMobile?10:9,fontWeight:700,
                border:avg===0?("1px solid "+T.border):"none",
              }}>{avg}%</span>;
            }
            return "";
          })()}
        </div>
      </div>

      {isSelected && (
        <div style={{height:detailHeight,background:T.bgSelected,borderBottom:`1px solid ${T.border}`,
          borderLeft:`3px solid ${c}`,padding:isMobile?"12px 14px":"10px 12px 10px 14px",boxSizing:"border-box",overflow:"hidden"}}>
          {isSector && (
            <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
              <div style={{width:12,height:12,borderRadius:3,background:item.color,flexShrink:0}}/>
              <span style={{fontSize:isMobile?13:12,color:T.textSub}}>
                {(() => { const kids=(allItems||[]).filter(i=>i.parentId===item.id); return `${kids.filter(k=>k.type==="task").length}개 작업, ${kids.filter(k=>k.type==="milestone").length}개 마일스톤`; })()}
              </span>
            </div>
          )}
          {item.type==="task" && (
            <div style={{marginBottom:isMobile?10:8}}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}>
                <span style={{fontSize:isMobile?11:10,color:T.textFaint,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.05em"}}>진행률</span>
                <span style={{fontSize:isMobile?12:11,fontWeight:700,color:c}}>{item.progress}%</span>
              </div>
              <div style={{height:isMobile?5:4,background:T.border,borderRadius:3,overflow:"hidden"}}>
                <div style={{height:"100%",width:item.progress+"%",background:c,borderRadius:3,transition:"width 0.4s"}}/>
              </div>
            </div>
          )}
          <div style={{display:"flex",gap:isMobile?7:5,flexWrap:"wrap"}}>
            {isSector && (
              <button onClick={e=>{e.stopPropagation();onToggle(item.id);}}
                style={{display:"flex",alignItems:"center",gap:4,padding:isMobile?"7px 12px":"4px 9px",background:themeKey==="dark"?"rgba(255,255,255,0.04)":"rgba(0,0,0,0.04)",border:"1px solid "+T.border,borderRadius:6,cursor:"pointer",fontSize:isMobile?12:11,color:T.textSub}}>
                {item.collapsed?<ChevronRight size={isMobile?12:10}/>:<ChevronDown size={isMobile?12:10}/>}{item.collapsed?"펼치기":"접기"}
              </button>
            )}
            {isSector && (
              <button onClick={e=>{e.stopPropagation();onEdit({type:"new-child",parentId:item.id});}}
                style={{display:"flex",alignItems:"center",gap:4,padding:isMobile?"7px 12px":"4px 9px",background:hexAlpha("#6366f1",0.12),border:"1px solid "+hexAlpha("#6366f1",0.3),borderRadius:6,cursor:"pointer",fontSize:isMobile?12:11,color:"#818cf8",fontWeight:600}}>
                <PlusIcon size={isMobile?12:10} color="#818cf8"/>작업
              </button>
            )}
            <button onClick={e=>{e.stopPropagation();onEdit(item);}}
              style={{display:"flex",alignItems:"center",gap:4,padding:isMobile?"7px 12px":"4px 9px",background:themeKey==="dark"?"rgba(255,255,255,0.04)":"rgba(0,0,0,0.04)",border:"1px solid "+T.border,borderRadius:6,cursor:"pointer",fontSize:isMobile?12:11,color:T.textSub}}>
              <EditIcon size={isMobile?12:10} color={T.textSub}/>편집
            </button>
            {!isSector && (
              item.link
                ? <button onClick={e=>{e.stopPropagation();window.open(item.link,"_blank");}}
                    style={{display:"flex",alignItems:"center",gap:4,padding:isMobile?"7px 12px":"4px 9px",background:hexAlpha("#6366f1",0.1),border:"1px solid "+hexAlpha("#6366f1",0.25),borderRadius:6,cursor:"pointer",fontSize:isMobile?12:11,color:"#818cf8"}}>
                    <LinkIcon size={isMobile?12:10} color="#818cf8"/>링크 열기
                  </button>
                : <button onClick={e=>{e.stopPropagation();onEdit(item);}}
                    style={{display:"flex",alignItems:"center",gap:4,padding:isMobile?"7px 12px":"4px 9px",background:themeKey==="dark"?"rgba(255,255,255,0.04)":"rgba(0,0,0,0.04)",border:"1px solid "+T.border,borderRadius:6,cursor:"pointer",fontSize:isMobile?12:11,color:T.textDim}}>
                    <LinkIcon size={isMobile?12:10} color={T.textDim}/>링크 추가
                  </button>
            )}
            <button onClick={e=>{e.stopPropagation();deleteItemConfirm(item, onDelete, onSelect);}}
              style={{display:"flex",alignItems:"center",gap:4,padding:isMobile?"7px 12px":"4px 9px",background:"rgba(248,113,113,0.08)",border:"1px solid rgba(248,113,113,0.25)",borderRadius:6,cursor:"pointer",fontSize:isMobile?12:11,color:"#f87171"}}>
              <TrashIcon size={isMobile?12:10} color="#f87171"/>삭제
            </button>
          </div>
        </div>
      )}

      {dropIndicator==="above" && (
        <div style={{position:"absolute",top:0,left:0,right:0,height:2,background:"#6366f1",zIndex:9,
          boxShadow:"0 0 6px #6366f1"}}/>
      )}
      {dropIndicator==="below" && (
        <div style={{position:"absolute",bottom:0,left:0,right:0,height:2,background:"#6366f1",zIndex:9,
          boxShadow:"0 0 6px #6366f1"}}/>
      )}
    </>
  );
}

function deleteItemConfirm(item, onDelete, onSelect) {
  // Use the global confirm dialog (set up via window.appConfirm) so it works
  // inside Teams iframe where window.confirm() may be blocked.
  if (window.appConfirm) {
    window.appConfirm(`"${item.name}" 삭제하시겠습니까?`, () => {
      onDelete(item.id);
      onSelect(null);
    });
    return;
  }
  // Fallback to native confirm
  let ok = false;
  try { ok = window.confirm(`"${item.name}" 삭제하시겠습니까?`); }
  catch(e) { ok = true; }
  if (ok) {
    onDelete(item.id);
    onSelect(null);
  }
}

// ─── KANBAN VIEW (mobile: vertical stacked sectors with horizontal columns) ───
const KANBAN_COLS = [
  { key:"todo", label:"할 일", color:"#475569" },
  { key:"inprogress", label:"진행 중", color:"#6366f1" },
  { key:"done", label:"완료", color:"#10b981" },
];

function getKStatus(item) {
  const p = item.progress || 0;
  if (p === 100) return "done";
  if (p === 0) return "todo";
  return "inprogress";
}

function KanbanView({ items, onEdit, onDelete, openAdd, openAddSector, themeKey="dark", isMobile=false, collapsedSectors, toggleSector,
  catCollapsed, toggleCategory, labelW, setLabelW, colResizing, setColResizing,
  onAddCategory, onEditCategory, onMoveCategory }) {
  const T = THEMES[themeKey];
  const [dragCard, setDragCard] = useState(null);
  const [dragOver, setDragOver] = useState(null);

  const categories = items.filter(i => i.type === "category")
    .slice().sort((a,b) => (a.order??0)-(b.order??0));

  // Sectors grouped by category, in order
  const sectorsByCategory = categories.map(cat => ({
    category: cat,
    sectors: items.filter(i => i.type === "sector" && i.parentId === cat.id)
      .slice().sort((a,b) => (a.order??0)-(b.order??0)),
  }));

  // Flat sector list (for kanban column count headers)
  const sectors = sectorsByCategory.flatMap(g => g.sectors);

  const getCards = (sectorId, statusKey) =>
    items.filter(i => i.parentId === sectorId && i.type === "task" && getKStatus(i) === statusKey)
      .slice().sort((a,b) => (a.order??0)-(b.order??0));

  const moveCard = (cardId, newStatus) => {
    const pm = { todo:0, inprogress:50, done:100 };
    onEdit({ _kanbanMove:true, id:cardId, progress:pm[newStatus] });
  };

  const fmt = d => d ? (d.getMonth()+1) + "/" + d.getDate() : "";

  // Resize effect for sector label column
  useEffect(() => {
    if (!colResizing || !setColResizing) return;
    const onMove = e => {
      const dx = e.clientX - colResizing.startX;
      const newW = Math.max(80, Math.min(400, colResizing.startWidth + dx));
      setLabelW(newW);
    };
    const onUp = () => setColResizing(null);
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
    return () => { window.removeEventListener("mousemove", onMove); window.removeEventListener("mouseup", onUp); };
  }, [colResizing, setColResizing, setLabelW]);

  // Sync category label heights to corresponding main category blocks
  // (so labels span exactly their category's vertical space)
  useEffect(() => {
    const apply = () => {
      categories.forEach(cat => {
        const block = document.querySelector(`[data-cat-block="${cat.id}"]`);
        const label = document.querySelector(`[data-cat-id="${cat.id}"]`);
        if (block && label) {
          const h = block.getBoundingClientRect().height;
          if (h > 0) label.style.height = h + "px";
        }
      });
    };
    apply();
    // Watch for changes
    const observers = [];
    categories.forEach(cat => {
      const block = document.querySelector(`[data-cat-block="${cat.id}"]`);
      if (block && typeof ResizeObserver !== "undefined") {
        const ro = new ResizeObserver(apply);
        ro.observe(block);
        observers.push(ro);
      }
    });
    // Fallback: re-apply after short delay (catches initial layout)
    const t = setTimeout(apply, 100);
    return () => { observers.forEach(o => o.disconnect()); clearTimeout(t); };
  }, [categories, items, collapsedSectors]);

  // Mobile vs desktop column widths
  const COL_W = isMobile ? 200 : 220;
  const SECTOR_LABEL_W = isMobile ? 130 : (labelW != null ? labelW : 160);
  const CATEGORY_COL_W = 26;
  const KANBAN_HEADER_H = 56; // header row height (must match below)

  return (
    <div style={{flex:1, display:"flex", flexDirection:"row", background:T.bgDeep, overflow:"hidden"}}>
      {/* === Category column (vertical labels) === */}
      <div style={{width:CATEGORY_COL_W, flexShrink:0, borderRight:`1px solid ${T.border}`, display:"flex", flexDirection:"column", background:T.bgDeep}}>
        {/* Top corner: + button */}
        <div style={{height:KANBAN_HEADER_H, flexShrink:0, borderBottom:`2px solid ${T.border}`, background:T.bgHeader, display:"flex",alignItems:"center",justifyContent:"center"}}>
          <button onClick={onAddCategory} title="대분류 추가"
            style={{
              width:20, height:20, padding:0,
              background:"transparent", border:`1px solid ${T.border}`,
              borderRadius:5, cursor:"pointer", color:T.textSub,
              display:"flex", alignItems:"center", justifyContent:"center",
              transition:"all 0.12s",
            }}
            onMouseEnter={e=>{ e.currentTarget.style.background=hexAlpha("#6366f1",0.15); e.currentTarget.style.borderColor="#6366f1"; }}
            onMouseLeave={e=>{ e.currentTarget.style.background="transparent"; e.currentTarget.style.borderColor=T.border; }}>
            <PlusIcon size={11} color={T.textSub}/>
          </button>
        </div>
        {/* Scrollable category labels — synced with main scroll */}
        <div className="hide-scrollbar" id="kanban-cat-col" style={{flex:1, overflowY:"auto", overflowX:"hidden", position:"relative"}}>
          {/* Categories rendered as flex column matching main content row order */}
          <div style={{display:"flex", flexDirection:"column"}}>
            {sectorsByCategory.map(({ category }, idx) => {
              const canUp = idx > 0;
              const canDown = idx < sectorsByCategory.length - 1;
              const catIsCollapsed = !!catCollapsed[category.id];
              return (
                <div key={category.id}
                  data-cat-id={category.id}
                  onClick={()=>toggleCategory(category.id)}
                  onContextMenu={e=>{ e.preventDefault(); onEditCategory(category.id); }}
                  title={`${category.name} · 클릭: ${catIsCollapsed?"펼치기":"접기"} · 우클릭: 편집`}
                  style={{
                    position:"relative",
                    flexShrink:0,
                    borderBottom:`2px solid ${T.border}`,
                    background: catIsCollapsed ? hexAlpha("#6366f1",0.04) : "transparent",
                    cursor:"pointer",
                    overflow:"hidden",
                    transition:"background 0.12s",
                    display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center",
                    minHeight: 80,
                  }}
                  onMouseEnter={e=>e.currentTarget.style.background=hexAlpha("#6366f1",0.10)}
                  onMouseLeave={e=>e.currentTarget.style.background = catIsCollapsed ? hexAlpha("#6366f1",0.04) : "transparent"}>
                  {canUp && (
                    <button
                      onClick={e=>{ e.stopPropagation(); onMoveCategory(category.id, -1); }}
                      title="위로"
                      style={{position:"absolute",top:1,left:0,right:0,height:14,display:"flex",alignItems:"center",justifyContent:"center",background:"transparent",border:"none",cursor:"pointer",color:T.textFaint,padding:0,zIndex:2}}>
                      <svg width="9" height="6" viewBox="0 0 9 6"><path d="M4.5 0L9 6H0z" fill="currentColor"/></svg>
                    </button>
                  )}
                  <div style={{
                    writingMode:"vertical-rl",
                    textOrientation:"upright",
                    fontSize:11, fontWeight:700, color:T.textSub,
                    letterSpacing:"0.05em",
                    lineHeight:1.15,
                    whiteSpace:"nowrap",
                    overflow:"hidden", textOverflow:"ellipsis",
                    padding: "18px 0",
                  }}>
                    {category.name}
                  </div>
                  {canDown && (
                    <button
                      onClick={e=>{ e.stopPropagation(); onMoveCategory(category.id, 1); }}
                      title="아래로"
                      style={{position:"absolute",bottom:1,left:0,right:0,height:14,display:"flex",alignItems:"center",justifyContent:"center",background:"transparent",border:"none",cursor:"pointer",color:T.textFaint,padding:0,zIndex:2}}>
                      <svg width="9" height="6" viewBox="0 0 9 6"><path d="M4.5 6L0 0H9z" fill="currentColor"/></svg>
                    </button>
                  )}
                </div>
              );
            })}
          </div>
        </div>
      </div>

      {/* === Main kanban area === */}
      <div style={{flex:1, display:"flex", flexDirection:"column", overflow:"hidden", minWidth:0}}>
      <div style={{flex:1, overflowX:"auto", overflowY:"hidden", display:"flex", flexDirection:"column", WebkitOverflowScrolling:"touch"}}>
        <div style={{display:"flex", flexShrink:0, borderBottom:`2px solid ${T.border}`, background:T.bgHeader, height:KANBAN_HEADER_H,
          minWidth: SECTOR_LABEL_W + KANBAN_COLS.length * COL_W}}>
          <div style={{
            width: SECTOR_LABEL_W, flexShrink:0,
            display:"flex", alignItems:"center", justifyContent:"flex-start",
            padding: isMobile?"0 10px":"0 14px", borderRight:`1px solid ${T.border}`, gap:6,
            position:"relative",
          }}>
            <span style={{fontSize:10,fontWeight:700,color:T.textDim,textTransform:"uppercase",letterSpacing:"0.05em"}}>섹터</span>
            {/* Resize handle on right edge */}
            {!isMobile && setColResizing && (
              <div
                onMouseDown={e=>{
                  e.preventDefault();
                  e.stopPropagation();
                  setColResizing({startX: e.clientX, startWidth: SECTOR_LABEL_W});
                }}
                onDoubleClick={()=>setLabelW(null)}
                title="드래그하여 너비 조절 / 더블클릭하여 자동"
                style={{
                  position:"absolute", right:-3, top:0, bottom:0, width:8,
                  cursor:"col-resize", zIndex:10,
                  background: colResizing ? hexAlpha("#6366f1",0.3) : "transparent",
                  transition:"background 0.15s",
                }}
                onMouseEnter={e=>{ if(!colResizing) e.currentTarget.style.background=hexAlpha("#6366f1",0.2); }}
                onMouseLeave={e=>{ if(!colResizing) e.currentTarget.style.background="transparent"; }}
              />
            )}
          </div>
          {KANBAN_COLS.map(col => (
            <div key={col.key} style={{
              width: COL_W, flexShrink:0,
              display:"flex", alignItems:"center", gap:8,
              padding: isMobile?"10px 12px":"12px 16px",
              borderRight:`1px solid ${T.borderLight}`,
              borderTop: `3px solid ${col.color}`,
            }}>
              <div style={{width:8,height:8,borderRadius:"50%",background:col.color,flexShrink:0}}/>
              <span style={{fontSize:isMobile?12:12,fontWeight:700,color:col.color}}>{col.label}</span>
              <span style={{
                fontSize:10,color:T.textDim,background:T.border,
                borderRadius:10,padding:"1px 8px",marginLeft:"auto"
              }}>
                {sectors.reduce((acc,s)=>acc+getCards(s.id,col.key).length,0)}
              </span>
            </div>
          ))}
        </div>

        <div className="kanban-main-scroll" style={{flex:1, overflowY:"auto", overflowX:"visible", minWidth: SECTOR_LABEL_W + KANBAN_COLS.length * COL_W, WebkitOverflowScrolling:"touch"}}
          onScroll={e => {
            // sync scroll to category column
            const catCol = document.getElementById("kanban-cat-col");
            if (catCol && Math.abs(catCol.scrollTop - e.target.scrollTop) > 1) {
              catCol.scrollTop = e.target.scrollTop;
            }
          }}>
          {sectorsByCategory.map(({ category, sectors: catSectors }) => {
            // Empty category: show placeholder block
            if (catSectors.length === 0) {
              return (
                <div key={`cat-${category.id}`} data-cat-block={category.id}
                  style={{minHeight:80, borderBottom:`2px solid ${hexAlpha("#6366f1",0.25)}`, display:"flex", alignItems:"center", justifyContent:"center", padding:"20px", color:T.textFaint, fontSize:12, fontStyle:"italic", background:hexAlpha("#6366f1",0.02)}}>
                  · {category.name} · 비어있음
                </div>
              );
            }
            return (
              <div key={`cat-${category.id}`} data-cat-block={category.id}
                style={{borderBottom:`3px solid ${hexAlpha("#6366f1",0.3)}`}}>
                {catSectors.map(sector => {
            const isCollapsed = collapsedSectors[sector.id] || !!catCollapsed[category.id];
            const totalTasks = items.filter(i=>i.parentId===sector.id&&i.type==="task").length;
            const doneTasks = items.filter(i=>i.parentId===sector.id&&i.type==="task"&&i.progress===100).length;
            const pct = totalTasks ? Math.round(doneTasks/totalTasks*100) : 0;

            return (
              <div key={sector.id} style={{
                display:"flex", flexDirection:"column",
                borderBottom:`2px solid ${hexAlpha(sector.color,0.25)}`,
                minWidth: SECTOR_LABEL_W + KANBAN_COLS.length * COL_W,
              }}>
                <div style={{
                  display:"flex", alignItems:"stretch",
                  background: hexAlpha(sector.color, 0.07),
                  borderBottom: isCollapsed ? "none" : `1px solid ${hexAlpha(sector.color,0.2)}`,
                  flexShrink:0,
                }}>
                  <div style={{
                    width: SECTOR_LABEL_W, flexShrink:0,
                    display:"flex", flexDirection:"column", justifyContent:"center",
                    padding: isMobile?"12px 10px":"14px 14px",
                    borderRight:`2px solid ${hexAlpha(sector.color,0.3)}`,
                    borderLeft:`4px solid ${sector.color}`,
                    gap:8,
                  }}>
                    <div style={{display:"flex",alignItems:"center",gap:8}}>
                      <div style={{width:10,height:10,borderRadius:3,background:sector.color,flexShrink:0}}/>
                      <span style={{fontSize:isMobile?12:13,fontWeight:700,color:T.text,flex:1,minWidth:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                        {sector.name}
                      </span>
                      <button
                        onClick={()=>toggleSector(sector.id)}
                        style={{width:24,height:24,display:"flex",alignItems:"center",justifyContent:"center",
                          background:"transparent",border:"none",cursor:"pointer",flexShrink:0,
                          color:T.textFaint}}
                        title={isCollapsed?"펼치기":"접기"}>
                        {isCollapsed ? <ChevronRight size={12}/> : <ChevronDown size={12}/>}
                      </button>
                    </div>
                    <div>
                      <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
                        <span style={{fontSize:9,color:T.textFaint,fontWeight:600,textTransform:"uppercase"}}>완료율</span>
                        <span style={{fontSize:9,fontWeight:700,color:pct===100?"#10b981":sector.color}}>{pct}%</span>
                      </div>
                      <div style={{height:3,background:themeKey==="dark"?"rgba(255,255,255,0.07)":"rgba(0,0,0,0.07)",borderRadius:2,overflow:"hidden"}}>
                        <div style={{height:"100%",width:pct+"%",background:sector.color,borderRadius:2,transition:"width 0.4s"}}/>
                      </div>
                    </div>
                    <button onClick={()=>openAdd("task",sector.id)} style={{
                      display:"flex",alignItems:"center",gap:5,padding:"6px 8px",
                      background:themeKey==="dark"?"rgba(255,255,255,0.04)":"rgba(0,0,0,0.04)",border:`1px dashed ${hexAlpha(sector.color,0.4)}`,
                      borderRadius:6,cursor:"pointer",fontSize:11,color:sector.color,fontWeight:600,
                      width:"100%",justifyContent:"center"
                    }}>
                      <PlusIcon size={10} color={sector.color}/>작업
                    </button>
                  </div>

                  {KANBAN_COLS.map(col => {
                    const cards = getCards(sector.id, col.key);
                    return (
                      <div key={col.key} style={{
                        width:COL_W, flexShrink:0,
                        display:"flex", alignItems:"center", justifyContent:"center",
                        borderRight:`1px solid ${T.borderLight}`,
                        padding:"8px",
                      }}>
                        {isCollapsed ? (
                          <div style={{
                            display:"flex",alignItems:"center",gap:6,
                            padding:"4px 12px",borderRadius:20,
                            background: cards.length>0 ? hexAlpha(col.color,0.12) : (themeKey==="dark"?"rgba(255,255,255,0.02)":"rgba(0,0,0,0.02)"),
                            border:`1px solid ${cards.length>0?hexAlpha(col.color,0.3):T.border}`,
                          }}>
                            <div style={{width:6,height:6,borderRadius:"50%",background:cards.length>0?col.color:T.textDim}}/>
                            <span style={{fontSize:12,fontWeight:700,color:cards.length>0?col.color:T.textDim}}>{cards.length}</span>
                          </div>
                        ) : (
                          <div style={{width:"100%",height:2,background:cards.length>0?hexAlpha(col.color,0.2):"transparent",borderRadius:1}}/>
                        )}
                      </div>
                    );
                  })}
                </div>

                {!isCollapsed && (
                  <div style={{display:"flex", alignItems:"flex-start"}}>
                    <div style={{width:SECTOR_LABEL_W,flexShrink:0,borderRight:`2px solid ${hexAlpha(sector.color,0.3)}`,borderLeft:`4px solid ${sector.color}`,minHeight:50,background:hexAlpha(sector.color,0.04)}}/>
                    {KANBAN_COLS.map(col => {
                      const cards = getCards(sector.id, col.key);
                      const isOver = dragOver && dragOver.sectorId===sector.id && dragOver.colKey===col.key;
                      return (
                        <div key={col.key}
                          onDragOver={e => { e.preventDefault(); setDragOver({sectorId:sector.id,colKey:col.key}); }}
                          onDragLeave={() => setDragOver(null)}
                          onDrop={e => { e.preventDefault(); if(dragCard) moveCard(dragCard, col.key); setDragOver(null); }}
                          style={{
                            width:COL_W, flexShrink:0, minHeight:100,
                            padding: isMobile?"8px 6px":"10px 8px",
                            display:"flex", flexDirection:"column", gap:7,
                            borderRight:`1px solid ${T.borderLight}`,
                            background: isOver ? hexAlpha(col.color,0.05) : "transparent",
                            transition:"background 0.15s",
                          }}>
                          {cards.map(card => {
                            const cc = effectiveColor(card, items);
                            return (
                              <div key={card.id}
                                draggable
                                onDragStart={() => setDragCard(card.id)}
                                onDragEnd={() => { setDragCard(null); setDragOver(null); }}
                                onClick={() => onEdit(card)}
                                style={{
                                  background:T.bgCard,
                                  border:`1px solid ${T.border}`,
                                  borderTop:`3px solid ${cc}`,
                                  borderRadius:8, padding:"10px 11px",
                                  cursor:"pointer",
                                  opacity: dragCard===card.id ? 0.35 : 1,
                                  transition:"all 0.15s, opacity 0.1s",
                                  boxShadow: dragCard===card.id?"none":"0 2px 10px rgba(0,0,0,0.15)"
                                }}>
                                <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:6}}>
                                  {card.type==="milestone"
                                    ? <span style={{fontSize:9,fontWeight:700,color:"#f59e0b",background:"rgba(245,158,11,0.15)",borderRadius:4,padding:"1px 6px"}}>마일스톤</span>
                                    : <span style={{fontSize:9,fontWeight:700,color:cc,background:hexAlpha(cc,0.15),borderRadius:4,padding:"1px 6px"}}>작업</span>
                                  }
                                  {card.link && <LinkIcon size={10} color="#6366f1"/>}
                                </div>
                                <div style={{fontSize:isMobile?13:12,fontWeight:600,color:T.text,lineHeight:1.4,marginBottom:6}}>
                                  {card.name}
                                </div>
                                <div style={{display:"flex",flexWrap:"wrap",gap:5,marginBottom:6}}>
                                  {card.start && (
                                    <span style={{fontSize:10,color:T.textFaint}}>
                                      {fmt(card.start)}{card.end && card.type!=="milestone" ? ` → ${fmt(card.end)}` : ""}
                                    </span>
                                  )}
                                </div>
                                {card.type==="task" && (
                                  <div>
                                    <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}>
                                      <span style={{fontSize:9,color:T.textDim}}>진행률</span>
                                      <span style={{fontSize:9,fontWeight:700,color:cc}}>{card.progress}%</span>
                                    </div>
                                    <div style={{height:3,background:T.border,borderRadius:2,overflow:"hidden"}}>
                                      <div style={{height:"100%",width:card.progress+"%",background:cc,borderRadius:2}}/>
                                    </div>
                                  </div>
                                )}
                              </div>
                            );
                          })}
                          {cards.length === 0 && (
                            <div style={{
                              height:56,display:"flex",alignItems:"center",justifyContent:"center",
                              border:`1px dashed ${T.border}`,borderRadius:8,color:T.border,fontSize:11
                            }}>없음</div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            );
                })}
              </div>
            );
          })}
        </div>
      </div>
      </div>{/* end main kanban area */}
    </div>
  );
}

function serializeItems(items) {
  const safeISO = d => {
    try { return d && !isNaN(d.getTime()) ? d.toISOString() : null; }
    catch(e) { return null; }
  };
  return items.map(item => ({
    ...item,
    start: safeISO(item.start),
    end: safeISO(item.end),
  }));
}

function deserializeItems(items) {
  if (!Array.isArray(items)) return [];
  const parsed = items.map(item => {
    if (!item) return null;
    const s = item.start ? new Date(item.start) : null;
    const e = item.end ? new Date(item.end) : null;
    return {
      ...item,
      start: s && !isNaN(s.getTime()) ? s : null,
      end: e && !isNaN(e.getTime()) ? e : null,
      color: item.color || undefined,
      progress: item.progress || 0,
      deps: item.deps || [],
    };
  }).filter(Boolean);

  let result = parsed;
  // Migration: legacy data without categories — auto-create default category
  // and assign all sectors to it
  const hasCategory = parsed.some(i => i.type === "category");
  if (!hasCategory && parsed.some(i => i.type === "sector")) {
    const maxId = Math.max(0, ...parsed.map(i => i.id || 0));
    const defaultCatId = maxId + 1;
    const defaultCategory = { id: defaultCatId, type: "category", name: "기본", collapsed: false, order: 0 };
    result = [defaultCategory, ...parsed.map(item =>
      item.type === "sector" && !item.parentId ? { ...item, parentId: defaultCatId } : item
    )];
  } else if (hasCategory) {
    // Repair orphan sectors: sectors whose parentId points to a missing/non-category item
    // get reassigned to the first available category (so they show up in flatRows)
    const categoryIds = new Set(parsed.filter(i => i.type === "category").map(i => i.id));
    const firstCat = parsed.find(i => i.type === "category" && (i.order ?? 0) === Math.min(...parsed.filter(c => c.type === "category").map(c => c.order ?? 0)));
    const fallbackCatId = firstCat?.id;
    if (fallbackCatId) {
      result = parsed.map(item => {
        if (item.type === "sector" && !categoryIds.has(item.parentId)) {
          return { ...item, parentId: fallbackCatId };
        }
        return item;
      });
    }
  }
  // CRITICAL: Bump uid counter so newly added items don't collide with
  // existing ids. Without this, uid() can return ids already in use,
  // causing setItems updates to be effectively no-ops or data corruption.
  syncUid(result);
  return result;
}

export default function GanttApp() {
  // Start with [] (not initialItems) so the sample 9-task gantt does NOT
  // flash before the real data finishes loading. A loading overlay is
  // rendered while isLoading is true; if the server has no data we fall
  // back to initialItems explicitly inside the load effect.
  const [items, _setItems] = useState([]);
  const [isLoading, setIsLoading] = useState(true);
  const isLoadedRef = useRef(false); // becomes true after initial load attempt completes
  const channelResolvedRef = useRef(false); // true once getChannelId has resolved (even to "default")
  const lastServerCountRef = useRef(0); // last known server item count for safety check
  const undoStackRef = useRef([]);
  const redoStackRef = useRef([]);
  const skipHistoryRef = useRef(false);
  const HISTORY_LIMIT = 10;
  const [historyVer, setHistoryVer] = useState(0);

  const dedupeItems = (arr) => {
    if (!Array.isArray(arr)) return arr;
    const seen = new Set();
    return arr.filter(item => {
      if (!item || seen.has(item.id)) return false;
      seen.add(item.id);
      return true;
    });
  };

  const setItems = useCallback((updater) => {
    _setItems(prev => {
      const raw = typeof updater === "function" ? updater(prev) : updater;
      const next = dedupeItems(raw);
      if (!skipHistoryRef.current && prev && prev.length > 0) {
        undoStackRef.current.push(prev);
        if (undoStackRef.current.length > HISTORY_LIMIT) undoStackRef.current.shift();
        redoStackRef.current = [];
        setHistoryVer(v=>v+1);
      }
      skipHistoryRef.current = false;
      return next;
    });
  }, []);

  const undo = useCallback(() => {
    if (undoStackRef.current.length === 0) return;
    const prev = undoStackRef.current.pop();
    skipHistoryRef.current = true;
    _setItems(current => {
      redoStackRef.current.push(current);
      if (redoStackRef.current.length > HISTORY_LIMIT) redoStackRef.current.shift();
      return prev;
    });
    setHistoryVer(v=>v+1);
  }, []);

  const redo = useCallback(() => {
    if (redoStackRef.current.length === 0) return;
    const next = redoStackRef.current.pop();
    skipHistoryRef.current = true;
    _setItems(current => {
      undoStackRef.current.push(current);
      if (undoStackRef.current.length > HISTORY_LIMIT) undoStackRef.current.shift();
      return next;
    });
    setHistoryVer(v=>v+1);
  }, []);

  const [activeTab, setActiveTab] = useState('list'); // changed default; we'll override based on viewport
  const [channelId, setChannelId] = useState("default");
  const [dbStatus, setDbStatus] = useState("idle");
  const [projectName, setProjectName] = useState("프로젝트 간트");
  const [updateAvailable, setUpdateAvailable] = useState(false);
  const [searchQuery, setSearchQuery] = useState("");
  const [themeKey, setThemeKey] = useState("dark");
  const T = THEMES[themeKey];

  const [editingName, setEditingName] = useState(false);
  const saveTimerRef = useRef(null);
  const importInputRef = useRef(null);

  const [zoomIdx, setZoomIdx] = useState(3);
  const [customDayWidth, setCustomDayWidth] = useState(null);
  const [viewStart, setViewStart] = useState(addDays(today,-7));
  const [showModal, setShowModal] = useState(false);
  const [confirmDialog, setConfirmDialog] = useState(null); // { message, onConfirm }
  // Popup shown after a save when new assignees were added — asks the user
  // whether to send them an email notification. Shape:
  //   { task: <item>, newAssignees: [{id,name,email}, ...] }
  const [assigneeNotifyDialog, setAssigneeNotifyDialog] = useState(null);
  // Category add/edit modal
  const [showCatModal, setShowCatModal] = useState(false);
  const [catForm, setCatForm] = useState(null); // { id?, name, sectorIds: [] }
  const [editItem, setEditItem] = useState(null);
  const [form, setForm] = useState(null);
  const [formError, setFormError] = useState("");
  const [dragState, setDragState] = useState(null);
  const [activeDragId, setActiveDragId] = useState(null);
  const [rowDragState, setRowDragState] = useState(null);
  const [chartPan, setChartPan] = useState(null);
  const zoomStateRef = useRef({});
  const [rowDragPos, setRowDragPos] = useState({x:0, y:0});

  const [selectedId, setSelectedId] = useState(null);
  const [hoveredId, setHoveredId] = useState(null);
  const [linkingFrom, setLinkingFrom] = useState(null);
  const [showMilestones, setShowMilestones] = useState(true);
  const [showLeftPanel, setShowLeftPanel] = useState(true);
  const [nameColWidth, setNameColWidth] = useState(null); // null = auto (1fr), or px number
  const [colResizing, setColResizing] = useState(null); // { startX, startWidth }
  const [kanbanCollapsed, setKanbanCollapsed] = useState({});
  const [kanbanCatCollapsed, setKanbanCatCollapsed] = useState({}); // independent of gantt category collapse
  const [kanbanLabelW, setKanbanLabelW] = useState(null); // null = default, or px (for resizable sector label column)
  const [kanbanColResizing, setKanbanColResizing] = useState(null);
  const toggleKanbanSector = useCallback(id => {
    setKanbanCollapsed(prev => ({...prev, [id]: !prev[id]}));
  }, []);
  const toggleKanbanCategory = useCallback(id => {
    setKanbanCatCollapsed(prev => ({...prev, [id]: !prev[id]}));
  }, []);
  // App-wide UI zoom (Ctrl+wheel)
  const [appZoom, setAppZoom] = useState(1.0);
  const [menuOpen, setMenuOpen] = useState(false);

  const [viewport, setViewport] = useState({w: window.innerWidth, h: window.innerHeight});
  useEffect(()=>{
    const onResize = ()=>setViewport({w:window.innerWidth,h:window.innerHeight});
    window.addEventListener("resize",onResize);
    return ()=>window.removeEventListener("resize",onResize);
  },[]);

  // App-wide zoom via Ctrl+wheel
  // Register an in-app confirm dialog so things like delete work even in
  // Teams iframe where window.confirm() may be blocked
  useEffect(()=>{
    window.appConfirm = (message, onConfirm) => {
      setConfirmDialog({ message, onConfirm });
    };
    return () => { delete window.appConfirm; };
  }, []);

  // Expose backup utilities to window so user can recover via DevTools console
  useEffect(()=>{
    window.ganttBackup = {
      list: () => {
        const key = "gantt_backup_" + channelId;
        const raw = localStorage.getItem(key);
        if (!raw) { console.log("백업 없음"); return []; }
        const list = JSON.parse(raw);
        console.table(list.map((b, i) => ({
          idx: i,
          time: b.time,
          itemCount: (b.data?.items || []).length,
        })));
        return list;
      },
      restore: (idx = 0) => {
        const key = "gantt_backup_" + channelId;
        const raw = localStorage.getItem(key);
        if (!raw) { console.error("백업 없음"); return; }
        const list = JSON.parse(raw);
        const target = list[idx];
        if (!target) { console.error("해당 인덱스 백업 없음"); return; }
        const d = target.data;
        if (Array.isArray(d)) {
          setItems(deserializeItems(d));
        } else if (d.items) {
          setItems(deserializeItems(d.items));
          if (d.projectName) setProjectName(d.projectName);
        }
        console.log("복구됨:", target.time, "(", (d.items || d).length, "개 항목)");
      },
    };
  }, [channelId]);
  useEffect(()=>{
    const onWheel = (e) => {
      if (e.ctrlKey || e.metaKey) {
        e.preventDefault();
        const dir = e.deltaY < 0 ? 1 : -1;
        setAppZoom(prev => {
          const step = 0.1;
          const next = Math.max(0.5, Math.min(2.0, prev + dir * step));
          return Math.round(next * 100) / 100;
        });
      }
    };
    const onKey = (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key === "0") {
        e.preventDefault();
        setAppZoom(1.0);
      } else if ((e.ctrlKey || e.metaKey) && (e.key === "=" || e.key === "+")) {
        e.preventDefault();
        setAppZoom(prev => Math.min(2.0, Math.round((prev + 0.1) * 100) / 100));
      } else if ((e.ctrlKey || e.metaKey) && e.key === "-") {
        e.preventDefault();
        setAppZoom(prev => Math.max(0.5, Math.round((prev - 0.1) * 100) / 100));
      }
    };
    // passive:false so we can preventDefault
    window.addEventListener("wheel", onWheel, {passive: false});
    window.addEventListener("keydown", onKey);
    return () => {
      window.removeEventListener("wheel", onWheel);
      window.removeEventListener("keydown", onKey);
    };
  }, []);

  const isMobile = viewport.w < 640;
  const isTablet = viewport.w >= 640 && viewport.w < 1024;
  const isLandscape = viewport.w > viewport.h;

  // On mobile, default to list tab on first mount; desktop defaults to gantt
  const initializedTabRef = useRef(false);
  useEffect(()=>{
    if (initializedTabRef.current) return;
    initializedTabRef.current = true;
    setActiveTab(isMobile ? 'list' : 'gantt');
  }, [isMobile]);

  // Safety: if viewport switches from mobile to desktop while on list tab, fall back to gantt
  useEffect(()=>{
    if (!isMobile && activeTab === 'list') setActiveTab('gantt');
  }, [isMobile, activeTab]);

  // (auto-scroll on fit→non-fit transition is defined below, after isFit is computed)
  const prevIsFitRef = useRef(true);
  const skipAutoScrollRef = useRef(false);

  useEffect(()=>{
    const checkVersion = async () => {
      try {
        const res = await fetch("/version.json?t=" + Date.now());
        const data = await res.json();
        if (data.version && data.version !== APP_VERSION) {
          setUpdateAvailable(true);
        }
      } catch(e) {}
    };
    checkVersion();
    const timer = setInterval(checkVersion, 5 * 60 * 1000);
    return () => clearInterval(timer);
  }, []);

  useEffect(()=>{
    const t = THEMES[themeKey];
    setGlobalTheme(themeKey);
    const root = document.documentElement;
    Object.entries(t).forEach(([k,v]) => root.style.setProperty("--t-"+k, v));
    document.body.style.background = t.bg;
    document.body.style.color = t.text;
  }, [themeKey]);

  useEffect(()=>{
    let cancelled = false;
    // Reset load flag — we're loading data for a (possibly new) channel.
    // While loading, save is disabled (prevents overwriting server data with stale local items).
    isLoadedRef.current = false;
    setIsLoading(true);
    // Helper to flip the loading flag in one place — keeps every exit path
    // (success, no-data, error, error+backup) consistent so the loading
    // overlay always disappears.
    const finishLoading = () => { if (!cancelled) setIsLoading(false); };
    getChannelId().then(async id => {
      if (cancelled) return;
      setChannelId(id);
      channelResolvedRef.current = true;
      setCachedChannelId(id);
      try {
        const data = await sbGet(id);
        if (cancelled) return;
        if (data) {
          if (Array.isArray(data)) {
            const loaded = deserializeItems(data);
            setItems(loaded);
            lastServerCountRef.current = loaded.length;
          } else if (data.items) {
            const loaded = deserializeItems(data.items);
            setItems(loaded);
            lastServerCountRef.current = loaded.length;
            if (data.projectName) setProjectName(data.projectName);
            if (data.uiState) {
              const ui = data.uiState;
              if (ui.activeTab && !isMobile) setActiveTab(ui.activeTab); // mobile keeps list default
              if (ui.zoomIdx !== undefined) setZoomIdx(ui.zoomIdx);
              if (ui.customDayWidth !== undefined) setCustomDayWidth(ui.customDayWidth);
              if (ui.showMilestones !== undefined) setShowMilestones(ui.showMilestones);
              if (ui.themeKey) setThemeKey(ui.themeKey);
              if (ui.kanbanCollapsed && typeof ui.kanbanCollapsed === "object") setKanbanCollapsed(ui.kanbanCollapsed);
              if (ui.kanbanCatCollapsed && typeof ui.kanbanCatCollapsed === "object") setKanbanCatCollapsed(ui.kanbanCatCollapsed);
              if (typeof ui.kanbanLabelW === "number") setKanbanLabelW(ui.kanbanLabelW);
              if (typeof ui.appZoom === "number" && ui.appZoom >= 0.5 && ui.appZoom <= 2) setAppZoom(ui.appZoom);
            }
          }
          // Local backup of successful load
          try {
            const backupKey = "gantt_backup_" + id;
            const existingRaw = localStorage.getItem(backupKey);
            const existing = existingRaw ? JSON.parse(existingRaw) : [];
            existing.unshift({ time: new Date().toISOString(), data });
            localStorage.setItem(backupKey, JSON.stringify(existing.slice(0, 10))); // keep last 10
          } catch {}
          isLoadedRef.current = true;
        } else {
          // No data on server but request succeeded — first time use, allow saves.
          // Seed with the default sample so the user has something to start from
          // (previously this relied on the items state still holding initialItems).
          setItems(initialItems);
          lastServerCountRef.current = 0;
          isLoadedRef.current = true;
        }
      } catch(e) {
        console.error("Supabase load error:", e);
        // CRITICAL: Don't set isLoadedRef=true on load failure!
        // This prevents the auto-save from overwriting server data with local defaults.
        setDbStatus("error");
        // Try to load from local backup
        let restored = false;
        try {
          const backupKey = "gantt_backup_" + id;
          const existingRaw = localStorage.getItem(backupKey);
          if (existingRaw) {
            const existing = JSON.parse(existingRaw);
            if (existing[0]?.data) {
              const d = existing[0].data;
              if (Array.isArray(d)) {
                setItems(deserializeItems(d));
              } else if (d.items) {
                setItems(deserializeItems(d.items));
                if (d.projectName) setProjectName(d.projectName);
              }
              restored = true;
              alert("⚠️ 서버에서 데이터를 불러올 수 없어 로컬 백업으로 복구했습니다.\n저장 기능은 일시적으로 비활성화됩니다.\n네트워크 확인 후 새로고침해주세요.");
            }
          }
        } catch {}
        // If we couldn't restore from backup either, at least show the
        // sample data instead of a permanently-empty screen.
        if (!restored) setItems(initialItems);
      } finally {
        finishLoading();
      }
    }).catch(() => {
      // getChannelId() should never reject (it always resolves), but be defensive.
      if (!cancelled) {
        setItems(initialItems);
        finishLoading();
      }
    });
    // NOTE: removed the second microsoftTeams.app.initialize() block here.
    // It duplicated work already done inside getChannelId() above and added
    // no behaviour beyond an extra round-trip on first paint.
    return () => { cancelled = true; };
  }, []);

  useEffect(()=>{
    // SAFETY: never save until initial load attempt completed.
    // This prevents overwriting server data with local defaults during network issues.
    if (!isLoadedRef.current) {
      setDbStatus("idle");
      return;
    }
    // SAFETY 2: don't save until getChannelId has actually resolved.
    // Otherwise we might save default-channel data into the wrong slot.
    if (!channelResolvedRef.current) {
      setDbStatus("idle");
      return;
    }
    if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
    setDbStatus("saving");
    saveTimerRef.current = setTimeout(async () => {
      const currentItemCount = items.length;
      // NOTE: Earlier we had a "50% data loss" confirmation here using window.confirm,
      // but window.confirm() is blocked in Teams iframe — it returns false silently,
      // causing legitimate saves to be skipped. The isLoadedRef guard above already
      // prevents the original data-wipe bug (saving defaults before load completes),
      // so we proceed straight to saving here.
      try {
        await sbSave(channelId, {
          items: serializeItems(items),
          projectName,
          uiState: { activeTab, zoomIdx, customDayWidth, showMilestones, themeKey, kanbanCollapsed, kanbanCatCollapsed, kanbanLabelW, appZoom }
        });
        lastServerCountRef.current = currentItemCount; // update after successful save
        // Local backup after successful save
        try {
          const backupKey = "gantt_backup_" + channelId;
          const existingRaw = localStorage.getItem(backupKey);
          const existing = existingRaw ? JSON.parse(existingRaw) : [];
          existing.unshift({
            time: new Date().toISOString(),
            data: { items: serializeItems(items), projectName }
          });
          localStorage.setItem(backupKey, JSON.stringify(existing.slice(0, 10)));
        } catch {}
        setDbStatus("saved");
        setTimeout(()=>setDbStatus("idle"), 2000);
      } catch(e) {
        console.error("Save error:", e);
        setDbStatus("error");
        setTimeout(async () => {
          try {
            await sbSave(channelId, {
              items: serializeItems(items),
              projectName,
              uiState: { activeTab, zoomIdx, customDayWidth, showMilestones, themeKey, kanbanCollapsed, kanbanCatCollapsed, kanbanLabelW, appZoom }
            });
            lastServerCountRef.current = currentItemCount;
            setDbStatus("saved");
            setTimeout(()=>setDbStatus("idle"), 2000);
          } catch(e2) {
            console.error("Retry failed:", e2);
          }
        }, 3000);
      }
    }, 1000);
  }, [items, projectName, channelId, kanbanCollapsed, kanbanCatCollapsed, kanbanLabelW, appZoom, activeTab, zoomIdx, customDayWidth, showMilestones, themeKey]);

  // Mobile-specific: list tab uses left-panel-style rendering full width
  // Gantt tab on mobile: chart only, no left panel
  // Width calculations
  // Left panel width grows when user enlarges the name column
  // Base: 480px desktop / 380px tablet / full viewport mobile
  // Other columns sum to ~192px + 22px category column on desktop
  const CATEGORY_COL_W = 26;
  const leftPanelBase = isMobile ? viewport.w : (isTablet ? 380 : 480);
  const otherColsWidth = 26 + 6 + 60 + 60 + 54 + 12 + CATEGORY_COL_W; // grid + paddings + category col
  // When user resizes name column, left panel grows/shrinks with it.
  // Auto mode (nameColWidth == null): use leftPanelBase.
  // Min width clamp keeps the panel usable even at small name widths.
  const leftPanelMinWidth = otherColsWidth + 80; // 80 = min name col
  const leftPanelWidth = isMobile
    ? viewport.w
    : (nameColWidth != null ? Math.max(leftPanelMinWidth, nameColWidth + otherColsWidth) : leftPanelBase);
  const leftPanelOverlay = false;

  const chartRef = useRef(null);
  const headerRef = useRef(null);
  const leftPanelListRef = useRef(null);
  const isSyncingScroll = useRef(false);

  const zoom = ZOOM_PRESETS[zoomIdx];
  const isFit = zoom.key === "fit";
  const baseDW = customDayWidth !== null ? customDayWidth : zoom.dw;

  const dataRange = useMemo(() => {
    const allDates = items.flatMap(i => [i.start, i.end]).filter(d => d && d instanceof Date && !isNaN(d.getTime()));
    if (!allDates.length) return { minD: addDays(today, -3), maxD: addDays(today, 30) };
    const minD = new Date(Math.min(...allDates.map(d => d.getTime())));
    const maxD = new Date(Math.max(...allDates.map(d => d.getTime())));
    if (isNaN(minD.getTime()) || isNaN(maxD.getTime())) return { minD: addDays(today, -3), maxD: addDays(today, 30) };
    // Pad both sides so right-edge labels (e.g. "Meetings (976 d)") fit and bars
    // don't touch the chart edges. Right side gets more padding because labels
    // are rendered to the right of bars.
    const totalDays = Math.ceil((maxD.getTime() - minD.getTime()) / 86400000) + 1;
    const leftPad = Math.max(7, Math.ceil(totalDays * 0.04));
    const rightPad = Math.max(60, Math.ceil(totalDays * 0.18));
    minD.setDate(minD.getDate() - leftPad);
    maxD.setDate(maxD.getDate() + rightPad);
    return { minD, maxD };
  }, [items]);

  // Extended range for non-fit zoom: ±5 years around data so users can scroll freely
  // back to past years (2023, 2024, etc) and forward to future years (2027, 2028, etc).
  // Fit zoom uses dataRange directly to keep the "fit" semantic of focusing on actual work.
  const EXTENDED_PADDING_DAYS = 365 * 5; // 5 years on each side
  const extendedRange = useMemo(() => {
    const minD = new Date(dataRange.minD); minD.setDate(minD.getDate() - EXTENDED_PADDING_DAYS);
    const maxD = new Date(dataRange.maxD); maxD.setDate(maxD.getDate() + EXTENDED_PADDING_DAYS);
    return { minD, maxD };
  }, [dataRange]);

  const dataRangeDays = Math.max(dateDiff(dataRange.minD, dataRange.maxD) + 1, 30);
  const extendedRangeDays = dateDiff(extendedRange.minD, extendedRange.maxD) + 1;
  // On mobile in gantt tab, chart spans full viewport width
  const chartW = chartRef.current?.clientWidth || (isMobile ? viewport.w : Math.max(viewport.w - leftPanelWidth - 20, 400));
  const DAY_WIDTH = isFit ? Math.max(0.15, chartW / dataRangeDays) : Math.max(0.15, baseDW);
  // Fit zoom: render only data range (so data fills the screen)
  // Other zoom levels: render extended range (±5yr) so users can scroll past/future
  const totalDays = isFit ? dataRangeDays : extendedRangeDays;

  zoomStateRef.current = { isFit, customDayWidth, zoom, dataRangeDays };

  const effectiveViewStart = isFit ? dataRange.minD : extendedRange.minD;
  const viewEnd = addDays(effectiveViewStart, totalDays - 1);

  // On mount and whenever entering Gantt view, auto-scroll the chart to the data start
  // (so users land on actual work, not 5 years of empty calendar).
  // Skipped when wheel-zoom is in progress (it handles its own scroll).
  useEffect(()=>{
    if (activeTab !== "gantt") return;
    if (skipAutoScrollRef.current) {
      skipAutoScrollRef.current = false;
      prevIsFitRef.current = isFit;
      return;
    }
    requestAnimationFrame(() => {
      const c = chartRef.current;
      if (!c) return;
      // In fit mode, viewStart IS dataRange.minD so target = 0
      // In non-fit, viewStart is extendedRange.minD so we need offset to data start
      const targetX = isFit ? 0 : Math.max(0, dateDiff(extendedRange.minD, dataRange.minD) * DAY_WIDTH - 20);
      // Only auto-scroll on initial mount of gantt
      if (c.scrollLeft < 1) {
        c.scrollLeft = targetX;
      }
      prevIsFitRef.current = isFit;
    });
  }, [activeTab, isFit, DAY_WIDTH, extendedRange.minD, dataRange.minD]);

  const flatRows = useMemo(()=>{
    const rows=[];
    const q = searchQuery.toLowerCase().trim();
    const categories = items.filter(i=>i.type==="category").slice().sort((a,b)=>(a.order??0)-(b.order??0));
    const categoryIds = new Set(categories.map(c => c.id));
    // Track which sectors have been rendered, so orphans (whose parentId points
    // to a non-category) can be added under the first category as a safety net.
    const renderedSectorIds = new Set();
    const renderSector = (sector, category) => {
      if (renderedSectorIds.has(sector.id)) return;
      renderedSectorIds.add(sector.id);
      const kids = items.filter(i=>i.parentId===sector.id&&(showMilestones||i.type!=="milestone"))
        .slice().sort((a,b)=>(a.order??0)-(b.order??0));
      const matchedKids = q ? kids.filter(k=>k.name.toLowerCase().includes(q)) : kids;
      const sectorMatch = q ? (sector.name.toLowerCase().includes(q) || category.name.toLowerCase().includes(q) || matchedKids.length > 0) : true;
      if (!sectorMatch) return;
      rows.push({item:sector,depth:0,categoryId:category.id});
      if (!sector.collapsed && !category.collapsed) {
        const showKids = q ? matchedKids : kids;
        showKids.forEach(child=>rows.push({item:child,depth:1,categoryId:category.id}));
      }
    };
    categories.forEach(category => {
      const sectorsInCat = items.filter(i=>i.type==="sector" && i.parentId===category.id)
        .slice().sort((a,b)=>(a.order??0)-(b.order??0));
      sectorsInCat.forEach(sector => renderSector(sector, category));
    });
    // Catch orphan sectors: those whose parentId doesn't match any category.
    // Show them under the first category so they're always visible.
    if (categories.length > 0) {
      const firstCat = categories[0];
      const orphans = items.filter(i =>
        i.type === "sector" &&
        !renderedSectorIds.has(i.id) &&
        !categoryIds.has(i.parentId)
      ).slice().sort((a,b)=>(a.order??0)-(b.order??0));
      orphans.forEach(sector => renderSector(sector, firstCat));
    }
    return rows;
  },[items,showMilestones,searchQuery]);

  // All categories (for rendering label column even when collapsed/empty)
  const allCategories = useMemo(() =>
    items.filter(i=>i.type==="category").slice().sort((a,b)=>(a.order??0)-(b.order??0))
  , [items]);

  const baseRowHeight = isMobile ? ROW_HEIGHT_MOBILE : ROW_HEIGHT;
  const baseDetailHeight = isMobile ? DETAIL_HEIGHT_MOBILE : DETAIL_HEIGHT_COMPACT;

  const getRowHeight = useCallback((rowIdx) => {
    const item = flatRows[rowIdx]?.item;
    return item && item.id === selectedId ? baseRowHeight + baseDetailHeight : baseRowHeight;
  }, [flatRows, selectedId, baseRowHeight, baseDetailHeight]);

  const getRowTop = useCallback((rowIdx) => {
    let top = 0;
    const MIN_HEIGHT = baseRowHeight;
    const myCatId = flatRows[rowIdx]?.categoryId;
    // Add spacer offsets for empty categories that come before this row's category
    for (const cat of allCategories) {
      if (cat.id === myCatId) break;
      const hasAny = flatRows.some(r => r.categoryId === cat.id);
      if (!hasAny) top += MIN_HEIGHT;
    }
    for (let i = 0; i < rowIdx; i++) top += getRowHeight(i);
    return top;
  }, [getRowHeight, flatRows, allCategories, baseRowHeight]);

  const totalChartHeight = useMemo(() => {
    let h = 0;
    for (let i = 0; i < flatRows.length; i++) h += getRowHeight(i);
    return h;
  }, [flatRows, getRowHeight]);

  // Compute pixel range (top, height) for each category, plus collapsed categories
  // shown as compact bars with no expanded rows
  const categoryRanges = useMemo(() => {
    const ranges = [];
    const MIN_HEIGHT = baseRowHeight; // empty cat still takes one row
    let yOffset = 0;
    let rowIdx = 0;
    allCategories.forEach(cat => {
      const startIdx = rowIdx;
      while (rowIdx < flatRows.length && flatRows[rowIdx].categoryId === cat.id) {
        rowIdx++;
      }
      let height = 0;
      for (let i = startIdx; i < rowIdx; i++) height += getRowHeight(i);
      if (height === 0) height = MIN_HEIGHT; // empty category placeholder
      ranges.push({ category: cat, top: yOffset, height, startIdx, endIdx: rowIdx, isEmpty: rowIdx === startIdx });
      yOffset += height;
    });
    return ranges;
  }, [allCategories, flatRows, getRowHeight, baseRowHeight]);

  // Total height (categories + their rows)
  const totalLeftHeight = useMemo(() => {
    return categoryRanges.reduce((s, r) => s + r.height, 0);
  }, [categoryRanges]);

  const days = useMemo(()=>{
    const arr=[]; const cur=new Date(effectiveViewStart);
    while(cur<=viewEnd){arr.push(new Date(cur));cur.setDate(cur.getDate()+1);}
    return arr;
  },[effectiveViewStart,viewEnd]);

  const headerGroups = useMemo(()=>{
    const g=[];
    days.forEach((d,i)=>{
      const label = DAY_WIDTH < 5 ? `${d.getFullYear()}`
        : DAY_WIDTH < 18 ? `${d.getFullYear()}년 ${d.getMonth()+1}월`
        : `${d.getFullYear()}년 ${d.getMonth()+1}월`;
      if (!g.length||g[g.length-1].label!==label) g.push({label,count:1});
      else g[g.length-1].count++;
    });
    return g;
  },[days,DAY_WIDTH]);

  useEffect(()=>{
    if (!dragState) return;
    const onMove=e=>{
      const delta=Math.round((e.clientX-dragState.startX)/DAY_WIDTH);
      setItems(prev=>prev.map(t=>{
        if(t.id!==dragState.id) return t;
        if(dragState.type==="move") return {...t,start:addDays(dragState.origStart,delta),end:addDays(dragState.origEnd,delta)};
        if(dragState.type==="resize-right"){const ne=addDays(dragState.origEnd,delta);return ne>=t.start?{...t,end:ne}:t;}
        if(dragState.type==="resize-left") {const ns=addDays(dragState.origStart,delta);return ns<=t.end?{...t,start:ns}:t;}
        return t;
      }));
    };
    const onUp=()=>setDragState(null);
    window.addEventListener("mousemove",onMove); window.addEventListener("mouseup",onUp);
    return ()=>{window.removeEventListener("mousemove",onMove);window.removeEventListener("mouseup",onUp);};
  },[dragState,DAY_WIDTH]);

  useEffect(()=>{
    if (!chartPan) return;
    const el = chartRef.current;
    if (!el) return;
    const onMove = e => {
      const dx = e.clientX - chartPan.startX;
      const dy = e.clientY - chartPan.startY;
      el.scrollLeft = chartPan.scrollLeft - dx;
      el.scrollTop = chartPan.scrollTop - dy;
    };
    const onUp = () => setChartPan(null);
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
    return () => { window.removeEventListener("mousemove", onMove); window.removeEventListener("mouseup", onUp); };
  }, [chartPan]);

  // Column resize for the name column in left panel
  useEffect(()=>{
    if (!colResizing) return;
    const onMove = e => {
      const dx = e.clientX - colResizing.startX;
      const newW = Math.max(80, Math.min(600, colResizing.startWidth + dx));
      setNameColWidth(newW);
    };
    const onUp = () => setColResizing(null);
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
    return () => { window.removeEventListener("mousemove", onMove); window.removeEventListener("mouseup", onUp); };
  }, [colResizing]);

  useEffect(()=>{
    const el = chartRef.current;
    if (!el) return;
    const handler = e => {
      if (e.ctrlKey || e.metaKey) return; // app-zoom handler takes priority
      if (e.deltaMode === 0 && e.deltaY !== 0 && Math.abs(e.deltaY) < 4) return;
      if (Math.abs(e.deltaY) > Math.abs(e.deltaX)) {
        e.preventDefault();
        e.stopPropagation();
        const z = zoomStateRef.current;
        const presets = [null, 2, 10, 28];
        const fitDW = Math.max(0.15, (el.clientWidth || 800) / (z.dataRangeDays || 100));
        const curDW = z.isFit ? fitDW : (z.customDayWidth !== null ? z.customDayWidth : (z.zoom?.dw || 10));

        // Use zoomIdx as the single source of truth for current level.
        let curLevel = z.zoom?.key === "fit" ? 0
          : z.zoom?.key === "year" ? 1
          : z.zoom?.key === "month" ? 2 : 3;

        const dir = e.deltaY < 0 ? 1 : -1;
        const newLevel = Math.max(0, Math.min(3, curLevel + dir));

        // Decide newDW and newCustomDW to apply
        let newDW = curDW;
        let newZoomIdx = curLevel;
        let newCustomDW = z.customDayWidth;
        if (newLevel !== curLevel) {
          if (newLevel === 0) { newDW = fitDW; newCustomDW = null; }
          else { newDW = presets[newLevel]; newCustomDW = presets[newLevel]; }
          newZoomIdx = newLevel;
        } else {
          if (curLevel === 0) {
            if (dir > 0) {
              newDW = presets[1]; newCustomDW = presets[1]; newZoomIdx = 1;
            }
          } else if (curLevel === 3) {
            if (dir > 0) {
              newDW = Math.min(50, curDW * 1.2); newCustomDW = newDW;
            } else {
              newDW = presets[2]; newCustomDW = presets[2]; newZoomIdx = 2;
            }
          }
        }

        // Apply state changes
        if (newZoomIdx !== curLevel) setZoomIdx(newZoomIdx);
        setCustomDayWidth(newCustomDW);

        // CRITICAL: synchronously update zoomStateRef so rapid wheel events
        // see the latest state instead of the stale closure
        zoomStateRef.current = {
          isFit: newZoomIdx === 0,
          customDayWidth: newCustomDW,
          zoom: ZOOM_PRESETS[newZoomIdx],
          dataRangeDays: z.dataRangeDays,
        };

        // After zoom, scroll to keep work centered in viewport.
        // Different modes have different effectiveViewStart (fit=dataRange.minD, non-fit=extendedRange.minD)
        // so we compute the target scroll based on the new mode's reference point.
        if (newDW !== curDW) {
          skipAutoScrollRef.current = true;
          // Use double rAF to ensure DOM has the new totalW after React re-render
          requestAnimationFrame(() => requestAnimationFrame(() => {
            const c = chartRef.current;
            if (!c) return;
            const newIsFit = newZoomIdx === 0;
            // The new effective viewStart depends on whether we're in fit mode now
            const newViewStart = newIsFit ? dataRange.minD : extendedRange.minD;
            // Data center day relative to the new viewStart
            const dataMidTime = (dataRange.minD.getTime() + dataRange.maxD.getTime()) / 2;
            const dataCenterDay = (dataMidTime - newViewStart.getTime()) / 86400000;
            c.scrollLeft = Math.max(0, dataCenterDay * newDW - c.clientWidth / 2);
          }));
        }
      }
    };
    el.addEventListener("wheel", handler, {passive:false});

    // ─── Pinch-to-zoom for mobile (matches desktop wheel-zoom logic) ───
    let pinchState = null; // { initialDist, initialDW, initialScrollLeft, initialCenterX, centerDayOffset }

    const getDist = (touches) => {
      const dx = touches[0].clientX - touches[1].clientX;
      const dy = touches[0].clientY - touches[1].clientY;
      return Math.sqrt(dx*dx + dy*dy);
    };
    const getCenter = (touches) => ({
      x: (touches[0].clientX + touches[1].clientX) / 2,
      y: (touches[0].clientY + touches[1].clientY) / 2,
    });

    const onTouchStart = e => {
      if (e.touches.length === 2) {
        const z = zoomStateRef.current;
        const fitDW = Math.max(0.15, (el.clientWidth || 800) / (z.dataRangeDays || 100));
        const curDW = z.isFit ? fitDW : (z.customDayWidth !== null ? z.customDayWidth : (z.zoom?.dw || 10));
        const center = getCenter(e.touches);
        const rect = el.getBoundingClientRect();
        // Day offset (in days) from viewStart that is currently under the pinch center
        const centerDayOffset = (center.x - rect.left + el.scrollLeft) / curDW;
        pinchState = {
          initialDist: getDist(e.touches),
          initialDW: curDW,
          initialFitDW: fitDW,
          initialCenterX: center.x - rect.left,
          centerDayOffset,
        };
      }
    };

    const onTouchMove = e => {
      if (e.touches.length === 2 && pinchState) {
        e.preventDefault();
        const newDist = getDist(e.touches);
        const rawScale = newDist / pinchState.initialDist;
        // Boost sensitivity: amplify scale non-linearly
        // Exponent of 2.5 means a 1.5x finger spread becomes ~2.76x zoom
        const SENSITIVITY = 2.5;
        const scale = Math.pow(rawScale, SENSITIVITY);
        // Compute new dayWidth, clamped
        const fitDW = pinchState.initialFitDW;
        let newDW = pinchState.initialDW * scale;
        newDW = Math.max(fitDW * 0.95, Math.min(50, newDW));
        // Determine which preset level this corresponds to
        let newLevel;
        if (newDW <= fitDW * 1.2) newLevel = 0;
        else if (newDW < 5) newLevel = 1;
        else if (newDW < 18) newLevel = 2;
        else newLevel = 3;
        // Apply
        if (newLevel === 0) {
          setZoomIdx(0);
          setCustomDayWidth(null);
        } else {
          setZoomIdx(newLevel);
          setCustomDayWidth(newDW);
        }
        // Maintain center: keep the same day under finger center
        // After zoom, scrollLeft should be: centerDayOffset * newDW - centerXInElement
        const targetScrollLeft = pinchState.centerDayOffset * newDW - pinchState.initialCenterX;
        // Use rAF to apply after state updates
        requestAnimationFrame(() => {
          if (el) el.scrollLeft = Math.max(0, targetScrollLeft);
        });
      }
    };

    const onTouchEnd = e => {
      if (e.touches.length < 2) pinchState = null;
    };

    el.addEventListener("touchstart", onTouchStart, {passive:true});
    el.addEventListener("touchmove", onTouchMove, {passive:false});
    el.addEventListener("touchend", onTouchEnd, {passive:true});
    el.addEventListener("touchcancel", onTouchEnd, {passive:true});

    return () => {
      el.removeEventListener("wheel", handler);
      el.removeEventListener("touchstart", onTouchStart);
      el.removeEventListener("touchmove", onTouchMove);
      el.removeEventListener("touchend", onTouchEnd);
      el.removeEventListener("touchcancel", onTouchEnd);
    };
  });

  // ─── Row drag (works for both mouse and touch) ───
  useEffect(()=>{
    if (!rowDragState) return;
    const getCoords = e => {
      const t = e.touches?.[0] || e.changedTouches?.[0] || e;
      return { x: t.clientX, y: t.clientY };
    };
    const onMove = e => {
      if (e.cancelable && e.touches) e.preventDefault();
      const {x, y} = getCoords(e);
      setRowDragPos({x, y});
      const listEl = leftPanelListRef.current;
      if (!listEl) return;
      const rect = listEl.getBoundingClientRect();
      const relY = y - rect.top + listEl.scrollTop;
      let accumulated = 0;
      let dropIdx = 0;
      let dropPos = "below";
      for (let i = 0; i < flatRows.length; i++) {
        const rh = getRowHeight(i);
        if (relY < accumulated + rh / 2) {
          dropIdx = i; dropPos = "above"; break;
        } else if (relY < accumulated + rh) {
          dropIdx = i; dropPos = "below"; break;
        }
        accumulated += rh;
        dropIdx = i; dropPos = "below";
      }
      // Auto-scroll on edge
      const edge = 60;
      if (y - rect.top < edge) listEl.scrollTop -= 8;
      else if (rect.bottom - y < edge) listEl.scrollTop += 8;

      setRowDragState(prev => ({...prev, dropIdx, dropPos}));
    };
    const onUp = e => {
      if (rowDragState.dropIdx !== undefined) {
        const draggedRow = flatRows.find(r => r.item.id === rowDragState.id);
        const targetRow = flatRows[rowDragState.dropIdx];
        if (draggedRow && targetRow && draggedRow.item.id !== targetRow.item.id) {
          setItems(prev => {
            const dragged = draggedRow.item;
            const target = targetRow.item;
            if (dragged.type === "sector") {
              if (target.type !== "sector") return prev;
              const sectors = prev.filter(i=>i.type==="sector").slice().sort((a,b)=>(a.order??0)-(b.order??0));
              const si = sectors.findIndex(s=>s.id===dragged.id);
              const ti = sectors.findIndex(s=>s.id===target.id);
              if (si<0||ti<0) return prev;
              const r=[...sectors]; r.splice(si,1);
              const insertAt = rowDragState.dropPos==="above" ? (ti>si?ti-1:ti) : (ti>=si?ti:ti+1);
              r.splice(insertAt,0,dragged);
              const map=Object.fromEntries(r.map((s,i)=>[s.id,i]));
              return prev.map(i=>i.type==="sector"?{...i,order:map[i.id]??i.order}:i);
            }
            const fullFlat = buildFlat(prev);
            const realFrom = fullFlat.findIndex(it => it.id === dragged.id);
            const realTargetIdx = fullFlat.findIndex(it => it.id === target.id);
            if (realFrom < 0 || realTargetIdx < 0) return prev;
            let realTo = realTargetIdx;
            if (rowDragState.dropPos === "below") realTo += 1;
            return applyFlatMove(prev, dragged.id, realFrom, realTo);
          });
        }
      }
      setRowDragState(null);
    };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
    window.addEventListener("touchmove", onMove, {passive:false});
    window.addEventListener("touchend", onUp);
    window.addEventListener("touchcancel", onUp);
    return () => {
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
      window.removeEventListener("touchmove", onMove);
      window.removeEventListener("touchend", onUp);
      window.removeEventListener("touchcancel", onUp);
    };
  }, [rowDragState, flatRows, getRowHeight]);

  const flatRowsForDrag = useMemo(()=>{
    const rows=[];
    items.filter(i=>i.type==="sector").slice().sort((a,b)=>(a.order??0)-(b.order??0))
      .forEach(sector=>{
        rows.push(sector);
        if (!sector.collapsed)
          items.filter(i=>i.parentId===sector.id).slice().sort((a,b)=>(a.order??0)-(b.order??0))
            .forEach(child=>rows.push(child));
      });
    return rows;
  },[items]);

  const buildFlat = (allItems) => {
    const flat = [];
    allItems.filter(i=>i.type==="sector").slice()
      .sort((a,b)=>(a.order??0)-(b.order??0))
      .forEach(sector => {
        flat.push(sector);
        allItems.filter(i=>i.parentId===sector.id).slice()
          .sort((a,b)=>(a.order??0)-(b.order??0))
          .forEach(child => flat.push(child));
      });
    return flat;
  };

  const applyFlatMove = (allItems, draggedId, fromIdx, toIdx) => {
    if (fromIdx === toIdx) return allItems;
    const flat = buildFlat(allItems);
    if (fromIdx < 0 || toIdx < 0 || fromIdx >= flat.length || toIdx > flat.length) return allItems;
    const dragged = flat[fromIdx];
    if (!dragged) return allItems;
    const newFlat = flat.filter((_, i) => i !== fromIdx);
    const insertAt = toIdx > fromIdx ? toIdx - 1 : toIdx;
    newFlat.splice(insertAt, 0, dragged);
    const getParent = (idx) => {
      for (let i = idx - 1; i >= 0; i--) {
        if (newFlat[i].type === "sector") return newFlat[i].id;
      }
      return null;
    };
    const newOrders = new Map();
    let sectorIdx = 0;
    const childCounters = {};
    newFlat.forEach((item, i) => {
      if (item.type === "sector") {
        newOrders.set(item.id, { parentId: undefined, order: sectorIdx++ });
        childCounters[item.id] = 0;
      } else {
        const pid = getParent(i);
        if (pid === null) return;
        if (childCounters[pid] === undefined) childCounters[pid] = 0;
        newOrders.set(item.id, { parentId: pid, order: childCounters[pid]++ });
      }
    });
    return allItems.map(item => {
      const update = newOrders.get(item.id);
      if (!update) return item;
      if (item.type === "sector") return { ...item, order: update.order };
      return { ...item, parentId: update.parentId, order: update.order };
    });
  };

  // Changes zoom level and re-scrolls chart to data area
  // (so users always land on actual work, not in empty calendar years)
  const changeZoom = useCallback((idx) => {
    skipAutoScrollRef.current = true;
    setZoomIdx(idx);
    setCustomDayWidth(null);
    // Use double rAF to ensure render completes first
    requestAnimationFrame(() => requestAnimationFrame(() => {
      const c = chartRef.current;
      if (!c) return;
      const newPreset = ZOOM_PRESETS[idx];
      const newIsFit = newPreset.key === "fit";
      const fitDW = Math.max(0.15, (c.clientWidth || 800) / dataRangeDays);
      const newDW = newIsFit ? fitDW : newPreset.dw;
      // In fit mode, viewStart = dataRange.minD so scrollLeft = 0
      // In other modes, viewStart = extendedRange.minD so offset to data
      if (newIsFit) {
        c.scrollLeft = 0;
      } else {
        // Center data in viewport
        const dataMidTime = (dataRange.minD.getTime() + dataRange.maxD.getTime()) / 2;
        const dataCenterDay = (dataMidTime - extendedRange.minD.getTime()) / 86400000;
        c.scrollLeft = Math.max(0, dataCenterDay * newDW - c.clientWidth / 2);
      }
    }));
  }, [dataRangeDays, extendedRange.minD, dataRange.minD, dataRange.maxD]);

  const moveStep = useCallback((id, dir) => {
    setItems(prev => {
      const flat = buildFlat(prev);
      const fromIdx = flat.findIndex(i => i.id === id);
      if (fromIdx < 0) return prev;
      const dragged = flat[fromIdx];
      if (dragged.type === "sector") {
        const sectors = prev.filter(i=>i.type==="sector").slice().sort((a,b)=>(a.order??0)-(b.order??0));
        const si = sectors.findIndex(s=>s.id===id);
        const ni = si + dir;
        if (ni < 0 || ni >= sectors.length) return prev;
        const r = [...sectors]; r.splice(si,1); r.splice(ni,0,dragged);
        const map = Object.fromEntries(r.map((s,i)=>[s.id,i]));
        return prev.map(i=>i.type==="sector"?{...i,order:map[i.id]??i.order}:i);
      }
      const toIdx = dir > 0 ? fromIdx + 2 : fromIdx + dir;
      if (toIdx < 0 || toIdx > flat.length) return prev;
      return applyFlatMove(prev, id, fromIdx, toIdx);
    });
  },[]);

  const handleRowTapWhileDragging = useCallback((targetId) => {
    if (!activeDragId || activeDragId === targetId) return;
    setItems(prev => {
      const flat = buildFlat(prev);
      const fromIdx = flat.findIndex(i=>i.id===activeDragId);
      const toIdx = flat.findIndex(i=>i.id===targetId);
      if (fromIdx < 0 || toIdx < 0) return prev;
      const dragged = flat[fromIdx];
      if (dragged.type === "sector") {
        const target = flat[toIdx];
        if (target.type !== "sector") return prev;
        const sectors = prev.filter(i=>i.type==="sector").slice().sort((a,b)=>(a.order??0)-(b.order??0));
        const si = sectors.findIndex(s=>s.id===activeDragId);
        const ti = sectors.findIndex(s=>s.id===targetId);
        if (si<0||ti<0) return prev;
        const r=[...sectors]; r.splice(si,1); r.splice(ti,0,dragged);
        const map=Object.fromEntries(r.map((s,i)=>[s.id,i]));
        return prev.map(i=>i.type==="sector"?{...i,order:map[i.id]??i.order}:i);
      }
      return applyFlatMove(prev, activeDragId, fromIdx, toIdx);
    });
    setActiveDragId(null);
  },[activeDragId]);

  useEffect(()=>{
    const onKey=e=>{
      if(e.key==="Escape"){setLinkingFrom(null);setActiveDragId(null);setSelectedId(null);return;}
      if((e.metaKey||e.ctrlKey)&&e.key.toLowerCase()==="z"&&!e.shiftKey){
        e.preventDefault(); undo(); return;
      }
      if(((e.metaKey||e.ctrlKey)&&e.key.toLowerCase()==="z"&&e.shiftKey)||((e.metaKey||e.ctrlKey)&&e.key.toLowerCase()==="y")){
        e.preventDefault(); redo(); return;
      }
    };
    window.addEventListener("keydown",onKey);
    return ()=>window.removeEventListener("keydown",onKey);
  },[undo, redo]);

  const assignOrder=(parsed,all)=>{
    if (parsed.type==="sector") {
      const max=Math.max(-1,...all.filter(i=>i.type==="sector").map(i=>i.order??0));
      return {...parsed,order:max+1};
    }
    const sibs=all.filter(i=>i.parentId===parsed.parentId);
    const max=Math.max(-1,...sibs.map(i=>i.order??0));
    return {...parsed,order:max+1};
  };

  const openAdd=(type="task",parentId=null)=>{
    const def=parentId||items.find(i=>i.type==="sector")?.id;
    const siblings = items.filter(i=>i.parentId===def && i.end && i.end instanceof Date && !isNaN(i.end.getTime()))
      .sort((a,b)=>b.end.getTime()-a.end.getTime());
    const autoStart = siblings.length > 0 ? addDays(siblings[0].end, 1) : today;
    setEditItem(null);
    setForm({type,parentId:def,name:"",start:toISO(autoStart),end:toISO(addDays(autoStart,7)),progress:0,link:"",deps:[],assignees:[]});
    setShowModal(true); setMenuOpen(false); setFormError("");
  };

  const openAddSector=()=>{
    setEditItem(null);
    const used=items.filter(i=>i.type==="sector").map(i=>i.color);
    const next=SECTOR_COLORS.find(c=>!used.includes(c))||SECTOR_COLORS[0];
    let firstCat = items.find(i=>i.type==="category");
    if (!firstCat) {
      // Auto-create default category if none exists
      const defaultCat = { id: uid(), type: "category", name: "기본", collapsed: false, order: 0 };
      setItems(prev => [...prev, defaultCat]);
      firstCat = defaultCat;
    }
    setForm({type:"sector",name:"",color:next,collapsed:false,parentId: firstCat.id});
    setShowModal(true); setMenuOpen(false); setFormError("");
  };

  const openEdit=item=>{
    setEditItem(item);
    setForm({...item,start:item.start?toISO(item.start):"",end:item.end?toISO(item.end):"",assignees:item.assignees||[]});
    setShowModal(true); setFormError("");
  };

  const saveForm=()=>{
    if (!form.name || !form.name.trim()) {
      setFormError("이름을 입력해주세요.");
      return;
    }
    // Validate dates for tasks/milestones
    if (form.type !== "sector") {
      const sd = parseDate(form.start);
      if (!sd) {
        setFormError("시작일이 올바르지 않습니다.");
        return;
      }
      if (form.type === "task") {
        const ed = parseDate(form.end);
        if (!ed) {
          setFormError("종료일이 올바르지 않습니다.");
          return;
        }
        if (ed < sd) {
          setFormError("종료일은 시작일 이후여야 합니다.");
          return;
        }
      }
    }
    setFormError("");
    let parsed;
    if (form.type==="sector") { const{_depSearch:__ds,...rest}=form; parsed={...rest}; }
    else { const{color:_c,_depSearch:__ds,...rest}=form; parsed={...rest,start:parseDate(form.start),end:parseDate(form.type==="milestone"?form.start:form.end),progress:form.type==="task"?(form.progress||0):0,deps:form.deps||[],assignees:Array.isArray(form.assignees)?form.assignees:[]}; }

    // Detect newly-added assignees (vs. what was on the original item) so we
    // can ask the user about email notifications after the save commits.
    // Only relevant for non-sector items.
    let newlyAdded = [];
    if (parsed.type !== "sector") {
      const before = (editItem?.assignees || []);
      const beforeIds = new Set(before.map(a => a.id || a.email));
      newlyAdded = (parsed.assignees || []).filter(a => !beforeIds.has(a.id || a.email));
    }

    if (editItem) {
      setItems(prev=>prev.map(i=>i.id===editItem.id?{...parsed,id:i.id}:i));
    } else {
      setItems(prev=>{
        const o=assignOrder({...parsed,id:uid()},prev);
        const next = [...prev, o];
        // Auto-expand parent sector + its category so the new item is visible
        if (parsed.type !== "sector" && parsed.parentId) {
          const parentSector = next.find(s => s.id === parsed.parentId);
          const parentCatId = parentSector?.parentId;
          // Also clear kanban-side collapse for these
          setKanbanCollapsed(c => ({...c, [parsed.parentId]: false}));
          if (parentCatId) setKanbanCatCollapsed(c => ({...c, [parentCatId]: false}));
          return next.map(i => {
            if (i.id === parsed.parentId && i.type === "sector") return {...i, collapsed: false};
            if (parentCatId && i.id === parentCatId && i.type === "category") return {...i, collapsed: false};
            return i;
          });
        }
        return next;
      });
    }
    setShowModal(false);
    // Surface the email-notification popup AFTER the save UI closes so the
    // dialog stack reads naturally: edit → saved → "want to email them?".
    if (newlyAdded.length > 0) {
      setAssigneeNotifyDialog({ task: parsed, newAssignees: newlyAdded });
    }
  };

  const deleteItem=id=>setItems(prev=>{
    const isSec=prev.find(i=>i.id===id&&i.type==="sector");
    if(isSec) return prev.filter(i=>i.id!==id&&i.parentId!==id);
    return prev.filter(i=>i.id!==id).map(i=>({...i,deps:(i.deps||[]).filter(d=>d!==id)}));
  });

  const toggleCollapse=id=>setItems(prev=>prev.map(i=>i.id===id?{...i,collapsed:!i.collapsed}:i));

  // Category helpers
  const openAddCategory = () => {
    setCatForm({ name:"", sectorIds:[] });
    setShowCatModal(true);
  };
  const openEditCategory = (id) => {
    const cat = items.find(i=>i.id===id);
    if (!cat) return;
    const sectorIds = items.filter(i => i.type==="sector" && i.parentId===id).map(s=>s.id);
    setCatForm({ id, name: cat.name, sectorIds });
    setShowCatModal(true);
  };
  const saveCategoryForm = () => {
    if (!catForm || !catForm.name.trim()) return;
    setItems(prev => {
      if (catForm.id) {
        // Edit existing: update name + reassign sectors
        return prev.map(i => {
          if (i.id === catForm.id) return { ...i, name: catForm.name.trim() };
          // Sectors: assign to this category if in sectorIds, otherwise leave their current parent if it's not this category
          if (i.type === "sector") {
            if (catForm.sectorIds.includes(i.id)) return { ...i, parentId: catForm.id };
            // If sector was in this category but unchecked, move to first other category
            if (i.parentId === catForm.id) {
              const otherCat = prev.find(c => c.type === "category" && c.id !== catForm.id);
              if (otherCat) return { ...i, parentId: otherCat.id };
            }
          }
          return i;
        });
      } else {
        // Create new category
        const maxOrder = Math.max(-1, ...prev.filter(i=>i.type==="category").map(i=>i.order??0));
        const newCat = { id:uid(), type:"category", name:catForm.name.trim(), collapsed:false, order:maxOrder+1 };
        return [...prev, newCat].map(i => {
          if (i.type === "sector" && catForm.sectorIds.includes(i.id)) {
            return { ...i, parentId: newCat.id };
          }
          return i;
        });
      }
    });
    setShowCatModal(false);
  };
  const renameCategory = (id) => openEditCategory(id);
  const deleteCategory = (id) => {
    const cat = items.find(i=>i.id===id);
    if (!cat) return;
    // Don't allow deleting last category
    const allCats = items.filter(i=>i.type==="category");
    if (allCats.length <= 1) {
      alert("최소 1개의 대분류는 유지해야 합니다.");
      return;
    }
    const sectorsInCat = items.filter(i=>i.type==="sector" && i.parentId===id);
    const targetCat = allCats.find(c=>c.id!==id);
    const doDelete = () => {
      setItems(prev => prev
        .filter(i => i.id !== id)
        .map(i => i.type==="sector" && i.parentId===id ? {...i, parentId:targetCat.id} : i)
      );
    };
    if (window.appConfirm) {
      window.appConfirm(`"${cat.name}" 대분류를 삭제하시겠습니까?\n속한 ${sectorsInCat.length}개 섹터는 "${targetCat.name}"으로 이동됩니다.`, doDelete);
    } else {
      let ok = false;
      try { ok = window.confirm(`"${cat.name}" 대분류를 삭제하시겠습니까?\n속한 ${sectorsInCat.length}개 섹터는 "${targetCat.name}"으로 이동됩니다.`); }
      catch(e) { ok = true; }
      if (ok) doDelete();
    }
  };
  const moveCategory = (id, dir) => {
    setItems(prev => {
      const cats = prev.filter(i=>i.type==="category").slice().sort((a,b)=>(a.order??0)-(b.order??0));
      const idx = cats.findIndex(c=>c.id===id);
      if (idx < 0) return prev;
      const ni = idx + dir;
      if (ni < 0 || ni >= cats.length) return prev;
      const r = [...cats]; const [removed] = r.splice(idx,1); r.splice(ni,0,removed);
      const map = Object.fromEntries(r.map((c,i)=>[c.id,i]));
      return prev.map(i => i.type==="category" ? {...i, order: map[i.id] ?? i.order} : i);
    });
  };

  const handleLinkClick=(e,itemId)=>{
    e.stopPropagation();
    if (!linkingFrom){setLinkingFrom(itemId);return;}
    if (linkingFrom===itemId){setLinkingFrom(null);return;}
    setItems(prev=>prev.map(i=>{ if(i.id!==itemId) return i; const deps=i.deps||[]; return deps.includes(linkingFrom)?{...i,deps:deps.filter(d=>d!==linkingFrom)}:{...i,deps:[...deps,linkingFrom]}; }));
    setLinkingFrom(null);
  };

  const importFromCSV = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        let text = e.target.result;
        if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
        const lines = text.split("\n").map(l => l.trimEnd()).filter(l => l.length > 0);
        if (lines.length < 2) return;
        const DQUOTE = String.fromCharCode(34);
        const COMMA = String.fromCharCode(44);
        const parseRow = (line) => {
          const cols = []; let cur = ""; let inQ = false;
          for (let i = 0; i < line.length; i++) {
            const ch = line[i];
            if (ch === DQUOTE) {
              if (inQ && line[i+1] === DQUOTE) { cur += DQUOTE; i++; }
              else { inQ = !inQ; }
            } else if (ch === COMMA && !inQ) { cols.push(cur); cur = ""; }
            else { cur += ch; }
          }
          cols.push(cur);
          return cols;
        };
        // Auto-detect format: if first column header is "대분류", use new format
        const headerRow = parseRow(lines[0]);
        const hasCategory = (headerRow[0] || "").trim() === "대분류";
        const rows = lines.slice(1).map(parseRow);
        const categoryMap = {}; // name -> id
        const sectorMap = {};   // "catName::sectorName" -> id (scoped per category)
        let maxId = Math.max(...items.map(i => i.id), 100);
        const newItems = [];
        // Always have at least one default category in new structure
        const ensureDefaultCategory = () => {
          if (!categoryMap["기본"]) {
            const cId = ++maxId;
            categoryMap["기본"] = cId;
            newItems.push({ id: cId, type: "category", name: "기본", collapsed: false, order: 0 });
          }
          return categoryMap["기본"];
        };
        rows.forEach(row => {
          let categoryName, sectorName, type, name, startStr, endStr, progress, link, depsStr;
          if (hasCategory) {
            categoryName = (row[0] || "").trim() || "기본";
            sectorName = (row[1] || "").trim();
            type = (row[2] || "").trim();
            name = (row[3] || "").trim();
            startStr = (row[4] || "").trim();
            endStr = (row[5] || "").trim();
            progress = (row[7] || "").trim(); // skip 일수 (col 6)
            link = (row[8] || "").trim();
            depsStr = (row[9] || "").trim();
          } else {
            categoryName = "기본";
            sectorName = (row[0] || "").trim();
            type = (row[1] || "").trim();
            name = (row[2] || "").trim();
            startStr = (row[3] || "").trim();
            endStr = (row[4] || "").trim();
            progress = (row[6] || row[5] || "").trim();
            link = (row[7] || row[6] || "").trim();
            depsStr = (row[8] || row[7] || "").trim();
          }
          if (!sectorName || !name) return;
          // Ensure category exists
          if (!categoryMap[categoryName]) {
            const cId = ++maxId;
            categoryMap[categoryName] = cId;
            const catCount = Object.keys(categoryMap).length;
            newItems.push({ id: cId, type: "category", name: categoryName, collapsed: false, order: catCount - 1 });
          }
          const catId = categoryMap[categoryName];
          // Ensure sector exists (scoped per category)
          const sectorKey = catId + "::" + sectorName;
          if (!sectorMap[sectorKey]) {
            const sId = ++maxId;
            sectorMap[sectorKey] = sId;
            const usedColors = newItems.filter(i => i.type === "sector").map(i => i.color);
            const color = SECTOR_COLORS.find(c => !usedColors.includes(c)) || SECTOR_COLORS[0];
            newItems.push({ id: sId, type: "sector", parentId: catId, name: sectorName, color, collapsed: false, order: newItems.filter(i => i.type === "sector" && i.parentId === catId).length });
          }
          const itemId = ++maxId;
          const itemType = type === "마일스톤" ? "milestone" : "task";
          newItems.push({
            id: itemId, type: itemType,
            parentId: sectorMap[sectorKey], name,
            start: startStr ? parseDate(startStr) : null,
            end: endStr ? parseDate(endStr) : (startStr ? parseDate(startStr) : null),
            progress: itemType === "task" ? (parseInt(progress) || 0) : 0,
            link, deps: [],
            order: newItems.filter(i => i.parentId === sectorMap[sectorKey]).length
          });
        });
        // Resolve dependency names → ids
        newItems.forEach(item => {
          const row = rows.find(r => {
            const idx = hasCategory ? 3 : 2;
            return (r[idx] || "").trim() === item.name;
          });
          if (!row) return;
          const depIdx = hasCategory ? 9 : (row[8] !== undefined ? 8 : 7);
          if (!row[depIdx]) return;
          const depNames = row[depIdx].split(";").map(s => s.trim()).filter(Boolean);
          item.deps = depNames.map(dn => newItems.find(i => i.name === dn)).filter(Boolean).map(i => i.id);
        });
        if (newItems.length > 0) {
          setItems(sortedByDate(newItems));
          alert("가져오기 완료! " + newItems.filter(i => i.type !== "sector" && i.type !== "category").length + "개 작업을 불러왔습니다.");
        }
      } catch(err) {
        alert("파일을 읽는 중 오류가 발생했습니다.");
        console.error(err);
      }
    };
    reader.readAsText(file, "UTF-8");
  };

  const exportToExcel=()=>{
    const categories=items.filter(i=>i.type==="category").slice().sort((a,b)=>(a.order??0)-(b.order??0));
    const rows=[["대분류","섹터","유형","작업명","시작일","종료일","일수","진행률(%)","링크","의존작업"]];
    categories.forEach(category=>{
      const sectorsInCat = items.filter(i=>i.type==="sector"&&i.parentId===category.id).slice().sort((a,b)=>(a.order??0)-(b.order??0));
      sectorsInCat.forEach(sector=>{
        items.filter(i=>i.parentId===sector.id).slice().sort((a,b)=>(a.order??0)-(b.order??0)).forEach(item=>{
          const depNames=(item.deps||[]).map(d=>items.find(x=>x.id===d)?.name||"").filter(Boolean).join(";");
          const days_count = item.start&&item.end ? Math.max(1,dateDiff(item.start,item.end)+1) : "";
          rows.push([
            category.name,
            sector.name,
            item.type==="milestone"?"마일스톤":"작업",
            item.name,
            item.start?toISO(item.start):"",
            item.end?toISO(item.end):"",
            days_count,
            item.type==="task"?item.progress:"",
            item.link||"",
            depNames
          ]);
        });
      });
    });
    // Also handle orphan sectors (no category) just in case
    const orphans = items.filter(i=>i.type==="sector"&&!categories.some(c=>c.id===i.parentId));
    orphans.forEach(sector=>{
      items.filter(i=>i.parentId===sector.id).forEach(item=>{
        const depNames=(item.deps||[]).map(d=>items.find(x=>x.id===d)?.name||"").filter(Boolean).join(";");
        const days_count = item.start&&item.end ? Math.max(1,dateDiff(item.start,item.end)+1) : "";
        rows.push([
          "(미분류)",
          sector.name,
          item.type==="milestone"?"마일스톤":"작업",
          item.name,
          item.start?toISO(item.start):"",
          item.end?toISO(item.end):"",
          days_count,
          item.type==="task"?item.progress:"",
          item.link||"",
          depNames
        ]);
      });
    });
    const csv=rows.map(r=>r.map(cell=>`"${String(cell).replace(/"/g,'""')}"`).join(",")).join("\n");
    const encoded = encodeURIComponent("\uFEFF" + csv);
    const a = document.createElement("a");
    a.href = "data:text/csv;charset=utf-8," + encoded;
    a.download = `간트차트_${toISO(today)}.csv`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
  };

  const todayX=dateDiff(effectiveViewStart,today)*DAY_WIDTH+DAY_WIDTH/2;
  const totalW=days.length*DAY_WIDTH;

  const SB = {
    nav: {padding:"6px 10px",background:"transparent",color:T.textSub,border:`1px solid ${T.border}`,borderRadius:7,cursor:"pointer",fontSize:12,fontWeight:600,display:"flex",alignItems:"center",gap:5},
    today: {padding:"6px 12px",background:"#6366f1",color:"#fff",border:"none",borderRadius:7,cursor:"pointer",fontSize:12,fontWeight:600},
    add: {padding:"7px 13px",background:"linear-gradient(135deg,#6366f1,#8b5cf6)",color:"#fff",border:"none",borderRadius:7,cursor:"pointer",fontSize:12,fontWeight:600,display:"flex",alignItems:"center",gap:5},
    sec: {padding:"7px 11px",background:"transparent",color:T.textSub,border:`1px solid ${T.border}`,borderRadius:7,cursor:"pointer",fontSize:12,fontWeight:600,display:"flex",alignItems:"center",gap:5},
    inp: {background:T.bgInput,border:`1px solid ${T.border}`,borderRadius:7,padding:"9px 11px",color:T.text,fontSize:13,outline:"none",fontFamily:"inherit"}
  };

  const renderToolbar = (vertical=false) => {
    return (
      <div style={{display:"flex",flexDirection:vertical?"column":"row",gap:vertical?6:5,alignItems:vertical?"stretch":"center",flexWrap:vertical?"nowrap":"wrap"}}>
        <button onClick={()=>setShowMilestones(v=>!v)}
          style={{...SB.sec,color:showMilestones?"#818cf8":T.textMuted,border:"1px solid "+(showMilestones?"#4338ca":T.border),justifyContent:vertical?"flex-start":"center"}}>
          <MilestoneIcon size={12} color={showMilestones?"#818cf8":T.textFaint}/>
          {"마일스톤 "+(showMilestones?"ON":"OFF")}
        </button>
        {linkingFrom&&<button onClick={()=>setLinkingFrom(null)} style={{...SB.sec,color:"#fbbf24",border:"1px solid #fbbf24",justifyContent:vertical?"flex-start":"center"}}>
          {"x 연결 취소"}
        </button>}
        {!vertical&&<div style={{width:1,height:22,background:T.border}}/>}
        <button onClick={()=>{if(chartRef.current){const chartW=chartRef.current.clientWidth;chartRef.current.scrollLeft = todayX - chartW/2;}}} style={{...SB.today,justifyContent:vertical?"flex-start":"center"}}>오늘</button>
        <div style={{display:"flex",border:`1px solid ${T.border}`,borderRadius:7,overflow:"hidden"}}>
          {ZOOM_PRESETS.map((z,i)=>(
            <button key={z.key} onClick={()=>changeZoom(i)}
              style={{padding:"6px 10px",border:"none",cursor:"pointer",fontSize:12,fontWeight:600,
                background:i===zoomIdx?"#6366f1":"transparent",
                color:i===zoomIdx?"#fff":T.textMuted,
                transition:"all 0.15s"}}>
              {z.label}
            </button>
          ))}
        </div>
        {!vertical&&<div style={{width:1,height:22,background:T.border}}/>}
        <button onClick={openAddSector} style={{...SB.add,justifyContent:vertical?"flex-start":"center"}}>
          <PlusIcon size={12} color="#fff"/>{" 섹터"}
        </button>
        <button onClick={()=>openAdd("task")} style={{...SB.add,justifyContent:vertical?"flex-start":"center"}}>
          <PlusIcon size={12} color="#fff"/>{" 작업"}
        </button>
        {!vertical&&<div style={{width:1,height:22,background:T.border}}/>}
        <button onClick={undo} disabled={undoStackRef.current.length===0}
          title="실행 취소 (Ctrl+Z)"
          style={{...SB.sec, padding:"6px 8px", opacity:undoStackRef.current.length===0?0.3:1, cursor:undoStackRef.current.length===0?"not-allowed":"pointer",justifyContent:vertical?"flex-start":"center"}}>
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M3 7v6h6"/><path d="M21 17a9 9 0 0 0-9-9 9 9 0 0 0-6.7 3L3 13"/>
          </svg>
        </button>
        <button onClick={redo} disabled={redoStackRef.current.length===0}
          title="다시 실행 (Ctrl+Shift+Z)"
          style={{...SB.sec, padding:"6px 8px", opacity:redoStackRef.current.length===0?0.3:1, cursor:redoStackRef.current.length===0?"not-allowed":"pointer",justifyContent:vertical?"flex-start":"center"}}>
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M21 7v6h-6"/><path d="M3 17a9 9 0 0 1 9-9 9 9 0 0 1 6.7 3L21 13"/>
          </svg>
        </button>
        <button onClick={()=>setThemeKey(k=>k==="dark"?"light":"dark")}
          title={themeKey==="dark"?"라이트 모드":"다크 모드"}
          style={{...SB.sec, padding:"6px 8px",justifyContent:vertical?"flex-start":"center"}}>
          {themeKey==="dark"
            ? <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <circle cx="12" cy="12" r="5"/><line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/>
                <line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/>
                <line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/>
                <line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/>
              </svg>
            : <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/>
              </svg>
          }
        </button>
        {!vertical&&<div style={{width:1,height:22,background:T.border}}/>}
        <input ref={importInputRef} type="file" accept=".csv" style={{display:"none"}}
          onChange={e=>{ if(e.target.files[0]){ importFromCSV(e.target.files[0]); e.target.value=""; } }}/>
        <button onClick={exportToExcel}
          style={{...SB.sec,color:"#22c55e",border:"1px solid #166534",justifyContent:vertical?"flex-start":"center"}}
          onMouseEnter={e=>e.currentTarget.style.background="rgba(34,197,94,0.1)"}
          onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
          <ExcelIcon size={13}/>{" 내보내기"}
        </button>
        <button onClick={()=>importInputRef.current?.click()}
          style={{...SB.sec,color:"#22c55e",border:"1px solid #166534",justifyContent:vertical?"flex-start":"center"}}
          onMouseEnter={e=>e.currentTarget.style.background="rgba(34,197,94,0.1)"}
          onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
            <polyline points="14 2 14 8 20 8"/>
            <line x1="12" y1="15" x2="12" y2="9"/>
            <polyline points="9 12 12 9 15 12"/>
          </svg>{" 가져오기"}
        </button>
      </div>
    );
  };

  // ─── List View Header (mobile) ───
  const renderListHeader = () => (
    <div style={{flexShrink:0,borderBottom:`1px solid ${T.border}`,background:T.bgHeader}}>
      <div style={{padding:"10px 12px",display:"flex",alignItems:"center",gap:8}}>
        <div style={{display:"flex",alignItems:"center",flex:1,gap:6,background:T.bgInput,border:`1px solid ${T.border}`,borderRadius:8,padding:"7px 10px"}}>
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke={T.textFaint} strokeWidth="2"><circle cx="11" cy="11" r="7"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
          <input value={searchQuery.trim()===""?"":searchQuery} onChange={e=>setSearchQuery(e.target.value)}
            placeholder={`작업/섹터 검색 (${flatRows.length}개)`}
            style={{flex:1,background:"transparent",border:"none",outline:"none",color:T.text,fontSize:13,fontFamily:"inherit",minWidth:0}}/>
          {searchQuery && searchQuery.trim() && (
            <span onClick={()=>setSearchQuery("")} style={{cursor:"pointer",color:T.textFaint,fontSize:16,padding:"0 4px",lineHeight:1}}>×</span>
          )}
        </div>
        <button onClick={()=>setShowMilestones(v=>!v)}
          style={{flexShrink:0,background:"transparent",border:`1px solid ${showMilestones?"#4338ca":T.border}`,borderRadius:7,padding:"7px 9px",cursor:"pointer",color:showMilestones?"#818cf8":T.textMuted,display:"flex",alignItems:"center",gap:4,fontSize:11,fontWeight:600}}
          title={showMilestones?"마일스톤 숨기기":"마일스톤 보기"}>
          <MilestoneIcon size={11} color={showMilestones?"#818cf8":T.textFaint}/>
          {showMilestones?"ON":"OFF"}
        </button>
      </div>
      <div style={{display:"grid",gridTemplateColumns:GRID_COLS_MOBILE,alignItems:"center",height:30,padding:"0 10px 0 6px",gap:6,borderTop:`1px solid ${T.borderLight}`}}>
        <div/>
        <div style={{fontSize:10,fontWeight:600,color:T.textFaint,paddingLeft:14}}>작업명</div>
        <div style={{fontSize:10,fontWeight:600,color:T.textFaint,textAlign:"center"}}>시작</div>
        <div style={{fontSize:10,fontWeight:600,color:T.textFaint,textAlign:"center"}}>종료</div>
        <div style={{fontSize:10,fontWeight:600,color:T.textFaint,textAlign:"center"}}>진행</div>
      </div>
    </div>
  );

  // ─── Main Render ───
  return (
    <div style={{fontFamily:"Pretendard, Noto Sans KR, sans-serif",background:T.bg,height:"100vh",color:T.text,display:"flex",flexDirection:"column",overflow:"hidden",
      cursor:rowDragState?"grabbing":(chartPan?"grabbing":"auto"), userSelect:rowDragState?"none":"auto",
      ...(appZoom !== 1.0 ? {zoom: appZoom} : {})}}>
      <style>{`
        .hide-scrollbar::-webkit-scrollbar{display:none;width:0;height:0;}
        @keyframes slideUp { from { transform: translateY(100%); } to { transform: translateY(0); } }
        @keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }
      `}</style>

      {updateAvailable && (
        <div style={{
          background:"linear-gradient(90deg,#6366f1,#8b5cf6)",
          color:"#fff", padding: isMobile?"8px 12px":"10px 20px",
          display:"flex", alignItems:"center", justifyContent:"space-between",
          flexShrink:0, zIndex:200, gap:8,
          boxShadow:"0 2px 12px rgba(99,102,241,0.4)"
        }}>
          <div style={{display:"flex",alignItems:"center",gap:8,minWidth:0,flex:1}}>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{flexShrink:0}}>
              <polyline points="23 4 23 10 17 10"/><polyline points="1 20 1 14 7 14"/>
              <path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/>
            </svg>
            <span style={{fontSize:isMobile?12:13,fontWeight:600,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>
              {isMobile ? "새 버전 있음" : "새 버전이 있어요! 새로고침하면 최신 버전으로 업데이트됩니다."}
            </span>
          </div>
          <div style={{display:"flex",gap:6,flexShrink:0}}>
            <button onClick={()=>window.location.reload()}
              style={{padding:isMobile?"5px 10px":"6px 16px",background:"#fff",color:"#6366f1",border:"none",
                borderRadius:7,cursor:"pointer",fontSize:isMobile?12:13,fontWeight:700,
                boxShadow:"0 2px 8px rgba(0,0,0,0.15)",whiteSpace:"nowrap"}}>
              새로고침
            </button>
            <button onClick={()=>setUpdateAvailable(false)}
              style={{padding:isMobile?"5px 8px":"6px 10px",background:"rgba(255,255,255,0.2)",color:"#fff",
                border:"1px solid rgba(255,255,255,0.3)",borderRadius:7,cursor:"pointer",fontSize:isMobile?12:13,fontWeight:600}}>
              ×
            </button>
          </div>
        </div>
      )}

      {/* TOP BAR */}
      <div style={{padding:isMobile?"8px 10px":"12px 18px",
        background:themeKey==="dark"?"linear-gradient(180deg,#0d111a,#0f1117)":"linear-gradient(180deg,#f0f1f5,#ffffff)",
        borderBottom:`1px solid ${T.border}`,display:"flex",alignItems:"center",gap:isMobile?8:12,flexShrink:0,position:"relative",zIndex:80}}>
        <div style={{display:"flex",alignItems:"center",gap:isMobile?8:10,minWidth:0,flex:isMobile?1:"unset"}}>
          <div style={{width:isMobile?28:30,height:isMobile?28:30,background:"linear-gradient(135deg,#6366f1,#8b5cf6)",borderRadius:7,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,boxShadow:"0 2px 8px rgba(99,102,241,0.4)"}}>
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round"><line x1="12" y1="20" x2="12" y2="10"/><line x1="18" y1="20" x2="18" y2="4"/><line x1="6" y1="20" x2="6" y2="16"/></svg>
          </div>
          <div style={{minWidth:0,flex:1}}>
            {editingName ? (
              <input
                autoFocus value={projectName}
                onChange={e=>setProjectName(e.target.value)}
                onBlur={()=>setEditingName(false)}
                onKeyDown={e=>{ if(e.key==="Enter"||e.key==="Escape") setEditingName(false); }}
                style={{
                  fontSize:isMobile?14:16, fontWeight:700, color:T.text,
                  letterSpacing:"-0.02em", background:"transparent",
                  border:"none", borderBottom:"2px solid #6366f1",
                  outline:"none", width: "100%", maxWidth: 220,
                  padding:"0 2px", fontFamily:"inherit"
                }}
              />
            ) : (
              <div onClick={()=>setEditingName(true)} title="클릭하여 이름 수정"
                style={{
                  fontSize:isMobile?14:16, fontWeight:700, color:T.text,
                  letterSpacing:"-0.02em", cursor:"text",
                  borderBottom:"2px solid transparent",
                  transition:"border-color 0.15s",
                  padding:"0 2px",
                  whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis"
                }}
                onMouseEnter={e=>e.currentTarget.style.borderBottomColor="#6366f144"}
                onMouseLeave={e=>e.currentTarget.style.borderBottomColor="transparent"}
              >
                {projectName}
              </div>
            )}
            {!isMobile&&<div style={{fontSize:11,color:T.textMuted,display:"flex",alignItems:"center",gap:6,marginTop:1}}>
              {items.filter(i=>i.type!=="sector").length}개 작업
              {dbStatus==="saving"&&<span style={{color:"#f59e0b",fontSize:10}}>● 저장 중...</span>}
              {dbStatus==="saved"&&<span style={{color:"#10b981",fontSize:10}}>● 저장됨</span>}
              {dbStatus==="error"&&<span style={{color:"#f87171",fontSize:10}}>● 저장 실패</span>}
            </div>}
            {isMobile && (
              <div style={{fontSize:10,color:T.textMuted,display:"flex",alignItems:"center",gap:5}}>
                {dbStatus==="saving"&&<span style={{color:"#f59e0b"}}>● 저장 중</span>}
                {dbStatus==="saved"&&<span style={{color:"#10b981"}}>● 저장됨</span>}
                {dbStatus==="error"&&<span style={{color:"#f87171"}}>● 실패</span>}
                {dbStatus==="idle"&&<span>v{APP_VERSION}</span>}
              </div>
            )}
          </div>
        </div>

        {/* TAB SWITCHER */}
        <div style={{display:"flex",background:T.bg,border:`1px solid ${T.border}`,borderRadius:8,padding:3,flexShrink:0}}>
          {[
            { key:"list", label:"목록", icon:ListIcon, mobileOnly:true },
            { key:"gantt", label:"간트", icon:GanttIcon },
            { key:"kanban", label:"칸반", icon:KanbanIcon },
          ].filter(tab => !tab.mobileOnly || isMobile).map(tab => {
            const I = tab.icon;
            return (
              <button key={tab.key} onClick={()=>{setActiveTab(tab.key);setMenuOpen(false);}}
                style={{padding: isMobile?"6px 9px":"6px 12px", border:"none", cursor:"pointer", borderRadius:6,
                  fontSize:12, fontWeight:600, transition:"all 0.15s", display:"flex", alignItems:"center", gap:5,
                  background: activeTab===tab.key ? "#6366f1" : "transparent",
                  color: activeTab===tab.key ? "#fff" : T.textMuted}}>
                <I size={13}/>
                {(!isMobile || activeTab===tab.key) && <span>{tab.label}</span>}
              </button>
            );
          })}
        </div>

        {isMobile ? (
          <div style={{display:"flex",gap:5,alignItems:"center",flexShrink:0}}>
            <button onClick={()=>openAdd("task")} style={{...SB.add,padding:"6px 10px"}}>
              <PlusIcon size={12} color="#fff"/>
            </button>
            <button onClick={()=>setMenuOpen(v=>!v)} style={{...SB.nav,padding:"6px 8px"}}>
              {menuOpen?<CloseIcon size={16}/>:<MenuIcon size={16}/>}
            </button>
          </div>
        ) : (
          <div style={{marginLeft:"auto"}}>
            {renderToolbar()}
          </div>
        )}
      </div>

      {isMobile&&menuOpen&&(
        <div style={{position:"absolute",top:isMobile?(updateAvailable?96:46):54,right:8,left:8,zIndex:90,background:T.bgCard,border:`1px solid ${T.border}`,borderRadius:10,padding:10,boxShadow:"0 8px 32px rgba(0,0,0,0.5)",animation:"fadeIn 0.15s"}}>
          {renderToolbar(true)}
        </div>
      )}

      {/* ─── LOADING OVERLAY ─── */}
      {/* Renders while we resolve the channel id and fetch from Supabase.
          Replaces the previous "flash of sample 9-task gantt" with a clean
          spinner so users no longer see the wrong project briefly. */}
      {isLoading && (
        <div style={{
          flex:1, display:"flex", flexDirection:"column",
          alignItems:"center", justifyContent:"center", gap:14,
          background:T.bg, color:T.textMuted,
        }}>
          <div style={{
            width:36, height:36, borderRadius:"50%",
            border:`3px solid ${T.border}`,
            borderTopColor:"#6366f1",
            animation:"ganttSpin 0.85s linear infinite",
          }}/>
          <div style={{fontSize:13, fontWeight:500}}>프로젝트 데이터를 불러오는 중...</div>
          <style>{`@keyframes ganttSpin { to { transform: rotate(360deg); } }`}</style>
        </div>
      )}

      {/* ─── LIST VIEW (mobile primary) ─── */}
      {!isLoading && activeTab==="list" && (
        <div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden",background:T.bg}}>
          {renderListHeader()}
          <div ref={leftPanelListRef} className="hide-scrollbar"
            style={{flex:1,overflowY:"auto",overflowX:"hidden",WebkitOverflowScrolling:"touch"}}>
            <div style={{height:totalLeftHeight,minHeight:"100%"}}>
              {flatRows.length === 0 && (
                <div style={{padding:"40px 20px",textAlign:"center",color:T.textFaint,fontSize:13}}>
                  {searchQuery.trim() ? "검색 결과가 없습니다" : "작업이 없습니다. 우측 상단 + 버튼으로 추가하세요."}
                </div>
              )}
              {flatRows.map(({item},i)=>(
                <div key={item.id} style={{height:getRowHeight(i),overflow:"hidden",flexShrink:0}}>
                  <LeftRow item={item}
                    hovered={hoveredId===item.id}
                    onHover={setHoveredId}
                    onEdit={it => { if(it?.type==="new-child") openAdd("task",it.parentId); else openEdit(it); }}
                    onDelete={deleteItem} onToggle={toggleCollapse}
                    allItems={items}
                    isSelected={selectedId===item.id}
                    onSelect={id=>setSelectedId(prev=>prev===id?null:id)}
                    isDragging={rowDragState?.id===item.id || activeDragId===item.id}
                    activeDragId={activeDragId}
                    onActivateDrag={id=>setActiveDragId(prev=>prev===id?null:id)}
                    onMoveUp={()=>moveStep(item.id,-1)}
                    onMoveDown={()=>moveStep(item.id,1)}
                    canUp={flatRowsForDrag.findIndex(r=>r.id===item.id)>0}
                    canDown={flatRowsForDrag.findIndex(r=>r.id===item.id)<flatRowsForDrag.length-1}
                    onTapWhileDragging={handleRowTapWhileDragging}
                    themeKey={themeKey}
                    isMobile={true}
                    onMouseDownDrag={e=>{
                      e.preventDefault();
                      e.stopPropagation();
                      const t = e.touches?.[0] || e;
                      setRowDragPos({x: t.clientX, y: t.clientY});
                      setRowDragState({id:item.id, startY:t.clientY, dropIdx:i, dropPos:"below"});
                    }}
                    dropIndicator={
                      rowDragState && rowDragState.dropIdx===i
                        ? rowDragState.dropPos
                        : null
                    }
                  />
                </div>
              ))}
              <div style={{height:80}}/>
            </div>
          </div>

          {/* FAB add button */}
          <button onClick={()=>openAdd("task")}
            style={{position:"fixed",bottom:20,right:16,width:54,height:54,borderRadius:"50%",
              background:"linear-gradient(135deg,#6366f1,#8b5cf6)",border:"none",color:"#fff",
              display:"flex",alignItems:"center",justifyContent:"center",cursor:"pointer",
              boxShadow:"0 6px 20px rgba(99,102,241,0.5)",zIndex:50,transition:"transform 0.15s"}}
            onTouchStart={e=>e.currentTarget.style.transform="scale(0.92)"}
            onTouchEnd={e=>e.currentTarget.style.transform="scale(1)"}>
            <PlusIcon size={22} color="#fff"/>
          </button>
        </div>
      )}

      {/* ─── KANBAN VIEW ─── */}
      {!isLoading && activeTab==="kanban" && (
        <KanbanView
          items={items}
          themeKey={themeKey}
          isMobile={isMobile}
          onEdit={item => {
            if (item._kanbanMove) {
              setItems(prev => prev.map(i => i.id===item.id ? {...i, progress:item.progress} : i));
              return;
            }
            openEdit(item);
          }}
          onDelete={deleteItem}
          openAdd={openAdd}
          openAddSector={openAddSector}
          collapsedSectors={kanbanCollapsed}
          toggleSector={toggleKanbanSector}
          catCollapsed={kanbanCatCollapsed}
          toggleCategory={toggleKanbanCategory}
          labelW={kanbanLabelW}
          setLabelW={setKanbanLabelW}
          colResizing={kanbanColResizing}
          setColResizing={setKanbanColResizing}
          onAddCategory={openAddCategory}
          onEditCategory={openEditCategory}
          onMoveCategory={moveCategory}
        />
      )}

      {/* ─── GANTT VIEW ─── */}
      {!isLoading && activeTab==="gantt" && (() => {
        return (
          <div style={{display:"flex",flex:1,overflow:"hidden",position:"relative"}}>
            {/* LEFT PANEL: only on desktop/tablet */}
            {!isMobile && (
              <div style={{
                width:leftPanelWidth, flexShrink:0,
                borderRight:`1px solid ${T.border}`,
                display:"flex", flexDirection:"row",
                background:T.bg,
              }}>
                {/* === Category column (vertical labels) === */}
                <div style={{width:CATEGORY_COL_W, flexShrink:0, borderRight:`1px solid ${T.border}`, display:"flex", flexDirection:"column", background:T.bgDeep}}>
                  {/* Top corner: + button to add category */}
                  <div style={{height:32+24, flexShrink:0, borderBottom:`1px solid ${T.border}`, background:T.bgHeader, display:"flex",alignItems:"center",justifyContent:"center"}}>
                    <button onClick={openAddCategory} title="대분류 추가"
                      style={{
                        width:20, height:20, padding:0,
                        background:"transparent", border:`1px solid ${T.border}`,
                        borderRadius:5, cursor:"pointer", color:T.textSub,
                        display:"flex", alignItems:"center", justifyContent:"center",
                        transition:"all 0.12s",
                      }}
                      onMouseEnter={e=>{ e.currentTarget.style.background=hexAlpha("#6366f1",0.15); e.currentTarget.style.borderColor="#6366f1"; }}
                      onMouseLeave={e=>{ e.currentTarget.style.background="transparent"; e.currentTarget.style.borderColor=T.border; }}>
                      <PlusIcon size={11} color={T.textSub}/>
                    </button>
                  </div>
                  {/* Scrollable category labels container - synced to list scroll */}
                  <div className="hide-scrollbar" style={{flex:1, overflowY:"auto", overflowX:"hidden", position:"relative"}}
                    ref={el => { if (el) el._isCategoryColumn = true; }}
                    onScroll={e => {
                      if (leftPanelListRef.current && Math.abs(leftPanelListRef.current.scrollTop - e.target.scrollTop) > 1) {
                        leftPanelListRef.current.scrollTop = e.target.scrollTop;
                      }
                    }}>
                    <div style={{height:totalLeftHeight, position:"relative"}}>
                      {categoryRanges.map(({ category, top, height }, idx) => {
                        const canUp = idx > 0;
                        const canDown = idx < categoryRanges.length - 1;
                        return (
                        <div key={category.id}
                          onClick={()=>toggleCollapse(category.id)}
                          onContextMenu={e=>{ e.preventDefault(); openEditCategory(category.id); }}
                          title={`${category.name} · 클릭: ${category.collapsed?"작업 펼치기":"작업 접기"} · 우클릭: 편집`}
                          style={{
                            position:"absolute", left:0, right:0, top, height,
                            display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center",
                            borderBottom:`1px solid ${T.borderLight}`,
                            background: category.collapsed ? hexAlpha("#6366f1",0.04) : "transparent",
                            cursor:"pointer",
                            overflow:"hidden",
                            transition:"background 0.12s",
                          }}
                          onMouseEnter={e=>e.currentTarget.style.background=hexAlpha("#6366f1",0.10)}
                          onMouseLeave={e=>e.currentTarget.style.background = category.collapsed ? hexAlpha("#6366f1",0.04) : "transparent"}>
                          {/* Move up button (top) */}
                          {canUp && (
                            <button
                              onClick={e=>{ e.stopPropagation(); moveCategory(category.id, -1); }}
                              title="위로"
                              style={{position:"absolute",top:1,left:0,right:0,height:14,display:"flex",alignItems:"center",justifyContent:"center",background:"transparent",border:"none",cursor:"pointer",color:T.textFaint,padding:0}}
                              onMouseEnter={e=>e.currentTarget.style.color="#818cf8"}
                              onMouseLeave={e=>e.currentTarget.style.color=T.textFaint}>
                              <svg width="9" height="6" viewBox="0 0 9 6"><path d="M4.5 0L9 6H0z" fill="currentColor"/></svg>
                            </button>
                          )}
                          {/* Vertical character stack */}
                          <div style={{
                            writingMode:"vertical-rl",
                            textOrientation:"upright",
                            fontSize:11, fontWeight:700, color:T.textSub,
                            letterSpacing:"0.05em",
                            lineHeight:1.15,
                            whiteSpace:"nowrap",
                            overflow:"hidden", textOverflow:"ellipsis",
                            maxHeight: Math.max(0, height - 36),
                          }}>
                            {category.name}
                          </div>
                          {/* Move down button (bottom) */}
                          {canDown && (
                            <button
                              onClick={e=>{ e.stopPropagation(); moveCategory(category.id, 1); }}
                              title="아래로"
                              style={{position:"absolute",bottom:1,left:0,right:0,height:14,display:"flex",alignItems:"center",justifyContent:"center",background:"transparent",border:"none",cursor:"pointer",color:T.textFaint,padding:0}}
                              onMouseEnter={e=>e.currentTarget.style.color="#818cf8"}
                              onMouseLeave={e=>e.currentTarget.style.color=T.textFaint}>
                              <svg width="9" height="6" viewBox="0 0 9 6"><path d="M4.5 6L0 0H9z" fill="currentColor"/></svg>
                            </button>
                          )}
                        </div>
                        );
                      })}
                    </div>
                  </div>
                </div>
                {/* === Right side: header + list === */}
                <div style={{flex:1, display:"flex", flexDirection:"column", minWidth:0}}>
                <div style={{flexShrink:0,borderBottom:`1px solid ${T.border}`,background:T.bgHeader}}>
                  <div style={{height:32,padding:"0 6px 0 4px",display:"flex",alignItems:"center",gap:5,borderBottom:`1px solid ${T.borderLight}`}}>
                    {searchQuery ? (
                      <div style={{display:"flex",alignItems:"center",flex:1,gap:4,background:T.bgInput,border:`1px solid ${T.border}`,borderRadius:6,padding:"3px 6px"}}>
                        <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke={T.textFaint} strokeWidth="2"><circle cx="11" cy="11" r="7"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
                        <input value={searchQuery} onChange={e=>setSearchQuery(e.target.value)} autoFocus
                          placeholder="검색..." style={{flex:1,background:"transparent",border:"none",outline:"none",color:T.text,fontSize:11,fontFamily:"inherit"}}/>
                        <span onClick={()=>setSearchQuery("")} style={{cursor:"pointer",color:T.textFaint,fontSize:11,padding:"0 2px"}}>×</span>
                      </div>
                    ) : (
                      <div style={{display:"flex",alignItems:"center",gap:6,flex:1}}>
                        <span style={{fontSize:10,fontWeight:600,color:T.textFaint,textTransform:"uppercase",letterSpacing:"0.05em",paddingLeft:4}}>작업 목록</span>
                        <button onClick={()=>setSearchQuery(" ")} style={{background:"transparent",border:"none",cursor:"pointer",color:T.textFaint,padding:4,marginLeft:"auto",display:"flex",alignItems:"center"}}
                          title="검색">
                          <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="11" cy="11" r="7"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
                        </button>
                      </div>
                    )}
                  </div>
                  <div style={{display:"grid",gridTemplateColumns:`26px 6px ${nameColWidth!=null?nameColWidth+"px":COL_NAME} ${COL_START} ${COL_END} ${COL_PROG}`,alignItems:"center",height:24,padding:"0 6px 0 4px"}}>
                    <div/><div/>
                    <div style={{fontSize:10,fontWeight:600,color:T.textFaint,paddingLeft:4,position:"relative",height:"100%",display:"flex",alignItems:"center"}}>
                      작업명
                      {/* Resize handle */}
                      <div
                        onMouseDown={e=>{
                          e.preventDefault();
                          e.stopPropagation();
                          const cell = e.currentTarget.parentElement;
                          const startWidth = nameColWidth != null ? nameColWidth : (cell?.getBoundingClientRect().width || 200);
                          setColResizing({startX: e.clientX, startWidth});
                        }}
                        onDoubleClick={()=>setNameColWidth(null)}
                        title="드래그하여 너비 조절 / 더블클릭하여 자동"
                        style={{
                          position:"absolute", right:-3, top:-2, bottom:-2, width:8,
                          cursor:"col-resize", zIndex:10,
                          background: colResizing ? hexAlpha("#6366f1",0.3) : "transparent",
                          transition:"background 0.15s",
                        }}
                        onMouseEnter={e=>{ if(!colResizing) e.currentTarget.style.background=hexAlpha("#6366f1",0.2); }}
                        onMouseLeave={e=>{ if(!colResizing) e.currentTarget.style.background="transparent"; }}
                      />
                    </div>
                    <div style={{fontSize:10,fontWeight:600,color:T.textFaint,textAlign:"center"}}>시작</div>
                    <div style={{fontSize:10,fontWeight:600,color:T.textFaint,textAlign:"center"}}>종료</div>
                    <div style={{fontSize:10,fontWeight:600,color:T.textFaint,textAlign:"center"}}>진행률</div>
                  </div>
                </div>
                <div ref={leftPanelListRef}
                  onWheel={e=>{
                    if (e.ctrlKey || e.metaKey) return; // app-zoom takes priority
                    if (Math.abs(e.deltaY) > Math.abs(e.deltaX) && chartRef.current) {
                      chartRef.current.scrollTop += e.deltaY;
                    }
                  }}
                  onScroll={e=>{
                    if (chartRef.current && Math.abs(chartRef.current.scrollTop - e.target.scrollTop) > 1) {
                      chartRef.current.scrollTop = e.target.scrollTop;
                    }
                  }}
                  className="hide-scrollbar" style={{flex:1,overflowY:"auto",overflowX:"hidden",WebkitOverflowScrolling:"touch"}}>
                  <div style={{height:totalLeftHeight}}>
                    {(() => {
                      // Render: walk through categories in order, emit their rows.
                      // Empty categories get a placeholder.
                      const elements = [];
                      let rowIdx = 0;
                      allCategories.forEach(cat => {
                        const startIdx = rowIdx;
                        while (rowIdx < flatRows.length && flatRows[rowIdx].categoryId === cat.id) rowIdx++;
                        if (rowIdx === startIdx) {
                          elements.push(
                            <div key={"cat-empty-"+cat.id} style={{height:baseRowHeight,borderBottom:`1px solid ${T.borderLight}`,display:"flex",alignItems:"center",padding:"0 14px"}}>
                              <span style={{fontSize:11,color:T.textFaint}}>· {cat.name} · 비어있음</span>
                            </div>
                          );
                          return;
                        }
                        for (let i = startIdx; i < rowIdx; i++) {
                          const { item } = flatRows[i];
                          elements.push(
                            <div key={item.id} style={{height:getRowHeight(i),overflow:"hidden",flexShrink:0}}>
                              <LeftRow item={item}
                                hovered={hoveredId===item.id}
                                onHover={setHoveredId}
                                onEdit={it => { if(it?.type==="new-child") openAdd("task",it.parentId); else openEdit(it); }}
                                onDelete={deleteItem} onToggle={toggleCollapse}
                                allItems={items}
                                isSelected={selectedId===item.id}
                                onSelect={id=>setSelectedId(prev=>prev===id?null:id)}
                                isDragging={rowDragState?.id===item.id || activeDragId===item.id}
                                activeDragId={activeDragId}
                                onActivateDrag={id=>setActiveDragId(prev=>prev===id?null:id)}
                                onMoveUp={()=>moveStep(item.id,-1)}
                                onMoveDown={()=>moveStep(item.id,1)}
                                canUp={flatRowsForDrag.findIndex(r=>r.id===item.id)>0}
                                canDown={flatRowsForDrag.findIndex(r=>r.id===item.id)<flatRowsForDrag.length-1}
                                onTapWhileDragging={handleRowTapWhileDragging}
                                themeKey={themeKey}
                                isMobile={false}
                                nameColWidth={nameColWidth}
                                onMouseDownDrag={e=>{
                                  e.preventDefault();
                                  e.stopPropagation();
                                  setRowDragPos({x: e.clientX, y: e.clientY});
                                  setRowDragState({id:item.id, startY:e.clientY, dropIdx:i, dropPos:"below"});
                                }}
                                dropIndicator={
                                  rowDragState && rowDragState.dropIdx===i
                                    ? rowDragState.dropPos
                                    : null
                                }
                              />
                            </div>
                          );
                        }
                      });
                      return elements;
                    })()}
                  </div>
                </div>
                </div>{/* end right-side wrapper */}
              </div>
            )}

            {/* CHART PANEL */}
            <div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden",minWidth:0}}>
              {/* Chart toolbar (mobile only): zoom + today + milestone */}
              {isMobile && (
                <div style={{flexShrink:0,padding:"8px 10px",borderBottom:`1px solid ${T.border}`,background:T.bgHeader,display:"flex",gap:6,alignItems:"center",overflowX:"auto"}} className="hide-scrollbar">
                  <button onClick={()=>{if(chartRef.current){const chartW=chartRef.current.clientWidth;chartRef.current.scrollLeft = todayX - chartW/2;}}} style={{...SB.today,padding:"7px 12px",flexShrink:0}}>오늘</button>
                  <div style={{display:"flex",border:`1px solid ${T.border}`,borderRadius:7,overflow:"hidden",flexShrink:0}}>
                    {ZOOM_PRESETS.map((z,i)=>(
                      <button key={z.key} onClick={()=>changeZoom(i)}
                        style={{padding:"7px 11px",border:"none",cursor:"pointer",fontSize:12,fontWeight:600,
                          background:i===zoomIdx?"#6366f1":"transparent",
                          color:i===zoomIdx?"#fff":T.textMuted}}>
                        {z.label}
                      </button>
                    ))}
                  </div>
                  <button onClick={()=>setShowMilestones(v=>!v)}
                    style={{...SB.sec,color:showMilestones?"#818cf8":T.textMuted,border:"1px solid "+(showMilestones?"#4338ca":T.border),padding:"7px 9px",flexShrink:0}}>
                    <MilestoneIcon size={12} color={showMilestones?"#818cf8":T.textFaint}/>
                  </button>
                </div>
              )}

              <div ref={headerRef} style={{overflowX:"hidden",flexShrink:0,borderBottom:`1px solid ${T.border}`,background:T.bgHeader}}>
                <div style={{width:totalW,display:"flex",flexDirection:"column"}}>
                  <div style={{display:"flex",height:32,background:T.bgHeader,borderBottom:`1px solid ${T.border}`}}>
                    {headerGroups.map((g,i)=>(
                      <div key={i} style={{width:g.count*DAY_WIDTH,flexShrink:0,display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:600,color:T.textSub,borderRight:`1px solid ${T.border}`}}>
                        {g.label}
                      </div>
                    ))}
                  </div>
                  <div style={{display:"flex",height:24,background:T.bgHeader}}>
                    {days.map((d,i)=>{
                      const isT=d.getTime()===today.getTime(), isW=d.getDay()===0||d.getDay()===6;
                      return (
                        <div key={i} style={{width:DAY_WIDTH,flexShrink:0,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",
                          background:isT?"rgba(99,102,241,0.15)":isW?"rgba(255,255,255,0.01)":"transparent",borderRight:DAY_WIDTH>=12?`1px solid ${T.borderLight}`:"none"}}>
                          {DAY_WIDTH>=14&&<span style={{fontSize:DAY_WIDTH<20?7:9,fontWeight:isT?700:500,color:isT?"#6366f1":T.textMuted}}>
                            {DAY_WIDTH>=8?(d.getMonth()+1)+"/"+d.getDate():""}
                          </span>}
                          {DAY_WIDTH>=22&&<span style={{fontSize:8,color:isT?"#6366f1":T.textFaint}}>
                            {["일","월","화","수","목","금","토"][d.getDay()]}
                          </span>}
                        </div>
                      );
                    })}
                  </div>
                </div>
              </div>

              <div ref={chartRef}
                onScroll={e=>{
                  if (headerRef.current) headerRef.current.scrollLeft = e.target.scrollLeft;
                  if (leftPanelListRef.current && Math.abs(leftPanelListRef.current.scrollTop - e.target.scrollTop) > 1) {
                    leftPanelListRef.current.scrollTop = e.target.scrollTop;
                  }
                }}
                style={{flex:1,overflowX:"scroll",overflowY:"scroll",position:"relative",WebkitOverflowScrolling:"touch",touchAction:"pan-x pan-y"}}>
                <div
                  onMouseDown={e=>{
                    if (e.button !== 0) return;
                    const tag = e.target.tagName;
                    if (tag === "BUTTON" || tag === "INPUT" || tag === "SELECT" || tag === "TEXTAREA") return;
                    // Walk up from the clicked element; abort pan if we hit a task bar or interactive control
                    let el = e.target;
                    let isInteractive = false;
                    while (el && el !== e.currentTarget) {
                      // Task bars use zIndex 6
                      const z = el.style?.zIndex;
                      if (z === "6" || z === 6) { isInteractive = true; break; }
                      if (el.dataset?.dragbtn) { isInteractive = true; break; }
                      // Has its own click handler (best-effort: check for cursor pointer)
                      el = el.parentElement;
                    }
                    if (isInteractive) return;
                    e.preventDefault();
                    const c = chartRef.current;
                    if (c) setChartPan({startX:e.clientX, startY:e.clientY, scrollLeft:c.scrollLeft, scrollTop:c.scrollTop});
                  }}
                  style={{width:totalW,minHeight:totalLeftHeight,position:"relative",
                    cursor:chartPan?"grabbing":"grab"}}>
                  {/* Pan capture layer — covers entire chart area below task bars (zIndex < 6) */}
                  <div
                    onMouseDown={e=>{
                      if (e.button !== 0) return;
                      e.preventDefault();
                      const c = chartRef.current;
                      if (c) setChartPan({startX:e.clientX, startY:e.clientY, scrollLeft:c.scrollLeft, scrollTop:c.scrollTop});
                    }}
                    style={{position:"absolute",inset:0,zIndex:5,cursor:chartPan?"grabbing":"grab"}}
                  />
                  {todayX>=0&&todayX<=totalW&&(
                    <div style={{position:"absolute",left:todayX,top:0,bottom:0,width:2,background:"#6366f1",zIndex:7,opacity:0.6,pointerEvents:"none"}}>
                      <div style={{position:"absolute",top:0,left:-14,background:themeKey==="dark"?"#0d111a":"#f0f1f5",border:"1px solid #6366f1",borderRadius:4,padding:"1px 5px",fontSize:9,fontWeight:700,color:"#818cf8"}}>오늘</div>
                    </div>
                  )}

                  {days.map((d,i)=>{
                    if (DAY_WIDTH < 5) {
                      if (d.getMonth()===0 && d.getDate()===1) {
                        return <div key={i} style={{position:"absolute",left:i*DAY_WIDTH,top:0,bottom:0,width:1,background:T.borderLight,zIndex:1}}/>;
                      }
                      if (d.getMonth()===6 && d.getDate()===1) {
                        return <div key={i} style={{position:"absolute",left:i*DAY_WIDTH,top:0,bottom:0,width:1,background:T.borderLight,opacity:0.5,zIndex:1}}/>;
                      }
                      return null;
                    }
                    if (d.getDay()===0||d.getDay()===6) {
                      return <div key={i} style={{position:"absolute",left:i*DAY_WIDTH,top:0,bottom:0,width:DAY_WIDTH,background:themeKey==="dark"?"rgba(255,255,255,0.015)":"rgba(0,0,0,0.015)",zIndex:1}}/>;
                    }
                    return null;
                  })}

                  {flatRows.map((_,i)=>{
                    const rh=getRowHeight(i), rt=getRowTop(i);
                    return <div key={i} style={{position:"absolute",left:0,top:rt,width:"100%",height:rh,borderBottom:`1px solid ${T.borderLight}`,zIndex:0}}/>;
                  })}

                  <DepArrows flatRows={flatRows} viewStart={effectiveViewStart} DAY_WIDTH={DAY_WIDTH} getRowTop={getRowTop} rowH={baseRowHeight}/>

                  {flatRows.map(({item},rowIdx)=>{
                    const top=getRowTop(rowIdx);
                    const rh = baseRowHeight;
                    if (item.type==="sector") {
                      const kids=items.filter(c=>c.parentId===item.id&&c.start&&c.end&&!isNaN(c.start.getTime())&&!isNaN(c.end.getTime()));
                      if (!kids.length) return null;
                      const ms=new Date(Math.min(...kids.map(c=>c.start.getTime())));
                      const me=new Date(Math.max(...kids.map(c=>c.end.getTime())));
                      if (isNaN(ms.getTime())||isNaN(me.getTime())) return null;
                      const l=dateDiff(effectiveViewStart,ms)*DAY_WIDTH;
                      const w=(dateDiff(ms,me)+1)*DAY_WIDTH-2;
                      return <div key={item.id} style={{position:"absolute",left:0,top:0,width:"100%",height:rh,zIndex:2,pointerEvents:"none"}}>
                        <div title={item.name+" ("+Math.max(1,dateDiff(ms,me)+1)+" d)"}
                          style={{position:"absolute",top:top+rh/2-Math.round(rh/8),left:l,width:Math.max(w,2),height:Math.round(rh/4),background:`linear-gradient(135deg,${item.color},${hexAlpha(item.color,0.6)})`,borderRadius:3,boxShadow:`0 2px 8px ${hexAlpha(item.color,0.3)}`,opacity:0.85}}/>
                        <div style={{position:"absolute",left:l+Math.max(w,2)+6,top:top+rh/2-7,height:14,fontSize:10,fontWeight:700,color:item.color,whiteSpace:"nowrap"}}>
                          {item.name} ({Math.max(1,dateDiff(ms,me)+1)} d)
                        </div>
                      </div>
                    }
                    if (item.type==="milestone") {
                      if (!item.start) return null;
                      const mc=effectiveColor(item,items);
                      const cx=(dateDiff(effectiveViewStart,item.start)+0.5)*DAY_WIDTH;
                      const cy=top+rh/2, sz=8;
                      return (
                        <div key={item.id} title={[
                            "Phase: "+(items.find(s=>s.id===item.parentId)?.name||""),
                            "Milestone: "+item.name,
                            "Date: "+(item.start?toISO(item.start):"")
                          ].join("\n")}
                          onMouseEnter={()=>setHoveredId(item.id)} onMouseLeave={()=>setHoveredId(null)}
                          onClick={e=>{e.stopPropagation();if(linkingFrom!==null){handleLinkClick(e,item.id);}else setSelectedId(prev=>prev===item.id?null:item.id);}}
                          style={{position:"absolute",left:cx-sz,top:cy-sz,width:sz*2,height:sz*2,zIndex:6,cursor:"pointer"}}>
                          <div style={{position:"absolute",inset:3,background:mc,transform:"rotate(45deg)",borderRadius:2,border:`2px solid ${themeKey==="dark"?"#0f1117":"#fff"}`,boxShadow:hoveredId===item.id?`0 0 12px ${mc}`:`0 2px 6px rgba(0,0,0,0.4)`}}/>
                          {<div style={{position:"absolute",left:sz*2+6,top:0,height:sz*2,display:"flex",alignItems:"center",fontSize:9,fontWeight:600,color:mc,whiteSpace:"nowrap",pointerEvents:"none"}}>{item.name}</div>}
                        </div>
                      );
                    }
                    if (!item.start || !item.end) return null;
                    const barColor=effectiveColor(item,items);
                    const left=dateDiff(effectiveViewStart,item.start)*DAY_WIDTH;
                    const width=Math.max((dateDiff(item.start,item.end)+1)*DAY_WIDTH-3,14);
                    return (
                      <div key={item.id} style={{display:"contents"}}>
                        <div onMouseEnter={()=>setHoveredId(item.id)} onMouseLeave={()=>setHoveredId(null)}
                          style={{position:"absolute",top:top+Math.round(rh*0.19),left,width,height:Math.round(rh*0.62),zIndex:6}}>
                          {!isMobile && <div onMouseDown={e=>{e.preventDefault();e.stopPropagation();setDragState({id:item.id,type:"resize-left",startX:e.clientX,origStart:item.start,origEnd:item.end});}}
                            style={{position:"absolute",left:0,top:0,width:7,height:"100%",cursor:"ew-resize",zIndex:2}}/>}
                          <div title={((s)=>s.find(x=>x.id===item.parentId)?.name+"\n"+item.name+" ("+Math.max(1,dateDiff(item.start,item.end)+1)+" d, "+(item.progress||0)+"%)")(items)}
                            onMouseDown={e=>{if(linkingFrom!==null||isMobile)return;e.preventDefault();e.stopPropagation();setDragState({id:item.id,type:"move",startX:e.clientX,origStart:item.start,origEnd:item.end});}}
                            onClick={e=>{e.stopPropagation();if(linkingFrom!==null){handleLinkClick(e,item.id);}else setSelectedId(prev=>prev===item.id?null:item.id);}}
                            style={{position:"absolute",inset:0,background:`linear-gradient(135deg,${barColor},${lighten(barColor,0.15)})`,
                              borderRadius:6,
                              border:themeKey==="light"?("1px solid "+hexAlpha(barColor,0.3)):"none",
                              boxShadow:hoveredId===item.id?`0 4px 20px ${hexAlpha(barColor,0.25)}`:"0 1px 3px rgba(0,0,0,0.2)",
                              cursor:linkingFrom!==null?"crosshair":(isMobile?"pointer":"pointer"),transition:"box-shadow 0.15s",overflow:"hidden"}}>
                            <div style={{position:"absolute",left:0,top:0,bottom:0,width:item.progress+"%",background:"rgba(0,0,0,0.25)"}}/>
                            {width>20&&item.progress>0&&<div style={{position:"absolute",left:0,top:0,right:0,bottom:0,display:"flex",alignItems:"center",justifyContent:"center",fontSize:9,fontWeight:700,color:"#fff",textShadow:"0 1px 2px rgba(0,0,0,0.6)"}}>
                              {item.progress}%
                            </div>}
                          </div>
                          {!isMobile && <div onMouseDown={e=>{e.preventDefault();e.stopPropagation();setDragState({id:item.id,type:"resize-right",startX:e.clientX,origStart:item.start,origEnd:item.end});}}
                            style={{position:"absolute",right:0,top:0,width:7,height:"100%",cursor:"ew-resize",zIndex:2}}/>}
                        </div>
                        <div style={{position:"absolute",
                          left:left+width+6,top:top+Math.round(rh*0.19),
                          height:Math.round(rh*0.62),
                          display:"flex",alignItems:"center",
                          fontSize:9,fontWeight:500,
                          color:themeKey==="dark"?hexAlpha(barColor,0.9):barColor,
                          whiteSpace:"nowrap",pointerEvents:"none",zIndex:5,
                          textShadow:themeKey==="dark"?"0 1px 3px rgba(0,0,0,0.8)":"none"
                        }}>
                          {item.name}{item.start&&item.end?" ("+Math.max(1,dateDiff(item.start,item.end)+1)+" d)":""}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>
          </div>
        );
      })()}

      {/* CATEGORY MODAL */}
      {showCatModal && catForm && (
        <div style={{position:"fixed",inset:0,background:themeKey==="dark"?"rgba(0,0,0,0.78)":"rgba(0,0,0,0.4)",
          display:"flex",alignItems: isMobile?"flex-end":"center",justifyContent:"center",zIndex:310,
          animation:"fadeIn 0.15s"}} onClick={()=>setShowCatModal(false)}>
          <div style={{background:T.bgCard,border:`1px solid ${T.border}`,
            borderRadius: isMobile ? "16px 16px 0 0" : 16,
            padding: isMobile?"14px 16px 18px":24,
            width: isMobile ? "100%" : 460, maxWidth: "100%",
            maxHeight: isMobile ? "92vh" : "85vh",
            overflowY:"auto",WebkitOverflowScrolling:"touch",
            animation: isMobile?"slideUp 0.25s ease":"fadeIn 0.15s"}}
            onClick={e=>e.stopPropagation()}>
            {isMobile && <div style={{width:36,height:4,background:T.border,borderRadius:2,margin:"0 auto 12px"}}/>}
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14}}>
              <h2 style={{margin:0,fontSize:isMobile?15:16,fontWeight:700,color:T.text}}>
                {catForm.id ? "대분류 편집" : "대분류 추가"}
              </h2>
              <button onClick={()=>setShowCatModal(false)} style={{...SB.nav,padding:"4px 6px"}}>
                <CloseIcon size={14}/>
              </button>
            </div>
            <div style={{display:"flex",flexDirection:"column",gap:13}}>
              <F label="이름">
                <input value={catForm.name} onChange={e=>setCatForm({...catForm, name:e.target.value})} style={SB.inp} autoFocus
                  placeholder="예: Agreement, Due Diligence..."/>
              </F>
              <F label={`포함할 섹터 (${catForm.sectorIds.length}개 선택)`}>
                {(()=>{
                  // Show all sectors. Disabled = currently belongs to another category but always selectable
                  const allSectors = items.filter(i => i.type==="sector").slice().sort((a,b)=>(a.order??0)-(b.order??0));
                  if (allSectors.length === 0) {
                    return <div style={{padding:"10px 12px",fontSize:12,color:T.textFaint,background:T.bg,border:`1px solid ${T.border}`,borderRadius:8}}>섹터가 없습니다. 먼저 섹터를 추가하세요.</div>;
                  }
                  const toggle = id => {
                    setCatForm(f => ({
                      ...f,
                      sectorIds: f.sectorIds.includes(id) ? f.sectorIds.filter(s=>s!==id) : [...f.sectorIds, id]
                    }));
                  };
                  return (
                    <div style={{background:T.bg,border:`1px solid ${T.border}`,borderRadius:8,maxHeight:240,overflowY:"auto",WebkitOverflowScrolling:"touch"}}>
                      {allSectors.map(sector => {
                        const checked = catForm.sectorIds.includes(sector.id);
                        const currentCat = items.find(c=>c.id===sector.parentId&&c.type==="category");
                        const inOtherCat = currentCat && (!catForm.id || currentCat.id !== catForm.id) && !checked;
                        const taskCount = items.filter(i=>i.parentId===sector.id&&i.type!=="sector").length;
                        return (
                          <div key={sector.id} onClick={()=>toggle(sector.id)}
                            style={{display:"flex",alignItems:"center",gap:10,
                              padding:"10px 12px",cursor:"pointer",
                              background: checked ? "rgba(99,102,241,0.10)" : "transparent",
                              borderBottom:`1px solid ${T.borderLight}`,
                              transition:"background 0.1s"}}>
                            <div style={{width:18,height:18,borderRadius:4,flexShrink:0,
                                border: checked ? "none" : `1.5px solid ${T.textDim}`,
                                background: checked ? "#6366f1" : "transparent",
                                display:"flex",alignItems:"center",justifyContent:"center",
                                transition:"all 0.12s"}}>
                              {checked && <svg width="11" height="9" viewBox="0 0 9 7" fill="none"><path d="M1 3.5L3.5 6L8 1" stroke="#fff" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"/></svg>}
                            </div>
                            <div style={{width:10,height:10,borderRadius:2,background:sector.color,flexShrink:0}}/>
                            <div style={{flex:1,minWidth:0}}>
                              <div style={{fontSize:13,color:checked?T.text:T.textSub,fontWeight:checked?700:500,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>
                                {sector.name}
                              </div>
                              <div style={{fontSize:10,color:T.textFaint,marginTop:2}}>
                                {taskCount}개 작업
                                {inOtherCat && <span style={{color:"#f59e0b",marginLeft:6}}>· 현재: {currentCat.name}</span>}
                              </div>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  );
                })()}
              </F>
              <div style={{fontSize:11,color:T.textFaint,padding:"6px 0"}}>
                💡 다른 대분류의 섹터를 선택하면 이 대분류로 이동됩니다.
              </div>
            </div>
            <div style={{display:"flex",gap:8,marginTop:20,justifyContent:"flex-end"}}>
              {catForm.id && (
                <button onClick={()=>{
                  deleteCategory(catForm.id);
                  setShowCatModal(false);
                }} style={{...SB.sec,color:"#f87171",border:"1px solid rgba(248,113,113,0.4)",marginRight:"auto"}}>
                  <TrashIcon size={11} color="#f87171"/> 삭제
                </button>
              )}
              <button onClick={()=>setShowCatModal(false)} style={SB.nav}>취소</button>
              <button onClick={saveCategoryForm} style={SB.add}>저장</button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL */}
      {showModal&&form&&(
        <div style={{position:"fixed",inset:0,background:themeKey==="dark"?"rgba(0,0,0,0.78)":"rgba(0,0,0,0.4)",
          display:"flex",alignItems: isMobile?"flex-end":"center",justifyContent:"center",zIndex:300,
          animation:"fadeIn 0.15s"}} onClick={()=>setShowModal(false)}>
          <div style={{background:T.bgCard,border:`1px solid ${T.border}`,
            borderRadius: isMobile ? "16px 16px 0 0" : 16,
            padding: isMobile?"14px 16px 18px":24,
            width: isMobile ? "100%" : 480, maxWidth: "100%",
            maxHeight: isMobile ? "92vh" : "85vh",
            overflowY:"auto",WebkitOverflowScrolling:"touch",
            animation: isMobile?"slideUp 0.25s ease":"fadeIn 0.15s"}}
            onClick={e=>e.stopPropagation()}>
            {isMobile && <div style={{width:36,height:4,background:T.border,borderRadius:2,margin:"0 auto 12px"}}/>}
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14}}>
              <h2 style={{margin:0,fontSize:isMobile?15:16,fontWeight:700,color:T.text}}>
                {editItem?"항목 편집":form.type==="sector"?"섹터 추가":"작업 추가"}
              </h2>
              <button onClick={()=>setShowModal(false)} style={{...SB.nav,padding:"4px 6px"}}>
                <CloseIcon size={14}/>
              </button>
            </div>
            <div style={{display:"flex",flexDirection:"column",gap:13}}>
              {form.type!=="sector"&&!editItem&&(
                <div style={{display:"flex",gap:0,background:T.bg,borderRadius:10,padding:4}}>
                  {[{key:"task",label:" 작업",desc:"기간이 있는 작업"},{key:"milestone",label:"◆ 마일스톤",desc:"한 시점의 일정"}].map(opt=>(
                    <button key={opt.key} onClick={()=>setForm(f=>({...f,type:opt.key,end:opt.key==="milestone"?f.start:f.end}))}
                      style={{flex:1,padding:isMobile?"10px 6px":"9px 6px",border:"none",borderRadius:8,cursor:"pointer",fontSize:12,
                        background:form.type===opt.key?(opt.key==="milestone"?"linear-gradient(135deg,#f59e0b,#fbbf24)":"linear-gradient(135deg,#6366f1,#8b5cf6)"):"transparent",
                        color:form.type===opt.key?"#fff":T.textMuted,fontWeight:form.type===opt.key?700:500,transition:"all 0.15s"}}>
                      <div>{opt.label}</div>
                      <div style={{fontSize:10,fontWeight:400,marginTop:2,opacity:0.8}}>{opt.desc}</div>
                    </button>
                  ))}
                </div>
              )}
              <F label="이름"><input value={form.name} onChange={e=>{setForm({...form,name:e.target.value}); if(formError) setFormError("");}} style={SB.inp} autoFocus/></F>
              {form.type!=="sector"&&(<>
                <div style={{display:"grid",gridTemplateColumns:form.type==="milestone"?"1fr":"1fr 1fr",gap:10}}>
                  <F label={form.type==="milestone"?"날짜":"시작일"}>
                    <input type="date" value={form.start} onChange={e=>setForm({...form,start:e.target.value,end:form.type==="milestone"?e.target.value:form.end})} style={SB.inp}/>
                    <div style={{fontSize:11,color:form.start?"#818cf8":T.textFaint,marginTop:4,fontFamily:"ui-monospace, monospace",fontWeight:600,letterSpacing:"0.02em"}}>
                      📅 {form.start || "선택안됨"}
                      <span style={{fontSize:9,color:T.textFaint,fontWeight:400,marginLeft:6}}>(연-월-일)</span>
                    </div>
                  </F>
                  {form.type!=="milestone"&&
                    <F label="종료일">
                      <input type="date" value={form.end} onChange={e=>setForm({...form,end:e.target.value})} style={SB.inp}/>
                      <div style={{fontSize:11,color:form.end?"#818cf8":T.textFaint,marginTop:4,fontFamily:"ui-monospace, monospace",fontWeight:600,letterSpacing:"0.02em"}}>
                        📅 {form.end || "선택안됨"}
                        <span style={{fontSize:9,color:T.textFaint,fontWeight:400,marginLeft:6}}>(연-월-일)</span>
                      </div>
                    </F>
                  }
                </div>
                <F label="섹터">
                  <select value={form.parentId||""} onChange={e=>setForm({...form,parentId:+e.target.value})} style={SB.inp}>
                    {(() => {
                      // Group sectors by category for clarity
                      const cats = items.filter(i=>i.type==="category").slice().sort((a,b)=>(a.order??0)-(b.order??0));
                      const opts = [];
                      cats.forEach(cat => {
                        const sectorsInCat = items.filter(i=>i.type==="sector" && i.parentId===cat.id)
                          .slice().sort((a,b)=>(a.order??0)-(b.order??0));
                        if (sectorsInCat.length > 0) {
                          opts.push(
                            <optgroup key={cat.id} label={"📁 " + cat.name}>
                              {sectorsInCat.map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
                            </optgroup>
                          );
                        }
                      });
                      return opts;
                    })()}
                  </select>
                  {(() => {
                    // Show which category this sector belongs to
                    const sector = items.find(s => s.id === form.parentId && s.type === "sector");
                    const cat = sector ? items.find(c => c.id === sector.parentId && c.type === "category") : null;
                    if (cat) {
                      return <div style={{fontSize:10,color:T.textFaint,marginTop:4,paddingLeft:2}}>📁 대분류: <span style={{color:"#818cf8",fontWeight:600}}>{cat.name}</span> (자동)</div>;
                    }
                    return null;
                  })()}
                </F>
                {form.type==="task"&&(<>
                  <F label={'진행률: ' + (form.progress||0) + '%'}>
                    <input type="range" min={0} max={100} value={form.progress||0} onChange={e=>setForm({...form,progress:+e.target.value})} style={{accentColor:"#6366f1",height:30}}/>
                  </F>
                </>)}
                <F label="링크 URL">
                  <input value={form.link||""} onChange={e=>setForm({...form,link:e.target.value})} style={SB.inp} placeholder="https://..."/>
                </F>
                {/* ─── 담당자 (Assignee) — Teams people picker, mention-style ─── */}
                <F label="담당자">
                  {(() => {
                    const assignees = Array.isArray(form.assignees) ? form.assignees : [];
                    const removeAssignee = (key) => {
                      setForm(f => ({...f, assignees: (f.assignees || []).filter(a => (a.id || a.email) !== key)}));
                    };
                    const addAssignees = (incoming) => {
                      if (!Array.isArray(incoming) || incoming.length === 0) return;
                      setForm(f => {
                        const cur = Array.isArray(f.assignees) ? f.assignees : [];
                        const seen = new Set(cur.map(a => a.id || a.email));
                        // Normalise picker output — Teams returns objectId/displayName/email,
                        // but we keep our own shape so legacy/manual entries are uniform.
                        const norm = incoming
                          .map(p => ({
                            id: p.objectId || p.id || p.email || "",
                            name: p.displayName || p.name || (p.email || "").split("@")[0] || "이름 없음",
                            email: p.email || "",
                          }))
                          .filter(a => a.id || a.email)
                          .filter(a => !seen.has(a.id || a.email));
                        return {...f, assignees: [...cur, ...norm]};
                      });
                    };
                    const openPicker = async () => {
                      try {
                        const picked = await pickTeamsPeople({
                          title: "담당자 선택 (멘션)",
                          setSelected: assignees.map(a => a.id).filter(Boolean),
                          singleSelect: false,
                        });
                        addAssignees(picked);
                      } catch(e) {
                        // Teams SDK unavailable → manual entry fallback so the
                        // feature still works in the web preview / dev environment.
                        const raw = window.prompt(
                          "Teams 멘션을 사용할 수 없습니다.\n수동 입력으로 담당자를 추가합니다.\n\n형식: 이름 <email@domain.com>\n여러 명은 쉼표로 구분하세요.",
                          ""
                        );
                        if (!raw || !raw.trim()) return;
                        const parts = raw.split(",").map(s => s.trim()).filter(Boolean);
                        const manual = parts.map(p => {
                          const m = p.match(/^(.+?)\s*<\s*(.+?)\s*>\s*$/);
                          if (m) return { name: m[1].trim(), email: m[2].trim() };
                          // Bare email
                          if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(p)) return { name: p.split("@")[0], email: p };
                          return null;
                        }).filter(Boolean);
                        if (manual.length === 0) {
                          alert("입력 형식이 올바르지 않습니다.");
                          return;
                        }
                        addAssignees(manual);
                      }
                    };
                    return (
                      <div style={{marginTop:4}}>
                        {assignees.length > 0 && (
                          <div style={{display:"flex",flexWrap:"wrap",gap:5,marginBottom:8}}>
                            {assignees.map(a => {
                              const key = a.id || a.email;
                              const initial = (a.name || a.email || "?").trim().charAt(0).toUpperCase();
                              return (
                                <div key={key}
                                  title={a.email || ""}
                                  style={{
                                    display:"inline-flex",alignItems:"center",gap:6,
                                    background:"rgba(99,102,241,0.14)",
                                    border:"1px solid rgba(99,102,241,0.4)",
                                    borderRadius:14, padding:"3px 4px 3px 3px",
                                    fontSize:11, color:T.text,
                                  }}>
                                  <span style={{
                                    width:18, height:18, borderRadius:"50%",
                                    background:"linear-gradient(135deg,#6366f1,#8b5cf6)",
                                    color:"#fff", fontSize:10, fontWeight:700,
                                    display:"flex", alignItems:"center", justifyContent:"center",
                                  }}>{initial}</span>
                                  <span style={{maxWidth:140,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>
                                    {a.name || a.email}
                                  </span>
                                  <span onClick={()=>removeAssignee(key)}
                                    style={{cursor:"pointer",color:T.textFaint,fontSize:14,padding:"0 4px",lineHeight:1}}>×</span>
                                </div>
                              );
                            })}
                          </div>
                        )}
                        <button type="button" onClick={openPicker}
                          style={{
                            display:"inline-flex",alignItems:"center",gap:6,
                            padding:"6px 10px",
                            background:"transparent",
                            border:`1px dashed ${T.border}`,
                            color:T.textMuted, borderRadius:8,
                            fontSize:11, fontWeight:600, cursor:"pointer",
                            transition:"all 0.15s",
                          }}
                          onMouseEnter={e=>{e.currentTarget.style.background="rgba(99,102,241,0.08)";e.currentTarget.style.color=T.text;}}
                          onMouseLeave={e=>{e.currentTarget.style.background="transparent";e.currentTarget.style.color=T.textMuted;}}
                        >
                          <span style={{fontSize:12,fontWeight:700}}>@</span> 담당자 추가 (Teams 멘션)
                        </button>
                        <div style={{fontSize:10,color:T.textFaint,marginTop:6,lineHeight:1.4}}>
                          저장 시 새로 추가된 담당자에게 메일 발송 여부를 안내합니다.
                        </div>
                      </div>
                    );
                  })()}
                </F>
                <F label="의존 작업 (복수 선택 가능)">
                  {(()=>{
                    const candidates = items.filter(i => i.type !== "sector" && i.id !== (editItem?.id || -1));
                    const sectors = items.filter(i => i.type === "sector").slice().sort((a,b)=>(a.order??0)-(b.order??0));
                    const deps = form.deps || [];
                    const depSearch = (form._depSearch || "").toLowerCase().trim();
                    const toggle = id => {
                      setForm(f => ({...f, deps: f.deps?.includes(id) ? f.deps.filter(d => d !== id) : [...(f.deps || []), id]}));
                    };
                    return (
                      <div style={{marginTop:4}}>
                        {deps.length > 0 && (
                          <div style={{display:"flex",flexWrap:"wrap",gap:5,marginBottom:8}}>
                            {deps.map(depId => {
                              const dep = items.find(i=>i.id===depId);
                              if (!dep) return null;
                              const dc = effectiveColor(dep, items);
                              return (
                                <span key={depId} style={{display:"flex",alignItems:"center",gap:5,
                                    background:T.border,borderRadius:20,padding:"3px 8px 3px 10px",
                                    fontSize:12,color:T.textSub,border:`1px solid ${hexAlpha(dc,0.4)}`}}>
                                  <span style={{width:6,height:6,borderRadius:"50%",background:dc}}/>
                                  {dep.name}
                                  <span onClick={()=>toggle(depId)} style={{cursor:"pointer",color:T.textFaint,fontSize:14,marginLeft:2}}>×</span>
                                </span>
                              );
                            })}
                          </div>
                        )}
                        {/* Search input */}
                        <div style={{display:"flex",alignItems:"center",gap:6,background:T.bgInput,border:`1px solid ${T.border}`,borderRadius:7,padding:"6px 10px",marginBottom:6}}>
                          <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke={T.textFaint} strokeWidth="2"><circle cx="11" cy="11" r="7"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
                          <input value={form._depSearch || ""} onChange={e=>setForm(f=>({...f, _depSearch: e.target.value}))}
                            placeholder="작업 검색..."
                            style={{flex:1,background:"transparent",border:"none",outline:"none",color:T.text,fontSize:12,fontFamily:"inherit",minWidth:0}}/>
                          {depSearch && (
                            <span onClick={()=>setForm(f=>({...f, _depSearch: ""}))} style={{cursor:"pointer",color:T.textFaint,fontSize:14,padding:"0 2px",lineHeight:1}}>×</span>
                          )}
                        </div>
                        <div style={{background:T.bg,border:`1px solid ${T.border}`,borderRadius:8,maxHeight:180,overflowY:"auto",WebkitOverflowScrolling:"touch"}}>
                          {candidates.length === 0 && (
                            <div style={{padding:"10px 12px",fontSize:12,color:T.textFaint}}>의존 가능한 작업이 없습니다</div>
                          )}
                          {(() => {
                            const visibleSectors = sectors.map(sector => {
                              const kids = candidates.filter(i => i.parentId === sector.id).slice().sort((a,b)=>(a.order??0)-(b.order??0));
                              const filteredKids = depSearch
                                ? kids.filter(k => k.name.toLowerCase().includes(depSearch) || sector.name.toLowerCase().includes(depSearch))
                                : kids;
                              return { sector, kids: filteredKids };
                            }).filter(s => s.kids.length > 0);
                            if (depSearch && visibleSectors.length === 0) {
                              return <div style={{padding:"10px 12px",fontSize:12,color:T.textFaint}}>검색 결과가 없습니다</div>;
                            }
                            return visibleSectors.map(({ sector, kids }) => (
                              <div key={sector.id}>
                                <div style={{padding:"5px 10px",fontSize:10,fontWeight:700,
                                    color:sector.color,textTransform:"uppercase",letterSpacing:"0.05em",
                                    background:"rgba(255,255,255,0.02)",borderBottom:`1px solid ${T.borderLight}`,
                                    display:"flex",alignItems:"center",gap:6}}>
                                  <span style={{width:6,height:6,borderRadius:1,background:sector.color}}/>
                                  {sector.name}
                                </div>
                                {kids.map(task => {
                                  const tc = effectiveColor(task, items);
                                  const checked = deps.includes(task.id);
                                  return (
                                    <div key={task.id} onClick={()=>toggle(task.id)}
                                      style={{display:"flex",alignItems:"center",gap:10,
                                        padding:"8px 12px",cursor:"pointer",
                                        background: checked ? "rgba(99,102,241,0.1)" : "transparent",
                                        borderBottom:`1px solid ${T.borderLight}`,
                                        transition:"background 0.1s"}}>
                                      <div style={{width:16,height:16,borderRadius:4,flexShrink:0,
                                          border: checked ? "none" : "1.5px solid #334155",
                                          background: checked ? "#6366f1" : "transparent",
                                          display:"flex",alignItems:"center",justifyContent:"center",
                                          transition:"all 0.12s"}}>
                                        {checked && <svg width="9" height="7" viewBox="0 0 9 7" fill="none"><path d="M1 3.5L3.5 6L8 1" stroke="#fff" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"/></svg>}
                                      </div>
                                      {task.type==="milestone"
                                        ? <div style={{width:8,height:8,background:tc,transform:"rotate(45deg)",borderRadius:1,flexShrink:0}}/>
                                        : <div style={{width:7,height:7,borderRadius:"50%",background:tc,flexShrink:0}}/>
                                      }
                                      <div style={{flex:1,minWidth:0}}>
                                        <div style={{fontSize:12,color: checked?T.text:T.textSub, fontWeight: checked?600:400,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>
                                          {task.name}
                                        </div>
                                        {task.start && (
                                          <div style={{fontSize:10,color:T.textFaint,marginTop:1}}>
                                            {task.type==="milestone" ? toISO(task.start) : toISO(task.start) + " ~ " + (task.end?toISO(task.end):"")}
                                          </div>
                                        )}
                                      </div>
                                    </div>
                                  );
                                })}
                              </div>
                            ));
                          })()}
                        </div>
                        {deps.length > 0 && (
                          <div style={{marginTop:6,textAlign:"right"}}>
                            <span onClick={()=>setForm(f=>({...f,deps:[]}))}
                              style={{fontSize:11,color:T.textFaint,cursor:"pointer",textDecoration:"underline"}}>
                              전체 해제
                            </span>
                          </div>
                        )}
                      </div>
                    );
                  })()}
                </F>
              </>)}
              {form.type==="sector"?(
                <>
                  <F label="대분류">
                    <select value={form.parentId||""} onChange={e=>setForm({...form,parentId:+e.target.value})} style={SB.inp}>
                      {items.filter(i=>i.type==="category").slice().sort((a,b)=>(a.order??0)-(b.order??0)).map(c=>(
                        <option key={c.id} value={c.id}>{c.name}</option>
                      ))}
                    </select>
                  </F>
                  <F label="섹터 색상">
                    <div style={{display:"flex",gap:8,flexWrap:"wrap",marginTop:4}}>
                      {SECTOR_COLORS.map(c=>(
                        <div key={c} onClick={()=>setForm({...form,color:c})}
                          style={{width:isMobile?32:26,height:isMobile?32:26,borderRadius:"50%",background:c,cursor:"pointer",
                            border:form.color===c?"3px solid #fff":"3px solid transparent",
                            boxShadow:form.color===c?`0 0 0 2px ${c}`:"none",transition:"all 0.15s"}}/>
                      ))}
                    </div>
                  </F>
                </>
              ):(
                <F label="색상 미리보기 (섹터 기반 자동 적용)">
                  {(()=>{
                    const sector=items.find(i=>i.id===form.parentId&&i.type==="sector");
                    if (!sector) return <span style={{fontSize:12,color:T.textFaint}}>섹터를 먼저 선택하세요</span>;
                    const siblings=items.filter(i=>i.parentId===sector.id&&i.type!=="sector").sort((a,b)=>(a.order??0)-(b.order??0));
                    const isNew=!editItem;
                    const cnt=isNew?siblings.length+1:siblings.length;
                    const idx=isNew?siblings.length:siblings.findIndex(i=>i.id===editItem?.id);
                    const pc=childColor(sector.color,Math.max(idx,0),cnt);
                    return (<div style={{display:"flex",alignItems:"center",gap:10,marginTop:4}}>
                      <div style={{width:28,height:28,borderRadius:7,background:pc,border:`2px solid ${T.border}`,flexShrink:0}}/>
                      <div>
                        <div style={{fontSize:12,color:T.textSub}}>섹터 <span style={{color:sector.color,fontWeight:700}}>{sector.name}</span> 안의 색상</div>
                        <div style={{fontSize:11,color:T.textFaint,marginTop:2}}>같은 섹터 안에서 자동으로 다른 음영이 적용됩니다</div>
                      </div>
                    </div>);
                  })()}
                </F>
              )}
            </div>
            {formError && (
              <div style={{
                marginTop:14, padding:"8px 12px",
                background:"rgba(248,113,113,0.12)",
                border:"1px solid rgba(248,113,113,0.4)",
                borderRadius:8, color:"#f87171",
                fontSize:12, fontWeight:600,
                display:"flex", alignItems:"center", gap:6,
              }}>
                <span>⚠</span> {formError}
              </div>
            )}
            <div style={{display:"flex",gap:8,marginTop:20,justifyContent:"flex-end"}}>
              {editItem && (
                <button onClick={()=>{
                  const doDel = () => {
                    deleteItem(editItem.id);
                    setSelectedId(null);
                    setShowModal(false);
                  };
                  if (window.appConfirm) {
                    window.appConfirm(`"${editItem.name}" 삭제하시겠습니까?`, doDel);
                  } else {
                    let ok = false;
                    try { ok = window.confirm(`"${editItem.name}" 삭제하시겠습니까?`); }
                    catch(e) { ok = true; }
                    if (ok) doDel();
                  }
                }} style={{...SB.sec,color:"#f87171",border:"1px solid rgba(248,113,113,0.4)",marginRight:"auto"}}>
                  <TrashIcon size={11} color="#f87171"/> 삭제
                </button>
              )}
              <button onClick={()=>{ setShowModal(false); setFormError(""); }} style={SB.nav}>취소</button>
              <button onClick={saveForm} style={SB.add}>저장</button>
            </div>
          </div>
        </div>
      )}

      {/* Row drag ghost */}
      {rowDragState && (() => {
        const draggingItem = flatRows.find(r => r.item.id === rowDragState.id)?.item;
        if (!draggingItem) return null;
        const isSector = draggingItem.type === "sector";
        const c = isSector ? draggingItem.color : effectiveColor(draggingItem, items);
        return (
          <div style={{
            position:"fixed",
            left: rowDragPos.x - 20,
            top: rowDragPos.y - baseRowHeight / 2,
            width: isMobile ? viewport.w - 30 : leftPanelWidth - 8,
            height: baseRowHeight,
            pointerEvents:"none", zIndex:1000,
            background: isSector ? "#1a2540" : "#161f35",
            border:`2px solid ${c}`,
            borderRadius:8,
            boxShadow:`0 8px 32px rgba(0,0,0,0.5), 0 0 0 1px ${hexAlpha(c,0.3)}`,
            display:"grid",
            gridTemplateColumns: isMobile ? GRID_COLS_MOBILE : GRID_COLS,
            alignItems:"center",
            padding: isMobile?"0 10px 0 6px":"0 6px 0 4px",
            gap: isMobile?6:0,
            opacity:0.92,
          }}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"center"}}>
              <GripIcon active={true}/>
            </div>
            {!isMobile && (
              <div style={{display:"flex",justifyContent:"center"}}>
                {isSector
                  ? <div style={{width:8,height:8,borderRadius:2,background:c}}/>
                  : draggingItem.type==="milestone"
                    ? <div style={{width:8,height:8,background:c,transform:"rotate(45deg)",borderRadius:1}}/>
                    : <div style={{width:7,height:7,borderRadius:"50%",background:c}}/>
                }
              </div>
            )}
            <div style={{paddingLeft: isSector?(isMobile?6:4):14, overflow:"hidden", display:"flex", alignItems:"center", gap:isMobile?7:5,
                borderLeft: isSector?"none":`2px solid ${T.border}`, marginLeft: isSector?0:4}}>
              {isMobile && (
                isSector
                  ? <div style={{width:9,height:9,borderRadius:2,background:c,flexShrink:0}}/>
                  : draggingItem.type==="milestone"
                    ? <div style={{width:9,height:9,background:c,transform:"rotate(45deg)",borderRadius:1,flexShrink:0}}/>
                    : <div style={{width:8,height:8,borderRadius:"50%",background:c,flexShrink:0}}/>
              )}
              <span style={{fontSize:isMobile?13:12, fontWeight:isSector?700:500, color:T.text,
                  whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis"}}>
                {draggingItem.name}
              </span>
            </div>
            <div style={{textAlign:"center",fontSize:11,color:T.textMuted}}>
              {draggingItem.start ? `${String(draggingItem.start.getMonth()+1).padStart(2,"0")}/${String(draggingItem.start.getDate()).padStart(2,"0")}` : ""}
            </div>
            <div style={{textAlign:"center",fontSize:11,color:T.textMuted}}>
              {draggingItem.end && draggingItem.type!=="milestone" ? `${String(draggingItem.end.getMonth()+1).padStart(2,"0")}/${String(draggingItem.end.getDate()).padStart(2,"0")}` : ""}
            </div>
            <div style={{textAlign:"center",fontSize:11,fontWeight:600,color:draggingItem.progress===100?"#10b981":"#6366f1"}}>
              {draggingItem.type==="task" ? draggingItem.progress+"%" : ""}
            </div>
          </div>
        );
      })()}

      {/* Drag mode toast */}
      {activeDragId && (() => {
        const item = items.find(i=>i.id===activeDragId);
        return (
          <div style={{position:"fixed",bottom:isMobile?80:22,left:"50%",transform:"translateX(-50%)",
            background:"rgba(99,102,241,0.97)",color:"#fff",padding:"10px 18px",borderRadius:40,
            fontSize:isMobile?12:13,fontWeight:700,zIndex:200,boxShadow:"0 4px 24px rgba(0,0,0,0.4)",
            display:"flex",alignItems:"center",gap:10,whiteSpace:"nowrap",pointerEvents:"none",
            maxWidth:"92vw",overflow:"hidden",textOverflow:"ellipsis"}}>
            <GripIcon active={true}/>
            "{item?.name}" 이동 중 — ▲▼ 또는 다른 행을 탭
          </div>
        );
      })()}

      {/* App zoom indicator (always visible, bottom-right) */}
      <div style={{position:"fixed",bottom:14,right:14,
        background:T.bgCard,color:T.textSub,padding:"6px 12px",borderRadius:20,
        fontSize:12,fontWeight:600,zIndex:150,
        border:`1px solid ${T.border}`,
        boxShadow:"0 4px 16px rgba(0,0,0,0.3)",
        display:"flex",alignItems:"center",gap:8,
        opacity: appZoom === 1.0 ? 0.55 : 1,
        transition:"opacity 0.2s"}}
        onMouseEnter={e=>e.currentTarget.style.opacity=1}
        onMouseLeave={e=>e.currentTarget.style.opacity = appZoom === 1.0 ? 0.55 : 1}>
        <span title="화면 배율">🔍 {Math.round(appZoom*100)}%</span>
        {appZoom !== 1.0 && (
          <button onClick={()=>setAppZoom(1.0)}
            title="원래 크기 (Ctrl+0)"
            style={{background:"transparent",border:"none",cursor:"pointer",color:T.textFaint,padding:"2px 4px",fontSize:10,fontWeight:600,borderRadius:4}}
            onMouseEnter={e=>e.currentTarget.style.background=hexAlpha("#6366f1",0.15)}
            onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
            초기화
          </button>
        )}
      </div>

      {/* In-app confirm dialog (replaces window.confirm which is blocked in Teams) */}
      {confirmDialog && (
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.6)",zIndex:400,
          display:"flex",alignItems:"center",justifyContent:"center",animation:"fadeIn 0.15s"}}
          onClick={()=>setConfirmDialog(null)}>
          <div style={{background:T.bgCard,border:`1px solid ${T.border}`,borderRadius:14,
            padding:"20px 22px",minWidth:280,maxWidth:420,boxShadow:"0 12px 40px rgba(0,0,0,0.5)"}}
            onClick={e=>e.stopPropagation()}>
            <div style={{fontSize:14,fontWeight:600,color:T.text,marginBottom:18,lineHeight:1.5,whiteSpace:"pre-wrap"}}>
              {confirmDialog.message}
            </div>
            <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
              <button onClick={()=>setConfirmDialog(null)} style={SB.nav}>취소</button>
              <button onClick={()=>{
                const fn = confirmDialog.onConfirm;
                setConfirmDialog(null);
                if (fn) fn();
              }} style={{...SB.add,background:"#ef4444",border:"1px solid #ef4444"}}>확인</button>
            </div>
          </div>
        </div>
      )}

      {/* Assignee email-notification dialog — shown after a save when one or
          more new assignees were added. Lets the user fire off a mailto:
          with the task details prefilled. */}
      {assigneeNotifyDialog && (() => {
        const dlg = assigneeNotifyDialog;
        const close = () => setAssigneeNotifyDialog(null);
        const recipientsWithEmail = (dlg.newAssignees || []).filter(a => a.email);
        const sendEmail = () => {
          if (recipientsWithEmail.length === 0) {
            alert("이메일 주소가 있는 담당자가 없습니다.");
            close();
            return;
          }
          const task = dlg.task || {};
          const fmt = d => {
            try {
              if (!d) return "";
              const dt = d instanceof Date ? d : new Date(d);
              if (isNaN(dt.getTime())) return "";
              return `${dt.getFullYear()}-${String(dt.getMonth()+1).padStart(2,"0")}-${String(dt.getDate()).padStart(2,"0")}`;
            } catch(e) { return ""; }
          };
          const lines = [];
          lines.push(`프로젝트: ${projectName}`);
          lines.push(`작업: ${task.name || "(이름 없음)"}`);
          if (task.type === "milestone") {
            if (task.start) lines.push(`날짜: ${fmt(task.start)}`);
          } else {
            if (task.start || task.end) lines.push(`기간: ${fmt(task.start)} ~ ${fmt(task.end)}`);
          }
          if (typeof task.progress === "number") lines.push(`진행률: ${task.progress}%`);
          if (task.link) lines.push(`링크: ${task.link}`);
          lines.push("");
          lines.push("위 작업의 담당자로 지정되었습니다. 확인 부탁드립니다.");
          const subject = `[${projectName}] 작업 담당자 지정 - ${task.name || ""}`;
          const body = lines.join("\n");
          const to = recipientsWithEmail.map(a => a.email).join(",");
          const url = `mailto:${encodeURIComponent(to)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
          // window.open in Teams iframe can be blocked — fall back to location
          const w = window.open(url, "_blank");
          if (!w) { try { window.location.href = url; } catch(e) {} }
          close();
        };
        return (
          <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.6)",zIndex:400,
            display:"flex",alignItems:"center",justifyContent:"center",animation:"fadeIn 0.15s"}}
            onClick={close}>
            <div style={{background:T.bgCard,border:`1px solid ${T.border}`,borderRadius:14,
              padding:"20px 22px",minWidth:300,maxWidth:460,boxShadow:"0 12px 40px rgba(0,0,0,0.5)"}}
              onClick={e=>e.stopPropagation()}>
              <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:12}}>
                <span style={{fontSize:18}}>📧</span>
                <div style={{fontSize:14,fontWeight:700,color:T.text}}>담당자 메일 알림</div>
              </div>
              <div style={{fontSize:12,color:T.textSub,marginBottom:10,lineHeight:1.5}}>
                새로 추가된 담당자에게 작업 알림 메일을 발송하시겠습니까?
              </div>
              <div style={{
                background:T.bg, border:`1px solid ${T.border}`, borderRadius:8,
                padding:"8px 10px", marginBottom:14, maxHeight:160, overflowY:"auto",
              }}>
                {(dlg.newAssignees || []).map(a => {
                  const key = a.id || a.email || a.name;
                  return (
                    <div key={key} style={{
                      display:"flex", alignItems:"center", gap:8,
                      padding:"4px 0", fontSize:12, color:T.text,
                    }}>
                      <span style={{
                        width:20, height:20, borderRadius:"50%",
                        background:"linear-gradient(135deg,#6366f1,#8b5cf6)",
                        color:"#fff", fontSize:10, fontWeight:700,
                        display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0,
                      }}>{(a.name || a.email || "?").trim().charAt(0).toUpperCase()}</span>
                      <span style={{fontWeight:600}}>{a.name || "(이름 없음)"}</span>
                      <span style={{color:T.textFaint, fontSize:11}}>
                        {a.email || "이메일 없음"}
                      </span>
                    </div>
                  );
                })}
              </div>
              {recipientsWithEmail.length === 0 && (
                <div style={{
                  fontSize:11, color:"#f59e0b", marginBottom:12,
                  padding:"6px 8px", background:"rgba(245,158,11,0.1)",
                  border:"1px solid rgba(245,158,11,0.3)", borderRadius:6,
                }}>
                  ⚠ 이메일 주소가 등록된 담당자가 없어 메일 발송이 불가능합니다.
                </div>
              )}
              <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
                <button onClick={close} style={SB.nav}>발송 안 함</button>
                <button onClick={sendEmail}
                  disabled={recipientsWithEmail.length === 0}
                  style={{
                    ...SB.add,
                    opacity: recipientsWithEmail.length === 0 ? 0.5 : 1,
                    cursor: recipientsWithEmail.length === 0 ? "not-allowed" : "pointer",
                  }}>
                  메일 보내기
                </button>
              </div>
            </div>
          </div>
        );
      })()}

      {/* Linking toast */}
      {linkingFrom&&(
        <div style={{position:"fixed",bottom:isMobile?80:22,left:"50%",transform:"translateX(-50%)",
          background:"#fbbf24",color:"#000",padding:"10px 18px",borderRadius:40,
          fontSize:13,fontWeight:700,zIndex:200,boxShadow:"0 4px 24px rgba(0,0,0,0.4)",
          display:"flex",alignItems:"center",gap:8,cursor:"pointer"}}
          onClick={()=>setLinkingFrom(null)}>
          <LinkIcon size={13} color="#000"/>
          연결할 작업을 클릭하세요
          <span style={{textDecoration:"underline",marginLeft:4}}>취소</span>
        </div>
      )}
    </div>
  );
}
