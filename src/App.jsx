// @ts-nocheck
import { useState, useRef, useEffect, useMemo, useCallback } from "react";

const APP_VERSION = "1.0.4"; // 배포할때마다 올려주세요

// --- Supabase ----------------------------------------------------------------
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
    Prefer: "return=minimal",
  };
  // Try PATCH first (update if exists)
  const patch = await fetch(`${SUPABASE_URL}/rest/v1/gantt_data?id=eq.${channelId}`, {
    method: "PATCH", headers, body,
  });
  if (patch.ok) return; // updated OK
  // If no row exists (204 with no rows updated), INSERT
  const post = await fetch(`${SUPABASE_URL}/rest/v1/gantt_data`, {
    method: "POST", headers, body,
  });
  if (!post.ok) {
    const err = await post.text();
    throw new Error("Save failed: " + post.status + " " + err);
  }
}

// Teams에서 채널 ID 가져오기 (없으면 기본값)
function getChannelId() {
  try {
    // Teams JS SDK가 로드돼 있으면 context에서 가져옴
    if (window.microsoftTeams) {
      return new Promise(resolve => {
        window.microsoftTeams.app.initialize().then(() => {
          window.microsoftTeams.app.getContext().then(ctx => {
            resolve(ctx.channel?.id || ctx.chat?.id || "default");
          }).catch(() => resolve("default"));
        }).catch(() => resolve("default"));
      });
    }
  } catch(e) {}
  return Promise.resolve("default");
}


// --- SVG Icons ----------------------------------------------------------------
const Icon = ({ d, size=14, stroke="#94a3b8", fill="none", sw=1.8, ...rest }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill={fill} stroke={stroke}
    strokeWidth={sw} strokeLinecap="round" strokeLinejoin="round" style={{ flexShrink:0 }} {...rest}>
    {Array.isArray(d) ? d.map((p,i)=><path key={i} d={p}/>) : <path d={d}/>}
  </svg>
);

const EditIcon = ({ size=14, color="#64748b" }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/>
    <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>
  </svg>
);

const TrashIcon = ({ size=14, color="#f87171" }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/>
    <path d="M10 11v6M14 11v6"/><path d="M9 6V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2"/>
  </svg>
);

const PlusIcon = ({ size=14, color="#94a3b8" }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/>
  </svg>
);

const LinkIcon = ({ size=13, color="#6366f1" }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/>
    <path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/>
  </svg>
);

const ChevronDown = ({ size=12 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="#475569" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><path d="M6 9l6 6 6-6"/></svg>
);

const ChevronRight = ({ size=12 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="#475569" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><path d="M9 18l6-6-6-6"/></svg>
);

const ExcelIcon = ({ size=14 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
    <polyline points="14 2 14 8 20 8"/>
    <line x1="12" y1="11" x2="12" y2="17"/>
    <polyline points="9 14 12 17 15 14"/>
  </svg>
);

const MilestoneIcon = ({ size=13, color="#818cf8" }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill={color} stroke={color} strokeWidth="1">
    <polygon points="12,2 22,12 12,22 2,12"/>
  </svg>
);

const MenuIcon = ({ size=18 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><line x1="3" y1="6" x2="21" y2="6"/><line x1="3" y1="12" x2="21" y2="12"/><line x1="3" y1="18" x2="21" y2="18"/></svg>
);

const CloseIcon = ({ size=16 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
);

// --- Color helpers ------------------------------------------------------------
const SECTOR_COLORS = ["#6366f1","#f43f5e","#10b981","#f59e0b","#3b82f6","#8b5cf6","#06b6d4","#ec4899"];

function hexToRgb(hex) {
  const h = hex.replace("#","");
  return { r:parseInt(h.slice(0,2),16), g:parseInt(h.slice(2,4),16), b:parseInt(h.slice(4,6),16) };
}

function lighten(hex, amount) {
  const {r,g,b} = hexToRgb(hex);
  const m = c => Math.round(c+(255-c)*amount);
  const x = c => c.toString(16).padStart(2,"0");
  return `#${x(m(r))}${x(m(g))}${x(m(b))}`;
}

function hexAlpha(hex, alpha) {
  if (!hex || !hex.startsWith("#")) return hex||"#6366f1";
  return hex + Math.round(alpha*255).toString(16).padStart(2,"0");
}

function childColor(sectorColor, index, total) {
  const t = total<=1 ? 0.28 : 0.14+(index/(total-1))*0.42;
  return lighten(sectorColor, t);
}

function effectiveColor(item, allItems) {
  if (item.type==="sector") return item.color;
  const sector = allItems.find(i=>i.id===item.parentId&&i.type==="sector");
  if (!sector) return "#6366f1";
  const siblings = allItems.filter(i=>i.parentId===sector.id&&i.type!=="sector");
  const idx = siblings.findIndex(i=>i.id===item.id);
  return childColor(sector.color, Math.max(idx,0), siblings.length);
}

// --- Constants ----------------------------------------------------------------
const ROW_HEIGHT = 52;
const DETAIL_HEIGHT = 110;

// --- Theme -------------------------------------------------------------------
const THEMES = {
  dark: {
    bg:          "#0f1117",
    bgDeep:      "#0b0e17",
    bgCard:      "#161b27",
    bgRow:       "#131820",
    bgRowHover:  "#1a1f2e",
    bgHeader:    "#0d111a",
    bgInput:     "#0f1117",
    bgSelected:  "#0f1b30",
    border:      "#1e2535",
    borderLight: "#1a2030",
    text:        "#e2e8f0",
    textSub:     "#94a3b8",
    textMuted:   "#64748b",
    textFaint:   "#475569",
    textDim:     "#334155",
  },
  light: {
    bg:          "#f0f2f7",
    bgDeep:      "#e8ebf2",
    bgCard:      "#ffffff",
    bgRow:       "#f7f8fc",
    bgRowHover:  "#eef0f8",
    bgHeader:    "#e8ebf4",
    bgInput:     "#ffffff",
    bgSelected:  "#eef2ff",
    border:      "#dde1ee",
    borderLight: "#e8ebf4",
    text:        "#1a1f36",
    textSub:     "#3d4466",
    textMuted:   "#5a6282",
    textFaint:   "#7b84a3",
    textDim:     "#a0a8c0",
  },
};

// Module-level theme ref - updated by GanttApp, used by child components
let T = THEMES.dark;
function setGlobalTheme(key) { T = THEMES[key]; } // height of expanded detail panel
const COL_NAME = "1fr";
const COL_START = "72px";
const COL_END = "72px";
const COL_PROG = "52px";
const GRID_COLS = `36px 10px ${COL_NAME} ${COL_START} ${COL_END} ${COL_PROG}`;

const ZOOM_LEVELS = [
  { key:"fit",     label:"전체", dayWidth:0,  totalDays:0   },
  { key:"year",    label:"연",   dayWidth:6,  totalDays:365 },
  { key:"quarter", label:"분기", dayWidth:12, totalDays:120 },
  { key:"month",   label:"월",   dayWidth:24, totalDays:60  },
  { key:"day",     label:"일",   dayWidth:38, totalDays:42  },
];

const today = new Date(); today.setHours(0,0,0,0);
const addDays = (d,n) => { const x=new Date(d); x.setDate(x.getDate()+n); return x; };
const dateDiff = (a,b) => Math.round((b-a)/86400000);
const toISO = d => d.toISOString().slice(0,10);
const parseDate = s => { if (!s || !s.trim()) return null; const d=new Date(s); if (isNaN(d.getTime())) return null; d.setHours(0,0,0,0); return d; };
let _uid=20; const uid=()=>++_uid;

// --- Sort helpers -------------------------------------------------------------
function sortedByDate(raw) {
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
  return [...sectors,...children];
}

// --- Initial data -------------------------------------------------------------
const initialItems = sortedByDate([
  { id:1, type:"sector", name:"기획", color:SECTOR_COLORS[0], collapsed:false },
  { id:2, type:"task", parentId:1, name:"요구사항 분석", start:addDays(today,-5), end:addDays(today,2), progress:70 },
  { id:3, type:"task", parentId:1, name:"기술 스택 결정", start:addDays(today,1), end:addDays(today,5), progress:0 },
  { id:4, type:"milestone", parentId:1, name:"기획 완료", start:addDays(today,4), end:addDays(today,4) },
  { id:5, type:"sector", name:"개발", color:SECTOR_COLORS[5], collapsed:false },
  { id:6, type:"task", parentId:5, name:"UI/UX 디자인", start:addDays(today,5), end:addDays(today,12), progress:0 },
  { id:7, type:"task", parentId:5, name:"프론트엔드 개발", start:addDays(today,10), end:addDays(today,22), progress:0 },
  { id:8, type:"task", parentId:5, name:"백엔드 API 개발", start:addDays(today,5), end:addDays(today,20), progress:0 },
  { id:9, type:"sector", name:"출시", color:SECTOR_COLORS[3], collapsed:false },
  { id:10, type:"task", parentId:9, name:"QA 테스트", start:addDays(today,20), end:addDays(today,27), progress:0 },
  { id:11, type:"milestone", parentId:9, name:"제품 런칭", start:addDays(today,28), end:addDays(today,28) },
]);

// --- Dep arrows ---------------------------------------------------------------
function DepArrows({ flatRows, viewStart, DAY_WIDTH, getRowTop }) {
  const arrows = [];
  flatRows.forEach((row, toIdx) => {
    (row.item.deps || []).forEach(depId => {
      const fromIdx = flatRows.findIndex(r => r.item.id === depId);
      if (fromIdx < 0) return;
      const from = flatRows[fromIdx].item;
      if (!from.end || !row.item.start) return;
      const x1 = (dateDiff(effectiveViewStart, from.end) + 1) * DAY_WIDTH;
      const y1 = (getRowTop ? getRowTop(fromIdx) : fromIdx * ROW_HEIGHT) + ROW_HEIGHT / 2;
      const x2 = dateDiff(effectiveViewStart, row.item.start) * DAY_WIDTH;
      const y2 = (getRowTop ? getRowTop(toIdx) : toIdx * ROW_HEIGHT) + ROW_HEIGHT / 2;
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

// --- Drag grip icon -----------------------------------------------------------
const GripIcon = ({active}) => (
  <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
    {[2,6,10].map(x=>[2,6,10].map(y=>(
      <circle key={x+'-'+y} cx={x} cy={y} r="1.3" fill={active?"#818cf8":"#475569"}/>
    )))}
  </svg>
);

const DragBtn = ({active, onActivate, onMoveUp, onMoveDown, canUp, canDown, onMouseDownDrag}) => (
  <div style={{display:"flex", alignItems:"center", gap:1, flexShrink:0}}>
    <button
      data-dragbtn="1"
      onClick={e=>{ e.stopPropagation(); e.preventDefault(); onActivate(); }}
      onMouseDown={e=>{ if(e.button===0) onMouseDownDrag(e); }}
      title="드래그하여 이동 / 클릭하여 이동 모드"
      style={{
        width:28, height:28, flexShrink:0,
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
          style={{width:22,height:13,display:"flex",alignItems:"center",justifyContent:"center",
            background:canUp?"rgba(99,102,241,0.3)":"rgba(255,255,255,0.04)",
            border:"none",borderRadius:"3px 3px 0 0",cursor:canUp?"pointer":"default",padding:0}}>
          <svg width="9" height="6" viewBox="0 0 9 6"><path d="M4.5 0L9 6H0z" fill={canUp?"#818cf8":"#334155"}/></svg>
        </button>
        <button data-dragbtn="1"
          onClick={e=>{ e.stopPropagation(); e.preventDefault(); onMoveDown(); }}
          disabled={!canDown}
          style={{width:22,height:13,display:"flex",alignItems:"center",justifyContent:"center",
            background:canDown?"rgba(99,102,241,0.3)":"rgba(255,255,255,0.04)",
            border:"none",borderRadius:"0 0 3px 3px",cursor:canDown?"pointer":"default",padding:0}}>
          <svg width="9" height="6" viewBox="0 0 9 6"><path d="M4.5 6L0 0H9z" fill={canDown?"#818cf8":"#334155"}/></svg>
        </button>
      </div>
    )}
  </div>
);

function Ib({onClick,title,children,sz=28}) {
  return (
    <button title={title} onClick={onClick}
      style={{width:sz,height:sz,background:"transparent",border:"none",cursor:"pointer",borderRadius:6,display:"flex",alignItems:"center",justifyContent:"center",padding:0}}
      onMouseEnter={e=>e.currentTarget.style.background="rgba(255,255,255,0.07)"}
      onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
      {children}
    </button>
  );
}

function F({label,children}) {
  return <label style={{display:"flex",flexDirection:"column",gap:5,fontSize:11,color:"#64748b",fontWeight:600,textTransform:"uppercase",letterSpacing:"0.05em"}}>{label}{children}</label>;
}

// --- Left panel row -----------------------------------------------------------
function LeftRow({item,hovered,onHover,onEdit,onDelete,onToggle,allItems,isDragging,isSelected,
  activeDragId,onActivateDrag,onMoveUp,onMoveDown,canUp,canDown,onTapWhileDragging,onSelect,
  onMouseDownDrag,dropIndicator,themeKey="dark"}) {
  const T = THEMES[themeKey];
  const isSector = item.type==="sector";
  const c = isSector ? item.color : effectiveColor(item,allItems||[]);
  const fmtD = d => d ? `${String(d.getFullYear()).slice(2)}/${String(d.getMonth()+1).padStart(2,'0')}/${String(d.getDate()).padStart(2,'0')}` : "";
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
  return (
    <>
      <div onMouseEnter={()=>onHover(item.id)} onMouseLeave={()=>onHover(null)}
        onClick={handleClick}
        style={{
          display:"grid", gridTemplateColumns:GRID_COLS,
          alignItems:"center", height:ROW_HEIGHT,
          padding:"0 6px 0 4px",
          borderBottom: isSelected ? "none" : "1px solid #1a2030",
          borderLeft: isSelected ? `3px solid ${c}` : "3px solid transparent",
          background: rowBg,
          opacity: isDragging ? 0.4 : 1,
          transition:"background 0.1s",
          cursor:"pointer", boxSizing:"border-box",
          position:"relative",
        }}>
        <DragBtn active={activeDragId===item.id} onActivate={()=>onActivateDrag(item.id)}
          onMoveUp={onMoveUp} onMoveDown={onMoveDown} canUp={canUp} canDown={canDown}
          onMouseDownDrag={onMouseDownDrag}/>
        <div style={{display:"flex",justifyContent:"center"}}>
          {isSector
            ? <div style={{width:8,height:8,borderRadius:2,background:c}}/>
            : item.type==="milestone"
            ? <div style={{width:8,height:8,background:c,transform:"rotate(45deg)",borderRadius:1}}/>
            : <div style={{width:7,height:7,borderRadius:"50%",background:c}}/>
          }
        </div>
        <div style={{overflow:"hidden",paddingLeft: isSector ? 4 : 14,display:"flex",alignItems:"center",gap:5,
          borderLeft: isSector ? "none" : "2px solid #1e2535", marginLeft: isSector ? 0 : 4}}>
          <span style={{fontSize:12,fontWeight:isSector?700:500,
            color:isSector?T.text:T.textSub,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>
            {item.name}
          </span>
          {item.link && <LinkIcon size={10} color="#6366f1"/>}
          {isSector && (
            <div style={{display:"flex",alignItems:"center",gap:2,marginLeft:"auto",flexShrink:0}}>
              <div
                onClick={e=>{e.stopPropagation();onToggle(item.id);}}
                style={{color:"#6366f1",padding:"4px 6px",cursor:"pointer",borderRadius:5,
                  display:"flex",alignItems:"center",
                  background:"transparent",transition:"background 0.12s"}}
                onMouseEnter={e=>e.currentTarget.style.background="rgba(99,102,241,0.15)"}
                onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                {item.collapsed ? <ChevronRight size={12}/> : <ChevronDown size={12}/>}
              </div>
            </div>
          )}
        </div>
        <div style={{textAlign:"center",fontSize:11,color: isSector?T.textFaint:T.textMuted,fontVariantNumeric:"tabular-nums"}}>
          {!isSector && item.start ? fmtD(item.start) : ""}
        </div>
        <div style={{textAlign:"center",fontSize:11,color: isSector?T.textFaint:T.textMuted,fontVariantNumeric:"tabular-nums"}}>
          {!isSector && item.type!=="milestone" && item.end ? fmtD(item.end) : ""}
        </div>
        <div style={{textAlign:"center",fontSize:11,fontWeight:600,
          color: item.type==="task" ? (item.progress===100?"#10b981":item.progress>0?"#818cf8":T.textDim) : "transparent"}}>
          {item.type==="task" ? item.progress+"%" : ""}
        </div>
      </div>
      {isSelected && (
        <div style={{height:DETAIL_HEIGHT,background:T.bgSelected,borderBottom:`1px solid ${T.border}`,
          borderLeft:`3px solid ${c}`,padding:"10px 12px 10px 14px",boxSizing:"border-box",overflow:"hidden"}}>
          {isSector && (
            <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
              <div style={{width:12,height:12,borderRadius:3,background:item.color,flexShrink:0}}/>
              <span style={{fontSize:12,color:T.textSub}}>
                {(() => { const kids=(allItems||[]).filter(i=>i.parentId===item.id); return '하위 작업 '+kids.length+'개'; })()}
              </span>
            </div>
          )}
          {item.type==="task" && (
            <div style={{marginBottom:8}}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}>
                <span style={{fontSize:10,color:T.textFaint,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.05em"}}>진행률</span>
                <span style={{fontSize:11,fontWeight:700,color:c}}>{item.progress}%</span>
              </div>
              <div style={{height:4,background:T.border,borderRadius:3,overflow:"hidden"}}>
                <div style={{height:"100%",width:item.progress+"%",background:c,borderRadius:3}}/>
              </div>
            </div>
          )}
          <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
            {isSector && (
              <button onClick={e=>{e.stopPropagation();onToggle(item.id);}}
                style={{display:"flex",alignItems:"center",gap:4,padding:"4px 9px",background:T.bgRowHover,border:"1px solid #1e2535",borderRadius:6,cursor:"pointer",fontSize:11,color:T.textMuted,fontWeight:500}}>
                {item.collapsed?<ChevronRight size={10}/>:<ChevronDown size={10}/>}{item.collapsed?"펼치기":"접기"}
              </button>
            )}
            {isSector && (
              <button onClick={e=>{e.stopPropagation();onEdit({type:"new-child",parentId:item.id});}}
                style={{display:"flex",alignItems:"center",gap:4,padding:"4px 9px",background:"rgba(99,102,241,0.1)",border:"1px solid #6366f133",borderRadius:6,cursor:"pointer",fontSize:11,color:"#818cf8",fontWeight:500}}>
                <PlusIcon size={10} color="#818cf8"/>작업 추가
              </button>
            )}
            <button onClick={e=>{e.stopPropagation();onEdit(item);}}
              style={{display:"flex",alignItems:"center",gap:4,padding:"4px 9px",background:themeKey==="dark"?"rgba(255,255,255,0.06)":"rgba(0,0,0,0.05)",border:"1px solid #1e2535",borderRadius:6,cursor:"pointer",fontSize:11,color:T.textSub,fontWeight:500}}>
              <EditIcon size={10} color={T.textSub}/>편집
            </button>
            {!isSector && (
              item.link
                ? <button onClick={e=>{e.stopPropagation();window.open(item.link,"_blank");}}
                    style={{display:"flex",alignItems:"center",gap:4,padding:"4px 9px",background:"rgba(99,102,241,0.12)",border:"1px solid #6366f155",borderRadius:6,cursor:"pointer",fontSize:11,color:"#818cf8",fontWeight:500}}>
                    <LinkIcon size={10} color="#818cf8"/>링크 열기
                  </button>
                : <button onClick={e=>{e.stopPropagation();onEdit(item);}}
                    style={{display:"flex",alignItems:"center",gap:4,padding:"4px 9px",background:"rgba(255,255,255,0.03)",border:"1px dashed #2a3245",borderRadius:6,cursor:"pointer",fontSize:11,color:T.textDim,fontWeight:500}}>
                    <LinkIcon size={10} color={T.textDim}/>링크 추가
                  </button>
            )}
            <button onClick={e=>{e.stopPropagation();onDelete(item.id);onSelect(null);}}
              style={{display:"flex",alignItems:"center",gap:4,padding:"4px 9px",background:"rgba(248,113,113,0.08)",border:"1px solid #f8717122",borderRadius:6,cursor:"pointer",fontSize:11,color:"#f87171",fontWeight:500}}>
              <TrashIcon size={10} color="#f87171"/>삭제
            </button>
          </div>
        </div>
      )}
      {dropIndicator==="above" && (
        <div style={{position:"absolute",top:0,left:0,right:0,height:2,background:"#6366f1",zIndex:20,borderRadius:1,
          boxShadow:"0 0 6px #6366f1"}}/>
      )}
      {dropIndicator==="below" && (
        <div style={{position:"absolute",bottom:0,left:0,right:0,height:2,background:"#6366f1",zIndex:20,borderRadius:1,
          boxShadow:"0 0 6px #6366f1"}}/>
      )}
    </>
  );
}

// --- Kanban View --------------------------------------------------------------
const KANBAN_COLS = [
  { key:"todo", label:"할 일", color:T.textFaint },
  { key:"inprogress", label:"진행 중", color:"#6366f1" },
  { key:"done", label:"완료", color:"#10b981" },
];

function getKStatus(item) {
  const p = item.progress || 0;
  if (p === 100) return "done";
  if (p === 0) return "todo";
  return "inprogress";
}

function KanbanView({ items, onEdit, onDelete, openAdd, openAddSector, themeKey="dark" }) {
  const T = THEMES[themeKey];
  const [dragCard, setDragCard] = useState(null);
  const [dragOver, setDragOver] = useState(null);
  const [collapsedSectors, setCollapsedSectors] = useState({});

  const sectors = items.filter(i => i.type === "sector")
    .slice().sort((a,b) => (a.order||0)-(b.order||0));
  const getCards = (sectorId, statusKey) =>
    items.filter(i => i.parentId === sectorId && i.type === "task" && getKStatus(i) === statusKey)
      .slice().sort((a,b) => (a.order||0)-(b.order||0));
  const moveCard = (cardId, newStatus) => {
    const pm = { todo:0, inprogress:50, done:100 };
    onEdit({ _kanbanMove:true, id:cardId, progress:pm[newStatus] });
  };
  // NOTE: getKStatus already correctly maps 100->done before review check
  const fmt = d => d ? (d.getMonth()+1) + "/" + d.getDate() : "";
  const toggleSector = id => setCollapsedSectors(prev => ({...prev, [id]: !prev[id]}));

  const COL_W = 220;
  const SECTOR_LABEL_W = 160;

  return (
    <div style={{flex:1, display:"flex", flexDirection:"column", background:T.bgDeep, overflow:"hidden"}}>

      {/* -- Scrollable container: header + body scroll together horizontally -- */}
      <div style={{flex:1, overflowX:"auto", overflowY:"hidden", display:"flex", flexDirection:"column"}}>

      {/* -- Fixed header: status columns -- */}
      <div style={{display:"flex", flexShrink:0, borderBottom:"2px solid #1e2535", background:T.bgHeader, position:"sticky", top:0, zIndex:10,
        minWidth: SECTOR_LABEL_W + KANBAN_COLS.length * COL_W}}>
        {/* top-left corner: sector label column */}
        <div style={{
          width: SECTOR_LABEL_W, flexShrink:0,
          display:"flex", alignItems:"center", justifyContent:"space-between",
          padding:"0 14px", borderRight:"1px solid #1e2535", gap:6
        }}>
          <span style={{fontSize:10,fontWeight:700,color:T.textDim,textTransform:"uppercase",letterSpacing:"0.08em"}}>섹터</span>
          <button onClick={openAddSector} style={{
            display:"flex",alignItems:"center",gap:4,padding:"4px 8px",
            background:themeKey==="dark"?"rgba(255,255,255,0.04)":"rgba(0,0,0,0.04)",border:"1px solid #1e2535",
            borderRadius:6,cursor:"pointer",color:T.textFaint,fontSize:11
          }}>
            <PlusIcon size={10} color={T.textFaint}/>섹터
          </button>
        </div>
        {/* status column headers */}
        {KANBAN_COLS.map(col => (
          <div key={col.key} style={{
            width: COL_W, flexShrink:0,
            display:"flex", alignItems:"center", gap:8,
            padding:"12px 16px",
            borderRight:"1px solid #1a2030",
            borderTop: `3px solid ${col.color}`,
          }}>
            <div style={{width:8,height:8,borderRadius:"50%",background:col.color,flexShrink:0}}/>
            <span style={{fontSize:12,fontWeight:700,color:col.color}}>{col.label}</span>
            <span style={{
              fontSize:10,color:T.textDim,background:T.border,
              borderRadius:10,padding:"1px 8px",marginLeft:"auto"
            }}>
              {sectors.reduce((acc,s)=>acc+getCards(s.id,col.key).length,0)}
            </span>
          </div>
        ))}
      </div>

      {/* -- Scrollable body -- */}
      <div style={{flex:1, overflowY:"auto", overflowX:"visible", minWidth: SECTOR_LABEL_W + KANBAN_COLS.length * COL_W}}>
        {sectors.map(sector => {
          const isCollapsed = collapsedSectors[sector.id];
          const totalTasks = items.filter(i=>i.parentId===sector.id&&i.type==="task").length;
          const doneTasks  = items.filter(i=>i.parentId===sector.id&&i.type==="task"&&i.progress===100).length;
          const pct = totalTasks ? Math.round(doneTasks/totalTasks*100) : 0;

          return (
            <div key={sector.id} style={{
              display:"flex", flexDirection:"column",
              borderBottom:`2px solid ${hexAlpha(sector.color,0.25)}`,
              minWidth: SECTOR_LABEL_W + KANBAN_COLS.length * COL_W,
            }}>
              {/* Sector row header */}
              <div style={{
                display:"flex", alignItems:"stretch",
                background: hexAlpha(sector.color, 0.07),
                borderBottom: isCollapsed ? "none" : `1px solid ${hexAlpha(sector.color,0.2)}`,
                flexShrink:0,
              }}>
                {/* Sector label cell */}
                <div style={{
                  width: SECTOR_LABEL_W, flexShrink:0,
                  display:"flex", flexDirection:"column", justifyContent:"center",
                  padding:"14px 14px",
                  borderRight:`2px solid ${hexAlpha(sector.color,0.3)}`,
                  borderLeft:`4px solid ${sector.color}`,
                  gap:8,
                }}>
                  <div style={{display:"flex",alignItems:"center",gap:8}}>
                    <div style={{width:10,height:10,borderRadius:3,background:sector.color,flexShrink:0}}/>
                    <span style={{fontSize:13,fontWeight:700,color:T.text,flex:1,minWidth:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                      {sector.name}
                    </span>
                    <button
                      onClick={()=>toggleSector(sector.id)}
                      style={{width:20,height:20,display:"flex",alignItems:"center",justifyContent:"center",
                        background:"transparent",border:"none",cursor:"pointer",flexShrink:0,borderRadius:4,
                        color:T.textFaint}}
                      title={isCollapsed?"펼치기":"접기"}>
                      {isCollapsed ? <ChevronRight size={12}/> : <ChevronDown size={12}/>}
                    </button>
                  </div>
                  {/* progress bar */}
                  <div>
                    <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
                      <span style={{fontSize:9,color:T.textFaint,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.05em"}}>{totalTasks}개 작업</span>
                      <span style={{fontSize:9,fontWeight:700,color:pct===100?"#10b981":sector.color}}>{pct}%</span>
                    </div>
                    <div style={{height:3,background:themeKey==="dark"?"rgba(255,255,255,0.07)":"rgba(0,0,0,0.06)",borderRadius:2,overflow:"hidden"}}>
                      <div style={{height:"100%",width:pct+"%",background:sector.color,borderRadius:2,transition:"width 0.4s"}}/>
                    </div>
                  </div>
                  {/* add task button */}
                  <button onClick={()=>openAdd("task",sector.id)} style={{
                    display:"flex",alignItems:"center",gap:5,padding:"5px 8px",
                    background:themeKey==="dark"?"rgba(255,255,255,0.04)":"rgba(0,0,0,0.04)",border:`1px solid ${hexAlpha(sector.color,0.3)}`,
                    borderRadius:6,cursor:"pointer",fontSize:11,color:sector.color,fontWeight:500,
                    width:"100%",justifyContent:"center"
                  }}>
                    <PlusIcon size={10} color={sector.color}/>작업 추가
                  </button>
                </div>

                {/* Status cells (collapsed: show mini counts) */}
                {KANBAN_COLS.map(col => {
                  const cards = getCards(sector.id, col.key);
                  return (
                    <div key={col.key} style={{
                      width:COL_W, flexShrink:0,
                      display:"flex", alignItems:"center", justifyContent:"center",
                      borderRight:"1px solid #1a2030",
                      padding:"8px",
                    }}>
                      {isCollapsed ? (
                        /* collapsed: badge only */
                        <div style={{
                          display:"flex",alignItems:"center",gap:6,
                          padding:"4px 12px",borderRadius:20,
                          background: cards.length>0 ? hexAlpha(col.color,0.12) : themeKey==="dark"?"rgba(255,255,255,0.03)":"rgba(99,102,241,0.04)",
                          border:`1px solid ${cards.length>0?hexAlpha(col.color,0.3):T.border}`,
                        }}>
                          <div style={{width:6,height:6,borderRadius:"50%",background:cards.length>0?col.color:T.textDim}}/>
                          <span style={{fontSize:12,fontWeight:700,color:cards.length>0?col.color:T.textDim}}>{cards.length}</span>
                        </div>
                      ) : (
                        /* expanded: just a thin accent line at top */
                        <div style={{width:"100%",height:2,background:cards.length>0?hexAlpha(col.color,0.2):"transparent",borderRadius:1}}/>
                      )}
                    </div>
                  );
                })}
              </div>

              {/* Cards row (hidden when collapsed) */}
              {!isCollapsed && (
                <div style={{display:"flex", alignItems:"flex-start"}}>
                  {/* empty sector label space */}
                  <div style={{width:SECTOR_LABEL_W,flexShrink:0,borderRight:`2px solid ${hexAlpha(sector.color,0.15)}`,alignSelf:"stretch"}}/>
                  {/* card columns */}
                  {KANBAN_COLS.map(col => {
                    const cards = getCards(sector.id, col.key);
                    const isOver = dragOver && dragOver.sectorId===sector.id && dragOver.colKey===col.key;
                    return (
                      <div key={col.key}
                        onDragOver={e => { e.preventDefault(); setDragOver({sectorId:sector.id, colKey:col.key}); }}
                        onDragLeave={() => setDragOver(null)}
                        onDrop={e => { e.preventDefault(); if(dragCard) moveCard(dragCard, col.key); setDragOver(null); setDragCard(null); }}
                        style={{
                          width:COL_W, flexShrink:0, minHeight:100,
                          padding:"10px 8px",
                          display:"flex", flexDirection:"column", gap:7,
                          borderRight:"1px solid #1a2030",
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
                                border:"1px solid #1e2535",
                                borderTop:`3px solid ${cc}`,
                                borderRadius:8, padding:"10px 11px",
                                cursor:"pointer",
                                opacity: dragCard===card.id ? 0.35 : 1,
                                transition:"all 0.15s, opacity 0.1s",
                                boxShadow: dragCard===card.id?"none":"0 2px 10px rgba(0,0,0,0.25)",
                              }}
                              onMouseEnter={e=>{ e.currentTarget.style.borderColor=T.border; e.currentTarget.style.transform="translateY(-1px)"; }}
                              onMouseLeave={e=>{ e.currentTarget.style.borderColor=T.border; e.currentTarget.style.transform=""; }}>
                              {/* type badge + link */}
                              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:7}}>
                                {card.type==="milestone"
                                  ? <span style={{fontSize:9,fontWeight:700,color:"#f59e0b",background:"rgba(245,158,11,0.12)",padding:"2px 7px",borderRadius:20,border:"1px solid rgba(245,158,11,0.25)"}}>&#9670; 마일스톤</span>
                                  : <span style={{fontSize:9,fontWeight:700,color:cc,background:hexAlpha(cc,0.12),padding:"2px 7px",borderRadius:20,border:`1px solid ${hexAlpha(cc,0.25)}`}}>작업</span>
                                }
                                {card.link && <LinkIcon size={10} color="#6366f1"/>}
                              </div>
                              {/* name */}
                              <div style={{fontSize:12,fontWeight:600,color:T.text,lineHeight:1.45,marginBottom:7}}>
                                {card.name}
                              </div>
                              {/* meta */}
                              <div style={{display:"flex",flexWrap:"wrap",gap:5,marginBottom:card.type==="task"?8:0}}>
                                
                                {card.start && (
                                  <span style={{fontSize:10,color:T.textFaint}}>
                                    {fmt(card.start)}{card.end && card.type!=="milestone" ? "-"+fmt(card.end) : ""}
                                  </span>
                                )}
                              </div>
                              {/* progress */}
                              {card.type==="task" && (
                                <div>
                                  <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}>
                                    <span style={{fontSize:9,color:T.textDim}}>진행률</span>
                                    <span style={{fontSize:9,fontWeight:700,color:cc}}>{card.progress}%</span>
                                  </div>
                                  <div style={{height:3,background:T.border,borderRadius:2,overflow:"hidden"}}>
                                    <div style={{height:"100%",width:card.progress+"%",background:`linear-gradient(90deg,${cc},${hexAlpha(cc,0.7)})`,borderRadius:2}}/>
                                  </div>
                                </div>
                              )}
                            </div>
                          );
                        })}
                        {cards.length === 0 && (
                          <div style={{
                            height:56,display:"flex",alignItems:"center",justifyContent:"center",
                            border:"1px dashed #1e2535",borderRadius:8,color:T.border,fontSize:11,
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
      </div>{/* end horizontal scroll wrapper */}
    </div>
  );
}

// --- Main App -----------------------------------------------------------------
// serialize/deserialize dates
function serializeItems(items) {
  const safeISO = d => {
    try { return d && !isNaN(d.getTime()) ? d.toISOString() : null; }
    catch(e) { return null; }
  };
  return items.map(item => ({
    ...item,
    start: safeISO(item.start),
    end:   safeISO(item.end),
  }));
}
function deserializeItems(items) {
  return items.map(item => ({
    ...item,
    start: item.start ? new Date(item.start) : null,
    end:   item.end   ? new Date(item.end)   : null,
  }));
}

export default function GanttApp() {
  const [items, setItems] = useState(initialItems);
  const [activeTab, setActiveTab] = useState('gantt');
  const [channelId, setChannelId] = useState("default");
  const [dbStatus, setDbStatus] = useState("idle"); // idle | saving | saved | error
  const [projectName, setProjectName] = useState("프로젝트 간트");
  const [updateAvailable, setUpdateAvailable] = useState(false);
  const [themeKey, setThemeKey] = useState("dark");
  const T = THEMES[themeKey]; // reactive for GanttApp
  const [editingName, setEditingName] = useState(false);
  const saveTimerRef = useRef(null);
  const importInputRef = useRef(null);
  const [zoomIdx, setZoomIdx] = useState(4);
  const [viewStart, setViewStart] = useState(addDays(today,-7));
  const [showModal, setShowModal] = useState(false);
  const [editItem, setEditItem] = useState(null);
  const [form, setForm] = useState(null);
  const [dragState, setDragState] = useState(null);
  const [activeDragId, setActiveDragId] = useState(null);
  const [rowDragState, setRowDragState] = useState(null);
  // rowDragState: { id, startY, currentIdx, dropIdx, dropPos }
  const [rowDragPos, setRowDragPos] = useState({x:0, y:0}); // ghost position
  const [selectedId, setSelectedId] = useState(null);
  const [hoveredId, setHoveredId] = useState(null);
  const [linkingFrom, setLinkingFrom] = useState(null);
  const [showMilestones, setShowMilestones] = useState(true);
  const [showLeftPanel, setShowLeftPanel] = useState(true);
  const [menuOpen, setMenuOpen] = useState(false);

  const [viewport, setViewport] = useState({w: window.innerWidth, h: window.innerHeight});
  useEffect(()=>{
    const onResize = ()=>setViewport({w:window.innerWidth,h:window.innerHeight});
    window.addEventListener("resize",onResize);
    return ()=>window.removeEventListener("resize",onResize);
  },[]);

  // -- 버전 체크: 5분마다 version.json 확인 ------------------------------
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

  // Apply theme CSS variables to :root
  useEffect(()=>{
    const t = THEMES[themeKey];
    setGlobalTheme(themeKey);
    const root = document.documentElement;
    Object.entries(t).forEach(([k,v]) => root.style.setProperty("--t-"+k, v));
    document.body.style.background = t.bg;
    document.body.style.color = t.text;
  }, [themeKey]);

  // -- Supabase: 채널 ID 가져오고 데이터 로드 + Teams 멤버 가져오기 ------------
  useEffect(()=>{
    getChannelId().then(async id => {
      setChannelId(id);
      try {
        const data = await sbGet(id);
        if (data) {
          if (Array.isArray(data)) {
            // 구버전 포맷 (배열)
            setItems(deserializeItems(data));
          } else if (data.items) {
            // 신버전 포맷 (객체)
            setItems(deserializeItems(data.items));
            if (data.projectName) setProjectName(data.projectName);
            if (data.uiState) {
              const ui = data.uiState;
              if (ui.activeTab) setActiveTab(ui.activeTab);
              if (ui.zoomIdx !== undefined) setZoomIdx(ui.zoomIdx);
              if (ui.showMilestones !== undefined) setShowMilestones(ui.showMilestones);
              if (ui.themeKey) setThemeKey(ui.themeKey);
              // collapsed sectors are stored in items themselves - already handled via items
            }
          }
        }
      } catch(e) {
        console.error("Supabase load error:", e);
      }
    });

    // Teams 멤버 목록 가져오기
    if (window.microsoftTeams) {
      window.microsoftTeams.app.initialize().then(() => {
        window.microsoftTeams.app.getContext().then(ctx => {
          // Teams Graph API로 멤버 가져오기 (가능한 경우)
          // 불가능할 경우 빈 배열 유지 -> 수동 입력 폴백
        }).catch(() => {});
      }).catch(() => {});
    }
  }, []);

  // -- Supabase: items/projectName 변경 시 자동 저장 (1초 디바운스) -----------
  useEffect(()=>{
    if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
    setDbStatus("saving");
    saveTimerRef.current = setTimeout(async () => {
      try {
        await sbSave(channelId, {
            items: serializeItems(items),
            projectName,
            uiState: { activeTab, zoomIdx, showMilestones, themeKey }
          });
        setDbStatus("saved");
        setTimeout(()=>setDbStatus("idle"), 2000);
      } catch(e) {
        console.error("Save error:", e);
        setDbStatus("error");
        // Retry once after 3 seconds
        setTimeout(async () => {
          try {
            await sbSave(channelId, {
              items: serializeItems(items),
              projectName,
              uiState: { activeTab, zoomIdx, showMilestones, themeKey }
            });
            setDbStatus("saved");
            setTimeout(()=>setDbStatus("idle"), 2000);
          } catch(e2) {
            console.error("Retry failed:", e2);
          }
        }, 3000);
      }
    }, 1000);
  }, [items, projectName, channelId]);

  const isMobile = viewport.w < 640;
  const isTablet = viewport.w >= 640 && viewport.w < 1024;
  const isLandscape = viewport.w > viewport.h;
  const leftPanelWidth = isMobile ? Math.min(viewport.w*0.96, 420) : isTablet ? 340 : 420;
  const leftPanelOverlay = isMobile;

  const chartRef = useRef(null);
  const headerRef = useRef(null);
  const leftPanelListRef = useRef(null);
  const isSyncingScroll = useRef(false);
  const zoom = ZOOM_LEVELS[zoomIdx];

  // -- "전체" fit view: compute range from all items ----------------------
  const fitRange = useMemo(() => {
    const allDates = items.flatMap(i => [i.start, i.end]).filter(Boolean);
    if (!allDates.length) return { start: addDays(today, -7), days: 42, dw: 38 };
    const minD = new Date(Math.min(...allDates.map(d => d.getTime())));
    const maxD = new Date(Math.max(...allDates.map(d => d.getTime())));
    minD.setDate(minD.getDate() - 7);
    maxD.setDate(maxD.getDate() + 7);
    const totalD = Math.max(dateDiff(minD, maxD) + 1, 14);
    const chartW = Math.max((chartRef.current?.clientWidth || (viewport.w - 450)), 400);
    const dw = Math.max(3, Math.floor(chartW / totalD));
    return { start: minD, days: totalD, dw };
  }, [items, zoomIdx, viewport]);

  const isFit = zoom.key === "fit";
  const DAY_WIDTH = isFit ? fitRange.dw : zoom.dayWidth;
  const totalDays = isFit ? fitRange.days : zoom.totalDays;
  const effectiveViewStart = isFit ? fitRange.start : viewStart;
  const viewEnd = addDays(effectiveViewStart, totalDays - 1);

  const flatRows = useMemo(()=>{
    const rows=[];
    items.filter(i=>i.type==="sector").slice().sort((a,b)=>(a.order??0)-(b.order??0))
      .forEach(sector=>{
        rows.push({item:sector,depth:0});
        if (!sector.collapsed)
          items.filter(i=>i.parentId===sector.id&&(showMilestones||i.type!=="milestone"))
            .slice().sort((a,b)=>(a.order??0)-(b.order??0))
            .forEach(child=>rows.push({item:child,depth:1}));
      });
    return rows;
  },[items,showMilestones]);

  // Dynamic row heights: selected row expands by DETAIL_HEIGHT
  const getRowHeight = useCallback((rowIdx) => {
    const item = flatRows[rowIdx]?.item;
    return item && item.id === selectedId ? ROW_HEIGHT + DETAIL_HEIGHT : ROW_HEIGHT;
  }, [flatRows, selectedId]);

  const getRowTop = useCallback((rowIdx) => {
    let top = 0;
    for (let i = 0; i < rowIdx; i++) top += getRowHeight(i);
    return top;
  }, [getRowHeight]);

  const totalChartHeight = useMemo(() => {
    let h = 0;
    for (let i = 0; i < flatRows.length; i++) h += getRowHeight(i);
    return h;
  }, [flatRows, getRowHeight]);

  const days = useMemo(()=>{
    const arr=[]; const cur=new Date(effectiveViewStart);
    while(cur<=viewEnd){arr.push(new Date(cur));cur.setDate(cur.getDate()+1);}
    return arr;
  },[effectiveViewStart,viewEnd]);

  const headerGroups = useMemo(()=>{
    const g=[];
    days.forEach((d,i)=>{
      const label = zoom.key==="year" ? `${d.getFullYear()}`
        : zoom.key==="quarter" ? `${d.getFullYear()} Q${Math.floor(d.getMonth()/3)+1}`
        : `${d.getFullYear()}년 ${d.getMonth()+1}월`;
      if (!g.length||g[g.length-1].label!==label) g.push({label,count:1});
      else g[g.length-1].count++;
    });
    return g;
  },[days,zoom.key]);

  const handleChartScroll = e => {
    if (headerRef.current) headerRef.current.scrollLeft = e.target.scrollLeft;
    if (!isSyncingScroll.current && leftPanelListRef.current) {
      isSyncingScroll.current = true;
      leftPanelListRef.current.scrollTop = e.target.scrollTop;
      isSyncingScroll.current = false;
    }
  };
  const handleLeftPanelScroll = e => {
    if (!isSyncingScroll.current && chartRef.current) {
      isSyncingScroll.current = true;
      chartRef.current.scrollTop = e.target.scrollTop;
      isSyncingScroll.current = false;
    }
  };

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

  // -- Row drag-and-drop via mouse ---------------------------------------------
  useEffect(()=>{
    if (!rowDragState) return;
    const onMove = e => {
      // update ghost position
      setRowDragPos({x: e.clientX, y: e.clientY});
      const listEl = leftPanelListRef.current;
      if (!listEl) return;
      const rect = listEl.getBoundingClientRect();
      const relY = e.clientY - rect.top + listEl.scrollTop;
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
      setRowDragState(prev => ({...prev, dropIdx, dropPos}));
    };
    const onUp = e => {
      if (rowDragState.dropIdx !== undefined) {
        const fromIdx = flatRows.findIndex(r => r.item.id === rowDragState.id);
        let toIdx = rowDragState.dropIdx;
        if (rowDragState.dropPos === "below") toIdx = toIdx + 1;
        if (toIdx > fromIdx) toIdx -= 1;
        if (fromIdx !== toIdx) {
          setItems(prev => {
            const dragged = flatRows[fromIdx].item;
            if (dragged.type === "sector") {
              const target = flatRows[rowDragState.dropIdx]?.item;
              if (!target || target.type !== "sector") return prev;
              const sectors = prev.filter(i=>i.type==="sector").slice().sort((a,b)=>(a.order??0)-(b.order??0));
              const si = sectors.findIndex(s=>s.id===dragged.id);
              const ti = sectors.findIndex(s=>s.id===target.id);
              if (si<0||ti<0) return prev;
              const r=[...sectors]; r.splice(si,1); r.splice(rowDragState.dropPos==="above"?ti:ti+1>si?ti:ti,0,dragged);
              const map=Object.fromEntries(r.map((s,i)=>[s.id,i]));
              return prev.map(i=>i.type==="sector"?{...i,order:map[i.id]??i.order}:i);
            }
            return applyFlatMove(prev, dragged.id, fromIdx, toIdx < 0 ? 0 : toIdx);
          });
        }
      }
      setRowDragState(null);
    };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
    return () => { window.removeEventListener("mousemove", onMove); window.removeEventListener("mouseup", onUp); };
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
    if (fromIdx < 0 || toIdx < 0 || fromIdx >= flat.length || toIdx >= flat.length) return allItems;
    const dragged = flat[fromIdx];
    const newFlat = [...flat];
    newFlat.splice(fromIdx, 1);
    newFlat.splice(toIdx, 0, dragged);
    const getParent = (idx) => {
      for (let i = idx; i >= 0; i--) {
        if (newFlat[i].type === "sector") return newFlat[i].id;
      }
      return allItems.find(i=>i.type==="sector")?.id ?? dragged.parentId;
    };
    const sectorOrder = {};
    const childOrder = {};
    let sIdx = 0;
    newFlat.forEach((item, i) => {
      if (item.type === "sector") {
        sectorOrder[item.id] = sIdx++;
        childOrder[item.id] = {};
      } else {
        const pid = getParent(i);
        if (!childOrder[pid]) childOrder[pid] = {};
        const cnt = Object.keys(childOrder[pid]).length;
        childOrder[pid][item.id] = cnt;
      }
    });
    return allItems.map(item => {
      if (item.type === "sector") {
        return { ...item, order: sectorOrder[item.id] ?? item.order };
      }
      const newParent = Object.keys(childOrder).find(pid => item.id in childOrder[pid]);
      if (newParent) {
        return { ...item, parentId: Number(newParent), order: childOrder[newParent][item.id] };
      }
      return item;
    });
  };

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
      let toIdx = fromIdx + dir;
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
    const onKey=e=>{if(e.key==="Escape"){setLinkingFrom(null);setActiveDragId(null);setSelectedId(null);}};
    window.addEventListener("keydown",onKey);
    return ()=>window.removeEventListener("keydown",onKey);
  },[]);

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
    setEditItem(null);
    setForm({type,parentId:def,name:"",start:toISO(today),end:toISO(addDays(today,7)),progress:0,link:"",deps:[]});
    setShowModal(true); setMenuOpen(false);
  };

  const openAddSector=()=>{
    setEditItem(null);
    const used=items.filter(i=>i.type==="sector").map(i=>i.color);
    const next=SECTOR_COLORS.find(c=>!used.includes(c))||SECTOR_COLORS[0];
    setForm({type:"sector",name:"",color:next,collapsed:false});
    setShowModal(true); setMenuOpen(false);
  };

  const openEdit=item=>{
    setEditItem(item);
    setForm({...item,start:item.start?toISO(item.start):"",end:item.end?toISO(item.end):""});
    setShowModal(true);
  };

  const saveForm=()=>{
    if (!form.name.trim()) return;
    let parsed;
    if (form.type==="sector") { parsed={...form}; }
    else { const{color:_c,...rest}=form; parsed={...rest,start:parseDate(form.start),end:parseDate(form.end||form.start)}; }
    if (editItem) setItems(prev=>prev.map(i=>i.id===editItem.id?{...parsed,id:i.id}:i));
    else setItems(prev=>{const o=assignOrder({...parsed,id:uid()},prev);return [...prev,o];});
    setShowModal(false);
  };

  const deleteItem=id=>setItems(prev=>{
    const isSec=prev.find(i=>i.id===id&&i.type==="sector");
    if(isSec) return prev.filter(i=>i.id!==id&&i.parentId!==id);
    return prev.filter(i=>i.id!==id).map(i=>({...i,deps:(i.deps||[]).filter(d=>d!==id)}));
  });

  const toggleCollapse=id=>setItems(prev=>prev.map(i=>i.id===id?{...i,collapsed:!i.collapsed}:i));

  const handleLinkClick=(e,itemId)=>{
    e.stopPropagation();
    if (!linkingFrom){setLinkingFrom(itemId);return;}
    if (linkingFrom===itemId){setLinkingFrom(null);return;}
    setItems(prev=>prev.map(i=>{ if(i.id!==itemId) return i; const deps=i.deps||[]; return deps.includes(linkingFrom)?i:{...i,deps:[...deps,linkingFrom]}; }));
    setLinkingFrom(null);
  };

  // -- CSV Import
  const importFromCSV = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        let text = e.target.result;
        if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
        const lines = text.split("\n").map(l => l.trimEnd()).filter(l => l.length > 0);
        if (lines.length < 2) return;

        const DQUOTE = String.fromCharCode(34);
        const COMMA  = String.fromCharCode(44);

        const parseRow = (line) => {
          const cols = [];
          let cur = "";
          let inQ = false;
          for (let i = 0; i < line.length; i++) {
            const ch = line[i];
            if (ch === DQUOTE) {
              if (inQ && line[i+1] === DQUOTE) { cur += DQUOTE; i++; }
              else { inQ = !inQ; }
            } else if (ch === COMMA && !inQ) {
              cols.push(cur); cur = "";
            } else {
              cur += ch;
            }
          }
          cols.push(cur);
          return cols;
        };

        const rows = lines.slice(1).map(parseRow);
        const sectorMap = {};
        let maxId = Math.max(...items.map(i => i.id), 100);
        const newItems = [];

        rows.forEach(row => {
          const sectorName = (row[0] || "").trim();
          const type       = (row[1] || "").trim();
          const name       = (row[2] || "").trim();
          const startStr   = (row[3] || "").trim();
          const endStr     = (row[4] || "").trim();
          const progress   = (row[5] || "").trim();
          const link       = (row[6] || "").trim();
          const depsStr    = (row[7] || "").trim();
          if (!sectorName || !name) return;

          if (!sectorMap[sectorName]) {
            const sId = ++maxId;
            sectorMap[sectorName] = sId;
            const usedColors = newItems.filter(i => i.type === "sector").map(i => i.color);
            const color = SECTOR_COLORS.find(c => !usedColors.includes(c)) || SECTOR_COLORS[0];
            newItems.push({ id: sId, type: "sector", name: sectorName, color, collapsed: false, order: Object.keys(sectorMap).length - 1 });
          }

          const itemId   = ++maxId;
          const itemType = type === "마일스톤" ? "milestone" : "task";
          newItems.push({
            id: itemId, type: itemType,
            parentId: sectorMap[sectorName],
            name,
            start: startStr ? parseDate(startStr) : null,
            end:   endStr   ? parseDate(endStr)   : (startStr ? parseDate(startStr) : null),
            progress: itemType === "task" ? (parseInt(progress) || 0) : 0,
            link, deps: [],
            order: newItems.filter(i => i.parentId === sectorMap[sectorName]).length
          });
        });

        newItems.forEach(item => {
          const row = rows.find(r => (r[2] || "").trim() === item.name);
          if (!row || !row[7]) return;
          const depNames = row[7].split(";").map(s => s.trim()).filter(Boolean);
          item.deps = depNames.map(dn => newItems.find(i => i.name === dn)).filter(Boolean).map(i => i.id);
        });

        if (newItems.length > 0) {
          setItems(sortedByDate(newItems));
          alert("가져오기 완료! " + newItems.filter(i => i.type !== "sector").length + "개 작업을 불러왔습니다.");
        }
      } catch(err) {
        alert("파일을 읽는 중 오류가 발생했습니다.");
        console.error(err);
      }
    };
    reader.readAsText(file, "UTF-8");
  };

  const exportToExcel=()=>{
    const sectors=items.filter(i=>i.type==="sector");
    const rows=[["섹터","유형","작업명","시작일","종료일","진행률(%)","링크","의존작업"]];
    sectors.forEach(sector=>{
      items.filter(i=>i.parentId===sector.id).forEach(item=>{
        const depNames=(item.deps||[]).map(d=>items.find(x=>x.id===d)?.name||"").filter(Boolean).join(";");
        rows.push([sector.name,item.type==="milestone"?"마일스톤":"작업",item.name,
          item.start?toISO(item.start):"",item.end?toISO(item.end):"",
          item.type==="task"?item.progress:"",item.link||"",depNames]);
      });
    });
    const csv=rows.map(r=>r.map(cell=>`"${String(cell).replace(/"/g,'""')}"`).join(",")).join("\n");
    const encoded = encodeURIComponent("\uFEFF" + csv);
    const a = document.createElement("a");
    a.href = "data:text/csv;charset=utf-8," + encoded;
    a.download = `간트차트_${toISO(today)}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  const todayX=dateDiff(effectiveViewStart,today)*DAY_WIDTH+DAY_WIDTH/2;
  const totalW=days.length*DAY_WIDTH;

  const SB = {
    nav: {padding:"6px 10px",background:"transparent",color:T.textSub,border:`1px solid ${T.border}`,borderRadius:7,cursor:"pointer",fontSize:12,display:"flex",alignItems:"center",gap:5,fontWeight:500},
    today: {padding:"6px 12px",background:"#6366f1",color:"#fff",border:"none",borderRadius:7,cursor:"pointer",fontSize:12,fontWeight:600},
    add: {padding:"7px 13px",background:"linear-gradient(135deg,#6366f1,#8b5cf6)",color:"#fff",border:"none",borderRadius:7,cursor:"pointer",fontSize:12,fontWeight:600,display:"flex",alignItems:"center",gap:5},
    sec: {padding:"7px 11px",background:"transparent",color:T.textSub,border:`1px solid ${T.border}`,borderRadius:7,cursor:"pointer",fontSize:12,display:"flex",alignItems:"center",gap:5,fontWeight:500},
    inp: {background:T.bgInput,border:`1px solid ${T.border}`,borderRadius:7,padding:"9px 11px",color:T.text,fontSize:13,outline:"none",width:"100%",boxSizing:"border-box",colorScheme:themeKey},
  };

  function ToolbarContent({vertical=false}) {
    return (
      <div style={{display:"flex",flexDirection:vertical?"column":"row",gap:vertical?6:5,alignItems:vertical?"stretch":"center",padding:vertical?"8px 10px":0}}>
        <button onClick={()=>setShowMilestones(v=>!v)}
          style={{...SB.sec,color:showMilestones?"#818cf8":T.textMuted,border:"1px solid "+(showMilestones?"#6366f133":T.border)}}>
          <MilestoneIcon size={12} color={showMilestones?"#818cf8":T.textFaint}/>
          {"마일스톤 "+(showMilestones?"ON":"OFF")}
        </button>
        {linkingFrom&&<button onClick={()=>setLinkingFrom(null)} style={{...SB.sec,color:"#fbbf24",border:"1px solid #92400e"}}>
          {"x 연결 취소"}
        </button>}
        {!vertical&&<div style={{width:1,height:22,background:T.border}}/>}
        <button onClick={()=>setViewStart(addDays(viewStart,-Math.max(1,Math.floor(totalDays/4))))} disabled={isFit} style={{...SB.nav,opacity:isFit?0.4:1}}>
          <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round"><path d="M15 18l-6-6 6-6"/></svg>
          이전
        </button>
        <button onClick={()=>setViewStart(addDays(today,-7))} disabled={isFit} style={{...SB.today,opacity:isFit?0.6:1}}>오늘</button>
        <button onClick={()=>setViewStart(addDays(viewStart,Math.max(1,Math.floor(totalDays/4))))} disabled={isFit} style={{...SB.nav,opacity:isFit?0.4:1}}>
          다음
          <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round"><path d="M9 18l6-6-6-6"/></svg>
        </button>
        <div style={{display:"flex",border:"1px solid #1e2535",borderRadius:7,overflow:"hidden"}}>
          {ZOOM_LEVELS.map((z,i)=>(
            <button key={z.key} onClick={()=>setZoomIdx(i)}
              style={{padding:"6px 10px",border:"none",cursor:"pointer",fontSize:12,fontWeight:600,
                background:i===zoomIdx?(z.key==="fit"?"linear-gradient(135deg,#10b981,#059669)":"#6366f1"):"transparent",
                color:i===zoomIdx?"#fff":T.textMuted,
                borderLeft:z.key==="fit"?("1px solid "+T.border):"none"}}>
              {z.label}
            </button>
          ))}
        </div>
        {!vertical&&<div style={{width:1,height:22,background:T.border}}/>}
        <button onClick={openAddSector} style={{...SB.sec,justifyContent:vertical?"flex-start":"center"}}>
          <PlusIcon size={12}/>{" 섹터"}
        </button>
        <button onClick={()=>openAdd("task")} style={{...SB.add,justifyContent:vertical?"flex-start":"center"}}>
          <PlusIcon size={12} color="#fff"/>{" 작업 추가"}
        </button>
        {!vertical&&<div style={{width:1,height:22,background:T.border}}/>}
        {/* Theme toggle */}
        <button onClick={()=>setThemeKey(k=>k==="dark"?"light":"dark")}
          title={themeKey==="dark"?"라이트 모드":"다크 모드"}
          style={{...SB.sec, padding:"6px 8px"}}>
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
        {/* 숨김 파일 입력 */}
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
  }

  return (
    <div style={{fontFamily:"Pretendard, Noto Sans KR, sans-serif",background:T.bg,height:"100vh",display:"flex",flexDirection:"column",overflow:"hidden",color:T.text,position:"relative",
      cursor:rowDragState?"grabbing":"auto", userSelect:rowDragState?"none":"auto"}}>
      {/* -- Update banner -- */}
      {updateAvailable && (
        <div style={{
          background:"linear-gradient(90deg,#6366f1,#8b5cf6)",
          color:"#fff", padding:"10px 20px",
          display:"flex", alignItems:"center", justifyContent:"space-between",
          flexShrink:0, zIndex:200,
          boxShadow:"0 2px 12px rgba(99,102,241,0.4)"
        }}>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2.5" strokeLinecap="round">
              <polyline points="23 4 23 10 17 10"/><polyline points="1 20 1 14 7 14"/>
              <path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/>
            </svg>
            <span style={{fontSize:13,fontWeight:600}}>새 버전이 있어요! 새로고침하면 최신 버전으로 업데이트됩니다. (데이터는 유지됩니다)</span>
          </div>
          <div style={{display:"flex",gap:8,flexShrink:0}}>
            <button
              onClick={()=>window.location.reload()}
              style={{padding:"6px 16px",background:"#fff",color:"#6366f1",border:"none",
                borderRadius:7,cursor:"pointer",fontSize:13,fontWeight:700,
                boxShadow:"0 2px 8px rgba(0,0,0,0.15)"}}>
              지금 새로고침
            </button>
            <button
              onClick={()=>setUpdateAvailable(false)}
              style={{padding:"6px 10px",background:"rgba(255,255,255,0.2)",color:"#fff",
                border:"1px solid rgba(255,255,255,0.3)",borderRadius:7,cursor:"pointer",fontSize:12}}>
              나중에
            </button>
          </div>
        </div>
      )}

      {/* -- TOP BAR -- */}
      <div style={{padding:isMobile?"10px 12px":"12px 18px",background:themeKey==="dark"?"linear-gradient(180deg,#161b27,#0f1117)":"linear-gradient(180deg,#ffffff,#f0f2f7)",borderBottom:`1px solid ${T.border}`,display:"flex",alignItems:"center",gap:12,flexShrink:0,flexWrap:"wrap",boxShadow:themeKey==="light"?"0 1px 8px rgba(99,102,241,0.08)":"none"}}>
        <div style={{display:"flex",alignItems:"center",gap:10,minWidth:0}}>
          {isMobile&&(
            <button onClick={()=>setShowLeftPanel(v=>!v)} style={{...SB.nav,padding:"6px 8px",flexShrink:0}}>
              <MenuIcon size={16}/>
            </button>
          )}
          <div style={{width:30,height:30,background:"linear-gradient(135deg,#6366f1,#8b5cf6)",borderRadius:8,boxShadow:themeKey==="light"?"0 2px 8px rgba(99,102,241,0.3)":"none",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2.5" strokeLinecap="round"><rect x="3" y="4" width="18" height="4" rx="1"/><rect x="3" y="10" width="12" height="4" rx="1"/><rect x="3" y="16" width="8" height="4" rx="1"/></svg>
          </div>
          <div style={{minWidth:0}}>
            {editingName ? (
              <input
                autoFocus
                value={projectName}
                onChange={e=>setProjectName(e.target.value)}
                onBlur={()=>setEditingName(false)}
                onKeyDown={e=>{ if(e.key==="Enter"||e.key==="Escape") setEditingName(false); }}
                style={{
                  fontSize:isMobile?14:16, fontWeight:700, color:T.text,
                  letterSpacing:"-0.02em", background:"transparent",
                  border:"none", borderBottom:"2px solid #6366f1",
                  outline:"none", width: Math.max(projectName.length*11, 120)+"px",
                  padding:"0 2px",
                }}
              />
            ) : (
              <div
                onClick={()=>setEditingName(true)}
                title="클릭하여 이름 수정"
                style={{
                  fontSize:isMobile?14:16, fontWeight:700, color:T.text,
                  letterSpacing:"-0.02em", cursor:"text",
                  borderBottom:"2px solid transparent",
                  transition:"border-color 0.15s",
                  padding:"0 2px",
                }}
                onMouseEnter={e=>e.currentTarget.style.borderBottomColor="#6366f144"}
                onMouseLeave={e=>e.currentTarget.style.borderBottomColor="transparent"}
              >
                {projectName}
              </div>
            )}
            {!isMobile&&<div style={{fontSize:11,color:T.textMuted,display:"flex",alignItems:"center",gap:6}}>
              {items.filter(i=>i.type!=="sector").length}개 작업
              {dbStatus==="saving"&&<span style={{color:"#f59e0b",fontSize:10}}>o 저장 중...</span>}
              {dbStatus==="saved"&&<span style={{color:"#10b981",fontSize:10}}>o 저장됨</span>}
              {dbStatus==="error"&&<span style={{color:"#f87171",fontSize:10}}>o 저장 실패</span>}
            </div>}
          </div>
        </div>
        <div style={{display:"flex",background:T.bg,border:"1px solid #1e2535",borderRadius:8,overflow:"hidden",flexShrink:0}}>
          {[
            { key:"gantt", label:"간트" },
            { key:"kanban", label:"칸반" },
          ].map(tab => (
            <button key={tab.key} onClick={()=>setActiveTab(tab.key)}
              style={{padding: isMobile?"6px 10px":"6px 14px", border:"none", cursor:"pointer",
                fontSize:12, fontWeight:600, transition:"all 0.15s", display:"flex", alignItems:"center", gap:5,
                background: activeTab===tab.key ? "#6366f1" : "transparent",
                color: activeTab===tab.key ? "#fff" : T.textMuted}}>
              
              {!isMobile && <span>{tab.label}</span>}
            </button>
          ))}
        </div>
        {isMobile ? (
          <div style={{display:"flex",gap:6,alignItems:"center",marginLeft:"auto"}}>
            <button onClick={()=>openAdd("task")} style={{...SB.add,padding:"6px 10px"}}>
              <PlusIcon size={12} color="#fff"/>{" 추가"}
            </button>
            <button onClick={()=>setMenuOpen(v=>!v)} style={{...SB.nav,padding:"6px 8px"}}>
              {menuOpen?<CloseIcon size={16}/>:<MenuIcon size={16}/>}
            </button>
          </div>
        ) : (
          <div style={{marginLeft:"auto"}}>
            <ToolbarContent/>
          </div>
        )}
      </div>

      {isMobile&&menuOpen&&(
        <div style={{position:"absolute",top:54,right:0,left:0,zIndex:90,background:T.bgCard,borderBottom:"1px solid #1e2535",boxShadow:"0 8px 32px rgba(0,0,0,0.5)"}}>
          <ToolbarContent vertical={true}/>
        </div>
      )}

      {/* -- KANBAN VIEW -- */}
      {activeTab==="kanban" && (
        <KanbanView
          items={items}
          themeKey={themeKey}
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
        />
      )}

      {/* -- GANTT MAIN -- */}
      {activeTab==="gantt" ? (() => {
        // Single scroll container: left panel columns + chart columns scroll together vertically
        // Only the chart scrolls horizontally
        const HEADER_H = 56; // two header rows (28+28)
        const LEFT_COL_H = 68; // left panel title(40) + column header(28)
        return (
          <div style={{display:"flex",flex:1,overflow:"hidden",position:"relative"}}>
            {leftPanelOverlay&&showLeftPanel&&(
              <div onClick={()=>setShowLeftPanel(false)}
                style={{position:"absolute",inset:0,background:"rgba(0,0,0,0.5)",zIndex:40}}/>
            )}

            {/* -- LEFT PANEL -- fixed-width column */}
            {(showLeftPanel||!leftPanelOverlay)&&(
              <div style={{
                width:leftPanelWidth, flexShrink:0,
                borderRight:"1px solid #1e2535",
                display:"flex", flexDirection:"column",
                background:T.bg,
                ...(leftPanelOverlay?{position:"absolute",left:0,top:0,bottom:0,zIndex:50,boxShadow:"4px 0 24px rgba(0,0,0,0.5)"}:{})
              }}>
                {/* Left header (matches chart header height exactly) */}
                <div style={{flexShrink:0,borderBottom:"1px solid #1e2535",background:T.bgHeader}}>
                  <div style={{height:40,padding:"0 6px 0 4px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                    <span style={{fontSize:11,fontWeight:600,color:T.textFaint,textTransform:"uppercase",letterSpacing:"0.08em",paddingLeft:8}}>작업 목록</span>
                    {leftPanelOverlay&&<button onClick={()=>setShowLeftPanel(false)} style={{...SB.nav,padding:"4px 6px"}}><CloseIcon size={14}/></button>}
                  </div>
                  <div style={{display:"grid",gridTemplateColumns:GRID_COLS,alignItems:"center",height:28,padding:"0 6px 0 4px"}}>
                    <div/><div/>
                    <div style={{fontSize:10,fontWeight:600,color:T.textFaint,paddingLeft:4}}>작업명</div>
                    <div style={{fontSize:10,fontWeight:600,color:T.textFaint,textAlign:"center"}}>시작</div>
                    <div style={{fontSize:10,fontWeight:600,color:T.textFaint,textAlign:"center"}}>종료</div>
                    <div style={{fontSize:10,fontWeight:600,color:T.textFaint,textAlign:"center"}}>진행</div>
                  </div>
                </div>
                {/* Left rows - overflow hidden, driven by the single scroll container below */}
                <div ref={leftPanelListRef} style={{flex:1,overflow:"hidden"}}>
                  <div style={{height:totalChartHeight}}>
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
                    ))}
                  </div>
                </div>
              </div>
            )}

            {/* -- CHART PANEL -- */}
            <div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden",minWidth:0}}>
              {/* Sticky date header - scrolls only horizontally */}
              <div ref={headerRef} style={{overflowX:"hidden",flexShrink:0,borderBottom:"1px solid #1e2535"}}>
                <div style={{width:totalW,display:"flex",flexDirection:"column"}}>
                  <div style={{display:"flex",height:40,background:T.bgHeader,borderBottom:"1px solid #1a2030",boxSizing:"border-box"}}>
                    {headerGroups.map((g,i)=>(
                      <div key={i} style={{width:g.count*DAY_WIDTH,flexShrink:0,display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:600,color:T.textFaint,borderRight:"1px solid #1a2030"}}>
                        {g.label}
                      </div>
                    ))}
                  </div>
                  <div style={{display:"flex",height:28,background:T.bgHeader}}>
                    {days.map((d,i)=>{
                      const isT=d.getTime()===today.getTime(), isW=d.getDay()===0||d.getDay()===6;
                      return (
                        <div key={i} style={{width:DAY_WIDTH,flexShrink:0,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",
                          background:isT?"rgba(99,102,241,0.15)":isW?"rgba(255,255,255,0.01)":"transparent"}}>
                          {zoom.key!=="year"&&<span style={{fontSize:DAY_WIDTH<14?7:9,fontWeight:isT?700:400,color:isT?"#818cf8":T.textDim}}>
                            {zoom.key==="day"||zoom.key==="month"?(d.getMonth()+1)+"/"+d.getDate():zoom.key==="quarter"?d.getDate():""}
                          </span>}
                          {zoom.key==="day"&&<span style={{fontSize:8,color:isT?"#6366f1":T.textDim}}>
                            {["일","월","화","수","목","금","토"][d.getDay()]}
                          </span>}
                        </div>
                      );
                    })}
                  </div>
                </div>
              </div>

              {/* Single scroll container - both vertical and horizontal */}
              <div ref={chartRef}
                onScroll={e=>{
                  // sync horizontal header
                  if (headerRef.current) headerRef.current.scrollLeft = e.target.scrollLeft;
                  // sync left panel vertical (no feedback loop guard needed - left is overflow:hidden)
                  if (leftPanelListRef.current) leftPanelListRef.current.scrollTop = e.target.scrollTop;
                }}
                onWheel={e=>{
                  if (Math.abs(e.deltaY) > Math.abs(e.deltaX)) {
                    e.preventDefault();
                    setZoomIdx(prev => e.deltaY < 0 ? Math.min(prev+1, ZOOM_LEVELS.length-1) : Math.max(prev-1, 0));
                  }
                }}
                style={{flex:1,overflowX:"scroll",overflowY:"scroll",position:"relative",WebkitOverflowScrolling:"touch"}}>
                <div style={{width:totalW,minHeight:totalChartHeight,position:"relative"}}>
                  {todayX>=0&&todayX<=totalW&&(
                    <div style={{position:"absolute",left:todayX,top:0,bottom:0,width:2,background:"linear-gradient(180deg,#6366f1,#6366f100)",zIndex:5,pointerEvents:"none"}}>
                      <div style={{position:"absolute",top:0,left:-14,background:"#6366f1",color:"#fff",fontSize:9,fontWeight:700,padding:"2px 5px",borderRadius:3,whiteSpace:"nowrap"}}>오늘</div>
                    </div>
                  )}
                  {days.map((d,i)=>(d.getDay()===0||d.getDay()===6)&&(
                    <div key={i} style={{position:"absolute",left:i*DAY_WIDTH,top:0,width:DAY_WIDTH,height:"100%",background:themeKey==="dark"?"rgba(255,255,255,0.012)":"rgba(99,102,241,0.03)",pointerEvents:"none"}}/>
                  ))}
                  {flatRows.map((_,i)=>{
                    const rh=getRowHeight(i), rt=getRowTop(i);
                    return <div key={i} style={{position:"absolute",left:0,top:rt,width:"100%",height:rh,borderBottom:`1px solid ${T.borderLight}`,pointerEvents:"none"}}/>;
                  })}
                  <DepArrows flatRows={flatRows} viewStart={effectiveViewStart} DAY_WIDTH={DAY_WIDTH} getRowTop={getRowTop}/>
                  {flatRows.map(({item},rowIdx)=>{
                    const top=getRowTop(rowIdx);
                    if (item.type==="sector") {
                      const kids=items.filter(c=>c.parentId===item.id&&c.start);
                      if (!kids.length) return null;
                      const ms=new Date(Math.min(...kids.map(c=>c.start.getTime())));
                      const me=new Date(Math.max(...kids.map(c=>c.end.getTime())));
                      const l=dateDiff(effectiveViewStart,ms)*DAY_WIDTH;
                      const w=(dateDiff(ms,me)+1)*DAY_WIDTH-2;
                      return <div key={item.id} style={{position:"absolute",top:top+ROW_HEIGHT/2-5,left:l,width:w,height:10,background:hexAlpha(item.color,0.18),borderRadius:3,border:`1px solid ${hexAlpha(item.color,0.35)}`,pointerEvents:"none"}}/>;
                    }
                    if (item.type==="milestone") {
                      if (!item.start) return null;
                      const mc=effectiveColor(item,items);
                      const cx=(dateDiff(effectiveViewStart,item.start)+0.5)*DAY_WIDTH;
                      const cy=top+ROW_HEIGHT/2, sz=12;
                      return (
                        <div key={item.id} title={item.name+(item.link?" ":"")}
                          onMouseEnter={()=>setHoveredId(item.id)} onMouseLeave={()=>setHoveredId(null)}
                          onClick={e=>{e.stopPropagation();if(linkingFrom!==null){handleLinkClick(e,item.id);}else{setSelectedId(prev=>prev===item.id?null:item.id);}}}
                          style={{position:"absolute",left:cx-sz,top:cy-sz,width:sz*2,height:sz*2,cursor:"pointer",zIndex:6}}>
                          <div style={{position:"absolute",inset:3,background:mc,transform:"rotate(45deg)",borderRadius:2}}/>
                          {DAY_WIDTH>=20&&<div style={{position:"absolute",top:sz*2+2,left:"50%",transform:"translateX(-50%)",fontSize:9,color:T.textSub,whiteSpace:"nowrap",fontWeight:600}}>{item.name}</div>}
                        </div>
                      );
                    }
                    if (!item.start) return null;
                    const barColor=effectiveColor(item,items);
                    const left=dateDiff(effectiveViewStart,item.start)*DAY_WIDTH;
                    const width=Math.max((dateDiff(item.start,item.end)+1)*DAY_WIDTH-3,14);
                    return (
                      <div key={item.id}
                        onMouseEnter={()=>setHoveredId(item.id)} onMouseLeave={()=>setHoveredId(null)}
                        style={{position:"absolute",top:top+10,left,width,height:32,userSelect:"none",zIndex:6}}>
                        <div onMouseDown={e=>{e.preventDefault();e.stopPropagation();setDragState({id:item.id,type:"resize-left",startX:e.clientX,origStart:item.start,origEnd:item.end});}}
                          style={{position:"absolute",left:0,top:0,width:7,height:"100%",cursor:"ew-resize",zIndex:2}}/>
                        <div
                          onMouseDown={e=>{if(linkingFrom!==null)return;e.preventDefault();e.stopPropagation();setDragState({id:item.id,type:"move",startX:e.clientX,origStart:item.start,origEnd:item.end});}}
                          onClick={e=>{e.stopPropagation();if(linkingFrom!==null){handleLinkClick(e,item.id);}else{setSelectedId(prev=>prev===item.id?null:item.id);}}}
                          style={{position:"absolute",inset:0,background:`linear-gradient(135deg,${barColor},${hexAlpha(barColor,0.8)})`,
                            borderRadius:6,
                            boxShadow:hoveredId===item.id?`0 4px 20px ${hexAlpha(barColor,0.27)}`:"none",
                            cursor:linkingFrom!==null?"crosshair":"pointer",transition:"box-shadow 0.2s"}}>
                          <div style={{position:"absolute",left:0,top:0,bottom:0,width:item.progress+"%",background:"rgba(255,255,255,0.15)",borderRadius:6,transition:"width 0.3s"}}/>
                          {width>48&&<div style={{position:"absolute",left:8,right:8,top:"50%",transform:"translateY(-50%)",fontSize:11,fontWeight:600,color:"#fff",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",textShadow:"0 1px 2px rgba(0,0,0,0.4)"}}>
                            {item.link&&" "}{item.name}
                          </div>}
                        </div>
                        <div onMouseDown={e=>{e.preventDefault();e.stopPropagation();setDragState({id:item.id,type:"resize-right",startX:e.clientX,origStart:item.start,origEnd:item.end});}}
                          style={{position:"absolute",right:0,top:0,width:7,height:"100%",cursor:"ew-resize",zIndex:2}}/>
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>
          </div>
        );
      })() : null}

      {/* -- MODAL -- */}
      {showModal&&form&&(
        <div style={{position:"fixed",inset:0,background:themeKey==="dark"?"rgba(0,0,0,0.78)":"rgba(30,35,60,0.55)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:100,padding:16}}>
          <div style={{background:T.bgCard,border:"1px solid #1e2535",borderRadius:16,padding:"22px 24px",width:"100%",maxWidth:480,maxHeight:"90vh",overflowY:"auto",boxShadow:themeKey==="dark"?"0 24px 80px rgba(0,0,0,0.6)":"0 24px 80px rgba(99,102,241,0.15)"}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:20}}>
              <h2 style={{margin:0,fontSize:16,fontWeight:700,color:T.text}}>
                {editItem?"항목 편집":form.type==="sector"?"섹터 추가":"작업 추가"}
              </h2>
              <button onClick={()=>setShowModal(false)} style={{...SB.nav,padding:"4px 6px"}}><CloseIcon size={14}/></button>
            </div>
            <div style={{display:"flex",flexDirection:"column",gap:13}}>
              {form.type!=="sector"&&!editItem&&(
                <div style={{display:"flex",gap:0,background:T.bg,borderRadius:10,padding:4}}>
                  {[{key:"task",label:" 시간막대",desc:"기간이 있는 작업"},{key:"milestone",label:"&#9670; 마일스톤",desc:"특정 날짜 이벤트"}].map(opt=>(
                    <button key={opt.key} onClick={()=>setForm(f=>({...f,type:opt.key,end:opt.key==="milestone"?f.start:f.end}))}
                      style={{flex:1,padding:"9px 6px",border:"none",borderRadius:8,cursor:"pointer",
                        background:form.type===opt.key?(opt.key==="milestone"?"linear-gradient(135deg,#f59e0b,#d97706)":"linear-gradient(135deg,#6366f1,#8b5cf6)"):"transparent",
                        color:form.type===opt.key?"#fff":T.textMuted,fontWeight:form.type===opt.key?600:400,transition:"all 0.15s",fontSize:12}}>
                      <div>{opt.label}</div>
                      <div style={{fontSize:10,fontWeight:400,marginTop:2,opacity:0.8}}>{opt.desc}</div>
                    </button>
                  ))}
                </div>
              )}
              <F label="이름"><input value={form.name} onChange={e=>setForm({...form,name:e.target.value})} style={SB.inp} placeholder="작업 이름을 입력하세요"/></F>
              {form.type!=="sector"&&(<>
                <div style={{display:"grid",gridTemplateColumns:form.type==="milestone"?"1fr":"1fr 1fr",gap:10}}>
                  <F label={form.type==="milestone"?"날짜":"시작일"}>
                    <input type="date" value={form.start} onChange={e=>setForm({...form,start:e.target.value})} style={SB.inp}/>
                  </F>
                  {form.type!=="milestone"&&<F label="종료일"><input type="date" value={form.end} onChange={e=>setForm({...form,end:e.target.value})} style={SB.inp}/></F>}
                </div>
                <F label="섹터">
                  <select value={form.parentId||""} onChange={e=>setForm({...form,parentId:+e.target.value})} style={SB.inp}>
                    {items.filter(i=>i.type==="sector").map(s=><option key={s.id} value={s.id}>{s.name}</option>)}
                  </select>
                </F>
                {form.type==="task"&&(<>
                  <F label={'진행률: ' + (form.progress||0) + '%'}>
                    <input type="range" min={0} max={100} value={form.progress||0} onChange={e=>setForm({...form,progress:+e.target.value})} style={{width:"100%",accentColor:"#6366f1"}}/>
                  </F>
                </>)}
                <F label="링크 URL">
                  <input value={form.link||""} onChange={e=>setForm({...form,link:e.target.value})} style={SB.inp} placeholder="https://..."/>
                </F>
                <F label="의존 작업 (복수 선택 가능)">
                  {(()=>{
                    const candidates = items.filter(i => i.type !== "sector" && i.id !== (editItem?.id ?? -1));
                    const sectors = items.filter(i => i.type === "sector");
                    const deps = form.deps || [];
                    const toggle = id => {
                      setForm(f => ({
                        ...f,
                        deps: f.deps?.includes(id) ? f.deps.filter(d => d !== id) : [...(f.deps||[]), id]
                      }));
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
                                  <span style={{width:6,height:6,borderRadius:"50%",background:dc,flexShrink:0,display:"inline-block"}}/>
                                  {dep.name}
                                  <span onClick={()=>toggle(depId)} style={{cursor:"pointer",color:"#f87171",fontWeight:700,fontSize:14,lineHeight:1}}>x</span>
                                </span>
                              );
                            })}
                          </div>
                        )}
                        <div style={{background:T.bg,border:"1px solid #1e2535",borderRadius:8,
                          maxHeight:180,overflowY:"auto",WebkitOverflowScrolling:"touch"}}>
                          {candidates.length === 0 && (
                            <div style={{padding:"10px 12px",fontSize:12,color:T.textFaint}}>선택 가능한 작업 없음</div>
                          )}
                          {sectors.map(sector => {
                            const kids = candidates.filter(i => i.parentId === sector.id);
                            if (!kids.length) return null;
                            return (
                              <div key={sector.id}>
                                <div style={{padding:"5px 10px",fontSize:10,fontWeight:700,
                                  color:sector.color,textTransform:"uppercase",letterSpacing:"0.06em",
                                  background:"rgba(255,255,255,0.02)",borderBottom:"1px solid #1a2030",
                                  display:"flex",alignItems:"center",gap:6}}>
                                  <span style={{width:6,height:6,borderRadius:1,background:sector.color,display:"inline-block"}}/>
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
                                        borderBottom:"1px solid #1a2030",
                                        transition:"background 0.1s"}}
                                      onMouseEnter={e=>{ if(!checked) e.currentTarget.style.background="rgba(255,255,255,0.03)"; }}
                                      onMouseLeave={e=>{ e.currentTarget.style.background = checked?"rgba(99,102,241,0.1)":"transparent"; }}>
                                      <div style={{width:16,height:16,borderRadius:4,flexShrink:0,
                                        border: checked ? "none" : "1.5px solid #334155",
                                        background: checked ? "#6366f1" : "transparent",
                                        display:"flex",alignItems:"center",justifyContent:"center",
                                        transition:"all 0.12s"}}>
                                        {checked && <svg width="9" height="7" viewBox="0 0 9 7"><path d="M1 3.5L3.5 6L8 1" stroke="#fff" strokeWidth="1.5" fill="none" strokeLinecap="round" strokeLinejoin="round"/></svg>}
                                      </div>
                                      {task.type==="milestone"
                                        ? <div style={{width:8,height:8,background:tc,transform:"rotate(45deg)",borderRadius:1,flexShrink:0}}/>
                                        : <div style={{width:7,height:7,borderRadius:"50%",background:tc,flexShrink:0}}/>
                                      }
                                      <div style={{flex:1,minWidth:0}}>
                                        <div style={{fontSize:12,color: checked?T.text:T.textSub,
                                          fontWeight: checked?600:400,
                                          whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>
                                          {task.name}
                                        </div>
                                        {task.start && (
                                          <div style={{fontSize:10,color:T.textFaint,marginTop:1}}>
                                            {task.type==="milestone" ? toISO(task.start) : toISO(task.start) + " - " + toISO(task.end)}
                                          </div>
                                        )}
                                      </div>
                                    </div>
                                  );
                                })}
                              </div>
                            );
                          })}
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
                <F label="섹터 색상">
                  <div style={{display:"flex",gap:8,flexWrap:"wrap",marginTop:4}}>
                    {SECTOR_COLORS.map(c=>(
                      <div key={c} onClick={()=>setForm({...form,color:c})}
                        style={{width:26,height:26,borderRadius:"50%",background:c,cursor:"pointer",
                          border:form.color===c?"3px solid #fff":"3px solid transparent",
                          boxShadow:form.color===c?`0 0 0 2px ${c}`:"none",transition:"all 0.15s"}}/>
                    ))}
                  </div>
                </F>
              ):(
                <F label="색상 미리보기 (섹터 기반 자동 적용)">
                  {(()=>{
                    const sector=items.find(i=>i.id===form.parentId&&i.type==="sector");
                    if (!sector) return <span style={{fontSize:12,color:T.textFaint}}>섹터를 먼저 선택하세요</span>;
                    const siblings=items.filter(i=>i.parentId===sector.id&&i.type!=="sector");
                    const isNew=!editItem;
                    const cnt=isNew?siblings.length+1:siblings.length;
                    const idx=isNew?siblings.length:siblings.findIndex(i=>i.id===editItem?.id);
                    const pc=childColor(sector.color,Math.max(idx,0),cnt);
                    return (<div style={{display:"flex",alignItems:"center",gap:10,marginTop:4}}>
                      <div style={{width:28,height:28,borderRadius:7,background:pc,border:"2px solid rgba(255,255,255,0.1)"}}/>
                      <div>
                        <div style={{fontSize:12,color:T.textSub}}>섹터 <span style={{color:sector.color,fontWeight:600}}>{sector.name}</span> 기반</div>
                        <div style={{fontSize:11,color:T.textFaint,marginTop:2}}>같은 섹터 안에서 밝기로 구분됩니다</div>
                      </div>
                    </div>);
                  })()}
                </F>
              )}
            </div>
            <div style={{display:"flex",gap:8,marginTop:20,justifyContent:"flex-end"}}>
              <button onClick={()=>setShowModal(false)} style={SB.nav}>취소</button>
              <button onClick={saveForm} style={SB.add}>저장</button>
            </div>
          </div>
        </div>
      )}

      {/* -- Row drag ghost element -- */}
      {rowDragState && (() => {
        const draggingItem = flatRows.find(r => r.item.id === rowDragState.id)?.item;
        if (!draggingItem) return null;
        const isSector = draggingItem.type === "sector";
        const c = isSector ? draggingItem.color : effectiveColor(draggingItem, items);
        return (
          <div style={{
            position:"fixed",
            left: rowDragPos.x - 20,
            top: rowDragPos.y - ROW_HEIGHT / 2,
            width: leftPanelWidth - 8,
            height: ROW_HEIGHT,
            pointerEvents:"none",
            zIndex:1000,
            background: isSector ? "#1a2540" : "#161f35",
            border:`2px solid ${c}`,
            borderRadius:8,
            boxShadow:`0 8px 32px rgba(0,0,0,0.5), 0 0 0 1px ${c}44`,
            display:"grid",
            gridTemplateColumns: GRID_COLS,
            alignItems:"center",
            padding:"0 6px 0 4px",
            opacity:0.92,
            backdropFilter:"blur(4px)",
          }}>
            {/* grip */}
            <div style={{display:"flex",alignItems:"center",justifyContent:"center"}}>
              <GripIcon active={true}/>
            </div>
            {/* dot */}
            <div style={{display:"flex",justifyContent:"center"}}>
              {isSector
                ? <div style={{width:8,height:8,borderRadius:2,background:c}}/>
                : draggingItem.type==="milestone"
                ? <div style={{width:8,height:8,background:c,transform:"rotate(45deg)",borderRadius:1}}/>
                : <div style={{width:7,height:7,borderRadius:"50%",background:c}}/>
              }
            </div>
            {/* name */}
            <div style={{paddingLeft: isSector?4:14, overflow:"hidden", display:"flex", alignItems:"center",
              borderLeft: isSector?"none":"2px solid #2a3555", marginLeft: isSector?0:4}}>
              <span style={{fontSize:12, fontWeight:isSector?700:500, color:T.text,
                whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis"}}>
                {draggingItem.name}
              </span>
            </div>
            {/* start */}
            <div style={{textAlign:"center",fontSize:11,color:T.textMuted}}>
              {draggingItem.start ? `${String(draggingItem.start.getFullYear()).slice(2)}/${String(draggingItem.start.getMonth()+1).padStart(2,'0')}/${String(draggingItem.start.getDate()).padStart(2,'0')}` : ""}
            </div>
            {/* end */}
            <div style={{textAlign:"center",fontSize:11,color:T.textMuted}}>
              {draggingItem.end && draggingItem.type!=="milestone" ? `${String(draggingItem.end.getFullYear()).slice(2)}/${String(draggingItem.end.getMonth()+1).padStart(2,'0')}/${String(draggingItem.end.getDate()).padStart(2,'0')}` : ""}
            </div>
            {/* progress */}
            <div style={{textAlign:"center",fontSize:11,fontWeight:600,color:draggingItem.progress===100?"#10b981":draggingItem.progress>0?"#818cf8":T.textDim}}>
              {draggingItem.type==="task" ? draggingItem.progress+"%" : ""}
            </div>
          </div>
        );
      })()}

      {/* Drag mode toast */}
      {activeDragId && (() => {
        const item = items.find(i=>i.id===activeDragId);
        return (
          <div style={{position:"fixed",bottom:isMobile?16:22,left:"50%",transform:"translateX(-50%)",
            background:"rgba(99,102,241,0.97)",color:"#fff",padding:"10px 18px",borderRadius:40,
            fontSize:isMobile?12:13,fontWeight:700,zIndex:200,boxShadow:"0 4px 24px rgba(0,0,0,0.4)",
            display:"flex",alignItems:"center",gap:10,whiteSpace:"nowrap",pointerEvents:"none"}}>
            <GripIcon active={true}/>
            "{item?.name}" 이동 중 - ^v 버튼 또는 다른 행을 탭
          </div>
        );
      })()}

      {/* Linking toast */}
      {linkingFrom&&(
        <div style={{position:"fixed",bottom:isMobile?16:22,left:"50%",transform:"translateX(-50%)",
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
