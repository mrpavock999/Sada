import { useState, useEffect, useMemo } from "react";
import {
  AreaChart, Area, XAxis, YAxis, Tooltip, ResponsiveContainer, CartesianGrid,
} from "recharts";

/* ─────────────── DB (Google Sheets — single source of truth) ───────────────
 * No local fallback. Every read/write hits the Sheet. Errors are surfaced.
 * POST uses Content-Type: text/plain to avoid CORS preflight (Apps Script
 * cannot respond to OPTIONS). The Apps Script reads e.postData.contents as JSON.
 *
 * AUTH: every request carries a shared-secret passcode. Apps Script rejects
 * mismatches with HTTP 200 + { error:"unauthorized" }. The passcode lives in
 * localStorage; the gate component below collects it on first load.
 */
const SHEETS_URL = "https://script.google.com/macros/s/AKfycbzaxRUVz7vRqDw6Mf0-xowdLhc9rEVWGSJTzrsoBvOSHPrE7h8uB0zzxWeUI1hq0_pQeQ/exec";
const AUTH_KEY = "ht_auth_v1";
const getAuth = () => { try { return localStorage.getItem(AUTH_KEY) || ""; } catch { return ""; } };
const setAuth = (v) => { try { v ? localStorage.setItem(AUTH_KEY, v) : localStorage.removeItem(AUTH_KEY); } catch {} };
class UnauthorizedError extends Error { constructor() { super("unauthorized"); this.code = 401; } }

async function dbFetchAll() {
  const url = SHEETS_URL + "?k=" + encodeURIComponent(getAuth());
  const res = await fetch(url, { method: "GET", redirect: "follow" });
  const text = await res.text();
  if (!res.ok) throw new Error(`HTTP ${res.status}: ${text.slice(0,200)}`);
  if (text.trim().startsWith("<")) {
    const m = text.match(/TypeError[^<]*|Error[^<]*/);
    throw new Error(`Apps Script error: ${m ? m[0] : "returned HTML, not JSON"}`);
  }
  let json;
  try { json = JSON.parse(text); }
  catch { throw new Error("Sheet returned non-JSON response"); }
  if (json && json.error === "unauthorized") throw new UnauthorizedError();
  return {
    entries: Array.isArray(json.entries) ? json.entries : [],
    schema:  Array.isArray(json.schema)  ? json.schema  : null,
    theme:   (json.theme && typeof json.theme === "object") ? json.theme : null,
  };
}

async function dbWrite(key, value) {
  // text/plain avoids CORS preflight; Apps Script still reads e.postData.contents
  const res = await fetch(SHEETS_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ key, value, k: getAuth() }),
    redirect: "follow",
  });
  const text = await res.text();
  if (!res.ok) throw new Error(`HTTP ${res.status}: ${text.slice(0,200)}`);
  if (text.trim().startsWith("<")) {
    const m = text.match(/TypeError[^<]*|Error[^<]*/);
    throw new Error(`Apps Script error: ${m ? m[0] : "returned HTML, not JSON"}`);
  }
  let json;
  try { json = JSON.parse(text); } catch { return { ok:true }; }
  if (json && json.error === "unauthorized") throw new UnauthorizedError();
  return json;
}

/* ─────────────── PALETTE ─────────────── */
const C = {
  bg:"#08080f", card:"#0f0f1a", card2:"#141422", card3:"#191930",
  border:"#1e1e30", border2:"#25253a", text:"#e2e8f0", muted:"#64748b", muted2:"#4a5568",
  purple:"#8b5cf6", violet:"#7c3aed", cyan:"#06b6d4", green:"#10b981", amber:"#f59e0b",
  red:"#f43f5e", pink:"#ec4899", white:"#fff",
};

/* ─────────────── SCHEMA (default + types) ─────────────── *
 * type: "toggle" | "number" | "stepper" | "time"
 * op (number/stepper/time): "gte" | "lte"
 * threshold: target value (for time, minutes-since-midnight)
 * holidayThreshold (number/stepper): override on holidays
 * skipHoliday: skip this criterion on holidays
 * inverted (toggle): met when value is FALSE (e.g. "no junk food")
 * weight: relative weight in score (default 1)
 * unit: free-text unit shown in input
 * step: numeric input step (default 1)
 * max (stepper): max stepper value
 */
const DEFAULT_SCHEMA = [
  { id:"wake",        icon:"🌅", label:"Wake Time",       type:"time",    op:"lte", threshold:390,  weight:1, enabled:true, targetLabel:"By" },
  { id:"sleepHours",  icon:"😴", label:"Sleep Hours",     type:"number",  op:"gte", threshold:7,    unit:"hrs", step:0.5, weight:1, enabled:true, targetLabel:"Goal" },
  { id:"meditation",  icon:"🧘", label:"Meditation",      type:"toggle",  weight:1, enabled:true },
  { id:"workout",     icon:"💪", label:"Workout",         type:"toggle",  weight:1, enabled:true },
  { id:"steps",       icon:"👟", label:"Steps",           type:"number",  op:"gte", threshold:10000, unit:"steps", weight:1, enabled:true, scoreMode:"graded", targetLabel:"Goal" },
  { id:"water",       icon:"💧", label:"Water",           type:"stepper", op:"gte", threshold:8,    max:16, unit:"cups", weight:1, enabled:true, targetLabel:"Goal" },
  { id:"calories",    icon:"🔥", label:"Calories",        type:"number",  op:"lte", threshold:2100, unit:"kcal", weight:1, enabled:true, targetLabel:"Max",
    message:"Stay under daily calorie cap",
    rules:[
      { id:"r_cal_workout", name:"Workout day boost", enabled:true, whenLogic:"all",
        when:[{ field:"workout", op:"truthy" }],
        set:{ threshold:2300, message:"Lift day — fuel up to 2300 kcal", targetLabel:"Max" } },
    ] },
  { id:"protein",     icon:"🍗", label:"Protein",         type:"number",  op:"gte", threshold:150,  unit:"g", weight:1, enabled:true, targetLabel:"Target", scoreMode:"graded",
    message:"Hit your protein target" },
  { id:"fats",        icon:"🥑", label:"Fats",            type:"number",  op:"lte", threshold:70,   unit:"g", weight:1, enabled:true, targetLabel:"Max",
    message:"Keep fats under cap" },
  { id:"junkFood",    icon:"🚫", label:"No Junk Food",    type:"toggle",  inverted:true, weight:1, enabled:true },
  { id:"workBlocks",  icon:"📋", label:"Work Blocks",     type:"stepper", op:"gte", threshold:3,    max:8, weight:1, skipHoliday:true, enabled:true, targetLabel:"Goal" },
  { id:"reading",     icon:"📖", label:"Reading",         type:"toggle",  weight:1, enabled:true },
  { id:"cleaning",    icon:"🧹", label:"Cleaning",        type:"toggle",  weight:1, enabled:true },
  { id:"movieMinutes",icon:"🎬", label:"Screen Limit",    type:"number",  op:"lte", threshold:90,   holidayThreshold:400, unit:"min", weight:1, enabled:true, targetLabel:"Limit" },
];

/* ─────────────── VALIDATION RULES ───────────────
 * Each field has optional `rules: [Rule]`. A Rule looks at OTHER fields in the
 * same entry and, when its conditions match, overrides this field's config
 * (threshold / weight / message / targetLabel / op / scoreMode / skip / holidayThreshold).
 * Rule = {
 *   id, name, enabled, whenLogic: "all" | "any",
 *   when: [{ field, op: "is"|"not"|"truthy"|"falsy"|"gte"|"lte"|"gt"|"lt", value }],
 *   set:  { threshold?, holidayThreshold?, op?, weight?, scoreMode?, message?, targetLabel?, skip? }
 * }
 */
function evalWhen(cond, entry) {
  if (!cond || !cond.field) return false;
  const v = entry?.[cond.field];
  switch (cond.op) {
    case "truthy": return !!v;
    case "falsy":  return !v;
    case "is":     return v === cond.value || String(v ?? "") === String(cond.value ?? "");
    case "not":    return !(v === cond.value || String(v ?? "") === String(cond.value ?? ""));
    case "gte":    return Number(v) >= Number(cond.value);
    case "lte":    return Number(v) <= Number(cond.value);
    case "gt":     return Number(v) >  Number(cond.value);
    case "lt":     return Number(v) <  Number(cond.value);
    default:       return false;
  }
}

function resolveField(field, entry) {
  const rules = Array.isArray(field.rules) ? field.rules : [];
  let cfg = { ...field };
  const activeRules = [];
  for (const r of rules) {
    if (r.enabled === false) continue;
    const conds = Array.isArray(r.when) ? r.when : [];
    if (conds.length === 0) continue;
    const ok = (r.whenLogic || "all") === "all"
      ? conds.every(c => evalWhen(c, entry))
      : conds.some(c => evalWhen(c, entry));
    if (!ok) continue;
    // strip undefined/blank overrides so we don't wipe field defaults
    const set = {};
    for (const [k, v] of Object.entries(r.set || {})) {
      if (v === undefined || v === null || v === "") continue;
      set[k] = v;
    }
    cfg = { ...cfg, ...set };
    activeRules.push(r);
  }
  return { cfg, activeRules };
}

function targetText(cfg) {
  if (cfg.type === "toggle") {
    return cfg.inverted ? "Avoid today" : "Complete today";
  }
  const lbl = cfg.targetLabel || (cfg.op === "lte" ? "Limit" : "Goal");
  if (cfg.type === "time") return `${lbl} ${cfg.op === "lte" ? "by" : "after"} ${minToTime(cfg.threshold || 0)}`;
  if (cfg.type === "number" || cfg.type === "stepper") {
    const u = cfg.unit ? ` ${cfg.unit}` : "";
    return `${lbl}: ${cfg.op === "lte" ? "≤" : "≥"} ${cfg.threshold}${u}`;
  }
  return "";
}

/* ─────────────── SCORE ENGINE ─────────────── */
// Robust boolean coercion. Sheets / xlsx round-trips often turn JS booleans
// into strings like "true"/"FALSE"/"0". Plain `!!v` then reads "false" as
// truthy, which made completed toggles appear met after a refresh. Always go
// through this helper before scoring or persisting a boolean.
const coerceBool = (v) => {
  if (v === undefined || v === null || v === "") return false;
  if (typeof v === "boolean") return v;
  if (typeof v === "number") return v !== 0;
  const s = String(v).trim().toLowerCase();
  if (s === "false" || s === "0" || s === "no" || s === "n" || s === "off") return false;
  return true;
};
function calcScore(entry, schema) {
  const criteria = {};
  let metW = 0, totalW = 0;
  for (const c of schema) {
    if (!c.enabled) continue;
    const { cfg } = resolveField(c, entry);
    if (cfg.skip) { criteria[c.id] = null; continue; }
    // A habit only counts from the day it started being tracked. Days before
    // that are not penalised (or rewarded) for this field.
    if (cfg.trackingSince && entry.date && String(entry.date) < String(cfg.trackingSince)) {
      criteria[c.id] = null; continue;
    }
    if (cfg.skipHoliday && entry.isHoliday) { criteria[c.id] = null; continue; }
    const v = entry[c.id];
    let pct = 0; // 0..1 partial credit
    if (cfg.type === "toggle") {
      const b = coerceBool(v);
      pct = (cfg.inverted ? !b : b) ? 1 : 0;
    } else if (cfg.type === "time") {
      if (v) {
        const [h, m] = String(v).split(":").map(Number);
        if (!Number.isNaN(h) && !Number.isNaN(m)) {
          const mins = h*60 + m;
          pct = (cfg.op === "lte" ? mins <= cfg.threshold : mins >= cfg.threshold) ? 1 : 0;
        }
      }
    } else if (cfg.type === "number" || cfg.type === "stepper") {
      const num = Number(v || 0);
      const thr = (cfg.holidayThreshold !== undefined && entry.isHoliday) ? cfg.holidayThreshold : cfg.threshold;
      const binary = cfg.op === "lte" ? num <= thr : num >= thr;
      if (cfg.scoreMode === "graded" && Number(thr) > 0) {
        if (cfg.op === "gte") pct = Math.max(0, Math.min(1, num / thr));
        else /* lte */ pct = num <= thr ? 1 : Math.max(0, 1 - (num - thr) / thr);
      } else {
        pct = binary ? 1 : 0;
      }
    }
    criteria[c.id] = pct >= 0.999;
    const w = Number(cfg.weight) || 1;
    totalW += w;
    metW += w * pct;
  }
  const score = totalW ? (metW / totalW) * 100 : 0;
  const met = Object.values(criteria).filter(Boolean).length;
  const total = Object.values(criteria).filter(v => v !== null && v !== undefined).length;
  return { score, criteria, met, total };
}

/* ─────────────── HELPERS ─────────────── */
// All date helpers use *local* time. Avoid toISOString() (UTC) which breaks
// in non-UTC timezones (e.g. shows yesterday's date until 5:30am IST).
const toLocalISO = (d) => {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
};
const getToday   = () => toLocalISO(new Date());
const addDays    = (ds, n) => {
  // ds = "YYYY-MM-DD". Construct a local-midnight Date, add days, format back.
  const [y, m, d] = ds.split("-").map(Number);
  const dt = new Date(y, m - 1, d);
  dt.setDate(dt.getDate() + n);
  return toLocalISO(dt);
};
const fmtShort   = s => new Date(s + "T00:00:00").toLocaleDateString("en-US", { month: "short", day: "numeric" });
const fmtFull    = s => new Date(s + "T00:00:00").toLocaleDateString("en-US", { weekday: "short", month: "short", day: "numeric" });
const scoreColor = s => s === 100 ? C.cyan : s >= 75 ? C.green : s >= 50 ? C.amber : s > 0 ? C.red : C.border2;
const scoreLabel = s => s === 100 ? "FLAWLESS" : s >= 87.5 ? "ELITE" : s >= 75 ? "GREAT" : s >= 62.5 ? "GOOD" : s >= 50 ? "FAIR" : s >= 25 ? "POOR" : "MISS";
const minToTime  = m => `${String(Math.floor(m/60)).padStart(2,"0")}:${String(m%60).padStart(2,"0")}`;
const timeToMin  = t => { if (!t) return 0; const [h,m] = t.split(":").map(Number); return h*60+m; };

// Normalise *anything* into a strict "HH:MM" string the <input type="time">
// will accept. Sheets loves to convert time strings into Date objects or
// fractional day numbers when round-tripped, breaking the input.
const normalizeTime = (v) => {
  if (v === undefined || v === null || v === "") return "";
  if (v instanceof Date && !isNaN(v)) {
    return `${String(v.getHours()).padStart(2,"0")}:${String(v.getMinutes()).padStart(2,"0")}`;
  }
  if (typeof v === "number" && isFinite(v)) {
    // 0..1 day-fraction (Excel/Sheets time) — guard against accidental serial dates.
    if (v >= 0 && v < 1) {
      const total = Math.round(v * 24 * 60);
      return `${String(Math.floor(total/60)).padStart(2,"0")}:${String(total%60).padStart(2,"0")}`;
    }
    return "";
  }
  const s = String(v).trim();
  // Already HH:MM or H:MM, optionally with seconds.
  const m = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  if (m) return `${String(Math.min(23, Number(m[1]))).padStart(2,"0")}:${m[2]}`;
  // ISO-ish "2026-04-29T06:30:00" → grab time part.
  const iso = s.match(/T(\d{2}):(\d{2})/);
  if (iso) return `${iso[1]}:${iso[2]}`;
  // Last resort: let Date parse it.
  const d = new Date(s);
  if (!isNaN(d)) return `${String(d.getHours()).padStart(2,"0")}:${String(d.getMinutes()).padStart(2,"0")}`;
  return "";
};

// Walk every entry and coerce each schema-declared `time` field into HH:MM,
// and every `toggle` field into a real boolean. Sheets/xlsx round-trips often
// store booleans as the strings "true"/"false" — `!!"false"` is truthy, so we
// must normalise before anything renders or scores.
const normalizeEntry = (entry, schema) => {
  if (!entry) return entry;
  let changed = false;
  const out = { ...entry };
  for (const c of schema) {
    if (c.type === "time") {
      const cur = out[c.id];
      if (cur === undefined || cur === null || cur === "") continue;
      const fixed = normalizeTime(cur);
      if (fixed !== cur) { out[c.id] = fixed; changed = true; }
    } else if (c.type === "toggle") {
      const cur = out[c.id];
      if (cur === undefined || cur === null || cur === "") continue;
      const fixed = coerceBool(cur);
      if (fixed !== cur) { out[c.id] = fixed; changed = true; }
    }
  }
  // `isHoliday` is also a boolean that round-trips through Sheets cells.
  if (out.isHoliday !== undefined && out.isHoliday !== null && typeof out.isHoliday !== "boolean") {
    const fixed = coerceBool(out.isHoliday);
    if (fixed !== out.isHoliday) { out.isHoliday = fixed; changed = true; }
  }
  // `criteria` may have been stored as JSON-stringified booleans too.
  if (out.criteria && typeof out.criteria === "object") {
    let critChanged = false;
    const fixedCrit = {};
    for (const k of Object.keys(out.criteria)) {
      const cv = out.criteria[k];
      if (cv === null) { fixedCrit[k] = null; continue; }
      const nb = coerceBool(cv);
      if (nb !== cv) critChanged = true;
      fixedCrit[k] = nb;
    }
    if (critChanged) { out.criteria = fixedCrit; changed = true; }
  }
  return changed ? out : entry;
};

const QUOTES = [
  "Discipline equals freedom.",
  "Small daily improvements lead to stunning results.",
  "You don't rise to your goals, you fall to your systems.",
  "The pain of discipline weighs ounces. Regret weighs tons.",
  "Habits are the compound interest of self-improvement.",
  "Don't break the chain.",
  "Win the morning, win the day.",
  "Be the person your future self thanks.",
];

/* ─────────────── PRIMITIVES ─────────────── */
const inputSt = {
  width:"100%", padding:"10px 12px", borderRadius:10, border:`1.5px solid ${C.border2}`,
  background:C.card2, color:C.text, fontSize:14, boxSizing:"border-box", outline:"none",
  fontFamily:"inherit", transition:"border-color .2s, box-shadow .2s",
};

function Toggle({ value, onChange, disabled }) {
  return (
    <button onClick={() => !disabled && onChange(!value)} disabled={disabled} className="ht-toggle" style={{
      width:48, height:26, borderRadius:13, border:"none", cursor: disabled ? "not-allowed" : "pointer",
      position:"relative", padding:0,
      background: disabled ? C.border2 : value ? `linear-gradient(135deg,${C.violet},${C.purple})` : C.border2,
      transition:"background .25s, box-shadow .25s", flexShrink:0, opacity: disabled ? 0.35 : 1,
      boxShadow: value && !disabled ? `0 0 14px ${C.purple}88` : "none",
    }}>
      <div style={{ position:"absolute", top:3, width:20, height:20, borderRadius:10, background:C.white,
        left: value ? 25 : 3, transition:"left .25s cubic-bezier(.4,.8,.4,1.4)", boxShadow:"0 1px 4px rgba(0,0,0,.5)" }}/>
    </button>
  );
}

function Stepper({ value, onChange, disabled, max=8 }) {
  const btn = (label, fn) => (
    <button onClick={() => !disabled && fn()} className="ht-step-btn" style={{
      width:32, height:32, borderRadius:8, border:`1px solid ${C.border2}`, background:C.card3,
      cursor: disabled ? "not-allowed" : "pointer", fontSize:18, fontWeight:700, color:C.muted,
      display:"flex", alignItems:"center", justifyContent:"center",
      transition:"all .15s",
    }}>{label}</button>
  );
  return (
    <div style={{ display:"flex", alignItems:"center", gap:10, opacity: disabled ? 0.35 : 1 }}>
      {btn("−", () => onChange(Math.max(0, value - 1)))}
      <span style={{ width:28, textAlign:"center", fontSize:18, fontWeight:800, color: value > 0 ? C.purple : C.muted }}>{value}</span>
      {btn("+", () => onChange(Math.min(max, value + 1)))}
    </div>
  );
}

function ScoreRing({ score, size = 130, animate = true }) {
  const [displayScore, setDisplayScore] = useState(animate ? 0 : score);
  useEffect(() => {
    if (!animate) { setDisplayScore(score); return; }
    let raf, start;
    const from = displayScore, to = score, dur = 800;
    const tick = (t) => {
      if (!start) start = t;
      const p = Math.min(1, (t - start)/dur);
      setDisplayScore(from + (to - from) * (1 - Math.pow(1-p, 3)));
      if (p < 1) raf = requestAnimationFrame(tick);
    };
    raf = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(raf);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [score]);

  const r = (size - 18) / 2, circ = 2 * Math.PI * r;
  const off = circ - (displayScore / 100) * circ, c = scoreColor(displayScore), cx = size / 2;
  const gid = `g-${size}-${Math.round(c.charCodeAt(1))}`;
  return (
    <svg width={size} height={size} style={{ display:"block" }}>
      <defs>
        <linearGradient id={gid} x1="0" y1="0" x2="1" y2="1">
          <stop offset="0%" stopColor={c}/>
          <stop offset="100%" stopColor={C.purple}/>
        </linearGradient>
      </defs>
      <circle cx={cx} cy={cx} r={r} fill="none" stroke={C.border2} strokeWidth={11}/>
      <circle cx={cx} cy={cx} r={r} fill="none" stroke={`url(#${gid})`} strokeWidth={11}
        strokeDasharray={circ} strokeDashoffset={off} strokeLinecap="round"
        transform={`rotate(-90 ${cx} ${cx})`}
        style={{ filter:`drop-shadow(0 0 8px ${c}aa)` }}/>
      <text x={cx} y={cx-6} textAnchor="middle" fill={c} fontSize={size*.2} fontWeight="800" style={{ fontFamily:"'Space Grotesk', sans-serif" }}>{Math.round(displayScore)}%</text>
      <text x={cx} y={cx+13} textAnchor="middle" fill={c} fontSize={size*.085} fontWeight="700" letterSpacing="2">{scoreLabel(displayScore)}</text>
    </svg>
  );
}

function CriteriaIcons({ criteria, schema, size=24 }) {
  return (
    <div style={{ display:"flex", flexWrap:"wrap", gap:4 }}>
      {schema.filter(c=>c.enabled).map(({ id, icon, label }) => {
        const v = criteria?.[id];
        if (v === null || v === undefined) return <div key={id} title={label} style={{ width:size, height:size, borderRadius:7, fontSize:size*.5, display:"flex", alignItems:"center", justifyContent:"center", background:C.border, border:`1px solid ${C.border2}`, opacity:.35 }}>{icon}</div>;
        return <div key={id} title={`${label}: ${v ? "✓" : "✗"}`} style={{ width:size, height:size, borderRadius:7, fontSize:size*.5, display:"flex", alignItems:"center", justifyContent:"center", background: v ? "rgba(16,185,129,.15)" : "rgba(244,63,94,.1)", border:`1px solid ${v ? "rgba(16,185,129,.4)" : "rgba(244,63,94,.3)"}`, filter: v ? "none" : "grayscale(70%) opacity(55%)" }}>{icon}</div>;
      })}
    </div>
  );
}

const Div = () => <div style={{ height:1, background:C.border, margin:"2px 0" }}/>;
const Label = ({ children, accent }) => <div style={{ fontSize:11, fontWeight:700, letterSpacing:1.2, color: accent || C.muted, textTransform:"uppercase", marginBottom:6, fontFamily:"'Space Grotesk', sans-serif" }}>{children}</div>;
const Card = ({ children, style, glow }) => (
  <div className="ht-card" style={{
    background:`linear-gradient(135deg,${C.card},${C.card2})`,
    border:`1px solid ${C.border}`, borderRadius:16, padding:16,
    display:"flex", flexDirection:"column", gap:14,
    boxShadow: glow ? `0 0 24px ${glow}33` : "0 4px 16px rgba(0,0,0,.3)",
    transition:"box-shadow .25s, border-color .25s, transform .25s",
    ...style
  }}>{children}</div>
);

/* Segmented radio (super cool) */
function Segmented({ options, value, onChange }) {
  return (
    <div style={{
      display:"inline-flex", padding:4, borderRadius:12, background:C.card3,
      border:`1px solid ${C.border2}`, gap:2,
    }}>
      {options.map(o => {
        const active = o.value === value;
        return (
          <button key={o.value} onClick={() => onChange(o.value)} style={{
            padding:"6px 14px", borderRadius:9, border:"none",
            background: active ? `linear-gradient(135deg,${C.violet},${C.purple})` : "transparent",
            color: active ? C.white : C.muted, cursor:"pointer", fontWeight:700, fontSize:12,
            transition:"all .2s", letterSpacing:.5,
            boxShadow: active ? `0 4px 12px ${C.purple}55` : "none",
            fontFamily:"inherit",
          }}>{o.label}</button>
        );
      })}
    </div>
  );
}

/* ─────────────── STATS ─────────────── */
function calcStreaks(scored) {
  const asc = [...scored].sort((a,b) => a.date.localeCompare(b.date));
  let best = 0, run = 0, perfect = 0;
  asc.forEach(e => {
    if (e.score >= 75) { run++; best = Math.max(best, run); } else run = 0;
    if (e.score === 100) perfect++;
  });
  let cur = 0;
  for (let i = asc.length - 1; i >= 0; i--) {
    if (asc[i].score >= 75) cur++; else break;
  }
  const last7 = asc.slice(-7);
  const last30 = asc.slice(-30);
  const avg = arr => arr.length ? Math.round(arr.reduce((s,e)=>s+e.score,0)/arr.length) : 0;
  return { cur, best, perfect, avg7: avg(last7), avg30: avg(last30), total: asc.length };
}

function StatCard({ icon, value, label, color }) {
  return (
    <div className="ht-stat" style={{
      flex:1, minWidth:0, background:`linear-gradient(135deg,${C.card2},${C.card3})`,
      border:`1px solid ${C.border}`, borderRadius:14, padding:"14px 10px", textAlign:"center",
      transition:"all .25s", cursor:"default",
    }}>
      <div style={{ fontSize:22, marginBottom:2 }}>{icon}</div>
      <div style={{ fontSize:22, fontWeight:800, color: color || C.text, lineHeight:1.1, fontFamily:"'Space Grotesk', sans-serif" }}>{value}</div>
      <div style={{ fontSize:10, color:C.muted, fontWeight:700, letterSpacing:1, textTransform:"uppercase", marginTop:2 }}>{label}</div>
    </div>
  );
}

/* ─────────────── CALENDAR ─────────────── */
function Calendar({ scored, onDay, cursor, setCursor, compact }) {
  // Optional controlled cursor; fall back to internal.
  const [internal, setInternal] = useState(() => { const d = new Date(); d.setDate(1); return d; });
  const cur = cursor || internal;
  const setCur = setCursor || setInternal;
  const map = useMemo(() => Object.fromEntries(scored.map(e => [e.date, e])), [scored]);
  const year = cur.getFullYear(), month = cur.getMonth();
  const firstDay = new Date(year, month, 1).getDay();
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const monthName = cur.toLocaleDateString("en-US", { month:"long", year:"numeric" });
  const today = getToday();

  const cells = [];
  for (let i = 0; i < firstDay; i++) cells.push(null);
  for (let d = 1; d <= daysInMonth; d++) {
    const ds = `${year}-${String(month+1).padStart(2,"0")}-${String(d).padStart(2,"0")}`;
    cells.push({ d, ds, e: map[ds] });
  }

  const navBtn = (label, fn) => (
    <button onClick={fn} className="ht-icon-btn" style={{
      width:32, height:32, borderRadius:8, border:`1px solid ${C.border2}`,
      background:C.card3, color:C.text, cursor:"pointer", fontSize:16, fontWeight:700, fontFamily:"inherit",
    }}>{label}</button>
  );

  return (
    <Card>
      {!compact && (
        <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between" }}>
          <Label accent={C.cyan}>📅 Calendar</Label>
          <div style={{ display:"flex", alignItems:"center", gap:8 }}>
            {navBtn("‹", () => setCur(new Date(year, month-1, 1)))}
            <span style={{ fontSize:13, fontWeight:700, color:C.text, minWidth:130, textAlign:"center", fontFamily:"'Space Grotesk', sans-serif" }}>{monthName}</span>
            {navBtn("›", () => setCur(new Date(year, month+1, 1)))}
          </div>
        </div>
      )}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(7,1fr)", gap:4 }}>
        {["S","M","T","W","T","F","S"].map((d,i) => (
          <div key={i} style={{ textAlign:"center", fontSize:10, fontWeight:700, color:C.muted, padding:"4px 0", letterSpacing:1 }}>{d}</div>
        ))}
        {cells.map((c, i) => {
          if (!c) return <div key={i}/>;
          const isToday = c.ds === today;
          const has = !!c.e;
          const score = c.e?.score || 0;
          const opacity = has ? Math.max(0.35, score/100) : 0.5;
          const bg = has ? scoreColor(score) : C.card3;
          const bgWithOpacity = has ? `${bg}${Math.round(opacity*255).toString(16).padStart(2,"0")}` : C.card3;
          return (
            <button key={i} onClick={() => onDay && onDay(c.ds)} className="ht-cal-cell" style={{
              aspectRatio:"1", borderRadius:8,
              border: isToday ? `2px solid ${C.cyan}` : `1px solid ${C.border}`,
              background: bgWithOpacity,
              color: has ? C.white : C.muted2, cursor:"pointer",
              fontSize:13, fontWeight:700, display:"flex", flexDirection:"column",
              alignItems:"center", justifyContent:"center", padding:0,
              boxShadow: isToday ? `0 0 12px ${C.cyan}88` : "none",
              transition:"all .2s", fontFamily:"'Space Grotesk', sans-serif",
              "--glow": has ? scoreColor(score) : C.purple,
            }}
            title={has ? `${c.ds} — ${Math.round(score)}%` : c.ds}
            >
              <div>{c.d}</div>
              {has && <div style={{ fontSize:8, opacity:.85, fontWeight:600 }}>{Math.round(score)}</div>}
            </button>
          );
        })}
      </div>
      <div style={{ display:"flex", alignItems:"center", justifyContent:"center", gap:8, fontSize:10, color:C.muted, marginTop:4 }}>
        <span>Less</span>
        {[15, 40, 60, 80, 100].map(v => (
          <div key={v} style={{ width:14, height:14, borderRadius:4, background:scoreColor(v), opacity: Math.max(0.35, v/100) }}/>
        ))}
        <span>More</span>
      </div>
    </Card>
  );
}

/* ─────────────── DAY DETAIL MODAL ───────────────
 * Click a calendar cell → see the full breakdown for that day with prev/next
 * navigation. Edit jumps to the Log tab pre-filled. Close dismisses.
 */
function DayDetailModal({ open, ds, scored, schema, onClose, onPrev, onNext, onEdit }) {
  // Keyboard nav: ←/→ shift, Esc closes.
  useEffect(() => {
    if (!open) return;
    const onKey = (e) => {
      if (e.key === "Escape") onClose?.();
      else if (e.key === "ArrowLeft") onPrev?.();
      else if (e.key === "ArrowRight") onNext?.();
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [open, onClose, onPrev, onNext]);
  if (!open || !ds) return null;
  const entry = scored.find(e => e.date === ds);
  const score = entry?.score || 0;
  const enabled = schema.filter(c => c.enabled);

  // pretty-print logged value per field
  const formatValue = (cfg, v) => {
    if (v === undefined || v === null || v === "") return "—";
    if (cfg.type === "toggle") return v ? "Yes" : "No";
    if (cfg.type === "time") return String(v);
    if (cfg.type === "number" || cfg.type === "stepper") return `${v}${cfg.unit ? " " + cfg.unit : ""}`;
    return String(v);
  };
  const targetSummary = (cfg) => {
    if (cfg.type === "toggle") return cfg.inverted ? "Avoid today" : "Complete today";
    const lbl = cfg.targetLabel || (cfg.op === "lte" ? "Limit" : "Goal");
    if (cfg.type === "time") return `${lbl} ${cfg.op === "lte" ? "≤" : "≥"} ${minToTime(cfg.threshold)}`;
    return `${lbl} ${cfg.op === "lte" ? "≤" : "≥"} ${cfg.threshold}${cfg.unit ? " " + cfg.unit : ""}`;
  };

  return (
    <div onClick={onClose} style={{
      position:"fixed", inset:0, background:"rgba(0,0,0,.7)", backdropFilter:"blur(8px)",
      display:"flex", alignItems:"center", justifyContent:"center", zIndex:100, padding:16,
      animation:"fadeIn .2s ease",
    }}>
      <div onClick={e => e.stopPropagation()} style={{
        background:`linear-gradient(135deg,${C.card},${C.card2})`,
        border:`1px solid ${C.border2}`, borderRadius:18, padding:0,
        maxWidth:560, width:"100%", maxHeight:"90vh", overflow:"hidden",
        display:"flex", flexDirection:"column",
        boxShadow:`0 20px 60px ${scoreColor(score)}55`, animation:"slideUp .25s ease",
      }}>
        {/* Header */}
        <div style={{
          padding:"16px 18px", display:"flex", alignItems:"center", gap:12,
          borderBottom:`1px solid ${C.border}`,
          background:`linear-gradient(135deg,${scoreColor(score)}15,transparent)`,
        }}>
          <button onClick={onPrev} title="Previous day" className="ht-icon-btn" style={{
            width:36, height:36, borderRadius:9, border:`1px solid ${C.border2}`,
            background:C.card3, color:C.text, cursor:"pointer", fontSize:18, fontWeight:700, fontFamily:"inherit",
          }}>‹</button>
          <div style={{ flex:1, minWidth:0, textAlign:"center" }}>
            <div style={{ fontSize:11, color:C.muted, fontWeight:700, letterSpacing:1.5, textTransform:"uppercase" }}>Day Detail</div>
            <div style={{ fontSize:17, fontWeight:800, color:C.text, fontFamily:"'Space Grotesk', sans-serif", marginTop:2 }}>{fmtFull(ds)}</div>
          </div>
          <button onClick={onNext} title="Next day" className="ht-icon-btn" style={{
            width:36, height:36, borderRadius:9, border:`1px solid ${C.border2}`,
            background:C.card3, color:C.text, cursor:"pointer", fontSize:18, fontWeight:700, fontFamily:"inherit",
          }}>›</button>
          <button onClick={onClose} title="Close" className="ht-icon-btn" style={{
            width:36, height:36, borderRadius:9, border:`1px solid ${C.border2}`,
            background:C.card3, color:C.text, cursor:"pointer", fontSize:20, fontWeight:700, fontFamily:"inherit",
          }}>×</button>
        </div>

        {/* Body */}
        <div style={{ padding:18, overflowY:"auto", display:"flex", flexDirection:"column", gap:14 }}>
          {entry ? (
            <>
              <div style={{ display:"flex", alignItems:"center", gap:16,
                padding:14, borderRadius:14,
                background:`linear-gradient(135deg,${C.card2},${C.card3})`,
                border:`1px solid ${scoreColor(score)}55`,
                boxShadow:`0 0 24px ${scoreColor(score)}22`,
              }}>
                <ScoreRing score={score} size={96}/>
                <div style={{ flex:1, minWidth:0 }}>
                  <div style={{ fontSize:11, color:C.muted, fontWeight:700, letterSpacing:1.5, textTransform:"uppercase" }}>Discipline</div>
                  <div style={{ fontSize:24, fontWeight:800, color:scoreColor(score), fontFamily:"'Space Grotesk', sans-serif" }}>
                    {Math.round(score)}% · {scoreLabel(score)}
                  </div>
                  <div style={{ fontSize:12, color:C.muted, marginTop:4 }}>
                    {entry.met}/{entry.total} habits met{entry.isHoliday ? " · 🏖️ Holiday" : ""}
                  </div>
                </div>
              </div>

              <div>
                <Label accent={C.purple}>📋 Subject Card</Label>
                <div style={{ display:"flex", flexDirection:"column", gap:6 }}>
                  {enabled.map(c => {
                    const { cfg } = resolveField(c, entry);
                    const status = entry.criteria?.[c.id]; // true | false | null
                    const skip = status === null || status === undefined;
                    const dotColor = skip ? C.muted2 : status ? C.green : C.red;
                    const grade = skip ? "—" : status ? "✓ MET" : "✗ MISS";
                    return (
                      <div key={c.id} style={{
                        display:"flex", alignItems:"center", gap:10,
                        padding:"10px 12px", borderRadius:10,
                        background:C.card3, border:`1px solid ${C.border}`,
                        borderLeft:`4px solid ${dotColor}`,
                        opacity: skip ? .55 : 1,
                      }}>
                        <div style={{ fontSize:20 }}>{cfg.icon}</div>
                        <div style={{ flex:1, minWidth:0 }}>
                          <div style={{ fontSize:13, fontWeight:700, color:C.text }}>{cfg.label}</div>
                          <div style={{ fontSize:10, color:C.muted, marginTop:2 }}>{targetSummary(cfg)}</div>
                        </div>
                        <div style={{ textAlign:"right" }}>
                          <div style={{ fontSize:13, color:C.text, fontWeight:700, fontFamily:"'Space Grotesk', sans-serif" }}>{formatValue(cfg, entry[c.id])}</div>
                          <div style={{ fontSize:9, fontWeight:800, color:dotColor, letterSpacing:1, marginTop:2 }}>{grade}</div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            </>
          ) : (
            <div style={{ padding:40, textAlign:"center" }}>
              <div style={{ fontSize:48 }}>📭</div>
              <div style={{ fontSize:15, fontWeight:700, color:C.text, marginTop:6, fontFamily:"'Space Grotesk', sans-serif" }}>No entry for this day</div>
              <div style={{ fontSize:12, color:C.muted, marginTop:4 }}>Hit Edit to log it now.</div>
            </div>
          )}
        </div>

        {/* Footer */}
        <div style={{
          padding:"12px 18px", borderTop:`1px solid ${C.border}`,
          display:"flex", gap:10, justifyContent:"flex-end",
        }}>
          <button onClick={onClose} className="ht-chip" style={{
            padding:"8px 16px", borderRadius:9, border:`1px solid ${C.border2}`,
            background:C.card3, color:C.muted, fontWeight:700, fontSize:13, cursor:"pointer", fontFamily:"inherit",
          }}>Close</button>
          <button onClick={onEdit} className="ht-cta" style={{
            padding:"8px 16px", borderRadius:9, border:"none",
            background:`linear-gradient(135deg,${C.violet},${C.purple})`, color:C.white,
            fontWeight:800, fontSize:13, cursor:"pointer", fontFamily:"inherit",
            boxShadow:`0 4px 12px ${C.purple}55`,
          }}>{entry ? "✏️ Edit Day" : "✨ Log Day"}</button>
        </div>
      </div>
    </div>
  );
}

/* ─────────────── INFO MODAL ─────────────── */
function InfoModal({ open, onClose, schema }) {
  if (!open) return null;
  const totalW = schema.filter(c=>c.enabled).reduce((s,c)=>s+(Number(c.weight)||1), 0);
  const desc = c => {
    if (c.type === "toggle") return c.inverted ? "Met when OFF" : "Met when ON";
    if (c.type === "time") return `${c.op === "lte" ? "≤" : "≥"} ${minToTime(c.threshold)}`;
    if (c.type === "stepper" || c.type === "number") {
      let s = `${c.op === "lte" ? "≤" : "≥"} ${c.threshold}${c.unit ? " " + c.unit : ""}`;
      if (c.holidayThreshold !== undefined) s += ` (Holiday: ≤ ${c.holidayThreshold})`;
      return s;
    }
    return "";
  };
  return (
    <div onClick={onClose} style={{
      position:"fixed", inset:0, background:"rgba(0,0,0,.7)", backdropFilter:"blur(8px)",
      display:"flex", alignItems:"center", justifyContent:"center", zIndex:100, padding:16,
      animation:"fadeIn .2s ease",
    }}>
      <div onClick={e=>e.stopPropagation()} style={{
        background:`linear-gradient(135deg,${C.card},${C.card2})`,
        border:`1px solid ${C.border2}`, borderRadius:18, padding:20,
        maxWidth:520, width:"100%", maxHeight:"85vh", overflowY:"auto",
        boxShadow:`0 20px 60px ${C.purple}44`, animation:"slideUp .25s ease",
      }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14 }}>
          <div>
            <div style={{ fontSize:18, fontWeight:800, color:C.text, fontFamily:"'Space Grotesk', sans-serif" }}>📊 How Scoring Works</div>
            <div style={{ fontSize:12, color:C.muted, marginTop:2 }}>Each habit met / total weight × 100</div>
          </div>
          <button onClick={onClose} className="ht-icon-btn" style={{
            width:32, height:32, borderRadius:8, border:`1px solid ${C.border2}`,
            background:C.card3, color:C.text, cursor:"pointer", fontSize:18, fontFamily:"inherit",
          }}>×</button>
        </div>
        <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
          {schema.filter(c=>c.enabled).map(c => (
            <div key={c.id} style={{
              display:"flex", alignItems:"center", gap:12, padding:"10px 12px",
              background:C.card3, borderRadius:10, border:`1px solid ${C.border}`,
            }}>
              <div style={{ fontSize:22 }}>{c.icon}</div>
              <div style={{ flex:1, minWidth:0 }}>
                <div style={{ fontSize:13, fontWeight:700, color:C.text }}>{c.label}</div>
                <div style={{ fontSize:11, color:C.muted, marginTop:2 }}>{desc(c)}</div>
              </div>
              <div style={{
                padding:"3px 10px", borderRadius:6, background:`${C.purple}22`,
                border:`1px solid ${C.purple}44`, fontSize:11, fontWeight:700, color:C.purple,
              }}>{Math.round((c.weight||1)/totalW*100)}%</div>
            </div>
          ))}
        </div>
        <div style={{ marginTop:16, padding:12, borderRadius:10, background:`${C.cyan}11`, border:`1px solid ${C.cyan}33`, fontSize:12, color:C.text }}>
          💡 <b>Tier:</b> 100% Flawless · ≥87.5% Elite · ≥75% Great · ≥62.5% Good · ≥50% Fair
        </div>
        <div style={{ marginTop:8, fontSize:11, color:C.muted, textAlign:"center" }}>
          Edit habits & weights in the <b style={{ color:C.purple }}>Engine</b> tab.
        </div>
      </div>
    </div>
  );
}

/* ─────────────── DASHBOARD ─────────────── */
function Dashboard({ scored, schema, onGoLog, onPickDate }) {
  const stats = useMemo(() => calcStreaks(scored), [scored]);
  const today = getToday();
  const todayEntry = scored.find(e => e.date === today);
  const todayScore = todayEntry?.score || 0;

  // Shared month cursor → drives Calendar, Trend chart, and Win-Rate widget.
  const [monthCursor, setMonthCursor] = useState(() => { const d = new Date(); d.setDate(1); return d; });
  const mYear = monthCursor.getFullYear();
  const mMonth = monthCursor.getMonth();
  const monthLabel = monthCursor.toLocaleDateString("en-US", { month:"long", year:"numeric" });
  const isCurrentMonth = mYear === new Date().getFullYear() && mMonth === new Date().getMonth();
  const monthEntries = useMemo(
    () => scored.filter(e => {
      const [y, m] = e.date.split("-").map(Number);
      return y === mYear && (m - 1) === mMonth;
    }),
    [scored, mYear, mMonth]
  );

  const trend = useMemo(() => {
    const asc = [...monthEntries].sort((a,b) => a.date.localeCompare(b.date));
    return asc.map(e => ({ date: fmtShort(e.date), score: Math.round(e.score) }));
  }, [monthEntries]);

  const criteriaStats = useMemo(() => {
    return schema.filter(c=>c.enabled).map(({ id, icon, label }) => {
      const valid = monthEntries.filter(e => e.criteria?.[id] !== null && e.criteria?.[id] !== undefined);
      const met = valid.filter(e => e.criteria?.[id]).length;
      const pct = valid.length ? Math.round((met / valid.length) * 100) : 0;
      return { id, icon, label, pct, n: valid.length };
    }).sort((a,b) => b.pct - a.pct);
  }, [monthEntries, schema]);

  const monthNav = (
    <div style={{ display:"flex", alignItems:"center", gap:6 }}>
      <button onClick={() => setMonthCursor(new Date(mYear, mMonth-1, 1))} className="ht-icon-btn" style={{
        width:28, height:28, borderRadius:7, border:`1px solid ${C.border2}`,
        background:C.card3, color:C.text, cursor:"pointer", fontSize:14, fontWeight:700, fontFamily:"inherit",
      }}>‹</button>
      <span style={{ fontSize:11, fontWeight:800, color:C.text, minWidth:110, textAlign:"center", letterSpacing:.5, fontFamily:"'Space Grotesk', sans-serif" }}>{monthLabel}</span>
      <button onClick={() => setMonthCursor(new Date(mYear, mMonth+1, 1))} className="ht-icon-btn" style={{
        width:28, height:28, borderRadius:7, border:`1px solid ${C.border2}`,
        background:C.card3, color:C.text, cursor:"pointer", fontSize:14, fontWeight:700, fontFamily:"inherit",
      }}>›</button>
      {!isCurrentMonth && (
        <button onClick={() => setMonthCursor((() => { const d = new Date(); d.setDate(1); return d; })())} style={{
          padding:"4px 10px", borderRadius:6, border:`1px solid ${C.cyan}33`,
          background:`${C.cyan}15`, color:C.cyan, fontSize:10, fontWeight:800, cursor:"pointer", fontFamily:"inherit",
        }}>Today</button>
      )}
    </div>
  );

  const quote = QUOTES[new Date().getDate() % QUOTES.length];
  const motivation =
    todayScore === 100 ? "🔥 Flawless. You're built different." :
    todayScore >= 75   ? "⚡ Strong work. Don't ease up." :
    todayScore >= 50   ? "💪 Halfway there. Push for more." :
    todayScore > 0     ? "⚠️ Today's slipping. Reclaim it." :
                         "🎯 Fresh slate. Start logging.";

  return (
    <div className="ht-dash">
      {/* Hero (full-width) */}
      <div className="ht-hero ht-card-glow" style={{
        gridColumn:"1 / -1",
        background:`linear-gradient(135deg,${C.card2},${C.card3})`,
        border:`1px solid ${scoreColor(todayScore)}55`, borderRadius:18, padding:20,
        display:"flex", alignItems:"center", gap:20,
        boxShadow:`0 0 32px ${scoreColor(todayScore)}22`,
      }}>
        <ScoreRing score={todayScore} size={140}/>
        <div style={{ flex:1, minWidth:0 }}>
          <div style={{ fontSize:11, color:C.muted, fontWeight:700, letterSpacing:1.5, textTransform:"uppercase" }}>Today</div>
          <div style={{ fontSize:20, fontWeight:800, color:C.text, marginTop:2, fontFamily:"'Space Grotesk', sans-serif" }}>{fmtFull(today)}</div>
          <div style={{ fontSize:13, color:scoreColor(todayScore), fontWeight:700, marginTop:8 }}>{motivation}</div>
          <div style={{ fontSize:12, color:C.muted, fontStyle:"italic", marginTop:4 }}>"{quote}"</div>
          <button onClick={onGoLog} className="ht-cta" style={{
            marginTop:12, padding:"10px 18px", borderRadius:10, border:"none",
            background:`linear-gradient(135deg,${C.violet},${C.purple})`, color:C.white,
            fontWeight:700, fontSize:13, cursor:"pointer", fontFamily:"inherit",
            boxShadow:`0 4px 12px ${C.purple}55`,
          }}>{todayEntry ? "✏️ Edit Today" : "✨ Log Today"}</button>
        </div>
      </div>

      {/* Stats grid */}
      <div className="ht-stats-grid" style={{ gridColumn:"1 / -1" }}>
        <StatCard icon="🔥" value={stats.cur} label="Streak" color={C.amber}/>
        <StatCard icon="🏆" value={stats.best} label="Best" color={C.cyan}/>
        <StatCard icon="💯" value={stats.perfect} label="Flawless" color={C.green}/>
        <StatCard icon="📈" value={`${stats.avg7}%`} label="7-Day Avg" color={C.purple}/>
        <StatCard icon="📊" value={`${stats.avg30}%`} label="30-Day Avg" color={C.pink}/>
        <StatCard icon="📅" value={stats.total} label="Total Days" color={C.text}/>
      </div>

      {/* Trend chart */}
      <Card>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", flexWrap:"wrap", gap:8 }}>
          <Label accent={C.cyan}>📈 {monthLabel} Trend</Label>
          {monthNav}
        </div>
        {trend.length > 0 ? (
          <div style={{ width:"100%", height:220 }}>
            <ResponsiveContainer>
              <AreaChart data={trend} margin={{ top:10, right:10, left:-20, bottom:0 }}>
                <defs>
                  <linearGradient id="scoreGrad" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="0%" stopColor={C.purple} stopOpacity={0.7}/>
                    <stop offset="100%" stopColor={C.purple} stopOpacity={0.05}/>
                  </linearGradient>
                </defs>
                <CartesianGrid stroke={C.border} strokeDasharray="3 3" vertical={false}/>
                <XAxis dataKey="date" stroke={C.muted} fontSize={10} tickLine={false} axisLine={false}/>
                <YAxis stroke={C.muted} fontSize={10} tickLine={false} axisLine={false} domain={[0,100]}/>
                <Tooltip
                  contentStyle={{ background:C.card3, border:`1px solid ${C.border2}`, borderRadius:8, fontSize:12 }}
                  labelStyle={{ color:C.text, fontWeight:700 }}
                  itemStyle={{ color:C.purple }}
                />
                <Area type="monotone" dataKey="score" stroke={C.purple} strokeWidth={2.5} fill="url(#scoreGrad)"/>
              </AreaChart>
            </ResponsiveContainer>
          </div>
        ) : (
          <div style={{ padding:"40px 20px", textAlign:"center", color:C.muted, fontSize:13 }}>
            📭 No entries logged in {monthLabel}.
          </div>
        )}
      </Card>

      {/* Calendar (shares the same monthCursor) */}
      <Calendar scored={scored} onDay={onPickDate} cursor={monthCursor} setCursor={setMonthCursor}/>

      {/* Habit win rate (full-width on desktop) — scoped to current month cursor */}
      <Card style={{ gridColumn:"1 / -1" }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", flexWrap:"wrap", gap:8 }}>
          <Label accent={C.green}>🎯 Habit Win Rate · {monthLabel}</Label>
          {monthNav}
        </div>
        {monthEntries.length === 0 ? (
          <div style={{ padding:"24px 8px", textAlign:"center", color:C.muted, fontSize:13 }}>
            📭 No entries to analyse for {monthLabel}.
          </div>
        ) : (
          <div className="ht-habit-grid">
            {criteriaStats.map(c => (
              <div key={c.id} style={{ display:"flex", alignItems:"center", gap:10 }}>
                <div style={{ width:24, fontSize:18 }}>{c.icon}</div>
                <div style={{ flex:1, minWidth:0 }}>
                  <div style={{ display:"flex", justifyContent:"space-between", marginBottom:4 }}>
                    <span style={{ fontSize:12, color:C.text, fontWeight:600 }}>{c.label}</span>
                    <span style={{ fontSize:12, color:scoreColor(c.pct), fontWeight:800, fontFamily:"'Space Grotesk', sans-serif" }}>{c.pct}% <span style={{ color:C.muted, fontWeight:600, fontSize:10 }}>({c.n})</span></span>
                  </div>
                  <div style={{ height:6, borderRadius:3, background:C.border, overflow:"hidden" }}>
                    <div style={{
                      height:"100%", width:`${c.pct}%`,
                      background:`linear-gradient(90deg,${scoreColor(c.pct)},${C.purple})`,
                      transition:"width .8s ease",
                      boxShadow:`0 0 8px ${scoreColor(c.pct)}aa`,
                    }}/>
                  </div>
                </div>
              </div>
            ))}
          </div>
        )}
      </Card>

      {scored.length === 0 && (
        <Card style={{ gridColumn:"1 / -1", alignItems:"center", textAlign:"center", padding:40 }}>
          <div style={{ fontSize:48 }}>🚀</div>
          <div style={{ fontSize:18, fontWeight:700, color:C.text, fontFamily:"'Space Grotesk', sans-serif" }}>No data yet</div>
          <div style={{ fontSize:13, color:C.muted }}>Log your first day to start tracking discipline.</div>
          <button onClick={onGoLog} className="ht-cta" style={{
            marginTop:6, padding:"10px 20px", borderRadius:10, border:"none",
            background:`linear-gradient(135deg,${C.violet},${C.purple})`, color:C.white,
            fontWeight:700, cursor:"pointer", fontFamily:"inherit",
          }}>Start Now</button>
        </Card>
      )}
    </div>
  );
}

/* ─────────────── LOG ENTRY (schema-driven) ─────────────── */
function CriterionField({ cfg, value, onChange, isHoliday }) {
  const disabled = (cfg.skipHoliday && isHoliday) || cfg.skip;
  if (disabled) return (
    <div style={{ fontSize:11, color:C.muted, fontStyle:"italic" }}>{cfg.skip ? "Skipped by rule" : "Skipped on holidays"}</div>
  );

  if (cfg.type === "toggle") {
    return <Toggle value={cfg.inverted ? !value : !!value} onChange={v => onChange(cfg.inverted ? !v : v)}/>;
  }
  if (cfg.type === "stepper") {
    return <Stepper value={Number(value) || 0} onChange={onChange} max={cfg.max || 16}/>;
  }
  if (cfg.type === "time") {
    return <input type="time" value={value || ""} onChange={e => onChange(e.target.value)} style={{ ...inputSt, width:140 }}/>;
  }
  // number
  return (
    <input type="number" step={cfg.step || 1} value={value ?? ""} placeholder="0"
      onChange={e => onChange(e.target.value)}
      style={{ ...inputSt, width:120, textAlign:"right" }}/>
  );
}

/* Target badge — surfaces the *current effective* target (after rules) per field */
function TargetBadge({ cfg, value }) {
  if (cfg.type === "toggle") return null;
  const lbl = cfg.targetLabel || (cfg.op === "lte" ? "Limit" : "Goal");
  let display, color = C.cyan, met = false;
  if (cfg.type === "time") {
    display = minToTime(cfg.threshold || 0);
    if (value) {
      const [h,m] = String(value).split(":").map(Number);
      if (!isNaN(h) && !isNaN(m)) {
        const mins = h*60+m;
        met = cfg.op === "lte" ? mins <= cfg.threshold : mins >= cfg.threshold;
      }
    }
  } else {
    const num = Number(value || 0);
    display = `${cfg.threshold}${cfg.unit ? " " + cfg.unit : ""}`;
    met = cfg.op === "lte" ? num <= cfg.threshold && num > 0 : num >= cfg.threshold;
    // for graded fields show progress
    if (cfg.scoreMode === "graded" && cfg.op === "gte" && cfg.threshold > 0) {
      const pct = Math.min(100, Math.round((num / cfg.threshold) * 100));
      display = `${num || 0} / ${cfg.threshold}${cfg.unit ? " " + cfg.unit : ""} · ${pct}%`;
    }
  }
  color = met ? C.green : (cfg.op === "lte" ? C.amber : C.cyan);
  return (
    <span style={{
      display:"inline-flex", alignItems:"center", gap:4,
      padding:"3px 8px", borderRadius:6, fontSize:10, fontWeight:800, letterSpacing:.5,
      background:`${color}15`, border:`1px solid ${color}44`, color, fontFamily:"'Space Grotesk', sans-serif",
      whiteSpace:"nowrap",
    }}>
      🎯 {lbl}: {display}
    </span>
  );
}

function LogEntry({ form, setForm, live, schema, onDateChange, onSave, saved, onQuickFill, hasYesterday }) {
  const f = (k, v) => setForm(p => ({ ...p, [k]: v }));
  const isFlawless = live.score === 100;

  // group fields into 2 columns on desktop
  const enabled = schema.filter(c => c.enabled);

  return (
    <div className="ht-log">
      {/* Live preview */}
      <div className="ht-live" style={{
        gridColumn:"1 / -1",
        background: isFlawless
          ? `linear-gradient(135deg,${C.cyan}22,${C.purple}22)`
          : `linear-gradient(135deg,${C.card},${C.card2})`,
        border:`1px solid ${isFlawless ? C.cyan : C.border}`,
        borderRadius:16, padding:18, display:"flex", alignItems:"center", gap:18,
        transition:"all .5s",
        boxShadow: isFlawless ? `0 0 32px ${C.cyan}44` : "none",
      }}>
        <ScoreRing score={live.score} size={110}/>
        <div style={{ flex:1, minWidth:0 }}>
          <Label>Live Preview</Label>
          <div style={{ fontSize:14, fontWeight:700, color:scoreColor(live.score), marginBottom:8, fontFamily:"'Space Grotesk', sans-serif" }}>
            {live.met}/{live.total} habits met
          </div>
          <CriteriaIcons criteria={live.criteria} schema={schema} size={24}/>
        </div>
        {isFlawless && (
          <div className="ht-pulse" style={{
            padding:"12px 16px", borderRadius:10,
            background:`linear-gradient(135deg,${C.cyan}33,${C.purple}33)`,
            border:`1px solid ${C.cyan}66`, color:C.cyan,
            fontWeight:800, letterSpacing:2, fontSize:14, whiteSpace:"nowrap",
          }}>
            🎯 FLAWLESS
          </div>
        )}
      </div>

      {/* Date / Holiday */}
      <Card style={{ gridColumn:"1 / -1" }}>
        <div style={{ display:"flex", flexWrap:"wrap", justifyContent:"space-between", alignItems:"center", gap:12 }}>
          <Label>📅 Date & Mode</Label>
          {hasYesterday && (
            <button onClick={onQuickFill} className="ht-chip" style={{
              fontSize:11, padding:"5px 12px", borderRadius:6, border:`1px solid ${C.border2}`,
              background:C.card3, color:C.cyan, cursor:"pointer", fontWeight:700, fontFamily:"inherit",
            }}>↻ Copy Yesterday</button>
          )}
        </div>
        <div style={{ display:"flex", flexWrap:"wrap", gap:14, alignItems:"center" }}>
          <input type="date" value={form.date} onChange={e => onDateChange(e.target.value)} style={{ ...inputSt, width:200 }}/>
          <div style={{ display:"flex", alignItems:"center", gap:10 }}>
            <span style={{ fontSize:13, color:C.text, fontWeight:600 }}>🏖️ Holiday / Weekend</span>
            <Toggle value={form.isHoliday} onChange={v => f("isHoliday", v)}/>
          </div>
        </div>
      </Card>

      {/* Habits as 2-col on desktop */}
      <div className="ht-fields" style={{ gridColumn:"1 / -1" }}>
        {enabled.map(c => {
          const { cfg, activeRules } = resolveField(c, form);
          const skipped = (cfg.skipHoliday && form.isHoliday) || cfg.skip;
          return (
            <div key={c.id} className="ht-field-row" style={{
              display:"flex", alignItems:"center", justifyContent:"space-between", gap:12,
              padding:"12px 14px",
              background: activeRules.length
                ? `linear-gradient(135deg,${C.card2},${C.card3})`
                : `linear-gradient(135deg,${C.card},${C.card2})`,
              border: activeRules.length ? `1px solid ${C.purple}55` : `1px solid ${C.border}`,
              borderRadius:12,
              transition:"border-color .2s, box-shadow .2s",
              opacity: skipped ? .45 : 1,
              boxShadow: activeRules.length ? `0 0 14px ${C.purple}22` : "none",
            }}>
              <div style={{ flex:1, minWidth:0 }}>
                <div style={{ display:"flex", alignItems:"center", gap:8, flexWrap:"wrap" }}>
                  <div style={{ fontSize:14, fontWeight:600, color:C.text }}>
                    <span style={{ marginRight:6 }}>{cfg.icon}</span>{cfg.label}
                  </div>
                  {!skipped && <TargetBadge cfg={cfg} value={form[c.id]}/>}
                  {activeRules.length > 0 && (
                    <span title={activeRules.map(r => r.name || "Rule").join(", ")} style={{
                      display:"inline-flex", alignItems:"center", gap:3,
                      padding:"2px 7px", borderRadius:5, fontSize:9, fontWeight:800, letterSpacing:.5,
                      background:`${C.purple}22`, border:`1px solid ${C.purple}55`, color:C.purple,
                      fontFamily:"'Space Grotesk', sans-serif",
                    }}>⚡ {activeRules.length} RULE{activeRules.length>1?"S":""}</span>
                  )}
                </div>
                {cfg.message && (
                  <div style={{ fontSize:11, color:C.text, marginTop:4, fontStyle:"italic", opacity:.8 }}>
                    {cfg.message}
                  </div>
                )}
              </div>
              <CriterionField cfg={cfg} value={form[c.id]} onChange={v => f(c.id, v)} isHoliday={form.isHoliday}/>
            </div>
          );
        })}
      </div>

      <button onClick={onSave} className="ht-save" style={{
        gridColumn:"1 / -1",
        width:"100%", padding:18, borderRadius:14, border:"none", cursor:"pointer",
        fontWeight:800, fontSize:16, color:C.white, fontFamily:"inherit", letterSpacing:.5,
        background: saved ? C.green : `linear-gradient(135deg,${C.violet},${C.purple})`,
        transition:"all .3s",
        boxShadow: saved ? `0 0 24px ${C.green}66` : `0 4px 16px ${C.purple}55`,
        transform: saved ? "scale(1.02)" : "scale(1)",
      }}>
        {saved ? "✓ Saved!" : "💾 Save Entry"}
      </button>
    </div>
  );
}

/* ─────────────── HISTORY ─────────────── */
function History({ scored, schema, onPick }) {
  const [filter, setFilter] = useState("all");
  const filtered = useMemo(() => {
    if (filter === "flawless") return scored.filter(e => e.score === 100);
    if (filter === "great") return scored.filter(e => e.score >= 75);
    if (filter === "poor") return scored.filter(e => e.score < 50);
    return scored;
  }, [scored, filter]);

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
      <div style={{ display:"flex", justifyContent:"center" }}>
        <Segmented value={filter} onChange={setFilter} options={[
          { value:"all", label:"All" },
          { value:"flawless", label:"💯 Flawless" },
          { value:"great", label:"⚡ Great" },
          { value:"poor", label:"⚠️ Poor" },
        ]}/>
      </div>
      {filtered.length === 0 ? (
        <Card style={{ alignItems:"center", padding:40 }}>
          <div style={{ fontSize:32 }}>📭</div>
          <div style={{ color:C.muted, fontSize:13 }}>No entries match this filter.</div>
        </Card>
      ) : (
        <div className="ht-history-grid">
          {filtered.map(e => (
            <button key={e.date} onClick={() => onPick(e)} className="ht-history-item" style={{
              padding:14, background:`linear-gradient(135deg,${C.card},${C.card2})`,
              border:`1px solid ${C.border}`, borderLeft:`4px solid ${scoreColor(e.score)}`,
              borderRadius:12, color:C.text, textAlign:"left", cursor:"pointer", fontFamily:"inherit",
              display:"flex", alignItems:"center", gap:12, transition:"all .2s",
            }}>
              <div style={{ flex:1, minWidth:0 }}>
                <div style={{ fontSize:13, fontWeight:700, color:C.text, fontFamily:"'Space Grotesk', sans-serif" }}>{fmtFull(e.date)}</div>
                <div style={{ marginTop:6 }}><CriteriaIcons criteria={e.criteria} schema={schema} size={18}/></div>
              </div>
              <div style={{ textAlign:"right" }}>
                <div style={{ fontSize:24, fontWeight:800, color:scoreColor(e.score), lineHeight:1, fontFamily:"'Space Grotesk', sans-serif" }}>{Math.round(e.score)}%</div>
                <div style={{ fontSize:9, color:scoreColor(e.score), fontWeight:700, letterSpacing:1, marginTop:2 }}>{scoreLabel(e.score)}</div>
              </div>
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

/* ─────────────── SCORE ENGINE EDITOR ─────────────── */
const EMOJI_PALETTE = ["🌅","💪","👟","🥗","🔥","🍗","🥑","💧","🧘","😴","📋","📖","🧹","🚫","🎬","🎯","⚡","🏆","💯","📊","🎨","🎵","💼","🚴","🏃","🧠","☕","🍎","🥦","💊","📝","📞","🛏️","🌙","☀️"];

const RULE_OPS = [
  { value:"is",     label:"=" },
  { value:"not",    label:"≠" },
  { value:"truthy", label:"is ON" },
  { value:"falsy",  label:"is OFF" },
  { value:"gte",    label:"≥" },
  { value:"lte",    label:"≤" },
  { value:"gt",     label:">" },
  { value:"lt",     label:"<" },
];

/* Renders the value input appropriate for the chosen condition field's type */
function ConditionValueInput({ allFields, cond, onChange }) {
  if (cond.op === "truthy" || cond.op === "falsy") {
    return <span style={{ fontSize:11, color:C.muted, padding:"6px 0" }}>(no value)</span>;
  }
  const tgt = allFields.find(f => f.id === cond.field);
  if (tgt?.type === "toggle") {
    return (
      <select value={String(cond.value ?? "true")} onChange={e => onChange(e.target.value === "true")}
        style={{ ...inputSt, width:90 }}>
        <option value="true">true</option>
        <option value="false">false</option>
      </select>
    );
  }
  if (tgt?.type === "time") {
    return <input type="time" value={cond.value || ""} onChange={e => onChange(e.target.value)} style={{ ...inputSt, width:120 }}/>;
  }
  return (
    <input type={tgt?.type === "number" || tgt?.type === "stepper" ? "number" : "text"}
      value={cond.value ?? ""} onChange={e => {
        const raw = e.target.value;
        const n = Number(raw);
        onChange(raw === "" ? "" : (isNaN(n) ? raw : n));
      }}
      style={{ ...inputSt, width:110 }}/>
  );
}

function RulesEditor({ field, allFields, onChange }) {
  const rules = Array.isArray(field.rules) ? field.rules : [];
  const others = allFields.filter(f => f.id !== field.id);

  const setRules = (next) => onChange({ ...field, rules: next });
  const updateRule = (i, patch) => setRules(rules.map((r, j) => j === i ? { ...r, ...patch } : r));
  const updateSet  = (i, patch) => updateRule(i, { set: { ...(rules[i].set || {}), ...patch } });
  const updateCond = (i, ci, patch) => updateRule(i, {
    when: rules[i].when.map((c, k) => k === ci ? { ...c, ...patch } : c)
  });

  const addRule = () => setRules([...rules, {
    id: `r_${Date.now()}_${Math.random().toString(36).slice(2,6)}`,
    name: "New rule",
    enabled: true,
    whenLogic: "all",
    when: others[0] ? [{ field: others[0].id, op: "truthy", value: "" }] : [],
    set: {},
  }]);
  const delRule = (i) => { if (confirm(`Delete rule "${rules[i].name || "rule"}"?`)) setRules(rules.filter((_,j)=>j!==i)); };
  const addCond = (i) => updateRule(i, {
    when: [...(rules[i].when || []), { field: others[0]?.id || "", op: "truthy", value: "" }],
  });
  const delCond = (i, ci) => updateRule(i, { when: rules[i].when.filter((_,k)=>k!==ci) });

  return (
    <div>
      <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:8 }}>
        <Label accent={C.purple}>⚡ Validation Rules ({rules.length})</Label>
        <button onClick={addRule} style={{
          padding:"5px 12px", borderRadius:7, border:`1px solid ${C.purple}55`,
          background:`${C.purple}22`, color:C.purple, fontWeight:700, fontSize:11,
          cursor:"pointer", fontFamily:"inherit",
        }}>＋ Add Rule</button>
      </div>
      {rules.length === 0 && (
        <div style={{
          padding:"10px 12px", borderRadius:8, border:`1px dashed ${C.border2}`,
          background:C.card3, fontSize:11, color:C.muted, lineHeight:1.5,
        }}>
          No rules. Add one to dynamically change <b>target</b>, <b>weight</b>, or <b>message</b> based on other fields. <i>Example: when Workout is ON, set Calories target to 2300.</i>
        </div>
      )}
      <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
        {rules.map((r, i) => (
          <div key={r.id} style={{
            padding:12, borderRadius:10,
            background: r.enabled === false ? C.card3 : `linear-gradient(135deg,${C.card2},${C.card3})`,
            border:`1px solid ${C.purple}33`, opacity: r.enabled === false ? .55 : 1,
          }}>
            {/* Header */}
            <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:8 }}>
              <input value={r.name || ""} placeholder="Rule name"
                onChange={e => updateRule(i, { name: e.target.value })}
                style={{ ...inputSt, fontWeight:700, flex:1 }}/>
              <Toggle value={r.enabled !== false} onChange={v => updateRule(i, { enabled: v })}/>
              <button onClick={() => delRule(i)} style={{
                width:28, height:28, borderRadius:7, border:`1px solid ${C.red}33`,
                background:`${C.red}11`, color:C.red, cursor:"pointer", fontFamily:"inherit",
              }}>🗑</button>
            </div>

            {/* WHEN */}
            <div style={{ marginBottom:8 }}>
              <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:6 }}>
                <Label accent={C.cyan}>WHEN — Match {(r.whenLogic||"all")==="all" ? "ALL" : "ANY"} of:</Label>
                <Segmented value={r.whenLogic || "all"} onChange={v => updateRule(i, { whenLogic: v })} options={[
                  { value:"all", label:"All" },
                  { value:"any", label:"Any" },
                ]}/>
              </div>
              <div style={{ display:"flex", flexDirection:"column", gap:6 }}>
                {(r.when || []).map((cond, ci) => (
                  <div key={ci} style={{ display:"flex", flexWrap:"wrap", gap:6, alignItems:"center" }}>
                    <select value={cond.field || ""} onChange={e => updateCond(i, ci, { field: e.target.value })}
                      style={{ ...inputSt, width:160 }}>
                      <option value="">— field —</option>
                      {others.map(f => <option key={f.id} value={f.id}>{f.icon} {f.label}</option>)}
                    </select>
                    <select value={cond.op} onChange={e => updateCond(i, ci, { op: e.target.value })}
                      style={{ ...inputSt, width:90 }}>
                      {RULE_OPS.map(o => <option key={o.value} value={o.value}>{o.label}</option>)}
                    </select>
                    <ConditionValueInput allFields={allFields} cond={cond} onChange={v => updateCond(i, ci, { value: v })}/>
                    <button onClick={() => delCond(i, ci)} title="Remove condition" style={{
                      width:26, height:26, borderRadius:6, border:`1px solid ${C.border2}`,
                      background:C.card3, color:C.muted, cursor:"pointer", fontFamily:"inherit",
                    }}>×</button>
                  </div>
                ))}
                <button onClick={() => addCond(i)} style={{
                  alignSelf:"flex-start", padding:"4px 10px", borderRadius:6,
                  border:`1px dashed ${C.border2}`, background:"transparent", color:C.muted,
                  fontSize:11, cursor:"pointer", fontFamily:"inherit", fontWeight:700,
                }}>＋ Add Condition</button>
              </div>
            </div>

            {/* THEN */}
            <div>
              <Label accent={C.green}>THEN — Override this field with:</Label>
              <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(140px, 1fr))", gap:8 }}>
                <div>
                  <div style={{ fontSize:10, color:C.muted, fontWeight:700, marginBottom:3 }}>Target</div>
                  <input type={field.type === "time" ? "time" : "number"}
                    placeholder="(no override)"
                    value={field.type === "time"
                      ? (r.set?.threshold !== undefined ? minToTime(r.set.threshold) : "")
                      : (r.set?.threshold ?? "")}
                    onChange={e => {
                      const v = e.target.value;
                      if (v === "") { const s = { ...r.set }; delete s.threshold; updateRule(i, { set: s }); }
                      else updateSet(i, { threshold: field.type === "time" ? timeToMin(v) : Number(v) });
                    }}
                    style={inputSt}/>
                </div>
                <div>
                  <div style={{ fontSize:10, color:C.muted, fontWeight:700, marginBottom:3 }}>Holiday Target</div>
                  <input type="number" placeholder="(no override)" value={r.set?.holidayThreshold ?? ""}
                    onChange={e => {
                      const v = e.target.value;
                      if (v === "") { const s = { ...r.set }; delete s.holidayThreshold; updateRule(i, { set: s }); }
                      else updateSet(i, { holidayThreshold: Number(v) });
                    }}
                    style={inputSt}/>
                </div>
                <div>
                  <div style={{ fontSize:10, color:C.muted, fontWeight:700, marginBottom:3 }}>Weight</div>
                  <input type="number" step="0.5" placeholder="(no override)" value={r.set?.weight ?? ""}
                    onChange={e => {
                      const v = e.target.value;
                      if (v === "") { const s = { ...r.set }; delete s.weight; updateRule(i, { set: s }); }
                      else updateSet(i, { weight: Number(v) });
                    }}
                    style={inputSt}/>
                </div>
                <div>
                  <div style={{ fontSize:10, color:C.muted, fontWeight:700, marginBottom:3 }}>Op</div>
                  <select value={r.set?.op ?? ""} onChange={e => {
                    const v = e.target.value;
                    if (v === "") { const s = { ...r.set }; delete s.op; updateRule(i, { set: s }); }
                    else updateSet(i, { op: v });
                  }} style={inputSt}>
                    <option value="">(no override)</option>
                    <option value="gte">≥</option>
                    <option value="lte">≤</option>
                  </select>
                </div>
                <div style={{ gridColumn:"1 / -1" }}>
                  <div style={{ fontSize:10, color:C.muted, fontWeight:700, marginBottom:3 }}>Custom Message (shown in Log form)</div>
                  <input placeholder="(no override)" value={r.set?.message ?? ""}
                    onChange={e => {
                      const v = e.target.value;
                      if (v === "") { const s = { ...r.set }; delete s.message; updateRule(i, { set: s }); }
                      else updateSet(i, { message: v });
                    }}
                    style={inputSt}/>
                </div>
                <div>
                  <div style={{ fontSize:10, color:C.muted, fontWeight:700, marginBottom:3 }}>Target Label</div>
                  <input placeholder="Goal / Limit / Max…" value={r.set?.targetLabel ?? ""}
                    onChange={e => {
                      const v = e.target.value;
                      if (v === "") { const s = { ...r.set }; delete s.targetLabel; updateRule(i, { set: s }); }
                      else updateSet(i, { targetLabel: v });
                    }}
                    style={inputSt}/>
                </div>
                <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                  <span style={{ fontSize:11, color:C.muted, fontWeight:700 }}>Skip field</span>
                  <Toggle value={!!r.set?.skip} onChange={v => updateSet(i, { skip: v })}/>
                </div>
              </div>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function FieldEditor({ field, allFields, onChange, onDelete, onMove }) {
  const [expanded, setExpanded] = useState(false);
  const update = (k, v) => onChange({ ...field, [k]: v });

  return (
    <Card style={{ padding:14, gap:10, opacity: field.enabled ? 1 : .55 }}>
      <div style={{ display:"flex", alignItems:"center", gap:10 }}>
        <button onClick={() => setExpanded(!expanded)} className="ht-icon-btn" style={{
          width:28, height:28, borderRadius:7, border:`1px solid ${C.border2}`,
          background:C.card3, color:C.muted, cursor:"pointer", fontSize:14, fontFamily:"inherit",
        }}>{expanded ? "▾" : "▸"}</button>
        <div style={{ fontSize:22 }}>{field.icon}</div>
        <input value={field.label} onChange={e => update("label", e.target.value)}
          style={{ ...inputSt, fontWeight:700, flex:1 }}/>
        <Toggle value={field.enabled} onChange={v => update("enabled", v)}/>
        <button onClick={() => onMove(-1)} className="ht-icon-btn" title="Move up" style={{ width:28, height:28, borderRadius:7, border:`1px solid ${C.border2}`, background:C.card3, color:C.muted, cursor:"pointer", fontFamily:"inherit" }}>↑</button>
        <button onClick={() => onMove(1)} className="ht-icon-btn" title="Move down" style={{ width:28, height:28, borderRadius:7, border:`1px solid ${C.border2}`, background:C.card3, color:C.muted, cursor:"pointer", fontFamily:"inherit" }}>↓</button>
        <button onClick={onDelete} className="ht-icon-btn" title="Delete" style={{ width:28, height:28, borderRadius:7, border:`1px solid ${C.red}33`, background:`${C.red}11`, color:C.red, cursor:"pointer", fontFamily:"inherit" }}>🗑</button>
      </div>

      {expanded && (
        <>
          <Div/>
          {/* Icon picker */}
          <div>
            <Label>Icon</Label>
            <div style={{ display:"flex", flexWrap:"wrap", gap:4 }}>
              {EMOJI_PALETTE.map(e => (
                <button key={e} onClick={() => update("icon", e)} style={{
                  width:32, height:32, borderRadius:7, fontSize:18, cursor:"pointer",
                  border: field.icon === e ? `2px solid ${C.purple}` : `1px solid ${C.border2}`,
                  background: field.icon === e ? `${C.purple}22` : C.card3,
                  fontFamily:"inherit",
                }}>{e}</button>
              ))}
            </div>
          </div>

          {/* Type */}
          <div>
            <Label>Type</Label>
            <Segmented value={field.type} onChange={v => update("type", v)} options={[
              { value:"toggle", label:"Toggle" },
              { value:"number", label:"Number" },
              { value:"stepper", label:"Stepper" },
              { value:"time", label:"Time" },
            ]}/>
          </div>

          {/* Operator */}
          {field.type !== "toggle" && (
            <div>
              <Label>Met When</Label>
              <Segmented value={field.op || "gte"} onChange={v => update("op", v)} options={[
                { value:"gte", label:"≥ Threshold" },
                { value:"lte", label:"≤ Threshold" },
              ]}/>
            </div>
          )}

          {/* Threshold */}
          {field.type === "time" ? (
            <div>
              <Label>Threshold (Time)</Label>
              <input type="time" value={minToTime(field.threshold || 0)}
                onChange={e => update("threshold", timeToMin(e.target.value))}
                style={{ ...inputSt, width:160 }}/>
            </div>
          ) : field.type !== "toggle" ? (
            <div style={{ display:"flex", gap:10, flexWrap:"wrap" }}>
              <div style={{ flex:"1 1 140px" }}>
                <Label>Threshold</Label>
                <input type="number" value={field.threshold ?? 0}
                  onChange={e => update("threshold", Number(e.target.value))}
                  style={inputSt}/>
              </div>
              <div style={{ flex:"1 1 100px" }}>
                <Label>Unit</Label>
                <input value={field.unit || ""} placeholder="g, kcal..."
                  onChange={e => update("unit", e.target.value)}
                  style={inputSt}/>
              </div>
              {field.type === "stepper" && (
                <div style={{ flex:"1 1 100px" }}>
                  <Label>Stepper Max</Label>
                  <input type="number" value={field.max || 8}
                    onChange={e => update("max", Number(e.target.value))}
                    style={inputSt}/>
                </div>
              )}
              <div style={{ flex:"1 1 140px" }}>
                <Label>Holiday Threshold</Label>
                <input type="number" value={field.holidayThreshold ?? ""}
                  placeholder="(optional)"
                  onChange={e => update("holidayThreshold", e.target.value === "" ? undefined : Number(e.target.value))}
                  style={inputSt}/>
              </div>
            </div>
          ) : (
            <div style={{ display:"flex", alignItems:"center", gap:10 }}>
              <Label>Inverted (met when OFF)</Label>
              <Toggle value={!!field.inverted} onChange={v => update("inverted", v)}/>
            </div>
          )}

          {/* Weight + skipHoliday + trackingSince */}
          <div style={{ display:"flex", gap:14, flexWrap:"wrap", alignItems:"center" }}>
            <div style={{ flex:"1 1 260px" }}>
              {(() => {
                const enabledFields = allFields.filter(f => f.enabled);
                const totalW = enabledFields.reduce((s, f) => s + (Number(f.weight) || 1), 0);
                const myW = Number(field.weight) || 1;
                const pct = totalW > 0 ? (myW / totalW) * 100 : 0;
                const setW = (n) => update("weight", Math.max(0.1, Math.round(n * 10) / 10));
                return (
                  <>
                    <div style={{ display:"flex", justifyContent:"space-between", alignItems:"baseline", marginBottom:6 }}>
                      <Label>Weight (relative importance)</Label>
                      <span style={{ fontSize:11, color:C.muted, fontWeight:700 }}>
                        ×{myW} · {pct.toFixed(1)}% of score
                      </span>
                    </div>
                    <div style={{ display:"flex", gap:8, alignItems:"center" }}>
                      <input type="range" min="0.5" max="5" step="0.5" value={myW}
                        onChange={e => setW(Number(e.target.value))}
                        style={{ flex:1, accentColor:C.purple }}/>
                      <input type="number" min="0.1" step="0.5" value={myW}
                        onChange={e => setW(Number(e.target.value) || 1)}
                        style={{ ...inputSt, width:80, padding:"6px 8px" }}/>
                    </div>
                    <div style={{ display:"flex", gap:4, marginTop:6, flexWrap:"wrap" }}>
                      {[
                        { v:0.5, lbl:"Low" },
                        { v:1,   lbl:"Normal" },
                        { v:2,   lbl:"High" },
                        { v:3,   lbl:"Critical" },
                      ].map(p => (
                        <button key={p.v} onClick={() => setW(p.v)}
                          style={{
                            padding:"3px 8px", borderRadius:6, fontSize:10, fontWeight:700,
                            cursor:"pointer",
                            background: myW === p.v ? C.purple : "transparent",
                            color: myW === p.v ? "#fff" : C.muted,
                            border:`1px solid ${myW === p.v ? C.purple : C.border2}`,
                          }}>
                          ×{p.v} {p.lbl}
                        </button>
                      ))}
                    </div>
                  </>
                );
              })()}
            </div>
            <div style={{ flex:"1 1 160px" }}>
              <Label>Tracking Since (excludes earlier days from score)</Label>
              <input type="date" value={field.trackingSince || ""}
                onChange={e => update("trackingSince", e.target.value || undefined)}
                style={inputSt}/>
            </div>
            <div style={{ display:"flex", alignItems:"center", gap:10 }}>
              <span style={{ fontSize:12, color:C.muted, fontWeight:700 }}>Skip on Holiday</span>
              <Toggle value={!!field.skipHoliday} onChange={v => update("skipHoliday", v)}/>
            </div>
          </div>

          {/* Display & scoring style */}
          <Div/>
          <div style={{ display:"flex", gap:10, flexWrap:"wrap" }}>
            <div style={{ flex:"2 1 240px" }}>
              <Label>Custom Message (helper text in Log form)</Label>
              <input value={field.message || ""} placeholder="e.g. Stay under target on rest days"
                onChange={e => update("message", e.target.value)}
                style={inputSt}/>
            </div>
            <div style={{ flex:"1 1 140px" }}>
              <Label>Target Label</Label>
              <input value={field.targetLabel || ""}
                placeholder={field.op === "lte" ? "Limit / Max / Cap" : "Goal / Min / Target"}
                onChange={e => update("targetLabel", e.target.value)}
                style={inputSt}/>
            </div>
          </div>
          {(field.type === "number" || field.type === "stepper") && (
            <div>
              <Label>Score Mode</Label>
              <Segmented value={field.scoreMode || "binary"} onChange={v => update("scoreMode", v)} options={[
                { value:"binary", label:"All or Nothing" },
                { value:"graded", label:"Partial Credit" },
              ]}/>
              <div style={{ fontSize:10, color:C.muted, marginTop:4 }}>
                {field.scoreMode === "graded"
                  ? `Earn proportional credit (e.g. ${field.op === "lte" ? "going over reduces score" : "value/target = credit"})`
                  : "Full credit only when target is reached."}
              </div>
            </div>
          )}

          {/* Validation Rules */}
          <Div/>
          <RulesEditor field={field} allFields={allFields} onChange={onChange}/>
        </>
      )}
    </Card>
  );
}

/* ─────────────── IMPORT PANEL (xlsx upload) ─────────────── */

// Excel serial → ISO date string (YYYY-MM-DD). Excel epoch quirk: day 60 = bogus 1900-02-29.
function excelSerialToDate(n) {
  if (typeof n !== "number" || !isFinite(n)) return null;
  const ms = Math.round((n - 25569) * 86400 * 1000); // 25569 = days from 1970-01-01 to 1900-01-01 epoch
  const d = new Date(ms);
  if (isNaN(d.getTime())) return null;
  return d.toISOString().split("T")[0];
}
// Excel/Sheets time → "HH:MM". Accepts:
//   - pure day-fraction (0 ≤ n < 1) e.g. 0.2548 → "06:07"
//   - serial date+time            e.g. 46139.2548 → "06:07" (drops the date)
//   - JS Date object              → uses local hours/minutes
//   - "HH:MM" / "H:MM" string     → returned padded
function fractionToTime(n) {
  if (n instanceof Date && !isNaN(n)) {
    return `${String(n.getHours()).padStart(2,"0")}:${String(n.getMinutes()).padStart(2,"0")}`;
  }
  if (typeof n === "string") {
    const m = n.match(/^(\d{1,2}):(\d{2})/);
    if (m) return `${String(Math.min(23, Number(m[1]))).padStart(2,"0")}:${m[2]}`;
    const num = Number(n);
    if (isFinite(num)) return fractionToTime(num);
    return "";
  }
  if (typeof n !== "number" || !isFinite(n)) return "";
  // Strip date portion if combined serial.
  const frac = n - Math.floor(n);
  const total = Math.round(frac * 24 * 60);
  const h = Math.floor(total / 60) % 24, m = total % 60;
  return `${String(h).padStart(2,"0")}:${String(m).padStart(2,"0")}`;
}

const COLUMN_HINTS = {
  date:        [/^date$/i],
  wake:        [/wake/i],
  isHoliday:   [/holiday|weekend/i],
  workout:     [/workout/i],
  steps:       [/^steps?$/i],
  calories:    [/calories?\s*actual|^calories?$/i],
  caloriesTarget:[/calories?\s*target/i],
  workBlocks:  [/work\s*blocks?/i],
  cleaning:    [/cleaning/i],
  junkFood:    [/junk/i],
  movieMinutes:[/movie|screen/i],
  reading:     [/reading|book/i],
  sleepHours:  [/sleep/i],
  meditation:  [/meditat/i],
  water:       [/water/i],
  protein:     [/protein/i],
  fats:        [/^fats?$/i],
  score:       [/^.*score.*$/i],
};

function autoMap(headers) {
  // returns { headerIndex: field | "_skip" }
  const map = {};
  headers.forEach((h, i) => {
    const txt = String(h || "").trim();
    if (!txt) { map[i] = "_skip"; return; }
    let matched = "_skip";
    for (const [field, patterns] of Object.entries(COLUMN_HINTS)) {
      if (patterns.some(p => p.test(txt))) { matched = field; break; }
    }
    map[i] = matched;
  });
  return map;
}

const KNOWN_FIELDS = [
  { value:"_skip", label:"— Skip —" },
  { value:"date", label:"date" },
  { value:"isHoliday", label:"isHoliday" },
  { value:"wake", label:"wake (HH:MM)" },
  { value:"sleepHours", label:"sleepHours" },
  { value:"meditation", label:"meditation" },
  { value:"workout", label:"workout" },
  { value:"steps", label:"steps" },
  { value:"water", label:"water" },
  { value:"calories", label:"calories" },
  { value:"protein", label:"protein" },
  { value:"fats", label:"fats" },
  { value:"junkFood", label:"junkFood" },
  { value:"workBlocks", label:"workBlocks" },
  { value:"reading", label:"reading" },
  { value:"cleaning", label:"cleaning" },
  { value:"movieMinutes", label:"movieMinutes" },
];

function coerceValue(field, raw, schemaField) {
  if (raw === null || raw === undefined || raw === "") return undefined;
  if (field === "date") {
    if (typeof raw === "number") return excelSerialToDate(raw);
    if (raw instanceof Date) return raw.toISOString().split("T")[0];
    const d = new Date(raw);
    return isNaN(d) ? String(raw) : d.toISOString().split("T")[0];
  }
  if (field === "isHoliday") return coerceBool(raw);
  // Prefer schema-declared type so user-added custom fields round-trip too.
  const type = schemaField?.type;
  if (type === "time" || field === "wake") return fractionToTime(raw);
  if (type === "toggle") return coerceBool(raw);
  if (type === "number" || type === "stepper") {
    const n = Number(raw);
    return isNaN(n) ? undefined : n;
  }
  // Legacy/no-schema fallback for known built-in ids.
  switch (field) {
    case "workout": case "meditation": case "junkFood":
    case "reading": case "cleaning":
      return coerceBool(raw);
    default: {
      const n = Number(raw);
      return isNaN(n) ? raw : n;
    }
  }
}

function ImportPanel({ schema, entries, onImport }) {
  const [rows, setRows] = useState(null);    // [{__row, ...rawCells}]
  const [headers, setHeaders] = useState([]);
  const [map, setMap] = useState({});
  const [mode, setMode] = useState("merge"); // merge | overwrite | skip
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState(null);
  const [filename, setFilename] = useState("");

  const onFile = async (file) => {
    if (!file) return;
    setBusy(true); setErr(null); setFilename(file.name);
    try {
      const XLSX = await import("xlsx");
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type:"array", cellDates:false });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const aoa = XLSX.utils.sheet_to_json(ws, { header:1, defval:null, blankrows:false });
      if (!aoa.length) throw new Error("Sheet is empty");
      const hdr = aoa[0].map(h => h == null ? "" : String(h));
      const data = aoa.slice(1).filter(r => r.some(c => c !== null && c !== ""));
      setHeaders(hdr);
      setRows(data);
      setMap(autoMap(hdr));
    } catch (e) {
      setErr(e.message || String(e));
      setRows(null);
    } finally { setBusy(false); }
  };

  const previewEntries = useMemo(() => {
    if (!rows) return [];
    return rows.map(r => {
      const obj = {};
      headers.forEach((_, i) => {
        const f = map[i];
        if (!f || f === "_skip") return;
        const sf = schema.find(s => s.id === f);
        const v = coerceValue(f, r[i], sf);
        if (v !== undefined) obj[f] = v;
      });
      return obj;
    }).filter(o => o.date);
  }, [rows, headers, map]);

  const existingDates = useMemo(() => new Set(entries.map(e => e.date)), [entries]);
  const newCount  = previewEntries.filter(e => !existingDates.has(e.date)).length;
  const dupCount  = previewEntries.length - newCount;

  const doImport = async () => {
    if (!previewEntries.length) return;
    let merged;
    if (mode === "overwrite") {
      const map2 = new Map(entries.map(e => [e.date, e]));
      previewEntries.forEach(e => map2.set(e.date, { ...map2.get(e.date), ...e }));
      merged = [...map2.values()];
    } else if (mode === "skip") {
      const have = new Set(entries.map(e => e.date));
      merged = [...entries, ...previewEntries.filter(e => !have.has(e.date))];
    } else { // merge: combine fields, new rows added
      const map2 = new Map(entries.map(e => [e.date, { ...e }]));
      previewEntries.forEach(e => {
        const cur = map2.get(e.date) || {};
        map2.set(e.date, { ...cur, ...e });
      });
      merged = [...map2.values()];
    }
    merged.sort((a,b) => a.date.localeCompare(b.date));
    setBusy(true); setErr(null);
    try {
      await onImport(merged);
      setRows(null); setHeaders([]); setMap({}); setFilename("");
    } catch (e) {
      setErr(e.message || String(e));
    } finally { setBusy(false); }
  };

  return (
    <Card>
      <Label accent={C.cyan}>📥 Import from Excel / CSV</Label>
      <div style={{ fontSize:13, color:C.text, lineHeight:1.6 }}>
        Upload an <code style={{ background:C.card3, padding:"2px 6px", borderRadius:4, color:C.cyan }}>.xlsx</code> or <code style={{ background:C.card3, padding:"2px 6px", borderRadius:4, color:C.cyan }}>.csv</code> file. Columns are auto-mapped to habit fields — adjust below, then import. Scores will be re-computed using your current Engine schema.
      </div>

      <label className="ht-chip" style={{
        display:"inline-flex", alignItems:"center", gap:8,
        padding:"10px 16px", borderRadius:10, border:`1.5px dashed ${C.purple}66`,
        background:`${C.purple}11`, color:C.purple, fontWeight:700, fontSize:13,
        cursor: busy ? "wait" : "pointer", width:"fit-content",
      }}>
        📂 {filename || "Choose file…"}
        <input type="file" accept=".xlsx,.xls,.csv" disabled={busy}
          onChange={e => onFile(e.target.files?.[0])}
          style={{ display:"none" }}/>
      </label>

      {err && (
        <div style={{ padding:12, borderRadius:8, background:`${C.red}15`, border:`1px solid ${C.red}55`, color:C.red, fontSize:13 }}>
          ⚠️ {err}
        </div>
      )}

      {rows && (
        <>
          <Div/>
          <Label>Column Mapping</Label>
          <div style={{ display:"grid", gridTemplateColumns:"1fr auto 1fr", gap:"8px 12px", alignItems:"center" }}>
            {headers.map((h, i) => (
              <div key={i} style={{ display:"contents" }}>
                <div style={{ fontSize:12, color:C.text, fontWeight:600, padding:"6px 10px", background:C.card3, borderRadius:6, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }} title={h}>
                  <span style={{ color:C.muted, fontFamily:"monospace", marginRight:6 }}>{String.fromCharCode(65+i)}</span>{h || <em style={{color:C.muted}}>(empty)</em>}
                </div>
                <div style={{ color:C.muted, fontSize:14 }}>→</div>
                <select value={map[i] || "_skip"}
                  onChange={e => setMap({ ...map, [i]: e.target.value })}
                  style={{ ...inputSt, padding:"6px 10px", fontSize:12 }}>
                  {KNOWN_FIELDS.map(f => <option key={f.value} value={f.value}>{f.label}</option>)}
                </select>
              </div>
            ))}
          </div>

          <Div/>
          <Label>Preview ({previewEntries.length} valid rows)</Label>
          <div style={{ maxHeight:200, overflow:"auto", border:`1px solid ${C.border}`, borderRadius:8 }}>
            <table style={{ width:"100%", fontSize:11, borderCollapse:"collapse" }}>
              <thead style={{ position:"sticky", top:0, background:C.card3 }}>
                <tr>
                  {Object.keys(previewEntries[0] || {}).map(k => (
                    <th key={k} style={{ padding:"6px 8px", textAlign:"left", color:C.muted, fontWeight:700, borderBottom:`1px solid ${C.border}` }}>{k}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {previewEntries.slice(0, 8).map((e, i) => (
                  <tr key={i} style={{ background: i%2 ? C.card2 : "transparent" }}>
                    {Object.keys(previewEntries[0] || {}).map(k => (
                      <td key={k} style={{ padding:"5px 8px", color:C.text, borderBottom:`1px solid ${C.border}`, whiteSpace:"nowrap" }}>{String(e[k] ?? "")}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div style={{ fontSize:11, color:C.muted }}>
            {newCount} new · {dupCount} overlap with existing dates
          </div>

          <Div/>
          <Label>Conflict Mode</Label>
          <Segmented value={mode} onChange={setMode} options={[
            { value:"merge",     label:"Merge fields" },
            { value:"overwrite", label:"Overwrite day" },
            { value:"skip",      label:"Skip existing" },
          ]}/>
          <div style={{ fontSize:11, color:C.muted, marginTop:-6 }}>
            {mode === "merge" && "Fill in missing fields per day; existing values kept."}
            {mode === "overwrite" && "Imported fields replace existing values for the same date."}
            {mode === "skip" && "Days that already exist are left untouched."}
          </div>

          <div style={{ display:"flex", gap:10, flexWrap:"wrap" }}>
            <button onClick={doImport} disabled={busy || !previewEntries.length} className="ht-cta" style={{
              padding:"10px 18px", borderRadius:10, border:"none",
              background: busy ? C.muted : `linear-gradient(135deg,${C.violet},${C.purple})`,
              color:C.white, fontWeight:700, fontSize:13,
              cursor: busy ? "wait" : "pointer", fontFamily:"inherit",
              boxShadow:`0 4px 12px ${C.purple}55`,
            }}>{busy ? "Importing…" : `📥 Import ${previewEntries.length} Rows`}</button>
            <button onClick={() => { setRows(null); setHeaders([]); setMap({}); setFilename(""); }} className="ht-chip" style={{
              padding:"10px 18px", borderRadius:10, border:`1px solid ${C.border2}`,
              background:C.card3, color:C.muted, fontWeight:700, fontSize:13, cursor:"pointer", fontFamily:"inherit",
            }}>Cancel</button>
          </div>
        </>
      )}
    </Card>
  );
}

function Engine({ schema, setSchema, entries, onImport }) {
  const totalW = schema.filter(c=>c.enabled).reduce((s,c)=>s+(Number(c.weight)||1), 0);

  const update = (i, f) => {
    const next = [...schema]; next[i] = f; setSchema(next);
  };
  const del = (i) => {
    if (!confirm(`Delete "${schema[i].label}"?`)) return;
    setSchema(schema.filter((_,j) => j !== i));
  };
  const move = (i, dir) => {
    const j = i + dir;
    if (j < 0 || j >= schema.length) return;
    const next = [...schema]; [next[i], next[j]] = [next[j], next[i]]; setSchema(next);
  };
  const add = () => {
    const id = `custom_${Date.now()}`;
    // Stamp the start date so days logged BEFORE this habit existed are not
    // retroactively scored against it.
    setSchema([...schema, { id, icon:"⭐", label:"New Habit", type:"toggle", weight:1, enabled:true, trackingSince: getToday() }]);
  };
  const reset = () => {
    if (!confirm("Reset all habits to defaults? Your entries are preserved.")) return;
    setSchema(DEFAULT_SCHEMA);
  };

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
      <Card>
        <Label accent={C.purple}>⚙️ Score Engine Editor</Label>
        <div style={{ fontSize:13, color:C.text, lineHeight:1.6 }}>
          Add, edit, reorder, and weight habits. Score = <b>met weight</b> ÷ <b>total weight</b> × 100.
        </div>
        <div style={{ display:"flex", gap:10, flexWrap:"wrap" }}>
          <button onClick={add} className="ht-cta" style={{
            padding:"8px 16px", borderRadius:10, border:"none",
            background:`linear-gradient(135deg,${C.violet},${C.purple})`, color:C.white,
            fontWeight:700, fontSize:13, cursor:"pointer", fontFamily:"inherit",
            boxShadow:`0 4px 12px ${C.purple}55`,
          }}>＋ Add Habit</button>
          <button onClick={reset} className="ht-chip" style={{
            padding:"8px 16px", borderRadius:10, border:`1px solid ${C.border2}`,
            background:C.card3, color:C.muted, fontWeight:700, fontSize:13, cursor:"pointer", fontFamily:"inherit",
          }}>↺ Reset to Defaults</button>
          <div style={{ marginLeft:"auto", padding:"8px 14px", borderRadius:10, background:`${C.cyan}15`, border:`1px solid ${C.cyan}33`, color:C.cyan, fontSize:12, fontWeight:700 }}>
            {schema.filter(c=>c.enabled).length} active · total weight {totalW}
          </div>
        </div>
      </Card>

      <ImportPanel schema={schema} entries={entries} onImport={onImport}/>

      <div className="ht-engine-grid">
        {schema.map((f, i) => (
          <FieldEditor
            key={f.id}
            field={f}
            allFields={schema}
            onChange={(nf) => update(i, nf)}
            onDelete={() => del(i)}
            onMove={(dir) => move(i, dir)}
          />
        ))}
      </div>
    </div>
  );
}

/* ─────────────── SETTINGS / THEME ───────────────
 * Lets the user customise core palette colors. Changes mutate the shared `C`
 * object live (bumping `themeVersion` re-renders the tree) and persist to
 * the Sheet under key="theme". On boot the theme is merged before first paint.
 */
const THEME_KEYS = [
  { k:"purple", label:"Primary",    hint:"Buttons, brand, headings" },
  { k:"violet", label:"Primary alt", hint:"Gradient pair w/ Primary" },
  { k:"cyan",   label:"Accent",     hint:"Today, links, info" },
  { k:"green",  label:"Success",    hint:"Streaks, met habits" },
  { k:"amber",  label:"Warning",    hint:"Streak fire, limits" },
  { k:"red",    label:"Danger",     hint:"Errors, missed habits" },
  { k:"pink",   label:"Highlight",  hint:"30-day avg etc" },
  { k:"bg",     label:"Background", hint:"Page background" },
  { k:"card",   label:"Surface 1",  hint:"Cards (top of gradient)" },
  { k:"card2",  label:"Surface 2",  hint:"Cards (bottom of gradient)" },
  { k:"card3",  label:"Surface 3",  hint:"Inputs, segmented bg" },
  { k:"text",   label:"Text",       hint:"Primary text color" },
  { k:"muted",  label:"Muted text", hint:"Helper / secondary text" },
  { k:"border", label:"Border",     hint:"Card borders" },
  { k:"border2",label:"Border alt", hint:"Inputs, dividers" },
];

// Each preset is a *full* theme override (surfaces + accents) so the look
// changes dramatically — not just the brand color. `bg`/`card*`/`text`/`border*`
// shift too, which is what makes them feel like different apps.
const PRESETS = {
  // Default — deep indigo dark with violet/cyan accents
  Midnight: {
    bg:"#08080f", card:"#0f0f1a", card2:"#141422", card3:"#191930",
    border:"#1e1e30", border2:"#25253a", text:"#e2e8f0", muted:"#64748b",
    purple:"#8b5cf6", violet:"#7c3aed", cyan:"#06b6d4", green:"#10b981",
    amber:"#f59e0b", red:"#f43f5e", pink:"#ec4899",
  },
  // Hot tropical — coral / magenta on warm dark plum
  Sunset: {
    bg:"#1a0a14", card:"#241019", card2:"#2e1320", card3:"#3a1828",
    border:"#3d1f2e", border2:"#4d2738", text:"#fde4d0", muted:"#a78597",
    purple:"#fb7185", violet:"#e11d48", cyan:"#fbbf24", green:"#a3e635",
    amber:"#fb923c", red:"#dc2626", pink:"#f472b6",
  },
  // Cool teal/blue on near-black ocean
  Ocean: {
    bg:"#020617", card:"#0c1a2e", card2:"#0f2540", card3:"#143155",
    border:"#1e3a5f", border2:"#2c4a73", text:"#cbd5e1", muted:"#64758b",
    purple:"#0ea5e9", violet:"#0369a1", cyan:"#22d3ee", green:"#14b8a6",
    amber:"#eab308", red:"#f87171", pink:"#67e8f9",
  },
  // Vibrant green / lime on deep forest
  Forest: {
    bg:"#06140d", card:"#0c1f15", card2:"#11291c", card3:"#163525",
    border:"#1d3d2c", border2:"#274d39", text:"#d4e8d8", muted:"#7a9786",
    purple:"#22c55e", violet:"#16a34a", cyan:"#a3e635", green:"#4ade80",
    amber:"#eab308", red:"#dc2626", pink:"#fb923c",
  },
  // Brutalist greyscale dark — only red survives as accent
  Mono: {
    bg:"#0a0a0a", card:"#141414", card2:"#1a1a1a", card3:"#222222",
    border:"#2a2a2a", border2:"#3a3a3a", text:"#f5f5f5", muted:"#737373",
    purple:"#fafafa", violet:"#d4d4d4", cyan:"#e5e5e5", green:"#a3a3a3",
    amber:"#d4d4d4", red:"#ef4444", pink:"#a3a3a3",
  },
  // Cyberpunk neon — magenta + cyan on jet black
  Neon: {
    bg:"#000000", card:"#0d0014", card2:"#150022", card3:"#1d0030",
    border:"#330055", border2:"#4d0080", text:"#f0aaff", muted:"#9966cc",
    purple:"#ff00ff", violet:"#cc00ff", cyan:"#00ffff", green:"#00ff88",
    amber:"#ffea00", red:"#ff0055", pink:"#ff66cc",
  },
  // Light mode — high contrast paper white
  Daylight: {
    bg:"#f8fafc", card:"#ffffff", card2:"#f1f5f9", card3:"#e2e8f0",
    border:"#cbd5e1", border2:"#94a3b8", text:"#0f172a", muted:"#475569",
    purple:"#7c3aed", violet:"#6d28d9", cyan:"#0891b2", green:"#059669",
    amber:"#d97706", red:"#dc2626", pink:"#db2777",
  },
  // Sepia / warm cream
  Parchment: {
    bg:"#f5efe1", card:"#fbf6e9", card2:"#efe6cf", card3:"#e3d6b3",
    border:"#c9b78a", border2:"#a8946b", text:"#3a2a14", muted:"#7a6a4f",
    purple:"#9a3412", violet:"#7c2d12", cyan:"#0c4a6e", green:"#365314",
    amber:"#a16207", red:"#7f1d1d", pink:"#9d174d",
  },
  // Cherry blossom — soft pink + sakura cream on midnight purple
  Sakura: {
    bg:"#1a0d1f", card:"#241229", card2:"#2e1834", card3:"#3a1f42",
    border:"#4a2754", border2:"#5e3268", text:"#ffd7e8", muted:"#b08aa8",
    purple:"#f9a8d4", violet:"#ec4899", cyan:"#a5f3fc", green:"#bef264",
    amber:"#fde68a", red:"#fb7185", pink:"#fbcfe8",
  },
  // Volcano — molten orange/red on charcoal
  Volcano: {
    bg:"#0c0908", card:"#161210", card2:"#1f1a17", card3:"#2a221d",
    border:"#3a2e26", border2:"#4d3d33", text:"#f4ede4", muted:"#a89484",
    purple:"#fb923c", violet:"#ea580c", cyan:"#fde047", green:"#84cc16",
    amber:"#f59e0b", red:"#b91c1c", pink:"#fb7185",
  },
};

const DEFAULT_THEME = {
  bg:"#08080f", card:"#0f0f1a", card2:"#141422", card3:"#191930",
  border:"#1e1e30", border2:"#25253a", text:"#e2e8f0", muted:"#64748b",
  purple:"#8b5cf6", violet:"#7c3aed", cyan:"#06b6d4", green:"#10b981",
  amber:"#f59e0b", red:"#f43f5e", pink:"#ec4899",
};

function ColorPicker({ value, onChange }) {
  return (
    <div style={{ display:"flex", alignItems:"center", gap:8 }}>
      <input type="color" value={value} onChange={e => onChange(e.target.value)}
        style={{ width:42, height:34, border:`1px solid ${C.border2}`, borderRadius:8, background:"transparent", cursor:"pointer", padding:2 }}/>
      <input value={value} onChange={e => onChange(e.target.value)}
        style={{ ...inputSt, width:110, fontFamily:"'Space Grotesk', monospace", fontSize:12 }}/>
    </div>
  );
}

function Settings({ theme, onChange, onReset }) {
  const set = (k, v) => onChange({ ...theme, [k]: v });
  return (
    <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
      <Card>
        <Label accent={C.purple}>🎨 Theme</Label>
        <div style={{ fontSize:13, color:C.text, lineHeight:1.6 }}>
          Customise the palette used everywhere. Changes are saved to the Sheet so they sync across devices.
        </div>
        <div>
          <Label>Quick Presets</Label>
          <div style={{ display:"flex", flexWrap:"wrap", gap:8 }}>
            {Object.entries(PRESETS).map(([name, p]) => (
              <button key={name} onClick={() => onChange(p)} className="ht-chip" style={{
                padding:"6px 12px", borderRadius:8, border:`1px solid ${C.border2}`,
                background:C.card3, color:C.text, fontWeight:700, fontSize:12, cursor:"pointer", fontFamily:"inherit",
                display:"flex", alignItems:"center", gap:6,
              }}>
                {/* Mini preview: surface + 4 accent dots */}
                <span style={{
                  width:14, height:14, borderRadius:4, background:p.bg,
                  border:`1px solid ${p.border2}`, flexShrink:0,
                }}/>
                <span style={{ display:"inline-flex", gap:2 }}>
                  {[p.purple, p.cyan, p.green, p.amber, p.red].map((c,i) => (
                    <span key={i} style={{ width:9, height:9, borderRadius:"50%", background:c }}/>
                  ))}
                </span>
                {name}
              </button>
            ))}
            <button onClick={onReset} className="ht-chip" style={{
              padding:"6px 12px", borderRadius:8, border:`1px solid ${C.red}33`,
              background:`${C.red}11`, color:C.red, fontWeight:700, fontSize:12, cursor:"pointer", fontFamily:"inherit",
            }}>↺ Reset to Default</button>
          </div>
        </div>
      </Card>

      <div className="ht-engine-grid">
        {THEME_KEYS.map(({ k, label, hint }) => (
          <Card key={k} style={{ padding:14, gap:8 }}>
            <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", gap:10 }}>
              <div>
                <div style={{ fontSize:13, fontWeight:700, color:C.text }}>{label}</div>
                <div style={{ fontSize:11, color:C.muted, marginTop:2 }}>{hint}</div>
              </div>
              <ColorPicker value={theme[k] || "#000000"} onChange={v => set(k, v)}/>
            </div>
            {/* Preview swatch */}
            <div style={{
              height:32, borderRadius:8, marginTop:4,
              background: ["bg","card","card2","card3","border","border2","text","muted"].includes(k)
                ? theme[k]
                : `linear-gradient(90deg,${theme[k]},${theme[k]}88)`,
              border: `1px solid ${C.border}`,
            }}/>
          </Card>
        ))}
      </div>

      <Card>
        <Label accent={C.cyan}>🔍 Live Preview</Label>
        <div style={{ display:"flex", flexWrap:"wrap", gap:10 }}>
          <button style={{ padding:"10px 18px", borderRadius:10, border:"none",
            background:`linear-gradient(135deg,${C.violet},${C.purple})`, color:C.white,
            fontWeight:700, fontSize:13, cursor:"pointer", fontFamily:"inherit",
            boxShadow:`0 4px 12px ${C.purple}55`, }}>Primary Button</button>
          <span style={{ padding:"6px 12px", borderRadius:6, background:`${C.cyan}15`, color:C.cyan, fontSize:12, fontWeight:700, border:`1px solid ${C.cyan}33` }}>Accent Pill</span>
          <span style={{ padding:"6px 12px", borderRadius:6, background:`${C.green}15`, color:C.green, fontSize:12, fontWeight:700, border:`1px solid ${C.green}33` }}>✓ Success</span>
          <span style={{ padding:"6px 12px", borderRadius:6, background:`${C.amber}15`, color:C.amber, fontSize:12, fontWeight:700, border:`1px solid ${C.amber}33` }}>⚠ Warning</span>
          <span style={{ padding:"6px 12px", borderRadius:6, background:`${C.red}15`, color:C.red, fontSize:12, fontWeight:700, border:`1px solid ${C.red}33` }}>✗ Danger</span>
          <ScoreRing score={87} size={70}/>
        </div>
      </Card>
    </div>
  );
}

/* ─────────────── APP ─────────────── */
const DFLT = () => ({ date:getToday(), isHoliday:false });

function SyncBadge({ sync }) {
  const { state, msg } = sync;
  const map = {
    idle:    { color: C.muted,  bg: C.card3,        dot: C.muted,  label: "Idle" },
    syncing: { color: C.amber,  bg: `${C.amber}15`, dot: C.amber,  label: "Syncing…" },
    saved:   { color: C.green,  bg: `${C.green}15`, dot: C.green,  label: "Synced" },
    error:   { color: C.red,    bg: `${C.red}15`,   dot: C.red,    label: "Sync Error" },
  };
  const s = map[state] || map.idle;
  return (
    <div className="ht-sync" title={msg || s.label} style={{
      display:"flex", alignItems:"center", gap:6,
      padding:"5px 10px", borderRadius:6,
      background: s.bg, border:`1px solid ${s.color}33`,
      color: s.color, fontSize:11, fontWeight:700, letterSpacing:.5,
      maxWidth:200, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap",
    }}>
      <span style={{
        width:8, height:8, borderRadius:"50%", background: s.dot,
        boxShadow: state==="syncing" ? `0 0 8px ${s.dot}` : "none",
        animation: state==="syncing" ? "pulse 1s ease-in-out infinite" : "none",
        flexShrink:0,
      }}/>
      <span className="ht-sync-text" style={{ overflow:"hidden", textOverflow:"ellipsis" }}>{msg || s.label}</span>
    </div>
  );
}

/* ─────────────── PASSCODE GATE ───────────────
 * Front-door for the app. Verifies the passcode by hitting the Sheets backend,
 * which holds the real secret in Script Properties (APP_SECRET). The passcode
 * is stored in localStorage and sent with every request after unlock.
 */
function PasscodeGate({ onUnlock }) {
  const [pc, setPc] = useState("");
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState(null);

  const submit = async (e) => {
    e?.preventDefault?.();
    if (!pc.trim()) return;
    setBusy(true); setErr(null);
    try {
      const url = SHEETS_URL + "?k=" + encodeURIComponent(pc);
      const res = await fetch(url, { method: "GET", redirect: "follow" });
      const text = await res.text();
      if (text.trim().startsWith("<")) throw new Error("Backend unreachable");
      const json = JSON.parse(text);
      if (json && json.error === "unauthorized") {
        setErr("Wrong passcode");
        return;
      }
      onUnlock(pc);
    } catch (e) {
      setErr(e.message || "Could not verify");
    } finally {
      setBusy(false);
    }
  };

  return (
    <div style={{
      minHeight:"100vh", background:"#08080f", color:"#e2e8f0",
      display:"flex", alignItems:"center", justifyContent:"center",
      padding:20, fontFamily:"'Inter',system-ui,sans-serif",
    }}>
      <form onSubmit={submit} style={{
        width:"100%", maxWidth:380,
        background:"#0f0f1a", border:"1px solid #1e1e30", borderRadius:18,
        padding:"28px 24px", boxShadow:"0 24px 64px rgba(0,0,0,.5)",
        display:"flex", flexDirection:"column", gap:14,
      }}>
        <div style={{ display:"flex", alignItems:"center", gap:10, justifyContent:"center" }}>
          <span style={{ fontSize:28, filter:"drop-shadow(0 0 12px #06b6d488)" }}>⚡</span>
          <span style={{
            fontFamily:"'Space Grotesk',sans-serif", fontSize:22, fontWeight:800, letterSpacing:2,
            background:"linear-gradient(135deg,#06b6d4,#8b5cf6)",
            WebkitBackgroundClip:"text", backgroundClip:"text", WebkitTextFillColor:"transparent",
          }}>DISCIPLINE</span>
        </div>
        <div style={{ textAlign:"center", color:"#64748b", fontSize:13, lineHeight:1.5 }}>
          Enter your passcode to unlock the tracker.
        </div>
        <input
          type="password"
          autoFocus
          inputMode="text"
          autoComplete="current-password"
          placeholder="Passcode"
          value={pc}
          onChange={e => { setPc(e.target.value); setErr(null); }}
          disabled={busy}
          style={{
            background:"#191930", border:`1.5px solid ${err ? "#f43f5e" : "#25253a"}`,
            color:"#e2e8f0", padding:"12px 14px", borderRadius:10,
            fontSize:15, fontFamily:"inherit", outline:"none",
            transition:"border-color .2s",
          }}
        />
        {err && (
          <div style={{ color:"#f43f5e", fontSize:12, fontWeight:600, textAlign:"center" }}>
            ⚠️ {err}
          </div>
        )}
        <button
          type="submit"
          disabled={busy || !pc.trim()}
          style={{
            background: busy ? "#25253a" : "linear-gradient(135deg,#7c3aed,#8b5cf6)",
            color:"#fff", border:"none", padding:"12px 16px", borderRadius:10,
            fontWeight:800, letterSpacing:.5, fontSize:14, cursor: busy ? "wait" : "pointer",
            boxShadow: busy ? "none" : "0 8px 24px #8b5cf688",
            fontFamily:"inherit", textTransform:"uppercase",
          }}
        >
          {busy ? "Verifying…" : "Unlock"}
        </button>
        <div style={{ fontSize:11, color:"#4a5568", textAlign:"center", lineHeight:1.5 }}>
          The passcode is stored only on this device.
        </div>
      </form>
    </div>
  );
}

export default function App() {
  const [tab, setTab] = useState("dashboard");
  const [menuOpen, setMenuOpen] = useState(false);
  const [entries, setEntries] = useState([]);
  const [form, setForm] = useState(DFLT());
  const [saved, setSaved] = useState(false);
  const [schema, setSchema] = useState(DEFAULT_SCHEMA);
  const [info, setInfo] = useState(false);
  // sync: { state: 'idle'|'syncing'|'saved'|'error', msg?: string, ts?: number }
  const [sync, setSync] = useState({ state: "idle" });
  const [bootError, setBootError] = useState(null);
  // Theme: starts from defaults; mutates global C live; bumps version to re-render.
  const [theme, setTheme] = useState(DEFAULT_THEME);
  const [, setThemeVersion] = useState(0);
  // Day-detail modal: { ds: "YYYY-MM-DD" } | null
  const [dayDetail, setDayDetail] = useState(null);
  // Auth gate: shown when no passcode yet, or after a 401 from the backend.
  const [locked, setLocked] = useState(() => !getAuth());

  // Re-trigger boot fetch when the user unlocks.
  const [bootNonce, setBootNonce] = useState(0);

  // Load EVERYTHING from the sheet on mount.
  useEffect(() => {
    if (locked) return; // wait until passcode is entered
    // Purge legacy localStorage from previous versions so it can't shadow Sheet truth.
    try {
      localStorage.removeItem("habit_entries_v3");
      localStorage.removeItem("habit_schema_v3");
      localStorage.removeItem("habit_tracker_entries_v2");
    } catch {}
    (async () => {
      setSync({ state: "syncing", msg: "Loading from Sheet…" });
      try {
        const { entries: e, schema: s, theme: t } = await dbFetchAll();
        const effSchema = (s && s.length) ? s : DEFAULT_SCHEMA;
        // Normalise time fields immediately so <input type="time"> never sees
        // a Date/fraction round-tripped from Sheets.
        setEntries(e.map(en => normalizeEntry(en, effSchema)));
        if (s && s.length) setSchema(s);
        else {
          // First-ever boot: push default schema to the sheet so all devices share it.
          try { await dbWrite("schema", DEFAULT_SCHEMA); } catch {}
        }
        if (t) {
          const merged = { ...DEFAULT_THEME, ...t };
          setTheme(merged);
          Object.assign(C, merged);
          setThemeVersion(v => v + 1);
        }
        setSync({ state: "saved", msg: `Synced · ${e.length} entries`, ts: Date.now() });
      } catch (err) {
        if (err && err.code === 401) {
          // Bad / missing passcode — lock the UI.
          setAuth("");
          setLocked(true);
          setSync({ state: "error", msg: "Wrong passcode" });
          return;
        }
        console.error(err);
        setBootError(err.message || String(err));
        setSync({ state: "error", msg: err.message || "Load failed" });
      }
    })();
  }, [locked, bootNonce]);

  // Mutate global palette + persist theme to sheet.
  const updateTheme = async (next) => {
    setTheme(next);
    Object.assign(C, next);
    setThemeVersion(v => v + 1);
    setSync({ state: "syncing", msg: "Saving theme…" });
    try {
      await dbWrite("theme", next);
      setSync({ state: "saved", msg: "Theme synced", ts: Date.now() });
    } catch (err) {
      if (err && err.code === 401) { setAuth(""); setLocked(true); return; }
      console.error(err);
      setSync({ state: "error", msg: err.message || "Theme save failed" });
    }
  };
  const resetTheme = () => updateTheme(DEFAULT_THEME);

  // Persist schema to sheet only when user actually edits it (via updateSchema).
  const updateSchema = async (next) => {
    setSchema(next);
    setSync({ state: "syncing", msg: "Saving schema…" });
    try {
      await dbWrite("schema", next);
      setSync({ state: "saved", msg: "Schema synced", ts: Date.now() });
    } catch (err) {
      if (err && err.code === 401) { setAuth(""); setLocked(true); return; }
      console.error(err);
      setSync({ state: "error", msg: err.message || "Schema save failed" });
    }
  };

  const onDateChange = d => {
    const ex = entries.find(e => e.date === d);
    setForm(ex ? { ...DFLT(), ...normalizeEntry(ex, schema) } : { ...DFLT(), date: d });
  };

  const onSave = async () => {
    const { score, criteria, met, total } = calcScore(form, schema);
    const entry = { ...form, score, criteria, met, total };
    const updated = [...entries.filter(e => e.date !== form.date), entry].sort((a, b) => a.date.localeCompare(b.date));
    setEntries(updated);
    setSync({ state: "syncing", msg: "Saving to Sheet…" });
    try {
      await dbWrite("entries", updated);
      setSaved(true);
      setSync({ state: "saved", msg: `Synced · ${updated.length} entries`, ts: Date.now() });
      setTimeout(() => { setSaved(false); setTab("dashboard"); }, 1200);
    } catch (err) {
      if (err && err.code === 401) { setAuth(""); setLocked(true); return; }
      console.error(err);
      setSync({ state: "error", msg: err.message || "Save failed" });
      alert(`Save FAILED — not stored in Sheet.\n\n${err.message || err}\n\nFix the Apps Script and try again.`);
    }
  };

  const scored = useMemo(
    () => entries.map(e => ({ ...e, ...calcScore(e, schema) })).sort((a, b) => b.date.localeCompare(a.date)),
    [entries, schema]
  );
  const live = useMemo(() => calcScore(form, schema), [form, schema]);

  const yesterday = useMemo(() => {
    const ds = addDays(getToday(), -1);
    return entries.find(e => e.date === ds);
  }, [entries]);

  const onImport = async (mergedEntries) => {
    // Re-score every entry with current schema, then write to Sheet.
    const rescored = mergedEntries.map(e => {
      const { score, criteria, met, total } = calcScore(e, schema);
      return { ...e, score, criteria, met, total };
    });
    setEntries(rescored);
    setSync({ state: "syncing", msg: `Importing ${rescored.length} entries…` });
    try {
      await dbWrite("entries", rescored);
      setSync({ state: "saved", msg: `Imported · ${rescored.length} entries`, ts: Date.now() });
    } catch (err) {
      if (err && err.code === 401) { setAuth(""); setLocked(true); throw err; }
      console.error(err);
      setSync({ state: "error", msg: err.message || "Import failed" });
      throw err;
    }
  };

  const onQuickFill = () => {
    if (!yesterday) return;
    setForm({ ...DFLT(), ...yesterday, date: form.date });
  };

  // Calendar click → open detail modal (no redirect).
  const pickDate = (ds) => setDayDetail({ ds });
  const editFromModal = () => {
    if (!dayDetail) return;
    const ex = entries.find(e => e.date === dayDetail.ds);
    setForm(ex ? { ...DFLT(), ...normalizeEntry(ex, schema) } : { ...DFLT(), date: dayDetail.ds });
    setDayDetail(null);
    setTab("log");
  };
  const shiftDay = (delta) => {
    setDayDetail(cur => cur ? { ds: addDays(cur.ds, delta) } : cur);
  };

  // Don't block the entire UI on the first network round-trip.
  // The SyncBadge surfaces "Syncing…" / "Synced" / errors. The dashboard,
  // log, engine and settings tabs all render fine with empty data, and
  // populate as soon as dbFetchAll() resolves. This trims perceived load
  // time from ~3s to instant on cold boots.
  // (We keep the splash for the very first render only, when nothing is
  // available yet AND the user has no cached state.)

  const tabs = [
    { k:"dashboard", label:"Dashboard", icon:"📊" },
    { k:"log",       label:"Log",       icon:"✍️" },
    { k:"history",   label:"History",   icon:"🗂️" },
    { k:"engine",    label:"Engine",    icon:"⚙️" },
    { k:"settings",  label:"Settings",  icon:"🎨" },
  ];

  if (locked) {
    return <PasscodeGate onUnlock={(pc) => {
      setAuth(pc);
      setLocked(false);
      setBootError(null);
      setBootNonce(n => n + 1);
    }}/>;
  }

  return (
    <div style={{ background:C.bg, minHeight:"100vh", color:C.text, paddingBottom:60 }}>
      <GlobalStyles/>
      <InfoModal open={info} onClose={() => setInfo(false)} schema={schema}/>
      <DayDetailModal
        open={!!dayDetail}
        ds={dayDetail?.ds}
        scored={scored}
        schema={schema}
        onClose={() => setDayDetail(null)}
        onPrev={() => shiftDay(-1)}
        onNext={() => shiftDay(1)}
        onEdit={editFromModal}
      />

      {/* Header */}
      <header className="ht-header">
        <div className="ht-header-inner">
          <div className="ht-brand">
            <span className="ht-brand-icon">⚡</span>
            <span className="ht-brand-text">DISCIPLINE</span>
          </div>

          <nav className="ht-tabs">
            {tabs.map(t => (
              <button key={t.k} onClick={() => setTab(t.k)}
                className={`ht-tab ${tab===t.k ? "active" : ""}`}>
                <span className="ht-tab-icon">{t.icon}</span>
                <span className="ht-tab-label">{t.label}</span>
              </button>
            ))}
          </nav>

          {/* Mobile-only: compact menu trigger that reveals a dropdown of tabs */}
          <button
            type="button"
            className="ht-menu-trigger"
            aria-haspopup="true"
            aria-expanded={menuOpen}
            onClick={() => setMenuOpen(o => !o)}
          >
            <span className="ht-tab-icon">{tabs.find(t => t.k === tab)?.icon}</span>
            <span className="ht-menu-trigger-label">{tabs.find(t => t.k === tab)?.label}</span>
            <span className="ht-menu-caret" aria-hidden="true">{menuOpen ? "▲" : "▼"}</span>
          </button>
          {menuOpen && (
            <>
              <div className="ht-menu-backdrop" onClick={() => setMenuOpen(false)} />
              <div className="ht-menu-sheet" role="menu">
                {tabs.map(t => (
                  <button
                    key={t.k}
                    role="menuitem"
                    onClick={() => { setTab(t.k); setMenuOpen(false); }}
                    className={`ht-menu-item ${tab===t.k ? "active" : ""}`}
                  >
                    <span className="ht-tab-icon">{t.icon}</span>
                    <span>{t.label}</span>
                    {tab===t.k && <span className="ht-menu-check">✓</span>}
                  </button>
                ))}
              </div>
            </>
          )}

          <div className="ht-header-actions">
            <SyncBadge sync={sync}/>
            <span className="ht-date-pill">{fmtShort(getToday())}</span>
            <button onClick={() => setInfo(true)} className="ht-info-btn" title="How scoring works">i</button>
            <button
              onClick={() => { if (confirm("Lock the app? You'll need the passcode to unlock.")) { setAuth(""); setLocked(true); } }}
              className="ht-info-btn" title="Lock"
              style={{ borderColor: `${C.muted}66`, background: `${C.muted}11`, color: C.muted, fontStyle:"normal", fontFamily:"inherit" }}
            >🔒</button>
          </div>
        </div>
      </header>

      {bootError && (
        <div style={{
          maxWidth:1400, margin:"12px auto 0", padding:"12px 16px",
          background:`${C.red}15`, border:`1px solid ${C.red}55`, borderRadius:12,
          color:C.red, fontSize:13, lineHeight:1.5,
        }}>
          <b>⚠️ Sheet sync failed:</b> {bootError}<br/>
          <span style={{ color:C.muted }}>The app cannot save until the Apps Script is fixed. See the README/instructions for the correct <code>Code.gs</code>.</span>
        </div>
      )}

      <main className="ht-main">
        {tab==="dashboard" && <Dashboard scored={scored} schema={schema} onGoLog={() => { onDateChange(getToday()); setTab("log"); }} onPickDate={pickDate}/>}
        {tab==="log" && <LogEntry form={form} setForm={setForm} live={live} schema={schema} onDateChange={onDateChange} onSave={onSave} saved={saved} onQuickFill={onQuickFill} hasYesterday={!!yesterday}/>}
        {tab==="history" && <History scored={scored} schema={schema} onPick={(e) => { setForm({ ...DFLT(), ...normalizeEntry(e, schema) }); setTab("log"); }}/>}
        {tab==="engine" && <Engine schema={schema} setSchema={updateSchema} entries={entries} onImport={onImport}/>}
        {tab==="settings" && <Settings theme={theme} onChange={updateTheme} onReset={resetTheme}/>}
      </main>

      <footer style={{ textAlign:"center", padding:"24px 16px 12px", color:C.muted2, fontSize:11 }}>
        ⚡ Discipline · Built for the relentless
      </footer>
    </div>
  );
}

/* ─────────────── GLOBAL STYLES ─────────────── */
function GlobalStyles() {
  return (
    <style>{`
      *,*::before,*::after { box-sizing: border-box; -webkit-tap-highlight-color: transparent; }
      :root { color-scheme: dark; }
      h1,h2,h3,h4,h5 { font-family: 'Space Grotesk', system-ui, sans-serif; }

      input, textarea, select, button { font-family: inherit; }
      input:focus, textarea:focus, select:focus {
        border-color: ${C.purple} !important;
        box-shadow: 0 0 0 3px ${C.purple}22, 0 0 18px ${C.purple}33 !important;
      }

      ::-webkit-scrollbar { width: 10px; height: 10px; }
      ::-webkit-scrollbar-track { background: ${C.bg}; }
      ::-webkit-scrollbar-thumb { background: ${C.border2}; border-radius: 5px; border: 2px solid ${C.bg}; }
      ::-webkit-scrollbar-thumb:hover { background: ${C.purple}; }

      @keyframes spin { to { transform: rotate(360deg); } }
      @keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }
      @keyframes slideUp { from { transform: translateY(20px); opacity: 0; } to { transform: translateY(0); opacity: 1; } }
      @keyframes pulse { 0%,100% { opacity: 1; } 50% { opacity: .6; } }
      @keyframes shimmer { 0% { background-position: -200% 0; } 100% { background-position: 200% 0; } }
      @keyframes glow { 0%,100% { box-shadow: 0 0 20px ${C.purple}33; } 50% { box-shadow: 0 0 32px ${C.purple}66; } }

      .ht-pulse { animation: pulse 2s ease-in-out infinite; }
      .ht-card-glow { animation: glow 3s ease-in-out infinite; }

      /* HEADER */
      .ht-header {
        position: sticky; top: 0; z-index: 50;
        background: linear-gradient(180deg, ${C.card}f5, ${C.card2}ee);
        border-bottom: 1px solid ${C.border};
        backdrop-filter: blur(14px); -webkit-backdrop-filter: blur(14px);
      }
      .ht-header-inner {
        max-width: 1400px; margin: 0 auto; padding: 12px 20px;
        display: flex; align-items: center; gap: 16px;
      }
      .ht-brand {
        display: flex; align-items: center; gap: 8px; font-family: 'Space Grotesk', sans-serif;
      }
      .ht-brand-icon { font-size: 22px; filter: drop-shadow(0 0 10px ${C.cyan}88); }
      .ht-brand-text {
        font-size: 18px; font-weight: 800; letter-spacing: 2px;
        background: linear-gradient(135deg, ${C.cyan}, ${C.purple});
        -webkit-background-clip: text; background-clip: text;
        -webkit-text-fill-color: transparent;
      }
      .ht-tabs {
        flex: 1; display: flex; justify-content: center; gap: 4px;
        background: ${C.card3}; padding: 4px; border-radius: 12px; border: 1px solid ${C.border2};
        max-width: 520px; margin: 0 auto;
      }
      .ht-tab {
        flex: 1; padding: 8px 14px; border: none; background: transparent;
        color: ${C.muted}; cursor: pointer; font-weight: 700; font-size: 12px;
        border-radius: 9px; transition: all .25s; letter-spacing: .5px;
        display: flex; align-items: center; justify-content: center; gap: 6px;
        font-family: inherit;
      }
      .ht-tab:hover { color: ${C.text}; background: ${C.card2}; }
      .ht-tab.active {
        color: ${C.white}; background: linear-gradient(135deg, ${C.violet}, ${C.purple});
        box-shadow: 0 4px 14px ${C.purple}66;
      }
      .ht-tab-icon { font-size: 14px; }
      .ht-tab-label { text-transform: uppercase; }

      /* Mobile menu trigger (hidden on desktop) */
      .ht-menu-trigger {
        display: none;
        align-items: center; gap: 8px;
        padding: 8px 12px; border-radius: 10px;
        background: ${C.card3}; border: 1px solid ${C.border2};
        color: ${C.text}; font-weight: 700; font-size: 12px; letter-spacing: .5px;
        text-transform: uppercase; cursor: pointer; font-family: inherit;
      }
      .ht-menu-trigger:hover { border-color: ${C.purple}66; color: ${C.purple}; }
      .ht-menu-trigger-label { flex: 1; text-align: left; }
      .ht-menu-caret { font-size: 9px; color: ${C.muted}; }
      .ht-menu-backdrop {
        position: fixed; inset: 0; z-index: 60; background: rgba(0,0,0,.35);
        backdrop-filter: blur(2px);
      }
      .ht-menu-sheet {
        position: absolute; z-index: 61;
        top: calc(100% + 6px); right: 12px; left: 12px;
        background: ${C.card2}; border: 1px solid ${C.border2}; border-radius: 14px;
        box-shadow: 0 16px 48px rgba(0,0,0,.5);
        padding: 6px; display: flex; flex-direction: column; gap: 2px;
        animation: htMenuIn .18s ease-out;
      }
      @keyframes htMenuIn {
        from { opacity: 0; transform: translateY(-6px); }
        to   { opacity: 1; transform: translateY(0); }
      }
      .ht-menu-item {
        display: flex; align-items: center; gap: 12px;
        padding: 12px 14px; border: none; background: transparent;
        color: ${C.text}; font-weight: 600; font-size: 14px; cursor: pointer;
        border-radius: 10px; font-family: inherit; text-align: left;
      }
      .ht-menu-item:hover { background: ${C.card3}; }
      .ht-menu-item.active {
        background: linear-gradient(135deg, ${C.violet}22, ${C.purple}22);
        color: ${C.purple};
      }
      .ht-menu-check { margin-left: auto; color: ${C.purple}; font-weight: 800; }

      .ht-header-actions { display: flex; align-items: center; gap: 8px; }
      .ht-date-pill {
        font-size: 11px; color: ${C.muted}; font-weight: 700;
        padding: 5px 10px; border-radius: 6px; background: ${C.card3};
        border: 1px solid ${C.border2}; letter-spacing: .5px;
      }
      .ht-info-btn {
        width: 32px; height: 32px; border-radius: 50%;
        border: 1.5px solid ${C.cyan}66; background: ${C.cyan}11;
        color: ${C.cyan}; cursor: pointer; font-weight: 800; font-size: 14px;
        font-family: 'Space Grotesk', serif; font-style: italic;
        transition: all .25s;
      }
      .ht-info-btn:hover {
        background: ${C.cyan}33; border-color: ${C.cyan};
        box-shadow: 0 0 16px ${C.cyan}aa; transform: scale(1.1);
      }

      .ht-main { max-width: 1400px; margin: 0 auto; padding: 20px; }

      /* CARD HOVER GLOW */
      .ht-card:hover {
        border-color: ${C.purple}66 !important;
        box-shadow: 0 8px 32px ${C.purple}33, 0 0 0 1px ${C.purple}22 !important;
        transform: translateY(-2px);
      }

      /* STAT HOVER */
      .ht-stat:hover {
        transform: translateY(-3px);
        border-color: ${C.purple}66 !important;
        box-shadow: 0 8px 24px ${C.purple}44 !important;
      }

      /* CTA BUTTONS */
      .ht-cta:hover {
        transform: translateY(-2px) scale(1.02);
        box-shadow: 0 8px 24px ${C.purple}88 !important;
        filter: brightness(1.1);
      }
      .ht-cta:active { transform: translateY(0) scale(.98); }

      .ht-chip:hover {
        border-color: ${C.cyan} !important; color: ${C.cyan} !important;
        box-shadow: 0 0 14px ${C.cyan}55;
      }

      .ht-icon-btn:hover {
        border-color: ${C.purple} !important; color: ${C.purple} !important;
        box-shadow: 0 0 12px ${C.purple}66;
        transform: scale(1.05);
      }
      .ht-step-btn:hover {
        background: ${C.purple}22 !important; color: ${C.purple} !important;
        border-color: ${C.purple} !important;
      }
      .ht-step-btn:active { transform: scale(.9); }

      .ht-toggle:hover:not(:disabled) {
        box-shadow: 0 0 20px ${C.purple}88;
      }

      .ht-cal-cell:hover {
        transform: scale(1.12);
        box-shadow: 0 0 14px var(--glow, ${C.purple})cc !important;
        z-index: 2; position: relative;
      }

      .ht-history-item:hover {
        transform: translateX(4px);
        border-color: ${C.purple}88 !important;
        box-shadow: -2px 4px 20px ${C.purple}33;
      }

      .ht-field-row:hover {
        border-color: ${C.purple}66 !important;
        box-shadow: 0 4px 16px ${C.purple}22;
      }

      .ht-save:hover {
        transform: translateY(-2px) scale(1.01);
        box-shadow: 0 12px 32px ${C.purple}aa !important;
        filter: brightness(1.1);
      }

      /* GRID LAYOUTS — desktop responsive */
      .ht-dash {
        display: grid; grid-template-columns: 1fr; gap: 14px;
      }
      .ht-stats-grid {
        display: grid; grid-template-columns: repeat(3, 1fr); gap: 8px;
      }
      .ht-habit-grid { display: grid; grid-template-columns: 1fr; gap: 10px; }
      .ht-fields { display: grid; grid-template-columns: 1fr; gap: 8px; }
      .ht-log { display: grid; grid-template-columns: 1fr; gap: 12px; }
      .ht-history-grid { display: grid; grid-template-columns: 1fr; gap: 10px; }
      .ht-engine-grid { display: grid; grid-template-columns: 1fr; gap: 12px; }

      @media (min-width: 720px) {
        .ht-stats-grid { grid-template-columns: repeat(6, 1fr); }
        .ht-fields { grid-template-columns: 1fr 1fr; }
        .ht-history-grid { grid-template-columns: 1fr 1fr; }
        .ht-habit-grid { grid-template-columns: 1fr 1fr; gap: 14px 24px; }
        .ht-engine-grid { grid-template-columns: 1fr 1fr; }
      }

      @media (min-width: 1024px) {
        .ht-dash {
          grid-template-columns: 1.2fr 1fr;
        }
        .ht-habit-grid { grid-template-columns: 1fr 1fr 1fr; }
        .ht-fields { grid-template-columns: 1fr 1fr 1fr; }
      }

      @media (min-width: 1280px) {
        .ht-fields { grid-template-columns: 1fr 1fr 1fr 1fr; }
      }

      /* MOBILE TWEAKS */
      @media (max-width: 640px) {
        .ht-header-inner { padding: 10px 12px; gap: 10px; position: relative; flex-wrap: nowrap; min-width: 0; }
        .ht-brand { flex-shrink: 0; }
        .ht-tabs { display: none; }
        .ht-menu-trigger { display: inline-flex; flex: 1 1 0; min-width: 0; overflow: hidden; }
        .ht-menu-trigger-label { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
        .ht-brand-text { font-size: 15px; letter-spacing: 1.5px; }
        .ht-brand-icon { font-size: 18px; }
        .ht-date-pill { display: none; }
        .ht-header-actions { gap: 6px; flex-shrink: 0; }
        .ht-info-btn { width: 30px; height: 30px; font-size: 13px; }
        /* Collapse sync pill to just the dot to free room for the menu */
        .ht-sync { padding: 5px 7px !important; max-width: none !important; }
        .ht-sync-text { display: none; }
        .ht-main { padding: 14px; }
        .ht-hero { flex-direction: column; text-align: center; gap: 12px !important; padding: 16px !important; }
        .ht-live { flex-wrap: wrap; }
      }
    `}</style>
  );
}
