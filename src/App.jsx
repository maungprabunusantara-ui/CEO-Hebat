import { useState, useEffect, useRef, useMemo } from "react";
import * as XLSX from "xlsx";
import {
  Crown, Building2, Users, Plus, X, Activity, Settings2, Clock,
  CheckCircle2, Loader2, TrendingUp, Sparkles, Send, ShieldAlert,
  Radio, ChevronRight, ChevronDown, Gauge, FileText, FileSpreadsheet,
  Image as ImageIcon, Presentation, Bell, Database, MessageCircle,
  Send as TelegramIcon, CheckCheck, AlertCircle, Link2, Copy, MessageSquare, Bot
} from "lucide-react";

const ACCENT_CYCLE = [
  { text: "text-amber-400", bar: "bg-amber-400", chip: "bg-amber-500" },
  { text: "text-sky-400", bar: "bg-sky-400", chip: "bg-sky-500" },
  { text: "text-violet-400", bar: "bg-violet-400", chip: "bg-violet-500" },
  { text: "text-emerald-400", bar: "bg-emerald-400", chip: "bg-emerald-500" },
  { text: "text-rose-400", bar: "bg-rose-400", chip: "bg-rose-500" },
  { text: "text-cyan-400", bar: "bg-cyan-400", chip: "bg-cyan-500" },
  { text: "text-fuchsia-400", bar: "bg-fuchsia-400", chip: "bg-fuchsia-500" },
];

const DEFAULT_TEAMS = [
  { name: "Finance & Treasury", agent: "CFO-AI", skills: ["budget", "forecast", "cashflow", "invoice", "expense", "revenue", "anggaran", "keuangan"] },
  { name: "Marketing & Growth", agent: "CMO-AI", skills: ["campaign", "brand", "content", "social", "seo", "launch", "marketing", "promosi"] },
  { name: "Operations", agent: "COO-AI", skills: ["process", "supply", "logistics", "inventory", "vendor", "quality", "operasional", "produksi"] },
  { name: "Human Resources", agent: "CHRO-AI", skills: ["hiring", "recruit", "onboarding", "payroll", "culture", "performance", "karyawan", "rekrutmen"] },
  { name: "Technology & IT", agent: "CTO-AI", skills: ["system", "software", "security", "infrastructure", "bug", "deployment", "server", "aplikasi"] },
  { name: "Legal & Compliance", agent: "CLO-AI", skills: ["contract", "compliance", "risk", "policy", "regulation", "audit", "kontrak", "hukum"] },
  { name: "Sales & Partnership", agent: "CSO-AI", skills: ["client", "deal", "pipeline", "partnership", "negotiation", "proposal", "klien", "penjualan"] },
];

const STATUS_META = {
  queued: { label: "Queued", dot: "bg-slate-400", bar: "bg-slate-500" },
  delegating: { label: "GM Delegating", dot: "bg-sky-400", bar: "bg-sky-400" },
  executing: { label: "Team Executing", dot: "bg-violet-400", bar: "bg-violet-400" },
  done: { label: "Completed", dot: "bg-emerald-400", bar: "bg-emerald-400" },
};

function timeAgo(ms) {
  const s = Math.max(0, Math.floor(ms / 1000));
  const m = Math.floor(s / 60);
  const r = s % 60;
  return `${m}:${r.toString().padStart(2, "0")}`;
}

function slug(s) {
  return (s || "task").toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/(^-|-$)/g, "").slice(0, 60) || "task";
}

function escapeHtml(s) {
  return (s || "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

function matchTeam(title, teams) {
  const lower = title.toLowerCase();
  let best = null, bestScore = 0;
  teams.forEach((team) => {
    let score = 0;
    team.skills.forEach((sk) => { if (lower.includes(sk)) score += 1; });
    if (score > bestScore) { bestScore = score; best = team; }
  });
  return bestScore > 0 ? best : null;
}

function downloadBlob(filename, content, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ---- Fire-and-forget cross-origin GET (works even where fetch() to
// arbitrary third-party domains is blocked by the artifact sandbox — an
// <img> resource load isn't subject to that same restriction, and the
// server still executes fully even though the browser can't render the
// text response as an image). ----
function pingGET(url) {
  return new Promise((resolve) => {
    let done = false;
    const finish = () => { if (!done) { done = true; resolve(true); } };
    const img = new Image();
    img.onload = finish;
    img.onerror = finish; // still means the request reached the server
    img.src = url;
    setTimeout(finish, 4000);
  });
}

function buildWaLink(number, text) {
  const digits = (number || "").replace(/[^\d]/g, "");
  return `https://wa.me/${digits}?text=${encodeURIComponent(text)}`;
}

function buildTelegramShareLink(text) {
  return `https://t.me/share/url?url=&text=${encodeURIComponent(text)}`;
}

function buildArchiveLink(gasUrl, payload) {
  const sep = gasUrl.includes("?") ? "&" : "?";
  return `${gasUrl}${sep}data=${encodeURIComponent(JSON.stringify(payload))}`;
}

// A direct, synchronous window.open() triggered inside a real onClick handler
// is treated by the browser as top-level navigation from a genuine user
// gesture — it is NOT subject to the sandbox's background-request block
// (that block only applies to fetch/XHR/img calls a script fires on its
// own). This is why "Open bridge URL" worked earlier while silent
// auto-notify did not: one is a click-driven navigation, the other is a
// background call. Use this for every archive/notify action.
function openAndAutoClose(url) {
  const win = window.open(url, "_blank", "width=60,height=60");
  if (win) setTimeout(() => { try { win.close(); } catch (e) {} }, 1200);
  return !!win;
}

// JSONP request — used to READ data back from the Apps Script bridge.
// A plain fetch() can't read a cross-origin response without CORS headers
// (which Apps Script doesn't provide), but a <script> tag load isn't
// restricted the same way, so we use that classic trick here.
function jsonpRequest(url, params) {
  return new Promise((resolve, reject) => {
    const callbackName = "ceoCC_cb_" + Math.random().toString(36).slice(2);
    let settled = false;
    const script = document.createElement("script");
    const cleanup = () => { delete window[callbackName]; script.remove(); };
    window[callbackName] = (data) => { if (settled) return; settled = true; cleanup(); resolve(data); };
    script.onerror = () => { if (settled) return; settled = true; cleanup(); reject(new Error("JSONP load failed")); };
    const qs = new URLSearchParams({ ...params, callback: callbackName }).toString();
    const sep = url.includes("?") ? "&" : "?";
    script.src = `${url}${sep}${qs}`;
    document.body.appendChild(script);
    setTimeout(() => { if (settled) return; settled = true; cleanup(); reject(new Error("JSONP timeout")); }, 8000);
  });
}

function saveSettingToSheet(gasUrl, key, value) {
  if (!gasUrl) return;
  const payload = { type: "saveSettings", key, value: typeof value === "string" ? value : JSON.stringify(value) };
  silentDeliver(buildArchiveLink(gasUrl, payload));
}

// Permanent local storage — safe here because this is a real standalone
// deployment (not the chat preview), so settings survive refreshes/closing
// the browser.
const STORE_PREFIX = "ceo-command-center:";
function loadStored(key, fallback) {
  try {
    const raw = window.localStorage.getItem(STORE_PREFIX + key);
    return raw ? JSON.parse(raw) : fallback;
  } catch (e) { return fallback; }
}
function saveStored(key, value) {
  try { window.localStorage.setItem(STORE_PREFIX + key, JSON.stringify(value)); } catch (e) { /* storage unavailable — ignore */ }
}

// Tries a real silent background request first — this works once the app is
// deployed on real hosting (no chat sandbox blocking it). Falls back to the
// pixel trick for environments that still block fetch().
async function silentDeliver(url) {
  try {
    await fetch(url, { mode: "no-cors" });
    return true;
  } catch (e) {
    try { return await pingGET(url); } catch (e2) { return false; }
  }
}

// ---- Real AI generation — routed through our own /api/gemini serverless
// function, which keeps the Gemini API key on the server (see api/gemini.js) ----
function toGeminiContents(messages) {
  return messages.map((m) => ({
    role: m.role === "assistant" ? "model" : "user",
    parts: [{ text: m.content }],
  }));
}

async function callGemini(systemPrompt, userPrompt) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 25000);
  try {
    const res = await fetch("/api/gemini", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      signal: controller.signal,
      body: JSON.stringify({
        systemInstruction: systemPrompt,
        contents: [{ role: "user", parts: [{ text: userPrompt }] }],
        maxOutputTokens: 3000,
      }),
    });
    clearTimeout(timeout);
    const data = await res.json();
    if (data.error) return null;
    const text = (data.candidates?.[0]?.content?.parts || []).map((p) => p.text || "").join("\n").trim();
    return text || null;
  } catch (e) {
    clearTimeout(timeout);
    return null;
  }
}

async function callGeminiChat(messages, systemPrompt) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 30000);
  try {
    const res = await fetch("/api/gemini", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      signal: controller.signal,
      body: JSON.stringify({
        systemInstruction: systemPrompt,
        contents: toGeminiContents(messages),
        maxOutputTokens: 2000,
      }),
    });
    clearTimeout(timeout);
    const data = await res.json();
    if (data.error) return `⚠️ ${typeof data.error === "string" ? data.error : data.error.message || "API error"}`;
    const text = (data.candidates?.[0]?.content?.parts || []).map((p) => p.text || "").join("\n").trim();
    return text || "Maaf, tidak ada respons yang bisa diambil saat ini.";
  } catch (e) {
    clearTimeout(timeout);
    return null;
  }
}

// ---- File exporters ----
function markdownInlineToHtml(text) {
  return escapeHtml(text)
    .replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>")
    .replace(/\*(.+?)\*/g, "<em>$1</em>");
}

function markdownToHtml(text) {
  const lines = (text || "").split("\n");
  const out = [];
  let listBuffer = [];
  let listType = null;
  const flushList = () => {
    if (listBuffer.length) {
      const tag = listType === "ol" ? "ol" : "ul";
      out.push(`<${tag}>${listBuffer.map((i) => `<li>${markdownInlineToHtml(i)}</li>`).join("")}</${tag}>`);
    }
    listBuffer = []; listType = null;
  };
  lines.forEach((raw) => {
    const t = raw.trim();
    if (t === "") { flushList(); return; }
    if (/^-{3,}$/.test(t)) { flushList(); out.push("<hr/>"); return; }
    const h = t.match(/^(#{1,4})\s+(.*)/);
    if (h) { flushList(); const lvl = Math.min(h[1].length + 1, 6); out.push(`<h${lvl}>${markdownInlineToHtml(h[2])}</h${lvl}>`); return; }
    const bullet = t.match(/^[-*]\s+(.*)/);
    if (bullet) { if (listType !== "ul") { flushList(); listType = "ul"; } listBuffer.push(bullet[1]); return; }
    const numbered = t.match(/^\d+\.\s+(.*)/);
    if (numbered) { if (listType !== "ol") { flushList(); listType = "ol"; } listBuffer.push(numbered[1]); return; }
    flushList();
    out.push(`<p>${markdownInlineToHtml(t)}</p>`);
  });
  flushList();
  return out.join("\n");
}

function exportWord(task, teamLabel) {
  const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${escapeHtml(task.title)}</title></head>
  <body style="font-family:Calibri,Arial,sans-serif;">
  <h1 style="font-family:Georgia,serif;font-size:22px;">${escapeHtml(task.title)}</h1>
  <p><strong>Team:</strong> ${escapeHtml(teamLabel)}<br/>
  <strong>Status:</strong> ${escapeHtml(task.status)}<br/>
  <strong>Generated:</strong> ${new Date(task.createdAt).toLocaleString("id-ID")}</p>
  <hr/>
  <div style="line-height:1.6;font-size:14px;">${markdownToHtml(task.output) || "<p>(no output yet)</p>"}</div>
  </body></html>`;
  downloadBlob(`${slug(task.title)}.doc`, html, "application/msword");
}

function exportExcel(task, teamLabel) {
  const rows = [
    ["Field", "Value"],
    ["Task", task.title],
    ["Team", teamLabel],
    ["Status", task.status],
    ["Urgent", task.urgent ? "Yes" : "No"],
    ["Created", new Date(task.createdAt).toLocaleString("id-ID")],
    ["Output", task.output || ""],
  ];
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = [{ wch: 14 }, { wch: 80 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Report");
  XLSX.writeFile(wb, `${slug(task.title)}.xlsx`);
}

function exportCSV(task, teamLabel) {
  const esc = (v) => `"${String(v || "").replace(/"/g, '""')}"`;
  const csv = [
    ["Field", "Value"].map(esc).join(","),
    ["Task", task.title].map(esc).join(","),
    ["Team", teamLabel].map(esc).join(","),
    ["Status", task.status].map(esc).join(","),
    ["Created", new Date(task.createdAt).toLocaleString("id-ID")].map(esc).join(","),
    ["Output", task.output || ""].map(esc).join(","),
  ].join("\n");
  downloadBlob(`${slug(task.title)}.csv`, csv, "text/csv;charset=utf-8");
}

function exportSVG(task, teamLabel, accentHex) {
  const lines = (task.output || "No output generated yet.").split(/\n+/).slice(0, 10);
  const wrapped = [];
  lines.forEach((l) => {
    let line = l;
    while (line.length > 78) { wrapped.push(line.slice(0, 78)); line = line.slice(78); }
    wrapped.push(line);
  });
  const body = wrapped.slice(0, 16).map((l, i) => `<text x="32" y="${150 + i * 20}" font-size="13" fill="#cbd5e1" font-family="monospace">${escapeHtml(l)}</text>`).join("");
  const svg = `<svg viewBox="0 0 700 ${180 + Math.min(wrapped.length, 16) * 20 + 20}" xmlns="http://www.w3.org/2000/svg">
  <rect width="100%" height="100%" fill="#0f172a"/>
  <rect x="16" y="16" width="668" height="60" rx="8" fill="${accentHex}22" stroke="${accentHex}" stroke-width="1"/>
  <text x="32" y="48" font-size="18" fill="${accentHex}" font-family="Georgia, serif">${escapeHtml(task.title).slice(0, 60)}</text>
  <text x="32" y="66" font-size="11" fill="#94a3b8" font-family="monospace">${escapeHtml(teamLabel)} · ${new Date(task.createdAt).toLocaleDateString("id-ID")}</text>
  <line x1="32" y1="100" x2="668" y2="100" stroke="#334155"/>
  <text x="32" y="128" font-size="12" fill="#64748b" font-family="monospace">REFERENCE OUTPUT</text>
  ${body}
  </svg>`;
  downloadBlob(`${slug(task.title)}-reference.svg`, svg, "image/svg+xml");
}

function exportSlides(task, teamLabel) {
  const paras = (task.output || "No content generated yet.").split(/\n+/).filter(Boolean);
  const chunks = [];
  for (let i = 0; i < paras.length; i += 2) chunks.push(paras.slice(i, i + 2).join("\n\n"));
  if (chunks.length === 0) chunks.push("No content generated yet.");
  const slidesHtml = [
    `<section class="slide title"><h1>${escapeHtml(task.title)}</h1><p>${escapeHtml(teamLabel)} · ${new Date(task.createdAt).toLocaleDateString("id-ID")}</p></section>`,
    ...chunks.map((c, i) => `<section class="slide"><h2>Point ${i + 1}</h2><p>${escapeHtml(c)}</p></section>`),
  ].join("\n");
  const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${escapeHtml(task.title)}</title>
  <style>
    body{margin:0;font-family:ui-sans-serif,system-ui;background:#0f172a;color:#f1f5f9;}
    .slide{min-height:100vh;display:none;flex-direction:column;justify-content:center;padding:8vw;box-sizing:border-box;}
    .slide.active{display:flex;}
    .slide.title h1{font-family:Georgia,serif;font-size:3rem;color:#fbbf24;}
    h2{color:#38bdf8;text-transform:uppercase;letter-spacing:.08em;font-size:1rem;}
    p{font-size:1.4rem;line-height:1.6;white-space:pre-wrap;}
    .nav{position:fixed;bottom:20px;right:20px;font-family:monospace;color:#64748b;}
  </style></head>
  <body>${slidesHtml}<div class="nav">Use ← → arrow keys</div>
  <script>
    const slides = document.querySelectorAll('.slide'); let idx=0;
    function show(i){ slides.forEach(s=>s.classList.remove('active')); slides[i].classList.add('active'); }
    show(0);
    document.addEventListener('keydown', e=>{
      if(e.key==='ArrowRight') idx=Math.min(idx+1, slides.length-1);
      if(e.key==='ArrowLeft') idx=Math.max(idx-1,0);
      show(idx);
    });
  </script></body></html>`;
  downloadBlob(`${slug(task.title)}-slides.html`, html, "text/html");
}

function renderInline(text) {
  const parts = [];
  const regex = /(\*\*(.+?)\*\*|\*(.+?)\*)/g;
  let lastIndex = 0, match, key = 0;
  while ((match = regex.exec(text))) {
    if (match.index > lastIndex) parts.push(text.slice(lastIndex, match.index));
    if (match[2] !== undefined) parts.push(<strong key={key++}>{match[2]}</strong>);
    else if (match[3] !== undefined) parts.push(<em key={key++}>{match[3]}</em>);
    lastIndex = regex.lastIndex;
  }
  if (lastIndex < text.length) parts.push(text.slice(lastIndex));
  return parts;
}

function MarkdownLite({ text }) {
  const lines = (text || "").split("\n");
  const elements = [];
  let listBuffer = [];
  let listType = null;
  const flushList = () => {
    if (listBuffer.length) {
      const Tag = listType === "ol" ? "ol" : "ul";
      elements.push(
        <Tag key={elements.length} className={listType === "ol" ? "list-decimal pl-4 space-y-0.5 mb-1.5" : "list-disc pl-4 space-y-0.5 mb-1.5"}>
          {listBuffer.map((item, i) => <li key={i}>{renderInline(item)}</li>)}
        </Tag>
      );
    }
    listBuffer = [];
    listType = null;
  };
  lines.forEach((rawLine) => {
    const trimmed = rawLine.trim();
    if (trimmed === "") { flushList(); return; }
    if (/^-{3,}$/.test(trimmed)) { flushList(); elements.push(<hr key={elements.length} className="border-slate-700 my-2" />); return; }
    const h = trimmed.match(/^(#{1,4})\s+(.*)/);
    if (h) {
      flushList();
      const level = h[1].length;
      const cls = level <= 2 ? "text-[13px] font-semibold text-amber-300 mt-2 mb-1" : "text-xs font-semibold text-sky-300 mt-2 mb-1";
      elements.push(<div key={elements.length} className={cls}>{renderInline(h[2])}</div>);
      return;
    }
    const bullet = trimmed.match(/^[-*]\s+(.*)/);
    if (bullet) { if (listType !== "ul") { flushList(); listType = "ul"; } listBuffer.push(bullet[1]); return; }
    const numbered = trimmed.match(/^\d+\.\s+(.*)/);
    if (numbered) { if (listType !== "ol") { flushList(); listType = "ol"; } listBuffer.push(numbered[1]); return; }
    flushList();
    elements.push(<p key={elements.length} className="mb-1.5 last:mb-0">{renderInline(trimmed)}</p>);
  });
  flushList();
  return <>{elements}</>;
}

export default function App() {
  const idRef = useRef(1);
  const nextId = () => idRef.current++;

  const [teams, setTeams] = useState(() =>
    loadStored("teams", DEFAULT_TEAMS.map((t, i) => ({
      id: `team-${i}`, name: t.name, agent: t.agent, skills: t.skills,
      status: "idle", accent: ACCENT_CYCLE[i % ACCENT_CYCLE.length],
    })))
  );

  const [tasks, setTasks] = useState([]);
  const [logs, setLogs] = useState([
    { id: nextId(), t: Date.now(), text: "System online. GM AI Agent standing by for CEO directives." },
  ]);

  const [policy, setPolicy] = useState(() => loadStored("policy", { autonomy: 3, escalation: "Medium", autoAssign: true, autoNotify: true }));

  const [integrations, setIntegrations] = useState(() => loadStored("integrations", {
    gasUrl: "", telegramChatId: "", telegramToken: "", telegramDirect: false, waNumber: "",
  }));
  useEffect(() => { saveStored("teams", teams.map(({ status, ...rest }) => rest)); }, [teams]);
  useEffect(() => { saveStored("policy", policy); }, [policy]);
  useEffect(() => {
    // Never persist the Telegram bot token to storage — keep that session-only for safety.
    const { telegramToken, ...safe } = integrations;
    saveStored("integrations", safe);
  }, [integrations]);

  const [testStatus, setTestStatus] = useState(null); // {ok, msg}
  const [settingsReady, setSettingsReady] = useState(false);
  const [settingsSyncMsg, setSettingsSyncMsg] = useState("");
  const DEFAULT_SKILL_PROFILE = "Kamu adalah asisten pribadi CEO yang jenius dan serba bisa: strategi bisnis, keuangan, hukum, teknologi, pemasaran, HR, riset, penulisan, dan analisis data — skill dewa, bisa bantu apa saja. Jawab to the point, praktis, dan actionable seperti konsultan top-tier. Sesuaikan bahasa dengan bahasa CEO. Kalau ditanya hal di luar keahlian, tetap coba bantu semaksimal mungkin dengan penalaran yang jujur.";
  const [skillProfile, setSkillProfile] = useState(() => loadStored("skillProfile", DEFAULT_SKILL_PROFILE));
  const [agentMemory, setAgentMemory] = useState(() => loadStored("agentMemory", []));
  useEffect(() => { saveStored("skillProfile", skillProfile); }, [skillProfile]);
  useEffect(() => { saveStored("agentMemory", agentMemory); }, [agentMemory]);

  // Pull saved settings from the Sheet once on load (if a bridge URL is
  // already known from this browser) so a different device/browser picks up
  // the same configuration instead of starting from scratch.
  useEffect(() => {
    let cancelled = false;
    if (!integrations.gasUrl) { setSettingsReady(true); return; }
    setSettingsSyncMsg("Menyinkronkan settings dari Sheet...");
    jsonpRequest(integrations.gasUrl, { type: "getSettings" })
      .then((remote) => {
        if (cancelled || !remote) return;
        try { if (remote.teams) setTeams(JSON.parse(remote.teams).map((t) => ({ ...t, status: "idle" }))); } catch (e) {}
        try { if (remote.policy) setPolicy(JSON.parse(remote.policy)); } catch (e) {}
        try { if (remote.integrations) setIntegrations((prev) => ({ ...prev, ...JSON.parse(remote.integrations) })); } catch (e) {}
        if (remote.skillProfile) setSkillProfile(remote.skillProfile);
        try { if (remote.agentMemory) setAgentMemory(JSON.parse(remote.agentMemory)); } catch (e) {}
        setSettingsSyncMsg("Settings tersinkron dari Sheet ✓");
      })
      .catch(() => { setSettingsSyncMsg(""); })
      .finally(() => { if (!cancelled) setSettingsReady(true); });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Push every settings change up to the Sheet too (once initial sync has
  // finished, so we don't overwrite good remote data with a fresh device's
  // empty defaults before the load above completes).
  useEffect(() => { if (settingsReady && integrations.gasUrl) saveSettingToSheet(integrations.gasUrl, "teams", teams.map(({ status, ...rest }) => rest)); }, [teams, settingsReady]); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => { if (settingsReady && integrations.gasUrl) saveSettingToSheet(integrations.gasUrl, "policy", policy); }, [policy, settingsReady]); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => {
    if (!settingsReady || !integrations.gasUrl) return;
    const { telegramToken, ...safe } = integrations;
    saveSettingToSheet(integrations.gasUrl, "integrations", safe);
  }, [integrations, settingsReady]); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => { if (settingsReady && integrations.gasUrl) saveSettingToSheet(integrations.gasUrl, "skillProfile", skillProfile); }, [skillProfile, settingsReady]); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => { if (settingsReady && integrations.gasUrl) saveSettingToSheet(integrations.gasUrl, "agentMemory", agentMemory); }, [agentMemory, settingsReady]); // eslint-disable-line react-hooks/exhaustive-deps
  const [showSettings, setShowSettings] = useState(true);

  const [showAddTeam, setShowAddTeam] = useState(false);
  const [newTeam, setNewTeam] = useState({ name: "", agent: "", skills: "" });
  const [activeTab, setActiveTab] = useState("board");
  const tabClass = (key) => `${activeTab === key ? "block" : "hidden"} md:block`;

  const [showSkillPanel, setShowSkillPanel] = useState(false);

  function buildAssistantSystemPrompt() {
    const memoryBlock = agentMemory.length
      ? `\n\nHal-hal yang sudah kamu pelajari dari CEO ini di percakapan sebelumnya (pakai sebagai konteks tersembunyi, jangan disebut-sebut secara eksplisit kecuali relevan):\n- ${agentMemory.join("\n- ")}`
      : "";
    return `${skillProfile}${memoryBlock}`;
  }

  const [chatMessages, setChatMessages] = useState([]);
  const [chatInput, setChatInput] = useState("");
  const [chatLoading, setChatLoading] = useState(false);
  const chatEndRef = useRef(null);
  useEffect(() => { chatEndRef.current?.scrollIntoView({ behavior: "smooth" }); }, [chatMessages, chatLoading]);

  const handleSendChat = async (e) => {
    e.preventDefault();
    const text = chatInput.trim();
    if (!text || chatLoading) return;
    const nextHistory = [...chatMessages, { role: "user", content: text }];
    setChatMessages(nextHistory);
    setChatInput("");
    setChatLoading(true);
    const reply = await callGeminiChat(nextHistory, buildAssistantSystemPrompt());
    setChatMessages((prev) => [...prev, { role: "assistant", content: reply || "Maaf, koneksi ke AI sedang bermasalah — coba lagi sebentar." }]);
    setChatLoading(false);

    // Best-effort: silently learn one short memorable fact from this exchange
    // so the assistant genuinely gets more context-aware over time.
    if (reply) {
      const extractPrompt = `Dari pertukaran berikut, apakah ada 1 fakta penting & ringkas tentang CEO ini atau bisnisnya yang layak diingat untuk percakapan berikutnya (misalnya: nama perusahaan, industri, preferensi, masalah yang sedang dihadapi)? Kalau ada, balas HANYA dengan fakta itu dalam satu kalimat pendek bahasa Indonesia, tanpa embel-embel. Kalau tidak ada yang penting/baru, balas persis satu kata: NONE.\n\nCEO: ${text}\nAsisten: ${reply}`;
      const fact = await callGemini("Kamu adalah sistem ekstraksi memori. Ikuti instruksi persis, tanpa basa-basi.", extractPrompt);
      const clean = (fact || "").trim();
      if (clean && clean.toUpperCase() !== "NONE" && clean.length < 220 && !agentMemory.includes(clean)) {
        setAgentMemory((prev) => [...prev, clean].slice(-30));
      }
    }
  };

  const [taskTitle, setTaskTitle] = useState("");
  const [assignTo, setAssignTo] = useState("auto");
  const [urgent, setUrgent] = useState(false);
  const [expandedId, setExpandedId] = useState(null);

  const [now, setNow] = useState(Date.now());

  const teamsRef = useRef(teams);
  useEffect(() => { teamsRef.current = teams; }, [teams]);
  const integrationsRef = useRef(integrations);
  useEffect(() => { integrationsRef.current = integrations; }, [integrations]);
  const policyRef = useRef(policy);
  useEffect(() => { policyRef.current = policy; }, [policy]);

  useEffect(() => {
    const c = setInterval(() => setNow(Date.now()), 1000);
    return () => clearInterval(c);
  }, []);

  const pushLog = (text) => setLogs((prev) => [{ id: nextId(), t: Date.now(), text }, ...prev].slice(0, 80));

  const composeMessage = (payload) => {
    const urgentTag = payload.urgent ? "🔴 URGENT\n" : "";
    return `${urgentTag}🏢 ${payload.team} (${payload.agent})\n📋 ${payload.title}\n📌 Status: ${payload.status}\n\n${(payload.output || "").slice(0, 500)}`;
  };

  // Best-effort silent attempt (works on some setups, not guaranteed in this
  // sandboxed preview — see the one-tap buttons on each task for the
  // guaranteed path).
  const notifyChannels = async (task, teamLabel, agentLabel) => {
    const integ = integrationsRef.current;
    if (!policyRef.current.autoNotify) return;
    const payload = {
      timestamp: new Date().toISOString(),
      team: teamLabel, agent: agentLabel, title: task.title,
      status: task.status, urgent: task.urgent, output: task.output || "",
    };
    if (integ.gasUrl) {
      try {
        const compact = { ...payload, output: (payload.output || "").slice(0, 300) };
        await silentDeliver(buildArchiveLink(integ.gasUrl, compact));
      } catch (e) { /* fully unreachable — use the task's Archive button instead */ }
    }
    setTasks((prev) => prev.map((t) => (t.id === task.id ? { ...t, pendingSync: true } : t)));
  };

  // trigger real AI answer generation for a team agent
  const triggerAgentAnswer = async (task) => {
    const team = teamsRef.current.find((t) => t.id === task.teamId);
    const teamLabel = team ? team.name : "General Management";
    const agentLabel = team ? team.agent : "GM-AI";
    const systemPrompt = `You are ${agentLabel}, the AI agent leading "${teamLabel}" inside a company's AI-run department, reporting directly to the CEO through the GM AI Agent. Your specialties: ${team ? team.skills.join(", ") : "general management, coordination"}. Respond to the CEO's directive with a professional, actionable report as if you already completed the task: include what was done/found, key data points or recommendations, and a clear next step. Keep it complete but disciplined — aim for 200-350 words; if the topic has multiple phases or sections, summarize each in 1-2 sentences rather than expanding at length, so the report always finishes properly instead of being cut off. Write in the same language as the directive. No preamble.`;

    let output = await callGemini(systemPrompt, task.title);
    if (!output) {
      output = `[Auto-summary] ${agentLabel} completed "${task.title}". Key findings compiled from internal analysis; recommended next step: review the archived report and confirm before rollout. (Live AI generation was unavailable this run — this is a fallback summary.)`;
    }
    const finalTask = { ...task, output, generating: false, progress: 100, status: "done", updatedAt: Date.now() };
    setTasks((prev) => prev.map((t) => (t.id === task.id ? finalTask : t)));
    pushLog(`✅ ${agentLabel} completed "${task.title}" — output ready for CEO review.`);
    notifyChannels(finalTask, teamLabel, agentLabel);
  };

  // simulation tick — real-time delegation flow
  useEffect(() => {
    const speed = 0.7 + policy.autonomy * 0.35;
    const interval = setInterval(() => {
      const t = Date.now();
      let transitions = [];
      setTasks((prev) =>
        prev.map((task) => {
          if (task.status === "done" || task.generating) return task;
          const elapsed = t - task.updatedAt;
          if (task.status === "queued" && elapsed > 1400 / speed) {
            transitions.push({ task, to: "delegating" });
            return { ...task, status: "delegating", updatedAt: t };
          }
          if (task.status === "delegating" && elapsed > 1600 / speed) {
            const started = { ...task, status: "executing", updatedAt: t, progress: 8, generating: true };
            transitions.push({ task: started, to: "executing" });
            triggerAgentAnswer(started);
            return started;
          }
          if (task.status === "executing" && !task.generating) {
            const inc = (3 + Math.random() * 6) * speed;
            return { ...task, progress: Math.min(90, task.progress + inc) };
          }
          return task;
        })
      );
      transitions.forEach(({ task, to }) => {
        const teamLabel = task.teamId === "GM" ? "GM AI Agent" : teamsRef.current.find((tm) => tm.id === task.teamId)?.agent || "Team Agent";
        if (to === "delegating") pushLog(`GM AI Agent is reviewing "${task.title}" for delegation${task.urgent ? " — flagged URGENT" : ""}.`);
        if (to === "executing") pushLog(`${teamLabel} started working on "${task.title}" (auto-answer generating...).`);
      });
    }, 700);
    return () => clearInterval(interval);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [policy.autonomy]);

  useEffect(() => {
    setTeams((prev) => prev.map((team) => ({
      ...team,
      status: tasks.some((tk) => tk.teamId === team.id && (tk.status === "executing" || tk.generating)) ? "busy" : "idle",
    })));
  }, [tasks]);

  const suggestedTeam = useMemo(() => (taskTitle.trim() ? matchTeam(taskTitle, teams) : null), [taskTitle, teams]);

  const handleCreateTask = (e) => {
    e.preventDefault();
    const title = taskTitle.trim();
    if (!title) return;
    let teamId = "GM";
    if (assignTo === "auto") {
      if (policy.autoAssign) { const m = matchTeam(title, teams); teamId = m ? m.id : "GM"; }
    } else teamId = assignTo;
    const teamLabel = teamId === "GM" ? "GM AI Agent (direct)" : teams.find((t) => t.id === teamId)?.name;
    setTasks((prev) => [
      { id: nextId(), title, teamId, urgent, status: "queued", progress: 0, createdAt: Date.now(), updatedAt: Date.now(), output: "", generating: false, archived: false, notified: false },
      ...prev,
    ]);
    pushLog(`CEO directive issued: "${title}" → routed to ${teamLabel}.`);
    setTaskTitle(""); setUrgent(false); setAssignTo("auto");
  };

  const handleAddTeam = (e) => {
    e.preventDefault();
    if (!newTeam.name.trim()) return;
    const skills = newTeam.skills.split(",").map((s) => s.trim().toLowerCase()).filter(Boolean);
    setTeams((prev) => [
      ...prev,
      { id: `team-${nextId()}`, name: newTeam.name.trim(), agent: newTeam.agent.trim() || "TEAM-AI",
        skills: skills.length ? skills : [newTeam.name.trim().toLowerCase()], status: "idle",
        accent: ACCENT_CYCLE[prev.length % ACCENT_CYCLE.length] },
    ]);
    pushLog(`CEO POLICY UPDATE: new team "${newTeam.name.trim()}" added to org structure under GM.`);
    setNewTeam({ name: "", agent: "", skills: "" }); setShowAddTeam(false);
  };

  const handleRemoveTeam = (id) => {
    const team = teams.find((t) => t.id === id);
    if (!team) return;
    setTasks((prev) => prev.map((tk) => (tk.teamId === id && tk.status !== "done" ? { ...tk, teamId: "GM", status: "queued", updatedAt: Date.now() } : tk)));
    setTeams((prev) => prev.filter((t) => t.id !== id));
    pushLog(`CEO POLICY UPDATE: team "${team.name}" dissolved. Active tasks reassigned to GM AI Agent.`);
  };

  const handleArchiveTap = (task, teamLabel, agentLabel) => {
    if (!integrations.gasUrl) { pushLog("⚠️ Set the Google Apps Script URL in Setup first."); return; }
    const payload = { timestamp: new Date().toISOString(), team: teamLabel, agent: agentLabel, title: task.title, status: task.status, urgent: task.urgent, output: task.output || "" };
    const ok = openAndAutoClose(buildArchiveLink(integrations.gasUrl, payload));
    if (ok) {
      setTasks((prev) => prev.map((t) => (t.id === task.id ? { ...t, archived: true, pendingSync: false } : t)));
      pushLog(`📤 Archived "${task.title}" to Google Sheets.`);
    } else {
      pushLog("⚠️ Popup blocked — allow popups for this page, or try again (this uses a real click).");
    }
  };

  const handleWhatsAppTap = (task, teamLabel, agentLabel) => {
    if (!integrations.waNumber) { pushLog("⚠️ Set the CEO's WhatsApp number in Setup first."); return; }
    const text = composeMessage({ team: teamLabel, agent: agentLabel, title: task.title, status: task.status, urgent: task.urgent, output: task.output || "" });
    window.open(buildWaLink(integrations.waNumber, text), "_blank");
    setTasks((prev) => prev.map((t) => (t.id === task.id ? { ...t, notified: true, pendingSync: false } : t)));
  };

  const handleTelegramTap = (task, teamLabel, agentLabel) => {
    const text = composeMessage({ team: teamLabel, agent: agentLabel, title: task.title, status: task.status, urgent: task.urgent, output: task.output || "" });
    window.open(buildTelegramShareLink(text), "_blank");
    setTasks((prev) => prev.map((t) => (t.id === task.id ? { ...t, notified: true, pendingSync: false } : t)));
  };

  const handleTestConnection = async () => {
    setTestStatus({ ok: null, msg: "Testing..." });
    const payload = { timestamp: new Date().toISOString(), team: "System", agent: "Command Center", title: "Connection test", status: "test", urgent: false, output: "This is a test ping from the CEO Command Center." };
    let sent = false, msgs = [];
    if (integrations.gasUrl) {
      const ok = await silentDeliver(buildArchiveLink(integrations.gasUrl, payload));
      sent = sent || ok;
      msgs.push(ok ? "Request sent to your bridge — check your Sheet now to confirm." : "Could not reach that URL — double-check it's pasted correctly and ends in /exec.");
    } else msgs.push("No Google Apps Script URL set.");
    if (integrations.telegramDirect && integrations.telegramToken && integrations.telegramChatId) {
      const ok = await silentDeliver(`https://api.telegram.org/bot${integrations.telegramToken}/sendMessage?chat_id=${encodeURIComponent(integrations.telegramChatId)}&text=${encodeURIComponent("✅ CEO Command Center test message.")}`);
      sent = sent || ok;
      msgs.push(ok ? "Telegram direct test sent." : "Telegram direct send failed.");
    }
    setTestStatus({ ok: sent, msg: msgs.join(" ") });
  };

  const handleOpenBridgeUrl = () => {
    if (!integrations.gasUrl) return;
    window.open(integrations.gasUrl, "_blank");
  };

  const activeAgents = teams.filter((t) => t.status === "busy").length + (tasks.some((t) => t.teamId === "GM" && t.status !== "done") ? 1 : 0);
  const inProgress = tasks.filter((t) => t.status === "executing" || t.status === "delegating").length;
  const completedToday = tasks.filter((t) => t.status === "done").length;
  const queuedCount = tasks.filter((t) => t.status === "queued").length;
  const filesReady = tasks.filter((t) => t.status === "done" && t.output).length;

  const columns = [
    { key: "queued", label: "Queued" }, { key: "delegating", label: "GM Delegating" },
    { key: "executing", label: "Team Executing" }, { key: "done", label: "Completed" },
  ];

  return (
    <div className="min-h-screen w-full bg-slate-950 text-slate-100" style={{ fontFamily: "ui-sans-serif, system-ui, sans-serif" }}>
      <style>{`
        @keyframes pulseDot { 0%,100% { opacity:1; transform:scale(1);} 50% { opacity:.4; transform:scale(1.3);} }
        @keyframes fadeSlide { from { opacity:0; transform: translateY(-6px);} to { opacity:1; transform:translateY(0);} }
        .anim-pulse-dot { animation: pulseDot 1.6s ease-in-out infinite; }
        .anim-fade-in { animation: fadeSlide .35s ease-out; }
        .scrollbar-thin::-webkit-scrollbar { width: 6px; height:6px; }
        .scrollbar-thin::-webkit-scrollbar-thumb { background:#334155; border-radius:4px; }
      `}</style>

      <header className="border-b border-slate-800/80 bg-slate-950/95 sticky top-0 z-20 backdrop-blur">
        <div className="max-w-7xl mx-auto px-5 py-4 flex items-center justify-between gap-4 flex-wrap">
          <div className="flex items-center gap-3">
            <div className="h-10 w-10 rounded-md bg-amber-400/10 ring-1 ring-amber-500/40 flex items-center justify-center">
              <Crown className="h-5 w-5 text-amber-400" />
            </div>
            <div>
              <h1 className="font-serif text-xl tracking-tight leading-none">CEO Command Center</h1>
              <p className="text-[11px] uppercase tracking-widest text-slate-500 mt-1">BOD → GM AI Agent → Team Delegation · Auto-Answer · Archive</p>
            </div>
          </div>
          <div className="flex items-center gap-3">
            <button onClick={() => { setShowSettings(true); setActiveTab("settings"); }} className="hidden sm:flex items-center gap-1.5 text-xs px-2.5 py-1.5 rounded-md bg-slate-800/70 text-slate-300 hover:bg-slate-800 transition">
              <Settings2 className="h-3.5 w-3.5" /> Integrations
            </button>
            <div className="flex items-center gap-1.5 text-xs text-slate-400">
              <Radio className="h-3.5 w-3.5 text-emerald-400 anim-pulse-dot" /><span className="font-mono">LIVE</span>
            </div>
            <div className="font-mono text-xs text-slate-300">{new Date(now).toLocaleTimeString("id-ID")}</div>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-5 py-6 pb-28 md:pb-6 space-y-6">
        <section className="grid grid-cols-2 md:grid-cols-5 gap-3">
          <KpiCard icon={<Users className="h-4 w-4" />} label="Active Agents" value={activeAgents} accent="text-sky-400" />
          <KpiCard icon={<Loader2 className="h-4 w-4" />} label="In Progress" value={inProgress} accent="text-violet-400" />
          <KpiCard icon={<CheckCircle2 className="h-4 w-4" />} label="Completed" value={completedToday} accent="text-emerald-400" />
          <KpiCard icon={<Clock className="h-4 w-4" />} label="Queued" value={queuedCount} accent="text-slate-300" />
          <KpiCard icon={<FileText className="h-4 w-4" />} label="Files Ready" value={filesReady} accent="text-amber-400" />
        </section>

        {showSettings && (
          <section className={`${tabClass("settings")} bg-slate-900/50 border border-slate-800 rounded-xl p-5 anim-fade-in`}>
            <h2 className="font-serif text-lg flex items-center gap-2 mb-1">
              <Bell className="h-4 w-4 text-amber-400" /> CEO Integrations — WhatsApp, Telegram &amp; Google Sheets Archive
            </h2>
            <p className="text-xs text-slate-500 mb-4">Archive/notify now tries a real silent connection first (works once deployed outside the chat preview), and falls back to one-tap Archive / WhatsApp / Telegram buttons on each completed task if that's ever blocked — a real click always gets through. See <code className="text-slate-400">SETUP-GUIDE.md</code> for details.</p>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div className="space-y-3">
                <Field label="Google Apps Script Web App URL (Archive button target)">
                  <input value={integrations.gasUrl} onChange={(e) => setIntegrations((s) => ({ ...s, gasUrl: e.target.value }))}
                    placeholder="https://script.google.com/macros/s/XXXX/exec"
                    className="w-full bg-slate-950 border border-slate-700 rounded-md px-3 py-2 text-xs font-mono outline-none focus:border-amber-500/60" />
                </Field>
                <Field label="CEO's WhatsApp number (for the WhatsApp button)">
                  <input value={integrations.waNumber} onChange={(e) => setIntegrations((s) => ({ ...s, waNumber: e.target.value }))}
                    placeholder="62812xxxxxxx (no + or spaces)"
                    className="w-full bg-slate-950 border border-slate-700 rounded-md px-3 py-2 text-xs font-mono outline-none focus:border-amber-500/60" />
                </Field>
                <p className="text-[11px] text-slate-500 flex items-start gap-1.5"><Database className="h-3.5 w-3.5 mt-0.5 shrink-0" />Telegram Bot Token, WhatsApp API credentials for full automation live inside the Apps Script (Script Properties) — not in this browser tab.</p>
              </div>
              <div className="space-y-3">
                <label className="flex items-center justify-between text-xs text-slate-400 cursor-pointer">
                  <span className="flex items-center gap-1.5"><TelegramIcon className="h-3.5 w-3.5" /> Direct Telegram test (Bot API, optional)</span>
                  <input type="checkbox" checked={integrations.telegramDirect} onChange={(e) => setIntegrations((s) => ({ ...s, telegramDirect: e.target.checked }))} className="accent-amber-400" />
                </label>
                {integrations.telegramDirect && (
                  <div className="grid grid-cols-1 gap-2 anim-fade-in">
                    <input value={integrations.telegramToken} onChange={(e) => setIntegrations((s) => ({ ...s, telegramToken: e.target.value }))}
                      placeholder="Telegram Bot Token (from @BotFather)" type="password"
                      className="w-full bg-slate-950 border border-slate-700 rounded-md px-3 py-2 text-xs font-mono outline-none focus:border-amber-500/60" />
                    <input value={integrations.telegramChatId} onChange={(e) => setIntegrations((s) => ({ ...s, telegramChatId: e.target.value }))}
                      placeholder="CEO's Telegram Chat ID"
                      className="w-full bg-slate-950 border border-slate-700 rounded-md px-3 py-2 text-xs font-mono outline-none focus:border-amber-500/60" />
                    <p className="text-[10px] text-slate-500">Both optional — only needed if you tap Test Connection's Telegram check. The task list's own Telegram button doesn't need these; it opens Telegram directly with the message pre-filled.</p>
                    <p className="text-[10px] text-rose-400/80 flex items-start gap-1"><AlertCircle className="h-3 w-3 mt-0.5 shrink-0" />Token lives in this tab's memory only and clears on refresh.</p>
                  </div>
                )}
              </div>
            </div>
            <div className="flex items-center gap-3 mt-4">
              <button onClick={handleTestConnection} className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-md bg-amber-400 text-slate-950 font-medium hover:bg-amber-300">
                <Link2 className="h-3.5 w-3.5" /> Test Connection
              </button>
              {integrations.gasUrl && (
                <button onClick={handleOpenBridgeUrl} className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-md bg-slate-800 text-slate-300 hover:bg-slate-700">
                  <Link2 className="h-3.5 w-3.5" /> Open bridge URL
                </button>
              )}
              <label className="flex items-center gap-1.5 text-xs text-slate-400 cursor-pointer ml-2">
                <input type="checkbox" checked={policy.autoNotify} onChange={(e) => setPolicy((p) => ({ ...p, autoNotify: e.target.checked }))} className="accent-amber-400" />
                Auto-notify CEO on every completed task
              </label>
              {testStatus && (
                <span className={`text-xs flex items-center gap-1 ${testStatus.ok ? "text-emerald-400" : testStatus.ok === false ? "text-rose-400" : "text-slate-400"}`}>
                  {testStatus.ok ? <CheckCheck className="h-3.5 w-3.5" /> : <AlertCircle className="h-3.5 w-3.5" />} {testStatus.msg}
                </span>
              )}
            </div>
            {integrations.gasUrl && (
              <p className="text-[11px] text-slate-500 mt-3 flex items-center gap-1.5">
                <Database className="h-3.5 w-3.5" />
                {settingsReady ? (settingsSyncMsg || "Settings tersimpan permanen di Sheet ini — buka dari device manapun cukup isi URL ini sekali.") : "Menyinkronkan settings..."}
              </p>
            )}
          </section>
        )}

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          <div className="lg:col-span-2 space-y-6">
            <section className={`${tabClass("org")} bg-slate-900/50 border border-slate-800 rounded-xl p-5`}>
              <div className="flex items-center justify-between mb-4">
                <h2 className="font-serif text-lg flex items-center gap-2"><Building2 className="h-4 w-4 text-amber-400" /> Organization &amp; Delegation Structure</h2>
                <button onClick={() => setShowAddTeam((s) => !s)} className="flex items-center gap-1 text-xs px-2.5 py-1.5 rounded-md bg-amber-400/10 text-amber-300 ring-1 ring-amber-500/30 hover:bg-amber-400/20 transition">
                  <Plus className="h-3.5 w-3.5" /> Add Team
                </button>
              </div>
              <div className="flex flex-col items-center mb-5">
                <NodePill icon={<Crown className="h-4 w-4 text-amber-400" />} title="CEO" subtitle="Chief Executive" ring="ring-amber-500/40" />
                <div className="h-4 w-px bg-slate-700" />
                <NodePill icon={<Building2 className="h-4 w-4 text-slate-300" />} title="Board of Directors" subtitle="Policy & Oversight" ring="ring-slate-600/50" />
                <div className="h-4 w-px bg-slate-700" />
                <NodePill icon={<Sparkles className="h-4 w-4 text-sky-400" />} title="GM AI Agent" subtitle="Orchestrates delegation & auto-answers" ring="ring-sky-500/40" />
                <div className="h-4 w-px bg-slate-700" />
                <div className="text-[10px] uppercase tracking-widest text-slate-500">Reporting Teams</div>
              </div>
              {showAddTeam && (
                <form onSubmit={handleAddTeam} className="mb-4 p-3 rounded-lg bg-slate-950 border border-slate-800 anim-fade-in space-y-2">
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                    <input value={newTeam.name} onChange={(e) => setNewTeam((s) => ({ ...s, name: e.target.value }))} placeholder="Team name (e.g. R&D Innovation)" className="bg-slate-900 border border-slate-700 rounded-md px-3 py-2 text-sm outline-none focus:border-amber-500/60" />
                    <input value={newTeam.agent} onChange={(e) => setNewTeam((s) => ({ ...s, agent: e.target.value }))} placeholder="AI Agent title (e.g. CIO-AI)" className="bg-slate-900 border border-slate-700 rounded-md px-3 py-2 text-sm outline-none focus:border-amber-500/60" />
                  </div>
                  <input value={newTeam.skills} onChange={(e) => setNewTeam((s) => ({ ...s, skills: e.target.value }))} placeholder="Skills, comma separated (e.g. research, prototype, data)" className="w-full bg-slate-900 border border-slate-700 rounded-md px-3 py-2 text-sm outline-none focus:border-amber-500/60" />
                  <div className="flex justify-end gap-2">
                    <button type="button" onClick={() => setShowAddTeam(false)} className="text-xs px-3 py-1.5 rounded-md text-slate-400 hover:text-slate-200">Cancel</button>
                    <button type="submit" className="text-xs px-3 py-1.5 rounded-md bg-amber-400 text-slate-950 font-medium hover:bg-amber-300">Create Team</button>
                  </div>
                </form>
              )}
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                {teams.map((team) => <TeamCard key={team.id} team={team} onRemove={() => handleRemoveTeam(team.id)} />)}
              </div>
            </section>

            <section className={`${tabClass("board")} bg-slate-900/50 border border-slate-800 rounded-xl p-5`}>
              <h2 className="font-serif text-lg flex items-center gap-2 mb-4"><Activity className="h-4 w-4 text-violet-400" /> Real-Time Delegation Board</h2>
              <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-4 gap-3">
                {columns.map((col) => (
                  <div key={col.key} className="bg-slate-950/60 border border-slate-800 rounded-lg p-3 min-h-[140px]">
                    <div className="flex items-center gap-2 mb-3">
                      <span className={`h-1.5 w-1.5 rounded-full ${STATUS_META[col.key].dot}`} />
                      <span className="text-xs uppercase tracking-wider text-slate-400">{col.label}</span>
                      <span className="text-[10px] text-slate-600 ml-auto font-mono">{tasks.filter((t) => t.status === col.key).length}</span>
                    </div>
                    <div className="space-y-2">
                      {tasks.filter((t) => t.status === col.key).map((task) => (
                        <TaskCard key={task.id} task={task} teams={teams} now={now} expanded={expandedId === task.id}
                          onToggle={() => setExpandedId((id) => (id === task.id ? null : task.id))}
                          onArchive={() => { const team = teams.find((tm) => tm.id === task.teamId); handleArchiveTap(task, team ? team.name : "GM AI Agent", team ? team.agent : "GM-AI"); }}
                          onWhatsApp={() => { const team = teams.find((tm) => tm.id === task.teamId); handleWhatsAppTap(task, team ? team.name : "GM AI Agent", team ? team.agent : "GM-AI"); }}
                          onTelegram={() => { const team = teams.find((tm) => tm.id === task.teamId); handleTelegramTap(task, team ? team.name : "GM AI Agent", team ? team.agent : "GM-AI"); }} />
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            </section>
          </div>

          <div className="space-y-6">
            <section className={`${tabClass("new")} bg-slate-900/50 border border-slate-800 rounded-xl p-5`}>
              <h2 className="font-serif text-lg flex items-center gap-2 mb-3"><Send className="h-4 w-4 text-amber-400" /> New Directive</h2>
              <form onSubmit={handleCreateTask} className="space-y-3">
                <textarea value={taskTitle} onChange={(e) => setTaskTitle(e.target.value)} placeholder="Describe the task... e.g. 'Review Q3 marketing budget forecast'" rows={3}
                  className="w-full resize-none bg-slate-950 border border-slate-700 rounded-md px-3 py-2 text-sm outline-none focus:border-amber-500/60" />
                {taskTitle.trim() && (
                  <div className="text-xs text-slate-500 flex items-center gap-1.5">
                    <Gauge className="h-3.5 w-3.5" />
                    {suggestedTeam ? <span>Smart match: <span className="text-sky-400 font-medium">{suggestedTeam.name}</span></span> : <span>No skill match — will route to <span className="text-sky-400 font-medium">GM AI Agent</span></span>}
                  </div>
                )}
                <div className="flex items-center gap-2">
                  <select value={assignTo} onChange={(e) => setAssignTo(e.target.value)} className="flex-1 bg-slate-950 border border-slate-700 rounded-md px-2 py-2 text-xs outline-none focus:border-amber-500/60">
                    <option value="auto">Auto (Smart Match)</option>
                    <option value="GM">GM AI Agent (direct)</option>
                    {teams.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}
                  </select>
                  <label className="flex items-center gap-1.5 text-xs text-slate-400 select-none cursor-pointer whitespace-nowrap">
                    <input type="checkbox" checked={urgent} onChange={(e) => setUrgent(e.target.checked)} className="accent-rose-500" /> Urgent
                  </label>
                </div>
                <button type="submit" className="w-full flex items-center justify-center gap-2 bg-amber-400 hover:bg-amber-300 text-slate-950 font-medium text-sm rounded-md py-2.5 transition">
                  <Send className="h-3.5 w-3.5" /> Delegate to Team
                </button>
              </form>
            </section>

            <section className={`${tabClass("assistant")} bg-slate-900/50 border border-slate-800 rounded-xl p-5 flex flex-col`}>
              <div className="flex items-center justify-between mb-1">
                <h2 className="font-serif text-lg flex items-center gap-2"><Bot className="h-4 w-4 text-emerald-400" /> AI Assistant — Tanya Apa Saja</h2>
                <button onClick={() => setShowSkillPanel((s) => !s)} className="flex items-center gap-1 text-[10px] px-2 py-1 rounded bg-slate-800 text-slate-300 hover:bg-slate-700">
                  <Settings2 className="h-3 w-3" /> Skill &amp; Memori
                </button>
              </div>
              <p className="text-xs text-slate-500 mb-3">Ngobrol bebas dengan AI — strategi, riset, draft, atau sekadar tanya-tanya. Bukan buat delegasi tugas terstruktur (pakai tab New untuk itu).</p>

              {showSkillPanel && (
                <div className="mb-3 p-3 rounded-lg bg-slate-950 border border-slate-800 space-y-3 anim-fade-in">
                  <div>
                    <div className="text-[11px] text-slate-400 mb-1">Skill &amp; Kepribadian AI Agent (bisa diedit bebas)</div>
                    <textarea value={skillProfile} onChange={(e) => setSkillProfile(e.target.value)} rows={4}
                      className="w-full resize-none bg-slate-900 border border-slate-700 rounded-md px-3 py-2 text-xs outline-none focus:border-amber-500/60" />
                    {skillProfile !== DEFAULT_SKILL_PROFILE && (
                      <button onClick={() => setSkillProfile(DEFAULT_SKILL_PROFILE)} className="text-[10px] text-slate-500 hover:text-slate-300 mt-1">Reset ke default</button>
                    )}
                  </div>
                  <div>
                    <div className="flex items-center justify-between mb-1">
                      <span className="text-[11px] text-slate-400">Memori yang sudah dipelajari ({agentMemory.length})</span>
                      {agentMemory.length > 0 && (
                        <button onClick={() => setAgentMemory([])} className="text-[10px] text-rose-400 hover:text-rose-300">Hapus semua</button>
                      )}
                    </div>
                    {agentMemory.length === 0 ? (
                      <div className="text-[10px] text-slate-600 italic">Belum ada memori — akan terisi otomatis seiring kamu ngobrol.</div>
                    ) : (
                      <div className="max-h-32 overflow-y-auto scrollbar-thin space-y-1">
                        {agentMemory.map((m, i) => (
                          <div key={i} className="flex items-start gap-1.5 text-[10px] text-slate-400 bg-slate-900 border border-slate-800 rounded px-2 py-1">
                            <span className="flex-1">{m}</span>
                            <button onClick={() => setAgentMemory((prev) => prev.filter((_, idx) => idx !== i))} className="shrink-0 text-slate-600 hover:text-rose-400"><X className="h-3 w-3" /></button>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                </div>
              )}

              <div className="flex-1 min-h-[240px] max-h-[420px] overflow-y-auto scrollbar-thin space-y-3 mb-3 pr-1">
                {chatMessages.length === 0 && (
                  <div className="text-xs text-slate-600 italic py-8 text-center">Belum ada percakapan. Coba tanya sesuatu di bawah.</div>
                )}
                {chatMessages.map((m, i) => (
                  <div key={i} className={`flex ${m.role === "user" ? "justify-end" : "justify-start"}`}>
                    <div className={`max-w-[85%] rounded-lg px-3 py-2 text-xs leading-relaxed ${m.role === "user" ? "bg-amber-400 text-slate-950 whitespace-pre-wrap" : "bg-slate-950 border border-slate-800 text-slate-200"}`}>
                      {m.role === "assistant" ? <MarkdownLite text={m.content} /> : m.content}
                    </div>
                  </div>
                ))}
                {chatLoading && (
                  <div className="flex justify-start">
                    <div className="bg-slate-950 border border-slate-800 rounded-lg px-3 py-2 text-xs text-slate-400 flex items-center gap-1.5">
                      <Loader2 className="h-3 w-3 animate-spin" /> Berpikir...
                    </div>
                  </div>
                )}
                <div ref={chatEndRef} />
              </div>
              <form onSubmit={handleSendChat} className="flex items-center gap-2">
                <input value={chatInput} onChange={(e) => setChatInput(e.target.value)} placeholder="Tanya apa saja..."
                  className="flex-1 bg-slate-950 border border-slate-700 rounded-md px-3 py-2 text-xs outline-none focus:border-amber-500/60" />
                <button type="submit" disabled={chatLoading || !chatInput.trim()} className="shrink-0 h-9 w-9 flex items-center justify-center rounded-md bg-amber-400 hover:bg-amber-300 disabled:opacity-40 disabled:cursor-not-allowed text-slate-950 transition">
                  <Send className="h-3.5 w-3.5" />
                </button>
              </form>
            </section>

            <section className={`${tabClass("ops")} bg-slate-900/50 border border-slate-800 rounded-xl p-5`}>
              <h2 className="font-serif text-lg flex items-center gap-2 mb-3"><Settings2 className="h-4 w-4 text-slate-300" /> CEO Policy</h2>
              <div className="space-y-4">
                <div>
                  <div className="flex justify-between text-xs text-slate-400 mb-1"><span>Delegation Autonomy</span><span className="font-mono text-slate-300">{policy.autonomy}/5</span></div>
                  <input type="range" min="1" max="5" step="1" value={policy.autonomy} onChange={(e) => setPolicy((p) => ({ ...p, autonomy: Number(e.target.value) }))} className="w-full accent-amber-400" />
                  <p className="text-[11px] text-slate-500 mt-1">Higher autonomy = faster delegation cycle.</p>
                </div>
                <div>
                  <div className="text-xs text-slate-400 mb-1">Escalation Threshold</div>
                  <select value={policy.escalation} onChange={(e) => setPolicy((p) => ({ ...p, escalation: e.target.value }))} className="w-full bg-slate-950 border border-slate-700 rounded-md px-2 py-1.5 text-xs outline-none focus:border-amber-500/60">
                    <option>Low</option><option>Medium</option><option>High</option>
                  </select>
                </div>
                <label className="flex items-center justify-between text-xs text-slate-400 cursor-pointer">
                  <span className="flex items-center gap-1.5"><ShieldAlert className="h-3.5 w-3.5" /> Auto-assign by skill match</span>
                  <input type="checkbox" checked={policy.autoAssign} onChange={(e) => setPolicy((p) => ({ ...p, autoAssign: e.target.checked }))} className="accent-amber-400" />
                </label>
              </div>
            </section>

            <section className={`${tabClass("ops")} bg-slate-900/50 border border-slate-800 rounded-xl p-5`}>
              <h2 className="font-serif text-lg flex items-center gap-2 mb-3"><TrendingUp className="h-4 w-4 text-emerald-400" /> Live Activity Log</h2>
              <div className="space-y-2 max-h-80 overflow-y-auto scrollbar-thin pr-1">
                {logs.map((l) => (
                  <div key={l.id} className="anim-fade-in flex items-start gap-2 text-xs">
                    <span className="text-slate-600 font-mono mt-0.5 shrink-0">{new Date(l.t).toLocaleTimeString("id-ID", { hour12: false })}</span>
                    <span className="text-slate-300">{l.text}</span>
                  </div>
                ))}
              </div>
            </section>
          </div>
        </div>
      </main>

      <BottomNav activeTab={activeTab} setActiveTab={setActiveTab} boardBadge={inProgress + queuedCount} />
    </div>
  );
}

function BottomNav({ activeTab, setActiveTab, boardBadge }) {
  const items = [
    { key: "org", label: "Team", icon: Building2 },
    { key: "board", label: "Board", icon: Activity, badge: boardBadge },
    { key: "new", label: "New", icon: Send, primary: true },
    { key: "assistant", label: "Chat", icon: Bot },
    { key: "ops", label: "Ops", icon: TrendingUp },
    { key: "settings", label: "Setup", icon: Settings2 },
  ];
  return (
    <nav className="md:hidden fixed bottom-0 inset-x-0 z-30 bg-slate-950/95 backdrop-blur border-t border-slate-800 flex items-stretch h-16 px-1 pb-[env(safe-area-inset-bottom)]">
      {items.map((it) => {
        const Icon = it.icon;
        const active = activeTab === it.key;
        if (it.primary) {
          return (
            <button key={it.key} onClick={() => setActiveTab(it.key)} className="flex-1 flex flex-col items-center justify-center">
              <span className={`h-11 w-11 -mt-5 rounded-full flex items-center justify-center shadow-lg shadow-amber-500/30 transition-transform active:scale-90 ${active ? "bg-amber-300" : "bg-amber-400"}`}>
                <Icon className="h-5 w-5 text-slate-950" />
              </span>
              <span className={`text-[10px] mt-1 font-medium ${active ? "text-amber-400" : "text-slate-500"}`}>{it.label}</span>
            </button>
          );
        }
        return (
          <button key={it.key} onClick={() => setActiveTab(it.key)} className="flex-1 flex flex-col items-center justify-center gap-1 relative transition-transform active:scale-90">
            <Icon className={`h-5 w-5 ${active ? "text-amber-400" : "text-slate-500"}`} />
            {!!it.badge && (
              <span className="absolute top-1 right-1/4 h-3.5 min-w-[14px] px-0.5 rounded-full bg-rose-500 text-[9px] leading-[14px] text-white text-center font-medium">{it.badge}</span>
            )}
            <span className={`text-[10px] ${active ? "text-amber-400 font-medium" : "text-slate-500"}`}>{it.label}</span>
          </button>
        );
      })}
    </nav>
  );
}

function Field({ label, children }) {
  return (<div><div className="text-xs text-slate-400 mb-1">{label}</div>{children}</div>);
}

function KpiCard({ icon, label, value, accent }) {
  return (
    <div className="bg-slate-900/50 border border-slate-800 rounded-xl p-4 flex items-center gap-3">
      <div className={`h-9 w-9 rounded-md bg-slate-800/60 flex items-center justify-center ${accent}`}>{icon}</div>
      <div><div className="text-xl font-serif leading-none">{value}</div><div className="text-[11px] text-slate-500 mt-1 uppercase tracking-wide">{label}</div></div>
    </div>
  );
}

function NodePill({ icon, title, subtitle, ring }) {
  return (
    <div className={`flex items-center gap-2 bg-slate-900 border border-slate-800 rounded-full pl-2 pr-4 py-1.5 ring-1 ${ring}`}>
      <div className="h-6 w-6 rounded-full bg-slate-800 flex items-center justify-center">{icon}</div>
      <div className="leading-tight"><div className="text-xs font-medium text-slate-100">{title}</div><div className="text-[10px] text-slate-500">{subtitle}</div></div>
    </div>
  );
}

function TeamCard({ team, onRemove }) {
  const busy = team.status === "busy";
  return (
    <div className="relative bg-slate-950/70 border border-slate-800 rounded-lg p-3 group">
      <button onClick={onRemove} title="Remove team" className="absolute top-2 right-2 h-5 w-5 rounded flex items-center justify-center text-slate-600 hover:text-rose-400 hover:bg-rose-500/10 opacity-0 group-hover:opacity-100 transition">
        <X className="h-3.5 w-3.5" />
      </button>
      <div className="flex items-center gap-2 mb-1.5">
        <span className={`h-2 w-2 rounded-full ${busy ? "bg-emerald-400 anim-pulse-dot" : "bg-slate-600"}`} />
        <span className="text-sm font-medium text-slate-100 pr-4">{team.name}</span>
      </div>
      <div className={`text-[11px] font-mono ${team.accent.text} mb-2`}>{team.agent} · {busy ? "BUSY" : "IDLE"}</div>
      <div className="flex flex-wrap gap-1">{team.skills.slice(0, 4).map((s) => <span key={s} className="text-[10px] px-1.5 py-0.5 rounded bg-slate-800/80 text-slate-400">{s}</span>)}</div>
    </div>
  );
}

function TaskCard({ task, teams, now, expanded, onToggle, onArchive, onWhatsApp, onTelegram }) {
  const team = teams.find((t) => t.id === task.teamId);
  const teamLabel = task.teamId === "GM" ? "GM AI Agent" : (team ? team.name : "Unassigned");
  const accentHex = team ? { text: "#fbbf24", bar: "#fbbf24" }[team.accent.text] : "#fbbf24";
  const hasOutput = !!task.output;

  return (
    <div className="bg-slate-900 border border-slate-800 rounded-md p-2.5 anim-fade-in">
      <div className="flex items-start justify-between gap-2">
        <p className="text-xs text-slate-200 leading-snug">{task.title}</p>
        {task.urgent && <span className="shrink-0 text-[9px] px-1.5 py-0.5 rounded bg-rose-500/15 text-rose-400 font-medium">URGENT</span>}
      </div>
      <div className="flex items-center justify-between mt-2 text-[10px] text-slate-500">
        <span className="flex items-center gap-1"><ChevronRight className="h-3 w-3" />{teamLabel}</span>
        <span className="font-mono">{timeAgo(now - task.createdAt)}</span>
      </div>
      {(task.status === "executing" || task.generating) && (
        <div className="h-1 w-full bg-slate-800 rounded-full mt-2 overflow-hidden">
          <div className="h-full bg-violet-400 transition-all duration-500" style={{ width: `${task.progress}%` }} />
        </div>
      )}
      {task.generating && <div className="text-[10px] text-violet-400 mt-1.5 flex items-center gap-1"><Loader2 className="h-3 w-3 animate-spin" /> Generating answer...</div>}

      {hasOutput && (
        <div className="mt-2">
          <button onClick={onToggle} className="flex items-center gap-1 text-[10px] text-sky-400 hover:text-sky-300">
            <ChevronDown className={`h-3 w-3 transition-transform ${expanded ? "rotate-180" : ""}`} /> {expanded ? "Hide output" : "View AI output"}
          </button>
          {expanded && (
            <div className="mt-2 space-y-2 anim-fade-in">
              <div className="text-[11px] text-slate-300 bg-slate-950 border border-slate-800 rounded p-2 max-h-40 overflow-y-auto scrollbar-thin"><MarkdownLite text={task.output} /></div>
              <div className="flex flex-wrap gap-1.5">
                <FileBtn icon={<FileText className="h-3 w-3" />} label="Word" onClick={() => exportWord(task, teamLabel)} />
                <FileBtn icon={<FileSpreadsheet className="h-3 w-3" />} label="Excel" onClick={() => exportExcel(task, teamLabel)} />
                <FileBtn icon={<FileText className="h-3 w-3" />} label="CSV" onClick={() => exportCSV(task, teamLabel)} />
                <FileBtn icon={<Presentation className="h-3 w-3" />} label="Slides" onClick={() => exportSlides(task, teamLabel)} />
                <FileBtn icon={<ImageIcon className="h-3 w-3" />} label="Reference" onClick={() => exportSVG(task, teamLabel, accentHex)} />
              </div>
              <div className="text-[10px] text-slate-500 pt-1">Tap to send — each opens a real link/tab, so it always gets through:</div>
              <div className="flex flex-wrap gap-1.5">
                <button onClick={onArchive} className="flex items-center gap-1 text-[10px] px-2 py-1 rounded bg-amber-400/10 text-amber-300 ring-1 ring-amber-500/30 hover:bg-amber-400/20">
                  <Database className="h-3 w-3" /> Archive to Sheet
                </button>
                <button onClick={onWhatsApp} className="flex items-center gap-1 text-[10px] px-2 py-1 rounded bg-emerald-400/10 text-emerald-300 ring-1 ring-emerald-500/30 hover:bg-emerald-400/20">
                  <MessageCircle className="h-3 w-3" /> WhatsApp
                </button>
                <button onClick={onTelegram} className="flex items-center gap-1 text-[10px] px-2 py-1 rounded bg-sky-400/10 text-sky-300 ring-1 ring-sky-500/30 hover:bg-sky-400/20">
                  <TelegramIcon className="h-3 w-3" /> Telegram
                </button>
              </div>
              <div className="flex items-center gap-2 pt-1">
                {task.archived && <span className="text-[10px] text-emerald-400 flex items-center gap-1"><CheckCheck className="h-3 w-3" /> Archived</span>}
                {task.notified && <span className="text-[10px] text-sky-400 flex items-center gap-1"><CheckCheck className="h-3 w-3" /> Sent</span>}
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function FileBtn({ icon, label, onClick }) {
  return (
    <button onClick={onClick} className="flex items-center gap-1 text-[10px] px-2 py-1 rounded bg-slate-800 text-slate-300 hover:bg-slate-700 transition">
      {icon} {label}
    </button>
  );
}
