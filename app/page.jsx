"use client";

import { useState, useRef, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, AreaChart, Area, LineChart, Line, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer
} from "recharts";

// ─── STATIC BUSINESS DATA ─────────────────────────────────────────────────────
const revenueData = [
  { month: "Aug", revenue: 18400, expenses: 11200, profit: 7200 },
  { month: "Sep", revenue: 22100, expenses: 12800, profit: 9300 },
  { month: "Oct", revenue: 19800, expenses: 11900, profit: 7900 },
  { month: "Nov", revenue: 31200, expenses: 14300, profit: 16900 },
  { month: "Dec", revenue: 41800, expenses: 18200, profit: 23600 },
  { month: "Jan", revenue: 28400, expenses: 13700, profit: 14700 },
  { month: "Feb", revenue: 33900, expenses: 15400, profit: 18500 },
];
const forecastData = [
  { month: "Feb", actual: 33900, forecast: null },
  { month: "Mar", actual: null, forecast: 36500 },
  { month: "Apr", actual: null, forecast: 39200 },
  { month: "May", actual: null, forecast: 43100 },
  { month: "Jun", actual: null, forecast: 47800 },
];
const categoryData = [
  { category: "Products", value: 54 },
  { category: "Services", value: 28 },
  { category: "Subscriptions", value: 12 },
  { category: "Other", value: 6 },
];
const alerts = [
  { id: 1, type: "positive", icon: "↑", title: "Revenue spike detected", body: "Last week's revenue was 23% above your 4-week average. Shopify orders surged on Thursday." },
  { id: 2, type: "warning", icon: "⚠", title: "Expense anomaly", body: "February operating costs are tracking 14% higher than January. Supplier invoices may need review." },
  { id: 3, type: "info", icon: "◈", title: "Forecast updated", body: "Based on Q4 trends, March revenue is projected at $36,500 — your strongest Q1 month ever." },
];
const FILE_TYPES = [
  { id: "excel", label: "Excel Spreadsheet", ext: ".xlsx, .xls", icon: "X", desc: "Bank statements, expense sheets, any Excel export", accept: ".xlsx,.xls", color: "#4ADE80" },
  { id: "csv",   label: "CSV File",          ext: ".csv",        icon: "C", desc: "Exported transactions from any bank or app",        accept: ".csv",      color: "#60A5FA" },
  { id: "pdf",   label: "PDF Statement",     ext: ".pdf",        icon: "P", desc: "Bank PDF statements - text extracted automatically", accept: ".pdf",      color: "#F87171" },
];

const PIE_COLORS = ["#E8C67A","#60A5FA","#4ADE80","#F87171","#A78BFA","#FBBF24","#34D399","#FB923C","#E879F9","#22D3EE"];

// ─── CLAUDE API ───────────────────────────────────────────────────────────────
async function askClaude(messages, systemPrompt) {
  const res = await fetch("/api/chat", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens: 1500,
      system: systemPrompt,
      messages,
    }),
  });
  const data = await res.json();
  if (data.error) throw new Error(data.error.message);
  return data.content?.map(b => b.text || "").join("") || "No response.";
}

// ─── HELPERS ──────────────────────────────────────────────────────────────────
const fmtMoney = n => `$${Math.abs(n).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
const fmtShort = n => `$${Math.abs(n).toLocaleString(undefined, { maximumFractionDigits: 0 })}`;

const ChartTip = ({ active, payload, label }) => {
  if (!active || !payload?.length) return null;
  return (
    <div style={{ background: "#141820", border: "1px solid rgba(255,255,255,0.1)", borderRadius: 8, padding: "10px 14px", fontSize: 12 }}>
      <p style={{ color: "#9AA3AF", margin: "0 0 5px" }}>{label}</p>
      {payload.map((p, i) => (
        <p key={i} style={{ color: p.color, margin: "2px 0" }}>
          {p.name}: {typeof p.value === "number" && p.value > 100 ? `$${p.value.toLocaleString()}` : p.value}
        </p>
      ))}
    </div>
  );
};

// ─── EXCEL PARSER — fully dynamic ─────────────────────────────────────────────
function parseExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

        // Find header row
        const keywords = ["date","description","amount","debit","credit","balance","narration","particulars","transaction","memo","ref","details","withdrawal","deposit"];
        let headerIdx = 0;
        for (let i = 0; i < Math.min(raw.length, 20); i++) {
          const row = raw[i].map(c => String(c).toLowerCase().trim());
          const matches = row.filter(c => keywords.some(k => c.includes(k))).length;
          if (matches >= 2) { headerIdx = i; break; }
        }

        const headers = raw[headerIdx].map(h => String(h).trim());
        const rows = raw.slice(headerIdx + 1).filter(r => r.some(c => c !== "" && c !== null && c !== undefined));

        // Detect column indices
        const find = (...terms) => {
          for (const t of terms) {
            const idx = headers.findIndex(h => h.toLowerCase().includes(t));
            if (idx !== -1) return idx;
          }
          return -1;
        };

        const dateCol    = find("date", "time");
        const descCol    = find("description", "narration", "particulars", "memo", "details", "ref", "transaction");
        const amtCol     = find("amount");
        const debitCol   = find("debit", "withdrawal", "dr");
        const creditCol  = find("credit", "deposit", "cr");
        const balanceCol = find("balance", "running bal");

        // All extra columns (not the ones we already use)
        const usedCols = new Set([dateCol, descCol, amtCol, debitCol, creditCol, balanceCol].filter(i => i !== -1));
        const extraCols = headers
          .map((h, i) => ({ name: h, idx: i }))
          .filter(c => !usedCols.has(c.idx) && c.name.trim() !== "");

        // Parse each row
        const transactions = [];
        for (const row of rows) {
          // Date
          let date = "";
          const dateRaw = dateCol !== -1 ? row[dateCol] : "";
          if (dateRaw !== "") {
            if (typeof dateRaw === "number") {
              try {
                const d = XLSX.SSF.parse_date_code(dateRaw);
                date = `${d.y}-${String(d.m).padStart(2,"0")}-${String(d.d).padStart(2,"0")}`;
              } catch { date = String(dateRaw); }
            } else {
              date = String(dateRaw).trim();
            }
          }

          const desc    = descCol    !== -1 ? String(row[descCol]    || "").trim() : "";
          const balance = balanceCol !== -1 ? parseFloat(String(row[balanceCol]).replace(/[^0-9.-]/g,"")) || null : null;

          // Amount
          let amount = 0;
          if (amtCol !== -1) {
            amount = parseFloat(String(row[amtCol]).replace(/[^0-9.-]/g,"")) || 0;
          } else {
            const debit  = parseFloat(String(row[debitCol]  || "0").replace(/[^0-9.-]/g,"")) || 0;
            const credit = parseFloat(String(row[creditCol] || "0").replace(/[^0-9.-]/g,"")) || 0;
            amount = credit > 0 ? credit : -debit;
          }

          // Extra column values
          const extras = {};
          for (const col of extraCols) {
            extras[col.name] = row[col.idx] !== undefined ? String(row[col.idx]).trim() : "";
          }

          if (date || desc || amount !== 0) {
            transactions.push({
              date, desc, amount, balance,
              type: amount >= 0 ? "credit" : "debit",
              extras,
            });
          }
        }

        resolve({ transactions, headers, extraCols });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// ─── AUTO-CATEGORISE ──────────────────────────────────────────────────────────
function categorise(desc) {
  const d = (desc || "").toLowerCase();
  if (/salary|payroll|wages|income|deposit|credit|inflow/.test(d))           return "Income";
  if (/rent|mortgage|lease/.test(d))                                          return "Rent/Mortgage";
  if (/grocery|supermarket|walmart|whole food|safeway|kroger|shoprite|spar/.test(d)) return "Groceries";
  if (/uber|lyft|taxi|gas|fuel|petrol|shell|bp|chevron|transport|bus|train/.test(d)) return "Transport";
  if (/netflix|spotify|hulu|amazon prime|disney|apple|subscription|dstv/.test(d))    return "Subscriptions";
  if (/restaurant|cafe|coffee|starbucks|mcdonald|food|pizza|doordash|grubhub|eatery/.test(d)) return "Food & Dining";
  if (/electricity|water|internet|phone|bill|utility|power|airtime|data/.test(d))    return "Utilities";
  if (/hospital|pharmacy|doctor|health|medical|dental|clinic/.test(d))               return "Healthcare";
  if (/amazon|shopping|store|mall|target|bestbuy|jumia|konga/.test(d))               return "Shopping";
  if (/transfer|zelle|venmo|paypal|wire|remit/.test(d))                              return "Transfers";
  if (/atm|withdrawal|cash/.test(d))                                                 return "Cash";
  if (/insurance/.test(d))                                                           return "Insurance";
  if (/gym|fitness|sport/.test(d))                                                   return "Health & Fitness";
  if (/school|tuition|education|university|college/.test(d))                         return "Education";
  if (/hotel|airbnb|travel|flight|airline|booking/.test(d))                          return "Travel";
  return "Other";
}

// ─── ANALYTICS ────────────────────────────────────────────────────────────────
function deriveAnalytics(transactions, extraCols) {
  const credits = transactions.filter(t => t.amount > 0);
  const debits  = transactions.filter(t => t.amount < 0);
  const totalIn  = credits.reduce((s, t) => s + t.amount, 0);
  const totalOut = debits.reduce((s, t)  => s + Math.abs(t.amount), 0);
  const net      = totalIn - totalOut;

  // Category breakdown
  const catMap = {};
  for (const t of debits) {
    const cat = categorise(t.desc);
    catMap[cat] = (catMap[cat] || 0) + Math.abs(t.amount);
  }
  const categories = Object.entries(catMap)
    .map(([name, value]) => ({ name, value: parseFloat(value.toFixed(2)) }))
    .sort((a, b) => b.value - a.value);

  // Monthly flow — use first 7 chars of date (YYYY-MM) or full date if short
  const monthMap = {};
  for (const t of transactions) {
    const raw = t.date || "";
    // Try to extract YYYY-MM or MM/YYYY or any month identifier
    let key = "Unknown";
    const isoMatch = raw.match(/(\d{4}[-/]\d{2})/);
    const dmyMatch = raw.match(/(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})/);
    if (isoMatch) {
      key = isoMatch[1].replace("/","-");
    } else if (dmyMatch) {
      const [, d, m, y] = dmyMatch;
      const year = y.length === 2 ? `20${y}` : y;
      key = `${year}-${m.padStart(2,"0")}`;
    } else if (raw) {
      key = raw.slice(0, 7);
    }
    if (!monthMap[key]) monthMap[key] = { month: key, in: 0, out: 0 };
    if (t.amount > 0) monthMap[key].in  += t.amount;
    else              monthMap[key].out += Math.abs(t.amount);
  }
  const monthlyFlow = Object.values(monthMap)
    .filter(m => m.month !== "Unknown")
    .sort((a, b) => a.month.localeCompare(b.month))
    .map(m => ({ ...m, in: parseFloat(m.in.toFixed(2)), out: parseFloat(m.out.toFixed(2)) }));

  // Balance trend
  const balMap = {};
  for (const t of transactions) {
    if (t.balance != null && t.date) balMap[t.date] = t.balance;
  }
  const balanceData = Object.entries(balMap)
    .sort(([a],[b]) => a.localeCompare(b))
    .map(([date, balance]) => ({ date, balance }));

  // Top expenses
  const topExpenses = [...debits]
    .sort((a,b) => Math.abs(b.amount) - Math.abs(a.amount))
    .slice(0, 5);

  // Dynamic extra column summaries (numeric ones)
  const extraSummaries = {};
  for (const col of extraCols) {
    const nums = transactions
      .map(t => parseFloat(String(t.extras[col.name] || "").replace(/[^0-9.-]/g,"")))
      .filter(n => !isNaN(n) && n !== 0);
    if (nums.length > 0) {
      extraSummaries[col.name] = {
        total: nums.reduce((s,n) => s+n, 0),
        avg: nums.reduce((s,n) => s+n, 0) / nums.length,
        count: nums.length,
      };
    }
  }

  return { totalIn, totalOut, net, categories, monthlyFlow, balanceData, topExpenses, extraSummaries };
}

// ─── PERSONAL BANKING TAB ─────────────────────────────────────────────────────
function PersonalBankTab({ sharedData, setSharedData }) {
  const data = sharedData;
  const setData = setSharedData;
  const [parsing, setParsing]     = useState(false);
  const [parseError, setError]    = useState("");
  const [aiInsight, setAiInsight] = useState("");
  const [aiLoading, setAiLoad]    = useState(false);
  const [question, setQuestion]   = useState("");
  const [qaHistory, setQaHistory] = useState([]);
  const [qaLoading, setQaLoad]    = useState(false);
  const [txFilter, setTxFilter]   = useState("All");
  const [visibleCount, setVisible]= useState(50);
  const fileRef  = useRef(null);
  const qaBottom = useRef(null);

  useEffect(() => { qaBottom.current?.scrollIntoView({ behavior: "smooth" }); }, [qaHistory]);

  const processFile = async (file) => {
    if (!file) return;
    const ext = file.name.split(".").pop().toLowerCase();
    if (!["xlsx","xls","csv","pdf"].includes(ext)) {
      setError("Please upload an Excel (.xlsx, .xls), CSV, or PDF file."); return;
    }
    setParsing(true); setError(""); setData(null); setAiInsight(""); setQaHistory([]);
    try {
      const { transactions, headers, extraCols } = await parseExcel(file);
      if (transactions.length === 0) throw new Error("No transactions found.");
      const withCats = transactions.map(t => ({ ...t, category: categorise(t.desc) }));
      const analytics = deriveAnalytics(withCats, extraCols);
      const parsed = { transactions: withCats, analytics, extraCols, fileName: file.name, rawHeaders: headers };
      setData(parsed);
      autoInsight(withCats, analytics, extraCols, file.name);
    } catch (err) {
      setError(`Could not parse file: ${err.message}. Make sure it has Date, Description, and Amount columns.`);
    }
    setParsing(false);
  };

  const autoInsight = async (txns, anal, extraCols, fileName) => {
    setAiLoad(true);
    const extraInfo = Object.entries(anal.extraSummaries || {})
      .map(([k,v]) => `${k}: total ${v.total.toFixed(2)}, avg ${v.avg.toFixed(2)}`)
      .join("; ");
    const summary = `
File: ${fileName}
Transactions: ${txns.length}
Total Credits: $${anal.totalIn.toFixed(2)}
Total Debits: $${anal.totalOut.toFixed(2)}
Net: $${anal.net.toFixed(2)}
Top spend categories: ${anal.categories.slice(0,6).map(c=>`${c.name} $${c.value}`).join(", ")}
Top 5 expenses: ${anal.topExpenses.map(t=>`${t.desc} $${Math.abs(t.amount).toFixed(2)}`).join(", ")}
Monthly: ${anal.monthlyFlow.map(m=>`${m.month} in:$${m.in} out:$${m.out}`).join(" | ")}
${extraInfo ? `Extra data found: ${extraInfo}` : ""}
    `.trim();
    const sys = `You are a warm and insightful personal finance advisor. Analyse this bank statement summary and write 4-6 sentences covering: overall cash flow health, the biggest spending categories to watch, any positive patterns, and 1-2 specific actionable tips. Reference the person's actual numbers. Write as flowing prose, no bullet points.`;
    try {
      const reply = await askClaude([{ role:"user", content:`Analyse my bank statement:\n\n${summary}` }], sys);
      setAiInsight(reply);
    } catch(e) {
      setAiInsight(`AI insight unavailable: ${e.message}. Check that your API key is set in .env.local.`);
    }
    setAiLoad(false);
  };

  const askQuestion = async () => {
    if (!question.trim() || qaLoading || !data) return;
    const q = question.trim();
    const newHistory = [...qaHistory, { role:"user", content:q }];
    setQaHistory(newHistory); setQuestion(""); setQaLoad(true);
    const { analytics, extraCols } = data;
    const extraInfo = Object.entries(analytics.extraSummaries || {})
      .map(([k,v]) => `${k}: total $${v.total.toFixed(2)}, avg $${v.avg.toFixed(2)}`)
      .join("; ");
    const ctx = `
Total Credits: $${analytics.totalIn.toFixed(2)}, Debits: $${analytics.totalOut.toFixed(2)}, Net: $${analytics.net.toFixed(2)}
Categories: ${analytics.categories.map(c=>`${c.name}: $${c.value}`).join(", ")}
Top expenses: ${analytics.topExpenses.map(t=>`${t.desc}: $${Math.abs(t.amount).toFixed(2)}`).join(", ")}
Monthly: ${analytics.monthlyFlow.map(m=>`${m.month} in:$${m.in} out:$${m.out}`).join(" | ")}
${extraInfo ? `Extra columns: ${extraInfo}` : ""}
Extra columns available: ${extraCols.map(c=>c.name).join(", ")}
    `.trim();
    const sys = `You are a personal finance assistant. Bank statement data:\n${ctx}\nAnswer helpfully and specifically using exact numbers. Keep answers to 2-4 sentences.`;
    const apiMsgs = newHistory.map(m => ({ role: m.role, content: m.content }));
    try {
      const reply = await askClaude(apiMsgs, sys);
      setQaHistory(prev => [...prev, { role:"assistant", content:reply }]);
    } catch(e) {
      setQaHistory(prev => [...prev, { role:"assistant", content:`Error: ${e.message}` }]);
    }
    setQaLoad(false);
  };

  const onDrop = useCallback((e) => {
    e.preventDefault(); setDragging(false);
    const file = e.dataTransfer.files[0];
    if (file) processFile(file);
  }, []);

  const card = { background:"rgba(255,255,255,0.03)", border:"1px solid rgba(255,255,255,0.07)", borderRadius:14, padding:"22px" };

  const allCats = data ? ["All", "Credits", "Debits", ...new Set(data.transactions.map(t=>t.category))] : [];
  const filteredTxns = data ? data.transactions.filter(t => {
    if (txFilter === "All")     return true;
    if (txFilter === "Credits") return t.type === "credit";
    if (txFilter === "Debits")  return t.type === "debit";
    return t.category === txFilter;
  }) : [];

  return (
    <div>
      <div style={{ marginBottom:22 }}>
        <h2 style={{ fontSize:18, fontFamily:"'DM Serif Display',serif", color:"#F0EEE8", margin:"0 0 4px" }}>Upload</h2>
        <p style={{ fontSize:12, color:"#6B7A8D", margin:"0 0 20px" }}>Choose a file type to upload — Claude will analyse your statement instantly</p>

        {/* File type options */}
        <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:14, marginBottom:24 }}>
          {[
            { label:"Excel", ext:".xlsx / .xls", icon:"📊", desc:"Bank exports from any bank app", color:"#4ADE80", accept:".xlsx,.xls" },
            { label:"CSV",   ext:".csv",         icon:"📋", desc:"Exported transactions CSV file", color:"#60A5FA", accept:".csv"      },
            { label:"PDF",   ext:".pdf",         icon:"📕", desc:"PDF bank statements",            color:"#F87171", accept:".pdf"      },
          ].map(ft => (
            <div key={ft.label}
              onClick={() => { if(fileRef.current){ fileRef.current.accept = ft.accept; fileRef.current.click(); } }}
              style={{ background:`${ft.color}0D`, border:`1px solid ${ft.color}33`, borderRadius:12, padding:"20px", cursor:"pointer", textAlign:"center", transition:"all 0.2s" }}
              onMouseEnter={e => e.currentTarget.style.background = `${ft.color}1A`}
              onMouseLeave={e => e.currentTarget.style.background = `${ft.color}0D`}
            >
              <div style={{ fontSize:30, marginBottom:10 }}>{ft.icon}</div>
              <p style={{ fontSize:14, fontWeight:600, color:"#F0EEE8", margin:"0 0 4px" }}>{ft.label}</p>
              <p style={{ fontSize:11, color:"#6B7A8D", margin:"0 0 10px" }}>{ft.desc}</p>
              <span style={{ fontSize:10, color:ft.color, fontFamily:"'DM Mono',monospace", background:`${ft.color}18`, padding:"3px 10px", borderRadius:99 }}>{ft.ext}</span>
            </div>
          ))}
        </div>
      </div>

      {/* ── HIDDEN FILE INPUT + PARSING STATE ── */}
      {!data && (
        <div>
          <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv,.pdf" style={{ display:"none" }} onChange={e=>processFile(e.target.files[0])} />
          {parsing && (
            <div style={{ ...card, textAlign:"center", padding:"32px", marginBottom:16 }}>
              <div style={{ display:"flex", gap:5, justifyContent:"center", marginBottom:12 }}>
                {[0,1,2].map(i=><div key={i} className="dp" style={{ width:8, height:8, borderRadius:"50%", background:"#E8C67A", animationDelay:`${i*0.22}s` }} />)}
              </div>
              <p style={{ color:"#E8C67A", fontSize:14 }}>Reading your file…</p>
            </div>
          )}
          {parseError && <p style={{ color:"#F87171", fontSize:13, marginBottom:16, padding:"12px 16px", background:"rgba(248,113,113,0.08)", borderRadius:8, border:"1px solid rgba(248,113,113,0.2)" }}>{parseError}</p>}
        </div>
      )}

      {/* ── RESULTS ── */}
      {data && (
        <div>
          {/* File bar */}
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:20 }}>
            <div style={{ display:"flex", alignItems:"center", gap:10 }}>
              <span style={{ fontSize:20 }}>📄</span>
              <div>
                <p style={{ color:"#F0EEE8", fontSize:14, fontWeight:500, margin:0 }}>{data.fileName}</p>
                <p style={{ color:"#6B7A8D", fontSize:11, margin:0 }}>
                  {data.transactions.length} transactions · Columns detected: {data.rawHeaders.filter(h=>h).join(", ")}
                </p>
              </div>
            </div>
            <button onClick={()=>{setData(null);setAiInsight("");setQaHistory([]);setTxFilter("All");}}
              style={{ background:"rgba(248,113,113,0.1)", border:"1px solid rgba(248,113,113,0.3)", borderRadius:8, padding:"6px 14px", fontSize:12, color:"#F87171", cursor:"pointer" }}>
              ✕ Remove
            </button>
          </div>

          {/* Stat Cards */}
          <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:14, marginBottom:20 }}>
            {[
              { label:"Total Income",   value:fmtShort(data.analytics.totalIn),  sub:`${data.transactions.filter(t=>t.type==="credit").length} credit transactions`, color:"#4ADE80" },
              { label:"Total Spending", value:fmtShort(data.analytics.totalOut), sub:`${data.transactions.filter(t=>t.type==="debit").length} debit transactions`,  color:"#F87171" },
              { label:"Net Cash Flow",  value:fmtShort(Math.abs(data.analytics.net)), sub: data.analytics.net >= 0 ? "▲ Positive balance" : "▼ Spending exceeds income", color: data.analytics.net>=0?"#4ADE80":"#F87171" },
              { label:"Months Covered", value: data.analytics.monthlyFlow.length || "—", sub: data.analytics.monthlyFlow.length ? `${data.analytics.monthlyFlow[0]?.month} → ${data.analytics.monthlyFlow[data.analytics.monthlyFlow.length-1]?.month}` : "Single period", color:"#E8C67A" },
            ].map(s => (
              <div key={s.label} style={{ background:"rgba(255,255,255,0.03)", border:"1px solid rgba(255,255,255,0.07)", borderRadius:12, padding:"18px 20px" }}>
                <p style={{ fontSize:10, color:"#6B7A8D", textTransform:"uppercase", letterSpacing:"0.12em", margin:"0 0 8px" }}>{s.label}</p>
                <p style={{ fontSize:24, fontFamily:"'DM Mono',monospace", color:s.color, margin:"0 0 4px" }}>{s.value}</p>
                <p style={{ fontSize:11, color:"#6B7A8D", margin:0 }}>{s.sub}</p>
              </div>
            ))}
          </div>

          {/* Charts Row */}
          <div style={{ display:"grid", gridTemplateColumns: data.analytics.monthlyFlow.length > 1 ? "3fr 2fr" : "1fr", gap:18, marginBottom:18 }}>
            {data.analytics.monthlyFlow.length > 1 && (
              <div style={card}>
                <p style={{ fontSize:14, fontWeight:500, color:"#F0EEE8", margin:"0 0 2px" }}>Monthly Cash Flow</p>
                <p style={{ fontSize:11, color:"#6B7A8D", margin:"0 0 18px" }}>Income vs Spending per month</p>
                <ResponsiveContainer width="100%" height={190}>
                  <BarChart data={data.analytics.monthlyFlow} margin={{ top:5, right:5, left:-20, bottom:0 }} barSize={12}>
                    <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.04)" vertical={false} />
                    <XAxis dataKey="month" tick={{ fontSize:10, fill:"#6B7A8D" }} axisLine={false} tickLine={false} />
                    <YAxis tick={{ fontSize:10, fill:"#6B7A8D" }} axisLine={false} tickLine={false} tickFormatter={v=>`$${(v/1000).toFixed(0)}k`} />
                    <Tooltip content={<ChartTip />} />
                    <Bar dataKey="in"  fill="#4ADE80" radius={[4,4,0,0]} name="Income" />
                    <Bar dataKey="out" fill="#F87171" radius={[4,4,0,0]} name="Spending" />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            )}

            {data.analytics.categories.length > 0 && (
              <div style={card}>
                <p style={{ fontSize:14, fontWeight:500, color:"#F0EEE8", margin:"0 0 2px" }}>Spending by Category</p>
                <p style={{ fontSize:11, color:"#6B7A8D", margin:"0 0 12px" }}>Auto-detected from descriptions</p>
                <ResponsiveContainer width="100%" height={120}>
                  <PieChart>
                    <Pie data={data.analytics.categories.slice(0,8)} cx="50%" cy="50%" innerRadius={32} outerRadius={52} dataKey="value" paddingAngle={3}>
                      {data.analytics.categories.slice(0,8).map((_,i) => <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />)}
                    </Pie>
                    <Tooltip formatter={v=>[`$${Number(v).toLocaleString()}`,"Spending"]} contentStyle={{ background:"#141820", border:"1px solid rgba(255,255,255,0.1)", borderRadius:8, fontSize:12 }} />
                  </PieChart>
                </ResponsiveContainer>
                <div style={{ display:"flex", flexDirection:"column", gap:5, marginTop:8 }}>
                  {data.analytics.categories.slice(0,7).map((c,i) => (
                    <div key={c.name} style={{ display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                      <div style={{ display:"flex", alignItems:"center", gap:6 }}>
                        <div style={{ width:7, height:7, borderRadius:"50%", background:PIE_COLORS[i % PIE_COLORS.length], flexShrink:0 }} />
                        <span style={{ fontSize:11, color:"#8A939F" }}>{c.name}</span>
                      </div>
                      <span style={{ fontSize:11, fontFamily:"'DM Mono',monospace", color:"#D8D4CC" }}>${c.value.toLocaleString()}</span>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>

          {/* Balance Trend */}
          {data.analytics.balanceData.length > 3 && (
            <div style={{ ...card, marginBottom:18 }}>
              <p style={{ fontSize:14, fontWeight:500, color:"#F0EEE8", margin:"0 0 2px" }}>Account Balance Over Time</p>
              <p style={{ fontSize:11, color:"#6B7A8D", margin:"0 0 16px" }}>Running balance from your statement</p>
              <ResponsiveContainer width="100%" height={140}>
                <AreaChart data={data.analytics.balanceData} margin={{ top:5, right:8, left:-12, bottom:0 }}>
                  <defs>
                    <linearGradient id="balGrad" x1="0" y1="0" x2="0" y2="1">
                      <stop offset="5%" stopColor="#E8C67A" stopOpacity={0.25} />
                      <stop offset="95%" stopColor="#E8C67A" stopOpacity={0} />
                    </linearGradient>
                  </defs>
                  <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.04)" />
                  <XAxis dataKey="date" tick={{ fontSize:9, fill:"#6B7A8D" }} axisLine={false} tickLine={false} />
                  <YAxis tick={{ fontSize:10, fill:"#6B7A8D" }} axisLine={false} tickLine={false} tickFormatter={v=>`$${(v/1000).toFixed(0)}k`} />
                  <Tooltip contentStyle={{ background:"#141820", border:"1px solid rgba(255,255,255,0.1)", borderRadius:8, fontSize:12 }} formatter={v=>[`$${Number(v).toLocaleString()}`,"Balance"]} />
                  <Area type="monotone" dataKey="balance" stroke="#E8C67A" strokeWidth={2} fill="url(#balGrad)" />
                </AreaChart>
              </ResponsiveContainer>
            </div>
          )}

          {/* Extra columns summary (dynamic) */}
          {Object.keys(data.analytics.extraSummaries || {}).length > 0 && (
            <div style={{ ...card, marginBottom:18 }}>
              <p style={{ fontSize:14, fontWeight:500, color:"#F0EEE8", margin:"0 0 2px" }}>Additional Data from Your File</p>
              <p style={{ fontSize:11, color:"#6B7A8D", margin:"0 0 16px" }}>Extra columns detected and summarised</p>
              <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(180px,1fr))", gap:12 }}>
                {Object.entries(data.analytics.extraSummaries).map(([name,v],i) => (
                  <div key={name} style={{ background:"rgba(255,255,255,0.03)", borderRadius:10, padding:"14px 16px", border:"1px solid rgba(255,255,255,0.06)" }}>
                    <p style={{ fontSize:10, color:"#6B7A8D", textTransform:"uppercase", letterSpacing:"0.1em", margin:"0 0 6px" }}>{name}</p>
                    <p style={{ fontSize:18, fontFamily:"'DM Mono',monospace", color:PIE_COLORS[i % PIE_COLORS.length], margin:"0 0 2px" }}>${v.total.toLocaleString(undefined,{maximumFractionDigits:0})}</p>
                    <p style={{ fontSize:11, color:"#6B7A8D", margin:0 }}>avg ${v.avg.toLocaleString(undefined,{maximumFractionDigits:0})} · {v.count} rows</p>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* AI Insight */}
          <div style={{ ...card, marginBottom:18, borderColor:"rgba(232,198,122,0.2)", background:"rgba(232,198,122,0.03)" }}>
            <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:12 }}>
              <div style={{ width:30, height:30, borderRadius:8, background:"rgba(232,198,122,0.18)", display:"flex", alignItems:"center", justifyContent:"center", fontSize:14, color:"#E8C67A" }}>✦</div>
              <p style={{ fontSize:14, fontWeight:500, color:"#E8C67A", margin:0 }}>AI Financial Insight</p>
              {aiLoading && <span style={{ fontSize:11, color:"#6B7A8D" }}>Analysing…</span>}
            </div>
            {aiLoading
              ? <div style={{ display:"flex", gap:5 }}>{[0,1,2].map(i=><div key={i} className="dp" style={{ width:7, height:7, borderRadius:"50%", background:"#E8C67A", animationDelay:`${i*0.22}s` }} />)}</div>
              : <p style={{ fontSize:14, color:"#C9C3B8", lineHeight:1.75, margin:0 }}>{aiInsight || "Add your ANTHROPIC_API_KEY to .env.local to enable AI insights."}</p>
            }
          </div>

          {/* Ask AI */}
          <div style={{ ...card, marginBottom:18 }}>
            <p style={{ fontSize:14, fontWeight:500, color:"#F0EEE8", margin:"0 0 4px" }}>Ask About Your Statement</p>
            <p style={{ fontSize:11, color:"#6B7A8D", margin:"0 0 14px" }}>Claude has full context of your uploaded data including all detected columns</p>
            <div style={{ display:"flex", gap:7, flexWrap:"wrap", marginBottom:14 }}>
              {["What's my biggest expense?","Am I saving money?","What can I cut back on?","How consistent is my income?"].map(s=>(
                <button key={s} onClick={()=>setQuestion(s)} style={{ background:"rgba(232,198,122,0.08)", border:"1px solid rgba(232,198,122,0.2)", borderRadius:99, padding:"5px 12px", fontSize:12, color:"#C9A84C", cursor:"pointer" }}>{s}</button>
              ))}
            </div>
            {qaHistory.length > 0 && (
              <div style={{ maxHeight:240, overflowY:"auto", display:"flex", flexDirection:"column", gap:10, marginBottom:14, paddingRight:4 }}>
                {qaHistory.map((m,i)=>(
                  <div key={i} style={{ display:"flex", flexDirection:m.role==="user"?"row-reverse":"row", gap:8, alignItems:"flex-start" }}>
                    <div style={{ width:26, height:26, borderRadius:6, flexShrink:0, background:m.role==="user"?"rgba(99,91,255,0.2)":"rgba(232,198,122,0.15)", display:"flex", alignItems:"center", justifyContent:"center", fontSize:10, color:m.role==="user"?"#A5A0FF":"#E8C67A" }}>{m.role==="user"?"U":"AI"}</div>
                    <div style={{ maxWidth:"80%", background:m.role==="user"?"rgba(99,91,255,0.1)":"rgba(255,255,255,0.04)", border:`1px solid ${m.role==="user"?"rgba(99,91,255,0.2)":"rgba(255,255,255,0.07)"}`, borderRadius:10, padding:"10px 14px", fontSize:13, lineHeight:1.65, color:"#D8D4CC", whiteSpace:"pre-wrap" }}>{m.content}</div>
                  </div>
                ))}
                {qaLoading && (
                  <div style={{ display:"flex", gap:8 }}>
                    <div style={{ width:26, height:26, borderRadius:6, background:"rgba(232,198,122,0.15)", display:"flex", alignItems:"center", justifyContent:"center", fontSize:10, color:"#E8C67A" }}>AI</div>
                    <div style={{ background:"rgba(255,255,255,0.04)", border:"1px solid rgba(255,255,255,0.07)", borderRadius:10, padding:"12px 14px", display:"flex", gap:5 }}>
                      {[0,1,2].map(i=><div key={i} className="dp" style={{ width:6, height:6, borderRadius:"50%", background:"#E8C67A", animationDelay:`${i*0.22}s` }} />)}
                    </div>
                  </div>
                )}
                <div ref={qaBottom} />
              </div>
            )}
            <div style={{ display:"flex", gap:10 }}>
              <input value={question} onChange={e=>setQuestion(e.target.value)} onKeyDown={e=>e.key==="Enter"&&askQuestion()}
                placeholder="Ask anything about your transactions…"
                style={{ flex:1, background:"rgba(255,255,255,0.05)", border:"1px solid rgba(255,255,255,0.1)", borderRadius:10, padding:"11px 16px", fontSize:14, color:"#F0EEE8", fontFamily:"inherit" }} />
              <button onClick={askQuestion} disabled={qaLoading||!question.trim()} style={{ background:qaLoading||!question.trim()?"rgba(232,198,122,0.15)":"linear-gradient(135deg,#E8C67A,#C9963A)", border:"none", borderRadius:10, padding:"11px 20px", fontSize:14, color:qaLoading||!question.trim()?"#6B7A8D":"#0C0F14", cursor:"pointer", fontWeight:600, fontFamily:"inherit" }}>Ask →</button>
            </div>
          </div>

          {/* Transaction Table */}
          <div style={card}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:16, flexWrap:"wrap", gap:10 }}>
              <div>
                <p style={{ fontSize:14, fontWeight:500, color:"#F0EEE8", margin:"0 0 2px" }}>All Transactions</p>
                <p style={{ fontSize:11, color:"#6B7A8D", margin:0 }}>{filteredTxns.length} records</p>
              </div>
              <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
                {allCats.slice(0,9).map(c=>(
                  <button key={c} onClick={()=>{setTxFilter(c);setVisible(50);}} style={{ padding:"4px 10px", borderRadius:99, fontSize:11, cursor:"pointer", border:"none", fontFamily:"inherit", background:txFilter===c?"rgba(232,198,122,0.2)":"rgba(255,255,255,0.05)", color:txFilter===c?"#E8C67A":"#6B7A8D", outline:txFilter===c?"1px solid rgba(232,198,122,0.4)":"1px solid transparent" }}>{c}</button>
                ))}
              </div>
            </div>

            {/* Dynamic headers — show extra cols if present */}
            <div style={{ borderTop:"1px solid rgba(255,255,255,0.05)" }}>
              <div style={{ display:"grid", gridTemplateColumns:`100px 1fr 110px ${data.extraCols.slice(0,2).map(()=>"100px").join(" ")} 100px`, gap:10, padding:"8px 10px" }}>
                {["Date","Description","Category",...data.extraCols.slice(0,2).map(c=>c.name),"Amount"].map(h=>(
                  <span key={h} style={{ fontSize:10, color:"#6B7A8D", textTransform:"uppercase", letterSpacing:"0.1em", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{h}</span>
                ))}
              </div>
              <div style={{ maxHeight:360, overflowY:"auto" }}>
                {filteredTxns.slice(0, visibleCount).map((tx,i)=>(
                  <div key={i} style={{ display:"grid", gridTemplateColumns:`100px 1fr 110px ${data.extraCols.slice(0,2).map(()=>"100px").join(" ")} 100px`, gap:10, padding:"10px 10px", borderRadius:7, alignItems:"center", background:i%2===0?"rgba(255,255,255,0.015)":"transparent" }}>
                    <span style={{ fontSize:11, color:"#6B7A8D", fontFamily:"'DM Mono',monospace" }}>{tx.date||"—"}</span>
                    <span style={{ fontSize:13, color:"#D8D4CC", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{tx.desc||"—"}</span>
                    <span style={{ fontSize:10, padding:"3px 8px", borderRadius:99, background:"rgba(255,255,255,0.05)", color:"#9AA3AF", width:"fit-content", whiteSpace:"nowrap" }}>{tx.category}</span>
                    {data.extraCols.slice(0,2).map(col=>(
                      <span key={col.name} style={{ fontSize:12, color:"#8A939F", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{tx.extras[col.name]||"—"}</span>
                    ))}
                    <span style={{ fontSize:13, fontFamily:"'DM Mono',monospace", fontWeight:500, color:tx.type==="credit"?"#4ADE80":"#F87171", textAlign:"right" }}>
                      {tx.type==="credit"?"+":"−"}{fmtMoney(tx.amount)}
                    </span>
                  </div>
                ))}
                {filteredTxns.length > visibleCount && (
                  <div style={{ textAlign:"center", padding:"14px 0" }}>
                    <button onClick={()=>setVisible(v=>v+50)} style={{ background:"rgba(255,255,255,0.05)", border:"1px solid rgba(255,255,255,0.1)", borderRadius:8, padding:"7px 20px", fontSize:12, color:"#9AA3AF", cursor:"pointer" }}>
                      Load more ({filteredTxns.length - visibleCount} remaining)
                    </button>
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}


// ─── SOURCES TAB ──────────────────────────────────────────────────────────────
function SourcesTab({ onNavigate }) {
  const fileRefs = { excel: useRef(null), csv: useRef(null), pdf: useRef(null) };
  const [uploaded, setUploaded] = useState({});
  const [pdfText, setPdfText] = useState("");
  const [pdfName, setPdfName] = useState("");
  const [pdfLoading, setPdfLoading] = useState(false);

  const handleFile = async (file, type) => {
    if (!file) return;
    if (type === "pdf") {
      setPdfLoading(true); setPdfName(file.name);
      // Read PDF as text using FileReader — extract raw text
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          // Convert PDF bytes to text via simple extraction
          const bytes = new Uint8Array(e.target.result);
          let text = "";
          for (let i = 0; i < bytes.length; i++) {
            if (bytes[i] >= 32 && bytes[i] < 127) text += String.fromCharCode(bytes[i]);
            else if (bytes[i] === 10 || bytes[i] === 13) text += "\n";
          }
          // Clean up extracted text
          text = text.replace(/[^\x20-\x7E\n]/g, " ").replace(/ {3,}/g, "  ").trim();
          setPdfText(text.slice(0, 8000)); // first 8000 chars
        } catch { setPdfText("Could not extract text from this PDF."); }
        setPdfLoading(false);
      };
      reader.readAsArrayBuffer(file);
    } else {
      setUploaded(prev => ({ ...prev, [type]: file.name }));
      // Switch to banking tab with file
      onNavigate("banking");
    }
  };

  const card = { background:"rgba(255,255,255,0.03)", border:"1px solid rgba(255,255,255,0.07)", borderRadius:14, padding:"22px" };

  return (
    <div>
      <div style={{ marginBottom:22 }}>
        <h2 style={{ fontSize:18, fontFamily:"'DM Serif Display',serif", color:"#F0EEE8", marginBottom:4 }}>Upload Your Data</h2>
        <p style={{ fontSize:12, color:"#6B7A8D" }}>Upload CSV, Excel, or PDF bank statements — Claude analyses them instantly</p>
      </div>

      {/* File type cards */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:18, marginBottom:24 }}>
        {FILE_TYPES.map(ft => (
          <div key={ft.id} style={{ ...card, borderColor: uploaded[ft.id] ? ft.color+"44" : "rgba(255,255,255,0.07)", background: uploaded[ft.id] ? ft.color+"08" : "rgba(255,255,255,0.03)", cursor:"pointer", transition:"all 0.2s" }}
            onClick={() => fileRefs[ft.id]?.current?.click()}
            onMouseEnter={e => e.currentTarget.style.borderColor = ft.color+"66"}
            onMouseLeave={e => e.currentTarget.style.borderColor = uploaded[ft.id] ? ft.color+"44" : "rgba(255,255,255,0.07)"}
          >
            <input ref={fileRefs[ft.id]} type="file" accept={ft.accept} style={{ display:"none" }} onChange={e => handleFile(e.target.files[0], ft.id)} />
            <div style={{ width:44, height:44, borderRadius:10, background: ft.color+"18", display:"flex", alignItems:"center", justifyContent:"center", fontSize:22, marginBottom:14 }}>{ft.icon}</div>
            <p style={{ fontSize:15, fontWeight:600, color:"#F0EEE8", marginBottom:6 }}>{ft.label}</p>
            <p style={{ fontSize:11, color:"#6B7A8D", marginBottom:14, lineHeight:1.5 }}>{ft.desc}</p>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center" }}>
              <span style={{ fontSize:10, color:"#4B5563", fontFamily:"'DM Mono',monospace", background:"rgba(255,255,255,0.05)", padding:"3px 8px", borderRadius:99 }}>{ft.ext}</span>
              {uploaded[ft.id]
                ? <span style={{ fontSize:11, color:"#4ADE80" }}>✓ Uploaded</span>
                : <span style={{ fontSize:11, color: ft.color, fontWeight:500 }}>Click to upload →</span>
              }
            </div>
            {uploaded[ft.id] && <p style={{ fontSize:11, color:"#6B7A8D", marginTop:8, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>📄 {uploaded[ft.id]}</p>}
          </div>
        ))}
      </div>

      {/* How it works */}
      <div style={{ ...card, marginBottom:24 }}>
        <p style={{ fontSize:14, fontWeight:500, color:"#F0EEE8", marginBottom:16 }}>How it works</p>
        <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:16 }}>
          {[
            { step:"01", title:"Export from your bank", desc:"Download your statement as Excel, CSV, or PDF from your bank's website or mobile app" },
            { step:"02", title:"Upload here",           desc:"Drag & drop or click to upload — works with any bank worldwide including Nigerian banks" },
            { step:"03", title:"Get AI insights",       desc:"Claude reads your data, categorises transactions, and gives you personalised financial insights" },
          ].map(s => (
            <div key={s.step} style={{ display:"flex", gap:14, alignItems:"flex-start" }}>
              <div style={{ width:32, height:32, borderRadius:8, background:"rgba(232,198,122,0.12)", display:"flex", alignItems:"center", justifyContent:"center", fontSize:11, fontFamily:"'DM Mono',monospace", color:"#E8C67A", flexShrink:0 }}>{s.step}</div>
              <div>
                <p style={{ fontSize:13, fontWeight:500, color:"#F0EEE8", marginBottom:4 }}>{s.title}</p>
                <p style={{ fontSize:12, color:"#6B7A8D", lineHeight:1.55 }}>{s.desc}</p>
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* Supported banks note */}
      <div style={{ ...card, borderColor:"rgba(96,165,250,0.2)", background:"rgba(96,165,250,0.04)" }}>
        <p style={{ fontSize:13, color:"#60A5FA", fontWeight:500, marginBottom:8 }}>Works with all banks</p>
        <p style={{ fontSize:13, color:"#8A939F", lineHeight:1.65 }}>
          Including Nigerian banks (GTBank, Access, Zenith, UBA, First Bank, Opay, Kuda) and international banks (Chase, Barclays, HSBC, Standard Chartered). 
          Just export your statement from your bank app or internet banking portal and upload it here.
        </p>
      </div>

      {/* PDF preview if uploaded */}
      {pdfName && (
        <div style={{ ...card, marginTop:24 }}>
          <p style={{ fontSize:14, fontWeight:500, color:"#F0EEE8", marginBottom:4 }}>📕 {pdfName}</p>
          {pdfLoading
            ? <p style={{ fontSize:13, color:"#6B7A8D" }}>Extracting text…</p>
            : <div>
                <p style={{ fontSize:11, color:"#6B7A8D", marginBottom:12 }}>Text extracted — Claude can now analyse this. Go to the <strong style={{ color:"#E8C67A", cursor:"pointer" }} onClick={()=>onNavigate("ask")}>Ask AI tab</strong> to ask questions about it.</p>
                <pre style={{ fontSize:11, color:"#8A939F", background:"rgba(255,255,255,0.03)", borderRadius:8, padding:14, overflowX:"auto", maxHeight:200, overflowY:"auto", lineHeight:1.6, whiteSpace:"pre-wrap" }}>{pdfText || "No readable text found in this PDF."}</pre>
              </div>
          }
        </div>
      )}
    </div>
  );
}

// ─── MAIN APP ─────────────────────────────────────────────────────────────────
export default function BizLens() {
  const [tab, setTab]           = useState("banking");
  const [sharedData, setSharedData] = useState(null); // uploaded file data shared across tabs
  const [msgs, setMsgs]         = useState([{ role:"assistant", content:"Hi! I'm BizLens AI. Start by uploading your bank statement in the Upload tab — then come back here to ask me anything about your data!" }]);
  const [input, setInput]       = useState("");
  const [loading, setLoad]      = useState(false);
  const bottomRef               = useRef(null);

  useEffect(()=>{ bottomRef.current?.scrollIntoView({ behavior:"smooth" }); },[msgs]);

  const send = async () => {
    if (!input.trim() || loading) return;
    const um = { role:"user", content:input.trim() };
    const next = [...msgs, um];
    setMsgs(next); setInput(""); setLoad(true);

    // Build system prompt from real uploaded data if available
    let sys = `You are BizLens AI, a friendly personal finance assistant. `;
    if (sharedData) {
      const a = sharedData.analytics;
      sys += `The user has uploaded a bank statement: ${sharedData.fileName}.
Total Income: $${a.totalIn.toFixed(2)}, Total Spending: $${a.totalOut.toFixed(2)}, Net: $${a.net.toFixed(2)}.
Top categories: ${a.categories.slice(0,5).map(c=>`${c.name} $${c.value}`).join(", ")}.
Top expenses: ${a.topExpenses.map(t=>`${t.desc} $${Math.abs(t.amount).toFixed(2)}`).join(", ")}.
Monthly flow: ${a.monthlyFlow.map(m=>`${m.month} in:$${m.in} out:$${m.out}`).join(" | ")}.
Answer questions specifically using their real numbers.`;
    } else {
      sys += `No file has been uploaded yet. Encourage the user to upload their bank statement in the Upload tab first. Be friendly and helpful.`;
    }

    try {
      const reply = await askClaude(next.filter(m=>m.role==="user").map(m=>({role:"user",content:m.content})), sys);
      setMsgs(prev=>[...prev,{role:"assistant",content:reply}]);
    } catch(e) {
      setMsgs(prev=>[...prev,{role:"assistant",content:`Error: ${e.message}.`}]);
    }
    setLoad(false);
  };

  const tabs = [
    { id:"banking",  label:"Upload",   icon:"⇪" },
    { id:"overview", label:"Overview", icon:"◉" },
    { id:"ask",      label:"Ask AI",   icon:"◈" },
    { id:"alerts",   label:"Insights", icon:"⬡" },
    { id:"forecast", label:"Forecast", icon:"↗" },
  ];

  const card = { background:"rgba(255,255,255,0.03)", border:"1px solid rgba(255,255,255,0.07)", borderRadius:14, padding:"22px" };

  return (
    <div style={{ minHeight:"100vh", background:"#0C0F14", color:"#D8D4CC", fontFamily:"'DM Sans','Segoe UI',sans-serif", display:"flex", flexDirection:"column" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Serif+Display&family=DM+Mono:wght@400;500&display=swap');
        *{box-sizing:border-box;margin:0;padding:0;}
        ::-webkit-scrollbar{width:4px;}
        ::-webkit-scrollbar-thumb{background:rgba(255,255,255,0.1);border-radius:99px;}
        input:focus{outline:none;}
        button{font-family:inherit;}
        .dp{animation:dp 1.4s ease-in-out infinite;}
        @keyframes dp{0%,100%{opacity:1;transform:scale(1)}50%{opacity:.3;transform:scale(.65)}}
      `}</style>

      {/* HEADER */}
      <header style={{ display:"flex", alignItems:"center", justifyContent:"space-between", padding:"0 32px", height:58, borderBottom:"1px solid rgba(255,255,255,0.06)", background:"rgba(12,15,20,0.97)", position:"sticky", top:0, zIndex:100 }}>
        <div style={{ display:"flex", alignItems:"center", gap:10 }}>
          <div style={{ width:28, height:28, background:"linear-gradient(135deg,#E8C67A,#C9963A)", borderRadius:7, display:"flex", alignItems:"center", justifyContent:"center", fontSize:13, fontWeight:700, color:"#0C0F14" }}>B</div>
          <span style={{ fontSize:16, fontFamily:"'DM Serif Display',serif", color:"#F0EEE8" }}>BizLens</span>
          <span style={{ fontSize:10, color:"#4ADE80", background:"rgba(74,222,128,0.1)", padding:"2px 8px", borderRadius:99 }}>LIVE</span>
        </div>
        <nav style={{ display:"flex", gap:2 }}>
          {tabs.map(t=>(
            <button key={t.id} onClick={()=>setTab(t.id)} style={{ padding:"5px 12px", borderRadius:8, fontSize:13, cursor:"pointer", border:"none", display:"flex", alignItems:"center", gap:5, transition:"all 0.15s", background:tab===t.id?"rgba(232,198,122,0.12)":"transparent", color:tab===t.id?"#E8C67A":"#6B7A8D", fontWeight:tab===t.id?500:400 }}>
              {t.icon} {t.label}
              {t.badge && <span style={{ fontSize:9, background:"rgba(96,165,250,0.2)", color:"#60A5FA", padding:"1px 5px", borderRadius:99 }}>{t.badge}</span>}
            </button>
          ))}
        </nav>
        <div style={{ display:"flex", alignItems:"center", gap:12 }}>
          <span style={{ fontSize:12, color:"#6B7A8D" }}><span style={{ color:"#4ADE80" }}>●</span> CSV · Excel · PDF</span>
          <a
            href="https://selar.com/63v3k3b339"
            target="_blank"
            rel="noopener noreferrer"
            style={{
              display:"flex", alignItems:"center", gap:6,
              background:"linear-gradient(135deg, rgba(232,198,122,0.15), rgba(201,150,58,0.15))",
              border:"1px solid rgba(232,198,122,0.35)",
              borderRadius:99, padding:"5px 14px",
              fontSize:12, fontWeight:600,
              color:"#E8C67A", textDecoration:"none",
              transition:"all 0.2s",
              cursor:"pointer",
              whiteSpace:"nowrap",
            }}
            onMouseEnter={e => {
              e.currentTarget.style.background = "linear-gradient(135deg, rgba(232,198,122,0.28), rgba(201,150,58,0.28))";
              e.currentTarget.style.borderColor = "rgba(232,198,122,0.6)";
              e.currentTarget.style.transform = "scale(1.04)";
            }}
            onMouseLeave={e => {
              e.currentTarget.style.background = "linear-gradient(135deg, rgba(232,198,122,0.15), rgba(201,150,58,0.15))";
              e.currentTarget.style.borderColor = "rgba(232,198,122,0.35)";
              e.currentTarget.style.transform = "scale(1)";
            }}
          >
            🍖 Buy Me Suya
          </a>
        </div>
      </header>

      <main style={{ flex:1, padding:"30px 32px", maxWidth:1200, width:"100%", margin:"0 auto" }}>

        {/* UPLOAD */}
        {tab==="banking" && <PersonalBankTab sharedData={sharedData} setSharedData={(d) => { setSharedData(d); if(d) setTab("overview"); }} />}

        {/* OVERVIEW */}
        {tab==="overview" && (
          <div>
            <div style={{ marginBottom:22 }}>
              <h2 style={{ fontSize:18, fontFamily:"'DM Serif Display',serif", color:"#F0EEE8", marginBottom:4 }}>Overview</h2>
              <p style={{ fontSize:12, color:"#6B7A8D" }}>{sharedData ? `${sharedData.fileName} · ${sharedData.transactions.length} transactions` : "Upload a file to see your data here"}</p>
            </div>

            {/* Empty state */}
            {!sharedData && (
              <div style={{ ...card, textAlign:"center", padding:"60px 24px", borderStyle:"dashed" }}>
                <div style={{ fontSize:40, marginBottom:16 }}>📂</div>
                <p style={{ color:"#F0EEE8", fontSize:16, fontWeight:500, marginBottom:8 }}>No data uploaded yet</p>
                <p style={{ color:"#6B7A8D", fontSize:13, marginBottom:20 }}>Upload your bank statement to see your spending overview, charts, and AI insights</p>
                <button onClick={()=>setTab("banking")} style={{ background:"linear-gradient(135deg,#E8C67A,#C9963A)", border:"none", borderRadius:8, padding:"10px 24px", fontSize:14, color:"#0C0F14", fontWeight:600, cursor:"pointer" }}>
                  ⇪ Upload Now
                </button>
              </div>
            )}

            {/* Real data from upload */}
            {sharedData && (
              <div>
                <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:14, marginBottom:22 }}>
                  {[
                    { label:"Total Income",   value:`$${sharedData.analytics.totalIn.toLocaleString(undefined,{maximumFractionDigits:0})}`,  delta:"Credits in period",    dc:"#4ADE80" },
                    { label:"Total Spending", value:`$${sharedData.analytics.totalOut.toLocaleString(undefined,{maximumFractionDigits:0})}`, delta:"Debits in period",     dc:"#F87171" },
                    { label:"Net Cash Flow",  value:`$${Math.abs(sharedData.analytics.net).toLocaleString(undefined,{maximumFractionDigits:0})}`, delta: sharedData.analytics.net >= 0 ? "▲ Positive" : "▼ Negative", dc: sharedData.analytics.net>=0?"#4ADE80":"#F87171" },
                    { label:"Transactions",   value:sharedData.transactions.length, delta:`${sharedData.transactions.filter(t=>t.type==="credit").length} in · ${sharedData.transactions.filter(t=>t.type==="debit").length} out`, dc:"#E8C67A" },
                  ].map(s=>(
                    <div key={s.label} style={{ ...card, padding:"18px 20px" }}>
                      <p style={{ fontSize:10, color:"#6B7A8D", textTransform:"uppercase", letterSpacing:"0.12em", marginBottom:8 }}>{s.label}</p>
                      <p style={{ fontSize:24, fontFamily:"'DM Mono',monospace", color:"#F0EEE8", marginBottom:4 }}>{s.value}</p>
                      <p style={{ fontSize:12, color:s.dc }}>{s.delta}</p>
                    </div>
                  ))}
                </div>
                <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:18, marginBottom:18 }}>
                  {sharedData.analytics.monthlyFlow.length > 1 && (
                    <div style={card}>
                      <p style={{ fontSize:14, fontWeight:500, color:"#F0EEE8", marginBottom:2 }}>Monthly Cash Flow</p>
                      <p style={{ fontSize:11, color:"#6B7A8D", marginBottom:18 }}>Income vs Spending</p>
                      <ResponsiveContainer width="100%" height={190}>
                        <BarChart data={sharedData.analytics.monthlyFlow} margin={{ top:5, right:8, left:-12, bottom:0 }} barSize={12}>
                          <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.04)" vertical={false} />
                          <XAxis dataKey="month" tick={{ fontSize:10, fill:"#6B7A8D" }} axisLine={false} tickLine={false} />
                          <YAxis tick={{ fontSize:10, fill:"#6B7A8D" }} axisLine={false} tickLine={false} tickFormatter={v=>`$${(v/1000).toFixed(0)}k`} />
                          <Tooltip content={<ChartTip />} />
                          <Bar dataKey="in"  fill="#4ADE80" radius={[4,4,0,0]} name="Income" />
                          <Bar dataKey="out" fill="#F87171" radius={[4,4,0,0]} name="Spending" />
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                  )}
                  {sharedData.analytics.categories.length > 0 && (
                    <div style={card}>
                      <p style={{ fontSize:14, fontWeight:500, color:"#F0EEE8", marginBottom:2 }}>Spending by Category</p>
                      <p style={{ fontSize:11, color:"#6B7A8D", marginBottom:12 }}>Where your money goes</p>
                      <ResponsiveContainer width="100%" height={120}>
                        <PieChart>
                          <Pie data={sharedData.analytics.categories.slice(0,8)} cx="50%" cy="50%" innerRadius={32} outerRadius={52} dataKey="value" paddingAngle={3}>
                            {sharedData.analytics.categories.slice(0,8).map((_,i)=><Cell key={i} fill={PIE_COLORS[i%PIE_COLORS.length]} />)}
                          </Pie>
                          <Tooltip formatter={v=>[`$${Number(v).toLocaleString()}`,"Spending"]} contentStyle={{ background:"#141820", border:"1px solid rgba(255,255,255,0.1)", borderRadius:8, fontSize:12 }} />
                        </PieChart>
                      </ResponsiveContainer>
                      <div style={{ display:"flex", flexDirection:"column", gap:5, marginTop:8 }}>
                        {sharedData.analytics.categories.slice(0,6).map((c,i)=>(
                          <div key={c.name} style={{ display:"flex", justifyContent:"space-between" }}>
                            <div style={{ display:"flex", alignItems:"center", gap:6 }}>
                              <div style={{ width:7, height:7, borderRadius:"50%", background:PIE_COLORS[i%PIE_COLORS.length] }} />
                              <span style={{ fontSize:11, color:"#8A939F" }}>{c.name}</span>
                            </div>
                            <span style={{ fontSize:11, fontFamily:"'DM Mono',monospace", color:"#D8D4CC" }}>${c.value.toLocaleString()}</span>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}
                </div>
                {sharedData.analytics.balanceData.length > 3 && (
                  <div style={card}>
                    <p style={{ fontSize:14, fontWeight:500, color:"#F0EEE8", marginBottom:2 }}>Balance Over Time</p>
                    <p style={{ fontSize:11, color:"#6B7A8D", marginBottom:16 }}>Running account balance</p>
                    <ResponsiveContainer width="100%" height={140}>
                      <AreaChart data={sharedData.analytics.balanceData} margin={{ top:5, right:8, left:-12, bottom:0 }}>
                        <defs>
                          <linearGradient id="balGrad" x1="0" y1="0" x2="0" y2="1">
                            <stop offset="5%" stopColor="#E8C67A" stopOpacity={0.25} />
                            <stop offset="95%" stopColor="#E8C67A" stopOpacity={0} />
                          </linearGradient>
                        </defs>
                        <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.04)" />
                        <XAxis dataKey="date" tick={{ fontSize:9, fill:"#6B7A8D" }} axisLine={false} tickLine={false} />
                        <YAxis tick={{ fontSize:10, fill:"#6B7A8D" }} axisLine={false} tickLine={false} tickFormatter={v=>`$${(v/1000).toFixed(0)}k`} />
                        <Tooltip contentStyle={{ background:"#141820", border:"1px solid rgba(255,255,255,0.1)", borderRadius:8, fontSize:12 }} formatter={v=>[`$${Number(v).toLocaleString()}`,"Balance"]} />
                        <Area type="monotone" dataKey="balance" stroke="#E8C67A" strokeWidth={2} fill="url(#balGrad)" />
                      </AreaChart>
                    </ResponsiveContainer>
                  </div>
                )}
              </div>
            )}
          </div>
        )}

        {/* ASK AI */}
        {tab==="ask" && (
          <div style={{ display:"flex", flexDirection:"column", height:"calc(100vh - 130px)" }}>
            <div style={{ marginBottom:18 }}>
              <h2 style={{ fontSize:18, fontFamily:"'DM Serif Display',serif", color:"#F0EEE8", marginBottom:4 }}>Ask AI</h2>
              <p style={{ fontSize:12, color:"#6B7A8D" }}>{sharedData ? `Analysing ${sharedData.fileName}` : "Upload a file first for personalised answers"}</p>
            </div>
            <div style={{ ...card, flex:1, display:"flex", flexDirection:"column", overflow:"hidden" }}>
              <div style={{ display:"flex", gap:7, flexWrap:"wrap", marginBottom:16 }}>
                {["Why did revenue spike in December?","What's my most profitable category?","Are expenses under control?","What should I focus on this quarter?"].map(s=>(
                  <button key={s} onClick={()=>setInput(s)} style={{ background:"rgba(232,198,122,0.08)", border:"1px solid rgba(232,198,122,0.2)", borderRadius:99, padding:"5px 12px", fontSize:12, color:"#C9A84C", cursor:"pointer" }}>{s}</button>
                ))}
              </div>
              <div style={{ flex:1, overflowY:"auto", display:"flex", flexDirection:"column", gap:12, paddingRight:4 }}>
                {msgs.map((m,i)=>(
                  <div key={i} style={{ display:"flex", flexDirection:m.role==="user"?"row-reverse":"row", gap:10, alignItems:"flex-start" }}>
                    <div style={{ width:28, height:28, borderRadius:7, flexShrink:0, background:m.role==="user"?"rgba(99,91,255,0.2)":"rgba(232,198,122,0.15)", display:"flex", alignItems:"center", justifyContent:"center", fontSize:11, color:m.role==="user"?"#A5A0FF":"#E8C67A" }}>{m.role==="user"?"U":"AI"}</div>
                    <div style={{ maxWidth:"78%", background:m.role==="user"?"rgba(99,91,255,0.1)":"rgba(255,255,255,0.04)", border:`1px solid ${m.role==="user"?"rgba(99,91,255,0.2)":"rgba(255,255,255,0.07)"}`, borderRadius:10, padding:"11px 14px", fontSize:14, lineHeight:1.65, color:"#D8D4CC", whiteSpace:"pre-wrap" }}>{m.content}</div>
                  </div>
                ))}
                {loading && (
                  <div style={{ display:"flex", gap:10 }}>
                    <div style={{ width:28, height:28, borderRadius:7, background:"rgba(232,198,122,0.15)", display:"flex", alignItems:"center", justifyContent:"center", fontSize:11, color:"#E8C67A" }}>AI</div>
                    <div style={{ background:"rgba(255,255,255,0.04)", border:"1px solid rgba(255,255,255,0.07)", borderRadius:10, padding:"12px 14px", display:"flex", gap:5 }}>
                      {[0,1,2].map(i=><div key={i} className="dp" style={{ width:6, height:6, borderRadius:"50%", background:"#E8C67A", animationDelay:`${i*0.22}s` }} />)}
                    </div>
                  </div>
                )}
                <div ref={bottomRef} />
              </div>
              <div style={{ marginTop:12, display:"flex", gap:10 }}>
                <input value={input} onChange={e=>setInput(e.target.value)} onKeyDown={e=>e.key==="Enter"&&send()}
                  placeholder="Ask about revenue, forecasts, expenses…"
                  style={{ flex:1, background:"rgba(255,255,255,0.05)", border:"1px solid rgba(255,255,255,0.1)", borderRadius:10, padding:"11px 16px", fontSize:14, color:"#F0EEE8" }} />
                <button onClick={send} disabled={loading||!input.trim()} style={{ background:loading||!input.trim()?"rgba(232,198,122,0.15)":"linear-gradient(135deg,#E8C67A,#C9963A)", border:"none", borderRadius:10, padding:"11px 20px", fontSize:14, color:loading||!input.trim()?"#6B7A8D":"#0C0F14", cursor:"pointer", fontWeight:600 }}>Send →</button>
              </div>
            </div>
          </div>
        )}

        {/* INSIGHTS */}
        {tab==="alerts" && (
          <div>
            <div style={{ marginBottom:22 }}>
              <h2 style={{ fontSize:18, fontFamily:"'DM Serif Display',serif", color:"#F0EEE8", marginBottom:4 }}>AI-Generated Insights</h2>
              <p style={{ fontSize:12, color:"#6B7A8D" }}>Automatically surfaced from connected data</p>
            </div>
            <div style={{ display:"flex", flexDirection:"column", gap:12, marginBottom:22 }}>
              {alerts.map(a=>(
                <div key={a.id} style={{ ...card, borderLeft:`3px solid ${a.type==="positive"?"#4ADE80":a.type==="warning"?"#FBBF24":"#60A5FA"}`, display:"flex", gap:16, alignItems:"flex-start" }}>
                  <div style={{ width:36, height:36, borderRadius:9, flexShrink:0, background:a.type==="positive"?"rgba(74,222,128,0.12)":a.type==="warning"?"rgba(251,191,36,0.12)":"rgba(96,165,250,0.12)", display:"flex", alignItems:"center", justifyContent:"center", fontSize:16, color:a.type==="positive"?"#4ADE80":a.type==="warning"?"#FBBF24":"#60A5FA" }}>{a.icon}</div>
                  <div><p style={{ fontWeight:500, color:"#F0EEE8", marginBottom:4 }}>{a.title}</p><p style={{ color:"#8A939F", fontSize:13, lineHeight:1.6 }}>{a.body}</p></div>
                </div>
              ))}
            </div>
            <div style={card}>
              <p style={{ fontSize:15, fontWeight:500, color:"#F0EEE8", marginBottom:18 }}>Performance Scores</p>
              {[{label:"Revenue Growth",score:82,color:"#E8C67A"},{label:"Expense Efficiency",score:67,color:"#F87171"},{label:"Profit Margin",score:91,color:"#4ADE80"},{label:"Revenue Consistency",score:74,color:"#60A5FA"}].map(m=>(
                <div key={m.label} style={{ marginBottom:14 }}>
                  <div style={{ display:"flex", justifyContent:"space-between", marginBottom:5 }}>
                    <span style={{ fontSize:13, color:"#9AA3AF" }}>{m.label}</span>
                    <span style={{ fontSize:13, fontFamily:"'DM Mono',monospace", color:m.color }}>{m.score}/100</span>
                  </div>
                  <div style={{ height:5, background:"rgba(255,255,255,0.06)", borderRadius:99 }}>
                    <div style={{ width:`${m.score}%`, height:"100%", background:m.color, borderRadius:99 }} />
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* FORECAST */}
        {tab==="forecast" && (
          <div>
            <div style={{ marginBottom:22 }}>
              <h2 style={{ fontSize:18, fontFamily:"'DM Serif Display',serif", color:"#F0EEE8", marginBottom:4 }}>Revenue Forecast</h2>
              <p style={{ fontSize:12, color:"#6B7A8D" }}>AI projection · Next 4 months</p>
            </div>
            <div style={{ ...card, marginBottom:18 }}>
              <ResponsiveContainer width="100%" height={250}>
                <LineChart data={forecastData} margin={{ top:10, right:18, left:-10, bottom:0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.04)" />
                  <XAxis dataKey="month" tick={{ fontSize:11, fill:"#6B7A8D" }} axisLine={false} tickLine={false} />
                  <YAxis tick={{ fontSize:11, fill:"#6B7A8D" }} axisLine={false} tickLine={false} tickFormatter={v=>`$${(v/1000).toFixed(0)}k`} />
                  <Tooltip content={<ChartTip />} />
                  <Line type="monotone" dataKey="actual" stroke="#E8C67A" strokeWidth={2.5} dot={{ fill:"#E8C67A", r:5 }} name="Actual" />
                  <Line type="monotone" dataKey="forecast" stroke="#60A5FA" strokeWidth={2.5} strokeDasharray="6 3" dot={{ fill:"#60A5FA", r:5 }} name="Forecast" />
                </LineChart>
              </ResponsiveContainer>
            </div>
            <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:14 }}>
              {[{month:"March",value:"$36,500",pct:"+7.7%"},{month:"April",value:"$39,200",pct:"+7.4%"},{month:"May",value:"$43,100",pct:"+10.0%"},{month:"June",value:"$47,800",pct:"+10.9%"}].map(f=>(
                <div key={f.month} style={card}>
                  <p style={{ fontSize:10, color:"#6B7A8D", textTransform:"uppercase", letterSpacing:"0.12em", marginBottom:8 }}>{f.month}</p>
                  <p style={{ fontSize:22, fontFamily:"'DM Mono',monospace", color:"#F0EEE8", marginBottom:4 }}>{f.value}</p>
                  <p style={{ fontSize:12, color:"#4ADE80" }}>▲ {f.pct} projected</p>
                </div>
              ))}
            </div>
          </div>
        )}

      </main>
    </div>
  );
}