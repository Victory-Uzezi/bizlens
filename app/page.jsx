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
const connectors = [
  { id: "shopify", label: "Shopify", icon: "⬡", status: "connected", color: "#96BF48" },
  { id: "stripe", label: "Stripe", icon: "⚡", status: "connected", color: "#635BFF" },
  { id: "sheets", label: "Google Sheets", icon: "⊞", status: "connected", color: "#34A853" },
  { id: "quickbooks", label: "QuickBooks", icon: "◈", status: "pending", color: "#2CA01C" },
  { id: "square", label: "Square", icon: "■", status: "idle", color: "#3E4348" },
  { id: "csv", label: "CSV Upload", icon: "⇪", status: "idle", color: "#3E4348" },
];

// ─── COLOURS ──────────────────────────────────────────────────────────────────
const PIE_COLORS = ["#E8C67A","#60A5FA","#4ADE80","#F87171","#A78BFA","#FBBF24","#34D399","#FB923C"];

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
  return data.content?.map(b => b.text || "").join("") || "No response.";
}

// ─── HELPERS ──────────────────────────────────────────────────────────────────
const fmt = n => `$${Math.abs(n).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
const fmtShort = n => `$${Math.abs(n).toLocaleString(undefined, { maximumFractionDigits: 0 })}`;

const Tip = ({ active, payload, label }) => {
  if (!active || !payload?.length) return null;
  return (
    <div style={{ background: "#141820", border: "1px solid rgba(255,255,255,0.1)", borderRadius: 8, padding: "10px 14px", fontSize: 12 }}>
      <p style={{ color: "#9AA3AF", margin: "0 0 5px" }}>{label}</p>
      {payload.map((p, i) => <p key={i} style={{ color: p.color, margin: "2px 0" }}>{p.name}: ${p.value?.toLocaleString()}</p>)}
    </div>
  );
};

// ─── PARSE EXCEL BANK STATEMENT ───────────────────────────────────────────────
function parseStatement(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

        // Find the header row (look for common bank statement column names)
        let headerIdx = 0;
        const keyWords = ["date","description","amount","debit","credit","balance","transaction","narration","particulars"];
        for (let i = 0; i < Math.min(raw.length, 15); i++) {
          const row = raw[i].map(c => String(c).toLowerCase());
          if (row.some(c => keyWords.some(k => c.includes(k)))) { headerIdx = i; break; }
        }

        const headers = raw[headerIdx].map(h => String(h).trim().toLowerCase());
        const rows = raw.slice(headerIdx + 1).filter(r => r.some(c => c !== ""));

        // Map columns intelligently
        const colIndex = (keywords) => {
          for (const kw of keywords) {
            const idx = headers.findIndex(h => h.includes(kw));
            if (idx !== -1) return idx;
          }
          return -1;
        };

        const dateCol    = colIndex(["date","time"]);
        const descCol    = colIndex(["description","narration","particulars","memo","details","ref"]);
        const amtCol     = colIndex(["amount"]);
        const debitCol   = colIndex(["debit","withdrawal","dr"]);
        const creditCol  = colIndex(["credit","deposit","cr"]);
        const balanceCol = colIndex(["balance","running"]);

        const transactions = [];
        for (const row of rows) {
          const dateRaw = dateCol !== -1 ? row[dateCol] : "";
          let date = "";
          if (dateRaw) {
            // Handle Excel serial date numbers
            if (typeof dateRaw === "number") {
              const d = XLSX.SSF.parse_date_code(dateRaw);
              date = `${d.y}-${String(d.m).padStart(2,"0")}-${String(d.d).padStart(2,"0")}`;
            } else {
              date = String(dateRaw).trim();
            }
          }
          const desc    = descCol !== -1 ? String(row[descCol] || "").trim() : "";
          const balance = balanceCol !== -1 ? parseFloat(String(row[balanceCol]).replace(/[^0-9.-]/g,"")) || null : null;

          let amount = 0;
          if (amtCol !== -1) {
            const raw = parseFloat(String(row[amtCol]).replace(/[^0-9.-]/g,"")) || 0;
            amount = raw;
          } else {
            const debit  = parseFloat(String(row[debitCol]  || "0").replace(/[^0-9.-]/g,"")) || 0;
            const credit = parseFloat(String(row[creditCol] || "0").replace(/[^0-9.-]/g,"")) || 0;
            amount = credit > 0 ? credit : -debit;
          }

          if (date || desc || amount) {
            transactions.push({ date, desc, amount, balance, type: amount >= 0 ? "credit" : "debit" });
          }
        }

        resolve(transactions);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// ─── AUTO-CATEGORISE ─────────────────────────────────────────────────────────
function categorise(desc) {
  const d = desc.toLowerCase();
  if (/salary|payroll|wages|income|deposit/.test(d))          return "Income";
  if (/rent|mortgage|lease/.test(d))                          return "Rent/Mortgage";
  if (/grocery|supermarket|walmart|whole food|safeway|kroger/.test(d)) return "Groceries";
  if (/uber|lyft|taxi|gas|fuel|petrol|shell|bp|chevron/.test(d))       return "Transport";
  if (/netflix|spotify|hulu|amazon prime|disney|subscription/.test(d)) return "Subscriptions";
  if (/restaurant|cafe|coffee|starbucks|mcdonald|food|pizza|doordash|grubhub/.test(d)) return "Food & Dining";
  if (/electricity|water|internet|phone|bill|utility/.test(d))         return "Utilities";
  if (/hospital|pharmacy|doctor|health|medical|dental/.test(d))        return "Healthcare";
  if (/amazon|shopping|store|mall|target|bestbuy/.test(d))             return "Shopping";
  if (/transfer|zelle|venmo|paypal/.test(d))                           return "Transfers";
  if (/atm|withdrawal|cash/.test(d))                                   return "Cash";
  if (/insurance/.test(d))                                             return "Insurance";
  if (/gym|fitness|sport/.test(d))                                     return "Health & Fitness";
  return "Other";
}

// ─── DERIVE ANALYTICS FROM TRANSACTIONS ──────────────────────────────────────
function deriveAnalytics(transactions) {
  const credits = transactions.filter(t => t.amount > 0);
  const debits  = transactions.filter(t => t.amount < 0);
  const totalIn  = credits.reduce((s, t) => s + t.amount, 0);
  const totalOut = debits.reduce((s, t)  => s + Math.abs(t.amount), 0);
  const net      = totalIn - totalOut;

  // Category breakdown (debits only)
  const catMap = {};
  for (const t of debits) {
    const cat = categorise(t.desc);
    catMap[cat] = (catMap[cat] || 0) + Math.abs(t.amount);
  }
  const categories = Object.entries(catMap)
    .map(([name, value]) => ({ name, value: parseFloat(value.toFixed(2)) }))
    .sort((a, b) => b.value - a.value);

  // Monthly flow
  const monthMap = {};
  for (const t of transactions) {
    const key = t.date ? t.date.slice(0, 7) : "Unknown";
    if (!monthMap[key]) monthMap[key] = { month: key, in: 0, out: 0 };
    if (t.amount > 0) monthMap[key].in  += t.amount;
    else              monthMap[key].out += Math.abs(t.amount);
  }
  const monthlyFlow = Object.values(monthMap)
    .sort((a, b) => a.month.localeCompare(b.month))
    .map(m => ({ ...m, in: parseFloat(m.in.toFixed(2)), out: parseFloat(m.out.toFixed(2)) }));

  // Balance trend (last balance per day)
  const balanceTrend = transactions
    .filter(t => t.balance != null)
    .reduce((acc, t) => {
      acc[t.date] = t.balance;
      return acc;
    }, {});
  const balanceData = Object.entries(balanceTrend)
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([date, balance]) => ({ date, balance }));

  // Top 5 largest expenses
  const topExpenses = [...debits]
    .sort((a, b) => Math.abs(b.amount) - Math.abs(a.amount))
    .slice(0, 5);

  return { totalIn, totalOut, net, categories, monthlyFlow, balanceData, topExpenses };
}

// ─── PERSONAL BANK STATEMENT TAB ─────────────────────────────────────────────
function PersonalBankTab() {
  const [transactions, setTransactions] = useState(null);
  const [analytics, setAnalytics]       = useState(null);
  const [fileName, setFileName]          = useState("");
  const [parsing, setParsing]            = useState(false);
  const [parseError, setParseError]      = useState("");
  const [dragging, setDragging]          = useState(false);
  const [aiInsight, setAiInsight]        = useState("");
  const [aiLoading, setAiLoading]        = useState(false);
  const [question, setQuestion]          = useState("");
  const [qaHistory, setQaHistory]        = useState([]);
  const [qaLoading, setQaLoading]        = useState(false);
  const [txFilter, setTxFilter]          = useState("All");
  const fileInputRef = useRef(null);
  const qaBottomRef  = useRef(null);

  useEffect(() => { qaBottomRef.current?.scrollIntoView({ behavior: "smooth" }); }, [qaHistory]);

  const processFile = async (file) => {
    if (!file) return;
    const ext = file.name.split(".").pop().toLowerCase();
    if (!["xlsx","xls","csv"].includes(ext)) {
      setParseError("Please upload an Excel (.xlsx / .xls) or CSV file."); return;
    }
    setParsing(true); setParseError(""); setTransactions(null); setAnalytics(null);
    setAiInsight(""); setQaHistory([]);
    try {
      const txns = await parseStatement(file);
      const cats = txns.map(t => ({ ...t, category: categorise(t.desc) }));
      const anal = deriveAnalytics(cats);
      setTransactions(cats);
      setAnalytics(anal);
      setFileName(file.name);
      // Auto-generate AI insight
      generateInsight(cats, anal);
    } catch (e) {
      setParseError("Could not parse this file. Make sure it has Date, Description, and Amount columns.");
    }
    setParsing(false);
  };

  const generateInsight = async (txns, anal) => {
    setAiLoading(true);
    const summary = `
Transactions: ${txns.length}
Total Credits: $${anal.totalIn.toFixed(2)}
Total Debits:  $${anal.totalOut.toFixed(2)}
Net:           $${anal.net.toFixed(2)}
Top spend categories: ${anal.categories.slice(0,5).map(c => `${c.name} $${c.value}`).join(", ")}
Top 5 expenses: ${anal.topExpenses.map(t => `${t.desc} $${Math.abs(t.amount).toFixed(2)}`).join(", ")}
Monthly flow: ${anal.monthlyFlow.map(m => `${m.month}: in $${m.in} out $${m.out}`).join(" | ")}
    `.trim();
    const sys = `You are a friendly and insightful personal finance advisor. Given a bank statement summary, provide a clear and encouraging 4–6 sentence analysis. Cover: overall cash flow health, biggest spending categories to watch, any positive patterns, and 1-2 specific actionable tips. Be warm and direct — no bullet lists, just flowing prose.`;
    try {
      const reply = await askClaude([{ role: "user", content: `Here is my bank statement summary:\n\n${summary}\n\nPlease analyse it.` }], sys);
      setAiInsight(reply);
    } catch { setAiInsight("Could not generate insight. Check your API configuration."); }
    setAiLoading(false);
  };

  const askQuestion = async () => {
    if (!question.trim() || qaLoading || !analytics) return;
    const q = question.trim();
    setQaHistory(prev => [...prev, { role: "user", content: q }]);
    setQuestion("");
    setQaLoading(true);
    const summary = `
Total Credits: $${analytics.totalIn.toFixed(2)}, Debits: $${analytics.totalOut.toFixed(2)}, Net: $${analytics.net.toFixed(2)}
Categories: ${analytics.categories.map(c => `${c.name}: $${c.value}`).join(", ")}
Top expenses: ${analytics.topExpenses.map(t => `${t.desc}: $${Math.abs(t.amount).toFixed(2)}`).join(", ")}
Monthly: ${analytics.monthlyFlow.map(m => `${m.month} in:$${m.in} out:$${m.out}`).join(" | ")}
    `.trim();
    const sys = `You are a personal finance assistant. You have the user's bank statement data:\n${summary}\nAnswer questions helpfully, specifically, and concisely (2-4 sentences). Use exact numbers from the data.`;
    const history = [...qaHistory, { role: "user", content: q }].filter(m => m.role === "user").map(m => ({ role: "user", content: m.content }));
    try {
      const reply = await askClaude(history, sys);
      setQaHistory(prev => [...prev, { role: "assistant", content: reply }]);
    } catch { setQaHistory(prev => [...prev, { role: "assistant", content: "Error connecting to AI. Check your configuration." }]); }
    setQaLoading(false);
  };

  const onDrop = useCallback((e) => {
    e.preventDefault(); setDragging(false);
    const file = e.dataTransfer.files[0];
    if (file) processFile(file);
  }, []);

  const card = { background: "rgba(255,255,255,0.03)", border: "1px solid rgba(255,255,255,0.07)", borderRadius: 14, padding: "22px" };
  const cats = transactions ? ["All", ...new Set(transactions.map(t => t.category))] : [];
  const visibleTxns = transactions
    ? (txFilter === "All" ? transactions : transactions.filter(t => t.category === txFilter))
    : [];

  return (
    <div>
      <div style={{ marginBottom: 22 }}>
        <h2 style={{ fontSize: 18, fontFamily: "'DM Serif Display',serif", color: "#F0EEE8", margin: "0 0 4px" }}>Personal Bank Statement Analyser</h2>
        <p style={{ fontSize: 12, color: "#6B7A8D", margin: 0 }}>Upload your bank statement Excel or CSV file — Claude will analyse your spending instantly</p>
      </div>

      {/* Upload Zone */}
      {!transactions && (
        <div
          onDragOver={e => { e.preventDefault(); setDragging(true); }}
          onDragLeave={() => setDragging(false)}
          onDrop={onDrop}
          style={{
            ...card,
            borderStyle: "dashed",
            borderColor: dragging ? "#E8C67A" : "rgba(255,255,255,0.12)",
            background: dragging ? "rgba(232,198,122,0.05)" : "rgba(255,255,255,0.02)",
            textAlign: "center", padding: "52px 24px",
            transition: "all 0.2s", cursor: "pointer",
          }}
          onClick={() => fileInputRef.current?.click()}
        >
          <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }}
            onChange={e => processFile(e.target.files[0])} />
          <div style={{ fontSize: 40, marginBottom: 14 }}>📊</div>
          <p style={{ color: "#F0EEE8", fontSize: 16, fontWeight: 500, marginBottom: 8 }}>
            {parsing ? "Reading your statement…" : "Drop your bank statement here"}
          </p>
          <p style={{ color: "#6B7A8D", fontSize: 13, marginBottom: 20 }}>
            Supports Excel (.xlsx, .xls) and CSV files exported from any bank
          </p>
          {!parsing && (
            <button style={{ background: "linear-gradient(135deg,#E8C67A,#C9963A)", border: "none", borderRadius: 8, padding: "10px 24px", fontSize: 14, color: "#0C0F14", fontWeight: 600, cursor: "pointer" }}>
              Browse Files
            </button>
          )}
          {parsing && <div style={{ color: "#E8C67A", fontSize: 13 }}>Parsing rows…</div>}
          {parseError && <p style={{ color: "#F87171", fontSize: 13, marginTop: 12 }}>{parseError}</p>}
        </div>
      )}

      {/* ── Results ── */}
      {transactions && analytics && (
        <div>
          {/* File + Reset header */}
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <span style={{ fontSize: 18 }}>📄</span>
              <div>
                <p style={{ color: "#F0EEE8", fontSize: 14, fontWeight: 500, margin: 0 }}>{fileName}</p>
                <p style={{ color: "#6B7A8D", fontSize: 12, margin: 0 }}>{transactions.length} transactions parsed</p>
              </div>
            </div>
            <button onClick={() => { setTransactions(null); setAnalytics(null); setFileName(""); setAiInsight(""); setQaHistory([]); }}
              style={{ background: "rgba(248,113,113,0.1)", border: "1px solid rgba(248,113,113,0.3)", borderRadius: 8, padding: "6px 14px", fontSize: 12, color: "#F87171", cursor: "pointer" }}>
              ✕ Remove file
            </button>
          </div>

          {/* Stat Cards */}
          <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 14, marginBottom: 20 }}>
            {[
              { label: "Total Income",   value: fmtShort(analytics.totalIn),  sub: "Credits in period",    color: "#4ADE80" },
              { label: "Total Spending", value: fmtShort(analytics.totalOut), sub: "Debits in period",     color: "#F87171" },
              { label: "Net Cash Flow",  value: fmtShort(analytics.net),      sub: analytics.net >= 0 ? "▲ Positive" : "▼ Negative", color: analytics.net >= 0 ? "#4ADE80" : "#F87171" },
              { label: "Transactions",   value: transactions.length,           sub: `${transactions.filter(t=>t.type==="credit").length} in · ${transactions.filter(t=>t.type==="debit").length} out`, color: "#E8C67A" },
            ].map(s => (
              <div key={s.label} style={{ background: "rgba(255,255,255,0.03)", border: "1px solid rgba(255,255,255,0.07)", borderRadius: 12, padding: "18px 20px" }}>
                <p style={{ fontSize: 10, color: "#6B7A8D", textTransform: "uppercase", letterSpacing: "0.12em", margin: "0 0 8px" }}>{s.label}</p>
                <p style={{ fontSize: 24, fontFamily: "'DM Mono',monospace", color: s.color, margin: "0 0 4px" }}>{s.value}</p>
                <p style={{ fontSize: 11, color: "#6B7A8D", margin: 0 }}>{s.sub}</p>
              </div>
            ))}
          </div>

          {/* Charts Row */}
          <div style={{ display: "grid", gridTemplateColumns: "3fr 2fr", gap: 18, marginBottom: 18 }}>
            {/* Monthly Flow */}
            {analytics.monthlyFlow.length > 1 && (
              <div style={card}>
                <p style={{ fontSize: 14, fontWeight: 500, color: "#F0EEE8", margin: "0 0 2px" }}>Monthly Cash Flow</p>
                <p style={{ fontSize: 11, color: "#6B7A8D", margin: "0 0 18px" }}>Income vs Spending by month</p>
                <ResponsiveContainer width="100%" height={180}>
                  <BarChart data={analytics.monthlyFlow} margin={{ top: 5, right: 5, left: -20, bottom: 0 }} barSize={12}>
                    <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.04)" vertical={false} />
                    <XAxis dataKey="month" tick={{ fontSize: 10, fill: "#6B7A8D" }} axisLine={false} tickLine={false} />
                    <YAxis tick={{ fontSize: 10, fill: "#6B7A8D" }} axisLine={false} tickLine={false} tickFormatter={v => `$${(v/1000).toFixed(0)}k`} />
                    <Tooltip content={<Tip />} />
                    <Bar dataKey="in"  fill="#4ADE80" radius={[4,4,0,0]} name="Income" />
                    <Bar dataKey="out" fill="#F87171" radius={[4,4,0,0]} name="Spending" />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            )}

            {/* Spend by Category Donut */}
            <div style={card}>
              <p style={{ fontSize: 14, fontWeight: 500, color: "#F0EEE8", margin: "0 0 2px" }}>Spending by Category</p>
              <p style={{ fontSize: 11, color: "#6B7A8D", margin: "0 0 12px" }}>Where your money goes</p>
              <ResponsiveContainer width="100%" height={110}>
                <PieChart>
                  <Pie data={analytics.categories.slice(0,8)} cx="50%" cy="50%" innerRadius={30} outerRadius={50} dataKey="value" paddingAngle={3}>
                    {analytics.categories.slice(0,8).map((_, i) => <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />)}
                  </Pie>
                  <Tooltip formatter={v => `$${v.toLocaleString()}`} contentStyle={{ background: "#141820", border: "1px solid rgba(255,255,255,0.1)", borderRadius: 8, fontSize: 12 }} />
                </PieChart>
              </ResponsiveContainer>
              <div style={{ display: "flex", flexDirection: "column", gap: 5, marginTop: 8 }}>
                {analytics.categories.slice(0,6).map((c, i) => (
                  <div key={c.name} style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                      <div style={{ width: 7, height: 7, borderRadius: "50%", background: PIE_COLORS[i] }} />
                      <span style={{ fontSize: 11, color: "#8A939F" }}>{c.name}</span>
                    </div>
                    <span style={{ fontSize: 11, fontFamily: "'DM Mono',monospace", color: "#D8D4CC" }}>${c.value.toLocaleString()}</span>
                  </div>
                ))}
              </div>
            </div>
          </div>

          {/* Balance Trend (if available) */}
          {analytics.balanceData.length > 2 && (
            <div style={{ ...card, marginBottom: 18 }}>
              <p style={{ fontSize: 14, fontWeight: 500, color: "#F0EEE8", margin: "0 0 2px" }}>Balance Over Time</p>
              <p style={{ fontSize: 11, color: "#6B7A8D", margin: "0 0 16px" }}>Running account balance</p>
              <ResponsiveContainer width="100%" height={140}>
                <AreaChart data={analytics.balanceData} margin={{ top: 5, right: 8, left: -12, bottom: 0 }}>
                  <defs>
                    <linearGradient id="balGrad" x1="0" y1="0" x2="0" y2="1">
                      <stop offset="5%" stopColor="#E8C67A" stopOpacity={0.25} />
                      <stop offset="95%" stopColor="#E8C67A" stopOpacity={0} />
                    </linearGradient>
                  </defs>
                  <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.04)" />
                  <XAxis dataKey="date" tick={{ fontSize: 10, fill: "#6B7A8D" }} axisLine={false} tickLine={false} />
                  <YAxis tick={{ fontSize: 10, fill: "#6B7A8D" }} axisLine={false} tickLine={false} tickFormatter={v => `$${(v/1000).toFixed(0)}k`} />
                  <Tooltip contentStyle={{ background: "#141820", border: "1px solid rgba(255,255,255,0.1)", borderRadius: 8, fontSize: 12 }} formatter={v => [`$${v.toLocaleString()}`, "Balance"]} />
                  <Area type="monotone" dataKey="balance" stroke="#E8C67A" strokeWidth={2} fill="url(#balGrad)" name="Balance" />
                </AreaChart>
              </ResponsiveContainer>
            </div>
          )}

          {/* AI Insight Card */}
          <div style={{ ...card, marginBottom: 18, borderColor: "rgba(232,198,122,0.2)", background: "rgba(232,198,122,0.04)" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
              <div style={{ width: 30, height: 30, borderRadius: 8, background: "rgba(232,198,122,0.18)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 14, color: "#E8C67A" }}>✦</div>
              <p style={{ fontSize: 14, fontWeight: 500, color: "#E8C67A", margin: 0 }}>AI Financial Insight</p>
            </div>
            {aiLoading
              ? <div style={{ display: "flex", gap: 5 }}>{[0,1,2].map(i => <div key={i} className="dp" style={{ width: 7, height: 7, borderRadius: "50%", background: "#E8C67A", animationDelay: `${i*0.22}s` }} />)}</div>
              : <p style={{ fontSize: 14, color: "#C9C3B8", lineHeight: 1.75, margin: 0 }}>{aiInsight}</p>
            }
          </div>

          {/* Ask AI about my statement */}
          <div style={{ ...card, marginBottom: 18 }}>
            <p style={{ fontSize: 14, fontWeight: 500, color: "#F0EEE8", margin: "0 0 4px" }}>Ask About Your Statement</p>
            <p style={{ fontSize: 11, color: "#6B7A8D", margin: "0 0 14px" }}>Claude has full context of your uploaded data</p>

            {/* Suggested questions */}
            <div style={{ display: "flex", gap: 7, flexWrap: "wrap", marginBottom: 14 }}>
              {["What's my biggest spending category?", "Am I saving enough?", "What can I cut back on?", "How consistent is my income?"].map(s => (
                <button key={s} onClick={() => setQuestion(s)} style={{ background: "rgba(232,198,122,0.08)", border: "1px solid rgba(232,198,122,0.2)", borderRadius: 99, padding: "5px 12px", fontSize: 12, color: "#C9A84C", cursor: "pointer" }}>{s}</button>
              ))}
            </div>

            {/* Q&A history */}
            {qaHistory.length > 0 && (
              <div style={{ maxHeight: 260, overflowY: "auto", display: "flex", flexDirection: "column", gap: 10, marginBottom: 14, paddingRight: 4 }}>
                {qaHistory.map((m, i) => (
                  <div key={i} style={{ display: "flex", flexDirection: m.role === "user" ? "row-reverse" : "row", gap: 8, alignItems: "flex-start" }}>
                    <div style={{ width: 26, height: 26, borderRadius: 6, flexShrink: 0, background: m.role === "user" ? "rgba(99,91,255,0.2)" : "rgba(232,198,122,0.15)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 10, color: m.role === "user" ? "#A5A0FF" : "#E8C67A" }}>{m.role === "user" ? "U" : "AI"}</div>
                    <div style={{ maxWidth: "80%", background: m.role === "user" ? "rgba(99,91,255,0.1)" : "rgba(255,255,255,0.04)", border: `1px solid ${m.role === "user" ? "rgba(99,91,255,0.2)" : "rgba(255,255,255,0.07)"}`, borderRadius: 10, padding: "10px 14px", fontSize: 13, lineHeight: 1.65, color: "#D8D4CC", whiteSpace: "pre-wrap" }}>{m.content}</div>
                  </div>
                ))}
                {qaLoading && (
                  <div style={{ display: "flex", gap: 8 }}>
                    <div style={{ width: 26, height: 26, borderRadius: 6, background: "rgba(232,198,122,0.15)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 10, color: "#E8C67A" }}>AI</div>
                    <div style={{ background: "rgba(255,255,255,0.04)", border: "1px solid rgba(255,255,255,0.07)", borderRadius: 10, padding: "12px 14px", display: "flex", gap: 5 }}>
                      {[0,1,2].map(i => <div key={i} className="dp" style={{ width: 6, height: 6, borderRadius: "50%", background: "#E8C67A", animationDelay: `${i*0.22}s` }} />)}
                    </div>
                  </div>
                )}
                <div ref={qaBottomRef} />
              </div>
            )}

            <div style={{ display: "flex", gap: 10 }}>
              <input value={question} onChange={e => setQuestion(e.target.value)} onKeyDown={e => e.key === "Enter" && askQuestion()}
                placeholder="e.g. How much did I spend on food last month?"
                style={{ flex: 1, background: "rgba(255,255,255,0.05)", border: "1px solid rgba(255,255,255,0.1)", borderRadius: 10, padding: "11px 16px", fontSize: 14, color: "#F0EEE8", fontFamily: "inherit" }} />
              <button onClick={askQuestion} disabled={qaLoading || !question.trim()} style={{ background: qaLoading || !question.trim() ? "rgba(232,198,122,0.15)" : "linear-gradient(135deg,#E8C67A,#C9963A)", border: "none", borderRadius: 10, padding: "11px 20px", fontSize: 14, color: qaLoading || !question.trim() ? "#6B7A8D" : "#0C0F14", cursor: "pointer", fontWeight: 600, fontFamily: "inherit" }}>Ask →</button>
            </div>
          </div>

          {/* Full Transaction Table */}
          <div style={card}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
              <div>
                <p style={{ fontSize: 14, fontWeight: 500, color: "#F0EEE8", margin: "0 0 2px" }}>All Transactions</p>
                <p style={{ fontSize: 11, color: "#6B7A8D", margin: 0 }}>{visibleTxns.length} records shown</p>
              </div>
              <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                {cats.slice(0, 8).map(c => (
                  <button key={c} onClick={() => setTxFilter(c)} style={{ padding: "4px 10px", borderRadius: 99, fontSize: 11, cursor: "pointer", border: "none", fontFamily: "inherit", background: txFilter === c ? "rgba(232,198,122,0.2)" : "rgba(255,255,255,0.05)", color: txFilter === c ? "#E8C67A" : "#6B7A8D", outline: txFilter === c ? "1px solid rgba(232,198,122,0.4)" : "1px solid transparent" }}>{c}</button>
                ))}
              </div>
            </div>

            <div style={{ borderTop: "1px solid rgba(255,255,255,0.05)" }}>
              <div style={{ display: "grid", gridTemplateColumns: "100px 1fr 110px 100px", gap: 12, padding: "8px 10px" }}>
                {["Date","Description","Category","Amount"].map(h => (
                  <span key={h} style={{ fontSize: 10, color: "#6B7A8D", textTransform: "uppercase", letterSpacing: "0.1em" }}>{h}</span>
                ))}
              </div>
              <div style={{ maxHeight: 340, overflowY: "auto" }}>
                {visibleTxns.slice(0, 100).map((tx, i) => (
                  <div key={i} style={{ display: "grid", gridTemplateColumns: "100px 1fr 110px 100px", gap: 12, padding: "10px 10px", borderRadius: 7, alignItems: "center", background: i % 2 === 0 ? "rgba(255,255,255,0.015)" : "transparent" }}>
                    <span style={{ fontSize: 11, color: "#6B7A8D", fontFamily: "'DM Mono',monospace" }}>{tx.date || "—"}</span>
                    <span style={{ fontSize: 13, color: "#D8D4CC", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{tx.desc || "—"}</span>
                    <span style={{ fontSize: 10, padding: "3px 8px", borderRadius: 99, background: "rgba(255,255,255,0.05)", color: "#9AA3AF", width: "fit-content" }}>{tx.category}</span>
                    <span style={{ fontSize: 13, fontFamily: "'DM Mono',monospace", fontWeight: 500, color: tx.type === "credit" ? "#4ADE80" : "#F87171", textAlign: "right" }}>
                      {tx.type === "credit" ? "+" : "−"}{fmt(tx.amount)}
                    </span>
                  </div>
                ))}
                {visibleTxns.length > 100 && <p style={{ textAlign: "center", color: "#6B7A8D", fontSize: 12, padding: "12px 0" }}>Showing first 100 of {visibleTxns.length} transactions</p>}
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── MAIN APP ─────────────────────────────────────────────────────────────────
export default function BizLens() {
  const [tab, setTab] = useState("overview");
  const [msgs, setMsgs] = useState([{ role: "assistant", content: "Hi! I'm BizLens AI. Ask me anything about your business metrics, forecasts, or upload a personal bank statement in the Banking tab for a full personal finance analysis." }]);
  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  const bottomRef = useRef(null);

  useEffect(() => { bottomRef.current?.scrollIntoView({ behavior: "smooth" }); }, [msgs]);

  const send = async () => {
    if (!input.trim() || loading) return;
    const um = { role: "user", content: input.trim() };
    const next = [...msgs, um];
    setMsgs(next); setInput(""); setLoading(true);
    const sys = `You are BizLens AI, an expert business analyst. Business data: Revenue Feb $33,900 | Expenses $15,400 | Margin 54.6% | Sources: Shopify, Stripe, Google Sheets | March forecast $36,500. Answer concisely and helpfully.`;
    try {
      const reply = await askClaude(next.filter(m => m.role === "user").map(m => ({ role: "user", content: m.content })), sys);
      setMsgs(prev => [...prev, { role: "assistant", content: reply }]);
    } catch { setMsgs(prev => [...prev, { role: "assistant", content: "Connection error. Check your API configuration." }]); }
    setLoading(false);
  };

  const tabs = [
    { id: "overview", label: "Overview",  icon: "◉" },
    { id: "banking",  label: "Banking",   icon: "⊛" },
    { id: "ask",      label: "Ask AI",    icon: "◈" },
    { id: "alerts",   label: "Insights",  icon: "⬡" },
    { id: "forecast", label: "Forecast",  icon: "↗" },
    { id: "sources",  label: "Sources",   icon: "⊞" },
  ];

  const card = { background: "rgba(255,255,255,0.03)", border: "1px solid rgba(255,255,255,0.07)", borderRadius: 14, padding: "22px" };

  return (
    <div style={{ minHeight: "100vh", background: "#0C0F14", color: "#D8D4CC", fontFamily: "'DM Sans','Segoe UI',sans-serif", display: "flex", flexDirection: "column" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Serif+Display&family=DM+Mono:wght@400;500&display=swap');
        * { box-sizing: border-box; margin: 0; padding: 0; }
        ::-webkit-scrollbar { width: 4px; }
        ::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.1); border-radius: 99px; }
        input:focus { outline: none; }
        button { font-family: inherit; }
        .dp { animation: dp 1.4s ease-in-out infinite; }
        @keyframes dp { 0%,100%{opacity:1;transform:scale(1)} 50%{opacity:.3;transform:scale(.65)} }
      `}</style>

      {/* HEADER */}
      <header style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "0 32px", height: 58, borderBottom: "1px solid rgba(255,255,255,0.06)", background: "rgba(12,15,20,0.97)", position: "sticky", top: 0, zIndex: 100 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ width: 28, height: 28, background: "linear-gradient(135deg,#E8C67A,#C9963A)", borderRadius: 7, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 13, fontWeight: 700, color: "#0C0F14" }}>B</div>
          <span style={{ fontSize: 16, fontFamily: "'DM Serif Display',serif", color: "#F0EEE8" }}>BizLens</span>
          <span style={{ fontSize: 10, color: "#4ADE80", background: "rgba(74,222,128,0.1)", padding: "2px 8px", borderRadius: 99 }}>LIVE</span>
        </div>
        <nav style={{ display: "flex", gap: 2 }}>
          {tabs.map(t => (
            <button key={t.id} onClick={() => setTab(t.id)} style={{ padding: "5px 12px", borderRadius: 8, fontSize: 13, cursor: "pointer", border: "none", display: "flex", alignItems: "center", gap: 5, transition: "all 0.15s", background: tab === t.id ? "rgba(232,198,122,0.12)" : "transparent", color: tab === t.id ? "#E8C67A" : "#6B7A8D", fontWeight: tab === t.id ? 500 : 400 }}>
              {t.icon} {t.label}
              {t.id === "banking" && <span style={{ fontSize: 9, background: "rgba(96,165,250,0.2)", color: "#60A5FA", padding: "1px 5px", borderRadius: 99 }}>NEW</span>}
            </button>
          ))}
        </nav>
        <span style={{ fontSize: 12, color: "#6B7A8D" }}><span style={{ color: "#4ADE80" }}>●</span> 3 sources</span>
      </header>

      <main style={{ flex: 1, padding: "30px 32px", maxWidth: 1200, width: "100%", margin: "0 auto" }}>

        {/* ── OVERVIEW ── */}
        {tab === "overview" && (
          <div>
            <div style={{ marginBottom: 22 }}>
              <h2 style={{ fontSize: 18, fontFamily: "'DM Serif Display',serif", color: "#F0EEE8", marginBottom: 4 }}>Business Overview</h2>
              <p style={{ fontSize: 12, color: "#6B7A8D" }}>Last 30 days · Shopify + Stripe + Google Sheets</p>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 14, marginBottom: 22 }}>
              {[
                { label: "Total Revenue", value: "$33,900", delta: "▲ 19% vs last month", dc: "#4ADE80" },
                { label: "Net Profit", value: "$18,500", delta: "▲ 26% vs last month", dc: "#4ADE80" },
                { label: "Total Expenses", value: "$15,400", delta: "▼ 12% increase", dc: "#F87171" },
                { label: "Profit Margin", value: "54.6%", delta: "▲ 3.2pts improvement", dc: "#4ADE80" },
              ].map(s => (
                <div key={s.label} style={{ ...card, padding: "18px 20px" }}>
                  <p style={{ fontSize: 10, color: "#6B7A8D", textTransform: "uppercase", letterSpacing: "0.12em", marginBottom: 8 }}>{s.label}</p>
                  <p style={{ fontSize: 24, fontFamily: "'DM Mono',monospace", color: "#F0EEE8", marginBottom: 4 }}>{s.value}</p>
                  <p style={{ fontSize: 12, color: s.dc }}>{s.delta}</p>
                </div>
              ))}
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 18, marginBottom: 18 }}>
              <div style={card}>
                <p style={{ fontSize: 14, fontWeight: 500, color: "#F0EEE8", marginBottom: 2 }}>Revenue vs Expenses</p>
                <p style={{ fontSize: 11, color: "#6B7A8D", marginBottom: 18 }}>7-month view</p>
                <ResponsiveContainer width="100%" height={190}>
                  <AreaChart data={revenueData} margin={{ top: 5, right: 8, left: -12, bottom: 0 }}>
                    <defs>
                      <linearGradient id="rg" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#E8C67A" stopOpacity={0.25} /><stop offset="95%" stopColor="#E8C67A" stopOpacity={0} /></linearGradient>
                      <linearGradient id="eg" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#F87171" stopOpacity={0.2} /><stop offset="95%" stopColor="#F87171" stopOpacity={0} /></linearGradient>
                    </defs>
                    <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.04)" />
                    <XAxis dataKey="month" tick={{ fontSize: 11, fill: "#6B7A8D" }} axisLine={false} tickLine={false} />
                    <YAxis tick={{ fontSize: 11, fill: "#6B7A8D" }} axisLine={false} tickLine={false} tickFormatter={v => `$${(v/1000).toFixed(0)}k`} />
                    <Tooltip content={<Tip />} />
                    <Area type="monotone" dataKey="revenue" stroke="#E8C67A" strokeWidth={2} fill="url(#rg)" name="Revenue" />
                    <Area type="monotone" dataKey="expenses" stroke="#F87171" strokeWidth={2} fill="url(#eg)" name="Expenses" />
                  </AreaChart>
                </ResponsiveContainer>
              </div>
              <div style={card}>
                <p style={{ fontSize: 14, fontWeight: 500, color: "#F0EEE8", marginBottom: 2 }}>Revenue by Category</p>
                <p style={{ fontSize: 11, color: "#6B7A8D", marginBottom: 18 }}>February breakdown</p>
                <ResponsiveContainer width="100%" height={190}>
                  <BarChart data={categoryData} margin={{ top: 5, right: 8, left: -12, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.04)" vertical={false} />
                    <XAxis dataKey="category" tick={{ fontSize: 11, fill: "#6B7A8D" }} axisLine={false} tickLine={false} />
                    <YAxis tick={{ fontSize: 11, fill: "#6B7A8D" }} axisLine={false} tickLine={false} tickFormatter={v => `${v}%`} />
                    <Tooltip contentStyle={{ background: "#141820", border: "1px solid rgba(255,255,255,0.1)", borderRadius: 8, fontSize: 12 }} />
                    <Bar dataKey="value" fill="#E8C67A" radius={[4,4,0,0]} name="Share %" />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
            <div style={card}>
              <p style={{ fontSize: 14, fontWeight: 500, color: "#F0EEE8", marginBottom: 2 }}>Monthly Net Profit</p>
              <p style={{ fontSize: 11, color: "#6B7A8D", marginBottom: 16 }}>7-month trend</p>
              <ResponsiveContainer width="100%" height={130}>
                <BarChart data={revenueData} margin={{ top: 5, right: 8, left: -12, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.04)" vertical={false} />
                  <XAxis dataKey="month" tick={{ fontSize: 11, fill: "#6B7A8D" }} axisLine={false} tickLine={false} />
                  <YAxis tick={{ fontSize: 11, fill: "#6B7A8D" }} axisLine={false} tickLine={false} tickFormatter={v => `$${(v/1000).toFixed(0)}k`} />
                  <Tooltip content={<Tip />} />
                  <Bar dataKey="profit" fill="#4ADE80" radius={[4,4,0,0]} name="Profit" />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>
        )}

        {/* ── PERSONAL BANKING ── */}
        {tab === "banking" && <PersonalBankTab />}

        {/* ── ASK AI ── */}
        {tab === "ask" && (
          <div style={{ display: "flex", flexDirection: "column", height: "calc(100vh - 130px)" }}>
            <div style={{ marginBottom: 18 }}>
              <h2 style={{ fontSize: 18, fontFamily: "'DM Serif Display',serif", color: "#F0EEE8", marginBottom: 4 }}>Ask Your Data</h2>
              <p style={{ fontSize: 12, color: "#6B7A8D" }}>Claude has full context of your business metrics</p>
            </div>
            <div style={{ ...card, flex: 1, display: "flex", flexDirection: "column", overflow: "hidden" }}>
              <div style={{ display: "flex", gap: 7, flexWrap: "wrap", marginBottom: 16 }}>
                {["Why did revenue spike in December?", "What's my most profitable category?", "Are expenses under control?", "What should I focus on this quarter?"].map(s => (
                  <button key={s} onClick={() => setInput(s)} style={{ background: "rgba(232,198,122,0.08)", border: "1px solid rgba(232,198,122,0.2)", borderRadius: 99, padding: "5px 12px", fontSize: 12, color: "#C9A84C", cursor: "pointer" }}>{s}</button>
                ))}
              </div>
              <div style={{ flex: 1, overflowY: "auto", display: "flex", flexDirection: "column", gap: 12, paddingRight: 4 }}>
                {msgs.map((m, i) => (
                  <div key={i} style={{ display: "flex", flexDirection: m.role === "user" ? "row-reverse" : "row", gap: 10, alignItems: "flex-start" }}>
                    <div style={{ width: 28, height: 28, borderRadius: 7, flexShrink: 0, background: m.role === "user" ? "rgba(99,91,255,0.2)" : "rgba(232,198,122,0.15)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 11, color: m.role === "user" ? "#A5A0FF" : "#E8C67A" }}>{m.role === "user" ? "U" : "AI"}</div>
                    <div style={{ maxWidth: "78%", background: m.role === "user" ? "rgba(99,91,255,0.1)" : "rgba(255,255,255,0.04)", border: `1px solid ${m.role === "user" ? "rgba(99,91,255,0.2)" : "rgba(255,255,255,0.07)"}`, borderRadius: 10, padding: "11px 14px", fontSize: 14, lineHeight: 1.65, color: "#D8D4CC", whiteSpace: "pre-wrap" }}>{m.content}</div>
                  </div>
                ))}
                {loading && (
                  <div style={{ display: "flex", gap: 10 }}>
                    <div style={{ width: 28, height: 28, borderRadius: 7, background: "rgba(232,198,122,0.15)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 11, color: "#E8C67A" }}>AI</div>
                    <div style={{ background: "rgba(255,255,255,0.04)", border: "1px solid rgba(255,255,255,0.07)", borderRadius: 10, padding: "12px 14px", display: "flex", gap: 5 }}>
                      {[0,1,2].map(i => <div key={i} className="dp" style={{ width: 6, height: 6, borderRadius: "50%", background: "#E8C67A", animationDelay: `${i*0.22}s` }} />)}
                    </div>
                  </div>
                )}
                <div ref={bottomRef} />
              </div>
              <div style={{ marginTop: 12, display: "flex", gap: 10 }}>
                <input value={input} onChange={e => setInput(e.target.value)} onKeyDown={e => e.key === "Enter" && send()}
                  placeholder="Ask about revenue, forecasts, expenses..." style={{ flex: 1, background: "rgba(255,255,255,0.05)", border: "1px solid rgba(255,255,255,0.1)", borderRadius: 10, padding: "11px 16px", fontSize: 14, color: "#F0EEE8" }} />
                <button onClick={send} disabled={loading || !input.trim()} style={{ background: loading || !input.trim() ? "rgba(232,198,122,0.15)" : "linear-gradient(135deg,#E8C67A,#C9963A)", border: "none", borderRadius: 10, padding: "11px 20px", fontSize: 14, color: loading || !input.trim() ? "#6B7A8D" : "#0C0F14", cursor: "pointer", fontWeight: 600 }}>Send →</button>
              </div>
            </div>
          </div>
        )}

        {/* ── INSIGHTS ── */}
        {tab === "alerts" && (
          <div>
            <div style={{ marginBottom: 22 }}>
              <h2 style={{ fontSize: 18, fontFamily: "'DM Serif Display',serif", color: "#F0EEE8", marginBottom: 4 }}>AI-Generated Insights</h2>
              <p style={{ fontSize: 12, color: "#6B7A8D" }}>Automatically surfaced from connected data</p>
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 12, marginBottom: 22 }}>
              {alerts.map(a => (
                <div key={a.id} style={{ ...card, borderLeft: `3px solid ${a.type==="positive"?"#4ADE80":a.type==="warning"?"#FBBF24":"#60A5FA"}`, display: "flex", gap: 16, alignItems: "flex-start" }}>
                  <div style={{ width: 36, height: 36, borderRadius: 9, flexShrink: 0, background: a.type==="positive"?"rgba(74,222,128,0.12)":a.type==="warning"?"rgba(251,191,36,0.12)":"rgba(96,165,250,0.12)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16, color: a.type==="positive"?"#4ADE80":a.type==="warning"?"#FBBF24":"#60A5FA" }}>{a.icon}</div>
                  <div><p style={{ fontWeight: 500, color: "#F0EEE8", marginBottom: 4 }}>{a.title}</p><p style={{ color: "#8A939F", fontSize: 13, lineHeight: 1.6 }}>{a.body}</p></div>
                </div>
              ))}
            </div>
            <div style={card}>
              <p style={{ fontSize: 15, fontWeight: 500, color: "#F0EEE8", marginBottom: 18 }}>Performance Scores</p>
              {[{ label:"Revenue Growth",score:82,color:"#E8C67A"},{label:"Expense Efficiency",score:67,color:"#F87171"},{label:"Profit Margin",score:91,color:"#4ADE80"},{label:"Revenue Consistency",score:74,color:"#60A5FA"}].map(m => (
                <div key={m.label} style={{ marginBottom: 14 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 5 }}>
                    <span style={{ fontSize: 13, color: "#9AA3AF" }}>{m.label}</span>
                    <span style={{ fontSize: 13, fontFamily: "'DM Mono',monospace", color: m.color }}>{m.score}/100</span>
                  </div>
                  <div style={{ height: 5, background: "rgba(255,255,255,0.06)", borderRadius: 99 }}>
                    <div style={{ width: `${m.score}%`, height: "100%", background: m.color, borderRadius: 99 }} />
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* ── FORECAST ── */}
        {tab === "forecast" && (
          <div>
            <div style={{ marginBottom: 22 }}>
              <h2 style={{ fontSize: 18, fontFamily: "'DM Serif Display',serif", color: "#F0EEE8", marginBottom: 4 }}>Revenue Forecast</h2>
              <p style={{ fontSize: 12, color: "#6B7A8D" }}>AI projection · Next 4 months</p>
            </div>
            <div style={{ ...card, marginBottom: 18 }}>
              <ResponsiveContainer width="100%" height={250}>
                <LineChart data={forecastData} margin={{ top: 10, right: 18, left: -10, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.04)" />
                  <XAxis dataKey="month" tick={{ fontSize: 11, fill: "#6B7A8D" }} axisLine={false} tickLine={false} />
                  <YAxis tick={{ fontSize: 11, fill: "#6B7A8D" }} axisLine={false} tickLine={false} tickFormatter={v => `$${(v/1000).toFixed(0)}k`} />
                  <Tooltip content={<Tip />} />
                  <Line type="monotone" dataKey="actual" stroke="#E8C67A" strokeWidth={2.5} dot={{ fill: "#E8C67A", r: 5 }} name="Actual" />
                  <Line type="monotone" dataKey="forecast" stroke="#60A5FA" strokeWidth={2.5} strokeDasharray="6 3" dot={{ fill: "#60A5FA", r: 5 }} name="Forecast" />
                </LineChart>
              </ResponsiveContainer>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 14 }}>
              {[{month:"March",value:"$36,500",pct:"+7.7%"},{month:"April",value:"$39,200",pct:"+7.4%"},{month:"May",value:"$43,100",pct:"+10.0%"},{month:"June",value:"$47,800",pct:"+10.9%"}].map(f => (
                <div key={f.month} style={card}>
                  <p style={{ fontSize: 10, color: "#6B7A8D", textTransform: "uppercase", letterSpacing: "0.12em", marginBottom: 8 }}>{f.month}</p>
                  <p style={{ fontSize: 22, fontFamily: "'DM Mono',monospace", color: "#F0EEE8", marginBottom: 4 }}>{f.value}</p>
                  <p style={{ fontSize: 12, color: "#4ADE80" }}>▲ {f.pct} projected</p>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* ── SOURCES ── */}
        {tab === "sources" && (
          <div>
            <div style={{ marginBottom: 22 }}>
              <h2 style={{ fontSize: 18, fontFamily: "'DM Serif Display',serif", color: "#F0EEE8", marginBottom: 4 }}>Data Sources</h2>
              <p style={{ fontSize: 12, color: "#6B7A8D" }}>Connect your business tools to power AI analytics</p>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 14 }}>
              {connectors.map(c => (
                <div key={c.id} style={{ ...card, display: "flex", alignItems: "center", gap: 14, borderColor: c.status==="connected"?"rgba(74,222,128,0.2)":"rgba(255,255,255,0.07)" }}>
                  <div style={{ width: 40, height: 40, borderRadius: 9, background: c.status==="connected"?c.color+"22":"rgba(255,255,255,0.04)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18, color: c.status==="connected"?c.color:"#3E4348" }}>{c.icon}</div>
                  <div style={{ flex: 1 }}>
                    <p style={{ fontWeight: 500, color: "#F0EEE8", marginBottom: 3 }}>{c.label}</p>
                    <p style={{ fontSize: 12, color: c.status==="connected"?"#4ADE80":c.status==="pending"?"#FBBF24":"#6B7A8D" }}>{c.status==="connected"?"● Connected":c.status==="pending"?"◌ Auth required":"○ Not connected"}</p>
                  </div>
                  <button style={{ background: c.status==="connected"?"rgba(74,222,128,0.1)":"rgba(232,198,122,0.1)", border: `1px solid ${c.status==="connected"?"rgba(74,222,128,0.3)":"rgba(232,198,122,0.3)"}`, borderRadius: 8, padding: "5px 12px", fontSize: 11, color: c.status==="connected"?"#4ADE80":"#E8C67A", cursor: "pointer" }}>{c.status==="connected"?"Manage":"Connect"}</button>
                </div>
              ))}
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
