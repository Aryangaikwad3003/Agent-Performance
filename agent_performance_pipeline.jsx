import { useState, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

// ── colour tokens ──────────────────────────────────────────────
const C = {
  bg: "#0F1117",
  surface: "#181C27",
  card: "#1E2333",
  border: "#2A3050",
  accent: "#4F8EF7",
  accentDim: "#1E3A6E",
  green: "#34D399",
  greenDim: "#0D3D2D",
  amber: "#FBBF24",
  amberDim: "#3D2D00",
  text: "#E8ECF4",
  muted: "#6B7592",
  danger: "#F87171",
};

const FILE_TYPES = [
  {
    id: "unified",
    label: "Unified Data",
    desc: "Large CRM export · ~50 k rows",
    icon: "🗄️",
    hint: "Needs: assigned_to · call_attempt_count · campaign_name · contact_name",
  },
  {
    id: "order",
    label: "Order Data",
    desc: "Pivot table with Grand Total per agent",
    icon: "📦",
    hint: "Needs: Agent Name · Grand Total columns",
  },
  {
    id: "quality",
    label: "Quality Score",
    desc: "CQ audit scores per agent",
    icon: "⭐",
    hint: "Needs: Agent Name · CQ Scores columns",
  },
];

// ── helpers ────────────────────────────────────────────────────
function readFileAsync(file) {
  return new Promise((res, rej) => {
    const reader = new FileReader();
    reader.onload = (e) => res(e.target.result);
    reader.onerror = () => rej(new Error("File read failed"));
    reader.readAsArrayBuffer(file);
  });
}

function parseWorkbook(ab) {
  return XLSX.read(ab, { type: "array", cellDates: false });
}

function sheetToRows(wb, sheetIdx = 0) {
  const ws = wb.Sheets[wb.SheetNames[sheetIdx]];
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
}

function sheetToObjects(wb, sheetIdx = 0) {
  const ws = wb.Sheets[wb.SheetNames[sheetIdx]];
  return XLSX.utils.sheet_to_json(ws, { defval: null, raw: true });
}

// ── core processing ────────────────────────────────────────────
async function processFiles(files, onProgress) {
  // ---------- 1. UNIFIED DATA ----------
  onProgress(5, "Reading Unified Data file…");
  const uAB = await readFileAsync(files.unified);

  onProgress(10, "Parsing Unified Data (this takes a moment for ~50 k rows)…");
  const uWB = parseWorkbook(uAB);
  const uRows = sheetToObjects(uWB);

  onProgress(20, "Aggregating Unified Data per agent…");

  const dialMap = {};      // agent → sum of call_attempt_count
  const campaignMap = {};  // agent → Set of unique campaign names
  const contactMap = {};   // agent → Set of unique contact names

  const CHUNK = 5000;
  for (let i = 0; i < uRows.length; i += CHUNK) {
    const slice = uRows.slice(i, i + CHUNK);
    for (const row of slice) {
      const agent = (row["assigned_to"] || "").toString().trim();
      if (!agent) continue;
      if (!dialMap[agent]) {
        dialMap[agent] = 0;
        campaignMap[agent] = new Set();
        contactMap[agent] = new Set();
      }
      dialMap[agent] += Number(row["call_attempt_count"]) || 0;
      if (row["campaign_name"]) campaignMap[agent].add(row["campaign_name"].toString().trim());
      if (row["contact_name"]) contactMap[agent].add(row["contact_name"].toString().trim());
    }
    // yield to paint
    await new Promise((r) => setTimeout(r, 0));
    const pct = 20 + Math.round(((i + CHUNK) / uRows.length) * 30);
    onProgress(Math.min(pct, 50), `Processing unified rows ${Math.min(i + CHUNK, uRows.length).toLocaleString()} / ${uRows.length.toLocaleString()}…`);
  }

  // ---------- 2. ORDER DATA ----------
  onProgress(55, "Reading Order Data file…");
  const oAB = await readFileAsync(files.order);
  const oWB = parseWorkbook(oAB);
  const oRawRows = sheetToRows(oWB);

  onProgress(60, "Extracting agent order totals…");
  // The pivot layout: row[1] is the header row, col[1]="Agent Name", col[11]="Grand Total"
  // Find header row dynamically
  let headerRow = null;
  let agentColIdx = -1;
  let grandTotalColIdx = -1;

  for (let i = 0; i < oRawRows.length; i++) {
    const r = oRawRows[i];
    const idx = r.findIndex((c) => c && c.toString().toLowerCase().includes("agent name"));
    if (idx !== -1) {
      headerRow = i;
      agentColIdx = idx;
      // find Grand Total col
      grandTotalColIdx = r.findIndex((c) => c && c.toString().toLowerCase().includes("grand total"));
      break;
    }
  }

  const orderMap = {}; // agent → grand total
  if (headerRow !== null) {
    for (let i = headerRow + 1; i < oRawRows.length; i++) {
      const r = oRawRows[i];
      const agentName = (r[agentColIdx] || "").toString().trim();
      if (!agentName || agentName.toLowerCase() === "grand total") continue;
      orderMap[agentName] = Number(r[grandTotalColIdx]) || 0;
    }
  }

  // Collect canonical agent list from Order data (the source of truth for agent list)
  const agentList = Object.keys(orderMap);

  // ---------- 3. QUALITY DATA ----------
  onProgress(65, "Reading Quality Score file…");
  const qAB = await readFileAsync(files.quality);
  const qWB = parseWorkbook(qAB);
  const qRows = sheetToObjects(qWB);

  onProgress(70, "Aggregating quality scores per agent…");
  // Find Agent Name and CQ Scores columns (case-insensitive)
  const qualCountMap = {};  // agent → count of audits
  const qualSumMap = {};    // agent → sum of CQ Scores

  let agentCol = null;
  let cqCol = null;
  if (qRows.length > 0) {
    const keys = Object.keys(qRows[0]);
    agentCol = keys.find((k) => k.toLowerCase().includes("agent")) || keys[0];
    cqCol = keys.find((k) => k.toLowerCase().includes("cq") || k.toLowerCase().includes("score")) || keys[1];
  }

  for (const row of qRows) {
    const agent = (row[agentCol] || "").toString().trim();
    if (!agent) continue;
    qualCountMap[agent] = (qualCountMap[agent] || 0) + 1;
    qualSumMap[agent] = (qualSumMap[agent] || 0) + (Number(row[cqCol]) || 0);
  }

  // ---------- 4. BUILD OUTPUT ----------
  onProgress(80, "Building output table…");

  // Merge all known agents (order list is canonical, but add any from unified not in order)
  const allAgents = [...new Set([
    ...agentList,
    ...Object.keys(dialMap),
  ])].filter((a) => a && a.toLowerCase() !== "grand total");

  // Helper: fuzzy-ish name match (trim + lowercase)
  function lookupAgent(map, name) {
    const key = Object.keys(map).find(
      (k) => k.trim().toLowerCase() === name.trim().toLowerCase()
    );
    return key ? map[key] : undefined;
  }

  const outputRows = allAgents.map((agent) => {
    const dialCount = lookupAgent(dialMap, agent) ?? 0;
    const campaignCount = campaignMap[agent.trim()]
      ? campaignMap[agent.trim()].size
      : (() => {
          const k = Object.keys(campaignMap).find(
            (k) => k.toLowerCase() === agent.trim().toLowerCase()
          );
          return k ? campaignMap[k].size : 0;
        })();
    const contactCount = contactMap[agent.trim()]
      ? contactMap[agent.trim()].size
      : (() => {
          const k = Object.keys(contactMap).find(
            (k) => k.toLowerCase() === agent.trim().toLowerCase()
          );
          return k ? contactMap[k].size : 0;
        })();
    const grandTotal = lookupAgent(orderMap, agent) ?? 0;
    const auditCount = lookupAgent(qualCountMap, agent) ?? 0;
    const avgCQ =
      auditCount > 0
        ? Math.round(((lookupAgent(qualSumMap, agent) ?? 0) / auditCount) * 100) / 100
        : 0;

    return {
      "Agent Name": agent,
      "Dial Count": dialCount,
      "Campaign Count": campaignCount,
      "Contacts Assigned": contactCount,
      "Total Orders Achieved": grandTotal,
      "Order Target": "",
      "Call Audit Count": auditCount,
      "Average CQ Score": avgCQ,
    };
  });

  onProgress(90, "Generating Excel output file…");

  // Build workbook with formatting
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(outputRows);

  // Column widths
  ws["!cols"] = [
    { wch: 22 }, { wch: 12 }, { wch: 16 }, { wch: 20 },
    { wch: 22 }, { wch: 14 }, { wch: 16 }, { wch: 18 },
  ];

  XLSX.utils.book_append_sheet(wb, ws, "Agent Performance");

  onProgress(95, "Finalising…");
  const wbOut = XLSX.write(wb, { bookType: "xlsx", type: "array" });

  onProgress(100, "Done!");
  return { wbOut, rowCount: outputRows.length };
}

// ── UI components ──────────────────────────────────────────────
function FileCard({ type, selected, file, onSelect, onFile }) {
  const inputRef = useRef();
  const isSelected = selected === type.id;
  const hasFile = !!file;

  return (
    <div
      onClick={() => onSelect(type.id)}
      style={{
        background: isSelected ? C.accentDim : C.card,
        border: `1.5px solid ${isSelected ? C.accent : hasFile ? C.green : C.border}`,
        borderRadius: 12,
        padding: "18px 20px",
        cursor: "pointer",
        transition: "all 0.18s ease",
        display: "flex",
        flexDirection: "column",
        gap: 8,
        userSelect: "none",
      }}
    >
      <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
        {/* radio indicator */}
        <div
          style={{
            width: 18, height: 18, borderRadius: "50%",
            border: `2px solid ${isSelected ? C.accent : C.muted}`,
            background: isSelected ? C.accent : "transparent",
            flexShrink: 0,
            display: "flex", alignItems: "center", justifyContent: "center",
          }}
        >
          {isSelected && (
            <div style={{ width: 7, height: 7, borderRadius: "50%", background: "#fff" }} />
          )}
        </div>
        <span style={{ fontSize: 20 }}>{type.icon}</span>
        <div>
          <div style={{ fontWeight: 700, fontSize: 15, color: C.text }}>{type.label}</div>
          <div style={{ fontSize: 12, color: C.muted }}>{type.desc}</div>
        </div>
        {hasFile && (
          <div style={{
            marginLeft: "auto", background: C.greenDim, color: C.green,
            borderRadius: 6, padding: "2px 9px", fontSize: 11, fontWeight: 600,
          }}>
            ✓ Loaded
          </div>
        )}
      </div>

      {hasFile && (
        <div style={{
          fontSize: 12, color: C.green, background: C.greenDim,
          borderRadius: 6, padding: "4px 10px",
          overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
        }}>
          📄 {file.name}
        </div>
      )}

      {isSelected && !hasFile && (
        <div style={{ fontSize: 11, color: C.accent, opacity: 0.8 }}>{type.hint}</div>
      )}

      <input
        ref={inputRef}
        type="file"
        accept=".xlsx,.xls,.csv"
        style={{ display: "none" }}
        onChange={(e) => {
          if (e.target.files[0]) onFile(type.id, e.target.files[0]);
        }}
      />

      {isSelected && (
        <button
          onClick={(ev) => { ev.stopPropagation(); inputRef.current.click(); }}
          style={{
            background: C.accent, color: "#fff", border: "none",
            borderRadius: 8, padding: "8px 0", fontSize: 13,
            fontWeight: 600, cursor: "pointer", marginTop: 4,
          }}
        >
          {hasFile ? "Replace file" : "Choose file (.xlsx / .csv)"}
        </button>
      )}
    </div>
  );
}

function ProgressBar({ pct, label }) {
  return (
    <div style={{ width: "100%" }}>
      <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
        <span style={{ fontSize: 13, color: C.muted }}>{label}</span>
        <span style={{ fontSize: 13, color: C.accent, fontWeight: 700 }}>{pct}%</span>
      </div>
      <div style={{ background: C.border, borderRadius: 999, height: 8, overflow: "hidden" }}>
        <div style={{
          width: `${pct}%`, height: "100%",
          background: `linear-gradient(90deg, ${C.accent}, ${C.green})`,
          borderRadius: 999, transition: "width 0.3s ease",
        }} />
      </div>
    </div>
  );
}

function LogLine({ lines }) {
  return (
    <div style={{
      background: C.bg, border: `1px solid ${C.border}`, borderRadius: 8,
      padding: "10px 14px", fontFamily: "monospace", fontSize: 12,
      color: C.muted, maxHeight: 130, overflowY: "auto",
    }}>
      {lines.map((l, i) => (
        <div key={i} style={{ color: i === lines.length - 1 ? C.accent : C.muted }}>
          {i === lines.length - 1 ? "▶ " : "✓ "}{l}
        </div>
      ))}
    </div>
  );
}

// ── main app ───────────────────────────────────────────────────
export default function AgentPipeline() {
  const [selected, setSelected] = useState("unified");
  const [files, setFiles] = useState({ unified: null, order: null, quality: null });
  const [status, setStatus] = useState("idle"); // idle | processing | done | error
  const [progress, setProgress] = useState(0);
  const [logs, setLogs] = useState([]);
  const [result, setResult] = useState(null);
  const [errMsg, setErrMsg] = useState("");

  const allLoaded = files.unified && files.order && files.quality;

  const addLog = useCallback((msg) => {
    setLogs((prev) => [...prev.slice(-40), msg]);
  }, []);

  function handleFile(typeId, file) {
    setFiles((prev) => ({ ...prev, [typeId]: file }));
    // auto-advance to next empty slot
    const order = ["unified", "order", "quality"];
    const next = order.find((id) => id !== typeId && !files[id]);
    if (next) setSelected(next);
  }

  async function handleProcess() {
    setStatus("processing");
    setProgress(0);
    setLogs([]);
    setErrMsg("");
    setResult(null);

    try {
      const { wbOut, rowCount } = await processFiles(files, (pct, msg) => {
        setProgress(pct);
        addLog(msg);
      });
      setResult({ wbOut, rowCount });
      setStatus("done");
    } catch (e) {
      setErrMsg(e.message || "Unknown error");
      setStatus("error");
    }
  }

  function handleDownload() {
    if (!result) return;
    const blob = new Blob([result.wbOut], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Agent_Performance_${new Date().toISOString().slice(0, 10)}.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
  }

  function handleReset() {
    setFiles({ unified: null, order: null, quality: null });
    setStatus("idle");
    setProgress(0);
    setLogs([]);
    setResult(null);
    setSelected("unified");
  }

  const loadedCount = [files.unified, files.order, files.quality].filter(Boolean).length;

  return (
    <div style={{
      minHeight: "100vh", background: C.bg, fontFamily: "'Inter', 'Segoe UI', sans-serif",
      color: C.text, padding: "32px 16px",
    }}>
      <div style={{ maxWidth: 640, margin: "0 auto" }}>

        {/* header */}
        <div style={{ marginBottom: 32, textAlign: "center" }}>
          <div style={{
            display: "inline-block", background: C.accentDim, color: C.accent,
            borderRadius: 8, padding: "4px 14px", fontSize: 12,
            fontWeight: 600, letterSpacing: "0.08em", marginBottom: 12,
            textTransform: "uppercase",
          }}>
            Agent Performance Pipeline
          </div>
          <h1 style={{ margin: 0, fontSize: 28, fontWeight: 800, letterSpacing: "-0.5px" }}>
            Consolidate your agent data
          </h1>
          <p style={{ color: C.muted, fontSize: 14, margin: "8px 0 0" }}>
            Upload 3 source files → get one clean performance sheet
          </p>
        </div>

        {/* output columns preview */}
        <div style={{
          background: C.card, border: `1px solid ${C.border}`, borderRadius: 12,
          padding: "12px 16px", marginBottom: 24,
          display: "flex", flexWrap: "wrap", gap: "6px 10px",
        }}>
          <span style={{ fontSize: 11, color: C.muted, width: "100%", marginBottom: 2 }}>
            OUTPUT COLUMNS
          </span>
          {[
            "Agent Name", "Dial Count", "Campaign Count", "Contacts Assigned",
            "Total Orders Achieved", "Order Target ✏️", "Call Audit Count", "Avg CQ Score",
          ].map((col) => (
            <span key={col} style={{
              background: C.surface, border: `1px solid ${C.border}`, borderRadius: 5,
              padding: "3px 9px", fontSize: 11, color: C.text, fontWeight: 500,
            }}>
              {col}
            </span>
          ))}
        </div>

        {/* step indicator */}
        <div style={{
          display: "flex", alignItems: "center", gap: 8, marginBottom: 16, fontSize: 13,
        }}>
          <div style={{
            background: loadedCount === 3 ? C.greenDim : C.accentDim,
            color: loadedCount === 3 ? C.green : C.accent,
            borderRadius: 6, padding: "3px 12px", fontWeight: 600,
          }}>
            {loadedCount} / 3 files loaded
          </div>
          {!allLoaded && (
            <span style={{ color: C.muted }}>Select a file type below, then upload it</span>
          )}
        </div>

        {/* file cards */}
        <div style={{ display: "flex", flexDirection: "column", gap: 12, marginBottom: 24 }}>
          {FILE_TYPES.map((type) => (
            <FileCard
              key={type.id}
              type={type}
              selected={selected}
              file={files[type.id]}
              onSelect={setSelected}
              onFile={handleFile}
            />
          ))}
        </div>

        {/* processing section */}
        {status === "idle" && (
          <button
            disabled={!allLoaded}
            onClick={handleProcess}
            style={{
              width: "100%", padding: "14px 0", borderRadius: 10,
              background: allLoaded ? C.accent : C.border,
              color: allLoaded ? "#fff" : C.muted,
              border: "none", fontSize: 15, fontWeight: 700,
              cursor: allLoaded ? "pointer" : "not-allowed",
              transition: "background 0.15s",
            }}
          >
            {allLoaded ? "⚙️  Process Files & Generate Report" : "Upload all 3 files to continue"}
          </button>
        )}

        {status === "processing" && (
          <div style={{
            background: C.card, border: `1px solid ${C.border}`, borderRadius: 12,
            padding: "20px 20px", display: "flex", flexDirection: "column", gap: 14,
          }}>
            <div style={{ fontWeight: 600, color: C.accent }}>⚙️ Processing…</div>
            <ProgressBar pct={progress} label="Overall progress" />
            <LogLine lines={logs} />
            <div style={{ fontSize: 12, color: C.muted }}>
              💡 All processing happens in your browser — no data leaves your device.
            </div>
          </div>
        )}

        {status === "done" && result && (
          <div style={{
            background: C.greenDim, border: `1.5px solid ${C.green}`, borderRadius: 12,
            padding: "20px", display: "flex", flexDirection: "column", gap: 14,
          }}>
            <div style={{ fontWeight: 700, fontSize: 16, color: C.green }}>
              ✅ Report ready — {result.rowCount} agents processed
            </div>
            <ProgressBar pct={100} label="Complete" />
            <LogLine lines={logs} />
            <div style={{ display: "flex", gap: 10 }}>
              <button
                onClick={handleDownload}
                style={{
                  flex: 1, padding: "12px 0", borderRadius: 9,
                  background: C.green, color: "#0a1a0f", border: "none",
                  fontSize: 14, fontWeight: 700, cursor: "pointer",
                }}
              >
                ⬇️ Download Agent_Performance.xlsx
              </button>
              <button
                onClick={handleReset}
                style={{
                  padding: "12px 18px", borderRadius: 9,
                  background: C.card, color: C.muted, border: `1px solid ${C.border}`,
                  fontSize: 13, cursor: "pointer",
                }}
              >
                Reset
              </button>
            </div>
          </div>
        )}

        {status === "error" && (
          <div style={{
            background: "#2A0D0D", border: `1.5px solid ${C.danger}`, borderRadius: 12,
            padding: "18px 20px",
          }}>
            <div style={{ color: C.danger, fontWeight: 700, marginBottom: 8 }}>
              ❌ Processing failed
            </div>
            <div style={{ fontFamily: "monospace", fontSize: 12, color: "#FFAAAA" }}>
              {errMsg}
            </div>
            <button
              onClick={handleReset}
              style={{
                marginTop: 12, padding: "8px 20px", borderRadius: 7,
                background: C.card, color: C.text, border: `1px solid ${C.border}`,
                cursor: "pointer", fontSize: 13,
              }}
            >
              Try again
            </button>
          </div>
        )}

        {/* footer note */}
        <div style={{
          marginTop: 28, textAlign: "center", fontSize: 11, color: C.muted, lineHeight: 1.6,
        }}>
          Accepts <strong style={{ color: C.text }}>.xlsx · .xls · .csv</strong> for all three files
          &nbsp;·&nbsp; Processed entirely in-browser &nbsp;·&nbsp;
          No data is uploaded to any server
        </div>
      </div>
    </div>
  );
}
